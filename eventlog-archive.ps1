<#
.NOTES

.SYNOPSIS
	Collect and archive Eventlogs on Windows servers

.DESCRIPTION
	This script will be used to automate the collection and archival of Windows event logs. When an eventlog exceeds 
	75% of the configured maximum size the log will be backed up, compressed, moved to the configured archive location
	and the log will be cleared. If no location is specified the script will default to the C:\ drive. It is recommended
	to set the archive path to another drive to move the logs from the default system drive. 
	
	In order to run continuously the script will created a scheduled task on the commputer to run every 30 minutes to
	to check the current status of event logs.   

	Status of the script will be written to the Application log [Evt_LogMaintenance]. 
	
.PARAMETER EventLogArchivePath
	The path the script will use to store the archived eventlogs. This parameter is only used durring the first run of the script.
	The value will be saved to the computers registry, and the value from the registry will be used on subsiquent runs. 

.EXAMPLE
	EventLogArchive.ps1 -EventLogArchivePath D:\EventLog_Archive
	This example is the script running for the first time. The Eventlog archive path will be set as "D:\Eventlog_Archive"
	in the registry. 
	
#>
Param (
    # Local folder to store Evt Data Collection 
    [parameter(Position=0, Mandatory=$False)][String]$EventLogArchivePath = "C:\EvtLogArchive",
	[parameter(Position=1, Mandatory=$False)][Switch]$Step = $False
)

If ($PSBoundParameters["Debug"]) {$DebugPreference = "Continue"}

# Add Zip assembly
Add-Type -assembly "system.io.compression.filesystem"

# Global Variables
[int]$ScriptVer = 1602
[string]$MachineName = ((get-wmiobject "Win32_ComputerSystem").Name)
[string]$EventSource = "Evt_LogMaintenance"
[string]$EventLogArchiveTemp = $($EventLogArchivePath+"\_temp")
[String]$CurrentScript = $MyInvocation.MyCommand.Definition

#region ScheduleTaskXML
[XML]$ScheduledTaskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2014-07-24T09:10:27.7100272</Date>
    <Author></Author>
    <Description>Automated Event Log Archive</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition>
        <Interval>PT30M</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <StartBoundary>2015-12-02T14:44:13</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>P3D</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>PowerShell.exe</Command>
      <Arguments>–Noninteractive –Noprofile –Command "REPLACE"</Arguments>
    </Exec>
  </Actions>
</Task>
"@
#endregion
#region Setup
Switch (Test-Path HKLM:\Software\Evt_Scripts\LogMaintenance) {
	$False {
		# Create script registry entries
		New-Item -Path HKLM:\Software\Evt_Scripts\LogMaintenance -Force
		New-ItemProperty -Path HKLM:\Software\Evt_Scripts\LogMaintenance -Name ScriptVersion -Value $ScriptVer
		New-ItemProperty -Path HKLM:\Software\Evt_Scripts\LogMaintenance -Name EventLogArchivePath -Value $EventLogArchivePath

		# Add Event Log entry for logging script actions
		eventcreate /ID 775 /L Application /T Information /SO $EventSource /D "Log Maintenance Script installation started"

		# Create event log archive directory
		Write-Debug "[SETUP]Creating Event Log archive directory: $EventLogArchivePath"
		try {
			New-Item $EventLogArchivePath -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $EventLogArchiveTemp -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $($EventLogArchivePath+"\_Script") -type directory -ErrorAction Stop -Force | Out-Null
			$EventMessage = "Created eventlog archive directory:`n`n$EventLogArchivePath"
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $EventMessage -Category 1 -EntryType Information			 
		}
		Catch {
			# Unable to create the archive directory. The script will end.
			$EventMessage = "Unable to create Event Log Archive directory.`n`n$_`nThe script will now end."
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $EventMessage -Category 1 -EntryType Error
			Write-Error "Unable to create Event Log Archvie directory: $_"
			Exit
		}
		# Copy script to archive location
		$ScriptCopyDest = $($EventLogArchivePath+"\_Script\"+$($CurrentScript.Split('\'))[4])
		Copy-Item $CurrentScript -Destination $ScriptCopyDest

		# Update Scheduled task XML with current logged on user
		$Creator = ($($env:userdomain)+"\"+$($env:username))
		$ScheduledTaskXML.task.RegistrationInfo.Author = $($Creator)
		Write-Debug "[SETUP]Task Author Account set as: $($Creator)"
		
		# Update Scheduled Task with path to script
		$TaskArguments = $ScheduledTaskXML.task.actions.exec.Arguments.replace("REPLACE", "$($ScriptCopyDest)")
		$ScheduledTaskXML.task.actions.exec.Arguments = $TaskArguments
		Write-Debug "[SETUP]Task Action Script Path: $($TaskArguments)"

		# Write Schedulded Task XML
		Write-Debug "[SETUP]Saving scheduled task XML to disk"
		$XMLExportPath = ($EventLogArchivePath+"\_Script\"+$MachineName+"-LogArchive.xml")
		$ScheduledTaskXML.save($XMLExportPath)

		# Create scheduled task
		Write-Debug "[SETUP]Creating Scheduled Task for Event Log Archive"
		schtasks /create /tn "Event Log Archive" /xml $XMLExportPath
		Start-Sleep -Seconds 10
		
		# First run scheduled task
		Write-Debug "[SETUP]Running task for first time"
		schtasks /Run /tn "Event Log Archive"
	}
	$True {	
		Write-Debug "No configuration needed"
		$EventLogArchivePath = ((Get-ItemProperty -Path HKLM:\Software\Evt_Scripts\LogMaintenance).EventLogArchivePath)
		Write-Debug "EventLog Archive path: $EventLogArchivePath"
		[string]$EventLogArchiveTemp = $($EventLogArchivePath+"\_temp")
		Write-Debug "EventLog Archive temp path: $EventLogArchiveTemp"
        $message = "Starting Evt EventLog Archive Tool"
		Write-EventLog -LogName Application -Source $EventSource -EventId 776 -Message $message -Category 1 -EntryType Information
	}
} 
#endregion
#region ArchiveLogs	
# Collect event log configuration and status from local computer
$EventLogConfig = Get-WmiObject Win32_NTEventlogFile | Select-Object LogFileName, Name, FileSize, MaxFileSize
Write-Debug "[$($EventLogConfig.count)] Event logs discovered"

# Process each discovered event log
foreach ($Log in $EventLogConfig){
	Write-Debug "Processing: $($Log.LogFileName)"

	# Determin size threshold to archive logs
	$LogSizeMB = ($Log.FileSize / 1mb)
	$LogMaxSizeMB = ($Log.MaxFileSize /1mb)
	$AlarmSize = ($LogMaxSizeMB - ($LogMaxSizeMB * .25))
	Write-Debug "$($Log.LogFileName) will be archived at $AlarmSize MB"

	# Check current log files against threshold
	Switch ($LogSizeMB -lt $AlarmSize){
		$True { Write-Debug "$($Log.LogfileName) Log below threshold"}
		$False{  
			# Event log archive location
			$EvtLogArchive = $($EventLogArchivePath+"\"+$($Log.logfilename))

			# Check / Create directory for log
			if ((Test-Path $EvtLogArchive) -eq $False){New-Item $EvtLogArchive -type directory -ErrorAction Stop -Force | Out-Null}

			# Export log to temp directory
			$tempFullPath = $EventLogArchiveTemp+"\"+$($Log.logfilename)+"_TEMP.evt"
			$tempEvtLog = Get-WmiObject Win32_NTEventlogFile | Where-Object {$_.logfilename -eq $($log.LogFileName)}
			$tempEvtLog.backupeventlog($tempFullPath)

			# Clear Security event log
			Write-Debug "Clearing log: $($Log.LogFileName)"
			Clear-EventLog -LogName $($Log.LogFileName)
		
			## ZIP exported event logs
			Write-Debug $EventLogArchiveTemp
			$ZipArchiveFile = ($EvtLogArchive+"\"+$MachineName+"_"+$($Log.LogFileName)+"_Archive_"+(Get-Date -Format MM.dd.yyyy-hhmm)+".zip")
		
			# Add Zip assembly
			#Add-Type -assembly "system.io.compression.filesystem"
			Write-Debug "Compressing archived log: $ZipArchiveFile"
			[io.compression.zipfile]::CreateFromDirectory($EventLogArchiveTemp, $ZipArchiveFile)		
		
			# Delete event log temp file
			Write-Debug "Removing temp event log file"
			try {Remove-Item $tempFullPath -ErrorAction Stop}	      
			catch {}    
            
			# Write event log entry
			$message = "Security log size ("+$LogSizeMB+"mb) exceeded 75% of configured maximum and was archived to: "+$ZipArchiveFile
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $message -Category 1
		}
	}
}
#endregion
#region RemoveOldLogs
Write-Debug "Searching for expired eventlog archives"

# Set archive retention
$DelDate = (Get-Date).AddDays(-182)

# Search event log archive directory for logs older than retention period
$ExpriedEventLogArchiveFiles = Get-ChildItem -Path $EventLogArchivePath -Recurse | Where-Object {$_.CreationTime -lt $DelDate -and $_.Name -like "*.zip"} | Select-Object Name, CreationTime, VersionInfo
Write-Debug "[$($ExpriedEventLogArchiveFiles.count)] Expried eventlog archives found"

if ($ExpriedEventLogArchiveFiles.count -ne 0){
	foreach ($OldLog in $ExpriedEventLogArchiveFiles){
		Write-Debug "Removing: $($OldLog.versioninfo.FileName)"
		try {
			$EventMessage = "Removing expired Eventlog backup:`n`n$($OldLog.versioninfo.FileName)"
			Remove-Item $OldLog.versioninfo.FileName -ErrorAction stop
			Write-EventLog -LogName Application -Source $EventSource -EventId 778 -Message $EventMessage -Category 1 -EntryType Information
		}
		catch {
			$EventMessage = "Unable to remove expired Eventlog backup:`n`n$($OldLog.versioninfo.FileName)`n`nError: $_"
			Write-EventLog -LogName Application -Source $EventSource -EventId 780 -Message $EventMessage -Category 9 -EntryType error
		}
	}		
}
#endregion
#region CheckForAutomaticallyArchivedLogs

# Check default Eventlog location for any old archived logs
[string]$DefaultEvtLogPath = "C:\Windows\System32\Winevt\Logs"
$AutoArchivedLogFiles = Get-ChildItem -Path $DefaultEvtLogPath | Where-Object {$_.Name -like "Archive-*"} | Select-Object Name, CreationTime, VersionInfo
Write-Debug "Searching $($DefaultEvtLogPath) for automatically archived eventlogs..."
Write-Debug "[$($AutoArchivedLogFiles.count)] Auto archive logs found..."

# If there are auto archived files process each file and copy to archive directory
if (($AutoArchivedLogFiles).count -gt 0){
	foreach ($AutoLog in $AutoArchivedLogFiles){
		$EventLogName = ($AutoLog.name.split('-')[1])
		$AutoArchivedEventLogPath = $($Autolog.VersionInfo.FileName).tostring()
		$EvtLogArchive = $($EventLogArchivePath+"\"+$EventLogName)
		
		# Check / Create directory for log
		if ((Test-Path $EvtLogArchive) -eq $False){New-Item $EvtLogArchive -type directory -ErrorAction Stop -Force | Out-Null}
		
		# Remove logs that are expired
		$DelDate = (Get-Date).AddDays(-182)
		$ExpiredAutoArchivedLogFiles = $AutoArchivedLogFiles | Where-Object {$_.CreationTime -lt $DelDate}
        
        write-debug "Moving archvie log to temp directory: [$AutoArchivedEventLogPath => $EventLogArchiveTemp]"
        $TTCM = (Measure-Command {Move-Item -Path $AutoArchivedEventLogPath -Destination $EventLogArchiveTemp})		
		Write-Debug "Time to complete move [MM:SS]: $($TTCM.Minutes):$($TTCM.Seconds)"


		## ZIP exported event logs
		Write-Debug $EventLogArchiveTemp
		$ZipArchiveFile = ($EvtLogArchive+"\"+$MachineName+"_"+$($EventLogName)+$($AutoLog.name)+".zip")
		
        Write-Debug "Log to Arc: $AutoArchivedEventLogPath"
		Write-Debug "Compressing archived log: $ZipArchiveFile"
        [io.compression.zipfile]::CreateFromDirectory($EventLogArchiveTemp, $ZipArchiveFile)
		
		$EventMessage = "EventLog Auto Archive compressed and moved:`n`n$ZipArchiveFile"
		Write-EventLog -LogName Application -Source $EventSource -EventId 777 -Message $EventMessage -Category 1 -EntryType Information

		# Remove copied logs from the temp directory
        Get-ChildItem -Path $EventLogArchiveTemp | Remove-Item 

		## For testing ##
		If ($Step -eq $True) {
			Write-Host -ForegroundColor Green "[AutoArchivedLogs] - Do you want to keep going?"
			$iii = read-host
		}

	}
}

#endregion

# reset default debug preference
$DebugPreference = "SilentlyContinue"

$message = "Stopping Evt EventLog Archive Tool"
Write-EventLog -LogName Application -Source $EventSource -EventId 779 -Message $message -Category 1 -EntryType Information
