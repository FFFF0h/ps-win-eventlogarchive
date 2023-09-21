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
	
.PARAMETER ArchiveRetentionDays
    Numbers of days of retention of archived log files. Older archived files will be purged. 
	Set to -1 for infinite retention. Default is 182.
	
.PARAMETER EvtLogPath
    Path to the event log files. Default is %SystemRoot%\System32\Winevt\Logs"

.EXAMPLE
	EventLogArchive.ps1 -EventLogArchivePath D:\EventLog_Archive
	This example is the script running for the first time. The Eventlog archive path will be set as "D:\Eventlog_Archive"
	in the registry. 
	
#>
Param (
    # Local folder to store Evt Data Collection 
    [parameter(Position=0, Mandatory=$False)][String]$EventLogArchivePath = "C:\EvtLogArchive",
	[parameter(Position=1, Mandatory=$False)][string]$EvtLogPath = "$($Env:SystemRoot)\System32\Winevt\Logs"
	[parameter(Position=2, Mandatory=$False)][int]$ArchiveRetentionDays = 182,
	[parameter(Position=3, Mandatory=$False)][Switch]$Dry = $False
)


### INIT ###
if ($PSBoundParameters["Debug"]) { $DebugPreference = "Continue" }

# Add Zip assembly
Add-Type -assembly "System.IO.Compression.FileSystem"

# Global Variables
[int]$ScriptVer = 1702
[string]$MachineName = ((Get-WmiObject "Win32_ComputerSystem").Name)
[string]$EventSource = "Evt_LogMaintenance"
[string]$EventLogArchiveTemp = $($EventLogArchivePath + "\_temp")
[string]$HKLMLogMaintenancePath = "HKLM:\Software\Evt_Scripts\LogMaintenance"
[String]$CurrentScript = $MyInvocation.MyCommand.Definition

# ScheduleTaskXML 
[XML]$ScheduledTaskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2023-01-01T00:00:00</Date>
    <Author></Author>
    <Description>Automated Event Log Archive</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition>
        <Interval>PT30M</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <StartBoundary>2023-01-01T00:00:00</StartBoundary>
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


### SETUP ###
Switch (Test-Path $HKLMLogMaintenancePath) 
{
	$False {
		# Create script registry entries
		New-Item -Path $HKLMLogMaintenancePath -Force
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name ScriptVersion -Value $ScriptVer
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name EventLogArchivePath -Value $EventLogArchivePath
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name ArchiveRetentionDays -Value $ArchiveRetentionDays
		
		# Add Event Log entry for logging script actions
		EventCreate /ID 775 /L Application /T Information /SO $EventSource /D "Log Maintenance Script installation started"

		# Create event log archive directory
		Write-Debug "[SETUP]Creating Event Log archive directory: $EventLogArchivePath"
		try 
		{
			New-Item $EventLogArchivePath -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $EventLogArchiveTemp -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $($EventLogArchivePath + "\_Script") -type directory -ErrorAction Stop -Force | Out-Null
			$EventMessage = "Created Event Log Archive directory:`n`n$EventLogArchivePath"
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $EventMessage -Category 1 -EntryType Information			 
		}
		Catch 
		{
			# Unable to create the archive directory. The script will end.
			$EventMessage = "Unable to create Event Log Archive directory.`n`n$_`nThe script will now end."
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $EventMessage -Category 1 -EntryType Error
			Write-Error "Unable to create Event Log Archive directory: $_"
			Exit
		}
		
		# Copy script to archive location
		$ScriptCopyDest = $($EventLogArchivePath + "\_Script\" + $($CurrentScript.Split('\'))[4])
		Copy-Item $CurrentScript -Destination $ScriptCopyDest

		# Update Scheduled task XML with current logged on user
		$Creator = ($($env:userdomain) + "\" + $($env:username))
		$ScheduledTaskXML.Task.RegistrationInfo.Author = $($Creator)
		Write-Debug "[SETUP]Task Author Account set as: $($Creator)"
		
		# Update Scheduled Task with path to script
		$TaskArguments = $ScheduledTaskXML.Task.Actions.Exec.Arguments.Replace("REPLACE", "$($ScriptCopyDest)")
		$ScheduledTaskXML.Task.Actions.Exec.Arguments = $TaskArguments
		Write-Debug "[SETUP]Task Action Script Path: $($TaskArguments)"

		# Write Scheduled Task XML
		Write-Debug "[SETUP]Saving scheduled task XML to disk"
		$XMLExportPath = ($EventLogArchivePath + "\_Script\" + $MachineName + "-LogArchive.xml")
		$ScheduledTaskXML.Save($XMLExportPath)

		# Create scheduled task
		Write-Debug "[SETUP]Creating Scheduled Task for Event Log Archive"
		Schtasks /create /tn "Event Log Archive" /xml $XMLExportPath
		Start-Sleep -Seconds 10
		
		# First run scheduled task
		Write-Debug "[SETUP]Running task for first time"
		Schtasks /Run /tn "Event Log Archive"
	}
	$True {	
		Write-Debug "No configuration needed"
		$EventLogArchivePath = ((Get-ItemProperty -Path $HKLMLogMaintenancePath).EventLogArchivePath)
		Write-Debug "EventLog Archive path: $EventLogArchivePath"
		
		[string]$EventLogArchiveTemp = $($EventLogArchivePath + "\_temp")
		Write-Debug "EventLog Archive temp path: $EventLogArchiveTemp"
		
		$ArchiveRetentionDays = ((Get-ItemProperty -Path $HKLMLogMaintenancePath).ArchiveRetentionDays)
		Write-Debug "EventLog Archive retention days: $(&{If ($ArchiveRetentionDays -eq -1) { "Infinite" } Else { $ArchiveRetentionDays }})"
		
        $message = "Starting Evt EventLog Archive Tool"
		Write-EventLog -LogName Application -Source $EventSource -EventId 776 -Message $message -Category 1 -EntryType Information
	}
} 


### ARCHIVE CURRENT LOGS ###
# Collect event log configuration and status from local computer
$EventLogConfig = Get-WmiObject Win32_NTEventlogFile | Select LogFileName, Name, FileSize, MaxFileSize
Write-Debug "[$($EventLogConfig.count)] Event logs discovered"

# Process each discovered event log
foreach ($Log in $EventLogConfig)
{
	Write-Debug "Processing: $($Log.LogFileName)"

	# Determine size threshold to archive logs
	$LogSizeMB = ($Log.FileSize / 1mb)
	$LogMaxSizeMB = ($Log.MaxFileSize / 1mb)
	$AlarmSize = ($LogMaxSizeMB - ($LogMaxSizeMB * .25))
	Write-Debug "$($Log.LogFileName) will be archived at $AlarmSize MB"

	# Check current log files against threshold
	Switch ($LogSizeMB -lt $AlarmSize)
	{
		$True { Write-Debug "$($Log.LogfileName) Log below threshold" }
		$False {  
			# Event log archive location
			$EvtLogArchive = $($EventLogArchivePath + "\" + $($Log.LogFileName))

			# Check / Create directory for log
			if ((Test-Path $EvtLogArchive) -eq $False) { New-Item $EvtLogArchive -type directory -ErrorAction Stop -Force | Out-Null }

			# Export log to temp directory
			$tempFullPath = $EventLogArchiveTemp + "\" + $($Log.LogFileName) + "_TEMP.evt"
			$tempEvtLog = Get-WmiObject Win32_NTEventlogFile | ? {$_.LogFileName -eq $($log.LogFileName)}
			$tempEvtLog.BackupEventLog($tempFullPath)

			# Clear Security event log
			Write-Debug "Clearing log: $($Log.LogFileName)"
			Clear-EventLog -LogName $($Log.LogFileName)
		
			## ZIP exported event logs
			Write-Debug $EventLogArchiveTemp
			$ZipArchiveFile = ($EvtLogArchive + "\" + $MachineName + "_" + $($Log.LogFileName) + "_Archive_" + (Get-Date -Format MM.dd.yyyy-hhmm) + ".zip")
		
			Write-Debug "Compressing archived log: $ZipArchiveFile"
			[IO.Compression.ZIPFile]::CreateFromDirectory($EventLogArchiveTemp, $ZipArchiveFile)		
		
			# Delete event log temp file
			Write-Debug "Removing temp event log file"
			try { Remove-Item $tempFullPath -ErrorAction Stop }	      
			catch {}    
            
			# Write event log entry
			$message = "Security log size (" + $LogSizeMB + "mb) exceeded 75% of configured maximum and was archived to: " + $ZipArchiveFile
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $message -Category 1
		}
	}
}


### REMOVE EXPIRED LOGS ###
if ($ArchiveRetentionDays -ge 0)
{
	Write-Debug "Searching for expired Event Log archives"
	
	# Set archive retention
	$DelDate = (Get-Date).AddDays(-$ArchiveRetentionDays)
	
	# Search event log archive directory for logs older than retention period
	$ExpriedEventLogArchiveFiles = Get-ChildItem -Path $EventLogArchivePath -Recurse | ?{$_.CreationTime -lt $DelDate -and $_.Name -like "*.zip"} | Select Name, CreationTime, VersionInfo
	Write-Debug "[$($ExpriedEventLogArchiveFiles.count)] Expried eventlog archives found"
	
	if ($ExpriedEventLogArchiveFiles.count -ne 0) 
	{
		foreach ($OldLog in $ExpriedEventLogArchiveFiles)
		{
			Write-Debug "Removing: $($OldLog.VersionInfo.FileName)"
			try 
			{
				$EventMessage = "Removing expired Eventlog backup:`n`n$($OldLog.VersionInfo.FileName)"
				Remove-Item $OldLog.versioninfo.FileName -ErrorAction stop
				Write-EventLog -LogName Application -Source $EventSource -EventId 778 -Message $EventMessage -Category 1 -EntryType Information
			}
			catch 
			{
				$EventMessage = "Unable to remove expired Eventlog backup:`n`n$($OldLog.VersionInfo.FileName)`n`nError: $_"
				Write-EventLog -LogName Application -Source $EventSource -EventId 780 -Message $EventMessage -Category 9 -EntryType error
			}
		}		
	}
}
else
{
	Write-Debug "Bypass Event Log archives deletion"
}


### CHECK FOR AUTOMATICALLY ARCHIVED LOGS ###
# Check default Event Log location for any old archived logs
$AutoArchivedLogFiles = Get-ChildItem -Path $EvtLogPath | ?{$_.Name -like "Archive-*"} | Select Name, CreationTime, VersionInfo
Write-Debug "Searching $($EvtLogPath) for automatically archived eventlogs..."
Write-Debug "[$($AutoArchivedLogFiles.Count)] Auto archive logs found..."

# If there are auto archived files process each file and copy to archive directory
if (($AutoArchivedLogFiles).Count -gt 0)
{
	foreach ($AutoLog in $AutoArchivedLogFiles)
	{
		$EventLogName = ($AutoLog.Name.Split('-')[1])
		$AutoArchivedEventLogFullPath = $($Autolog.VersionInfo.FileName).ToString()
		$EvtLogArchive = $($EventLogArchivePath + "\" + $EventLogName)
		
		# Check / Create directory for log
		if ((Test-Path $EvtLogArchive) -eq $False)
		{
		    New-Item $EvtLogArchive -type directory -ErrorAction Stop -Force | Out-Null
		}
        
        write-debug "Moving archive log to temp directory: [$AutoArchivedEventLogFullPath => $EventLogArchiveTemp]"
        $TTCM = (Measure-Command {Move-Item -Path $AutoArchivedEventLogFullPath -Destination $EventLogArchiveTemp})		
		Write-Debug "Time to complete move [MM:SS]: $($TTCM.Minutes):$($TTCM.Seconds)"


		## ZIP exported event logs
		Write-Debug $EventLogArchiveTemp
		$ZipArchiveFile = ($EvtLogArchive + "\" + $MachineName + "_" + $EventLogName + $($AutoLog.Name) + ".zip")
		
        Write-Debug "Log to Arc: $AutoArchivedEventLogFullPath"
		Write-Debug "Compressing archived log: $ZipArchiveFile"
        [IO.Compression.ZIPFile]::CreateFromDirectory($EventLogArchiveTemp, $ZipArchiveFile)
		
		$EventMessage = "EventLog Auto Archive compressed and moved:`n`n$ZipArchiveFile"
		Write-EventLog -LogName Application -Source $EventSource -EventId 777 -Message $EventMessage -Category 1 -EntryType Information

		# Remove copied logs from the temp directory
        Get-ChildItem -Path $EventLogArchiveTemp | Remove-Item 
	}
}


### ENDING ###
# reset default debug preference
$DebugPreference = "SilentlyContinue"

$message = "Stopping Evt EventLog Archive Tool"
Write-EventLog -LogName Application -Source $EventSource -EventId 779 -Message $message -Category 1 -EntryType Information
