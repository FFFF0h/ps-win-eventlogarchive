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
	
.PARAMETER DestinationPath
	The path the script will use to store the archived eventlogs. This parameter is only used durring the first run of the script.
	The value will be saved to the computers registry, and the value from the registry will be used on subsiquent runs. 
	
.PARAMETER ArchiveRetentionDays
    Numbers of days of retention of archived log files. Older archived files will be purged. 
	Set to -1 for infinite retention. Default is 182. This parameter is only used durring the first run of the script.
	The value will be saved to the computers registry, and the value from the registry will be used on subsiquent runs. 
	
.PARAMETER EventLogSourcePath
    Path to the event log files. Default is %SystemRoot%\System32\Winevt\Logs". This parameter is only used durring the first run of the script.
	The value will be saved to the computers registry, and the value from the registry will be used on subsiquent runs. 
	
.PARAMETER Dry
    When present, do not remove any logs file.

.EXAMPLE
	EventLog-Archive.ps1 -DestinationPath D:\EventLog_Archive
	This example is the script running for the first time. The Eventlog archive path will be set as "D:\Eventlog_Archive"
	in the registry. 
	
#>
Param (
    # Local folder to store Evt Data Collection 
    [parameter(Position=0, Mandatory=$False)][String]$DestinationPath = "C:\EventLogArchive",
	[parameter(Position=1, Mandatory=$False)][string]$EventLogSourcePath = "$($Env:SystemRoot)\System32\Winevt\Logs",
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
[string]$TempDestinationPath = $($DestinationPath + "\_temp")
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
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name DestinationPath -Value $DestinationPath
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name ArchiveRetentionDays -Value $ArchiveRetentionDays
		
		# Add Event Log entry for logging script actions
		EventCreate /ID 775 /L Application /T Information /SO $EventSource /D "Log Maintenance Script installation started"

		# Create event log archive directory
		Write-Debug "[SETUP]Creating Event Log archive directory: $DestinationPath"
		try 
		{
			New-Item $DestinationPath -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $TempDestinationPath -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $($DestinationPath + "\_Script") -type directory -ErrorAction Stop -Force | Out-Null
			$eventMessage = "Created Event Log Archive directory:`n`n$DestinationPath"
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $eventMessage -Category 1 -EntryType Information			 
		}
		Catch 
		{
			# Unable to create the archive directory. The script will end.
			$eventMessage = "Unable to create Event Log Archive directory.`n`n$_`nThe script will now end."
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $eventMessage -Category 1 -EntryType Error
			Write-Error "Unable to create Event Log Archive directory: $_"
			Exit
		}
		
		# Copy script to archive location
		$scriptCopyDestinationination = $($DestinationPath + "\_Script\" + $($CurrentScript.Split('\'))[4])
		Copy-Item $CurrentScript -Destination $scriptCopyDestination

		# Update Scheduled task XML with current logged on user
		$creator = ($($env:UserDomain) + "\" + $($env:UserName))
		$ScheduledTaskXML.Task.RegistrationInfo.Author = $($creator)
		Write-Debug "[SETUP]Task Author Account set as: $($creator)"
		
		# Update Scheduled Task with path to script
		$taskArguments = $ScheduledTaskXML.Task.Actions.Exec.Arguments.Replace("REPLACE", "$($scriptCopyDestination)")
		$ScheduledTaskXML.Task.Actions.Exec.Arguments = $taskArguments
		Write-Debug "[SETUP]Task Action Script Path: $($taskArguments)"

		# Write Scheduled Task XML
		Write-Debug "[SETUP]Saving scheduled task XML to disk"
		$XMLExportPath = ($DestinationPath + "\_Script\" + $MachineName + "-LogArchive.xml")
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
		$DestinationPath = ((Get-ItemProperty -Path $HKLMLogMaintenancePath).DestinationPath)
		Write-Debug "EventLog Archive path: $DestinationPath"
		
		[string]$TempDestinationPath = $($DestinationPath + "\_temp")
		Write-Debug "EventLog Archive temp path: $TempDestinationPath"
		
		$ArchiveRetentionDays = ((Get-ItemProperty -Path $HKLMLogMaintenancePath).ArchiveRetentionDays)
		Write-Debug "EventLog Archive retention days: $(&{If ($ArchiveRetentionDays -eq -1) { "Infinite" } Else { $ArchiveRetentionDays }})"
		
        $message = "Starting Evt EventLog Archive Tool"
		Write-EventLog -LogName Application -Source $EventSource -EventId 776 -Message $message -Category 1 -EntryType Information
	}
} 


### ARCHIVE CURRENT LOGS ###
# Collect event log configuration and status from local computer
$eventLogConfig = Get-WmiObject Win32_NTEventlogFile | Select LogfileName, Name, FileSize, MaxFileSize
Write-Debug "[$($eventLogConfig.count)] Event logs discovered"

# Process each discovered event log
foreach ($Log in $eventLogConfig)
{
	Write-Debug "Processing: $($Log.LogFileName)"

	# Determine size threshold to archive logs
	$logSizeMB = ($Log.FileSize / 1mb)
	$logMaxSizeMB = ($Log.MaxFileSize / 1mb)
	$logAlarmSize = ($logMaxSizeMB - ($logMaxSizeMB * .25))
	Write-Debug "$($Log.LogfileName) will be archived at $logAlarmSize MB"

	# Check current log files against threshold
	Switch ($logSizeMB -lt $logAlarmSize)
	{
		$True { Write-Debug "$($Log.LogfileName) Log below threshold" }
		$False {  
			# Event log archive location
			$eventLogArchiveFolderPath = $($DestinationPath + "\" + $($Log.LogfileName))

			# Check / Create directory for log
			if ((Test-Path $eventLogArchiveFolderPath) -eq $False) 
			{ 
				New-Item $eventLogArchiveFolderPath -type directory -ErrorAction Stop -Force | Out-Null 
			}

			# Export log to temp directory
			$tempEventLogFullPath = $TempDestinationPath + "\" + $($Log.LogfileName) + "_TEMP.evt"
			$tempEvtLog = Get-WmiObject Win32_NTEventlogFile | ? {$_.LogfileName -eq $($log.LogfileName)}
			$tempEvtLog.BackupEventLog($tempEventLogFullPath)

			# Clear event log
			Write-Debug "Clearing log: $($Log.LogfileName)"
			if (!$Dry) { Clear-EventLog -LogName $($Log.LogfileName) }
		
			## ZIP exported event logs
			Write-Debug $TempDestinationPath
			$zipArchiveFile = ($eventLogArchiveFolderPath + "\" + $MachineName + "_" + $($Log.LogfileName) + "_Archive_" + (Get-Date -Format MM.dd.yyyy-hhmm) + ".zip")
		
			Write-Debug "Compressing archived log: $zipArchiveFile"
			[IO.Compression.ZIPFile]::CreateFromDirectory($TempDestinationPath, $zipArchiveFile)		
		
			# Delete event log temp file
			Write-Debug "Removing temp event log file"
			try { Remove-Item $tempEventLogFullPath -ErrorAction Stop }	      
			catch {}    
            
			# Write event log entry
			$message = "Security log size (" + $logSizeMB + "mb) exceeded 75% of configured maximum and was archived to: " + $zipArchiveFile
			Write-EventLog -LogName Application -Source $EventSource -EventId 775 -Message $message -Category 1
		}
	}
}


### REMOVE EXPIRED LOGS ###
if ($ArchiveRetentionDays -ge 0)
{
	Write-Debug "Searching for expired Event Log archives"
	
	# Set archive retention
	$DeletionDate = (Get-Date).AddDays(-$ArchiveRetentionDays)
	
	# Search event log archive directory for logs older than retention period
	$expiredEventLogArchiveFiles = Get-ChildItem -Path $DestinationPath -Recurse | ?{$_.CreationTime -lt $DeletionDate -and $_.Name -like "*.zip"} | Select Name, CreationTime, VersionInfo
	Write-Debug "[$($expiredEventLogArchiveFiles.Count)] Expired Eventlog archives found"
	
	if ($expiredEventLogArchiveFiles.count -ne 0) 
	{
		foreach ($currentExpiredLog in $expiredEventLogArchiveFiles)
		{
			Write-Debug "Removing: $($currentExpiredLog.VersionInfo.FileName)"
			try 
			{
				$eventMessage = "Removing expired Eventlog backup:`n`n$($currentExpiredLog.VersionInfo.FileName)"
				if (!$Dry) { Remove-Item $currentExpiredLog.versioninfo.FileName -ErrorAction stop }
				Write-EventLog -LogName Application -Source $EventSource -EventId 778 -Message $eventMessage -Category 1 -EntryType Information
			}
			catch 
			{
				$eventMessage = "Unable to remove expired Eventlog backup:`n`n$($currentExpiredLog.VersionInfo.FileName)`n`nError: $_"
				Write-EventLog -LogName Application -Source $EventSource -EventId 780 -Message $eventMessage -Category 9 -EntryType error
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
$autoArchivedLogFiles = Get-ChildItem -Path $EventLogSourcePath | ?{$_.Name -like "Archive-*"} | Select Name, CreationTime, VersionInfo
Write-Debug "Searching $($EventLogSourcePath) for automatically archived eventlogs..."
Write-Debug "[$($autoArchivedLogFiles.Count)] Auto archive logs found..."

# If there are auto archived files process each file and copy to archive directory
if (($autoArchivedLogFiles).Count -gt 0)
{
	foreach ($currentAutoLog in $autoArchivedLogFiles)
	{
		$eventLogName = ($currentAutoLog.Name.Split('-')[1])
		$autoArchivedEventLogFullPath = $($currentAutoLog.VersionInfo.FileName).ToString()
		$eventLogArchiveFolderPath = $($DestinationPath + "\" + $eventLogName)
		
		# Check / Create directory for log
		if ((Test-Path $eventLogArchiveFolderPath) -eq $False)
		{
		    New-Item $eventLogArchiveFolderPath -type directory -ErrorAction Stop -Force | Out-Null
		}
        
        write-debug "Moving archive log to temp directory: [$autoArchivedEventLogFullPath => $TempDestinationPath]"
        if (!$Dry)
		{
			$TTCM = (Measure-Command {Move-Item -Path $autoArchivedEventLogFullPath -Destination $TempDestinationPath})		
		}
		else
		{
			$TTCM = (Measure-Command {Copy-Item -Path $autoArchivedEventLogFullPath -Destination $TempDestinationPath})		
		}
		Write-Debug "Time to complete move [MM:SS]: $($TTCM.Minutes):$($TTCM.Seconds)"


		## ZIP exported event logs
		Write-Debug $TempDestinationPath
		$zipArchiveFile = ($eventLogArchiveFolderPath + "\" + $MachineName + "_" + $eventLogName + $($currentAutoLog.Name) + ".zip")
		
        Write-Debug "Log to Arc: $autoArchivedEventLogFullPath"
		Write-Debug "Compressing archived log: $zipArchiveFile"
        [IO.Compression.ZIPFile]::CreateFromDirectory($TempDestinationPath, $zipArchiveFile)
		
		$eventMessage = "EventLog Auto Archive compressed and moved:`n`n$zipArchiveFile"
		Write-EventLog -LogName Application -Source $EventSource -EventId 777 -Message $eventMessage -Category 1 -EntryType Information

		# Remove copied logs from the temp directory
        Get-ChildItem -Path $TempDestinationPath | Remove-Item
	}
}


### ENDING ###
# reset default debug preference
$DebugPreference = "SilentlyContinue"

$message = "Stopping Evt EventLog Archive Tool"
Write-EventLog -LogName Application -Source $EventSource -EventId 779 -Message $message -Category 1 -EntryType Information
