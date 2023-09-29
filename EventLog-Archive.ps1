<#
.NOTES
	????/??/?? - Version 1602 - Daniel Hibbert - Initial version.
	2023/09/29 - Version 1802 - Laurent Le Guillermic - Enhancements & bug corrections.

.SYNOPSIS
	Collect and archive Eventlogs on Windows servers

.DESCRIPTION
	This script will be used to automate the collection and archival of Windows event logs. When an eventlog exceeds 
	75% of the configured maximum size the log will be backed up, compressed, moved to the configured archive location
	and the log will be cleared. If no location is specified the script will default to the C:\ drive. It is recommended
	to set the archive path to another drive to move the logs from the default system drive. 
	
	In order to run continuously the script will created a scheduled task on the commputer to run every 30 minutes to
	to check the current status of event logs.   

	Status of the script will be written to the Application log [EventLog_LogMaintenance]. 
	
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
	
.PARAMETER Install
	Create a scheduled task to run the script every 30min and save parameters to the registry.
	
.PARAMETER Remove
	Remove the scheduled task and remove saved parameters.
	
.PARAMETER Dry
    When present, do not remove any logs file.
	
.PARAMETER Debug
	Show debug informations.

.EXAMPLE
	EventLog-Archive.ps1 -DestinationPath D:\EventLog_Archive
	This example is the script running for the first time. The Eventlog archive path will be set as "D:\Eventlog_Archive"
	in the registry. 
	
#>
Param (
    # Local folder to store EventLog Data Collection 
    [parameter(Position=0, Mandatory=$False)][String]$DestinationPath = "C:\EventLogArchive",
	[parameter(Position=1, Mandatory=$False)][string]$EventLogSourcePath = "$($Env:SystemRoot)\System32\Winevt\Logs",
	[parameter(Position=2, Mandatory=$False)][int]$ArchiveRetentionDays = 182,
	[parameter(Position=3, Mandatory=$False)][Switch]$Install = $False,
	[parameter(Position=4, Mandatory=$False)][Switch]$Remove = $False,
	[parameter(Position=5, Mandatory=$False)][Switch]$Dry = $False
)


### INIT ###
if ($PSBoundParameters["Debug"]) { $DebugPreference = "Continue" }

# Add Zip assembly
Add-Type -assembly "System.IO.Compression.FileSystem"

# Global Variables
[int]$ScriptVer = 1802
[string]$MachineName = ((Get-WmiObject "Win32_ComputerSystem").Name)
[string]$EventSource = "EventLog_LogMaintenance"
[string]$TempDestinationPath = $($DestinationPath + "\_temp")
[string]$HKLMLogMaintenancePath = "HKLM:\Software\EventLogScripts\LogMaintenance"
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

### CHECK PARAMETERS ###
if (!(Test-Path $EventLogSourcePath))
{
	Write-Error "[SETUP]The EventLog Source path $EventLogSourcePath doesn't exists !"
	Exit
}

### SETUP ###
Switch (Test-Path $HKLMLogMaintenancePath) 
{
	$False {
		# Exit if remove switch and the script is not installed !
		if ($PSBoundParameters.ContainsKey("Remove")) 
		{ 
			Write-Error "[SETUP]The script is not installed, nothing to remove !"
			Exit 
		}

		# If no install switch, don't install it and exit if Destination path doesn't exists !
		if (!$PSBoundParameters.ContainsKey("Install")) 
		{ 
			if (!((Test-Path $DestinationPath) -and (Test-Path "$DestinationPath\_temp")))
			{
				Write-Error "[SETUP]The EventLog Destination path $DestinationPath doesn't exists !"
				Exit
			}
			
			Break 
		}
		
		# Create script registry entries
		Write-Debug "[SETUP]Creating Event Log archive registry entries: $HKLMLogMaintenancePath"
		New-Item -Path $HKLMLogMaintenancePath -Force | Out-Null
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name ScriptVersion -Value $ScriptVer | Out-Null
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name EventLogSourcePath -Value $EventLogSourcePath | Out-Null
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name DestinationPath -Value $DestinationPath | Out-Null
		New-ItemProperty -Path $HKLMLogMaintenancePath -Name ArchiveRetentionDays -Value $ArchiveRetentionDays | Out-Null
		
		# Add Event Log entry for logging script actions
		eventcreate /ID 775 /L Application /T Information /SO $EventSource /D "Log Maintenance Script installation started" | Out-Null

		# Create event log archive directory
		Write-Debug "[SETUP]Creating Event Log archive directory: $DestinationPath"
		try 
		{
			New-Item $DestinationPath -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $TempDestinationPath -type directory -ErrorAction Stop -Force | Out-Null
			New-Item $($DestinationPath + "\_script") -type directory -ErrorAction Stop -Force | Out-Null
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
		$scriptCopyDestination = $($DestinationPath + "\_script\" + $($CurrentScript.Split('\'))[-1])
		Write-Debug "[SETUP]Copying script to: $scriptCopyDestination"
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
		$XMLExportPath = ($DestinationPath + "\_script\" + $MachineName + "-LogArchive.xml")
		$ScheduledTaskXML.Save($XMLExportPath)

		# Create scheduled task
		Write-Debug "[SETUP]Creating Scheduled Task for Event Log Archive"
		schtasks /create /tn "Event Log Archive" /xml $XMLExportPath
		Start-Sleep -Seconds 10
		
		# First run scheduled task
		Write-Debug "[SETUP]Running task for first time"
		Schtasks /Run /tn "Event Log Archive"
	}
	$True {	
		# Exit if install switch and the script is already installed !
		if ($PSBoundParameters.ContainsKey("Install")) 
		{ 
			Write-Error "[SETUP]The script is already installed !"
			Exit 
		} 	
	
		# Remove script
		if ($Remove)
		{
			# Remove script registry entries
			Write-Debug "[SETUP]Removing Registry entries for Event Log Archive"
			Remove-Item -Path $HKLMLogMaintenancePath -Force
				
			# Create scheduled task
			Write-Debug "[SETUP]Removing Scheduled Task for Event Log Archive"
			schtasks /Delete /tn "Event Log Archive" /F
			Start-Sleep -Seconds 10
			Exit
		}	
	
		# Read configuration
		Write-Debug "Reading saved configuration"
		
		# Source path
		$readEventLogSourcePath = $true
		if ($PSBoundParameters.ContainsKey("EventLogSourcePath"))
		{
			if (Test-Path $EventLogSourcePath)
			{
				Write-Debug "Forcing EventLog Source path: $EventLogSourcePath"
				$readEventLogSourcePath = $false
			}
			else
			{
				Write-Warning "Can't force EventLog Source path : $EventLogSourcePath doesn't exists !"
			}
		}
		
		if ($readEventLogSourcePath)
		{
			$EventLogSourcePath = ((Get-ItemProperty -Path $HKLMLogMaintenancePath).EventLogSourcePath)
			#Write-Debug "EventLog Source path: $EventLogSourcePath"
		}
		
		# Destination path
		$readDestinationPath = $true
		if ($PSBoundParameters.ContainsKey("DestinationPath"))
		{
			if ((Test-Path $DestinationPath) -and (Test-Path "$DestinationPath\_temp"))
			{
				Write-Debug "Forcing EventLog Destination path: $DestinationPath"
				$readDestinationPath = $false
			}
			else
			{
				Write-Warning "Can't force EventLog Destination path : $DestinationPath doesn't exists !"
			}
			
		}
		
		if ($readDestinationPath)
		{
			$DestinationPath = ((Get-ItemProperty -Path $HKLMLogMaintenancePath).DestinationPath)
			#Write-Debug "EventLog Destination path: $DestinationPath"
		}
		
		[string]$TempDestinationPath = $($DestinationPath + "\_temp")
		#Write-Debug "EventLog Archive temporary path: $TempDestinationPath"
		
		# Archive retention days
		if ($PSBoundParameters.ContainsKey("ArchiveRetentionDays"))
		{
			Write-Debug "Forcing EventLog Archive retention days: $(&{If ($ArchiveRetentionDays -eq -1) { "Infinite" } Else { $ArchiveRetentionDays }})"
		}
		else
		{
			$ArchiveRetentionDays = ((Get-ItemProperty -Path $HKLMLogMaintenancePath).ArchiveRetentionDays)
			#Write-Debug "EventLog Archive retention days: $(&{If ($ArchiveRetentionDays -eq -1) { "Infinite" } Else { $ArchiveRetentionDays }})"
		}
		
        $message = "Starting EventLog Archive Tool"
		Write-EventLog -LogName Application -Source $EventSource -EventId 776 -Message $message -Category 1 -EntryType Information
	}
} 

# Show parameters
Write-Debug "EventLog Source path: $EventLogSourcePath"
Write-Debug "EventLog Destination path: $DestinationPath"
Write-Debug "EventLog Archive temporary path: $TempDestinationPath"
Write-Debug "EventLog Archive retention days: $(&{If ($ArchiveRetentionDays -eq -1) { "Infinite" } Else { $ArchiveRetentionDays }})"


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

$message = "Stopping EventLog Archive Tool"
Write-EventLog -LogName Application -Source $EventSource -EventId 779 -Message $message -Category 1 -EntryType Information
