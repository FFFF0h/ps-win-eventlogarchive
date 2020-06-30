# PowersShell - Windows Event Log Archiver

## Background

I Received a request to archive all of the event logs on server, and maintain the archived logs on the server for up to six months. To meet these requirements the following script will create a schedule task that will run every 30 minutes. Each event log will be checked, and if the size exceeds 75% of the confgured maximum log file size the log will be exported and compressed to the configured archive directory. The script will query for any event logs that were automatically archived by the system in the defualt event log location (C:\Windows\System32\Winevt\logs) and will compress and archive the files to the configured archive location. After the archived event logs have exceeded the configured retention period it will be automatically removed the next time the task is run.

All of the actions of the script will be written to the application log on the local system. This will allow monitoring of the status of the script with SCOM or other monitoring tools.

### Moved from PSGallery

I originally published this script on PowerShell Gallery.

[Archive Windows Event Logs - w/ Logging](https://gallery.technet.microsoft.com/scriptcenter/Archive-Windows-Event-Logs-f2acb98a)

![PSG Comment](/static/PSG_comment.jpg)

## Description

 This script will be used to automate the collection and archival of Windows event logs. When an eventlog exceeds  75% of the configured maximum size the log will be backed up, compressed, moved to the configured archive location and the log will be cleared. If no location is specified the script will default to the C:\ drive. It is recommended to set the archive path to another drive to move the logs from the default system drive.  In order to run continuously the script will created a scheduled task on the commputer to run every 30 minutes to to check the current status of event logs.   
Status of the script will be written to the Application log [MSP_LogMaintenance].

### Event ID reference

|||
|---|---|
|775|Script setup operations|
|776|Script Start|
|777|EventLog|
|778|Remove expired EventLog|
|779|Script end|

## Example

This example is the script running for the first time. The Eventlog archive path will be set as "D:\Eventlog_Archive" in the registry.

```powershell
    EventLog_Archive.ps1 -EventLogArchivePath D:\EventLog_Archive
```
