# Backup-Reaper

PowerShell script to synchronize Reaper project and data directories to one or more backup locations using Robocopy, with optional pre-flight checks, scheduler-friendly quiet mode, and configurable tuning.

## Quick Start

1. Open PowerShell in this repository folder.
2. Create a config file or use command-line parameters:

```powershell
.\Sync-ReaperBackups.ps1 -SourcePaths @("C:\Reaper\Projects", "C:\Reaper\Data") `
    -BackupRoots @("E:\Backups\Reaper")
```

3. Verify output:
   - Files appear in the backup root directory
   - A log file is written (if `-LogPath` is specified)
   - A Windows notification appears on completion (unless `-SchedulerFriendly` is set)

## Run Daily With Task Scheduler

Create a daily task at 2:00 AM using a config file:

```powershell
$scriptPath = "C:\utils\Backup-Reaper\Sync-ReaperBackups.ps1"
$configPath = "C:\utils\Backup-Reaper\backup.config.json"

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ConfigPath `"$configPath`""
$trigger = New-ScheduledTaskTrigger -Daily -At 2:00AM

Register-ScheduledTask -TaskName "Reaper Backup Sync" -Action $action -Trigger $trigger `
    -Description "Daily Reaper project backup synchronization"
```

To remove the task later:

```powershell
Unregister-ScheduledTask -TaskName "Reaper Backup Sync" -Confirm:$false
```

## Script

- `Sync-ReaperBackups.ps1` – Main backup synchronization script
- `Edit-BackupConfigGui.ps1` – Windows GUI editor for backup config files
- `backup.config.example.json` – Example configuration file

## What It Does

- Validates that all source directories exist and are accessible
- Performs pre-flight checks on destination volumes (reachability and free space)
- Synchronizes each source directory to each backup root using Robocopy
- By default, creates an exact mirror of the source (deletes files that no longer exist in source)
- Optionally preserves files that exist only in the backup
- Writes timestamped logs with optional severity levels
- Sends Windows notifications on completion (success, partial failure, or fatal error)
- Returns clear exit codes for scripting and task scheduler monitoring

## Pre-Flight Checks

During startup, the script:

- **Source validation**: Confirms all source paths are accessible directories
- **Destination validation**: Checks that backup roots are reachable (local drives, UNC paths, etc.)
- **Free space check**: Estimates total source size and verifies local destination volumes have sufficient space
- **Logging validation**: Ensures the log directory can be created if needed

Optional: Skip the expensive source size scan with `-SkipPreflightSizeScan` for faster startup on large libraries.

## Robocopy Tuning

The script offers configurable multithreading levels:

- **Normal** (default): `/MT:8` – Conservative, suitable for most systems
- **High**: `/MT:16` – Increased parallelism for systems with spare I/O capacity
- **Higher**: `/MT:32` – Maximum parallelism for dedicated or high-bandwidth targets

Set via `-RobocopyTuning` parameter or in the config file. Use `Higher` for backup drives with dedicated bandwidth; use `Normal` for backup over network or to avoid source disk thrashing.

## Reliability Features

- **Path validation**: Resolves all paths before any copy operations begin
- **Directory creation**: Creates destination directories as needed (unless in dry-run mode)
- **Atomic writes**: Robocopy handles file copying with retries and graceful failure handling
- **Comprehensive logging**: Per-job, per-failure, and final outcome logged with timestamps
- **Scheduler-safe**: Fully suppresses console output and notifications when `-SchedulerFriendly` is set
- **Exit codes**: Returns 0 for success, 1 for fatal errors, 2 for partial failures

## Requirements

- Windows
- PowerShell 5.1+ (PowerShell 7 also supported)
- Robocopy (included with Windows)
- Permission to read source and write destination/log directories

Optional:

- BurntToast module for richer toast notifications

Install BurntToast:

```powershell
Install-Module BurntToast -Scope CurrentUser
```

## Usage

### Basic Command Line

```powershell
.\Sync-ReaperBackups.ps1 -SourcePaths @("C:\Reaper\Projects") `
    -BackupRoots @("E:\Backups\Reaper")
```

### Config File (Recommended)

```powershell
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json
```

Edit `backup.config.json` to set all parameters once:

```json
{
  "SourcePaths": [
    "C:\\Reaper\\Projects",
    "C:\\Reaper\\Data"
  ],
  "BackupRoots": [
    "E:\\Backups\\Reaper",
    "F:\\ReaperArchive"
  ],
  "PreserveTargetOnlyFiles": false,
  "DryRun": false,
  "SchedulerFriendly": false,
  "SkipPreflightSizeScan": false,
  "RobocopyTuning": "Normal",
  "LogPath": "C:\\Logs\\Backup-Reaper\\backup.log",
  "TaskIdentifier": null
}
```

### GUI Config Editor

Use the built-in Windows GUI to create or update your JSON config file:

```powershell
.\Edit-BackupConfigGui.ps1
```

Or open an existing config directly:

```powershell
.\Edit-BackupConfigGui.ps1 -ConfigPath .\backup.config.json
```

In the editor:

- Use `Open Existing...` to select an existing JSON config file
- Use `New Config...` to start from defaults and choose where to create a new config file
- On startup (without `-ConfigPath`), choose whether to open existing or start new
- Enter one source path per line in `SourcePaths`
- Enter one destination root per line in `BackupRoots`
- Use `Add Folder...` buttons for Windows folder-picking in both path lists
- Toggle switches for `DryRun`, `SchedulerFriendly`, and related options
- Choose `RobocopyTuning` from the dropdown (`Normal`, `High`, `Higher`)
- Set `LogPath`, then save to `backup.config.json`
- Manage an associated scheduled task from the `Scheduled Task` section:
  - `Create / Update Task` to register a daily task for the selected config
  - The editor stores a stable `TaskIdentifier` in the config and uses it to resolve the matching scheduled task before falling back to the task name
  - Review all configured daily trigger times in the list, use `Add Time` to add from the picker, and `Remove Selected Time` to delete selected entries
  - `Force -SchedulerFriendly for scheduled runs` to keep task runs quiet (recommended)
  - `Task Status` to view current state, run times, and last result
  - `Run Now` to start the task immediately
  - `Remove Task` to delete the scheduled task

### Multiple Backups

Specify multiple destination roots to sync the same source to multiple locations:

```powershell
.\Sync-ReaperBackups.ps1 -SourcePaths @("C:\Reaper\Projects", "C:\Reaper\Data") `
    -BackupRoots @("E:\Backups\Reaper", "F:\ReaperArchive", "\\NAS\BackupShare")
```

### Preserve Non-Source Files

Keep files in the backup that don't exist in the source (instead of deleting them):

```powershell
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -PreserveTargetOnlyFiles
```

### Dry Run

Preview what would be copied without making changes:

```powershell
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -DryRun
```

Or use the PowerShell `-WhatIf` parameter:

```powershell
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -WhatIf
```

### Scheduler-Friendly Mode

Suppress all console output and notifications (for cron/Task Scheduler runs):

```powershell
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -SchedulerFriendly
```

### High-Performance Tuning

Increase Robocopy multithreading for large libraries on fast drives:

```powershell
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -RobocopyTuning Higher
```

### Skip Expensive Pre-Flight Scan

When you trust your destination capacity, skip the source size calculation:

```powershell
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -SkipPreflightSizeScan
```

### Show Help

```powershell
Get-Help .\Sync-ReaperBackups.ps1 -Full
```

## Exit Behavior

- **Exit 0**: All backups completed successfully
- **Exit 1**: Fatal error (invalid paths, inaccessible sources/destinations, or insufficient free space)
- **Exit 2**: Partial failure (one or more backup jobs failed, others succeeded); all completed jobs are logged

When `-SchedulerFriendly` is set, all output and error messages are suppressed from the console but are still written to the log file.

## Notes

- Running as Administrator may improve consistency when source directories contain open files, as some metadata is preserved more reliably
- Pre-flight checks are always run unless `-SkipPreflightSizeScan` is set
- The script uses Robocopy's `/MIR` (mirror) by default, which deletes destination files not in the source; use `-PreserveTargetOnlyFiles` to change this behavior
- Log files grow over time; consider rotating them externally or using a log archive tool
- For very large Reaper libraries (100+ GB), consider using `-RobocopyTuning Higher` and `-SkipPreflightSizeScan` for faster execution

## Additional Links

- [Code](https://github.com/yourusername/Backup-Reaper)
- [Issues](https://github.com/yourusername/Backup-Reaper/issues)
- [Pull requests](https://github.com/yourusername/Backup-Reaper/pulls)
