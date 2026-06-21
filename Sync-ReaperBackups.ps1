<#
.SYNOPSIS
Synchronizes Reaper project and data directories to one or more backup locations.

.DESCRIPTION
Copies each source directory to each backup root using Robocopy.

By default, the destination is an exact mirror of the source directory. Files and
folders that no longer exist in the source are removed from the backup.

Use -PreserveTargetOnlyFiles to keep files and folders that exist only on the
destination while still updating changed and new content from the source.

.EXAMPLE
.\Sync-ReaperBackups.ps1 -SourcePaths @("C:\Reaper\Projects", "C:\Reaper\Data") `
    -BackupRoots @("E:\Backups\Reaper", "F:\ReaperArchive")

.EXAMPLE
.\Sync-ReaperBackups.ps1 -SourcePaths @("C:\Reaper\Projects", "C:\Reaper\Data") `
    -BackupRoots @("E:\Backups\Reaper") -PreserveTargetOnlyFiles

.EXAMPLE
.\Sync-ReaperBackups.ps1 -SourcePaths @("C:\Reaper\Projects", "C:\Reaper\Data") `
    -BackupRoots @("E:\Backups\Reaper") -WhatIf

.EXAMPLE
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json

.EXAMPLE
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -DryRun

.EXAMPLE
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -SchedulerFriendly

.EXAMPLE
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -RobocopyTuning High

.EXAMPLE
.\Sync-ReaperBackups.ps1 -ConfigPath .\backup.config.json -SkipPreflightSizeScan
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateNotNullOrEmpty()]
    [string[]]$SourcePaths,

    [ValidateNotNullOrEmpty()]
    [string[]]$BackupRoots,

    [switch]$PreserveTargetOnlyFiles,

    [switch]$DryRun,

    [switch]$SchedulerFriendly,

    [switch]$SkipPreflightSizeScan,

    [ValidateSet('Normal', 'High', 'Higher')]
    [string]$RobocopyTuning = 'Normal',

    [string]$LogPath,

    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

function Resolve-ConfigPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
}

function Read-BackupConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config file '$Path' was not found."
    }

    $config = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if (-not $config) {
        throw "Config file '$Path' is empty or invalid."
    }

    return $config
}

if ($ConfigPath) {
    $ConfigPath = Resolve-ConfigPath -Path $ConfigPath
    $config = Read-BackupConfiguration -Path $ConfigPath

    if (-not $SourcePaths -and $config.SourcePaths) {
        $SourcePaths = @($config.SourcePaths)
    }

    if (-not $BackupRoots -and $config.BackupRoots) {
        $BackupRoots = @($config.BackupRoots)
    }

    if (-not $PSBoundParameters.ContainsKey('PreserveTargetOnlyFiles') -and $null -ne $config.PreserveTargetOnlyFiles) {
        $PreserveTargetOnlyFiles = [bool]$config.PreserveTargetOnlyFiles
    }

    if (-not $PSBoundParameters.ContainsKey('DryRun') -and $null -ne $config.DryRun) {
        $DryRun = [bool]$config.DryRun
    }

    if (-not $PSBoundParameters.ContainsKey('SchedulerFriendly') -and $null -ne $config.SchedulerFriendly) {
        $SchedulerFriendly = [bool]$config.SchedulerFriendly
    }

    if (-not $PSBoundParameters.ContainsKey('SkipPreflightSizeScan') -and $null -ne $config.SkipPreflightSizeScan) {
        $SkipPreflightSizeScan = [bool]$config.SkipPreflightSizeScan
    }

    if (-not $PSBoundParameters.ContainsKey('RobocopyTuning') -and $config.RobocopyTuning) {
        $RobocopyTuning = [string]$config.RobocopyTuning
    }

    if (-not $LogPath -and $config.LogPath) {
        $LogPath = [string]$config.LogPath
    }
}

if (-not $SourcePaths -or $SourcePaths.Count -eq 0) {
    throw 'SourcePaths must be specified either as a parameter or in the config file.'
}

if (-not $BackupRoots -or $BackupRoots.Count -eq 0) {
    throw 'BackupRoots must be specified either as a parameter or in the config file.'
}

$isDryRun = $DryRun -or $WhatIfPreference
$isSchedulerFriendly = $SchedulerFriendly
$skipPreflightSizeScan = $SkipPreflightSizeScan

function Resolve-ExistingDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    if (-not $item.PSIsContainer) {
        throw "Source path '$Path' is not a directory."
    }

    return $item.FullName
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($isDryRun) {
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -ItemType Directory -Path $Path -Force
    }
}

function Get-RobocopyThreadCount {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Normal', 'High', 'Higher')]
        [string]$Tuning
    )

    switch ($Tuning) {
        'Normal' { return 8 }
        'High' { return 16 }
        'Higher' { return 32 }
    }
}

function Get-DirectorySizeBytes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $size = 0L
    Get-ChildItem -LiteralPath $Path -Force -File -Recurse -ErrorAction Stop | ForEach-Object {
        $size += [int64]$_.Length
    }

    return $size
}

function Format-Size {
    param(
        [Parameter(Mandatory = $true)]
        [int64]$Bytes
    )

    $units = 'B', 'KB', 'MB', 'GB', 'TB', 'PB'
    $value = [double]$Bytes
    $unitIndex = 0

    while ($value -ge 1024 -and $unitIndex -lt ($units.Count - 1)) {
        $value /= 1024
        $unitIndex++
    }

    return ('{0:N2} {1}' -f $value, $units[$unitIndex])
}

function Test-BackupPreflight {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$SourcePaths,

        [Parameter(Mandatory = $true)]
        [string[]]$BackupRoots,

        [switch]$SkipSizeScan
    )

    $sourceSizes = @{}
    $totalSourceBytes = 0L

    if (-not $SkipSizeScan) {
        foreach ($sourcePath in $SourcePaths) {
            $sourceSize = Get-DirectorySizeBytes -Path $sourcePath
            $sourceSizes[$sourcePath] = $sourceSize
            $totalSourceBytes += $sourceSize
        }
    }

    foreach ($backupRoot in $BackupRoots) {
        $rootItem = Get-Item -LiteralPath $backupRoot -ErrorAction SilentlyContinue

        if ($rootItem -and -not $rootItem.PSIsContainer) {
            throw "Backup root '$backupRoot' exists but is not a directory."
        }

        $driveRoot = [System.IO.Path]::GetPathRoot($backupRoot)
        if (-not $driveRoot) {
            throw "Unable to determine the target volume for backup root '$backupRoot'."
        }

        if ($driveRoot -notmatch '^[A-Za-z]:\\$') {
            if (-not $isSchedulerFriendly) {
                Write-Warning "Skipping free-space preflight for non-local path '$backupRoot'."
            }
            continue
        }

        $drive = [System.IO.DriveInfo]::new($driveRoot)

        if (-not $drive.IsReady) {
            throw "Backup root '$backupRoot' is not reachable."
        }

        $minimumRequiredBytes = $totalSourceBytes
        if ($minimumRequiredBytes -lt 1GB) {
            $minimumRequiredBytes = 1GB
        }

        if ([int64]$drive.AvailableFreeSpace -lt $minimumRequiredBytes) {
            throw "Backup root '$backupRoot' does not have enough free space. Required at least $(Format-Size -Bytes $minimumRequiredBytes), available $(Format-Size -Bytes ([int64]$drive.AvailableFreeSpace))."
        }
    }

    return [pscustomobject]@{
        TotalSourceBytes = $totalSourceBytes
        SizeScanSkipped = [bool]$SkipSizeScan
        SourceSizes = $sourceSizes
    }
}

function Send-WindowsNotification {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $IsWindows) {
        return
    }

    try {
        $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
        $escapedMessage = [System.Security.SecurityElement]::Escape($Message)

        if (Get-Module -ListAvailable -Name BurntToast) {
            Import-Module BurntToast -ErrorAction Stop
            New-BurntToastNotification -Text $Title, $Message | Out-Null
            return
        }

        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop | Out-Null
        $xml = @"
<toast>
  <visual>
    <binding template="ToastGeneric">
            <text>$escapedTitle</text>
            <text>$escapedMessage</text>
    </binding>
  </visual>
</toast>
"@

        $toastXml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $toastXml.LoadXml($xml)
        $toast = [Windows.UI.Notifications.ToastNotification]::new($toastXml)
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Backup-Reaper')
        $notifier.Show($toast)
    }
    catch {
        Write-Warning "Unable to display Windows notification: $($_.Exception.Message)"
    }
}

function Write-LogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if (-not $logWriter) {
        return
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    if ($Level -eq 'INFO') {
        $logWriter.WriteLine("[$timestamp] $Message")
    }
    else {
        $logWriter.WriteLine("[$timestamp] $Level $Message")
    }
}

function Invoke-RobocopySync {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [switch]$PreserveTargetOnlyFiles,

        [ValidateSet('Normal', 'High', 'Higher')]
        [string]$RobocopyTuning
    )

    Ensure-Directory -Path $DestinationPath

    $robocopyArgs = @(
        $SourcePath
        $DestinationPath
        '/COPY:DAT'
        '/DCOPY:DAT'
        '/FFT'
        '/R:2'
        '/W:5'
        '/XJ'
        '/NDL'
        '/NJH'
        '/NJS'
    )

    $threadCount = Get-RobocopyThreadCount -Tuning $RobocopyTuning
    $robocopyArgs += "/MT:$threadCount"

    if ($PreserveTargetOnlyFiles) {
        $robocopyArgs += '/E'
    }
    else {
        $robocopyArgs += '/MIR'
    }

    if ($isDryRun) {
        return 0
    }

    if ($PSCmdlet.ShouldProcess($DestinationPath, "Sync from $SourcePath")) {
        & robocopy @robocopyArgs | Out-Null
        $exitCode = $LASTEXITCODE

        if ($exitCode -ge 8) {
            throw "Robocopy failed while syncing '$SourcePath' to '$DestinationPath' (exit code $exitCode)."
        }

        return $exitCode
    }
}

$logWriter = $null
$script:BackupFailureMessage = $null
$script:ExitCode = 0
if ($LogPath) {
    $logDirectory = Split-Path -Path $LogPath -Parent
    if ($logDirectory) {
        Ensure-Directory -Path $logDirectory
    }

    $logWriter = [System.IO.StreamWriter]::new($LogPath, $true, [System.Text.UTF8Encoding]::new($false))
    $logWriter.AutoFlush = $true
}

$resolvedSources = @()
$jobFailureMessages = New-Object System.Collections.Generic.List[string]
$completedJobs = 0
$totalJobs = 0
$fatalErrorMessage = $null

try {
    $resolvedSources = foreach ($sourcePath in $SourcePaths) {
        Resolve-ExistingDirectory -Path $sourcePath
    }

    $preflight = Test-BackupPreflight -SourcePaths $resolvedSources -BackupRoots $BackupRoots -SkipSizeScan:$skipPreflightSizeScan

    if (-not $isSchedulerFriendly) {
        if ($preflight.SizeScanSkipped) {
            $modeLabel = if ($isDryRun) { 'Dry run pre-flight check passed' } else { 'Pre-flight check passed' }
            Write-LogEntry -Message "$modeLabel. Source size scan skipped by request."
            Write-Host "$modeLabel. Source size scan skipped by request."
        }
        else {
            $modeLabel = if ($isDryRun) { 'Dry run pre-flight check passed' } else { 'Pre-flight check passed' }
            Write-LogEntry -Message "$modeLabel. Estimated source data size: $(Format-Size -Bytes $preflight.TotalSourceBytes)"
            Write-Host "$modeLabel. Estimated source data size: $(Format-Size -Bytes $preflight.TotalSourceBytes)"
        }
    }

    $totalJobs = $resolvedSources.Count * $BackupRoots.Count

    if ($isDryRun) {
        if (-not $isSchedulerFriendly) {
            Write-LogEntry -Message 'Dry run enabled. No files or folders will be created, changed, or deleted.'
            Write-Host 'Dry run enabled. No files or folders will be created, changed, or deleted.'
        }
    }

    foreach ($backupRoot in $BackupRoots) {
        Ensure-Directory -Path $backupRoot

        foreach ($sourcePath in $resolvedSources) {
            $destinationPath = Join-Path -Path $backupRoot -ChildPath (Split-Path -Path $sourcePath -Leaf)
            $message = "Syncing '$sourcePath' to '$destinationPath'"
            $jobSucceeded = $false
            $progressPercent = if ($totalJobs -gt 0) { [math]::Round($completedJobs / $totalJobs * 100, 0) } else { 0 }

            Write-LogEntry -Message $message

            if (-not $isSchedulerFriendly) {
                Write-Progress -Id 0 -Activity 'Backing up Reaper directories' -Status $message -PercentComplete $progressPercent
                Write-Host $message
            }

            try {
                $null = Invoke-RobocopySync -SourcePath $sourcePath -DestinationPath $destinationPath -PreserveTargetOnlyFiles:$PreserveTargetOnlyFiles -RobocopyTuning $RobocopyTuning
                $jobSucceeded = $true
            }
            catch {
                $failureMessage = "Failed '$sourcePath' to '$destinationPath': $($_.Exception.Message)"
                $jobFailureMessages.Add($failureMessage)

                if (-not $isSchedulerFriendly) {
                    Write-Warning $failureMessage
                }

                Write-LogEntry -Message $failureMessage -Level 'ERROR'
            }

            $completedJobs++

            $finishedPercent = if ($totalJobs -gt 0) { [math]::Round($completedJobs / $totalJobs * 100, 0) } else { 100 }
            $progressStatus = if ($jobSucceeded) {
                "Completed '$sourcePath' -> '$destinationPath'"
            }
            else {
                "Completed with errors '$sourcePath' -> '$destinationPath'"
            }

            if (-not $isSchedulerFriendly) {
                Write-Progress -Id 0 -Activity 'Backing up Reaper directories' -Status $progressStatus -PercentComplete $finishedPercent
            }
        }
    }

    if ($jobFailureMessages.Count -gt 0) {
        $script:BackupFailureMessage = "Completed with $($jobFailureMessages.Count) failure(s) out of $totalJobs job(s). First error: $($jobFailureMessages[0])"
        $script:ExitCode = 2
    }
}
catch {
    $fatalErrorMessage = $_.Exception.Message
    $script:BackupFailureMessage = $fatalErrorMessage
    $script:ExitCode = 1

    if (-not $isSchedulerFriendly) {
        Write-Error $fatalErrorMessage
    }
}
finally {
    if (-not $isSchedulerFriendly) {
        Write-Progress -Id 0 -Activity 'Backing up Reaper directories' -Completed
    }

    if ($logWriter) {
        if ($fatalErrorMessage) {
            Write-LogEntry -Message $fatalErrorMessage -Level 'ERROR'
        }
        elseif ($jobFailureMessages.Count -gt 0) {
            Write-LogEntry -Message $script:BackupFailureMessage -Level 'ERROR'
        }
        elseif ($isDryRun) {
            Write-LogEntry -Message 'DRY RUN COMPLETE'
        }
        else {
            Write-LogEntry -Message 'COMPLETE'
        }

        $logWriter.Dispose()
    }

    if (-not $isSchedulerFriendly) {
        if ($fatalErrorMessage) {
            Send-WindowsNotification -Title 'Reaper backup failed' -Message $fatalErrorMessage
        }
        elseif ($jobFailureMessages.Count -gt 0) {
            Send-WindowsNotification -Title 'Reaper backup completed with errors' -Message $script:BackupFailureMessage
        }
        elseif ($isDryRun) {
            Send-WindowsNotification -Title 'Reaper backup dry run complete' -Message "Dry run completed for $($resolvedSources.Count) source folder(s) across $($BackupRoots.Count) destination root(s)."
        }
        else {
            Send-WindowsNotification -Title 'Reaper backup complete' -Message "Backed up $($resolvedSources.Count) source folder(s) to $($BackupRoots.Count) destination root(s)."
        }
    }
}

if ($fatalErrorMessage) {
    exit 1
}

if ($jobFailureMessages.Count -gt 0) {
    exit 2
}

exit $script:ExitCode