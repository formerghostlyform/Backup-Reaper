<#
.SYNOPSIS
Opens a GUI editor for Backup-Reaper configuration files.

.DESCRIPTION
Provides a Windows Forms interface to load, edit, validate, and save
Backup-Reaper JSON config files used by Sync-ReaperBackups.ps1.

.EXAMPLE
.\Edit-BackupConfigGui.ps1

.EXAMPLE
.\Edit-BackupConfigGui.ps1 -ConfigPath .\backup.config.json
#>

[CmdletBinding()]
param(
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

if (-not $IsWindows) {
    throw 'This GUI is supported only on Windows.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function New-DefaultConfig {
    return [ordered]@{
        SourcePaths = @(
            'C:\Reaper\Projects'
            'C:\Reaper\Data'
        )
        BackupRoots = @(
            'E:\Backups\Reaper'
        )
        PreserveTargetOnlyFiles = $false
        DryRun = $false
        SchedulerFriendly = $false
        SkipPreflightSizeScan = $false
        PostValidation = $false
        RobocopyTuning = 'Normal'
        LogPath = 'C:\Logs\Backup-Reaper\backup.log'
        LogRetentionDays = 90
        TaskIdentifier = $null
    }
}

function ConvertTo-ConfigObject {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$InputObject
    )

    $result = New-DefaultConfig

    if ($null -ne $InputObject.SourcePaths) {
        $result.SourcePaths = @($InputObject.SourcePaths | ForEach-Object { [string]$_ })
    }

    if ($null -ne $InputObject.BackupRoots) {
        $result.BackupRoots = @($InputObject.BackupRoots | ForEach-Object { [string]$_ })
    }

    if ($null -ne $InputObject.PreserveTargetOnlyFiles) {
        $result.PreserveTargetOnlyFiles = [bool]$InputObject.PreserveTargetOnlyFiles
    }

    if ($null -ne $InputObject.DryRun) {
        $result.DryRun = [bool]$InputObject.DryRun
    }

    if ($null -ne $InputObject.SchedulerFriendly) {
        $result.SchedulerFriendly = [bool]$InputObject.SchedulerFriendly
    }

    if ($null -ne $InputObject.SkipPreflightSizeScan) {
        $result.SkipPreflightSizeScan = [bool]$InputObject.SkipPreflightSizeScan
    }

    if ($null -ne $InputObject.PostValidation) {
        $result.PostValidation = [bool]$InputObject.PostValidation
    }

    if ($null -ne $InputObject.RobocopyTuning -and [string]$InputObject.RobocopyTuning) {
        $result.RobocopyTuning = [string]$InputObject.RobocopyTuning
    }

    if ($null -ne $InputObject.LogPath -and [string]$InputObject.LogPath) {
        $result.LogPath = [string]$InputObject.LogPath
    }

    if ($null -ne $InputObject.LogRetentionDays) {
        $result.LogRetentionDays = [int]$InputObject.LogRetentionDays
    }

    if ($null -ne $InputObject.TaskIdentifier -and [string]$InputObject.TaskIdentifier) {
        $result.TaskIdentifier = [string]$InputObject.TaskIdentifier
    }

    return $result
}

function Update-SourcePathsWarning {
    $paths = Convert-MultilineToArray -Text $txtSourcePaths.Text
    $leafNames = @($paths | ForEach-Object { Split-Path -Path $_ -Leaf } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $duplicates = @($leafNames | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })

    if ($duplicates.Count -gt 0) {
        $lblSourcePathsWarning.Text = "Warning: duplicate folder name(s) '$($duplicates -join "', '")' — these sources will collide under each backup root."
        $lblSourcePathsWarning.Visible = $true
    }
    else {
        $lblSourcePathsWarning.Text = ''
        $lblSourcePathsWarning.Visible = $false
    }
}

function Convert-MultilineToArray {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    return @(
        $Text -split "`r?`n" |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Resolve-ConfigFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location) -ChildPath $PathValue))
}

function Get-SyncScriptPath {
    $scriptDirectory = Split-Path -Path $PSCommandPath -Parent
    return (Join-Path -Path $scriptDirectory -ChildPath 'Sync-ReaperBackups.ps1')
}

function Get-DailyTimeString {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$TimeValue
    )

    return $TimeValue.ToString('HH:mm')
}

function Test-DailyTimeString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TimeText
    )

    try {
        [void][datetime]::ParseExact(
            $TimeText,
            'HH:mm',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None
        )
        return $true
    }
    catch {
        return $false
    }
}

function Get-LastNonEmptyLine {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $lines = @(
        $Text -split "`r?`n" |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    if ($lines.Count -eq 0) {
        return $null
    }

    return $lines[-1]
}

function Select-FolderPath {
    param(
        [string]$Title,
        [string]$InitialPath
    )

    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = $Title
    $folderDialog.ShowNewFolderButton = $true

    if (-not [string]::IsNullOrWhiteSpace($InitialPath) -and (Test-Path -LiteralPath $InitialPath -PathType Container)) {
        $folderDialog.SelectedPath = $InitialPath
    }

    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderDialog.SelectedPath
    }

    return $null
}

function Add-PathToTextbox {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$TextBox,

        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.Text = $PathValue
    }
    else {
        $TextBox.Text = $TextBox.Text.TrimEnd() + [Environment]::NewLine + $PathValue
    }
}

function Read-ConfigFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    if (-not (Test-Path -LiteralPath $PathValue)) {
        throw "Config file '$PathValue' was not found."
    }

    $jsonText = Get-Content -LiteralPath $PathValue -Raw -ErrorAction Stop
    $parsed = $jsonText | ConvertFrom-Json -ErrorAction Stop
    return (ConvertTo-ConfigObject -InputObject $parsed)
}

function Validate-ConfigValues {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    if ($Config.SourcePaths.Count -eq 0) {
        throw 'At least one source path is required.'
    }

    if ($Config.BackupRoots.Count -eq 0) {
        throw 'At least one backup root is required.'
    }

    $validTuning = @('Normal', 'High', 'Higher')
    if ($validTuning -notcontains $Config.RobocopyTuning) {
        throw "RobocopyTuning must be one of: $($validTuning -join ', ')."
    }

    if ([string]::IsNullOrWhiteSpace($Config.LogPath)) {
        throw 'LogPath is required.'
    }
}

function ConvertTo-JsonConfig {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    $taskIdentifier = if ([string]::IsNullOrWhiteSpace([string]$Config.TaskIdentifier)) { $null } else { [string]$Config.TaskIdentifier }

    $ordered = [ordered]@{
        SourcePaths = $Config.SourcePaths
        BackupRoots = $Config.BackupRoots
        PreserveTargetOnlyFiles = [bool]$Config.PreserveTargetOnlyFiles
        DryRun = [bool]$Config.DryRun
        SchedulerFriendly = [bool]$Config.SchedulerFriendly
        SkipPreflightSizeScan = [bool]$Config.SkipPreflightSizeScan
        PostValidation = [bool]$Config.PostValidation
        RobocopyTuning = [string]$Config.RobocopyTuning
        LogPath = [string]$Config.LogPath
        LogRetentionDays = [int]$Config.LogRetentionDays
        TaskIdentifier = $taskIdentifier
    }

    return ($ordered | ConvertTo-Json -Depth 5)
}

function Set-FormConfigValues {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    $txtSourcePaths.Text = ($Config.SourcePaths -join [Environment]::NewLine)
    $txtBackupRoots.Text = ($Config.BackupRoots -join [Environment]::NewLine)
    $chkPreserveTargetOnlyFiles.Checked = [bool]$Config.PreserveTargetOnlyFiles
    $chkDryRun.Checked = [bool]$Config.DryRun
    $chkSchedulerFriendly.Checked = [bool]$Config.SchedulerFriendly
    $chkSkipPreflightSizeScan.Checked = [bool]$Config.SkipPreflightSizeScan
    $chkPostValidation.Checked = [bool]$Config.PostValidation
    $cmbRobocopyTuning.SelectedItem = [string]$Config.RobocopyTuning
    $txtLogPath.Text = [string]$Config.LogPath
    $nudLogRetentionDays.Value = [Math]::Max(0, [int]$Config.LogRetentionDays)
    $script:taskIdentifier = if ($null -ne $Config.TaskIdentifier -and -not [string]::IsNullOrWhiteSpace([string]$Config.TaskIdentifier)) { [string]$Config.TaskIdentifier } else { $null }
}

function Get-FormConfigValues {
    $config = [ordered]@{
        SourcePaths = (Convert-MultilineToArray -Text $txtSourcePaths.Text)
        BackupRoots = (Convert-MultilineToArray -Text $txtBackupRoots.Text)
        PreserveTargetOnlyFiles = $chkPreserveTargetOnlyFiles.Checked
        DryRun = $chkDryRun.Checked
        SchedulerFriendly = $chkSchedulerFriendly.Checked
        SkipPreflightSizeScan = $chkSkipPreflightSizeScan.Checked
        PostValidation = $chkPostValidation.Checked
        RobocopyTuning = [string]$cmbRobocopyTuning.SelectedItem
        LogPath = $txtLogPath.Text.Trim()
        LogRetentionDays = [int]$nudLogRetentionDays.Value
        TaskIdentifier = if ([string]::IsNullOrWhiteSpace($script:taskIdentifier)) { $null } else { [string]$script:taskIdentifier }
    }

    Validate-ConfigValues -Config $config
    return $config
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Backup-Reaper Config Editor'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(860, 886)
$form.MinimumSize = New-Object System.Drawing.Size(760, 826)

$lblCurrentFile = New-Object System.Windows.Forms.Label
$lblCurrentFile.AutoSize = $true
$lblCurrentFile.Location = New-Object System.Drawing.Point(16, 15)
$lblCurrentFile.Text = 'Config file: (unsaved/new)'
$form.Controls.Add($lblCurrentFile)

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.AutoSize = $false
$lblMode.Location = New-Object System.Drawing.Point(568, 30)
$lblMode.Size = New-Object System.Drawing.Size(260, 20)
$lblMode.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblMode.Anchor = 'Top,Right'
$lblMode.ForeColor = [System.Drawing.Color]::FromArgb(25, 90, 40)
$lblMode.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblMode.Text = 'Mode: New unsaved config'
$form.Controls.Add($lblMode)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = 'Open Existing...'
$btnLoad.Size = New-Object System.Drawing.Size(120, 30)
$btnLoad.Location = New-Object System.Drawing.Point(16, 42)
$form.Controls.Add($btnLoad)

$btnNewConfig = New-Object System.Windows.Forms.Button
$btnNewConfig.Text = 'New Config...'
$btnNewConfig.Size = New-Object System.Drawing.Size(110, 30)
$btnNewConfig.Location = New-Object System.Drawing.Point(142, 42)
$form.Controls.Add($btnNewConfig)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Save'
$btnSave.Size = New-Object System.Drawing.Size(95, 30)
$btnSave.Location = New-Object System.Drawing.Point(258, 42)
$form.Controls.Add($btnSave)

$btnSaveAs = New-Object System.Windows.Forms.Button
$btnSaveAs.Text = 'Save As...'
$btnSaveAs.Size = New-Object System.Drawing.Size(95, 30)
$btnSaveAs.Location = New-Object System.Drawing.Point(359, 42)
$form.Controls.Add($btnSaveAs)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = 'Reset Defaults'
$btnReset.Size = New-Object System.Drawing.Size(120, 30)
$btnReset.Location = New-Object System.Drawing.Point(460, 42)
$form.Controls.Add($btnReset)

$groupPaths = New-Object System.Windows.Forms.GroupBox
$groupPaths.Text = 'Paths'
$groupPaths.Location = New-Object System.Drawing.Point(16, 86)
$groupPaths.Size = New-Object System.Drawing.Size(812, 330)
$groupPaths.Anchor = 'Top,Left,Right'
$form.Controls.Add($groupPaths)

$lblSourcePaths = New-Object System.Windows.Forms.Label
$lblSourcePaths.AutoSize = $true
$lblSourcePaths.Location = New-Object System.Drawing.Point(16, 28)
$lblSourcePaths.Text = 'SourcePaths (one path per line)'
$groupPaths.Controls.Add($lblSourcePaths)

$btnAddSourcePath = New-Object System.Windows.Forms.Button
$btnAddSourcePath.Text = 'Add Folder...'
$btnAddSourcePath.Size = New-Object System.Drawing.Size(105, 24)
$btnAddSourcePath.Location = New-Object System.Drawing.Point(284, 24)
$btnAddSourcePath.Anchor = 'Top,Left'
$groupPaths.Controls.Add($btnAddSourcePath)

$txtSourcePaths = New-Object System.Windows.Forms.TextBox
$txtSourcePaths.Multiline = $true
$txtSourcePaths.ScrollBars = 'Vertical'
$txtSourcePaths.Location = New-Object System.Drawing.Point(19, 50)
$txtSourcePaths.Size = New-Object System.Drawing.Size(370, 228)
$txtSourcePaths.Anchor = 'Top,Left,Bottom'
$groupPaths.Controls.Add($txtSourcePaths)

$lblSourcePathsWarning = New-Object System.Windows.Forms.Label
$lblSourcePathsWarning.AutoSize = $false
$lblSourcePathsWarning.Size = New-Object System.Drawing.Size(370, 36)
$lblSourcePathsWarning.Location = New-Object System.Drawing.Point(19, 283)
$lblSourcePathsWarning.Anchor = 'Bottom,Left'
$lblSourcePathsWarning.ForeColor = [System.Drawing.Color]::FromArgb(160, 80, 0)
$lblSourcePathsWarning.Font = New-Object System.Drawing.Font('Segoe UI', 8.25)
$lblSourcePathsWarning.Text = ''
$lblSourcePathsWarning.Visible = $false
$groupPaths.Controls.Add($lblSourcePathsWarning)

$lblBackupRoots = New-Object System.Windows.Forms.Label
$lblBackupRoots.AutoSize = $true
$lblBackupRoots.Location = New-Object System.Drawing.Point(418, 28)
$lblBackupRoots.Text = 'BackupRoots (one path per line)'
$groupPaths.Controls.Add($lblBackupRoots)

$btnAddBackupRoot = New-Object System.Windows.Forms.Button
$btnAddBackupRoot.Text = 'Add Folder...'
$btnAddBackupRoot.Size = New-Object System.Drawing.Size(105, 24)
$btnAddBackupRoot.Location = New-Object System.Drawing.Point(686, 24)
$btnAddBackupRoot.Anchor = 'Top,Right'
$groupPaths.Controls.Add($btnAddBackupRoot)

$txtBackupRoots = New-Object System.Windows.Forms.TextBox
$txtBackupRoots.Multiline = $true
$txtBackupRoots.ScrollBars = 'Vertical'
$txtBackupRoots.Location = New-Object System.Drawing.Point(421, 50)
$txtBackupRoots.Size = New-Object System.Drawing.Size(370, 255)
$txtBackupRoots.Anchor = 'Top,Left,Bottom,Right'
$groupPaths.Controls.Add($txtBackupRoots)

$groupOptions = New-Object System.Windows.Forms.GroupBox
$groupOptions.Text = 'Options'
$groupOptions.Location = New-Object System.Drawing.Point(16, 424)
$groupOptions.Size = New-Object System.Drawing.Size(812, 160)
$groupOptions.Anchor = 'Top,Left,Right'
$form.Controls.Add($groupOptions)

$chkPreserveTargetOnlyFiles = New-Object System.Windows.Forms.CheckBox
$chkPreserveTargetOnlyFiles.Text = 'PreserveTargetOnlyFiles'
$chkPreserveTargetOnlyFiles.Location = New-Object System.Drawing.Point(19, 28)
$chkPreserveTargetOnlyFiles.AutoSize = $true
$groupOptions.Controls.Add($chkPreserveTargetOnlyFiles)

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = 'DryRun'
$chkDryRun.Location = New-Object System.Drawing.Point(260, 28)
$chkDryRun.AutoSize = $true
$groupOptions.Controls.Add($chkDryRun)

$chkSchedulerFriendly = New-Object System.Windows.Forms.CheckBox
$chkSchedulerFriendly.Text = 'SchedulerFriendly'
$chkSchedulerFriendly.Location = New-Object System.Drawing.Point(355, 28)
$chkSchedulerFriendly.AutoSize = $true
$groupOptions.Controls.Add($chkSchedulerFriendly)

$chkSkipPreflightSizeScan = New-Object System.Windows.Forms.CheckBox
$chkSkipPreflightSizeScan.Text = 'SkipPreflightSizeScan'
$chkSkipPreflightSizeScan.Location = New-Object System.Drawing.Point(523, 28)
$chkSkipPreflightSizeScan.AutoSize = $true
$groupOptions.Controls.Add($chkSkipPreflightSizeScan)

$chkPostValidation = New-Object System.Windows.Forms.CheckBox
$chkPostValidation.Text = 'PostValidation'
$chkPostValidation.Location = New-Object System.Drawing.Point(19, 55)
$chkPostValidation.AutoSize = $true
$groupOptions.Controls.Add($chkPostValidation)

$lblTuning = New-Object System.Windows.Forms.Label
$lblTuning.Text = 'RobocopyTuning'
$lblTuning.AutoSize = $true
$lblTuning.Location = New-Object System.Drawing.Point(19, 94)
$groupOptions.Controls.Add($lblTuning)

$cmbRobocopyTuning = New-Object System.Windows.Forms.ComboBox
$cmbRobocopyTuning.DropDownStyle = 'DropDownList'
$cmbRobocopyTuning.Location = New-Object System.Drawing.Point(19, 115)
$cmbRobocopyTuning.Size = New-Object System.Drawing.Size(160, 25)
$cmbRobocopyTuning.Items.AddRange(@('Normal', 'High', 'Higher'))
$groupOptions.Controls.Add($cmbRobocopyTuning)

$lblLogRetentionDays = New-Object System.Windows.Forms.Label
$lblLogRetentionDays.Text = 'Retention (days)'
$lblLogRetentionDays.AutoSize = $true
$lblLogRetentionDays.Location = New-Object System.Drawing.Point(200, 94)
$groupOptions.Controls.Add($lblLogRetentionDays)

$nudLogRetentionDays = New-Object System.Windows.Forms.NumericUpDown
$nudLogRetentionDays.Location = New-Object System.Drawing.Point(200, 115)
$nudLogRetentionDays.Size = New-Object System.Drawing.Size(90, 25)
$nudLogRetentionDays.Minimum = 0
$nudLogRetentionDays.Maximum = 3650
$nudLogRetentionDays.Value = 90
$groupOptions.Controls.Add($nudLogRetentionDays)

$lblLogPath = New-Object System.Windows.Forms.Label
$lblLogPath.Text = 'LogPath'
$lblLogPath.AutoSize = $true
$lblLogPath.Location = New-Object System.Drawing.Point(308, 94)
$groupOptions.Controls.Add($lblLogPath)

$txtLogPath = New-Object System.Windows.Forms.TextBox
$txtLogPath.Location = New-Object System.Drawing.Point(308, 115)
$txtLogPath.Size = New-Object System.Drawing.Size(394, 25)
$txtLogPath.Anchor = 'Top,Left,Right'
$groupOptions.Controls.Add($txtLogPath)

$btnBrowseLog = New-Object System.Windows.Forms.Button
$btnBrowseLog.Text = 'Browse...'
$btnBrowseLog.Size = New-Object System.Drawing.Size(88, 27)
$btnBrowseLog.Location = New-Object System.Drawing.Point(708, 114)
$btnBrowseLog.Anchor = 'Top,Right'
$groupOptions.Controls.Add($btnBrowseLog)

$groupTask = New-Object System.Windows.Forms.GroupBox
$groupTask.Text = 'Scheduled Task'
$groupTask.Location = New-Object System.Drawing.Point(16, 592)
$groupTask.Size = New-Object System.Drawing.Size(812, 240)
$groupTask.Anchor = 'Left,Right,Bottom'
$form.Controls.Add($groupTask)

$lblTaskName = New-Object System.Windows.Forms.Label
$lblTaskName.Text = 'Task name'
$lblTaskName.AutoSize = $true
$lblTaskName.Location = New-Object System.Drawing.Point(19, 24)
$groupTask.Controls.Add($lblTaskName)

$txtTaskName = New-Object System.Windows.Forms.TextBox
$txtTaskName.Location = New-Object System.Drawing.Point(19, 52)
$txtTaskName.Size = New-Object System.Drawing.Size(290, 25)
$txtTaskName.Text = 'Reaper Backup Sync'
$groupTask.Controls.Add($txtTaskName)

$lblTaskTime = New-Object System.Windows.Forms.Label
$lblTaskTime.Text = 'Configured daily times (HH:mm)'
$lblTaskTime.AutoSize = $true
$lblTaskTime.Location = New-Object System.Drawing.Point(329, 24)
$groupTask.Controls.Add($lblTaskTime)

$lstTaskTimes = New-Object System.Windows.Forms.ListBox
$lstTaskTimes.Location = New-Object System.Drawing.Point(329, 52)
$lstTaskTimes.Size = New-Object System.Drawing.Size(202, 100)
$lstTaskTimes.IntegralHeight = $false
$lstTaskTimes.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
[void]$lstTaskTimes.Items.Add('02:00')
$groupTask.Controls.Add($lstTaskTimes)

$dtTaskTime = New-Object System.Windows.Forms.DateTimePicker
$dtTaskTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$dtTaskTime.ShowUpDown = $true
$dtTaskTime.Width = 110
$dtTaskTime.Location = New-Object System.Drawing.Point(544, 52)
$dtTaskTime.Value = [datetime]::Today.AddHours(2)
$groupTask.Controls.Add($dtTaskTime)

$btnTaskAddTime = New-Object System.Windows.Forms.Button
$btnTaskAddTime.Text = 'Add Time'
$btnTaskAddTime.Size = New-Object System.Drawing.Size(85, 27)
$btnTaskAddTime.Location = New-Object System.Drawing.Point(660, 51)
$groupTask.Controls.Add($btnTaskAddTime)

$btnTaskRemoveTime = New-Object System.Windows.Forms.Button
$btnTaskRemoveTime.Text = 'Remove Selected Time'
$btnTaskRemoveTime.Size = New-Object System.Drawing.Size(145, 27)
$btnTaskRemoveTime.Location = New-Object System.Drawing.Point(544, 84)
$groupTask.Controls.Add($btnTaskRemoveTime)

$btnTaskCreateUpdate = New-Object System.Windows.Forms.Button
$btnTaskCreateUpdate.Text = 'Create / Update Task'
$btnTaskCreateUpdate.Size = New-Object System.Drawing.Size(145, 30)
$btnTaskCreateUpdate.Location = New-Object System.Drawing.Point(19, 175)
$groupTask.Controls.Add($btnTaskCreateUpdate)

$btnTaskStatus = New-Object System.Windows.Forms.Button
$btnTaskStatus.Text = 'Task Status'
$btnTaskStatus.Size = New-Object System.Drawing.Size(90, 30)
$btnTaskStatus.Location = New-Object System.Drawing.Point(169, 175)
$groupTask.Controls.Add($btnTaskStatus)

$btnTaskRunNow = New-Object System.Windows.Forms.Button
$btnTaskRunNow.Text = 'Run Now'
$btnTaskRunNow.Size = New-Object System.Drawing.Size(90, 30)
$btnTaskRunNow.Location = New-Object System.Drawing.Point(265, 175)
$groupTask.Controls.Add($btnTaskRunNow)

$btnTaskRemove = New-Object System.Windows.Forms.Button
$btnTaskRemove.Text = 'Remove Task'
$btnTaskRemove.Size = New-Object System.Drawing.Size(110, 28)
$btnTaskRemove.Location = New-Object System.Drawing.Point(19, 79)
$groupTask.Controls.Add($btnTaskRemove)

$chkTaskSchedulerFriendly = New-Object System.Windows.Forms.CheckBox
$chkTaskSchedulerFriendly.Text = 'Force -SchedulerFriendly for scheduled runs'
$chkTaskSchedulerFriendly.AutoSize = $false
$chkTaskSchedulerFriendly.Size = New-Object System.Drawing.Size(295, 32)
$chkTaskSchedulerFriendly.Checked = $true
$chkTaskSchedulerFriendly.Location = New-Object System.Drawing.Point(19, 112)
$groupTask.Controls.Add($chkTaskSchedulerFriendly)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(16, 842)
$lblStatus.Anchor = 'Left,Bottom'
$lblStatus.Text = 'Ready'
$form.Controls.Add($lblStatus)

$openDialog = New-Object System.Windows.Forms.OpenFileDialog
$openDialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
$openDialog.Title = 'Open backup config'

$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
$saveDialog.Title = 'Save backup config'
$saveDialog.DefaultExt = 'json'

$script:currentConfigPath = $null
$script:taskIdentifier = $null

function Update-Status {
    param([string]$Message)

    $lblStatus.Text = $Message
}

function Update-CurrentFileLabel {
    if ($script:currentConfigPath) {
        $lblCurrentFile.Text = "Config file: $script:currentConfigPath"
        $lblMode.Text = 'Mode: Editing existing file'
        $lblMode.ForeColor = [System.Drawing.Color]::FromArgb(30, 70, 140)
    }
    else {
        $lblCurrentFile.Text = 'Config file: (unsaved/new)'
        $lblMode.Text = 'Mode: New unsaved config'
        $lblMode.ForeColor = [System.Drawing.Color]::FromArgb(25, 90, 40)
    }
}

function Load-ConfigIntoForm {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    $config = Read-ConfigFile -PathValue $PathValue
    Set-FormConfigValues -Config $config

    try {
        $resolvedTask = Resolve-ScheduledTask -TaskName $txtTaskName.Text.Trim() -TaskIdentifier $script:taskIdentifier
        if ($resolvedTask -and $resolvedTask.TaskName -and $txtTaskName.Text.Trim() -ne $resolvedTask.TaskName) {
            $txtTaskName.Text = $resolvedTask.TaskName
        }
    }
    catch {
    }

    Sync-TaskTimesFromExistingTask -Silent
    $script:currentConfigPath = $PathValue
    Update-CurrentFileLabel
    Update-Status "Loaded $PathValue"
}

function Save-ConfigFromForm {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    $config = Get-FormConfigValues
    $json = ConvertTo-JsonConfig -Config $config
    Set-Content -LiteralPath $PathValue -Value $json -Encoding utf8 -ErrorAction Stop
    $script:currentConfigPath = $PathValue
    Update-CurrentFileLabel
    Update-Status "Saved $PathValue"
}

function Get-TaskNameValue {
    $taskName = $txtTaskName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($taskName)) {
        throw 'Task name cannot be empty.'
    }

    return $taskName
}

function Get-OrCreateTaskIdentifier {
    if ([string]::IsNullOrWhiteSpace($script:taskIdentifier)) {
        $script:taskIdentifier = ([guid]::NewGuid()).Guid
    }

    return $script:taskIdentifier
}

function Resolve-ScheduledTask {
    param(
        [string]$TaskName,
        [string]$TaskIdentifier
    )

    if (-not [string]::IsNullOrWhiteSpace($TaskIdentifier)) {
        $escapedIdentifier = [regex]::Escape($TaskIdentifier)
        $task = Get-ScheduledTask -ErrorAction Stop | Where-Object {
            $descriptionMatches = -not [string]::IsNullOrWhiteSpace($_.Description) -and $_.Description -match $escapedIdentifier
            $argumentMatches = @(
                $_.Actions |
                    ForEach-Object { [string]$_.Arguments }
            ) -match $escapedIdentifier

            $descriptionMatches -or $argumentMatches
        } | Select-Object -First 1

        if ($task) {
            return $task
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($TaskName)) {
        return (Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop)
    }

    throw 'Scheduled task was not found.'
}

function Test-TaskExists {
    param(
        [string]$TaskName,
        [string]$TaskIdentifier
    )

    if ([string]::IsNullOrWhiteSpace($TaskName) -and [string]::IsNullOrWhiteSpace($TaskIdentifier)) {
        return $false
    }

    try {
        $null = Resolve-ScheduledTask -TaskName $TaskName -TaskIdentifier $TaskIdentifier
        return $true
    }
    catch {
        return $false
    }
}

function Update-TaskActionButtons {
    $taskName = $txtTaskName.Text.Trim()
    $taskExists = Test-TaskExists -TaskName $taskName -TaskIdentifier $script:taskIdentifier
    $btnTaskStatus.Enabled = $taskExists
    $btnTaskRunNow.Enabled = $taskExists
    $btnTaskRemove.Enabled = $taskExists
}

function Set-TaskTimesList {
    param(
        [string[]]$TimeValues
    )

    $lstTaskTimes.Items.Clear()

    $normalized = @($TimeValues | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($normalized.Count -eq 0) {
        [void]$lstTaskTimes.Items.Add('02:00')
        return
    }

    foreach ($timeValue in $normalized) {
        [void]$lstTaskTimes.Items.Add($timeValue)
    }
}

function Get-TaskDailyTriggerTimes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [string]$TaskIdentifier
    )

    $task = Resolve-ScheduledTask -TaskName $TaskName -TaskIdentifier $TaskIdentifier

    return @(
        $task.Triggers |
            Where-Object { $_.CimClass.CimClassName -eq 'MSFT_TaskDailyTrigger' } |
            ForEach-Object {
                if ($_.StartBoundary) {
                    try {
                        ([datetime]$_.StartBoundary).ToString('HH:mm')
                    }
                    catch {
                        $null
                    }
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}

function Sync-TaskTimesFromExistingTask {
    param(
        [switch]$Silent
    )

    $taskName = $txtTaskName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($taskName)) {
        return
    }

    if (-not (Test-TaskExists -TaskName $taskName -TaskIdentifier $script:taskIdentifier)) {
        return
    }

    try {
        $taskTimes = Get-TaskDailyTriggerTimes -TaskName $taskName -TaskIdentifier $script:taskIdentifier
        if ($taskTimes.Count -gt 0) {
            Set-TaskTimesList -TimeValues $taskTimes
            if (-not $Silent) {
                Update-Status "Loaded trigger times from task '$taskName'."
            }
        }
        elseif (-not $Silent) {
            Update-Status "Task '$taskName' has no daily triggers to load."
        }
    }
    catch {
        if (-not $Silent) {
            Update-Status "Unable to read triggers from task '$taskName'."
        }
    }
}

function Ensure-ConfigPathForTask {
    if ($script:currentConfigPath -and (Test-Path -LiteralPath $script:currentConfigPath)) {
        return $script:currentConfigPath
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        'The config has not been saved to a file yet. Save it now so the task can reference it?',
        'Save config first',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        return $null
    }

    $saveDialog.FileName = 'backup.config.json'
    try {
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Save-ConfigFromForm -PathValue $saveDialog.FileName
            return $script:currentConfigPath
        }
    }
    finally {
        $saveDialog.FileName = ''
    }

    return $null
}

function New-TaskAction {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigFilePath,

        [string]$TaskIdentifier,

        [bool]$ForceSchedulerFriendly
    )

    $syncScriptPath = Get-SyncScriptPath
    if (-not (Test-Path -LiteralPath $syncScriptPath)) {
        throw "Expected script was not found: $syncScriptPath"
    }

    $argument = "-NoProfile -ExecutionPolicy Bypass -File `"$syncScriptPath`" -ConfigPath `"$ConfigFilePath`""
    if (-not [string]::IsNullOrWhiteSpace($TaskIdentifier)) {
        $argument += " -TaskIdentifier `"$TaskIdentifier`""
    }
    if ($ForceSchedulerFriendly) {
        $argument += ' -SchedulerFriendly'
    }

    return (New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument)
}

function New-TaskTrigger {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$DailyTimes
    )

    return @(
        $DailyTimes | ForEach-Object {
            New-ScheduledTaskTrigger -Daily -At $_
        }
    )
}

function Get-TaskTimesFromForm {
    $timeValues = @(
        $lstTaskTimes.Items |
            ForEach-Object { [string]$_ }
    )

    if ($timeValues.Count -eq 0) {
        throw 'At least one daily trigger time is required (HH:mm).'
    }

    $invalid = @($timeValues | Where-Object { -not (Test-DailyTimeString -TimeText $_) })
    if ($invalid.Count -gt 0) {
        throw "Invalid time value(s): $($invalid -join ', '). Use 24-hour HH:mm format."
    }

    return @($timeValues | Sort-Object -Unique)
}

function Add-TaskTimeFromPicker {
    $selectedTime = Get-DailyTimeString -TimeValue $dtTaskTime.Value
    $existing = @(
        $lstTaskTimes.Items |
            ForEach-Object { [string]$_ }
    )

    if ($existing -contains $selectedTime) {
        return $false
    }

    [void]$lstTaskTimes.Items.Add($selectedTime)
    $sorted = @(
        $lstTaskTimes.Items |
            ForEach-Object { [string]$_ } |
            Sort-Object -Unique
    )

    $lstTaskTimes.Items.Clear()
    foreach ($timeValue in $sorted) {
        [void]$lstTaskTimes.Items.Add($timeValue)
    }

    $selectedIndex = $lstTaskTimes.Items.IndexOf($selectedTime)
    if ($selectedIndex -ge 0) {
        $lstTaskTimes.SelectedIndex = $selectedIndex
    }

    return $true
}

function Remove-SelectedTaskTimes {
    if ($lstTaskTimes.SelectedItems.Count -eq 0) {
        return 0
    }

    $toRemove = @(
        $lstTaskTimes.SelectedItems |
            ForEach-Object { [string]$_ }
    )
    $removedCount = $toRemove.Count

    foreach ($timeValue in $toRemove) {
        [void]$lstTaskTimes.Items.Remove($timeValue)
    }

    return $removedCount
}

function New-TaskPrincipal {
    $userId = "$env:USERDOMAIN\$env:USERNAME"
    return (New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited)
}

function Format-TaskStatusText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [string]$TaskIdentifier
    )

    $task = Resolve-ScheduledTask -TaskName $TaskName -TaskIdentifier $TaskIdentifier
    $taskInfo = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop
    $nextRun = if ($taskInfo.NextRunTime -and $taskInfo.NextRunTime -ne [datetime]::MinValue) { $taskInfo.NextRunTime } else { 'N/A' }
    $lastRun = if ($taskInfo.LastRunTime -and $taskInfo.LastRunTime -ne [datetime]::MinValue) { $taskInfo.LastRunTime } else { 'N/A' }
    $dailyTriggerTimes = @(
        $task.Triggers |
            Where-Object { $_.CimClass.CimClassName -eq 'MSFT_TaskDailyTrigger' } |
            ForEach-Object {
                if ($_.StartBoundary) {
                    ([datetime]$_.StartBoundary).ToString('HH:mm')
                }
            } |
            Sort-Object -Unique
    )
    $triggerSummary = if ($dailyTriggerTimes.Count -gt 0) { $dailyTriggerTimes -join ', ' } else { 'N/A' }

    return @(
        "Task: $TaskName"
        "State: $($task.State)"
        "Enabled: $($task.Settings.Enabled)"
        "Daily triggers: $triggerSummary"
        "Next run: $nextRun"
        "Last run: $lastRun"
        "Last result: $($taskInfo.LastTaskResult)"
    ) -join [Environment]::NewLine
}

function Start-NewConfig {
    Set-FormConfigValues -Config (New-DefaultConfig)
    $script:currentConfigPath = $null
    Set-TaskTimesList -TimeValues @('02:00')
    Sync-TaskTimesFromExistingTask -Silent
    Update-CurrentFileLabel
    Update-Status 'Started a new config (unsaved).'
}

function Show-StartupWorkflowDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Choose config workflow'
    $dialog.StartPosition = 'CenterParent'
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.ShowInTaskbar = $false
    $dialog.ClientSize = New-Object System.Drawing.Size(380, 132)
    $dialog.TopMost = $true

    $lblPrompt = New-Object System.Windows.Forms.Label
    $lblPrompt.AutoSize = $false
    $lblPrompt.Location = New-Object System.Drawing.Point(16, 16)
    $lblPrompt.Size = New-Object System.Drawing.Size(348, 44)
    $lblPrompt.Text = 'Select how you want to begin editing Backup-Reaper config files.'
    $dialog.Controls.Add($lblPrompt)

    $btnSelectExisting = New-Object System.Windows.Forms.Button
    $btnSelectExisting.Text = 'Select Existing...'
    $btnSelectExisting.Size = New-Object System.Drawing.Size(136, 30)
    $btnSelectExisting.Location = New-Object System.Drawing.Point(78, 76)
    $btnSelectExisting.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($btnSelectExisting)

    $btnNew = New-Object System.Windows.Forms.Button
    $btnNew.Text = 'New'
    $btnNew.Size = New-Object System.Drawing.Size(80, 30)
    $btnNew.Location = New-Object System.Drawing.Point(220, 76)
    $btnNew.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($btnNew)

    $dialog.AcceptButton = $btnSelectExisting
    $dialog.CancelButton = $btnNew

    try {
        return $dialog.ShowDialog($form)
    }
    finally {
        $dialog.Dispose()
    }
}

$btnLoad.Add_Click({
    try {
        if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Load-ConfigIntoForm -PathValue $openDialog.FileName
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Load error', 'OK', 'Error') | Out-Null
        Update-Status 'Load failed.'
    }
})

$btnAddSourcePath.Add_Click({
    try {
        $initialSourcePath = Get-LastNonEmptyLine -Text $txtSourcePaths.Text
        $selectedPath = Select-FolderPath -Title 'Select source folder' -InitialPath $initialSourcePath
        if (-not [string]::IsNullOrWhiteSpace($selectedPath)) {
            Add-PathToTextbox -TextBox $txtSourcePaths -PathValue $selectedPath
            Update-Status 'Added source folder path.'
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Add source path error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to add source folder path.'
    }
})

$btnAddBackupRoot.Add_Click({
    try {
        $initialBackupPath = Get-LastNonEmptyLine -Text $txtBackupRoots.Text
        $selectedPath = Select-FolderPath -Title 'Select backup root folder' -InitialPath $initialBackupPath
        if (-not [string]::IsNullOrWhiteSpace($selectedPath)) {
            Add-PathToTextbox -TextBox $txtBackupRoots -PathValue $selectedPath
            Update-Status 'Added backup root path.'
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Add backup root error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to add backup root path.'
    }
})

$btnNewConfig.Add_Click({
    try {
        Start-NewConfig

        $saveDialog.FileName = 'backup.config.json'
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Save-ConfigFromForm -PathValue $saveDialog.FileName
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'New config error', 'OK', 'Error') | Out-Null
        Update-Status 'New config action failed.'
    }
    finally {
        $saveDialog.FileName = ''
    }
})

$btnSave.Add_Click({
    try {
        if (-not $script:currentConfigPath) {
            if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                return
            }

            $script:currentConfigPath = $saveDialog.FileName
        }

        Save-ConfigFromForm -PathValue $script:currentConfigPath
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Save error', 'OK', 'Error') | Out-Null
        Update-Status 'Save failed.'
    }
})

$btnSaveAs.Add_Click({
    try {
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Save-ConfigFromForm -PathValue $saveDialog.FileName
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Save As error', 'OK', 'Error') | Out-Null
        Update-Status 'Save As failed.'
    }
})

$btnReset.Add_Click({
    Set-FormConfigValues -Config (New-DefaultConfig)
    Update-Status 'Reset fields to defaults.'
})

$btnBrowseLog.Add_Click({
    try {
        $directory = Split-Path -Path $txtLogPath.Text -Parent
        if ([string]::IsNullOrWhiteSpace($directory) -or -not (Test-Path -LiteralPath $directory)) {
            $directory = [Environment]::GetFolderPath('MyDocuments')
        }

        $saveDialog.InitialDirectory = $directory
        $saveDialog.FileName = if ([string]::IsNullOrWhiteSpace($txtLogPath.Text)) { 'backup.log' } else { [System.IO.Path]::GetFileName($txtLogPath.Text) }
        $saveDialog.Filter = 'Log files (*.log)|*.log|All files (*.*)|*.*'
        $saveDialog.Title = 'Choose log file path'

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtLogPath.Text = $saveDialog.FileName
            Update-Status 'Updated log file path.'
        }
    }
    finally {
        $saveDialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
        $saveDialog.Title = 'Save backup config'
        $saveDialog.FileName = ''
    }
})

$btnTaskCreateUpdate.Add_Click({
    try {
        $taskName = Get-TaskNameValue
        $configPath = Ensure-ConfigPathForTask
        if (-not $configPath) {
            Update-Status 'Task creation canceled. Config file is required.'
            return
        }

        $taskIdentifier = Get-OrCreateTaskIdentifier
        Save-ConfigFromForm -PathValue $configPath

        $action = New-TaskAction -ConfigFilePath $configPath -TaskIdentifier $taskIdentifier -ForceSchedulerFriendly $chkTaskSchedulerFriendly.Checked
        $triggerTimes = Get-TaskTimesFromForm
        $trigger = New-TaskTrigger -DailyTimes $triggerTimes
        $principal = New-TaskPrincipal
        $description = "Daily Reaper project backup synchronization. TaskIdentifier: $taskIdentifier"

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Description $description -Principal $principal -Force | Out-Null
        Update-TaskActionButtons
        Sync-TaskTimesFromExistingTask -Silent
        Update-Status "Scheduled task '$taskName' created or updated at: $($triggerTimes -join ', ')."
        [System.Windows.Forms.MessageBox]::Show("Task '$taskName' is configured.", 'Task created/updated', 'OK', 'Information') | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task create/update error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to create or update task.'
    }
})

$btnTaskAddTime.Add_Click({
    try {
        if (Add-TaskTimeFromPicker) {
            Update-Status 'Added daily trigger time.'
        }
        else {
            Update-Status 'Selected time is already in the trigger list.'
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Add trigger time error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to add trigger time.'
    }
})

$btnTaskRemoveTime.Add_Click({
    try {
        $removedCount = Remove-SelectedTaskTimes
        if ($removedCount -gt 0) {
            Update-Status "Removed $removedCount selected trigger time(s)."
        }
        else {
            Update-Status 'Select one or more time lines to remove.'
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Remove trigger time error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to remove selected trigger time(s).'
    }
})

$btnTaskStatus.Add_Click({
    try {
        $taskName = Get-TaskNameValue
        $statusText = Format-TaskStatusText -TaskName $taskName -TaskIdentifier $script:taskIdentifier
        [System.Windows.Forms.MessageBox]::Show($statusText, 'Task status', 'OK', 'Information') | Out-Null
        Update-Status "Loaded status for task '$taskName'."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task status error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to read task status.'
    }
})

$btnTaskRunNow.Add_Click({
    try {
        $taskName = Get-TaskNameValue
        $task = Resolve-ScheduledTask -TaskName $taskName -TaskIdentifier $script:taskIdentifier
        Start-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop
        Update-Status "Started task '$($task.TaskName)'."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task run error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to start task.'
    }
})

$btnTaskRemove.Add_Click({
    try {
        $taskName = Get-TaskNameValue
        $task = Resolve-ScheduledTask -TaskName $taskName -TaskIdentifier $script:taskIdentifier
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove scheduled task '$($task.TaskName)'?",
            'Confirm task removal',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
        Update-TaskActionButtons
        Update-Status "Removed task '$($task.TaskName)'."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task removal error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to remove task.'
    }
})

$txtTaskName.Add_TextChanged({
    Update-TaskActionButtons
    Sync-TaskTimesFromExistingTask -Silent
})

$txtSourcePaths.Add_TextChanged({
    Update-SourcePathsWarning
})

Set-FormConfigValues -Config (New-DefaultConfig)
Set-TaskTimesList -TimeValues @('02:00')
Update-CurrentFileLabel
Update-TaskActionButtons
Sync-TaskTimesFromExistingTask -Silent
Update-SourcePathsWarning

if ($ConfigPath) {
    try {
        $resolvedPath = Resolve-ConfigFilePath -PathValue $ConfigPath
        Load-ConfigIntoForm -PathValue $resolvedPath
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Startup load error', 'OK', 'Warning') | Out-Null
        Update-Status 'Using defaults. Startup config load failed.'
    }
}
else {
    $startupChoice = Show-StartupWorkflowDialog

    if ($startupChoice -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Load-ConfigIntoForm -PathValue $openDialog.FileName
            }
            else {
                Start-NewConfig
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Startup open error', 'OK', 'Warning') | Out-Null
            Start-NewConfig
        }
    }
    else {
        Start-NewConfig
    }
}

[void]$form.ShowDialog()