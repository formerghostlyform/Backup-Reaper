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
        RobocopyTuning = 'Normal'
        LogPath = 'C:\Logs\Backup-Reaper\backup.log'
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

    if ($null -ne $InputObject.RobocopyTuning -and [string]$InputObject.RobocopyTuning) {
        $result.RobocopyTuning = [string]$InputObject.RobocopyTuning
    }

    if ($null -ne $InputObject.LogPath -and [string]$InputObject.LogPath) {
        $result.LogPath = [string]$InputObject.LogPath
    }

    return $result
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

    $ordered = [ordered]@{
        SourcePaths = $Config.SourcePaths
        BackupRoots = $Config.BackupRoots
        PreserveTargetOnlyFiles = [bool]$Config.PreserveTargetOnlyFiles
        DryRun = [bool]$Config.DryRun
        SchedulerFriendly = [bool]$Config.SchedulerFriendly
        SkipPreflightSizeScan = [bool]$Config.SkipPreflightSizeScan
        RobocopyTuning = [string]$Config.RobocopyTuning
        LogPath = [string]$Config.LogPath
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
    $cmbRobocopyTuning.SelectedItem = [string]$Config.RobocopyTuning
    $txtLogPath.Text = [string]$Config.LogPath
}

function Get-FormConfigValues {
    $config = [ordered]@{
        SourcePaths = (Convert-MultilineToArray -Text $txtSourcePaths.Text)
        BackupRoots = (Convert-MultilineToArray -Text $txtBackupRoots.Text)
        PreserveTargetOnlyFiles = $chkPreserveTargetOnlyFiles.Checked
        DryRun = $chkDryRun.Checked
        SchedulerFriendly = $chkSchedulerFriendly.Checked
        SkipPreflightSizeScan = $chkSkipPreflightSizeScan.Checked
        RobocopyTuning = [string]$cmbRobocopyTuning.SelectedItem
        LogPath = $txtLogPath.Text.Trim()
    }

    Validate-ConfigValues -Config $config
    return $config
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Backup-Reaper Config Editor'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(860, 790)
$form.MinimumSize = New-Object System.Drawing.Size(760, 720)

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
$txtSourcePaths.Size = New-Object System.Drawing.Size(370, 255)
$txtSourcePaths.Anchor = 'Top,Left,Bottom'
$groupPaths.Controls.Add($txtSourcePaths)

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
$groupOptions.Size = New-Object System.Drawing.Size(812, 178)
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

$lblTuning = New-Object System.Windows.Forms.Label
$lblTuning.Text = 'RobocopyTuning'
$lblTuning.AutoSize = $true
$lblTuning.Location = New-Object System.Drawing.Point(19, 67)
$groupOptions.Controls.Add($lblTuning)

$cmbRobocopyTuning = New-Object System.Windows.Forms.ComboBox
$cmbRobocopyTuning.DropDownStyle = 'DropDownList'
$cmbRobocopyTuning.Location = New-Object System.Drawing.Point(19, 88)
$cmbRobocopyTuning.Size = New-Object System.Drawing.Size(160, 25)
$cmbRobocopyTuning.Items.AddRange(@('Normal', 'High', 'Higher'))
$groupOptions.Controls.Add($cmbRobocopyTuning)

$lblLogPath = New-Object System.Windows.Forms.Label
$lblLogPath.Text = 'LogPath'
$lblLogPath.AutoSize = $true
$lblLogPath.Location = New-Object System.Drawing.Point(208, 67)
$groupOptions.Controls.Add($lblLogPath)

$txtLogPath = New-Object System.Windows.Forms.TextBox
$txtLogPath.Location = New-Object System.Drawing.Point(211, 88)
$txtLogPath.Size = New-Object System.Drawing.Size(491, 25)
$txtLogPath.Anchor = 'Top,Left,Right'
$groupOptions.Controls.Add($txtLogPath)

$btnBrowseLog = New-Object System.Windows.Forms.Button
$btnBrowseLog.Text = 'Browse...'
$btnBrowseLog.Size = New-Object System.Drawing.Size(88, 27)
$btnBrowseLog.Location = New-Object System.Drawing.Point(708, 87)
$btnBrowseLog.Anchor = 'Top,Right'
$groupOptions.Controls.Add($btnBrowseLog)

$groupTask = New-Object System.Windows.Forms.GroupBox
$groupTask.Text = 'Scheduled Task'
$groupTask.Location = New-Object System.Drawing.Point(16, 610)
$groupTask.Size = New-Object System.Drawing.Size(812, 120)
$groupTask.Anchor = 'Left,Right,Bottom'
$form.Controls.Add($groupTask)

$lblTaskName = New-Object System.Windows.Forms.Label
$lblTaskName.Text = 'Task name'
$lblTaskName.AutoSize = $true
$lblTaskName.Location = New-Object System.Drawing.Point(19, 29)
$groupTask.Controls.Add($lblTaskName)

$txtTaskName = New-Object System.Windows.Forms.TextBox
$txtTaskName.Location = New-Object System.Drawing.Point(19, 49)
$txtTaskName.Size = New-Object System.Drawing.Size(290, 25)
$txtTaskName.Text = 'Reaper Backup Sync'
$groupTask.Controls.Add($txtTaskName)

$lblTaskTime = New-Object System.Windows.Forms.Label
$lblTaskTime.Text = 'Daily at'
$lblTaskTime.AutoSize = $true
$lblTaskTime.Location = New-Object System.Drawing.Point(329, 29)
$groupTask.Controls.Add($lblTaskTime)

$dtTaskTime = New-Object System.Windows.Forms.DateTimePicker
$dtTaskTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$dtTaskTime.ShowUpDown = $true
$dtTaskTime.Width = 110
$dtTaskTime.Location = New-Object System.Drawing.Point(329, 49)
$dtTaskTime.Value = [datetime]::Today.AddHours(2)
$groupTask.Controls.Add($dtTaskTime)

$btnTaskCreateUpdate = New-Object System.Windows.Forms.Button
$btnTaskCreateUpdate.Text = 'Create / Update Task'
$btnTaskCreateUpdate.Size = New-Object System.Drawing.Size(145, 30)
$btnTaskCreateUpdate.Location = New-Object System.Drawing.Point(468, 45)
$groupTask.Controls.Add($btnTaskCreateUpdate)

$btnTaskStatus = New-Object System.Windows.Forms.Button
$btnTaskStatus.Text = 'Task Status'
$btnTaskStatus.Size = New-Object System.Drawing.Size(90, 30)
$btnTaskStatus.Location = New-Object System.Drawing.Point(620, 45)
$groupTask.Controls.Add($btnTaskStatus)

$btnTaskRunNow = New-Object System.Windows.Forms.Button
$btnTaskRunNow.Text = 'Run Now'
$btnTaskRunNow.Size = New-Object System.Drawing.Size(80, 30)
$btnTaskRunNow.Location = New-Object System.Drawing.Point(716, 45)
$groupTask.Controls.Add($btnTaskRunNow)

$btnTaskRemove = New-Object System.Windows.Forms.Button
$btnTaskRemove.Text = 'Remove Task'
$btnTaskRemove.Size = New-Object System.Drawing.Size(110, 28)
$btnTaskRemove.Location = New-Object System.Drawing.Point(19, 82)
$groupTask.Controls.Add($btnTaskRemove)

$chkTaskSchedulerFriendly = New-Object System.Windows.Forms.CheckBox
$chkTaskSchedulerFriendly.Text = 'Force -SchedulerFriendly for scheduled runs'
$chkTaskSchedulerFriendly.AutoSize = $true
$chkTaskSchedulerFriendly.Checked = $true
$chkTaskSchedulerFriendly.Location = New-Object System.Drawing.Point(140, 86)
$groupTask.Controls.Add($chkTaskSchedulerFriendly)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(16, 740)
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

function Test-TaskExists {
    param(
        [string]$TaskName
    )

    if ([string]::IsNullOrWhiteSpace($TaskName)) {
        return $false
    }

    try {
        $null = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Update-TaskActionButtons {
    $taskName = $txtTaskName.Text.Trim()
    $taskExists = Test-TaskExists -TaskName $taskName
    $btnTaskStatus.Enabled = $taskExists
    $btnTaskRunNow.Enabled = $taskExists
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

        [bool]$ForceSchedulerFriendly
    )

    $syncScriptPath = Get-SyncScriptPath
    if (-not (Test-Path -LiteralPath $syncScriptPath)) {
        throw "Expected script was not found: $syncScriptPath"
    }

    $argument = "-NoProfile -ExecutionPolicy Bypass -File `"$syncScriptPath`" -ConfigPath `"$ConfigFilePath`""
    if ($ForceSchedulerFriendly) {
        $argument += ' -SchedulerFriendly'
    }

    return (New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument)
}

function New-TaskTrigger {
    $dailyAt = Get-DailyTimeString -TimeValue $dtTaskTime.Value
    return (New-ScheduledTaskTrigger -Daily -At $dailyAt)
}

function New-TaskPrincipal {
    $userId = "$env:USERDOMAIN\$env:USERNAME"
    return (New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited)
}

function Format-TaskStatusText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskName
    )

    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction Stop
    $nextRun = if ($taskInfo.NextRunTime -and $taskInfo.NextRunTime -ne [datetime]::MinValue) { $taskInfo.NextRunTime } else { 'N/A' }
    $lastRun = if ($taskInfo.LastRunTime -and $taskInfo.LastRunTime -ne [datetime]::MinValue) { $taskInfo.LastRunTime } else { 'N/A' }

    return @(
        "Task: $TaskName"
        "State: $($task.State)"
        "Enabled: $($task.Settings.Enabled)"
        "Next run: $nextRun"
        "Last run: $lastRun"
        "Last result: $($taskInfo.LastTaskResult)"
    ) -join [Environment]::NewLine
}

function Start-NewConfig {
    Set-FormConfigValues -Config (New-DefaultConfig)
    $script:currentConfigPath = $null
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

        $action = New-TaskAction -ConfigFilePath $configPath -ForceSchedulerFriendly $chkTaskSchedulerFriendly.Checked
        $trigger = New-TaskTrigger
        $principal = New-TaskPrincipal
        $description = 'Daily Reaper project backup synchronization.'

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Description $description -Principal $principal -Force | Out-Null
        Update-TaskActionButtons
        Update-Status "Scheduled task '$taskName' created or updated."
        [System.Windows.Forms.MessageBox]::Show("Task '$taskName' is configured.", 'Task created/updated', 'OK', 'Information') | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task create/update error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to create or update task.'
    }
})

$btnTaskStatus.Add_Click({
    try {
        $taskName = Get-TaskNameValue
        $statusText = Format-TaskStatusText -TaskName $taskName
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
        Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
        Update-Status "Started task '$taskName'."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task run error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to start task.'
    }
})

$btnTaskRemove.Add_Click({
    try {
        $taskName = Get-TaskNameValue
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove scheduled task '$taskName'?",
            'Confirm task removal',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
        Update-TaskActionButtons
        Update-Status "Removed task '$taskName'."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Task removal error', 'OK', 'Error') | Out-Null
        Update-Status 'Unable to remove task.'
    }
})

$txtTaskName.Add_TextChanged({
    Update-TaskActionButtons
})

Set-FormConfigValues -Config (New-DefaultConfig)
Update-CurrentFileLabel
Update-TaskActionButtons

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