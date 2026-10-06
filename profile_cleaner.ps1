<#
.SYNOPSIS
    GUI User Profile Manager (Pure PowerShell)
.DESCRIPTION
    Scans a local or remote PC for stale user profiles unused for X days.
    Displays profiles with their size and last login date.
    Provides options to backup via 7-Zip and/or delete stale profiles via CIM.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# --- MAIN FORM ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Windows Stale User Profile Manager (PowerShell)"
$Form.Size = New-Object System.Drawing.Size(960, 720)$Form.StartPosition = "CenterScreen"
$Form.MinimumSize = New-Object System.Drawing.Size(850, 600)

# --- TOP GROUP (Scan Inputs) ---
$GroupScan = New-Object System.Windows.Forms.GroupBox
$GroupScan.Text = " Scan Target & Criteria "
$GroupScan.Location = New-Object System.Drawing.Point(12, 10)
$GroupScan.Size = New-Object System.Drawing.Size(920, 75)$GroupScan.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Computer Name
$lblComp = New-Object System.Windows.Forms.Label
$lblComp.Text = "Computer Name:"
$lblComp.Location = New-Object System.Drawing.Point(15, 28)
$lblComp.AutoSize =$true
$GroupScan.Controls.Add($lblComp)

$txtComp = New-Object System.Windows.Forms.TextBox
$txtComp.Text = "localhost"
$txtComp.Location = New-Object System.Drawing.Point(120, 25)$txtComp.Size = New-Object System.Drawing.Size(140, 23)
$GroupScan.Controls.Add($txtComp)

# Stale Days
$lblDays = New-Object System.Windows.Forms.Label
$lblDays.Text = "Stale Threshold (Days):"
$lblDays.Location = New-Object System.Drawing.Point(280, 28)
$lblDays.AutoSize =$true
$GroupScan.Controls.Add($lblDays)

$numDays = New-Object System.Windows.Forms.NumericUpDown
$numDays.Value = 90
$numDays.Maximum = 3650$numDays.Minimum = 1
$numDays.Location = New-Object System.Drawing.Point(420, 25)$numDays.Size = New-Object System.Drawing.Size(65, 23)
$GroupScan.Controls.Add($numDays)

# 7-Zip Path
$lbl7z = New-Object System.Windows.Forms.Label
$lbl7z.Text = "7-Zip Path:"
$lbl7z.Location = New-Object System.Drawing.Point(505, 28)
$lbl7z.AutoSize =$true
$GroupScan.Controls.Add($lbl7z)

$txt7z = New-Object System.Windows.Forms.TextBox
$txt7z.Text = "C:\Program Files\7-Zip\7z.exe"
$txt7z.Location = New-Object System.Drawing.Point(575, 25)$txt7z.Size = New-Object System.Drawing.Size(190, 23)
$GroupScan.Controls.Add($txt7z)

# Scan Button
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan Profiles"
$btnScan.Location = New-Object System.Drawing.Point(780, 20)$btnScan.Size = New-Object System.Drawing.Size(125, 38)
$btnScan.Font = New-Object System.Drawing.Font($btnScan.Font, [System.Drawing.FontStyle]::Bold)
$GroupScan.Controls.Add($btnScan)

$Form.Controls.Add($GroupScan)

# --- MIDDLE GRID ACTIONS ---
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = New-Object System.Drawing.Point(12, 95)$btnSelectAll.Size = New-Object System.Drawing.Size(90, 25)
$Form.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Deselect All"
$btnDeselectAll.Location = New-Object System.Drawing.Point(108, 95)$btnDeselectAll.Size = New-Object System.Drawing.Size(90, 25)
$Form.Controls.Add($btnDeselectAll)

# --- DATA GRID ---
$Grid = New-Object System.Windows.Forms.DataGridView
$Grid.Location = New-Object System.Drawing.Point(12, 125)$Grid.Size = New-Object System.Drawing.Size(920, 420)
$Grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right$Grid.AllowUserToAddRows = $false$Grid.AllowUserToDeleteRows = $false$Grid.SelectionMode = "FullRowSelect"
$Grid.AutoSizeColumnsMode = "Fill"
$Grid.MultiSelect =$false

# Columns
$colChk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colChk.HeaderText = "[X]"
$colChk.Name = "Select"
$colChk.Width = 45
$Grid.Columns.Add($colChk) | Out-Null

$Grid.Columns.Add("Username", "Username") | Out-Null
$Grid.Columns.Add("LastLogin", "Last Login Date") | Out-Null
$Grid.Columns.Add("SizeGB", "Size (GB)") | Out-Null
$Grid.Columns.Add("Path", "Profile Path") | Out-Null
$Grid.Columns.Add("SID", "SID") | Out-Null

$Grid.Columns["Username"].ReadOnly = $true
$Grid.Columns["LastLogin"].ReadOnly = $true
$Grid.Columns["SizeGB"].ReadOnly = $true
$Grid.Columns["Path"].ReadOnly = $true
$Grid.Columns["SID"].ReadOnly = $true
$Grid.Columns["SID"].Visible = $false

$Form.Controls.Add($Grid)

# --- BOTTOM GROUP (Actions & Settings) ---
$GroupAction = New-Object System.Windows.Forms.GroupBox
$GroupAction.Text = " Actions & Options "
$GroupAction.Location = New-Object System.Drawing.Point(12, 555)
$GroupAction.Size = New-Object System.Drawing.Size(920, 85)$GroupAction.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Backup Directory
$lblBackupDir = New-Object System.Windows.Forms.Label
$lblBackupDir.Text = "Backup Destination (on Target PC):"
$lblBackupDir.Location = New-Object System.Drawing.Point(15, 28)
$lblBackupDir.AutoSize =$true
$GroupAction.Controls.Add($lblBackupDir)

$txtBackupDir = New-Object System.Windows.Forms.TextBox
$txtBackupDir.Text = "C:\ProfileBackups"
$txtBackupDir.Location = New-Object System.Drawing.Point(210, 25)$txtBackupDir.Size = New-Object System.Drawing.Size(220, 23)
$GroupAction.Controls.Add($txtBackupDir)

# Dry Run Checkbox
$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run Mode (Simulate actions only)"
$chkDryRun.Checked =$true
$chkDryRun.Location = New-Object System.Drawing.Point(15, 55)$chkDryRun.AutoSize = $true$chkDryRun.Font = New-Object System.Drawing.Font($chkDryRun.Font, [System.Drawing.FontStyle]::Bold)$chkDryRun.ForeColor = [System.Drawing.Color]::DarkBlue
$GroupAction.Controls.Add($chkDryRun)

# Action Buttons
$btnBackup = New-Object System.Windows.Forms.Button
$btnBackup.Text = "Backup Selected"
$btnBackup.Location = New-Object System.Drawing.Point(450, 25)$btnBackup.Size = New-Object System.Drawing.Size(140, 42)
$GroupAction.Controls.Add($btnBackup)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete Selected"
$btnDelete.Location = New-Object System.Drawing.Point(600, 25)
$btnDelete.Size = New-Object System.Drawing.Size(140, 42)$btnDelete.ForeColor = [System.Drawing.Color]::DarkRed
$GroupAction.Controls.Add($btnDelete)

$btnBackupDelete = New-Object System.Windows.Forms.Button
$btnBackupDelete.Text = "Backup & Delete"
$btnBackupDelete.Location = New-Object System.Drawing.Point(750, 25)$btnBackupDelete.Size = New-Object System.Drawing.Size(155, 42)
$btnBackupDelete.Font = New-Object System.Drawing.Font($btnBackupDelete.Font, [System.Drawing.FontStyle]::Bold)
$GroupAction.Controls.Add($btnBackupDelete)

$Form.Controls.Add($GroupAction)

# --- STATUS BAR ---
$StatusBar = New-Object System.Windows.Forms.StatusStrip
$StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$StatusLabel.Text = "Ready"
$StatusBar.Items.Add($StatusLabel) | Out-Null
$Form.Controls.Add($StatusBar)


# ==========================================
# LOGIC & EVENT HANDLERS
# ==========================================

# Helper: Update Status Text
function Set-Status ($Text) {$StatusLabel.Text = $Text$StatusBar.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# 1. SCAN PROFILES
$btnScan.Add_Click({
    $Computer =$txtComp.Text.Trim()
    $Days = [int]$numDays.Value
    
    if ([string]::IsNullOrWhiteSpace($Computer)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name.", "Error", "OK", "Error")
        return
    }

    $Grid.Rows.Clear()
    Set-Status "Scanning $Computer for profiles older than$Days days... Please wait."
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $ScanScript = {
        param([int]$CutoffDays)
        $CutoffDate = (Get-Date).AddDays(-$CutoffDays)

        Get-CimInstance -ClassName Win32_UserProfile | Where-Object {
            -not $_.Special -and -not $_.Loaded -and$_.LastUseTime -and ($_.LastUseTime -lt$CutoffDate)
        } | ForEach-Object {
            $Path =$_.LocalPath
            $SizeGB = 0
            if (Test-Path -Path $Path) {
                $SizeSum = (Get-ChildItem -Path$Path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                if ($SizeSum) {
                    $SizeGB = [math]::Round($SizeSum / 1GB, 2)
                }
            }
            [PSCustomObject]@{
                SID         = $_.SID
                Username    = Split-Path -Path $Path -Leaf
                LastUseTime = $_.LastUseTime.ToString("yyyy-MM-dd HH:mm:ss")
                SizeGB      = $SizeGB
                LocalPath   = $Path
            }
        }
    }

    try {
        if ($Computer -match "^(localhost|127\.0\.0\.1|$env:COMPUTERNAME)$") {
            $Results = & $ScanScript -CutoffDays$Days
        } else {
            $Results = Invoke-Command -ComputerName$Computer -ScriptBlock $ScanScript -ArgumentList$Days -ErrorAction Stop
        }

        if ($Results) {
            foreach ($item in $Results) {$Grid.Rows.Add($false,$item.Username, $item.LastUseTime, $item.SizeGB, $item.LocalPath, $item.SID) | Out-Null
            }
            Set-Status "Scan complete: Found $($Grid.Rows.Count) stale profile(s)."
        } else {
            Set-Status "Scan complete: No stale user profiles found."
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Scan failed: $($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "Scan failed."
    }
    finally {
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# 2. SELECT / DESELECT ALL
$btnSelectAll.Add_Click({
    foreach ($row in$Grid.Rows) { $row.Cells["Select"].Value = $true }
})

$btnDeselectAll.Add_Click({
    foreach ($row in$Grid.Rows) { $row.Cells["Select"].Value = $false }
})

# Helper: Get Selected Rows
function Get-SelectedRows {
    return @($Grid.Rows | Where-Object { $_.Cells["Select"].Value -eq $true })
}

# 3. BACKUP ACTION
function Invoke-ProfileBackup ($SelectedRows) {
    $Computer =$txtComp.Text.Trim()
    $BackupDir =$txtBackupDir.Text.Trim()
    $SevenZipExe =$txt7z.Text.Trim()
    $IsDryRun =$chkDryRun.Checked

    $BackupScript = {
        param([string]$Path, [string]$ZipPath, [string]$ExePath)
        if (-not (Test-Path $ExePath)) {
            throw "7-Zip executable not found at '$ExePath' on target machine."
        }
        $Dir = Split-Path -Path$ZipPath -Parent
        if (-not (Test-Path $Dir)) { New-Item -Path$Dir -ItemType Directory -Force | Out-Null }

        $args = @("a", "-t7z", "-ssw", "-mx5", "`"$ZipPath`"", "`"$Path\*`"")
        $p = Start-Process -FilePath $ExePath -ArgumentList$args -Wait -NoNewWindow -PassThru
        return ($p.ExitCode -eq 0)
    }

    $SuccessCount = 0
    foreach ($row in$SelectedRows) {
        $User =$row.Cells["Username"].Value
        $Path =$row.Cells["Path"].Value
        $TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $ZipPath = "$BackupDir\${User}_Backup_${TimeStamp}.7z"

        if ($IsDryRun)
