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
$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = "Windows Stale User Profile Manager (PowerShell)"
$script:Form.Size = New-Object System.Drawing.Size(960, 720)$script:Form.StartPosition = "CenterScreen"
$script:Form.MinimumSize = New-Object System.Drawing.Size(850, 600)

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

$script:txtComp = New-Object System.Windows.Forms.TextBox
$script:txtComp.Text = "localhost"
$script:txtComp.Location = New-Object System.Drawing.Point(120, 25)$script:txtComp.Size = New-Object System.Drawing.Size(140, 23)
$GroupScan.Controls.Add($script:txtComp)

# Stale Days
$lblDays = New-Object System.Windows.Forms.Label
$lblDays.Text = "Stale Threshold (Days):"
$lblDays.Location = New-Object System.Drawing.Point(280, 28)
$lblDays.AutoSize =$true
$GroupScan.Controls.Add($lblDays)

$script:numDays = New-Object System.Windows.Forms.NumericUpDown
$script:numDays.Value = 90
$script:numDays.Maximum = 3650$script:numDays.Minimum = 1
$script:numDays.Location = New-Object System.Drawing.Point(420, 25)$script:numDays.Size = New-Object System.Drawing.Size(65, 23)
$GroupScan.Controls.Add($script:numDays)

# 7-Zip Path
$lbl7z = New-Object System.Windows.Forms.Label
$lbl7z.Text = "7-Zip Path:"
$lbl7z.Location = New-Object System.Drawing.Point(505, 28)
$lbl7z.AutoSize =$true
$GroupScan.Controls.Add($lbl7z)

$script:txt7z = New-Object System.Windows.Forms.TextBox
$script:txt7z.Text = "C:\Program Files\7-Zip\7z.exe"
$script:txt7z.Location = New-Object System.Drawing.Point(575, 25)$script:txt7z.Size = New-Object System.Drawing.Size(190, 23)
$GroupScan.Controls.Add($script:txt7z)

# Scan Button
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan Profiles"
$btnScan.Location = New-Object System.Drawing.Point(780, 20)$btnScan.Size = New-Object System.Drawing.Size(125, 38)
$btnScan.Font = New-Object System.Drawing.Font($btnScan.Font, [System.Drawing.FontStyle]::Bold)
$GroupScan.Controls.Add($btnScan)

$script:Form.Controls.Add($GroupScan)

# --- MIDDLE GRID ACTIONS ---
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = New-Object System.Drawing.Point(12, 95)$btnSelectAll.Size = New-Object System.Drawing.Size(90, 25)
$script:Form.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Deselect All"
$btnDeselectAll.Location = New-Object System.Drawing.Point(108, 95)$btnDeselectAll.Size = New-Object System.Drawing.Size(90, 25)
$script:Form.Controls.Add($btnDeselectAll)

# --- DATA GRID ---
$script:Grid = New-Object System.Windows.Forms.DataGridView
$script:Grid.Location = New-Object System.Drawing.Point(12, 125)$script:Grid.Size = New-Object System.Drawing.Size(920, 420)
$script:Grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right$script:Grid.AllowUserToAddRows = $false$script:Grid.AllowUserToDeleteRows = $false$script:Grid.SelectionMode = "FullRowSelect"
$script:Grid.AutoSizeColumnsMode = "Fill"
$script:Grid.MultiSelect =$false

# Columns
$colChk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colChk.HeaderText = "[X]"
$colChk.Name = "Select"
$colChk.Width = 45
[void]$script:Grid.Columns.Add($colChk)

[void]$script:Grid.Columns.Add("Username", "Username")
[void]$script:Grid.Columns.Add("LastLogin", "Last Login Date")
[void]$script:Grid.Columns.Add("SizeGB", "Size (GB)")
[void]$script:Grid.Columns.Add("Path", "Profile Path")
[void]$script:Grid.Columns.Add("SID", "SID")

$script:Grid.Columns["Username"].ReadOnly = $true
$script:Grid.Columns["LastLogin"].ReadOnly = $true
$script:Grid.Columns["SizeGB"].ReadOnly = $true
$script:Grid.Columns["Path"].ReadOnly = $true
$script:Grid.Columns["SID"].ReadOnly = $true
$script:Grid.Columns["SID"].Visible = $false

$script:Form.Controls.Add($script:Grid)

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

$script:txtBackupDir = New-Object System.Windows.Forms.TextBox
$script:txtBackupDir.Text = "C:\ProfileBackups"
$script:txtBackupDir.Location = New-Object System.Drawing.Point(210, 25)$script:txtBackupDir.Size = New-Object System.Drawing.Size(220, 23)
$GroupAction.Controls.Add($script:txtBackupDir)

# Dry Run Checkbox
$script:chkDryRun = New-Object System.Windows.Forms.CheckBox
$script:chkDryRun.Text = "Dry Run Mode (Simulate actions only)"
$script:chkDryRun.Checked =$true
$script:chkDryRun.Location = New-Object System.Drawing.Point(15, 55)$script:chkDryRun.AutoSize = $true$script:chkDryRun.Font = New-Object System.Drawing.Font($script:chkDryRun.Font, [System.Drawing.FontStyle]::Bold)$script:chkDryRun.ForeColor = [System.Drawing.Color]::DarkBlue
$GroupAction.Controls.Add($script:chkDryRun)

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

$script:Form.Controls.Add($GroupAction)

# --- STATUS BAR ---
$script:StatusBar = New-Object System.Windows.Forms.StatusStrip
$script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusLabel.Text = "Ready"
[void]$script:StatusBar.Items.Add($script:StatusLabel)
$script:Form.Controls.Add($script:StatusBar)


# ==========================================
# LOGIC & EVENT HANDLERS
# ==========================================

# Helper: Update Status Text
function Set-Status ([string]$Text) {$script:StatusLabel.Text = $Text$script:StatusBar.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# Helper: Get Selected Rows safely
function Get-SelectedRows {
    $selected = @()
    foreach ($row in$script:Grid.Rows) {
        $val =$row.Cells["Select"].Value
        if ($null -ne$val -and [bool]$val -eq$true) {
            $selected +=$row
        }
    }
    return $selected
}

# 1. SCAN PROFILES
$btnScan.Add_Click({
    $Computer =$script:txtComp.Text.Trim()
    $Days = [int]$script:numDays.Value
    
    if ([string]::IsNullOrWhiteSpace($Computer)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name.", "Error", "OK", "Error")
        return
    }

    $script:Grid.Rows.Clear()
    Set-Status "Scanning $Computer for profiles older than$Days days... Please wait."
    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

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
            foreach ($item in$Results) {
                [void]$script:Grid.Rows.Add($false,$item.Username, $item.LastUseTime, $item.SizeGB, $item.LocalPath, $item.SID)
            }
            Set-Status "Scan complete: Found $($script:Grid.Rows.Count) stale profile(s)."
        } else {
            Set-Status "Scan complete: No stale user profiles found."
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Scan failed: $($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "Scan failed."
    }
    finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# 2. SELECT / DESELECT ALL
$btnSelectAll.Add_Click({
    foreach ($row in$script:Grid.Rows) { $row.Cells["Select"].Value = $true }
})

$btnDeselectAll.Add_Click({
    foreach ($row in$script:Grid.Rows) { $row.Cells["Select"].Value = $false }
})

# 3. BACKUP ACTION
function Invoke-ProfileBackup ($SelectedRows) {
    $Computer =$script:txtComp.Text.Trim()
    $BackupDir =$script:txtBackupDir.Text.Trim()
    $SevenZipExe =$script:txt7z.Text.Trim()
    $IsDryRun =$script:chkDryRun.Checked

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

        if ($IsDryRun) {
            Set-Status "[DRY RUN] Would backup '$User' to '$ZipPath'"
            Start-Sleep -Milliseconds 300
            $SuccessCount++
        } else {
            Set-Status "Backing up $User to$ZipPath..."
            try {
                if ($Computer -match "^(localhost|127\.0\.0\.1|$env:COMPUTERNAME)$") {
                    $ok = & $BackupScript -Path$Path -ZipPath $ZipPath -ExePath$SevenZipExe
                } else {
                    $ok = Invoke-Command -ComputerName$Computer -ScriptBlock $BackupScript -ArgumentList$Path, $ZipPath,$SevenZipExe -ErrorAction Stop
                }
                if ($ok) {$SuccessCount++ }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Backup failed for $User: $($_.Exception.Message)", "Backup Error", "OK", "Error")
            }
        }
    }
    return $SuccessCount
}

# 4. DELETE ACTION
function Invoke-ProfileDelete ($SelectedRows) {
    $Computer =$script:txtComp.Text.Trim()
    $IsDryRun =$script:chkDryRun.Checked

    $DeleteScript = {
        param([string]$TargetSID)$profile = Get-CimInstance -ClassName Win32_UserProfile | Where-Object { $_.SID -eq$TargetSID }
        if ($profile) {
            Remove-CimInstance -InputObject $profile -ErrorAction Stop
            return $true
        }
        return $false
    }

    $RowsToRemove = @()
    foreach ($row in$SelectedRows) {
        $User =$row.Cells["Username"].Value
        $SID =$row.Cells["SID"].Value

        if ($IsDryRun) {
            Set-Status "[DRY RUN] Would delete CIM profile for '$User' (SID:$SID)"
            Start-Sleep -Milliseconds 300
            $RowsToRemove +=$row
        } else {
            Set-Status "Deleting profile $User..."
            try {
                if ($Computer -match "^(localhost|127\.0\.0\.1|$env:COMPUTERNAME)$") {
                    & $DeleteScript -TargetSID$SID | Out-Null
                } else {
                    Invoke-Command -ComputerName $Computer -ScriptBlock $DeleteScript -ArgumentList$SID -ErrorAction Stop | Out-Null
                }
                $RowsToRemove +=$row
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Deletion failed for $User: $($_.Exception.Message)", "Delete Error", "OK", "Error")
            }
        }
    }

    # Remove deleted rows from UI
    foreach ($r in$RowsToRemove) {
        $script:Grid.Rows.Remove($r)
    }
    return $RowsToRemove.Count
}

# BUTTON CLICK HANDLERS
$btnBackup.Add_Click({$Selected = Get-SelectedRows
    if ($Selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one profile.", "Notice", "OK", "Information")
        return
    }
    $count = Invoke-ProfileBackup$Selected
    Set-Status "Backup completed for $count profile(s)."
})

$btnDelete.Add_Click({$Selected = Get-SelectedRows
    if ($Selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one profile.", "Notice", "OK", "Information")
        return
    }

    $msg = "Are you sure you want to PERMANENTLY delete $($Selected.Count) selected profile(s)?"
    if ($script:chkDryRun.Checked) {$msg = "[DRY RUN] Simulate deletion of $($Selected.Count) profile(s)?" }
    
    $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Deletion", "YesNo", "Warning")
    if ($confirm -eq "Yes") {
        $count = Invoke-ProfileDelete$Selected
        Set-Status "Deletion completed for $count profile(s)."
    }
})

$btnBackupDelete.Add_Click({$Selected = Get-SelectedRows
    if ($Selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one profile.", "Notice", "OK", "Information")
        return
    }

    $msg = "Backup AND Delete $($Selected.Count) selected profile(s)?"
    if ($script:chkDryRun.Checked) {$msg = "[DRY RUN] Simulate Backup & Delete of $($Selected.Count) profile(s)?" }

    $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Action", "YesNo", "Question")
    if ($confirm -eq "Yes") {
        $backedUp = Invoke-ProfileBackup$Selected
        if ($backedUp -gt 0) {
            $deleted = Invoke-ProfileDelete$Selected
            Set-Status "Completed Backup ($backedUp) & Deletion ($deleted)."
        }
    }
})

# SHOW FORM
[void]$script:Form.ShowDialog()
