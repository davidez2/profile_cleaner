#Requires -Version 5.1
<#
.SYNOPSIS
    GUI tool: find stale user profiles on a remote PC, back them up with 7-Zip ON THAT PC, then delete them.

.DESCRIPTION
    - Scan: lists profiles via CIM (WinRM with DCOM fallback). Last activity = the newer of
      Win32_UserProfile.LastUseTime and the NTUSER.DAT last-write time (read over the C$ share).
    - Backup: 7-Zip runs ON THE REMOTE PC (PowerShell remoting / WinRM) against the local profile folder
      and writes the .7z archive to a folder on the remote PC. No profile data crosses the network.
    - A profile is deleted only if its archive was created AND passed "7z t".
    - Deletion uses Remove-CimInstance on Win32_UserProfile (removes folder + registry entry).
    - DRY RUN is ON by default: nothing is archived or deleted. It does check remoting, 7-Zip,
      the backup drive and estimates profile sizes against free space.

    Requirements: run elevated, as (or with credentials of) a local admin on the remote PC.
    WinRM (Enable-PSRemoting) must be enabled on the remote PC for the backup step.
    7-Zip must be installed on the remote PC, OR give a local 7za.exe (standalone "7-Zip Extra")
    / 7z.exe and the script copies it to the remote temp folder for the run and removes it afterwards.
#>

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ----------------------------------------------------------------- state
$script:Cred        = $null
$script:CimSession  = $null
$script:PSSess      = $null
$script:Cancel      = $false
$script:Mappings    = @()
$script:LogFile     = $null
$script:RemoteTemp  = $null
$script:ProfileInfo = @{}     # SID -> hashtable (Name, LocalPath, Unc)

# ----------------------------------------------------------------- helpers
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0}  [{1}]  {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    if ($script:LogFile) { try { Add-Content -Path $script:LogFile -Value $line -ErrorAction Stop } catch { } }
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-CimConn {
    param([string]$Computer)
    if ($script:CimSession) { Remove-CimSession $script:CimSession -ErrorAction SilentlyContinue; $script:CimSession = $null }
    $p = @{ ComputerName = $Computer; ErrorAction = 'Stop' }
    if ($script:Cred) { $p.Credential = $script:Cred }
    try {
        $script:CimSession = New-CimSession @p
        Write-Log "CIM connected to $Computer (WSMan)."
    } catch {
        Write-Log "CIM over WSMan failed ($($_.Exception.Message.Trim())). Trying DCOM..." 'WARN'
        $p.SessionOption = New-CimSessionOption -Protocol Dcom
        $script:CimSession = New-CimSession @p
        Write-Log "CIM connected to $Computer (DCOM)."
    }
    return $script:CimSession
}

function Get-PSConn {
    param([string]$Computer)
    if ($script:PSSess) { Remove-PSSession $script:PSSess -ErrorAction SilentlyContinue; $script:PSSess = $null }
    $p = @{ ComputerName = $Computer; ErrorAction = 'Stop' }
    if ($script:Cred) { $p.Credential = $script:Cred }
    try {
        $script:PSSess = New-PSSession @p
        Write-Log "PowerShell remoting session opened to $Computer."
    } catch {
        throw "Cannot open a PowerShell remoting session to $Computer (needed to run 7-Zip remotely). Is WinRM enabled there (Enable-PSRemoting)? Details: $($_.Exception.Message.Trim())"
    }
}

function ConvertTo-UncPath {
    param([string]$Computer, [string]$LocalPath)
    '\\{0}\{1}${2}' -f $Computer, $LocalPath.Substring(0, 1), $LocalPath.Substring(2)   # C:\Users\bob -> \\PC\C$\Users\bob
}

function Connect-Share {
    # Only used by the scan, to read NTUSER.DAT timestamps. Not needed when running as the current user.
    param([string]$Computer, [string]$DriveLetter)
    if (-not $script:Cred) { return }
    $remote = "\\$Computer\$DriveLetter`$"
    try {
        New-SmbMapping -RemotePath $remote -UserName $script:Cred.UserName `
            -Password $script:Cred.GetNetworkCredential().Password -ErrorAction Stop | Out-Null
        $script:Mappings += $remote
    } catch {
        Write-Log "Could not map $remote ($($_.Exception.Message.Trim())); NTUSER.DAT dates may be missing." 'WARN'
    }
}

function Disconnect-Shares {
    foreach ($m in $script:Mappings) { Remove-SmbMapping -RemotePath $m -Force -UpdateProfile -ErrorAction SilentlyContinue }
    $script:Mappings = @()
}

function Invoke-Remote {
    # Runs a scriptblock in the remote session as a job so the GUI stays responsive and Cancel works.
    # Returns the job output, or $null if cancelled. $KillMatch = text in a 7z command line to kill on cancel.
    param([scriptblock]$Script, [object[]]$ArgList, [string]$KillMatch)
    $job = Invoke-Command -Session $script:PSSess -ScriptBlock $Script -ArgumentList $ArgList -AsJob
    while ($job.State -in 'Running', 'NotStarted') {
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:Cancel) {
            Stop-Job $job -ErrorAction SilentlyContinue
            if ($KillMatch) {
                try {
                    Invoke-Command -Session $script:PSSess -ArgumentList $KillMatch -ScriptBlock {
                        param($m)
                        Get-CimInstance Win32_Process -Filter "Name LIKE '7z%'" |
                            Where-Object { $_.CommandLine -like "*$m*" } |
                            ForEach-Object { Invoke-CimMethod -InputObject $_ -MethodName Terminate | Out-Null }
                    } | Out-Null
                } catch { }
            }
            Remove-Job $job -Force -ErrorAction SilentlyContinue
            return $null
        }
        Start-Sleep -Milliseconds 250
    }
    try { $res = Receive-Job $job -ErrorAction Stop } finally { Remove-Job $job -Force -ErrorAction SilentlyContinue }
    return $res
}

# --- remote scriptblocks (executed on the target PC)
$sbPrepare = {
    param($Exe, $BackupDir)
    $root = [IO.Path]::GetPathRoot($BackupDir)
    [pscustomobject]@{
        ExeFound  = (Test-Path -LiteralPath $Exe)
        DirExists = (Test-Path -LiteralPath $BackupDir)
        DriveOk   = (Test-Path -LiteralPath $root)
        FreeBytes = $(try { ([IO.DriveInfo]$root).AvailableFreeSpace } catch { $null })
        TempDir   = $env:TEMP
    }
}
$sbEnsureDir = {
    param($Dir)
    if (-not (Test-Path -LiteralPath $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
}
$sbSize = {
    param($Path)
    (Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
}
$sbBackup = {
    param($Exe, $Src, $Archive, $SkipTemp)
    $parent = Split-Path $Src -Parent
    $leaf   = Split-Path $Src -Leaf
    Set-Location -LiteralPath $parent
    $a = @('a', '-t7z', '-mx=3', '-mmt=2', '-snl', '-bso0', '-bsp0', $Archive, $leaf)
    if ($SkipTemp) {
        $a += "-xr!$leaf\AppData\Local\Temp"
        $a += "-xr!$leaf\AppData\Local\Microsoft\Windows\INetCache"
        $a += "-xr!$leaf\AppData\Local\Google\Chrome\User Data\*\Cache"
        $a += "-xr!$leaf\AppData\Local\Microsoft\Edge\User Data\*\Cache"
    }
    $out = & $Exe @a 2>&1
    [pscustomobject]@{ Code = $LASTEXITCODE; Output = (($out | Select-Object -Last 8) -join "`r`n") }
}
$sbTest = {
    param($Exe, $Archive)
    $out = & $Exe t -bso0 -bsp0 $Archive 2>&1
    $code = $LASTEXITCODE
    $size = $null
    if (Test-Path -LiteralPath $Archive) { $size = (Get-Item -LiteralPath $Archive).Length }
    [pscustomobject]@{ Code = $code; Size = $size; Output = (($out | Select-Object -Last 8) -join "`r`n") }
}
$sbExists   = { param($Path) Test-Path -LiteralPath $Path }
$sbMkTemp   = { $d = Join-Path $env:TEMP ('7z_' + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $d | Out-Null; $d }
$sbRmTemp   = { param($Path) Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue }
$sbRmFile   = { param($Path) Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue }
$sbSignals  = {
    # Collects every timestamp we can use as a "last activity" signal, ON the remote PC (no SMB needed)
    param($Items)
    $u32 = { param($v) [BitConverter]::ToUInt32([BitConverter]::GetBytes([int32]$v), 0) }
    foreach ($i in $Items) {
        $nt = $null; $uc = $null; $un = $null
        try { $nt = (Get-Item -LiteralPath (Join-Path $i.Path 'NTUSER.DAT') -Force -ErrorAction Stop).LastWriteTime } catch { }
        try { $uc = (Get-Item -LiteralPath (Join-Path $i.Path 'AppData\Local\Microsoft\Windows\UsrClass.dat') -Force -ErrorAction Stop).LastWriteTime } catch { }
        try {
            $k = Get-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($i.SID)" -ErrorAction Stop
            if ($k.LocalProfileUnloadTimeHigh -or $k.LocalProfileUnloadTimeLow) {
                $ft = ([int64](& $u32 $k.LocalProfileUnloadTimeHigh) -shl 32) -bor [int64](& $u32 $k.LocalProfileUnloadTimeLow)
                if ($ft -gt 0) { $un = [DateTime]::FromFileTime($ft) }
            }
        } catch { }
        [pscustomobject]@{ SID = $i.SID; NTUSER = $nt; UsrClass = $uc; Unload = $un }
    }
}

function Set-Busy {
    param([bool]$Busy)
    $btnScan.Enabled = -not $Busy
    $btnRun.Enabled  = -not $Busy
    $btnCancel.Enabled = $Busy
    $form.Cursor = if ($Busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
}

function Format-Size($b) { if ($null -eq $b) { '?' } elseif ($b -ge 1GB) { '{0:N1} GB' -f ($b / 1GB) } else { '{0:N0} MB' -f ($b / 1MB) } }

function Format-Date($d) { if ($d) { ([datetime]$d).ToString('yyyy-MM-dd HH:mm') } else { '-' } }

function Get-AdLastLogon {
    # Domain accounts only: lastLogonTimestamp from AD (replicates every ~14 days, fine for a 90-day test)
    param([string]$Sid)
    try {
        $s = [adsisearcher]"(objectSid=$Sid)"
        [void]$s.PropertiesToLoad.Add('lastlogontimestamp')
        $r = $s.FindOne()
        if ($r -and $r.Properties['lastlogontimestamp'].Count) { return [DateTime]::FromFileTime([int64]$r.Properties['lastlogontimestamp'][0]) }
    } catch { }
    return $null
}

function Get-LastActivity {
    # Returns @{ Time = <datetime or $null>; Source = <text> } according to the chosen basis
    param([string]$Basis, $LastUse, $Ntuser, $UsrClass, $Unload, $Ad)
    switch -Wildcard ($Basis) {
        'NTUSER*' {
            if ($Ntuser)   { return @{ Time = $Ntuser;   Source = 'NTUSER.DAT' } }
            if ($UsrClass) { return @{ Time = $UsrClass; Source = 'UsrClass.dat' } }
            if ($LastUse)  { return @{ Time = $LastUse;  Source = 'LastUseTime (fallback)' } }
            return @{ Time = $null; Source = 'none' }
        }
        'LastUseTime*' {
            if ($LastUse) { return @{ Time = $LastUse; Source = 'LastUseTime' } }
            return @{ Time = $null; Source = 'none' }
        }
        'AD*' {
            if ($Ad)     { return @{ Time = $Ad;     Source = 'AD lastLogon' } }
            if ($Ntuser) { return @{ Time = $Ntuser; Source = 'NTUSER.DAT (no AD data)' } }
            return @{ Time = $null; Source = 'none' }
        }
        default {   # newest of everything = safest
            $best = $null; $src = 'none'
            foreach ($c in @(@($LastUse, 'LastUseTime'), @($Ntuser, 'NTUSER.DAT'), @($UsrClass, 'UsrClass.dat'), @($Unload, 'Profile unload'), @($Ad, 'AD lastLogon'))) {
                if ($c[0] -and (-not $best -or $c[0] -gt $best)) { $best = $c[0]; $src = $c[1] }
            }
            return @{ Time = $best; Source = $src }
        }
    }
}

# ----------------------------------------------------------------- scan
function Start-Scan {
    $computer = $txtComputer.Text.Trim()
    if (-not $computer) { [void][System.Windows.Forms.MessageBox]::Show('Enter a computer name.'); return }

    $script:Cancel = $false
    Set-Busy $true
    try {
        $grid.Rows.Clear()
        $script:ProfileInfo = @{}
        $threshold = [int]$numDays.Value
        $exclude   = $txtExclude.Text.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }

        Write-Log "Scanning $computer for profiles idle more than $threshold days..."
        $session  = Get-CimConn -Computer $computer
        $profiles = Get-CimInstance -CimSession $session -ClassName Win32_UserProfile -ErrorAction Stop |
                    Where-Object { -not $_.Special -and $_.LocalPath }

        $profiles | ForEach-Object { $_.LocalPath.Substring(0, 1) } | Select-Object -Unique |
            ForEach-Object { Connect-Share -Computer $computer -DriveLetter $_ }

        $basis = [string]$cmbBasis.SelectedItem
        $useAD = $chkAD.Checked

        # Collect timestamps ON the remote PC in one call (more reliable than reading C$ file by file)
        $signals = @{}
        try {
            Get-PSConn -Computer $computer
            $items = @($profiles | ForEach-Object { @{ SID = $_.SID; Path = $_.LocalPath } })
            $res = Invoke-Remote -Script $sbSignals -ArgList @(, $items)
            foreach ($r in @($res)) { $signals[$r.SID] = $r }
            Write-Log "Collected profile timestamps remotely for $($signals.Count) profile(s)."
        } catch {
            Write-Log "Remote timestamp collection failed ($($_.Exception.Message.Trim())). Falling back to the C$ share (NTUSER.DAT only)." 'WARN'
        }
        if ($script:Cancel) { return }

        Write-Log "Staleness basis: $basis"
        $now = Get-Date
        foreach ($p in $profiles) {
            [System.Windows.Forms.Application]::DoEvents()
            $name = Split-Path $p.LocalPath -Leaf
            $unc  = ConvertTo-UncPath -Computer $computer -LocalPath $p.LocalPath

            $lastUse = $p.LastUseTime
            $sig = $signals[$p.SID]
            $ntuser = $null; $usrcls = $null; $unload = $null
            if ($sig) { $ntuser = $sig.NTUSER; $usrcls = $sig.UsrClass; $unload = $sig.Unload }
            else { try { $ntuser = (Get-Item -LiteralPath (Join-Path $unc 'NTUSER.DAT') -Force -ErrorAction Stop).LastWriteTime } catch { } }

            $adLogon = $null
            if (($useAD -or $basis -like 'AD*') -and $p.SID -like 'S-1-5-21-*') { $adLogon = Get-AdLastLogon -Sid $p.SID }

            $act  = Get-LastActivity -Basis $basis -LastUse $lastUse -Ntuser $ntuser -UsrClass $usrcls -Unload $unload -Ad $adLogon
            $last = $act.Time
            $days = if ($last) { [int]($now - $last).TotalDays } else { $null }

            $select = $false
            if ($p.Loaded)                    { $status = 'Loaded / in use - skipped' }
            elseif ($exclude -contains $name) { $status = 'Excluded by name' }
            elseif ($null -eq $days)          { $status = 'No date found - review manually' }
            elseif ($days -gt $threshold)     { $status = 'STALE'; $select = $true }
            else                              { $status = 'Active' }

            $script:ProfileInfo[$p.SID] = @{ Name = $name; LocalPath = $p.LocalPath; Unc = $unc }

            [void]$grid.Rows.Add(
                $select, $name, $p.SID,
                (Format-Date $lastUse),
                (Format-Date $ntuser),
                $(if ($null -ne $days) { $days } else { '?' }),
                $status,
                (Format-Date $usrcls), (Format-Date $unload), (Format-Date $adLogon), $act.Source)

            $row = $grid.Rows[$grid.Rows.Count - 1]
            if ($status -eq 'STALE') { $row.DefaultCellStyle.BackColor = [Drawing.Color]::MistyRose }
            elseif ($status -ne 'Active') { $row.DefaultCellStyle.ForeColor = [Drawing.Color]::Gray }
            if ($status -ne 'STALE' -and $status -ne 'Active') { $row.Cells[0].ReadOnly = $true }
        }
        $stale = @($grid.Rows | Where-Object { $_.Cells[6].Value -eq 'STALE' }).Count
        Write-Log "Scan complete: $($grid.Rows.Count) profiles, $stale stale."
    } catch {
        Write-Log "Scan failed: $($_.Exception.Message)" 'ERROR'
    } finally {
        Set-Busy $false
    }
}

# ----------------------------------------------------------------- run (backup on remote PC + delete)
function Start-Cleanup {
    $computer  = $txtComputer.Text.Trim()
    $dry       = $chkDry.Checked
    $remote7z  = $txtRemote7z.Text.Trim()
    $local7z   = $txtLocal7z.Text.Trim()
    $backup    = $txtBackup.Text.Trim().TrimEnd('\')

    $selected = @($grid.Rows | Where-Object { $_.Cells[0].Value -eq $true })
    if ($selected.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('No profiles selected. Run a scan and tick the profiles to process.'); return }
    if ($backup -notmatch '^[A-Za-z]:\\.+') { [void][System.Windows.Forms.MessageBox]::Show('Backup folder must be a local path on the remote PC, e.g. D:\ProfileBackups'); return }
    foreach ($row in $selected) {
        $lp = $script:ProfileInfo[$row.Cells[2].Value].LocalPath
        if ($backup -like "$lp*") { [void][System.Windows.Forms.MessageBox]::Show("The backup folder cannot be inside a profile that will be deleted ($lp)."); return }
    }

    if (-not $dry) {
        $msg = "LIVE RUN on $computer`r`n`r`n$($selected.Count) profile(s) will be backed up on that PC to:`r`n$backup`r`n`r`nand then PERMANENTLY DELETED.`r`n`r`nContinue?"
        if ([System.Windows.Forms.MessageBox]::Show($msg, 'Confirm deletion', 'YesNo', 'Warning', 'Button2') -ne 'Yes') { return }
    }

    $script:Cancel = $false
    Set-Busy $true
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logBase = if ($PSScriptRoot) { $PSScriptRoot } else { $env:TEMP }
    $localLog = Join-Path $logBase "ProfileCleanup_${computer}_$stamp.log"
    if (-not $dry) { $script:LogFile = $localLog }
    try {
        $mode = if ($dry) { 'DRY RUN (no changes will be made)' } else { 'LIVE' }
        Write-Log "=== Starting $mode on $computer - $($selected.Count) profile(s) ==="

        $cim = Get-CimConn -Computer $computer
        Get-PSConn -Computer $computer

        # ---- prepare: check 7-Zip + backup drive on the remote PC
        $prep = Invoke-Remote -Script $sbPrepare -ArgList @($remote7z, $backup)
        if ($null -eq $prep) { return }
        if (-not $prep.DriveOk) { throw "Drive for '$backup' does not exist on $computer." }
        Write-Log ("Backup target on {0}: {1}  (free space: {2}){3}" -f $computer, $backup, (Format-Size $prep.FreeBytes), $(if ($prep.DirExists) { '' } else { '  [folder will be created]' }))

        $exe = $remote7z
        if (-not $prep.ExeFound) {
            if ($local7z -and (Test-Path -LiteralPath $local7z)) {
                if ($dry) {
                    Write-Log "7-Zip not found at '$remote7z' on $computer. A live run would copy '$local7z' to its temp folder." 'WARN'
                } else {
                    $script:RemoteTemp = Invoke-Remote -Script $sbMkTemp -ArgList @()
                    Write-Log "7-Zip not found on $computer; copying '$local7z' to $($script:RemoteTemp)..."
                    Copy-Item -LiteralPath $local7z -Destination $script:RemoteTemp -ToSession $script:PSSess -ErrorAction Stop
                    $dll = Join-Path (Split-Path $local7z -Parent) '7z.dll'
                    if (Test-Path -LiteralPath $dll) { Copy-Item -LiteralPath $dll -Destination $script:RemoteTemp -ToSession $script:PSSess -ErrorAction Stop }
                    $exe = Join-Path $script:RemoteTemp (Split-Path $local7z -Leaf)
                }
            } else {
                throw "7-Zip not found at '$remote7z' on $computer, and no valid local 7z/7za.exe was given to copy over."
            }
        }
        if (-not $dry) { Invoke-Remote -Script $sbEnsureDir -ArgList @($backup) | Out-Null }

        $skipTemp = $chkTemp.Checked
        $ok = 0; $failed = 0
        $totalEst = 0

        foreach ($row in $selected) {
            if ($script:Cancel) { Write-Log 'Cancelled by user.' 'WARN'; break }
            $sid  = $row.Cells[2].Value
            $info = $script:ProfileInfo[$sid]
            $name = $info.Name
            $archive = '{0}\{1}_{2}_{3}.7z' -f $backup, $computer, $name, $stamp

            # Safety re-check: profile may have been loaded since the scan
            $cur = Get-CimInstance -CimSession $cim -ClassName Win32_UserProfile -Filter "SID='$sid'" -ErrorAction SilentlyContinue
            if (-not $cur)   { Write-Log "[$name] profile no longer exists - skipped." 'WARN'; continue }
            if ($cur.Loaded) { Write-Log "[$name] profile is loaded (user logged on) - skipped." 'WARN'; $row.Cells[6].Value = 'Skipped (loaded)'; continue }

            if ($dry) {
                Write-Log "[DRY RUN] [$name] measuring $($info.LocalPath) on $computer ..."
                $size = Invoke-Remote -Script $sbSize -ArgList @($info.LocalPath)
                if ($script:Cancel) { Write-Log 'Cancelled by user.' 'WARN'; break }
                $totalEst += [double]$size
                Write-Log "[DRY RUN] [$name] size on disk: $(Format-Size $size)"
                Write-Log "[DRY RUN] [$name] would run 7-Zip on $computer: $($info.LocalPath)  ->  $archive"
                Write-Log "[DRY RUN] [$name] would test the archive, then delete the profile (SID $sid)"
                $row.Cells[6].Value = 'Dry run OK'
                $ok++
                continue
            }

            # ---- backup (runs on the remote PC)
            Write-Log "[$name] backing up on $computer -> $archive ..."
            $r = Invoke-Remote -Script $sbBackup -ArgList @($exe, $info.LocalPath, $archive, $skipTemp) -KillMatch $archive
            if ($null -eq $r) {
                Write-Log "[$name] cancelled during backup; removing partial archive." 'WARN'
                Invoke-Command -Session $script:PSSess -ScriptBlock $sbRmFile -ArgumentList $archive -ErrorAction SilentlyContinue
                $row.Cells[6].Value = 'Cancelled'
                break
            }
            if ($r.Code -ne 0) {
                # 1 = warning (usually locked/unreadable files). Never delete on anything but a clean backup.
                Write-Log "[$name] 7z exit code $($r.Code) - backup NOT clean. Profile kept.`r`n    $($r.Output)" 'ERROR'
                $row.Cells[6].Value = "Backup failed ($($r.Code))"; $failed++; continue
            }

            # ---- verify (also on the remote PC)
            Write-Log "[$name] verifying archive..."
            $t = Invoke-Remote -Script $sbTest -ArgList @($exe, $archive) -KillMatch $archive
            if ($null -eq $t) { Write-Log "[$name] cancelled during verify. Profile kept." 'WARN'; $row.Cells[6].Value = 'Cancelled'; break }
            if ($t.Code -ne 0) {
                Write-Log "[$name] archive failed integrity test (code $($t.Code)). Profile kept.`r`n    $($t.Output)" 'ERROR'
                $row.Cells[6].Value = 'Verify failed'; $failed++; continue
            }
            Write-Log "[$name] archive OK ($(Format-Size $t.Size)): $archive"

            # ---- delete
            try {
                Write-Log "[$name] deleting profile..."
                $cur | Remove-CimInstance -ErrorAction Stop
                $left = Invoke-Remote -Script $sbExists -ArgList @($info.LocalPath)
                if ($left) {
                    Write-Log "[$name] profile registration removed, but some files remain at $($info.LocalPath)." 'WARN'
                    $row.Cells[6].Value = 'Deleted (leftovers)'
                } else {
                    Write-Log "[$name] profile deleted."
                    $row.Cells[6].Value = 'Deleted'
                }
                $ok++
            } catch {
                Write-Log "[$name] delete failed: $($_.Exception.Message)" 'ERROR'
                $row.Cells[6].Value = 'Delete failed'; $failed++
            }
        }

        if ($dry -and $totalEst -gt 0) {
            $msg = "Dry run: total uncompressed size of selected profiles = $(Format-Size $totalEst); free space on target drive = $(Format-Size $prep.FreeBytes)."
            if ($prep.FreeBytes -and $totalEst -gt $prep.FreeBytes) { Write-Log "$msg Archives are compressed, but this may NOT fit." 'WARN' } else { Write-Log $msg }
        }
        Write-Log "=== Finished ($mode): $ok succeeded, $failed failed ==="
    } catch {
        Write-Log "Run failed: $($_.Exception.Message)" 'ERROR'
    } finally {
        if ($script:RemoteTemp -and $script:PSSess) {
            try { Invoke-Command -Session $script:PSSess -ScriptBlock $sbRmTemp -ArgumentList $script:RemoteTemp -ErrorAction SilentlyContinue } catch { }
            $script:RemoteTemp = $null
        }
        # keep a copy of the log next to the archives on the remote PC
        if (-not $dry -and $script:PSSess -and (Test-Path -LiteralPath $localLog)) {
            try { Copy-Item -LiteralPath $localLog -Destination $backup -ToSession $script:PSSess -ErrorAction Stop } catch { }
        }
        if (-not $dry) { Write-Log "Local log: $localLog" }
        $script:LogFile = $null
        Set-Busy $false
    }
}

# ----------------------------------------------------------------- GUI
$form = New-Object Windows.Forms.Form
$form.Text = 'Stale Profile Cleanup (remote 7-Zip)'
$form.Size = New-Object Drawing.Size(980, 780)
$form.MinimumSize = New-Object Drawing.Size(900, 680)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object Drawing.Font('Segoe UI', 9)

function New-Label($text, $x, $y, $w = 120) {
    $l = New-Object Windows.Forms.Label; $l.Text = $text; $l.Location = "$x,$y"; $l.Size = "$w,22"; $l.TextAlign = 'MiddleLeft'; $form.Controls.Add($l); $l
}

# Row 1: computer / credentials / days
[void](New-Label 'Remote computer:' 12 14)
$txtComputer = New-Object Windows.Forms.TextBox; $txtComputer.Location = '190,14'; $txtComputer.Size = '170,22'; $form.Controls.Add($txtComputer)
$btnCred = New-Object Windows.Forms.Button; $btnCred.Text = 'Credentials...'; $btnCred.Location = '370,12'; $btnCred.Size = '110,26'; $form.Controls.Add($btnCred)
$lblCred = New-Label '(current user)' 485 14 150
[void](New-Label 'Idle more than (days):' 645 14 130)
$numDays = New-Object Windows.Forms.NumericUpDown; $numDays.Location = '780,14'; $numDays.Size = '70,22'; $numDays.Minimum = 1; $numDays.Maximum = 3650; $numDays.Value = 90; $form.Controls.Add($numDays)

# Row 2: 7-Zip on remote PC
[void](New-Label '7-Zip path on remote PC:' 12 46 175)
$txtRemote7z = New-Object Windows.Forms.TextBox; $txtRemote7z.Location = '190,46'; $txtRemote7z.Size = '650,22'; $txtRemote7z.Anchor = 'Top,Left,Right'
$txtRemote7z.Text = 'C:\Program Files\7-Zip\7z.exe'; $form.Controls.Add($txtRemote7z)

# Row 3: local fallback
[void](New-Label 'Local 7za.exe (fallback):' 12 78 175)
$txtLocal7z = New-Object Windows.Forms.TextBox; $txtLocal7z.Location = '190,78'; $txtLocal7z.Size = '650,22'; $txtLocal7z.Anchor = 'Top,Left,Right'
$txtLocal7z.Text = (@("$env:ProgramFiles\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1)
$form.Controls.Add($txtLocal7z)
$btnLocal7z = New-Object Windows.Forms.Button; $btnLocal7z.Text = 'Browse...'; $btnLocal7z.Location = '850,77'; $btnLocal7z.Size = '105,26'; $btnLocal7z.Anchor = 'Top,Right'; $form.Controls.Add($btnLocal7z)

# Row 4: backup folder ON THE REMOTE PC
[void](New-Label 'Backup folder ON REMOTE PC:' 12 110 175)
$txtBackup = New-Object Windows.Forms.TextBox; $txtBackup.Location = '190,110'; $txtBackup.Size = '650,22'; $txtBackup.Anchor = 'Top,Left,Right'; $txtBackup.Text = 'C:\ProfileBackups'; $form.Controls.Add($txtBackup)

# Row 5: exclusions
[void](New-Label 'Never touch (names):' 12 142 175)
$txtExclude = New-Object Windows.Forms.TextBox; $txtExclude.Location = '190,142'; $txtExclude.Size = '765,22'; $txtExclude.Anchor = 'Top,Left,Right'
$txtExclude.Text = 'Administrator,Public,Default,Default User,All Users,defaultuser0'; $form.Controls.Add($txtExclude)

# Row 6: options
$chkDry = New-Object Windows.Forms.CheckBox; $chkDry.Text = 'DRY RUN (no backup, no delete)'; $chkDry.Location = '190,174'; $chkDry.Size = '260,24'; $chkDry.Checked = $true; $chkDry.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold); $form.Controls.Add($chkDry)
$chkTemp = New-Object Windows.Forms.CheckBox; $chkTemp.Text = 'Skip temp / browser cache folders in backup'; $chkTemp.Location = '460,174'; $chkTemp.Size = '320,24'; $chkTemp.Checked = $true; $form.Controls.Add($chkTemp)

# Row 7: buttons
$btnScan = New-Object Windows.Forms.Button; $btnScan.Text = '1. Scan'; $btnScan.Location = '12,206'; $btnScan.Size = '110,30'; $form.Controls.Add($btnScan)
$btnRun  = New-Object Windows.Forms.Button; $btnRun.Text = '2. Run selected'; $btnRun.Location = '130,206'; $btnRun.Size = '140,30'; $form.Controls.Add($btnRun)
$btnCancel = New-Object Windows.Forms.Button; $btnCancel.Text = 'Cancel'; $btnCancel.Location = '278,206'; $btnCancel.Size = '90,30'; $btnCancel.Enabled = $false; $form.Controls.Add($btnCancel)

[void](New-Label 'Idle time based on:' 385 210 110)
$cmbBasis = New-Object Windows.Forms.ComboBox; $cmbBasis.Location = '497,209'; $cmbBasis.Size = '270,24'; $cmbBasis.DropDownStyle = 'DropDownList'
[void]$cmbBasis.Items.AddRange(@('NTUSER.DAT (recommended)', 'Newest of all signals (safest)', 'LastUseTime only (unreliable)', 'AD lastLogon (domain users)'))
$cmbBasis.SelectedIndex = 0
$form.Controls.Add($cmbBasis)
$chkAD = New-Object Windows.Forms.CheckBox; $chkAD.Text = 'Also query AD'; $chkAD.Location = '777,208'; $chkAD.Size = '150,26'; $form.Controls.Add($chkAD)

# Grid
$grid = New-Object Windows.Forms.DataGridView
$grid.Location = '12,244'; $grid.Size = '943,240'; $grid.Anchor = 'Top,Left,Right'
$grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'; $grid.AutoSizeColumnsMode = 'Fill'; $grid.BackgroundColor = [Drawing.Color]::White
$colSel = New-Object Windows.Forms.DataGridViewCheckBoxColumn; $colSel.HeaderText = 'Select'; $colSel.FillWeight = 40
[void]$grid.Columns.Add($colSel)
foreach ($c in @(@('User', 90), @('SID', 150), @('LastUseTime (WMI)', 90), @('NTUSER.DAT modified', 90), @('Days idle', 50), @('Status', 120), @('UsrClass.dat', 90), @('Profile unload', 90), @('AD lastLogon', 90), @('Basis used', 90))) {
    $col = New-Object Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]; $col.ReadOnly = $true
    [void]$grid.Columns.Add($col)
}
$form.Controls.Add($grid)
$grid.Columns[6].DisplayIndex = $grid.Columns.Count - 1   # keep Status as the last visible column

# Log
$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Location = '12,492'; $txtLog.Size = '943,240'; $txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
$txtLog.Font = New-Object Drawing.Font('Consolas', 9); $txtLog.BackColor = [Drawing.Color]::White
$form.Controls.Add($txtLog)

# ----------------------------------------------------------------- events
$btnCred.Add_Click({
    $c = Get-Credential -Message 'Admin account for the remote PC (Cancel = use current Windows login)'
    $script:Cred = $c
    $lblCred.Text = if ($c) { $c.UserName } else { '(current user)' }
})
$btnLocal7z.Add_Click({
    $d = New-Object Windows.Forms.OpenFileDialog; $d.Filter = '7z.exe / 7za.exe|7z*.exe|All files|*.*'
    if ($d.ShowDialog() -eq 'OK') { $txtLocal7z.Text = $d.FileName }
})
$chkDry.Add_CheckedChanged({
    $btnRun.Text = if ($chkDry.Checked) { '2. Run (dry run)' } else { '2. Run (LIVE!)' }
    $btnRun.ForeColor = if ($chkDry.Checked) { [Drawing.Color]::Black } else { [Drawing.Color]::Firebrick }
})
$btnScan.Add_Click({ Start-Scan })
$btnRun.Add_Click({ Start-Cleanup })
$btnCancel.Add_Click({ $script:Cancel = $true; Write-Log 'Cancel requested...' 'WARN' })
$form.Add_FormClosing({
    Disconnect-Shares
    if ($script:PSSess)     { Remove-PSSession $script:PSSess -ErrorAction SilentlyContinue }
    if ($script:CimSession) { Remove-CimSession $script:CimSession -ErrorAction SilentlyContinue }
})

$chkDry.Checked = $false; $chkDry.Checked = $true   # fire the handler to set the button text
Write-Log 'Ready. Enter a computer name and click Scan. DRY RUN is enabled by default.'
[void]$form.ShowDialog()
