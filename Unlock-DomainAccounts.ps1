#Requires -Version 5.1
<#
.SYNOPSIS
    GUI tool that lists locked-out Active Directory accounts and unlocks them.

.DESCRIPTION
    - Shows every account that is locked out RIGHT NOW (not just accounts that have ever been locked),
      using the same flag Active Directory itself uses, so accounts whose lockout has already expired
      are not listed.
    - By default it queries the PDC emulator, which has the most up-to-date lockout information.
      Type a domain controller name to use a specific one.
    - Unlock the selected accounts, or every account currently shown (after filtering).
    - "Find lockout source" searches the DC's Security log for event 4740 and shows which computer
      the bad passwords came from.
    - Optional auto-refresh, live filter, and export to CSV.
    - Every unlock is written to UnlockAccounts_<date>.log next to the script (who unlocked what, and when).

    No ActiveDirectory (RSAT) module is needed: it uses System.DirectoryServices.

    Requirements
      - Run from a domain-joined PC (or give a DC name and credentials).
      - To unlock: rights to unlock accounts (Domain Admins, Account Operators, or a delegated
        "Reset password / write lockoutTime" permission on the OUs).
      - For "Find lockout source": permission to read the Security log of the DC
        (Event Log Readers or admin), with remote event log access allowed through its firewall.
        Event 4740 is only logged if "Audit User Account Management" is enabled on the DCs
        (it is by default).
#>

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.DirectoryServices
[System.Windows.Forms.Application]::EnableVisualStyles()

# ----------------------------------------------------------------- state
$script:Cred    = $null
$script:Server  = $null          # DC actually used for the last refresh
$script:Busy    = $false
$script:LastRefresh = $null
$logBase = if ($PSScriptRoot) { $PSScriptRoot } else { $env:TEMP }
$script:AuditFile = Join-Path $logBase ('UnlockAccounts_{0:yyyyMMdd}.log' -f (Get-Date))

# ----------------------------------------------------------------- helpers
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO', [switch]$Audit)
    $line = '{0}  [{1}]  {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    if ($Audit -or $Level -eq 'ERROR') {
        try { Add-Content -LiteralPath $script:AuditFile -Value $line -ErrorAction Stop } catch { }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function New-DirEntry {
    param([string]$Path)
    if ($script:Cred) {
        New-Object System.DirectoryServices.DirectoryEntry($Path, $script:Cred.UserName, $script:Cred.GetNetworkCredential().Password)
    } else {
        New-Object System.DirectoryServices.DirectoryEntry($Path)
    }
}

function Get-First {
    param($SearchResult, [string]$Name)
    if ($SearchResult.Properties[$Name].Count -gt 0) { $SearchResult.Properties[$Name][0] }
}

function ConvertFrom-AdTime {
    # AD large-integer time -> DateTime, or $null for "never" / not set
    param($Value)
    if ($null -eq $Value) { return $null }
    try { $n = [int64]$Value } catch { return $null }
    if ($n -le 0 -or $n -ge 2650467743999999999) { return $null }
    return [DateTime]::FromFileTime($n)
}

function Get-DbValue { param($Value) if ($null -eq $Value) { [DBNull]::Value } else { $Value } }

function Get-LockoutDuration {
    # Domain lockout duration (negative 100ns ticks; Int64.MinValue = locked until an admin unlocks). Only used as a fallback.
    param($Root)
    try {
        $s = New-Object System.DirectoryServices.DirectorySearcher($Root, '(objectClass=domainDNS)', [string[]]@('lockoutduration'), 'Base')
        $r = $s.FindOne()
        if ($r -and $r.Properties['lockoutduration'].Count) { return [int64]$r.Properties['lockoutduration'][0] }
    } catch { }
    return $null
}

function Get-LockedAccounts {
    $server = $txtServer.Text.Trim()
    if (-not $server -and -not $script:Cred) {
        try { $server = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().PdcRoleOwner.Name }
        catch { Write-Log "Could not find the PDC emulator ($($_.Exception.Message.Trim())); using the default domain controller." 'WARN' }
    }
    $script:Server = $server
    $sp = if ($server) { "$server/" } else { '' }

    $rootDse = New-DirEntry "LDAP://${sp}RootDSE"
    $nc = [string]$rootDse.Properties['defaultNamingContext'][0]
    if (-not $nc) { throw 'Could not read the domain naming context.' }
    $root = New-DirEntry "LDAP://${sp}$nc"

    $searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
    $searcher.PageSize = 1000
    # candidates: accounts that have a lockout time; the real "locked right now" check is done below
    $searcher.Filter = '(&(objectCategory=person)(objectClass=user)(lockoutTime>=1))'
    foreach ($p in 'displayname', 'cn', 'samaccountname', 'lockouttime', 'badpwdcount', 'badpasswordtime',
                   'lastlogontimestamp', 'useraccountcontrol', 'msds-user-account-control-computed',
                   'pwdlastset', 'description', 'distinguishedname', 'adspath') {
        [void]$searcher.PropertiesToLoad.Add($p)
    }

    $durationLoaded = $false
    $duration = $null
    $results = $searcher.FindAll()
    try {
        foreach ($r in $results) {
            $lockTime = [int64](Get-First $r 'lockouttime')

            # Authoritative: constructed attribute, bit 0x10 = account is locked out right now
            $computed = Get-First $r 'msds-user-account-control-computed'
            if ($null -ne $computed) {
                $locked = (([int]$computed -band 0x10) -ne 0)
            } else {
                # Fallback (approximate): compare lockout time with the domain lockout duration
                if (-not $durationLoaded) { $duration = Get-LockoutDuration $root; $durationLoaded = $true }
                $locked = ($null -eq $duration) -or ($duration -eq [int64]::MinValue) -or ($duration -eq 0) -or
                          (($lockTime + [math]::Abs($duration)) -gt (Get-Date).ToFileTime())
            }
            if (-not $locked) { continue }

            $uac = Get-First $r 'useraccountcontrol'
            $dn  = [string](Get-First $r 'distinguishedname')
            $display = [string](Get-First $r 'displayname')
            if (-not $display) { $display = [string](Get-First $r 'cn') }

            [pscustomobject]@{
                DisplayName = $display
                Username    = [string](Get-First $r 'samaccountname')
                LockedSince = ConvertFrom-AdTime $lockTime
                BadPwdCount = [int](Get-First $r 'badpwdcount')
                LastBadPwd  = ConvertFrom-AdTime (Get-First $r 'badpasswordtime')
                LastLogon   = ConvertFrom-AdTime (Get-First $r 'lastlogontimestamp')
                Enabled     = (-not (([int]$uac) -band 2))
                PwdLastSet  = ConvertFrom-AdTime (Get-First $r 'pwdlastset')
                Description = [string](Get-First $r 'description')
                Location    = ($dn -replace '^(?:\\.|[^,\\])+,', '')
                AdsPath     = [string](Get-First $r 'adspath')
            }
        }
    } finally {
        $results.Dispose()
        $searcher.Dispose()
    }
}

function Set-Busy {
    param([bool]$IsBusy)
    $script:Busy = $IsBusy
    foreach ($b in $btnRefresh, $btnUnlock, $btnUnlockAll, $btnSource, $btnExport) { $b.Enabled = -not $IsBusy }
    $form.Cursor = if ($IsBusy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
}

function Update-Status {
    if ($null -eq $script:LastRefresh) { return }
    $dc = if ($script:Server) { $script:Server } else { '(default DC)' }
    $lblStatus.Text = '{0} locked account(s), {1} shown   |   DC: {2}   |   refreshed {3:HH:mm:ss}' -f $table.Rows.Count, $view.Count, $dc, $script:LastRefresh
}

function Update-LockedList {
    if ($script:Busy) { return }
    Set-Busy $true
    try {
        # remember the selection so an auto-refresh does not drop it
        $keep = @($grid.SelectedRows | ForEach-Object { $_.DataBoundItem.Row['Username'] })

        $accounts = @(Get-LockedAccounts)

        $table.BeginLoadData()
        $table.Clear()
        foreach ($a in $accounts) {
            $row = $table.NewRow()
            $row['DisplayName'] = Get-DbValue $a.DisplayName
            $row['Username']    = Get-DbValue $a.Username
            $row['LockedSince'] = Get-DbValue $a.LockedSince
            $row['BadPwdCount'] = Get-DbValue $a.BadPwdCount
            $row['LastBadPwd']  = Get-DbValue $a.LastBadPwd
            $row['LastLogon']   = Get-DbValue $a.LastLogon
            $row['Enabled']     = Get-DbValue $a.Enabled
            $row['PwdLastSet']  = Get-DbValue $a.PwdLastSet
            $row['Description'] = Get-DbValue $a.Description
            $row['Location']    = Get-DbValue $a.Location
            $row['AdsPath']     = Get-DbValue $a.AdsPath
            $table.Rows.Add($row)
        }
        $table.EndLoadData()

        $grid.ClearSelection()
        foreach ($gr in $grid.Rows) {
            if ($keep -contains $gr.DataBoundItem.Row['Username']) { $gr.Selected = $true }
        }

        $script:LastRefresh = Get-Date
        Update-Status
        if ($accounts.Count -eq 0) { $lblStatus.Text = 'No locked accounts.   ' + $lblStatus.Text }
    } catch {
        $lblStatus.Text = 'Refresh failed - see the log below.'
        Write-Log "Refresh failed: $($_.Exception.Message)" 'ERROR'
    } finally {
        Set-Busy $false
    }
}

function Unlock-Rows {
    param($Rows)
    $Rows = @($Rows)
    if ($Rows.Count -eq 0) { return }

    $by = if ($script:Cred) { $script:Cred.UserName } else { [Security.Principal.WindowsIdentity]::GetCurrent().Name }
    $dc = if ($script:Server) { $script:Server } else { '(default DC)' }
    $ok = 0; $fail = 0

    Set-Busy $true
    try {
        foreach ($rv in $Rows) {
            $user = [string]$rv.Row['Username']
            $path = [string]$rv.Row['AdsPath']
            $de = $null
            try {
                $de = New-DirEntry $path
                $de.Properties['lockoutTime'].Value = 0          # 0 = unlocked
                $de.CommitChanges()

                # verify with the same flag AD uses
                $de.RefreshCache([string[]]@('msDS-User-Account-Control-Computed'))
                $flag = $de.Properties['msDS-User-Account-Control-Computed']
                if ($flag.Count -gt 0 -and (([int]$flag[0] -band 0x10) -ne 0)) {
                    Write-Log "$user : unlock was accepted but the account still shows as locked." 'WARN' -Audit
                    $fail++
                } else {
                    Write-Log "UNLOCKED $user (by $by, on $dc)" 'INFO' -Audit
                    $ok++
                }
            } catch {
                Write-Log "Could not unlock $user : $($_.Exception.Message.Trim())" 'ERROR'
                $fail++
            } finally {
                if ($de) { $de.Dispose() }
            }
        }
    } finally {
        Set-Busy $false
    }
    Write-Log "Unlock finished: $ok unlocked, $fail failed."
    Update-LockedList
}

function Find-LockoutSource {
    $rows = @($grid.SelectedRows | ForEach-Object { $_.DataBoundItem })
    if ($rows.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('Select one or more accounts first.'); return }
    if (-not $script:Server) { [void][System.Windows.Forms.MessageBox]::Show('Refresh the list first so the tool knows which domain controller to ask.'); return }

    $hours = [int]$numHours.Value
    $ms = [int64]$hours * 3600000
    Set-Busy $true
    try {
        foreach ($rv in $rows) {
            $user = [string]$rv.Row['Username']
            if ($user -match "'") { Write-Log "$user : name contains a quote, skipped." 'WARN'; continue }
            Write-Log "[$user] searching the Security log on $($script:Server) for lockout events (last $hours h)..."

            $xpath = "*[System[(EventID=4740) and TimeCreated[timediff(@SystemTime) <= $ms]] and EventData[Data[@Name='TargetUserName']='$user']]"
            $p = @{ ComputerName = $script:Server; LogName = 'Security'; FilterXPath = $xpath; MaxEvents = 20; ErrorAction = 'Stop' }
            if ($script:Cred) { $p.Credential = $script:Cred }

            try {
                $events = @(Get-WinEvent @p)
            } catch {
                if ($_.FullyQualifiedErrorId -like '*NoMatchingEventsFound*' -or $_.Exception.Message -like '*No events were found*') {
                    $events = @()
                } else {
                    Write-Log "[$user] could not read the event log: $($_.Exception.Message.Trim())" 'ERROR'
                    continue
                }
            }

            if ($events.Count -eq 0) {
                Write-Log "[$user] no lockout (4740) events found on $($script:Server). The lockout may have been logged on a different DC, or it is older than $hours h."
                continue
            }
            foreach ($ev in $events) {
                $xml = [xml]$ev.ToXml()
                $caller = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetDomainName' }).'#text'
                Write-Log ('[{0}] locked out at {1:yyyy-MM-dd HH:mm:ss} - caller computer: {2}' -f $user, $ev.TimeCreated, $caller)
            }
        }
    } finally {
        Set-Busy $false
    }
}

function Set-RowFilter {
    $t = $txtFilter.Text.Trim()
    if (-not $t) {
        $view.RowFilter = ''
    } else {
        $e = $t.Replace("'", "''").Replace('[', '[[]').Replace('%', '[%]').Replace('*', '[*]')
        $view.RowFilter = "DisplayName LIKE '%$e%' OR Username LIKE '%$e%' OR Description LIKE '%$e%'"
    }
    Update-Status
}

# ----------------------------------------------------------------- data
$table = New-Object System.Data.DataTable 'Locked'
foreach ($c in @(@('DisplayName', [string]), @('Username', [string]), @('LockedSince', [datetime]), @('BadPwdCount', [int]),
                 @('LastBadPwd', [datetime]), @('LastLogon', [datetime]), @('Enabled', [bool]), @('PwdLastSet', [datetime]),
                 @('Description', [string]), @('Location', [string]), @('AdsPath', [string]))) {
    [void]$table.Columns.Add($c[0], $c[1])
}
$view = $table.DefaultView
$view.Sort = 'LockedSince DESC'

# ----------------------------------------------------------------- GUI
$form = New-Object Windows.Forms.Form
$form.Text = 'Locked Domain Accounts'
$form.Size = New-Object Drawing.Size(1000, 720)
$form.MinimumSize = New-Object Drawing.Size(900, 560)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object Drawing.Font('Segoe UI', 9)

function New-Label($text, $x, $y, $w = 120) {
    $l = New-Object Windows.Forms.Label; $l.Text = $text; $l.Location = "$x,$y"; $l.Size = "$w,22"; $l.TextAlign = 'MiddleLeft'; $form.Controls.Add($l); $l
}

# Row 1: DC / credentials / refresh
[void](New-Label 'Domain controller (blank = PDC):' 12 14 200)
$txtServer = New-Object Windows.Forms.TextBox; $txtServer.Location = '215,14'; $txtServer.Size = '190,22'; $form.Controls.Add($txtServer)
$btnCred = New-Object Windows.Forms.Button; $btnCred.Text = 'Credentials...'; $btnCred.Location = '415,12'; $btnCred.Size = '110,26'; $form.Controls.Add($btnCred)
$lblCred = New-Label '(current user)' 532 14 190
$btnRefresh = New-Object Windows.Forms.Button; $btnRefresh.Text = 'Refresh'; $btnRefresh.Location = '845,12'; $btnRefresh.Size = '110,26'; $btnRefresh.Anchor = 'Top,Right'; $form.Controls.Add($btnRefresh)

# Row 2: filter / auto-refresh
[void](New-Label 'Filter:' 12 46 60)
$txtFilter = New-Object Windows.Forms.TextBox; $txtFilter.Location = '75,46'; $txtFilter.Size = '230,22'; $form.Controls.Add($txtFilter)
$chkAuto = New-Object Windows.Forms.CheckBox; $chkAuto.Text = 'Auto-refresh every'; $chkAuto.Location = '330,46'; $chkAuto.Size = '130,24'; $form.Controls.Add($chkAuto)
$numAuto = New-Object Windows.Forms.NumericUpDown; $numAuto.Location = '462,47'; $numAuto.Size = '60,22'; $numAuto.Minimum = 10; $numAuto.Maximum = 3600; $numAuto.Value = 30; $form.Controls.Add($numAuto)
[void](New-Label 'seconds' 526 46 60)

# Row 3: actions
$btnUnlock = New-Object Windows.Forms.Button; $btnUnlock.Text = 'Unlock selected'; $btnUnlock.Location = '12,78'; $btnUnlock.Size = '150,30'; $form.Controls.Add($btnUnlock)
$btnUnlockAll = New-Object Windows.Forms.Button; $btnUnlockAll.Text = 'Unlock all shown'; $btnUnlockAll.Location = '170,78'; $btnUnlockAll.Size = '150,30'; $form.Controls.Add($btnUnlockAll)
$btnSource = New-Object Windows.Forms.Button; $btnSource.Text = 'Find lockout source'; $btnSource.Location = '328,78'; $btnSource.Size = '170,30'; $form.Controls.Add($btnSource)
[void](New-Label 'look back (hours):' 506 82 110)
$numHours = New-Object Windows.Forms.NumericUpDown; $numHours.Location = '618,83'; $numHours.Size = '60,22'; $numHours.Minimum = 1; $numHours.Maximum = 720; $numHours.Value = 24; $form.Controls.Add($numHours)
$btnExport = New-Object Windows.Forms.Button; $btnExport.Text = 'Export to CSV'; $btnExport.Location = '845,78'; $btnExport.Size = '110,30'; $btnExport.Anchor = 'Top,Right'; $form.Controls.Add($btnExport)

# Grid
$grid = New-Object Windows.Forms.DataGridView
$grid.Location = '12,118'; $grid.Size = '943,335'; $grid.Anchor = 'Top,Bottom,Left,Right'
$grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.ReadOnly = $true; $grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $true
$grid.AutoSizeColumnsMode = 'Fill'; $grid.BackgroundColor = [Drawing.Color]::White
# Columns are defined explicitly (not auto-generated) so they exist no matter when the grid is bound.
# The AdsPath column of the table is deliberately not shown; it is read from the row when unlocking.
$grid.AutoGenerateColumns = $false
foreach ($d in @(
        @('DisplayName', 'Display name',      100, 'text'),
        @('Username',    'Username',           80, 'text'),
        @('LockedSince', 'Locked since',       80, 'date'),
        @('BadPwdCount', 'Bad pwd count',      45, 'text'),
        @('LastBadPwd',  'Last bad password',  80, 'date'),
        @('LastLogon',   'Last logon',         80, 'date'),
        @('Enabled',     'Enabled',            40, 'check'),
        @('PwdLastSet',  'Password last set',  80, 'date'),
        @('Description', 'Description',       110, 'text'),
        @('Location',    'Location (OU)',     140, 'text'))) {
    $col = if ($d[3] -eq 'check') { New-Object Windows.Forms.DataGridViewCheckBoxColumn } else { New-Object Windows.Forms.DataGridViewTextBoxColumn }
    $col.Name = $d[0]; $col.DataPropertyName = $d[0]; $col.HeaderText = $d[1]; $col.FillWeight = $d[2]; $col.ReadOnly = $true
    if ($d[3] -eq 'date') { $col.DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm' }
    [void]$grid.Columns.Add($col)
}
$form.Controls.Add($grid)
$grid.DataSource = $view

# Status + log
$lblStatus = New-Object Windows.Forms.Label; $lblStatus.Location = '12,458'; $lblStatus.Size = '943,22'; $lblStatus.Anchor = 'Bottom,Left,Right'; $lblStatus.TextAlign = 'MiddleLeft'
$lblStatus.Text = 'Click Refresh to load the locked accounts.'
$form.Controls.Add($lblStatus)

$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Location = '12,484'; $txtLog.Size = '943,190'; $txtLog.Anchor = 'Bottom,Left,Right'
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
$txtLog.Font = New-Object Drawing.Font('Consolas', 9); $txtLog.BackColor = [Drawing.Color]::White
$form.Controls.Add($txtLog)

$timer = New-Object Windows.Forms.Timer
$timer.Interval = 30000

# ----------------------------------------------------------------- events
$btnCred.Add_Click({
    $c = Get-Credential -Message 'Account to use for the AD query and to unlock accounts (Cancel = current Windows login)'
    $script:Cred = $c
    $lblCred.Text = if ($c) { $c.UserName } else { '(current user)' }
})
$btnRefresh.Add_Click({ Update-LockedList })
$txtFilter.Add_TextChanged({ Set-RowFilter })

$chkAuto.Add_CheckedChanged({
    $timer.Interval = [int]$numAuto.Value * 1000
    $timer.Enabled = $chkAuto.Checked
})
$numAuto.Add_ValueChanged({ $timer.Interval = [int]$numAuto.Value * 1000 })
$timer.Add_Tick({ Update-LockedList })

$btnUnlock.Add_Click({
    $rows = @($grid.SelectedRows | ForEach-Object { $_.DataBoundItem })
    if ($rows.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('Select one or more accounts first.'); return }
    if ($rows.Count -gt 1) {
        $r = [System.Windows.Forms.MessageBox]::Show("Unlock $($rows.Count) selected accounts?", 'Confirm unlock', 'YesNo', 'Question')
        if ($r -ne 'Yes') { return }
    }
    Unlock-Rows $rows
})
$btnUnlockAll.Add_Click({
    $rows = @($view | ForEach-Object { $_ })
    if ($rows.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('There are no accounts to unlock.'); return }
    $r = [System.Windows.Forms.MessageBox]::Show("Unlock ALL $($rows.Count) account(s) currently shown?", 'Confirm unlock', 'YesNo', 'Warning', 'Button2')
    if ($r -ne 'Yes') { return }
    Unlock-Rows $rows
})
$btnSource.Add_Click({ Find-LockoutSource })

$btnExport.Add_Click({
    if ($view.Count -eq 0) { return }
    $d = New-Object Windows.Forms.SaveFileDialog; $d.Filter = 'CSV|*.csv'; $d.FileName = ('LockedAccounts_{0:yyyyMMdd-HHmm}.csv' -f (Get-Date))
    if ($d.ShowDialog() -eq 'OK') {
        $out = foreach ($rv in $view) {
            [pscustomobject][ordered]@{
                DisplayName = $rv['DisplayName']; Username = $rv['Username']; LockedSince = $rv['LockedSince']
                BadPwdCount = $rv['BadPwdCount']; LastBadPassword = $rv['LastBadPwd']; LastLogon = $rv['LastLogon']
                Enabled = $rv['Enabled']; PasswordLastSet = $rv['PwdLastSet']; Description = $rv['Description']; Location = $rv['Location']
            }
        }
        $out | Export-Csv -LiteralPath $d.FileName -NoTypeInformation -Encoding UTF8
        Write-Log "Exported $($view.Count) account(s) to $($d.FileName)"
    }
})

$form.Add_Shown({ Update-LockedList })
$form.Add_FormClosing({ $timer.Stop(); $timer.Dispose() })

Write-Log "Ready. Unlock actions are recorded in $($script:AuditFile)"
[void]$form.ShowDialog()
