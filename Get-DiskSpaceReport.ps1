#Requires -Version 5.1
<#
.SYNOPSIS
    Reports free disk space for every Active Directory computer whose name starts with one of the
    prefixes listed in a CSV file, and saves the report as a CSV.

.DESCRIPTION
    1. Reads computer-name prefixes from a CSV file (one prefix per row).
    2. Searches Active Directory for computer accounts whose name starts with each prefix.
       This uses System.DirectoryServices, so the ActiveDirectory (RSAT) module is NOT required.
    3. Queries every matching computer in parallel for its fixed disks (Win32_LogicalDisk) over WinRM,
       falling back to DCOM, and writes one row per disk to the report.

    Nothing disappears silently: computers that are offline or unreachable get a row with their
    status and error, and prefixes that match no computer get a "No AD match" row.

    Report columns:
      Prefix, ComputerName, DnsHostName, OperatingSystem, AdLastLogon, Status, Drive, VolumeName,
      SizeGB, FreeGB, FreePercent, Warning, Protocol, Error, ScanTime

    Status values: OK, Offline (no ping reply), Error, No fixed disks, No AD match for prefix.
    Warning is "LOW" when FreePercent is below -WarnPercent.

    Requirements: run from a domain-joined PC as an account that is a local administrator on the
    target computers (needed to read disk information remotely), with WinRM or DCOM/WMI reachable.

.PARAMETER InputCsv
    CSV file with the prefixes. By default the column named "Prefix" is used; if there is no such
    column, the first column is used. The first row is treated as a header (see -NoHeader).

.PARAMETER Column
    Name of the column that holds the prefixes (overrides the default).

.PARAMETER NoHeader
    The CSV has no header row: every row, including the first, is a prefix.

.PARAMETER OutputCsv
    Where to save the report. Default: DiskSpaceReport_<timestamp>.csv in the current folder.

.PARAMETER SearchBase
    Limit the AD search to an OU, e.g. 'OU=Workstations,DC=corp,DC=local'. Default: whole domain.

.PARAMETER Server
    Domain controller or domain name to query. Default: the current domain.

.PARAMETER IncludeDisabled
    Also include disabled computer accounts (excluded by default).

.PARAMETER Credential
    Alternate credentials for both the AD search and the disk queries.

.PARAMETER ThrottleLimit
    How many computers to query at the same time. Default 20.

.PARAMETER TimeoutSec
    Per-computer timeout for the remote queries. Default 30.

.PARAMETER WarnPercent
    Disks with less free space than this percentage are flagged "LOW". Default 10.

.PARAMETER SkipPing
    Do not ping first. Use this if ICMP is blocked in your network (otherwise such computers
    would be reported as offline).

.EXAMPLE
    .\Get-DiskSpaceReport.ps1 -InputCsv .\prefixes.csv

.EXAMPLE
    .\Get-DiskSpaceReport.ps1 -InputCsv .\prefixes.csv -OutputCsv C:\Reports\disks.csv `
        -SearchBase 'OU=Workstations,DC=corp,DC=local' -WarnPercent 15

.EXAMPLE
    # CSV is just a plain list of prefixes with no header row
    .\Get-DiskSpaceReport.ps1 -InputCsv .\prefixes.csv -NoHeader -SkipPing
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [string]$InputCsv,

    [string]$Column,
    [switch]$NoHeader,

    [string]$OutputCsv = (Join-Path (Get-Location).Path ('DiskSpaceReport_{0:yyyyMMdd-HHmmss}.csv' -f (Get-Date))),

    [string]$SearchBase,
    [string]$Server,
    [switch]$IncludeDisabled,
    [pscredential]$Credential,

    [ValidateRange(1, 128)][int]$ThrottleLimit = 20,
    [ValidateRange(5, 600)][int]$TimeoutSec = 30,
    [ValidateRange(0, 100)][double]$WarnPercent = 10,
    [switch]$SkipPing
)

$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------- helpers
function New-ReportRow {
    # Every row has the same columns, in this order.
    param([hashtable]$Values)
    $row = [ordered]@{
        Prefix          = ''
        ComputerName    = ''
        DnsHostName     = ''
        OperatingSystem = ''
        AdLastLogon     = ''
        Status          = ''
        Drive           = ''
        VolumeName      = ''
        SizeGB          = $null
        FreeGB          = $null
        FreePercent     = $null
        Warning         = ''
        Protocol        = ''
        Error           = ''
        ScanTime        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    [pscustomobject]$row
}

function New-DirEntry {
    param([string]$Path)
    if ($Credential) {
        New-Object System.DirectoryServices.DirectoryEntry($Path, $Credential.UserName, $Credential.GetNetworkCredential().Password)
    } else {
        New-Object System.DirectoryServices.DirectoryEntry($Path)
    }
}

function Get-FirstValue {
    param($SearchResult, [string]$Name)
    if ($SearchResult.Properties[$Name].Count -gt 0) { $SearchResult.Properties[$Name][0] }
}

function Find-AdComputers {
    # All computer accounts whose name starts with $Prefix
    param([string]$Prefix)

    # escape LDAP filter special characters
    $esc = $Prefix -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29'
    $filter = "(&(objectCategory=computer)(name=$esc*))"
    if (-not $IncludeDisabled) {
        $filter = "(&(objectCategory=computer)(name=$esc*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
    }

    $searcher = New-Object System.DirectoryServices.DirectorySearcher($script:SearchRoot, $filter)
    $searcher.PageSize = 1000
    $searcher.SearchScope = 'Subtree'
    foreach ($p in 'name', 'dnshostname', 'operatingsystem', 'lastlogontimestamp', 'distinguishedname') {
        [void]$searcher.PropertiesToLoad.Add($p)
    }

    $results = $searcher.FindAll()
    try {
        foreach ($r in $results) {
            $lastLogon = ''
            $ll = Get-FirstValue $r 'lastlogontimestamp'
            if ($ll -and [int64]$ll -gt 0) { $lastLogon = [DateTime]::FromFileTime([int64]$ll).ToString('yyyy-MM-dd') }

            [pscustomobject]@{
                Prefix            = $Prefix
                Name              = [string](Get-FirstValue $r 'name')
                DnsHostName       = [string](Get-FirstValue $r 'dnshostname')
                OperatingSystem   = [string](Get-FirstValue $r 'operatingsystem')
                LastLogon         = $lastLogon
                DistinguishedName = [string](Get-FirstValue $r 'distinguishedname')
            }
        }
    } finally {
        $results.Dispose()
        $searcher.Dispose()
    }
}

# --------------------------------------------------------------------------- 1. read the prefixes
if (-not (Test-Path -LiteralPath $InputCsv)) { throw "Input file not found: $InputCsv" }

$lines = @(Get-Content -LiteralPath $InputCsv | Where-Object { $_.Trim() })
if ($lines.Count -eq 0) { throw 'The input file is empty.' }

$rows = @(if ($NoHeader) { $lines | ConvertFrom-Csv -Header 'Prefix' } else { $lines | ConvertFrom-Csv })
if ($rows.Count -eq 0) { throw 'No data rows found after the header row. If the file has no header row, use -NoHeader.' }

$cols = @($rows[0].PSObject.Properties.Name)
if ($Column) {
    $col = $cols | Where-Object { $_ -ieq $Column } | Select-Object -First 1
    if (-not $col) { throw "Column '$Column' not found. Columns in the file: $($cols -join ', ')" }
} elseif ($cols -contains 'Prefix') {
    $col = $cols | Where-Object { $_ -ieq 'Prefix' } | Select-Object -First 1
} else {
    $col = $cols[0]
    Write-Warning "No 'Prefix' column found; using the first column '$col' and treating the first row as a header. If your file has no header row, run again with -NoHeader."
}

$prefixes = @($rows | ForEach-Object { ([string]$_.$col).Trim().TrimEnd('*').Trim() } |
              Where-Object { $_ } | Sort-Object -Unique)
if ($prefixes.Count -eq 0) { throw "No prefixes found in column '$col'." }

foreach ($p in $prefixes | Where-Object { $_.Length -lt 3 }) {
    Write-Warning "Prefix '$p' is very short and may match a large number of computers."
}
Write-Host "Read $($prefixes.Count) prefix(es) from '$InputCsv' (column '$col')."

# --------------------------------------------------------------------------- 2. search Active Directory
Write-Host 'Searching Active Directory...'
$report = New-Object System.Collections.Generic.List[object]
$computers = New-Object System.Collections.Generic.List[object]
$seen = @{}

try {
    $serverPart = if ($Server) { "$Server/" } else { '' }
    if (-not $SearchBase) {
        $rootDse = New-DirEntry "LDAP://${serverPart}RootDSE"
        $SearchBase = [string]$rootDse.Properties['defaultNamingContext'][0]
        if (-not $SearchBase) { throw 'could not read the default naming context.' }
    }
    $script:SearchRoot = New-DirEntry "LDAP://${serverPart}$SearchBase"

    foreach ($prefix in $prefixes) {
        $found = @(Find-AdComputers -Prefix $prefix)
        foreach ($c in $found) {
            # a computer that matches several prefixes is only queried once
            if (-not $seen.ContainsKey($c.DistinguishedName)) {
                $seen[$c.DistinguishedName] = $true
                $computers.Add($c)
            }
        }
        Write-Host ('  {0,-24} {1} computer(s)' -f $prefix, $found.Count)
        if ($found.Count -eq 0) {
            $report.Add((New-ReportRow @{ Prefix = $prefix; Status = 'No AD match for prefix' }))
        }
    }
} catch {
    throw "Active Directory search failed: $($_.Exception.Message)"
}

Write-Host "Found $($computers.Count) computer(s) in total."

# --------------------------------------------------------------------------- 3. query the disks in parallel
# Runs inside a runspace; returns one hashtable per disk (or one status hashtable on failure).
$worker = {
    param($Target, $Cred, $TimeoutSec, $SkipPing, $WarnPercent)
    $ErrorActionPreference = 'Stop'

    if (-not $SkipPing) {
        $alive = $false
        try { $alive = Test-Connection -ComputerName $Target -Count 1 -Quiet -ErrorAction SilentlyContinue } catch { }
        if (-not $alive) { return @{ Status = 'Offline (no ping reply)' } }
    }

    $disks = $null
    $proto = ''
    $errors = @()
    foreach ($p in 'WSMan', 'Dcom') {
        $sess = $null
        try {
            $sp = @{
                ComputerName        = $Target
                SessionOption       = (New-CimSessionOption -Protocol $p)
                OperationTimeoutSec = $TimeoutSec
            }
            if ($Cred) { $sp.Credential = $Cred }
            $sess = New-CimSession @sp
            $disks = @(Get-CimInstance -CimSession $sess -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -OperationTimeoutSec $TimeoutSec)
            $proto = $p
            $errors = @()
            break
        } catch {
            $errors += ('{0}: {1}' -f $p, $_.Exception.Message.Trim())
        } finally {
            if ($sess) { Remove-CimSession $sess -ErrorAction SilentlyContinue }
        }
    }

    if ($null -eq $disks)   { return @{ Status = 'Error'; Error = ($errors -join ' | ') } }
    if ($disks.Count -eq 0) { return @{ Status = 'No fixed disks'; Protocol = $proto } }

    foreach ($d in ($disks | Sort-Object DeviceID)) {
        $size = [double]$d.Size
        $free = [double]$d.FreeSpace
        $pct  = if ($size -gt 0) { [math]::Round(100 * $free / $size, 1) } else { $null }
        @{
            Status      = 'OK'
            Drive       = $d.DeviceID
            VolumeName  = $d.VolumeName
            SizeGB      = [math]::Round($size / 1GB, 2)
            FreeGB      = [math]::Round($free / 1GB, 2)
            FreePercent = $pct
            Warning     = $(if ($null -ne $pct -and $pct -lt $WarnPercent) { 'LOW' } else { '' })
            Protocol    = $proto
        }
    }
}

if ($computers.Count -gt 0) {
    Write-Host "Querying disk space (up to $ThrottleLimit computers at a time)..."
    $pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
    $pool.Open()
    $jobs = @()
    try {
        foreach ($c in $computers) {
            $target = if ($c.DnsHostName) { $c.DnsHostName } else { $c.Name }
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($worker.ToString())
            [void]$ps.AddArgument($target)
            [void]$ps.AddArgument($Credential)
            [void]$ps.AddArgument($TimeoutSec)
            [void]$ps.AddArgument([bool]$SkipPing)
            [void]$ps.AddArgument($WarnPercent)
            $jobs += [pscustomobject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Comp = $c }
        }

        # wait, with a progress bar
        while ($true) {
            $done = @($jobs | Where-Object { $_.Handle.IsCompleted }).Count
            Write-Progress -Activity 'Querying disk space' -Status "$done of $($jobs.Count) computers" `
                -PercentComplete ([int](100 * $done / $jobs.Count))
            if ($done -ge $jobs.Count) { break }
            Start-Sleep -Milliseconds 500
        }
        Write-Progress -Activity 'Querying disk space' -Completed

        # collect
        foreach ($j in $jobs) {
            $c = $j.Comp
            $base = @{
                Prefix          = $c.Prefix
                ComputerName    = $c.Name
                DnsHostName     = $c.DnsHostName
                OperatingSystem = $c.OperatingSystem
                AdLastLogon     = $c.LastLogon
            }
            try {
                $results = @($j.PS.EndInvoke($j.Handle))
            } catch {
                $msg = $_.Exception.Message
                if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
                $results = @(@{ Status = 'Error'; Error = $msg })
            }
            if ($results.Count -eq 0) { $results = @(@{ Status = 'Error'; Error = 'No data returned' }) }

            foreach ($r in $results) {
                $h = $r.psobject.BaseObject
                if ($h -isnot [hashtable]) { continue }
                $values = @{} + $base
                foreach ($k in $h.Keys) { $values[$k] = $h[$k] }
                $report.Add((New-ReportRow $values))
            }
            $j.PS.Dispose()
        }
    } finally {
        $pool.Close()
        $pool.Dispose()
    }
}

# --------------------------------------------------------------------------- 4. save the report
$fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputCsv)
$dir = Split-Path -Parent $fullPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$sorted = $report | Sort-Object @{ Expression = { [string]::IsNullOrEmpty($_.ComputerName) } }, ComputerName, Drive
$sorted | Export-Csv -LiteralPath $fullPath -NoTypeInformation -Encoding UTF8

# --------------------------------------------------------------------------- summary
$okComputers  = @($report | Where-Object { $_.Status -eq 'OK' } | Select-Object -ExpandProperty ComputerName -Unique).Count
$offline      = @($report | Where-Object { $_.Status -like 'Offline*' }).Count
$failed       = @($report | Where-Object { $_.Status -eq 'Error' }).Count
$noMatch      = @($report | Where-Object { $_.Status -like 'No AD match*' }).Count
$lowDisks     = @($report | Where-Object { $_.Warning -eq 'LOW' }).Count

Write-Host ''
Write-Host 'Summary'
Write-Host "  Computers found in AD : $($computers.Count)"
Write-Host "  Queried successfully  : $okComputers"
Write-Host "  Offline               : $offline"
Write-Host "  Errors                : $failed"
Write-Host "  Prefixes with no match: $noMatch"
Write-Host "  Disks below $WarnPercent% free : $lowDisks"
Write-Host ''
Write-Host "Report saved to: $fullPath"
