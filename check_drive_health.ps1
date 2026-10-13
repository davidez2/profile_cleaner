# Define the output path for your CSV report
\$OutputPath = "C:\Temp\Network_Disk_Health_Report.csv"

# Ensure the destination folder exists
\(TargetDir = Split-Path\)OutputPath
if (-not (Test-Path \(TargetDir)) { New-Item -ItemType Directory -Path\)TargetDir -Force | Out-Null }

# 1. Fetch active computer objects from Active Directory
# Filters for Windows operating systems and machines that have logged in within the last 90 days to avoid dead accounts
Write-Host "Fetching computer list from Active Directory..." -ForegroundColor Cyan
\(LastLogonThreshold = (Get-Date).AddDays(-90)\)Computers = Get-ADComputer -Filter {OperatingSystem -like "*Windows*" -and Enabled -eq \$true} -Properties LastLogonDate | 
    Where-Object { \(_.LastLogonDate -gt\)LastLogonThreshold } | 
    Select-Object -ExpandProperty Name

Write-Host "Found \((\)Computers.Count) active computers to scan.`n" -ForegroundColor Green

# 2. Iterate through each machine and poll physical disk metrics
$Report = foreach ($Computer in $Computers) {
    Write-Host "Querying $Computer..." -ForegroundColor Yellow
    
    # Test network connectivity first to avoid long timeout delays
    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
        Write-Host "  [-] $Computer is Offline (Ping failed)." -ForegroundColor Red
        [PSCustomObject]@{
            ComputerName      = $Computer
            DeviceID          = "N/A"
            FriendlyName      = "N/A"
            MediaType         = "N/A"
            BusType           = "N/A"
            HealthStatus      = "Offline"
            OperationalStatus = "Ping Failed"
            Temperature_C     = "N/A"
            Wear_Percentage   = "N/A"
        }
        continue
    }

    try {
        # Establish a common CIM session configuration to fetch both physical profiles and telemetry counters
        $CimOptions = New-CimSessionOption -Protocol Dcom
        $CimSession = New-CimSession -ComputerName $Computer -SessionOption $CimOptions -ErrorAction Stop
        
        # Query basic drive infrastructure profile
        $Disks = Get-CimInstance -CimSession $CimSession -ClassName MSFT_PhysicalDisk -Namespace Root\Microsoft\Windows\Storage -ErrorAction Stop
        
        foreach ($Disk in $Disks) {
            # Safely cross-reference the storage reliability counter logs for SMART data elements
            $Reliability = Get-CimInstance -CimSession $CimSession -Namespace Root\Microsoft\Windows\Storage -Query "SELECT Temperature, Wear FROM MSFT_StorageReliabilityCounter WHERE DeviceId='$( $Disk.DeviceId )'" -ErrorAction SilentlyContinue

            [PSCustomObject]@{
                ComputerName      = $Computer
                DeviceID          = $Disk.DeviceId
                FriendlyName      = $Disk.FriendlyName
                MediaType         = if ($Disk.MediaType -eq 3) { "HDD" } elseif ($Disk.MediaType -eq 4) { "SSD" } else { "Unspecified ($($Disk.MediaType))" }
                BusType           = switch ($Disk.BusType) { 17 { "NVMe" } 11 { "SATA" } 8 { "SCSI" } 16 { "SAS" } default { "Other ($($Disk.BusType))" } }
                HealthStatus      = switch ($Disk.HealthStatus) { 0 { "Healthy" } 1 { "Warning" } 2 { "Unhealthy" } default { "Unknown" } }
                OperationalStatus = $Disk.OperationalStatus -join ', '
                Temperature_C     = if ($Reliability.Temperature) { $Reliability.Temperature } else { "N/A" }
                Wear_Percentage   = if ($Reliability.Wear -ne $null) { $Reliability.Wear } else { "N/A" }
            }
        }
    }
    catch {
        Write-Host "  [-] RPC/WMI Query Failed on $Computer. Exception: $_" -ForegroundColor Red
        [PSCustomObject]@{
            ComputerName      = $Computer
            DeviceID          = "N/A"
            FriendlyName      = "N/A"
            MediaType         = "N/A"
            BusType           = "N/A"
            HealthStatus      = "Error"
            OperationalStatus = "WMI/CIM Connection Refused"
            Temperature_C     = "N/A"
            Wear_Percentage   = "N/A"
        }
    }
    finally {
        if ($CimSession) { Remove-CimSession $CimSession }
    }
}

# 3. Export data payload cleanly to a structured CSV
$Report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nScan Complete. Consolidated health log saved to: \$OutputPath" -ForegroundColor Green
