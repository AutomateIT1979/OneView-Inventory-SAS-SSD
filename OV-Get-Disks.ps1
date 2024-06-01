# Import required modules and show import progress
Write-Host "Importing HPE OneView module..." -ForegroundColor Yellow
Import-Module HPEOneView.660 -Verbose

Write-Host "Importing ImportExcel module..." -ForegroundColor Yellow
Import-Module ImportExcel -Verbose

# Define script paths and file names
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$csvPath = Join-Path $scriptPath "Appliances_liste.csv"
$appliancesList = Import-Csv $csvPath

# Define function for error handling and logging
function Write-ErrorLog {
    param([string]$errorMessage)
    Write-Host $errorMessage -ForegroundColor Red
    $errorMessage | Add-Content $logFile
}

# Initialize data collection and error log file path
$dataCollection = [System.Collections.Generic.List[object]]::new()
$logFile = Join-Path $scriptPath "error_log.txt"

# Check if log file exists, create it if not
if (-not (Test-Path $logFile)) {
    New-Item -Path $logFile -ItemType File -Force
}

# Connect to all appliances and gather information
foreach ($appliance in $appliancesList) {
    $fqdn = $appliance.Appliance_FQDN

    try {
        Write-Host "Connecting to OneView appliance: $fqdn" -ForegroundColor Green
        $connectedSession = Connect-OVMgmt -Hostname $fqdn -Credential (Get-Credential -Message "Enter OneView credentials")

        $servers = Get-OVServer -Connection $connectedSession | Where-Object { $_.model -match 'Gen10' }

        foreach ($server in $servers) {
            $localStorageUri = $server.uri + '/localStorage'
            $localStorageDetails = Send-OVRequest -uri $localStorageUri

            if ($localStorageDetails) {
                foreach ($drive in $localStorageDetails.Data.PhysicalDrives) {
                    $info = [PSCustomObject]@{
                        ApplianceFQDN              = $connectedSession.Name
                        ServerName                 = $server.serverName.ToUpper()
                        Name                       = ($server.Name -split ', ')[0]
                        BayNumber                  = ($server.Name -split ', ')[-1]
                        ServerStatus               = $server.Status
                        ServerPower                = $server.PowerState
                        ProcessorCoreCount         = $server.ProcessorCoreCount
                        ProcessorCount             = $server.ProcessorCount
                        ProcessorSpeedMhz          = $server.ProcessorSpeedMhz
                        ProcessorType              = $server.ProcessorType
                        ServerSerialNumber         = $server.SerialNumber
                        ServerModel                = $server.Model
                        AdapterType                = $localStorageDetails.Data.AdapterType
                        CurrentOperatingMode       = $localStorageDetails.Data.CurrentOperatingMode
                        FirmwareVersion            = $localStorageDetails.Data.FirmwareVersion.Current.VersionString
                        InternalPortCount          = $localStorageDetails.Data.InternalPortCount
                        Location                   = $localStorageDetails.Data.Location
                        LocationFormat             = $localStorageDetails.Data.LocationFormat
                        LogicalDriveNumbers        = ($localStorageDetails.Data.LogicalDrives.LogicalDriveNumber) -join ', '
                        RaidValues                 = ($localStorageDetails.Data.LogicalDrives.Raid) -join ', '
                        Model                      = $localStorageDetails.Data.Model
                        DriveBlockSizeBytes        = $drive.BlockSizeBytes
                        LogicalCapacityGB          = [math]::Round(($drive.CapacityLogicalBlocks * $drive.BlockSizeBytes) / 1e9, 2)
                        DriveEncryptedDrive        = $drive.EncryptedDrive
                        DriveFirmwareVersion       = $drive.FirmwareVersion.Current.VersionString
                        DriveInterfaceType         = $drive.InterfaceType
                        DriveMediaType             = $drive.MediaType
                        DriveLocation              = $drive.Location
                        DriveModel                 = $drive.Model
                        DriveSerialNumber          = $drive.SerialNumber
                        DriveStatus                = $drive.Status.Health
                        DriveState                 = $drive.Status.State
                        "Drive Life Remaining (%)" = "{0}%" -f (100 - $drive.SSDEnduranceUtilizationPercentage)
                    }
                    $dataCollection.Add($info)
                }
            }
        }
    }
    catch {
        Write-ErrorLog "Error connecting to appliance ${fqdn}: $($_.Exception.Message)"
    }
    finally {
        Write-Host "Disconnecting from OneView appliance: $fqdn" -ForegroundColor Green
        Disconnect-OVMgmt -Hostname $fqdn
    }
}

# Sorting and exporting data to CSV and Excel
$sortedData = $dataCollection | Sort-Object -Property ApplianceFQDN, BayNumber -Descending

# Export data to CSV file (append mode)
$sortedData | Export-Csv -Path "$scriptPath\LocalStorageDetails.csv" -NoTypeInformation -Append

# Export data to Excel file (append mode)
$sortedData | Export-Excel -Path "$scriptPath\LocalStorageDetails.xlsx" -Show -AutoSize -Append

# Display completion message to the user with the path to the exported files
Write-Host "`t• Data collection completed. The data has been exported to the following files:"
Write-Host "`t• CSV file: $scriptPath\LocalStorageDetails.csv" -ForegroundColor Green
Write-Host "`t• Excel file: $scriptPath\LocalStorageDetails.xlsx" -ForegroundColor Green
