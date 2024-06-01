# Import required modules
Import-Module HPEOneView.660
Import-Module ImportExcel

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

# Connect to all appliances and gather information
foreach ($appliance in $appliancesList) {
    $fqdn = $appliance.Appliance_FQDN

    try {
        Connect-OVMgmt -Hostname $fqdn -Credential (Get-Credential -Message "Enter OneView credentials")
    }
    catch {
        Log-Error "Error connecting to appliance ${fqdn}: $($_.Exception.Message)"
        continue  # Skip to the next appliance if connection fails
    }

    foreach ($connection in $Global:ConnectedSessions) {
        try {
            Set-OVApplianceConnection $connection
            $servers = Get-OVServer | Where-Object { $_.model -match 'Gen10' }

            foreach ($server in $servers) {
                $localStorageUri = $server.uri + '/localStorage'
                $localStorageDetails = Send-OVRequest -uri $localStorageUri

                if ($localStorageDetails) {
                    foreach ($drive in $localStorageDetails.Data.PhysicalDrives) {
                        $info = [PSCustomObject]@{
                            ApplianceFQDN              = $connection.Name
                            ServerName                 = $server.serverName.ToUpper()
                            Name                       = $server.Name -split ', ' | Select-Object -First 1
                            BayNumber                  = $server.Name -split ', ' | Select-Object -Last 1
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
                            LogicalDriveNumbers        = ($localStorageDetails.Data.LogicalDrives | ForEach-Object { $_.LogicalDriveNumber }) -join ', '
                            RaidValues                 = ($localStorageDetails.Data.LogicalDrives | ForEach-Object { $_.Raid }) -join ', '
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
            Write-ErrorLog "Error processing appliance $($connection.Name): $($_.Exception.Message)"
        }
    }
}

# Disconnect from all appliances
Disconnect-OVMgmt

# Sorting and exporting data to CSV and Excel
$sortedData = $dataCollection | Sort-Object -Property ApplianceFQDN, BayNumber -Descending
# Export data to CSV and Excel files
$sortedData | Export-Csv -Path "$scriptPath\LocalStorageDetails.csv" -NoTypeInformation
$sortedData | Export-Excel -Path "$scriptPath\LocalStorageDetails.xlsx" -Show -AutoSize
# Display completion message
Write-Host "Audit completed and data exported to LocalStorageDetails.csv and LocalStorageDetails.xlsx"
