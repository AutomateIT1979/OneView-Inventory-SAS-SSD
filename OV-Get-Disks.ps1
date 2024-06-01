# Import required modules and show import progress
Write-Host "Importing HPE OneView module..." -ForegroundColor Yellow
Import-Module HPEOneView.660 -Verbose

Write-Host "Importing ImportExcel module..." -ForegroundColor Yellow
Import-Module ImportExcel -Verbose

# Define script paths and file names
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$csvPath = Join-Path $scriptPath "Appliances_liste.csv"
$appliances = Import-Csv $csvPath

# Initialize an array to hold the collected data
$data = @()

# Log file for errors
$logFile = Join-Path -Path $scriptPath -ChildPath "error_log.txt"

# Define credentials for connecting to OneView appliances
$credential = Get-Credential -Message "Enter OneView credentials"

# Loop through each appliance and retrieve the required information
foreach ($appliance in $appliances) {
    $fqdn = $appliance.Appliance_FQDN
    try {
        # Connect to the OneView appliance
        Connect-OVMgmt -Hostname $fqdn -Credential $credential

        # Get the server objects for Gen10 servers
        $servers = Get-OVServer | Where-Object { $_.model -match 'Gen10' }

        foreach ($server in $servers) {
            # Construct the URI for local storage details
            $localStorageUri = $server.uri + '/localStorage'

            # Retrieve the local storage details (using Send-OVRequest)
            $localStorageDetails = Send-OVRequest -Uri $localStorageUri -Method GET

            # Check if localStorageDetails is not null
            if ($null -ne $localStorageDetails) {
                foreach ($drive in $localStorageDetails.Data.PhysicalDrives) {
                    $info = [PSCustomObject]@{
                                                ApplianceFQDN              = $fqdn
                        ServerName                 = $server.Name
                        ServerStatus               = $server.Status
                        ServerPower                = $server.PowerState
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
                        # Calculate the logical capacity in GB
                        LogicalCapacityGB = [math]::Round(($drive.CapacityLogicalBlocks * $drive.BlockSizeBytes) / 1e9, 2)
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

                    # Add the collected information to the data array
                    $data += $info
                }
            }
        }

        Disconnect-OVMgmt -Hostname $fqdn
    }
    catch {
        $errorMessage = "Error processing appliance ${fqdn}: $($_.Exception.Message)"
        Write-Warning $errorMessage
        $errorMessage | Add-Content -Path $logFile
    }
}

# Sorting and exporting data to CSV and Excel
$sortedData = $data | Sort-Object -Property ApplianceFQDN, Servername -Descending

# Export data to CSV file (append mode)
$sortedData | Export-Csv -Path "$scriptPath\LocalStorageDetails.csv" -NoTypeInformation -Append

# Export data to Excel file (append mode)
$sortedData | Export-Excel -Path "$scriptPath\LocalStorageDetails.xlsx" -Show -AutoSize -Append

# Display completion message to the user with the path to the exported files
Write-Host "`t• Data collection completed. The data has been exported to the following files:"
Write-Host "`t• CSV file: $scriptPath\LocalStorageDetails.csv" -ForegroundColor Green
Write-Host "`t• Excel file: $scriptPath\LocalStorageDetails.xlsx" -ForegroundColor Green
