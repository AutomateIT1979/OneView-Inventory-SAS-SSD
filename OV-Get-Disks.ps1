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
                foreach ($drive in $localStorageDetails.PhysicalDrives) {
                    $info = [PSCustomObject]@{
                        ApplianceFQDN             = $fqdn
                        Servername                = $server.name.Trim()  # Remove extra spaces
                        AdapterType               = $localStorageDetails.AdapterType
                        CurrentOperatingMode      = $localStorageDetails.CurrentOperatingMode
                        ExternalPortCount         = $localStorageDetails.ExternalPortCount
                        FirmwareVersion           = $localStorageDetails.FirmwareVersion.Current
                        InternalPortCount         = $localStorageDetails.InternalPortCount
                        Location                  = $localStorageDetails.Location
                        LocationFormat            = $localStorageDetails.LocationFormat
                        Model                     = $localStorageDetails.Model
                        Name                      = $localStorageDetails.Name
                        SerialNumber              = $localStorageDetails.SerialNumber
                        Status                    = $localStorageDetails.Status
                        Drive_BlockSizeBytes       = $drive.BlockSizeBytes
                        Drive_CapacityLogicalBlocks = $drive.CapacityLogicalBlocks
                        Drive_CapacityMiB          = $drive.CapacityMiB
                        Drive_EncryptedDrive       = $drive.EncryptedDrive
                        Drive_FirmwareVersion      = $drive.FirmwareVersion
                        Drive_Location            = $drive.Location
                        Drive_Model               = $drive.Model
                        Drive_SerialNumber        = $drive.SerialNumber
                        Drive_Status              = $drive.Status
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
