# Import necessary PowerShell modules
Import-Module HPEOneView.660
Import-Module HPEOneView.800
Import-Module ImportExcel

# Path to the CSV file containing appliance FQDNs (assuming it's in the same folder as the script)
$scriptPath = $PSScriptRoot
$csvPath = Join-Path -Path $scriptPath -ChildPath "Appliances_liste.csv"

# Read the appliances list from the CSV file
$appliances = Import-Csv -Path $csvPath

# Prompt the user for OneView credentials
$credential = Get-Credential -Message "Enter OneView credentials"

# Initialize an array to hold the collected data
$data = @()

# Loop through each appliance and retrieve the required information
foreach ($appliance in $appliances) {
    $fqdn = $appliance.Appliance_FQDN

    # Connect to the OneView appliance using the provided credentials
    Connect-OVMgmt -Hostname $fqdn -Credential $credential

    # Get the server objects for Gen10 servers
    $servers = Get-OVServer | Where-Object { $_.model -match 'Gen10' }

    foreach ($server in $servers) {
        # Construct the URI for local storage details
        $localStorageUri = $server.uri + '/localStorage'

        # Retrieve the local storage details
        $localStorageDetails = Send-OVRequest -Uri $localStorageUri -Method GET

        # Extract the necessary information
        $info = [PSCustomObject]@{
            AdapterType                                   = $localStorageDetails.AdapterType
            BackupPowerSourceStatus                       = $localStorageDetails.BackupPowerSourceStatus
            CacheMemorySizeMiB                            = $localStorageDetails.CacheMemorySizeMiB
            CurrentOperatingMode                          = $localStorageDetails.CurrentOperatingMode
            ExternalPortCount                             = $localStorageDetails.ExternalPortCount
            FirmwareVersion                               = $localStorageDetails.FirmwareVersion.Current
            InternalPortCount                             = $localStorageDetails.InternalPortCount
            Location                                      = $localStorageDetails.Location
            LocationFormat                                = $localStorageDetails.LocationFormat
            Model                                         = $localStorageDetails.Model
            Name                                          = $localStorageDetails.Name
            PhysicalDrives                                = ($localStorageDetails.PhysicalDrives | ForEach-Object {
                [PSCustomObject]@{
                    BlockSizeBytes       = $_.BlockSizeBytes
                    CapacityLogicalBlocks = $_.CapacityLogicalBlocks
                    CapacityMiB          = $_.CapacityMiB
                    EncryptedDrive       = $_.EncryptedDrive
                    FirmwareVersion      = $_.FirmwareVersion
                    Location             = $_.Location
                    Model                = $_.Model
                    SerialNumber         = $_.SerialNumber
                    Status               = $_.Status
                }
            })
            SerialNumber                                  = $localStorageDetails.SerialNumber
            Status                                        = $localStorageDetails.Status
        }

        # Add the collected information to the data array
        $data += $info
    }

    # Disconnect from the OneView appliance
    Disconnect-OVMgmt
}

# Export the collected data to a CSV file
$data | Export-Csv -Path "LocalStorageDetails.csv" -NoTypeInformation

# Export the collected data to an Excel file
$data | Export-Excel -Path "LocalStorageDetails.xlsx" -AutoSize

Write-Output "Audit completed and data exported to LocalStorageDetails.csv and LocalStorageDetails.xlsx"
