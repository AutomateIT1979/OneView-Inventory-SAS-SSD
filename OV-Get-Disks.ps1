# Import the appropriate HPEOneView module (choose 660 or 800)
Import-Module HPEOneView.800
Import-Module ImportExcel

# Get the full path of the current script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Path to the CSV file containing appliance FQDNs (relative to the script)
$csvPath = Join-Path -Path $scriptPath -ChildPath "Appliances_liste.csv"

# Read the appliances list from the CSV file
$appliances = Import-Csv -Path $csvPath

# Prompt the user for OneView credentials
$credential = Get-Credential -Message "Enter OneView credentials"

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
            # Construct the URI for local storage details (corrected)
            $localStorageUri = "${server.uri}/localStorage"

            # Retrieve the local storage details
            $localStorageDetails = Send-OVRequest -Uri $localStorageUri -Method GET

            # Check if localStorageDetails is not null
            if ($null -ne $localStorageDetails) {
                # Extract the necessary information
                $info = [PSCustomObject]@{
                    AdapterType               = $localStorageDetails.AdapterType
                    BackupPowerSourceStatus   = $localStorageDetails.BackupPowerSourceStatus
                    CacheMemorySizeMiB        = $localStorageDetails.CacheMemorySizeMiB
                    CurrentOperatingMode      = $localStorageDetails.CurrentOperatingMode
                    ExternalPortCount         = $localStorageDetails.ExternalPortCount
                    FirmwareVersion           = $localStorageDetails.FirmwareVersion.Current
                    InternalPortCount         = $localStorageDetails.InternalPortCount
                    Location                  = $localStorageDetails.Location
                    LocationFormat            = $localStorageDetails.LocationFormat
                    Model                     = $localStorageDetails.Model
                    Name                      = $localStorageDetails.Name
                    PhysicalDrives            = ($localStorageDetails.PhysicalDrives | ForEach-Object {
                        "{$_.BlockSizeBytes},{$_.CapacityLogicalBlocks},{$_.CapacityMiB},{$_.EncryptedDrive},{$_.FirmwareVersion},{$_.Location},{$_.Model},{$_.SerialNumber},{$_.Status}"
                    }) -join ','
                    SerialNumber               = $localStorageDetails.SerialNumber
                    Status                    = $localStorageDetails.Status
                }

                # Add the collected information to the data array
                $data += $info
            }
        } 
    }
    catch {
        $errorMessage = "Error processing appliance ${fqdn}: $($_.Exception.Message)"  # Corrected error message
        Write-Warning $errorMessage
        $errorMessage | Add-Content -Path $logFile
    }
    finally {
        # Always disconnect, even if an error occurs
        Disconnect-OVMgmt
    }
}

# Export the collected data to a CSV file
$data | Export-Csv -Path (Join-Path -Path $scriptPath -ChildPath "LocalStorageDetails.csv") -NoTypeInformation

# Export the collected data to an Excel file
$data | Export-Excel -Path (Join-Path -Path $scriptPath -ChildPath "LocalStorageDetails.xlsx") -AutoSize

Write-Output "Audit completed and data exported to LocalStorageDetails.csv and LocalStorageDetails.xlsx"
