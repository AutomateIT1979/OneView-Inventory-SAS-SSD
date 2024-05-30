# Import the appropriate HPEOneView module (choose 400, 660, or 800)
Import-Module HPEOneView.660
Import-Module ImportExcel
# Get the full path of the current script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
# Path to the CSV file containing appliance FQDNs (relative to the script)
$csvPath = Join-Path -Path $scriptPath -ChildPath "Appliances_liste.csv"
# Read the appliances list from the CSV file
$appliances = Import-Csv -Path $csvPath
# Prompt the user for OneView credentials
$credential = Get-Credential -Message "Enter OneView credentials"
# Initialize a list to hold the collected data
$data = New-Object System.Collections.Generic.List[object]
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
            $localStorageUri = $server.uri + '/localStorage'  
            # Retrieve the local storage details (using Send-OVRequest)
            $localStorageDetails = Send-OVRequest -uri $localStorageUri
            # Check if localStorageDetails is not null
            if ($null -ne $localStorageDetails) {
                # Extract the necessary information
                $info = [PSCustomObject]@{
                    ApplianceFQDN             = $fqdn
                    AdapterType               = $localStorageDetails.Data.AdapterType                 
                    CacheMemorySizeMiB        = $localStorageDetails.Data.CacheMemorySizeMiB
                    CurrentOperatingMode      = $localStorageDetails.Data.CurrentOperatingMode
                    FirmwareVersion           = $localStorageDetails.Data.FirmwareVersion.Current
                    InternalPortCount         = $localStorageDetails.Data.InternalPortCount
                    Location                  = $localStorageDetails.Data.Location
                    LocationFormat            = $localStorageDetails.Data.LocationFormat
                    Model                     = $localStorageDetails.Data.Model
                    Name                      = $localStorageDetails.Data.Name
                    PhysicalDrives            = ($localStorageDetails.Data.PhysicalDrives | ForEach-Object {
                        "{$_.BlockSizeBytes},{$_.CapacityLogicalBlocks},{$_.CapacityMiB},{$_.EncryptedDrive},{$_.FirmwareVersion},{$_.Location},{$_.Model},{$_.SerialNumber},{$_.Status}"
                    }) -join ','
                    SerialNumber               = $localStorageDetails.Data.SerialNumber
                }

                $info | Add-Member -NotePropertyName Health -NotePropertyValue $localStorageDetails.Data.Status.Health
                $info | Add-Member -NotePropertyName State -NotePropertyValue $localStorageDetails.Data.Status.State
                # Add the collected information to the data list
                $data.Add($info)
            }
        } 
    }
    catch {
        $errorMessage = "Error processing appliance ${fqdn}: $($_.Exception.Message)" 
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
