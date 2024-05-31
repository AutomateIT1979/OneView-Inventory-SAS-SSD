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
                foreach ($drive in $localStorageDetails.Data.PhysicalDrives) {
                    # Extract the necessary information
                    $info = [PSCustomObject]@{
                        # Server information (from the server object) and local storage details (from localStorageDetails)
                        ApplianceFQDN              = $fqdn
                        ServerName                 = $server.Name -split ', ' | Select-Object -First 1
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
                        # Calculate the logical capacity in GB
                        LogicalCapacityGB = [math]::Round(($drive.CapacityLogicalBlocks * $drive.BlockSizeBytes) / 1e9, 2)
                        DriveEncryptedDrive        = $drive.EncryptedDrive
                        DriveFirmwareVersion       = $drive.FirmwareVersion.Current.VersionString
                        DriveInterfaceType         = $drive.InterfaceType
                        DriveMediaType             = $drive.MediaType
                        DriveLocation              = $drive.Location
                        DriveModel                 = $drive.Model
                        # Get the drive serial number
                        DriveSerialNumber          = $drive.SerialNumber
                        # Get the drive status (health)
                        DriveStatus                = $drive.Status.Health
                        # Get the drive state
                        DriveState                 = $drive.Status.State
                        # Show the remaining life of the SSD in percentage
                        "Drive Life Remaining (%)" = "{0}%" -f (100 - $drive.SSDEnduranceUtilizationPercentage)
                    }
                    # Add the collected information to the data list
                    $data.Add($info)
                }
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
    }
}
# Disconnect after all servers for a given appliance have been processed
Disconnect-OVMgmt
# Sort the collected data by ApplianceFQDN and BayNumber
$sortedData = $data | Sort-Object -Property ApplianceFQDN, BayNumber -Descending
# Export the collected data to a CSV file
$sortedData | Export-Csv -Path (Join-Path -Path $scriptPath -ChildPath "LocalStorageDetails.csv") -NoTypeInformation
# Export the collected data to an Excel file
$sortedData | Export-Excel -Path (Join-Path -Path $scriptPath -ChildPath "LocalStorageDetails.xlsx") -AutoSize
Write-Output "Audit completed and data exported to LocalStorageDetails.csv and LocalStorageDetails.xlsx"
