Import-Module HPEOneView.660
Import-Module ImportExcel

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$csvPath = Join-Path $scriptPath "Appliances_liste.csv"
$appliances = Import-Csv $csvPath

$credential = Get-Credential -Message "Enter OneView credentials"
$data = [System.Collections.Generic.List[object]]::new()
$logFile = Join-Path $scriptPath "error_log.txt"

foreach ($appliance in $appliances) {
    $fqdn = $appliance.Appliance_FQDN

    try {
        # Connect to the OneView appliance
        Connect-OVMgmt -Hostname $fqdn -Credential $credential

        $servers = Get-OVServer | Where-Object { $_.model -match 'Gen10' }

        foreach ($server in $servers) {
            $localStorageUri = $server.uri + '/localStorage'
            $localStorageDetails = Send-OVRequest -uri $localStorageUri
           
            if ($localStorageDetails) {
                foreach ($drive in $localStorageDetails.Data.PhysicalDrives) {
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
                        LogicalCapacityGB          = [math]::Round(($drive.CapacityLogicalBlocks * $drive.BlockSizeBytes) / 1e9, 2)
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
                    $data.Add($info) 
                }
            }
        }
    }
    catch {
        $errorMessage = "Error processing appliance ${fqdn}: $($_.Exception.Message)"
        Write-Warning $errorMessage
        $errorMessage | Add-Content $logFile
    }
    finally {
        # Always disconnect after processing each appliance
        Disconnect-OVMgmt
    } 
}

$sortedData = $data | Sort-Object -Property ApplianceFQDN, BayNumber -Descending
$sortedData | Export-Csv (Join-Path $scriptPath "LocalStorageDetails.csv") -NoTypeInformation
$sortedData | Export-Excel (Join-Path $scriptPath "LocalStorageDetails.xlsx") -AutoSize
Write-Output "Audit completed and data exported to LocalStorageDetails.csv and LocalStorageDetails.xlsx"
