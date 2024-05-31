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
    
    # Check if already connected to this appliance
    $existingConnection = $Global:ConnectedSessions | Where-Object { $_.Name -eq $fqdn }

    if (-not $existingConnection) {  # Connect only if not already connected
        try {
            Connect-OVMgmt -Hostname $fqdn -Credential $credential
        }
        catch {
            $errorMessage = "Error connecting to appliance ${fqdn}: $($_.Exception.Message)"
            Write-Warning $errorMessage
            $errorMessage | Add-Content $logFile
            continue  # Skip to the next appliance if connection fails
        }
    }

    try {
        $servers = Get-OVServer | Where-Object { $_.model -match 'Gen10' }

        foreach ($server in $servers) {
            $localStorageUri = $server.uri + '/localStorage'
            $localStorageDetails = Send-OVRequest -uri $localStorageUri
            
            if ($localStorageDetails) {
                foreach ($drive in $localStorageDetails.Data.PhysicalDrives) {
                    $info = [PSCustomObject]@{
                        ApplianceFQDN             = $fqdn
                        ServerName               = $server.Name -split ', ' | Select-Object -First 1
                        BayNumber                = $server.Name -split ', ' | Select-Object -Last 1
                        ServerStatus             = $server.Status
                        ServerPower              = $server.PowerState
                        ProcessorCoreCount       = $server.ProcessorCoreCount
                        ProcessorCount           = $server.ProcessorCount
                        ProcessorSpeedMhz        = $server.ProcessorSpeedMhz
                        ProcessorType            = $server.ProcessorType
                        ServerSerialNumber       = $server.SerialNumber
                        ServerModel              = $server.Model
                        AdapterType              = $localStorageDetails.Data.AdapterType
                        CurrentOperatingMode     = $localStorageDetails.Data.CurrentOperatingMode
                        FirmwareVersion          = $localStorageDetails.Data.FirmwareVersion.Current.VersionString
                        InternalPortCount        = $localStorageDetails.Data.InternalPortCount
                        Location                 = $localStorageDetails.Data.Location
                        LocationFormat           = $localStorageDetails.Data.LocationFormat
                        LogicalDriveNumbers      = ($localStorageDetails.Data.LogicalDrives | ForEach-Object { $_.LogicalDriveNumber }) -join ', '
                        RaidValues               = ($localStorageDetails.Data.LogicalDrives | ForEach-Object { $_.Raid }) -join ', '
                        Model                    = $localStorageDetails.Data.Model
                        DriveBlockSizeBytes       = $drive.BlockSizeBytes
                        LogicalCapacityGB        = [math]::Round(($drive.CapacityLogicalBlocks * $drive.BlockSizeBytes) / 1e9, 2)
                        DriveEncryptedDrive      = $drive.EncryptedDrive
                        DriveFirmwareVersion     = $drive.FirmwareVersion.Current.VersionString
                        DriveInterfaceType       = $drive.InterfaceType
                        DriveMediaType           = $drive.MediaType
                        DriveLocation            = $drive.Location
                        DriveModel               = $drive.Model
                        DriveSerialNumber        = $drive.SerialNumber
                        DriveStatus              = $drive.Status.Health
                        DriveState               = $drive.Status.State
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
        # No explicit disconnect needed. The HPE OneView PowerShell module handles disconnections automatically.
    }
}

$sortedData = $data | Sort-Object -Property ApplianceFQDN, BayNumber -Descending
$sortedData | Export-Csv (Join-Path $scriptPath "LocalStorageDetails.csv") -NoTypeInformation
$sortedData | Export-Excel (Join-Path $scriptPath "LocalStorageDetails.xlsx") -AutoSize
Write-Output "Audit completed and data exported to LocalStorageDetails.csv and LocalStorageDetails.xlsx"
