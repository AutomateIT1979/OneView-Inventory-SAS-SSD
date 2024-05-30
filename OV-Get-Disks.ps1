[CmdletBinding()]
Param (
    [Parameter(HelpMessage = "Filter by media type (All, SSD, HDD). Defaults to All.")]
    [string]$mediaType = 'All'
)

# Function to fetch drive details
Function Get-DriveDetails {
    param (
        $drive,
        $mediaType
    )
    $data = $null
    $mediaFilter = ($mediaType -eq 'All') -or ($drive.driveMedia -eq $mediaType)
    if ($mediaFilter) {
        $sn = $drive.SerialNumber
        if ($sn) {
            $interface = $drive.Interface
            $media = $drive.Model
            $model = $drive.Model
            $fw = $drive.FirmwareVersion
            $ssdPercentUsage = [int]$drive.SSDEnduranceUtilizationPercentage
            $ph = $drive.PowerOnHours
            $powerOnHours = ""
            $ssdUsage = ""
            if ($media -match 'SSD') {
                $timeSpan = New-TimeSpan -Hours $ph
                $years = [math]::floor($timeSpan.Days / 365)
                $months = [math]::floor(($timeSpan.Days % 365) / 30)
                $days = ($timeSpan.Days % 365) % 30
                $hours = $timeSpan.Hours
                $powerOnHours = "$years years-$months months-$days days-$hours hours"
                $ssdUsage = "$ssdPercentUsage%"
            }
            $data = [PSCustomObject]@{
                Name = $drive.Location
                Interface = $interface
                MediaType = $media
                Model = $model
                SerialNumber = $sn
                Firmware = $fw
                SSDUsage = $ssdUsage
                PowerOnHours = $powerOnHours
            }
        }
    }
    return $data
}

# Function to fetch disk inventory from a single server
Function Get-ServerInventory {
    param (
        $server,
        $mediaType
    )
    $inventory = @()
    $localStorageUri = $server.uri + "/localStorage"
    $lStorage = Send-OVRequest -Uri $localStorageUri
    Write-Host "Local Storage Data for $($server.Name): $($lStorage | Out-String)"
    foreach ($drive in $lStorage.PhysicalDrives) {
        $driveData = Get-DriveDetails -drive $drive -mediaType $mediaType
        if ($driveData) {
            Write-Host "Drive Data for $($drive.Location): $($driveData | Out-String)"
            $inventory += $driveData
        }
    }
    return $inventory
}

$date = (Get-Date).ToString('MM_dd_yyyy')

# Get the path to the current script directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
# Construct the full path to the appliance list CSV file
$applianceListPath = Join-Path $scriptPath "Appliances_liste.csv"

# Read appliance FQDNs from CSV
$appliances = Import-Csv -Path $applianceListPath

foreach ($appliance in $appliances) {
    $hostName = $appliance.Appliance_FQDN # Get FQDN from the column

    # Prompt for credentials
    $credentials = Get-Credential -Message "Please enter your OneView credentials for $hostName"

    $diskInventory = @()

    try {
        # Connect to OneView with the provided credentials
        Connect-OVMgmt -Hostname $hostName -Credential $credentials -loginAcknowledge:$true 

        $outFile = "$hostName-$date-disk_Inventory.csv"
        $errorFile = "$hostName-$date-errors.txt"

        # Collect Server inventory
        $serverList = Get-OVServer | Where-Object { $_.model -match 'Gen10' }

        foreach ($server in $serverList) {
            $sName = $server.Name
            $sModel = $server.Model
            $sSN = $server.SerialNumber

            $serverInventory = Get-ServerInventory -server $server -mediaType $mediaType

            foreach ($drive in $serverInventory) {
                $driveDetails = [PSCustomObject]@{
                    ServerName = $sName
                    ServerModel = $sModel
                    ServerSerialNumber = $sSN
                    DriveName = $drive.Name
                    Interface = $drive.Interface
                    MediaType = $drive.MediaType
                    Model = $drive.Model
                    SerialNumber = $drive.SerialNumber
                    Firmware = $drive.Firmware
                    SSDUsage = $drive.SSDUsage
                    PowerOnHours = $drive.PowerOnHours
                }
                Write-Host "Collected Drive Details: $($driveDetails | Out-String)"
                $diskInventory += $driveDetails
            }
        }

        $diskInventory | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
        Write-Host "Disk inventory successfully written to $outFile"

    } catch {
        Write-Host -ForegroundColor Red "Error: $_"
        $_ | Out-File -FilePath $errorFile -Append
    } finally {
        Disconnect-OVMgmt
    }
}
