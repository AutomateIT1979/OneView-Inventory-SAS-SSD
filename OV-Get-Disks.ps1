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
    $data = @()
    $mediaFilter = ($mediaType -eq 'All') -or ($drive.driveMedia -eq $mediaType)
    Write-Host "Processing drive: $($drive.Name) - Media Filter: $mediaFilter"
    if ($mediaFilter) {
        $sn = $drive.serialNumber
        if ($sn) {
            $interface = $drive.deviceInterface
            $media = $drive.driveMedia
            $model = $drive.model
            $fw = $drive.firmwareVersion
            $ssdPercentUsage = [int]$drive.SSDEnduranceUtilizationPercentage
            $ph = $drive.PowerOnHours
            $powerOnHours = $ssdUsage = ""
            if ($media -eq 'SSD') {
                $timeSpan = new-timespan -hours $ph
                $years = [math]::floor($timeSpan.Days / 365)
                $months = [math]::floor(($timeSpan.Days % 365) / 30)
                $days = ($timeSpan.Days % 365) % 30
                $hours = $timeSpan.Hours
                $powerOnHours = "$years years-$months months-$days days-$hours hours"
                $ssdUsage = "$ssdPercentUsage%"
            }
            $data += "$drive.Name,$interface,$media,$model,$sn,$fw,$ssdUsage,$powerOnHours"
        }
    }
    return $data
}

# Function to fetch disk inventory from servers
Function Get-ServerInventory {
    param (
        $server,
        $mediaType
    )
    $inventory = @()
    $lStorageUri = $server.subResources.LocalStorage.uri
    $lStorage = Send-OVRequest -uri $lStorageUri
    Write-Host "Local Storage Data: $($lStorage | Out-String)"
    foreach ($drive in $lStorage.data.PhysicalDrives) {
        $driveData = Get-DriveDetails -drive $drive -mediaType $mediaType
        Write-Host "Drive Data: $($driveData | Out-String)"
        $inventory += $driveData
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

    $diskInventory = @("Server,serverModel,serverSN,Interface,MediaType,SerialNumber,firmware,ssdEnduranceUtilizationPercentage,powerOnHours")

    Write-Host -ForegroundColor Cyan "---- Connecting to OneView --> $hostName"
    try {
        # Connect to OneView with the provided credentials
        Connect-OVMgmt -Hostname $hostName -Credential $credentials -loginAcknowledge:$true 

        $outFile = "$hostName-$date-disk_Inventory.csv"
        $errorFile = "$hostName-$date-errors.txt"

        # Set Message (slightly improved clarity)
        $diskMessage = "disks"
        if ($mediaType -ne 'All') { $diskMessage += " and $mediaType media" }

        # Collect Server inventory
        $serverList = Get-OVServer | Where-Object { $_.mpModel -notlike '*ilo3*' }

        foreach ($server in $serverList) {
            $data = @()
            $sName = $server.Name
            $sModel = $server.Model
            $sSN = $server.SerialNumber
            $serverPrefix = "$sName,$sModel,$sSN"

            Write-Host "---- Collecting $diskMessage information on server ---> $sName"
            $data = Get-ServerInventory -server $server -mediaType $mediaType

            if ($data) {
                $data = $data | ForEach-Object { "$serverPrefix,$_" }
                $diskInventory += $data
            } else {
                Write-Host -ForegroundColor Yellow "------ No matching $diskMessage found on $sName...."
            }
        }

        $diskInventory | Out-File $outFile -Encoding UTF8

        Write-Host -ForegroundColor Cyan "Disk Inventory on server complete --> file: $outFile`n" 
    } catch {
        Write-Host -ForegroundColor Red "Error: $_"
        $_ | Out-File -FilePath $errorFile -Append
    } finally {
        Disconnect-OVMgmt
    }
}
