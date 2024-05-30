[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true, HelpMessage = "Path to CSV file containing OneView instances.")]
    [string]$csvFilePath,
    [Parameter(HelpMessage = "Filter by device interface (All, SAS, SATA, etc.). Defaults to All.")]
    [string]$interfaceType = 'All',
    [Parameter(HelpMessage = "Filter by media type (All, SSD, HDD). Defaults to All.")]
    [string]$mediaType = 'All'
)

# Function to fetch drive details
Function Get-DriveDetails {
    param (
        $drive,
        $interfaceType,
        $mediaType
    )
    $data = @()
    $interfaceFilter = ($interfaceType -eq 'All') -or ($drive.deviceInterface -eq $interfaceType)
    $mediaFilter = ($mediaType -eq 'All') -or ($drive.driveMedia -eq $mediaType)
    if ($interfaceFilter -and $mediaFilter) {
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
        $interfaceType,
        $mediaType
    )
    $inventory = @()
    $lStorageUri = $server.subResources.LocalStorage.uri
    $lStorage = Send-OVRequest -uri $lStorageUri
    foreach ($drive in $lStorage.data.PhysicalDrives) {
        $inventory += Get-DriveDetails -drive $drive -interfaceType $interfaceType -mediaType $mediaType
    }
    return $inventory
}

$date = (Get-Date).ToString('MM_dd_yyyy')

# Read OneView instances from CSV
$oneviewInstances = Import-Csv -Path $csvFilePath

foreach ($instance in $oneviewInstances) {
    $hostName = $instance.hostName
    $userName = $instance.userName
    $password = $instance.password
    $authLoginDomain = $instance.authLoginDomain

    $diskInventory = @("Server,serverModel,serverSN,Interface,MediaType,SerialNumber,firmware,ssdEnduranceUtilizationPercentage,powerOnHours")

    # Securely store credentials
    $securePassword = $password | ConvertTo-SecureString -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword

    Write-Host -ForegroundColor Cyan "---- Connecting to OneView --> $hostName"
    try {
        Connect-HPOVMgmt -Hostname $hostName -loginAcknowledge:$true -AuthLoginDomain $authLoginDomain -Credential $cred

        $outFile = "$hostName-$date-disk_Inventory.csv"
        $errorFile = "$hostName-$date-errors.txt"

        # Set Message (slightly improved clarity)
        $diskMessage = "disks"
        if ($interfaceType -ne 'All') { $diskMessage += " with $interfaceType interface" }
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
            $data = Get-ServerInventory -server $server -interfaceType $interfaceType -mediaType $mediaType

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
        Disconnect-HPOVMgmt
    }
}
