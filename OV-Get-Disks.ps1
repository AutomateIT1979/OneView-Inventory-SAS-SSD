[CmdletBinding()]
Param (
    [string]$serverUri = "/rest/server-hardware/39313738-3034-5A43-4A31-343230364638/localStorage",
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
        $sn = $drive.serialNumber
        if ($sn) {
            $interface = $drive.deviceInterface
            $media = $drive.driveMedia
            $model = $drive.model
            $fw = $drive.firmwareVersion
            $ssdPercentUsage = [int]$drive.SSDEnduranceUtilizationPercentage
            $ph = $drive.PowerOnHours
            $powerOnHours = ""
            $ssdUsage = ""
            if ($media -eq 'SSD') {
                $timeSpan = New-TimeSpan -Hours $ph
                $years = [math]::floor($timeSpan.Days / 365)
                $months = [math]::floor(($timeSpan.Days % 365) / 30)
                $days = ($timeSpan.Days % 365) % 30
                $hours = $timeSpan.Hours
                $powerOnHours = "$years years-$months months-$days days-$hours hours"
                $ssdUsage = "$ssdPercentUsage%"
            }
            $data = [PSCustomObject]@{
                Name = $drive.Name
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
        $serverUri,
        $mediaType
    )
    $inventory = @()
    $lStorage = Send-OVRequest -Uri $serverUri
    Write-Host "Local Storage Data: $($lStorage | Out-String)"
    foreach ($drive in $lStorage.data.PhysicalDrives) {
        $driveData = Get-DriveDetails -drive $drive -mediaType $mediaType
        if ($driveData) {
            Write-Host "Drive Data: $($driveData | Out-String)"
            $inventory += $driveData
        }
    }
    return $inventory
}

# Example usage: Fetch and display inventory for the specified server URI
$inventory = Get-ServerInventory -serverUri $serverUri -mediaType $mediaType
$inventory | Format-Table -AutoSize
