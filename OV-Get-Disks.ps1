# Clear the console window
Clear-Host
# Create a string of 4 spaces
$Spaces = [string]::new(' ', 4)
# Define the script version
$ScriptVersion = "1.0"
# Get the directory from which the script is being executed
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
# Define the location of the script file
$ScriptFile = Join-Path -Path $scriptDirectory -ChildPath $MyInvocation.MyCommand.Name
# Define the variable to store date and time information for creation and last modification
$Created = (Get-ItemProperty -Path $ScriptFile -Name CreationTime).CreationTime.ToString("dd/MM/yyyy")
# Get the parent directory of the script's directory
$parentDirectory = Split-Path -Parent $scriptDirectory
# Define the logging function Directory
$loggingFunctionsDirectory = Join-Path -Path $parentDirectory -ChildPath "Logging_Function"
# Construct the path to the Logging_Functions.ps1 script
$loggingFunctionsPath = Join-Path -Path $loggingFunctionsDirectory -ChildPath "Logging_Functions.ps1"
# Script Header main script
$HeaderMainScript = @"
Author: CHARCHOUF Sabri
Description: This script creates Networks in HPE OneView using the HPE OneView PowerShell Library.
Created: $Created
Last Modified : $((Get-Item $PSCommandPath).LastWriteTime.ToString("dd/MM/yyyy"))
"@
# Display the header information in the console with a design
$consoleWidth = $Host.UI.RawUI.WindowSize.Width
$line = "─" * ($consoleWidth - 2)
Write-Host "+$line+" -ForegroundColor DarkGray
# Split the header into lines and display each part in different colors
$HeaderMainScript -split "`n" | ForEach-Object {
    $parts = $_ -split ": ", 2
    Write-Host "`t" -NoNewline
    Write-Host $parts[0] -NoNewline -ForegroundColor DarkGray
    Write-Host ": " -NoNewline
    Write-Host $parts[1] -ForegroundColor Cyan
}
Write-Host "+$line+" -ForegroundColor DarkGray
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Logging_Functions]---------------------------------------------------------------
# Check if the Logging_Functions.ps1 script exists
if (Test-Path -Path $loggingFunctionsPath) {
    # Dot-source the Logging_Functions.ps1 script
    . $loggingFunctionsPath
    # Write a message to the console indicating that the logging functions have been loaded
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Logging functions have been loaded." -ForegroundColor Green
}
else {
    # Write an error message to the console indicating that the logging functions script could not be found
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The logging functions script could not be found at $loggingFunctionsPath" -ForegroundColor Red
    # Stop the script execution
    exit
}
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Initialize task]-----------------------------------------------------------------
# Initialize task counter with script scope
$script:taskNumber = 1
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Import Required Modules]----------------------------------------------------------
# Define the function to import required modules if they are not already imported
function Import-ModulesIfNotExists {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ModuleNames
    )
    # Start logging
    Start-Log -ScriptVersion $ScriptVersion -ScriptPath $PSCommandPath
    # Task 1: Checking required modules
    Write-Host "`n$Spaces$($taskNumber). Checking required modules:`n" -ForegroundColor DarkGreen
    # Log the task
    Write-Log -Message "Checking required modules." -Level "Info" -NoConsoleOutput
    # Increment $script:taskNumber after the function call
    $script:taskNumber++
    # Total number of modules to check
    $totalModules = $ModuleNames.Count
    # Initialize the current module counter
    $currentModuleNumber = 0
    foreach ($ModuleName in $ModuleNames) {
        $currentModuleNumber++
        # Simple text output for checking required modules
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Checking module " -NoNewline -ForegroundColor DarkGray
        Write-Host "$currentModuleNumber" -NoNewline -ForegroundColor White
        Write-Host " of " -NoNewline -ForegroundColor DarkGray
        Write-Host "${totalModules}" -NoNewline -ForegroundColor Cyan
        Write-Host ": $ModuleName" -ForegroundColor White
        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
                Write-Host "`t• " -NoNewline -ForegroundColor White
                Write-Host "Module " -NoNewline -ForegroundColor White
                Write-Host "$ModuleName" -NoNewline -ForegroundColor Red
                Write-Host " is not installed." -ForegroundColor White
                Write-Log -Message "Module '$ModuleName' is not installed." -Level "Error" -NoConsoleOutput
                continue
            }
            # Check if the module is already imported
            if (Get-Module -Name $ModuleName) {
                Write-Host "`t• " -NoNewline -ForegroundColor White
                Write-Host "Module " -NoNewline -ForegroundColor DarkGray
                Write-Host "$ModuleName" -NoNewline -ForegroundColor Yellow
                Write-Host " is already imported." -ForegroundColor DarkGray
                Write-Log -Message "Module '$ModuleName' is already imported." -Level "Info" -NoConsoleOutput
                continue
            }
            # Try to import the module
            Import-Module $ModuleName -ErrorAction Stop
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Module " -NoNewline -ForegroundColor DarkGray
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Green
            Write-Host " imported successfully." -ForegroundColor DarkGray
            Write-Log -Message "Module '[$ModuleName]' imported successfully." -Level "OK" -NoConsoleOutput
        }
        catch {
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Failed to import module " -NoNewline
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Red
            Write-Host ": $_" -ForegroundColor Red
            Write-Log -Message "Failed to import module '[$ModuleName]': $_" -Level "Error" -NoConsoleOutput
        }
        # Add a delay to slow down the progress bar
        Start-Sleep -Seconds 1
    }
}
# Import the required modules
# Link to HPE OneView PowerShell Library: https://www.powershellgallery.com/packages/HPEOneView.800/8.0.3642.2784
Import-ModulesIfNotExists -ModuleNames 'HPEOneView.660', 'Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility', 'ImportExcel'
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Appliances list]-----------------------------------------------------------------
# Task 2: import Appliances list from the CSV file.
Write-Host "`n$Spaces$($taskNumber). Importing Appliances list from the CSV file:`n" -ForegroundColor DarkGreen
$csvPath = Join-Path $scriptDirectory "Appliances_liste.csv"
$appliances = Import-Csv $csvPath
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Confirm Import CSV file]---------------------------------------------------------
# Confirm that the CSV file was imported successfully
if ($appliances) {
    # Get the total number of appliances
    $totalAppliances = $appliances.Count
    # Log the total number of appliances
    Write-Log -Message "There are $totalAppliances appliances in the CSV file." -Level "Info" -NoConsoleOutput
    # Display if the CSV file was imported successfully
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The CSV file was imported successfully." -ForegroundColor Green
    # Display the total number of appliances
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Total number of appliances:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $totalappliances" -NoNewline -ForegroundColor Cyan
    Write-Host "" # This is to add a newline after the above output
    # Log the successful import of the CSV file
    Write-Log -Message "The CSV file was imported successfully." -Level "OK" -NoConsoleOutput
}
else {
    # Display an error message if the CSV file failed to import
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Failed to import the CSV file." -ForegroundColor Red
    # Log the failure to import the CSV file
    Write-Log -Message "Failed to import the CSV file." -Level "Error" -NoConsoleOutput
}
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Credential folder]---------------------------------------------------------------
# Define the path to the credential folder
$credentialFolder = Join-Path -Path $parentDirectory -ChildPath "Credential"
# increment $script:taskNumber after the function call
$script:taskNumber++
# Task 3: Check if credential folder exists
Write-Host "`n$Spaces$($taskNumber). Checking for credential folder:`n" -ForegroundColor DarkGreen
# Log the task
Write-Log -Message "Checking for credential folder." -Level "Info" -NoConsoleOutput
# Check if the credential folder exists, if not say it at console and create it, if already exist say it at console
if (Test-Path -Path $credentialFolder) {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Credential folder already exists at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $credentialFolder" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "Credential folder already exists at $credentialFolder" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Credential folder does not exist." -NoNewline -ForegroundColor Red
    Write-Host " Creating now..." -ForegroundColor DarkGray
    Write-Log -Message "Credential folder does not exist, creating now..." -Level "Info" -NoConsoleOutput
    # Create the credential folder if it does not exist already
    New-Item -ItemType Directory -Path $credentialFolder | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Credential folder created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $credentialFolder" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "Credential folder created at $credentialFolder" -Level "OK" -NoConsoleOutput
}
# Define the path to the credential file
$credentialFile = Join-Path -Path $credentialFolder -ChildPath "credential.txt"
# increment $script:taskNumber after the function call
$script:taskNumber++
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Check CSV & Excel Folders exists]------------------------------------------------
# Task 4: Check CSV & Excel Folders exists.
Write-Host "`n$Spaces$($taskNumber). Check CSV & Excel Folders exists:`n" -ForegroundColor DarkGreen
# Check if the credential file exists
if (-not (Test-Path -Path $credentialFile)) {
    # Prompt the user to enter their login and password
    $credential = Get-Credential -Message "Please enter your login and password."
    # Save the credential to the credential file
    $credential | Export-Clixml -Path $credentialFile
}
else {
    # Load the credential from the credential file
    $credential = Import-Clixml -Path $credentialFile
}
# Define the directories for the CSV and Excel files
$csvDir = Join-Path -Path $script:ReportsDir -ChildPath 'CSV'
$excelDir = Join-Path -Path $script:ReportsDir -ChildPath 'Excel'
# Check if the CSV directory exists
if (Test-Path -Path $csvDir) {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory already exists at:" -NoNewline -ForegroundColor DarkGray
    write-host " $csvDir" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "CSV directory already exists at $csvDir" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory does not exist." -NoNewline -ForegroundColor Red
    Write-Host " Creating now..." -ForegroundColor DarkGray
    Write-Log -Message "CSV directory does not exist, creating now..." -Level "Info" -NoConsoleOutput
    # Create the CSV directory if it does not exist already
    New-Item -ItemType Directory -Path $csvDir | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $csvDir" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "CSV directory created at $csvDir" -Level "OK" -NoConsoleOutput
}
# Check if the Excel directory exists
if (Test-Path -Path $excelDir) {
    # Write a message to the console
    write-host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory already exists at:" -NoNewline -ForegroundColor DarkGray
    write-host " $excelDir" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "Excel directory already exists at $excelDir" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory does not exist at" -NoNewline -ForegroundColor Red
    Write-Host " $excelDir" -ForegroundColor DarkGray
    # Write a message to the log file
    Write-Log -Message "Excel directory does not exist at $excelDir, creating now..." -Level "Info" -NoConsoleOutput
    # Create the Excel directory if it does not exist already
    New-Item -ItemType Directory -Path $excelDir | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $excelDir" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "Excel directory created at $excelDir" -Level "OK" -NoConsoleOutput
}
# Increment $script:taskNumber after the function call
$script:taskNumber++
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Data collection]-----------------------------------------------------------------
# Initialize an array to hold the collected data
$data = @()
# Log file for errors
$logFile = Join-Path -Path $scriptDirectory -ChildPath "error_log.txt"
# Loop through each appliance and retrieve the required information
foreach ($appliance in $appliances) {
    $fqdn = $appliance.Appliance_FQDN
    try {
        # Connect to the OneView appliance
        Connect-OVMgmt -Hostname $fqdn -Credential $credential
        # Get the server objects for Gen10 servers
        $servers = Get-OVServer | Where-Object { $_.model -match 'Gen10' }
        foreach ($server in $servers) {
            # Construct the URI for local storage details
            $localStorageUri = $server.uri + '/localStorage'
            # Retrieve the local storage details (using Send-OVRequest)
            $localStorageDetails = Send-OVRequest -Uri $localStorageUri -Method GET
            # Check if localStorageDetails is not null
            if ($null -ne $localStorageDetails) {
                foreach ($drive in $localStorageDetails.Data.PhysicalDrives) {
                    $info = [PSCustomObject]@{
                        ApplianceFQDN              = $fqdn
                        # Extract the server name and make it in CAPS
                        ServerName                 = $server.serverName.ToUpper()
                        # Extract the server bay number 
                        BayNumber                  = $server.Name.Split(', ')[1]
                        # Extract the server Status {Critical, Warning, OK}
                        ServerStatus               = $server.Status
                        # Extract the server Power State {On, Off, Standby}
                        ServerPower                = $server.PowerState
                        # Extract the server Serial Number
                        ServerSerialNumber         = $server.SerialNumber
                        # Extract the server Model
                        ServerModel                = $server.Model
                        # Extract the Adapter Type {HBA, RAID}
                        AdapterType                = $localStorageDetails.Data.AdapterType
                        # Extract the Model of the Adapter
                        Model                      = $localStorageDetails.Data.Model
                        # Extract the Current Operating Mode {RAID, HBA}
                        CurrentOperatingMode       = $localStorageDetails.Data.CurrentOperatingMode
                        # Extract the Firmware Version of the Adapter                 
                        FirmwareVersion            = $localStorageDetails.Data.FirmwareVersion.Current.VersionString
                        # Extract the Internal Port Count of the Adapter
                        InternalPortCount          = $localStorageDetails.Data.InternalPortCount
                        # Extract the location of the adapter {Slot}
                        Location                   = $localStorageDetails.Data.Location
                        # Extract the Location Format {PCI}
                        LocationFormat             = $localStorageDetails.Data.LocationFormat
                        # Extract the Logical Drive Count
                        LogicalDriveNumbers        = ($localStorageDetails.Data.LogicalDrives | ForEach-Object { $_.LogicalDriveNumber }) -join ', '
                        # Extract the RAID Values {RAID 1, RAID 5, RAID 6}
                        RaidValues                 = ($localStorageDetails.Data.LogicalDrives | ForEach-Object { $_.Raid }) -join ', '
                        # Extract the Size of the Block in Bytes {512, 4096}
                        DriveBlockSizeBytes        = $drive.BlockSizeBytes
                        # Calculate the logical capacity in GB
                        LogicalCapacityGB          = [math]::Round(($drive.CapacityLogicalBlocks * $drive.BlockSizeBytes) / 1e9, 2)
                        # Check if the drive is encrypted
                        DriveEncryptedDrive        = $drive.EncryptedDrive
                        # Extract the Firmware Version of the Drive as it should be started with "HP"
                        DriveFirmwareVersion       = $drive.FirmwareVersion.Current.VersionString
                        # Extract the Interface Type {SAS, SATA}
                        DriveInterfaceType         = $drive.InterfaceType
                        # Extract the MediaType {SSD, HDD}
                        DriveMediaType             = $drive.MediaType
                        # Extract the Location of the Drive {Drive Bay}
                        DriveLocation              = $drive.Location
                        # Extract the Model of the Drive
                        DriveModel                 = $drive.Model
                        # Extract the Serial Number of the Drive
                        DriveSerialNumber          = $drive.SerialNumber
                        # Extract the Status of the Drive {Critical, Warning, OK}
                        DriveStatus                = $drive.Status.Health
                        # Extract the State of the Drive {Enabled, Disabled}
                        DriveState                 = $drive.Status.State
                        # Extract the Life Remaining of the Drive in Percentage
                        "Drive Life Remaining (%)" = "{0}%" -f (100 - $drive.SSDEnduranceUtilizationPercentage)
                    }
                    # Add the collected information to the data array
                    $data += $info
                }
            }
        }
        # Disconnect from the OneView appliance
        Disconnect-OVMgmt -Hostname $fqdn
    }
    catch {
        # Log the error message and continue to the next appliance in the list if an error occurs during data collection
        $errorMessage = "Error processing appliance ${fqdn}: $($_.Exception.Message)"
        Write-Warning $errorMessage
        $errorMessage | Add-Content -Path $logFile
    }
}
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Close Excel]---------------------------------------------------------------------
# Task 5: Closing Excel
Write-Host "`n$Spaces$($taskNumber). Closing Excel:`n" -ForegroundColor DarkGreen
# Log the task
Write-Log -Message "Closing Excel." -Level "Info" -NoConsoleOutput
# Get all Excel processes
$excelProcesses = Get-Process -Name Excel -ErrorAction SilentlyContinue
# If there are any Excel processes
if ($excelProcesses) {
    # Stop all Excel processes
    $excelProcesses | ForEach-Object {
        Stop-Process -Id $_.Id -Force
    }
    # Write a message to the console
    Write-Host "`t• All running Excel processes have been closed." -NoNewline -ForegroundColor DarkGray
    Write-Host " ✔" -ForegroundColor Green
}
else {
    # Write a message to the console
    Write-Host "No Excel processes are currently running."-NoNewline -ForegroundColor DarkGray
    Write-Host " ℹ" -ForegroundColor Cyan
}
# -------------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------- [Export Data to Excel]------------------------------------------------------------
# Task 6: Export Data to Excel
Write-Host "`n$Spaces$($taskNumber). Exporting Data to Excel:`n" -ForegroundColor DarkGreen
# Log the task
Write-Log -Message "Exporting Data to Excel." -Level "Info" -NoConsoleOutput
# Increment $script:taskNumber after the function call
$script:taskNumber++
# Sorting and exporting data to CSV and Excel
$sortedData = $data | Sort-Object -Property ApplianceFQDN, Servername
# Export data to CSV file (append mode)
$csvPath = Join-Path $csvDir -ChildPath "LocalStorageDetails.csv"
$csvExported = $false
while (-not $csvExported) {
    try {
        $sortedData | Export-Csv -Path $csvPath -NoTypeInformation -Append
        $csvExported = $true
    }
    catch {
        Write-Warning "Failed to export data to the CSV file. Retrying..."
        Start-Sleep -Seconds 1
    }
}
# Import data to Excel file (append mode) and apply VBA macro
$excelPath = Join-Path $excelDir -ChildPath "LocalStorageDetails.xlsx"
$worksheetName = "LocalStorageDetails"
try {
    if (Test-Path -Path $csvPath) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($csvPath)
        $worksheet = $workbook.Worksheets.Item(1)
        # Rename worksheet
        $worksheet.Name = $worksheetName
        # Add VBA macro to highlight selected row and column
        $vbaCode = @"
        Private Sub Worksheet_SelectionChange(ByVal Target As Range)
            Dim selectedRow As Range
            Dim selectedColumn As Range
            ' Clear previous highlighting
            Cells.Interior.ColorIndex = xlNone
            ' Highlight selected row
            Set selectedRow = Rows(Target.Row)
            selectedRow.Interior.Color = RGB(255, 255, 0) ' Yellow color
            ' Highlight selected column
            Set selectedColumn = Columns(Target.Column)
            selectedColumn.Interior.Color = RGB(255, 255, 0) ' Yellow color
        End Sub
"@
        $vbaModule = $workbook.VBProject.VBComponents.Add(1)
        $vbaModule.CodeModule.AddFromString($vbaCode)
        # Save and close the Excel file
        $workbook.SaveAs($excelPath)
        $workbook.Close()
        # Quit Excel application
        $excel.Quit()
        Write-Output "Excel file saved successfully."
    }
    else {
        Write-Warning "CSV file not found at $csvPath. Skipping Excel export."
    }
}
catch {
    Write-Warning "Failed to import data to Excel and apply VBA macro. Error: $($_.Exception.Message)"
}
