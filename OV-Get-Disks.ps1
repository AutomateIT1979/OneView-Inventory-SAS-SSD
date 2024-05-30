
# Get the server object for Gen10 servers and select the first one
$server = Get-OVServer | Where-Object { $_.model -match 'Gen10' } | Select-Object -First 1

# Construct the URI for local storage details
$localStorageUri = $server.uri + '/localStorage'

# Retrieve the local storage details
$localStorageDetails = Invoke-OVCommand -uri $localStorageUri

# Display the local storage details
$localStorageDetails.Data
