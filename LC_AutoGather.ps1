# Define the version number
$versionNumber = "1.2.2"

# Define the progress title
$progressTitle = "Created by Patrick Moon. Version: $versionNumber"

# Define the list of possible clients
$clients = @(
    "2ndAndC",
    "DJC",
    "Faith",
    "MorrisonCPA",
    "OpenArms",
    "PLFNYA",
    "PLFNYE",
    "PLFNYF",
    "SBS",
    "SFleet",
    "LC"
    "Other"
)

#show progress
function outputProgress {
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Status,
         [Parameter(Mandatory=$true, Position=1)]
         [int] $Progress
    )

    Write-Progress -Activity $progressTitle -Status $Status -PercentComplete $Progress
}

# Check execution policy
if ((Get-ExecutionPolicy) -eq "Restricted") {
    Write-Warning "The script execution policy is set to Restricted. To run this script, you need to change the execution policy to RemoteSigned or Unrestricted. Please run Set-ExecutionPolicy and try again."
    exit
}

# Prompt user to select a client if no argument is provided
Write-Host "Choose a client to gather information from:" -ForegroundColor Cyan
Write-Host "Please select a client:"
for ($i=0; $i -lt $clients.Length; $i++) {
    Write-Host "$i. $($clients[$i])"
}
$clientIndex = Read-Host "Enter the number corresponding to the client you want to select"
$client = $clients[$clientIndex]

outputProgress "Getting Date..." 05
# Get the current date and format it as yyyy-MM-dd
$date = Get-Date -Format "yyyy-MM-dd"

outputProgress "Getting OS version..." 10
# Get the Windows version number
$osVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version

outputProgress "Getting HOSTNAME tag..." 15
# Get the hostname for the computer
$hostname = $env:computername

outputProgress "Getting service tag..." 20
# Get the Dell service tag
$serviceTag = Invoke-Command -ScriptBlock {
    Get-CimInstance -ClassName win32_bios | Select-Object -ExpandProperty SerialNumber
}

outputProgress "Getting IP addresses..." 30
# Get all network IP addresses
$ipAddresses = Invoke-Command -ScriptBlock {
    Get-NetIPAddress | Where-Object {$_.AddressFamily -eq "IPv4"} | Select-Object InterfaceAlias,IpAddress
}

outputProgress "Getting current user..." 40
# Get the currently logged in user
$currentUser = $env:USERNAME

outputProgress "Getting shared drives..." 50
# Get a list of shared drives and their locations
$shares = Invoke-Command -ScriptBlock {
    (Get-SmbShare | Where-Object {$_.ScopeName -eq "Default"}).Name
}

outputProgress "Getting remote shares..." 60
# Get a list of remote shares and their paths
$remoteShares = Invoke-Command -ScriptBlock {
    (Get-PSDrive -PSProvider FileSystem | Where-Object {$_.DisplayRoot -like "\\*\*"}).DisplayRoot
}

outputProgress "Getting printers..." 70
# Get a list of printers and their names
$printers = Invoke-Command -ScriptBlock {
    Get-Printer | Select-Object Name
}

outputProgress "Getting drives..." 80
# Get a list of all drives and their size and free space
$drives = Invoke-Command -ScriptBlock {
    Get-PSDrive -PSProvider 'FileSystem' | Select-Object Name, @{Name="Size(GB)";Expression={[math]::Round($_.Used/1GB)}}, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.Free/1GB)}}
}

outputProgress "Getting domain..." 90
# Get the domain the computer is connected to
switch ($client) {
    "SBS" {
        $domain = "SBS"
    }
    default { $domain = $env:USERDOMAIN}
}

outputProgress "Finished gathering data!" 100
# Create a custom object to store the service tag, IP addresses, machine name, logged in user, printers, shared drives, remote shares, and folders
$serviceInfoObj = @{
    "Date Create"    = $date
    "Script Version" = $versionNumber
    "OS Version"     = $osVersion
    "ClientName" = $client
    "Domain"     = $domain
    "ServiceTag" = $serviceTag
    "IPAddresses" = $ipAddresses
    "LoggedInUser" = $currentUser
    "Printers" = $printers
    "Shares" = $shares
    "RemoteShares" = $remoteShares
    "Drives" = $drives
    "Hostname" = $hostname
}

# Convert the object to a JSON string
$serviceInfoJson = ConvertTo-Json $serviceInfoObj -Depth 4

# Define the default file path and name to the user's desktop
#$jsonFilePath = "$env:USERPROFILE\Desktop\SystemInfo_$hostname.json"
$jsonFilePath = ".\SystemInfo~$client~$domain~$hostname.json"


# Write the service information to the CSV file
$serviceInfoJson | Out-File -FilePath $jsonFilePath -Encoding ascii

Write-Host "Service information was saved to $jsonFilePath"