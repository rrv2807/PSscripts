Add-Type -AssemblyName System.Windows.Forms

# Define function to convert Windows version to friendly name
function Convert-WindowsVersionToFriendlyName {
    param (
        [string]$Version
    )

    # Map of Windows 10 versions to friendly names
    $versionMap = @{
        "10.0.10240" = "1507"
        "10.0.10586" = "1511"
        "10.0.14393" = "1607"
        "10.0.15063" = "1703"
        "10.0.16299" = "1709"
        "10.0.17134" = "1803"
        "10.0.17763" = "1809"
        "10.0.18362" = "1903"
        "10.0.19041" = "2004"
        "10.0.19042" = "20H2"
        "10.0.19043" = "21H1"
		"10.0.19045" = "22H1"
		"10.0.19058" = "22H2"
        "10.0.22000" = "22H2"  # Example, add more mappings as needed
    }

    # Attempt to match the version
    if ($versionMap.ContainsKey($Version)) {
        return $versionMap[$Version]
    } else {
        return "Unknown"
    }
}

# Function to get computer names from a text file
function Get-ComputerNamesFromFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select a text file containing computer names"
    $openFileDialog.Filter = "Text files (*.txt)|*.txt"
    $openFileDialog.Multiselect = $false

    $result = $openFileDialog.ShowDialog()
    if ($result -ne "OK") {
        Write-Host "File selection canceled."
        exit
    }

    $filePath = $openFileDialog.FileName
    return Get-Content $filePath
}

# Function to get computer names from the command shell
function Get-ComputerNamesFromCommandLine {
    $computerNames = (Read-Host "Enter computer names to inventory, separated by commas").Split(',')
    return $computerNames
}

# Prompt user to choose input method
$inputMethod = Read-Host "Choose input method: Enter 'file' to select a text file or 'cmd' to enter names in command shell"

if ($inputMethod -eq 'file') {
    $computerNames = Get-ComputerNamesFromFile
} elseif ($inputMethod -eq 'cmd') {
    $computerNames = Get-ComputerNamesFromCommandLine
} else {
    Write-Host "Invalid input method."
    exit
}

# Initialize an array to store results
$results = @()

# Loop through each computer name
foreach ($computerName in $computerNames) {
    # Trim any leading/trailing whitespace from the computer name
    $computerName = $computerName.Trim()

    # Check if the computer is reachable
    if (-not (Test-Connection -ComputerName $computerName -Count 1 -Quiet)) {
        Write-Host "Computer '$computerName' is not reachable."
        $result = [PSCustomObject]@{
            ComputerName = $computerName
            Processor = "N/A"
            RAM_GB = "N/A"
            Disks = "N/A"
            WindowsVersion = "N/A"
            LastUpdateInstalled = "N/A"
            Motherboard = "N/A"
        }
        $results += $result
        continue  # Skip to the next computer name
    }

    try {
        # Get processor information
        $processor = Get-WmiObject -ComputerName $computerName Win32_Processor | Select-Object -ExpandProperty Name

        # Get RAM information
        $ram = Get-WmiObject -ComputerName $computerName Win32_ComputerSystem | Select-Object -ExpandProperty TotalPhysicalMemory
        $ramGB = [math]::Round($ram / 1GB, 2)

        # Get disk space information
        $disks = Get-WmiObject -ComputerName $computerName Win32_LogicalDisk -Filter "DriveType=3" |
                 ForEach-Object {
                     "Drive: $($_.DeviceID), Size: $([math]::Round($_.Size / 1GB, 2)) GB, Free Space: $([math]::Round($_.FreeSpace / 1GB, 2)) GB"
                 }
        $disksInfo = $disks -join "; "

        # Get Windows version information
        $windowsVersion = (Get-CimInstance -ComputerName $computerName Win32_OperatingSystem).Version

        # Convert Windows version to friendly format
        $friendlyWindowsVersion = Convert-WindowsVersionToFriendlyName -Version $windowsVersion

        # Get last update installed information
        $lastUpdateInstalled = Get-WmiObject -ComputerName $computerName Win32_QuickFixEngineering | 
                               Sort-Object -Property InstalledOn -Descending | 
                               Select-Object -First 1 -ExpandProperty InstalledOn
        if (-not $lastUpdateInstalled) {
            $lastUpdateInstalled = "No updates found"
        }

        # Get motherboard information
        $motherboard = Get-WmiObject -ComputerName $computerName Win32_BaseBoard | 
                       Select-Object -ExpandProperty Product

        # Create result object
        $result = [PSCustomObject]@{
            ComputerName = $computerName
            Processor = $processor
			Motherboard = $motherboard
            RAM_GB = $ramGB
            Disks = $disksInfo
            WindowsVersion = $friendlyWindowsVersion
            LastUpdateInstalled = $lastUpdateInstalled
        }

        # Add result to results array
        $results += $result
    }
    catch {
        Write-Host "Error occurred while retrieving information from computer '$computerName': $_"
    }
}

# Display results in Grid View
$results | Out-GridView -Title "Inventory Results"
