# Shields Up!
# This simple script disables ALL radios - be they wireless, wired, Bluetooth, etc.
# This is handy for instances where you think a machine might be compromised and you want to
# ensure it doesn't talk to anything else on any other networks.

# Check for admin privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script requires administrative privileges. Please run as administrator."
    exit
}

# Specify log file
$logFile = "NetworkAdapterActions.log"

# Function to write to log
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -Append $logFile
}

# Prompt user for action (enable or disable)
$action = Read-Host "Do you want to enable or disable network interfaces? (Enter 'enable' or 'disable')"

# Validate user input
if ($action -ne 'enable' -and $action -ne 'disable') {
    Write-Host "Invalid input. Please enter 'enable' or 'disable'."
    exit
}

# Prompt for confirmation
$confirmation = Read-Host "This script will $action all network interfaces. Are you sure you want to continue? (Y/N)"
if ($confirmation -ne 'Y') {
    Write-Host "Exiting script."
    exit
}

try {
    # Log the initial state of network adapters
    $adapters = Get-NetAdapter
    foreach ($adapter in $adapters) {
        Write-Log "Initial state of $($adapter.Name): $($adapter.Status)"
    }

    # Perform the chosen action
    if ($action -eq 'disable') {
        # Disable all network adapters
        Disable-NetAdapter -Name * -Confirm:$false
        Write-Host "The interfaces - such as wired/wireless/Bluetooth - are now all disabled."
        Write-Log "All network interfaces have been disabled."
    } else {
        # Enable all network adapters
        Enable-NetAdapter -Name * -Confirm:$false
        Write-Host "All interfaces are now enabled."
        Write-Log "All network interfaces have been enabled."
    }
} catch {
    Write-Host "An error occurred: $($_.Exception.Message)"
    Write-Log "Error: $($_.Exception.Message)"
}
