# This is a quick egress test PowerShell script to see what ports (1-1024) are open from your network to the outside world
# The results get exported to a file called openports.txt that will get dropped in the same folder the script is run from
# Source: Black Hills Infosec (https://www.blackhillsinfosec.com/poking-holes-in-the-firewall-egress-testing-with-allports-exposed/)
#
#1..1024 | % {$test= new-object system.Net.Sockets.TcpClient; $wait = $test.beginConnect("allports.exposed",$_,$null,$null); ($wait.asyncwaithandle.waitone(250,$false)); if($test.Connected){echo "$_ open"}else{echo "$_ closed"}} | select-string " " > openports.txt

# I re-wrote this script below based on the script from above.

param (
    [string]$domain = "allports.exposed",
    [int]$startPort = 1,
    [int]$endPort = 1024,
    [int]$timeout = 250
)

# Output file
$outputFile = "openports.txt"

# Clear the output file if it exists
if (Test-Path $outputFile) {
    Clear-Content $outputFile
}

# Loop through the specified port range
for ($port = $startPort; $port -le $endPort; $port++) {
    Write-Progress -Activity "Scanning ports" -Status "Scanning port $port" -PercentComplete (($port - $startPort) / ($endPort - $startPort) * 100)
    
    # Attempt to make a connection
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $connect = $tcpClient.BeginConnect($domain, $port, $null, $null)
        $connected = $connect.AsyncWaitHandle.WaitOne($timeout, $false)
        
        # Write the result to the output file
        if ($connected) {
            "$port open" | Out-File -Append $outputFile
        } else {
            "$port closed" | Out-File -Append $outputFile
        }
    } catch {
        "$port error" | Out-File -Append $outputFile
    }
}

Write-Host "Scan complete. Results saved to $outputFile"
