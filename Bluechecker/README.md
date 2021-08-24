# Bluechecker

![Bluechecker](https://ctrla1tdel.files.wordpress.com/2020/04/blu.gif)


BlueChecker will help you audit PowerShell and check for any suspicious activity. At the end it will then generate a report.
Default location: C:\Temp\report.html

Simply download the script or run remotely using:

powershell –nop –c “iex(New-Object Net.WebClient).DownloadString(‘https://raw.githubusercontent.com/securethelogs/Bluechecker/master/BlueChecker.ps1’)”

Once ran, BlueChecker will check for:

- Powershell status
- Evidence of downgrading
- Registry and GP set for PowerShell auditing
- Malicious scripts using keywords
- Firewall spesific to Powershell
- Event logs for Module logging and script block logging.



 For More Information, visit: https://securethelogs.com/hacking-with-powershell-blue-team/


