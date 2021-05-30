# scri%00pt - Privilege Escalation

scri%00pt is PowerShell script designed to help penetration testers quickly identify potential privilege escalation vectors on Windows systems. It is written using PowerShell 2.0 so 'should' run on every Windows version since Windows 7.  Will absolutely tigger ATP in current state.

## Usage:


**Run from within CMD shell and write out to file.**
```
CMD C:\temp> powershell.exe -ExecutionPolicy Bypass -File .\scri%00pt-enum.ps1 -OutputFilename scri%00pt-Enum.txt
```
**Run from within CMD shell and write out to screen.**
```
CMD C:\temp> powershell.exe -ExecutionPolicy Bypass -File .\scri%00pt-enum.ps1
```
**Run from within PS Shell and write out to file.**
```
PS C:\temp> .\scri%00pt-enum.ps1 -OutputFileName scri%00pt-Enum.txt
```

## Current Features
  - Network Information (interfaces, arp, netstat)
  - Firewall Status and Rules
  - Running Processes
  - Files and Folders with Full Control or Modify Access
  - Mapped Drives
  - Potentially Interesting Files
  - Unquoted Service Paths
  - Recent Documents
  - System Install Files 
  - AlwaysInstallElevated Registry Key Check
  - Stored Credentials
  - Installed Applications
  - Potentially Vulnerable Services
  - MuiCache Files
  - Scheduled Tasks