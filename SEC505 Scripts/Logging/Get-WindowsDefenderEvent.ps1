<#
.SYNOPSIS
 Get interesting Event Log messages produced by Windows Defender AV.

.NOTES
 For a list of Windows Defender event ID numbers and their meaning, see:
 https://docs.microsoft.com/en-us/windows/security/threat-protection/windows-defender-antivirus/troubleshoot-windows-defender-antivirus
#> 


(Param $ComputerName = "$env:COMPUTERNAME", $MaxEvents = 1000)

$XPath = '*[System[(EventID=1006 or EventID=1015 or EventID=1116 or EventID=1121 or EventID=1007 or EventID=1117 or EventID=5001 or EventID=5010 or EventID=5008)]]'

Get-WinEvent -ComputerName $ComputerName -FilterXPath $XPath -MaxEvents $MaxEvents -LogName 'Microsoft-Windows-Windows Defender/Operational' 


<#
# This also works, if you don't like XPath syntax, but it's slower:
Get-WinEvent -MaxEvents 100 -LogName 'Microsoft-Windows-Windows Defender/Operational' |
Where-Object { @(1006,1015,1116,1121,1007,1117,5001,5010,5008) -contains $_.Id } 
#>

