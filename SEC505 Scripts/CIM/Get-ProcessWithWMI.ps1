##############################################################################
#.SYNOPSIS
#   Demo how to retrieve process information through WMI.
#.NOTES
#   The idle process has a $null CreationDate.
#   Date: 30.May.2007
#   Author: Jason Fossen, Enclave Consulting LLC
#   Legal: 0BSD
##############################################################################


Param ($Computer)


function Get-ProcessWithWMI ($Computer) 
{
    Get-CimInstance -Query "SELECT * FROM Win32_Process" -ComputerName $Computer |
    Select-Object Name,ProcessID,CreationDate,CommandLine
}


Get-ProcessWithWMI -Computer $Computer



