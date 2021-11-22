##############################################################################
#.SYNOPSIS
#   Get driver information through WMI.
#.NOTES
#    Date: 30.May.2007
# Version: 1.0
#  Author: Jason Fossen, Enclave Consulting LLC
#   Legal: 0BSD
##############################################################################

param ($Computer)

function Get-DriverWithWMI ($Computer) 
{
    Get-CimInstance -Query "SELECT * FROM Win32_SystemDriver" -ComputerName $Computer |
    Select-Object Name,DisplayName,PathName,ServiceType,State,StartMode
}

Get-DriverWithWMI -Computer $Computer

