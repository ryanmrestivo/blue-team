##############################################################################
#.SYNOPSIS
#   Demo WMI.
#.NOTES
#   Legal: Script provided "AS IS" without warranties or guarantees of any
#          kind.  USE AT YOUR OWN RISK.  Public domain, no rights reserved.
##############################################################################

param ($computer = ".")

function Get-DriverWithWMI ($computer = ".") 
{
    get-wmiobject -query "SELECT * FROM Win32_SystemDriver" -computername $computer |
    select-object Name,DisplayName,PathName,ServiceType,State,StartMode
}

get-driverwithwmi $computer

