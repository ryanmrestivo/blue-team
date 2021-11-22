##############################################################################
#.SYNOPSIS
#   Demo WMI.
#.NOTES
#   Legal: Script provided "AS IS" without warranties or guarantees of any
#          kind.  USE AT YOUR OWN RISK.  Public domain, no rights reserved.
##############################################################################



param ($computer = ".")

function Get-ProcessWithWMI ($computer = ".") 
{
    get-wmiobject -query "SELECT * FROM Win32_Process" -computername $computer |
    select-object Name,ProcessID,CommandLine,
                  @{Name="Domain"; Expression={$_.GetOwner().domain}}, 
                  @{Name="User"; Expression={$_.GetOwner().user}}, 
                  @{Name="CreationDate"; Expression={  if ($_.CreationDate -ne $null) {$_.ConvertToDateTime($_.CreationDate)} 
                                                       else {$null}  }  }
}

get-processwithwmi $computer

