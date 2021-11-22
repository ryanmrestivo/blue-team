##############################################################################
#.SYNOPSIS
#   Demo WMI.
#.NOTES
#   Legal: Script provided "AS IS" without warranties or guarantees of any
#          kind.  USE AT YOUR OWN RISK.  Public domain, no rights reserved.
##############################################################################

param ($computer = ".", $commandline = "calc.exe")


function Create-ProcessWithWMI ($computer = ".", $commandline = $(throw "Enter the command to execute.") ) 
{
    $ProcessClass = get-wmiobject -query "SELECT * FROM Meta_Class WHERE __Class = 'Win32_Process'" -namespace "root\cimv2" -computername $computer
    # $ProcessClass = [WMICLASS] "root\cimv2:Win32_Process"   # The equally valid, shorter, but more obscure method.
    $results = $ProcessClass.Create( $commandline )

    if ($results.ReturnValue -eq 0) { $results.ProcessID }  # Or just return $true if you don't want the PID.
    else { $false ; throw "Failed to create process!" }
}


create-processwithwmi -computer $computer -commandline $commandline






