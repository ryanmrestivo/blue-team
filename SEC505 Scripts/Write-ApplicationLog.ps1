##############################################################################
#.SNOPSIS
#   Write to the Application event log with a COM object.
#.NOTES
#   Date: 21.May.2007
#   Version: 1.0
#   Author: Jason Fossen, Enclave Consulting LLC (BlueTeamPowerShell.com)
#   Legal: 0BSD
##############################################################################



Param ([String] $message = "Your data here.", $type = "Information", $computer = "localhost")


Function Write-ApplicationLog ([String] $message = "text", $type = "Info", $computer = "localhost") 
{
    Switch -regex ($type) {
        'error'   {$type = 1}
        'warning' {$type = 2}
        'info'    {$type = 4}
        default   {$type = 4}
    }

    $WshShell = new-object -com "WScript.Shell"

    $ErrCode = $WshShell.LogEvent($type, $message, $computer)   
    $ErrCode
}

write-applicationlog -message $message -type $type -computer $computer


