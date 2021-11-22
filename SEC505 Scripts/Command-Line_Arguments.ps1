##############################################################################
#.SYNOPSIS
#      Demo how command line arguments are processed.
#.NOTES
#      Date: 27.Mar.2007
#      Author: Jason Fossen, Enclave Consulting LLC (BlueTeamPowerShell.com)
#      Legal: 0BSD
##############################################################################


# All arguments go into the $args object array:

$Args.GetType().FullName    # System.Object[]

$Args.Length                # Number of arguments

$Args[0]                    # First argument
$Args[1]                    # Second argument
$Args[0..1]                 # First two arguments
$Args[-1]                   # Last argument
$Args[-2..-1]               # Last two arguments

[System.String] $Args       # All arguments in a single string
       [String] $Args       # All arguments in a single string 


# Because $Args is an array, you can enumerate through it:

ForEach ($Arg in $Args) { $Arg }



# If you know you'll pass in three arguments, they can be assigned on one line:

$1st, $2nd, $3rd = $Args[0..2]

$1st + " : " + $2nd + " : " + $3rd



# If you reference an argument that wasn't passed in, the variable will be $Null:

If ($Args[18] -eq $Null) { "It's null." } Else { "Not null." }



