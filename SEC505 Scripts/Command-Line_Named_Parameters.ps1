##############################################################################
#.SYNOPSIS
#   Demo how command line parameters are processed.
#.NOTES
#    Date: 14.May.2007
#  Author: Jason Fossen (BlueTeamPowerShell.com)
#   Legal: 0BSD
##############################################################################


# If you wish to pass in named parameters instead of using $Args, the
# param keyword must be the first executable line in the script, i.e.,
# the first statement after any blank or comment lines.

# The types of the parameters can be constrained with casts, e.g.,
# [String], [Int], [DateTime], etc.

# Default values can be assigned to some, all or none of the named
# parameters with the "=" sign.  Defaults can be overwritten by
# passing in values for some or all of the parameters.

param ([String] $word = "Cat", [Int] $number = 3)
$word * $number


