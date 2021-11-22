##############################################################################
#.SYNOPSIS
#   Select files by size and sorts them from largest to smallest.
##############################################################################

param ([String] $Location = $pwd, [Int] $SizeMB = 5, [String] $Filter = "*")

get-childitem $location -recurse -filter $filter | 
    where-object { $_.length -ge ($sizeMB * 1MB) } | 
    sort-object -desc length



# "MB" is built-in PowerShell shorthand for "1024 * 1024".
