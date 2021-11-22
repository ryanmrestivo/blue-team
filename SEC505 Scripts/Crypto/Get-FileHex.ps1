##############################################################################
#.SYNOPSIS
#   Show the hex of a file's contents.
#.PARAMETER Path
#   Path to the file.
#.PARAMETER Width
#   How many file bytes displayed per line of output.
#.NOTES
# Updated: 16.May.2021
# Version: 1.1
#  Author: Jason Fossen, Enclave Consulting LLC
#   Legal: 0BSD
##############################################################################

param ( $Path = $(throw "Enter path to file!"), [Int] $Width = 16 )


function Get-FileHex ( $path = $(throw "Enter path to file!"), [Int] $width = 16 ) {
    $linecounter = 0  
    $padwidth = $width * 3
    $placeholder = "."     # What to print when byte is not a letter or digit.

    if ($PSVersionTable.PSEdition -eq 'Core')
    { $splat = @{ Path = $path ; ReadCount = $width ; AsByteStream = $True } } 
    else 
    { $splat = @{ Path = $path ; ReadCount = $width ; Encoding = 'Byte' } }

    get-content @splat | 
    foreach-object { 
        $paddedhex = $asciitext = $null
        $bytes = $_        # Array of [Byte] objects that is $width items in length.
 
        foreach ($byte in $bytes) { 
            $simplehex = [String]::Format("{0:X}", $byte)     # Convert to hex.
            $paddedhex += $simplehex.PadLeft(2,"0") + " "     # Pad with zeros to force 2-digit length
        } 
 
        # Total bytes in file unlikely to be evenly divisible by $width, so fix last line.
        if ($paddedhex.length -lt $padwidth) 
           { $paddedhex = $paddedhex.PadRight($padwidth," ") }

        foreach ($byte in $bytes) { 
            if ( [Char]::IsLetterOrDigit($byte) -or 
                 [Char]::IsPunctuation($byte) -or 
                 [Char]::IsSymbol($byte) ) 
               { $asciitext += [Char] $byte }                 # Cast raw byte to a character.
            else 
               { $asciitext += $placeholder }
        }
        
        $offsettext = [String]::Format("{0:X}", $linecounter) # Linecounter to hex too.
        $offsettext = $offsettext.PadLeft(9,"0") + "h:"       # Pad hex linecounter with zeros.
        $linecounter += $width                                # Increment linecounter.

        "$offsettext $paddedhex $asciitext"           
    }
}


Get-FileHex -path $path -width $width

