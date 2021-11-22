param ($Path)

function ConvertTo-ASCII ($Path) 
{
    #.SYNOPSIS
    #  Convert text file to ASCII text.
    $Path = (Resolve-Path -Path $Path -ErrorAction Stop).Path
    # Do not remove parentheses or else the file will become empty!
    (Get-Content -Path $Path -ReadCount 0) | Out-File -Encoding ascii -FilePath $Path 
}

ConvertTo-ASCII -Path $Path 
