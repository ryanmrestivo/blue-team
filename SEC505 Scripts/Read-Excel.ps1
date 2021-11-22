##############################################################################
#.SYNOPSIS
#   Reads one or more cells out of an Excel spreadsheet.
#.NOTES
#    Date: 2.Jul.2007
#  Author: Jason Fossen (BlueTeamPowerShell.com)
#   Legal: 0BSD
##############################################################################

param ($path = $(throw "Enter full path to Excel spreadsheet file."), 
       $sheetname = "Sheet1", $firstcell = "A1", $lastcell = "A1")


function Read-Excel ($path = $(throw "Enter full path to Excel spreadsheet file."), 
                     $sheetname = "Sheet1", $firstcelll = "A1", $lastcell = "A1") 
{
    # Assume present working directory if full path to file not given.
    if ($path -notmatch "\\") { $path = "$pwd\$path" }
 
    $excel = new-object -com "Excel.Application"
    $excel.Visible = $false
    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false 

    $workbooks = $excel.workbooks
    $workbook = $workbooks.open($path)
    $activesheet = $workbook.worksheets.item($sheetname)

    $range = $activesheet.range($firstcell,$lastcell)
    $range.Value2 

    # Encourage the garbage collector to kill the process quicker.
    ForEach ($item In $workbooks) { $item.Saved = $true ; $item.Close() }
    $excel.Quit()
    $range = $activesheet = $workbook = $workbooks = $excel = $null
}


Read-Excel -path $path -sheetname $sheetname -firstcell $firstcell -lastcell $lastcell


