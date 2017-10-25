#Get a list of files to copy from
$Files = GCI 'c:\Users\avsavenkov\Documents\Mind\powershell\Scripts\Excel\' | ?{$_.Extension -Match "xlsx?"} | select -ExpandProperty FullName
#$File = GCI 'c:\Users\avsavenkov\Documents\Mind\powershell\Scripts\Excel\iem_tscm_data_4_20160613.xlsx' | ?{$_.Extension -Match "xlsx?"} | select -ExpandProperty FullName

#Launch Excel, and make it do as its told (supress confirmations)
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True
$Excel.DisplayAlerts = $False
$xlCellTypeLastCell = 11
#Open up a new workbook
$Dest = $Excel.Workbooks.Add()

#Loop through files, opening each, selecting the Used range, and only grabbing the first 6 columns of it. Then find next available row on the destination worksheet and paste the data
ForEach($File in $Files[0..4]){
    
  #  $Rng = $ws.UsedRange | Out-Null


 #   If(($Dest.ActiveSheet.UsedRange.Count -eq 1) -and ([String]::IsNullOrEmpty($Dest.ActiveSheet.Range("A1").Value2))){ #If there is only 1 used cell and it is blank select A1
 #       [void]$source.ActiveSheet.Range("A1","F$(($Source.ActiveSheet.UsedRange.Rows|Select -Last 1).Row)").Copy()
 #       [void]$Dest.Activate()
 #       [void]$Dest.ActiveSheet.Range("A1").Select()
 #   }Else{ #If there is data go to the next empty row and select Column A
 #       [void]$source.ActiveSheet.Range("A2","F$(($Source.ActiveSheet.UsedRange.Rows|Select -Last 1).Row)").Copy()
  $objWorksheet = $Dest.Worksheets(1)

$objWorksheet.Activate()
$objRange = $objWorksheet.UsedRange
$objRange.SpecialCells($xlCellTypeLastCell).Activate()
$intNewRow = $Excel.ActiveCell.Row + 1
$objWorksheet.Range("A$intNewRow").Select()

$Source = $Excel.Workbooks.Open($File,$true,$true)
    $ws = $Source.WorkSheets.item('Squirrel SQL Export')
    $ws.activate()
    Start-Sleep 1
    $ws.AutoFilterMode = $False
    $ws.Range("L1").AutoFilter(12,'=',7)
    $ws.Range("AD1").AutoFilter(30,'=',7)
    $ws.UsedRange.Delete()
    $ws.AutoFilterMode = $False
    
  $Source.ActiveSheet.UsedRange.Select()
$Source.ActiveSheet.UsedRange.Copy()

        [void]$Dest.Activate()
 #       [void]$Dest.ActiveSheet.Range("A$(($Dest.ActiveSheet.UsedRange.Rows|Select -last 1).row+1)").Select()
 #   }
    [void]$Dest.ActiveSheet.Paste()
    $Source.Close()
}
$Dest.SaveAs("c:\Users\avsavenkov\Documents\Mind\powershell\Scripts\Output\test.xlsx",51)
$Dest.close()
$Excel.Quit()