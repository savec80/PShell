<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function set-pressuredata
{
    [CmdletBinding()]
    [Alias("press")]
    Param
    (
        # Input your blood pressure in format: systolic (maximum)/diastolic (minimum)/Pulse
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $rightarm,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $leftarm,

        # Path to save data in csv
        [String]
        $Path="$home\pressure",
        
        # Format
        [ValidateSet(“Csv”,”Excel”)] 
        $Format="Excel",
        
        # Format
        [ValidateSet(“Morning”,”Evening”)] 
        $time
    )

    Begin
    {
    [psCustomObject]$table = @{'RightArm'=$rightarm;
                               'LeftArm'=$leftarm
                              }
     $xlsFile = 'c:\Users\avsavenkov\Documents\Mind\powershell\press.xlsx'
            $Excel = New-Object -ComObject Excel.Application
            $excel.DisplayAlerts = $False
            If (Test-Path $xlsFile) {$ExcelWorkBook = $Excel.Workbooks.Open($XLSFile)}
            else                    {$ExcelWorkBook = $excel.Workbooks.Add()}
            #$ExcelWorkBook = $Excel.Workbooks.Open($XLSFile)
            $ExcelWorkSheet = $ExcelWorkBook.sheets.item("Sheet1")
            $ExcelWorkSheet.Activate()
    }
    Process
    {
    try {
     switch ($Format)
     {
      "csv" {
                Get-Date | Export-Csv "$Path.csv" -NoTypeInformation -Append

                $table.GetEnumerator()| select | Export-Csv "$Path.csv" -NoTypeInformation -Append
                break
            }
      "Excel" {
  
# Go to the first empty row
$LastRow = $ExcelWorkSheet.UsedRange.rows.count

$date  =$(get-date).ToShortDateString()
$dateR  =$date +" R"
$dateL  =$date+" L"
if ($ExcelWorkSheet.cells.Item($lastRow,1).text -eq $dateL) {
    $lastRowR = $lastRow-1
    $lastRowL = $lastRow
    if ($time -eq "Morning") {
        $ExcelWorkSheet.cells.Item($LastRowR,1) = $dateR
        $ExcelWorkSheet.cells.Item($LastRowL,1) = $dateL
        $ExcelWorkSheet.cells.Item($LastRowL,2) = $table.LeftArm
        $ExcelWorkSheet.cells.Item($LastRowR,2) = $table.RightArm
    }
    elseif ($time -eq "Evening") {
        $ExcelWorkSheet.cells.Item($LastRowR,1) = $dateR
        $ExcelWorkSheet.cells.Item($LastRowL,1) = $dateL
        $ExcelWorkSheet.cells.Item($LastRowL,3) = $table.LeftArm
        $ExcelWorkSheet.cells.Item($LastRowR,3) = $table.RightArm
    }
    Else {break}
}
else {
    $lastRowR = $lastRow+1
    $lastRowL = $lastRowR+1
    if ($time -eq "Morning") {
        $ExcelWorkSheet.cells.Item($LastRowR,1) = $dateR
        $ExcelWorkSheet.cells.Item($LastRowL,1) = $dateL
        $ExcelWorkSheet.cells.Item($LastRowL,2) = $table.LeftArm
        $ExcelWorkSheet.cells.Item($LastRowR,2) = $table.RightArm
    }
    elseif ($time -eq "Evening") {
        $ExcelWorkSheet.cells.Item($LastRowR,1) = $dateR
        $ExcelWorkSheet.cells.Item($LastRowL,1) = $dateL
        $ExcelWorkSheet.cells.Item($LastRowL,3) = $table.LeftArm
        $ExcelWorkSheet.cells.Item($LastRowR,3) = $table.RightArm
    }
    Else {break}

    }
#$ExcelWorkBook.Save()
#$ExcelWorkBook.SaveAs($xlsFile)
#$ExcelWorkBook.Close()
#$Excel.Quit()
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
#Remove-Variable $Excel
              
#                  [void]$wkbk.PSBase.GetType().InvokeMember("SaveAs",[Reflection.BindingFlags]::InvokeMethod, $null, $wkbk, $sfile, $newci)
#                  [void]$wkbk.PSBase.GetType().InvokeMember("Close",[Reflection.BindingFlags]::InvokeMethod, $null, $wkbk, 0, $newci)
#                  $xl.Quit()


              }
     }
    
} catch {write-host "some error"}
    }
    End
    {
    
        $ExcelWorkBook.SaveAs($xlsFile)
        
        $ExcelWorkBook.Close()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWorkSheet)
        $Excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
        [System.GC]::Collect()
    }
}

