[CmdletBinding()]
param()
$data = Import-Csv 'C:\Users\avsavenkov\Documents\Mind\powershell\PShell\Lunch\data.csv'
Write-Debug "Imported csv data"
$totalqty  = 0
$totalsold = 0
$totalbought = 0
foreach ($line in $data) {
    if ($line.transaction -eq 'buy') {
        Write-Debug "ended buy trans (we sold)"
        $totalqty -= $line.qty
        $totalsold += $line.total
    }
    else {
        $totalqty += $line.qty
        $totalbought += $line.total
        Write-Debug "ended sell trans (we bought)"
    }
}
Write-Debug "output: $totalqty, $totalbought, $totalsold`
                        $($totalbought-$totalsold)"
"totalqty, totalbought, totalsold, totalamt" | Out-File 'C:\Users\avsavenkov\Documents\Mind\powershell\PShell\Lunch\data.txt'
"$totalqty, $totalbought, $totalsold, $($totalbought-$totalsold)" | Out-File 'C:\Users\avsavenkov\Documents\Mind\powershell\PShell\Lunch\data.txt' -Append