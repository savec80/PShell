[xml]$XMLDocument = Get-Content "C:\Docs\Projects\Python\Parsing\DB12_Analyses.xml"
$path2save = "C:\Docs\Projects\Python\Parsing\Multiple Analyses.txt"
foreach ($analises in $XMLDocument.bes.Analysis) {
Try {
    if ($analises.Relevance) {
        $title = $analises.Title
        $title + ':' | Out-File -Append -FilePath $path2save
        foreach ($relevance in $analises.Relevance) {
            if ($relevance.'#cdata-section') {
                $resultRelevance += $relevance.'#cdata-section'
                $relevance.'#cdata-section'.Trim() | Out-File -Append -FilePath $path2save
            }
            else {
                $resultRelevance += $relevance
                $relevance.Trim() | Out-File -Append -FilePath $path2save
            }
        }    
    }
    }
Catch {
    continue
}
}