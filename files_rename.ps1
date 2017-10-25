 $fold = Get-Item -Path C:\Docs\Projects\BFI\*
 foreach ($f in $fold) {
    $files = Get-Item -Path ($f.FullName+"\*") -Filter bravo*
        foreach ($file in $files) {
            $file -match "bravo_0_(.*)"
            Rename-Item -Path $file -NewName $Matches[1] -Force
        }
 }