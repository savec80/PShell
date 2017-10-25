foreach ( $KeyPath in $RegPath )
{
    $_Path = (invoke-command {Get-ChildItem $KeyPath -Recurse -ErrorAction SilentlyContinue})

    ForEach ($key in $_Path)
    {

        $local = @()
        $remote = @()
        ForEach ( $Property in $key.Property )
        {           
            if ( $Key.Name -like "*HKEY_LOCAL_MACHINE*" )
            {
                $KeyPath = [regex]::Replace($key.Name,[regex]"HKEY_LOCAL_MACHINE\\","HKLM:")
            } elseif ( $Key.Name -like "*HKEY_CURRENT_USER*" ) {
                $KeyPath = [regex]::Replace($key.Name,[regex]"HKEY_CURRENT_USER\\","HKCU:")
            } else {
                Write-Host "I was unable to set the Registry Hive path, exiting........"
                exit 1
            }

            $lentry = (invoke-command {Get-ItemProperty -Path $KeyPath -Name $Property})
            $lfound = New-Object -TypeName PSObject
            $lfound | Add-Member -Type NoteProperty -Name Path -Value $KeyPath
            $lfound | Add-Member -Type NoteProperty -Name Name -Value $Property
            $lfound | Add-Member -Type NoteProperty -Name Data -Value $lentry.$Property

            $local += $lfound

            $rentry = (invoke-command -computername remoteservername -Script {Get-ItemProperty -Path $args[0] -Name $args[1]} -Args $KeyPath, $Property -ErrorVariable errmsg 2>$null)

            if ( $errmsg -like "*does not exist at path*" )
            {
                $Value = "KEY IS MISSING"
            } else {
                $Value = $rentry.$Property
            }

            $rfound = New-Object -TypeName PSObject
            $rfound | Add-Member -Type NoteProperty -Name Path -Value $KeyPath
            $rfound | Add-Member -Type NoteProperty -Name Name -Value $Property
            $rfound | Add-Member -Type NoteProperty -Name Data -Value $Value

            $remote += $rfound

        }
        $compare = Compare-Object -ReferenceObject $local -DifferenceObject $remote -Property Path,Name,Data
        $results += $compare
    }
}