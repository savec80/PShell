    function get-osinfo {
        Param
        (
            # Param1 help description
            [string]
            $ComputerName = "localhost"
        )
        Get-CimInstance -ClassName Win32_OperatingSystem `
                        -computername $ComputerNameS
    }

    function get-diskinfo {
        Param (
            [string]
            $computername = "localhost",
            [int]
            $minfreepercent = 10
        )
        $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "Drivetype=3"
        foreach ($disk in $disks) {
            $perFree = ($disk.FreeSpace/$disk.Size)*100
            if ($perFree -ge $minfreepercent) {
                $OK = $true
            }
            else {
                $OK = $false
            }
            $disk | Select DeviceID, VolumeName, Size, FreeSpace, `
                            @{Name="Ok";Expression={$OK}}, @{Name="FreePercent";Expression={$perFree}}
        }
    }

function new-drives {
    Param ()
    New-PSDrive -Name AppData -PSProvider FileSystem -Root $env:APPDATA -Scope script
    New-PSDrive -Name Temp -PSProvider FileSystem -Root $env:TEMP -Scope script

    $mydocs = Join-Path -Path $env:USERPROFILE -ChildPath Documents
    New-PSDrive -Name Docs -PSProvider FileSystem -Root $mydocs -Scope script
}

new-drives
dir temp: | Measure-Object -Property length -Sum