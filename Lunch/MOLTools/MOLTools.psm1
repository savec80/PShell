$MOLErrorLogPreference = 'c:\temp\mol-retries.txt'
$MOLConnectionString = "server=localhost\SQLEXPRESS;database=inventory;trusted_connection=True"

Import-Module MOLDatabase

function Get-MOLComputerNamesFromDatabase {
<#
.SYNOPSIS
Read computernames from database
#>
    Get-MOLDatabaseData -ConnectionString $MOLConnectionString -isSQLServer -query "SELECT computername FROM computers"
}

function Set-MOLInventoryInDatase {
<#
Accept the output of GET_MOLSystemInfo and saves the result back to database
#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [object[]]$inputObject
    )
    Process {
        foreach ($obj in $inputObject) {
            $query = "UPDATE computers SET
                      osversion = '$($obj.osversion)',
                      spversion = '$($obj.spversion)',
                      manufacturer = '$($obj.manufacturer)',
                      model = '$($obj.model)'
                      WHERE computername = '$($obj.computername)'"
            Write-Verbose "query is: $query"
            Invoke-MOLDatabaseQuery -connection $MOLConnectionString -isSQLServer -query $query
        }
    }
}

function get-MOLsysteminfo {
    [cmdletbinding()]
    Param (
        [Parameter(
            Mandatory=$True,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName=$True,
            HelpMessage="Computer Name or IP address")]
        [Alias('hostname')]
        [ValidateCount(1,10)]
        [string[]]$ComputerNames = "localhost",
        
        [string]$errorlog = "$MOLErrorLogPreference",

        [switch]$LogErrors
    )
    begin {
        Write-verbose "errorlog is $errorlog"
    }
    process {
        Write-Verbose "beggining process block"
        foreach ($computer in $ComputerNames) {
            Write-Verbose "getting data from $computer"
            try {
                $semaphore = $True
                $data_CS = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop -ErrorVariable err `
                         | Select-Object -Property Name, Workgroup, @{Name="AdminPasswordStatus";Expression={if ($_.AdminPasswordStatus -eq 1){"Disabled"}`
                                                              elseif ($_.AdminPasswordStatus -eq 2){"Enabled"} elseif ($_.AdminPasswordStatus -eq 3){"NA"}`
                                                              else{"Unknown"}}}, Model, Manufacturer
            } Catch {
                $semaphore = $False
                Write-Warning "$computer failed"
                Write-Warning "error: $err.message"
                if ($LogErrors) {
                    $computer | Out-File $errorlog -Append
                    Write-Warning "error logged to $errorlog"
                }
            }

            if ($semaphore) {
                $data_bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $Computer
                $data_OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $Computer
                #Add-Member -InputObject $data_CS -MemberType NoteProperty -Name BIOS_Version -Value $data_bios.version -Force
                #Add-Member -InputObject $data_CS -MemberType NoteProperty -Name OS_Version -Value $data_os.version -Force
                #Add-Member -InputObject $data_CS -MemberType NoteProperty -Name OS_ServicePackMajorVersion -Value $data_os.ServicePackMajorVersion -Force        
                $props = @{'ComputerName' = $computer;
                          'OSVersion' = $data_OS.version;
                          'Workgroup' = $data_CS.Workgroup;
                          'SPVersion' = $data_OS.ServicePackMajorVersion;
                          'BIOSVersion' = $data_bios.Version;
                          'Manufacturer' = $data_CS.Manufacturer;
                          'Model' = $data_CS.Model;
                          'AdminPasswordStatus' = $data_CS.AdminPasswordStatus}
                Write-Verbose "wmi query complete"
                $data = New-Object -TypeName psobject -Property $props
                $data.PSobject.TypeNames.Insert(0, 'MOL.ComputerSystemInfo')
                Write-Output -InputObject $data
            }
        } #foreach
    } #process
    end {}
}

function get-MOLdiskdetails {
    [cmdletbinding()]
    Param ([Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [string[]]$computernames = "localhost",
        [string]$errorlog = 'c:\temp\errors.log',
        [Switch]$LogError
    )
    begin {}
    Process {
        foreach ($computer in $computernames) {
            Write-Verbose "connecting to $computer"
            try {
                $disks = Get-CimInstance -ClassName Win32_Volume -ComputerName $computer -Filter "DriveType=3" -ErrorAction Stop
                foreach ($disk in $disks) {
                $FreeSpace = "{0:N2}" -f ($disk.FreeSpace/1Gb)
                $Size = "{0:N2}" -f ($disk.Capacity/1Gb)
                $props = @{computername = $computer;
                            Drive = $disk.Name;
                            Freespace = $FreeSpace;
                            Size = $Size}
                $data = New-Object -TypeName psobject -Property $props
                $data.PSobject.TypeNames.Insert(0, 'MOL.DiskInfo')
                Write-Output $data
                }
                Remove-Variable $disks
             } Catch {
                if ($LogError) {
                    Write-Verbose "loggind error to $errorlog"
                    $msg = "failed to get data from $computer. $($_.Exception.Message)"
                    Write-Error $msg
                    $computer | Out-File -FilePath $errorlog -Append
                    }
                }
        }
    }
    end {}
}

function get-MOLservicedetails {
    [cmdletbinding()]
    Param ([Parameter(Mandatory=$true)]
        [string[]]$computernames = "localhost",
        [string]$errorlog = 'c:\temp\errors.log',
        [switch]$LogErrors
    )
    begin {
        Write-Verbose "start processing"
        if ((test-path $errorlog) -and $LogErrors) {Remove-Item $errorlog}
    }
    Process {
        foreach ($computer in $computernames) {
            try {
                Write-Verbose "connect to $computer"
                $services = Get-CimInstance -ClassName win32_service -ComputerName $computer -Filter 'State="Running"' -ea Stop
                foreach ($service in $services) {
                  $props = @{
                    ComputerName = $services[0].SystemName
                    DisplayName = $service.DisplayName
                    Name = $service.name
                  }
                  Write-Verbose "process $service"
                  $proc_data = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId=$($service.ProcessId)" -ErrorAction Stop
                  $props.Add("ThreadCount", $proc_data.ThreadCount)
                  $props.Add("ProcessName", $proc_data.ProcessName)
                  $props.Add("VMSize", $proc_data.VirtualSize)
                  $props.Add("PeakPageFile", $proc_data.PeakPageFileUsage)
                  $data = New-Object -TypeName psobject -Property $props
                  $data.PSobject.TypeNames.Insert(0, 'MOL.ServiceInfo')
                  Write-Output $data
                } #foreach
            } catch {
                Write-Verbose "failed connect to $computer"
                $msg = "error message is: $_.Exception.Message"
                Write-Error $msg
                if ($LogErrors) {
                    Write-Verbose "logging error to $errorlog"
                    $computer | Out-File -FilePath $errorlog -Append
                }
            } #catch
        } #foreach
    }
    end {}
}

function Get-RemoteSMBShare {
<#
.EXAMPLE
get-remotesmbshare -computername localhost
#>
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$True)]
        [Alias('hostname')]
        [ValidateCount(1,5)]
        [string[]]$computername,

        [string]$erropath = 'c:\temp\MOLshareErr.txt',

        [switch]$LogError
    )
    foreach ($computer in $computername) {
        try {
            $data = Invoke-Command -ScriptBlock {gwmi -ClassName win32_share} -ComputerName $computer -ErrorAction Stop 
        } Catch {
            Write-Verbose "fail to connect to $computer"
            Write-Verbose "error is: $_.Exception.Message"
            if ($LogError) {
                Out-File -FilePath $erropath -InputObject $computer -Append
            }
        }
        if ($data) {$data; Remove-Variable data}
    }
}

function Restart-MOLComputer {
    [CmdletBinding(SupportsShouldProcess=$true,
                   ConfirmImpact='High')]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [string[]]$Computername
    )
    PROCESS {
        foreach ($Computer in $Computername) {
            Invoke-WMIMethod -Class Win32_OperatingSystem -Name reboot -ComputerName $Computer
        }
    }
}

function Set-MOLServicePassword {
    [CmdletBinding(SupportsShouldProcess=$true,
                   ConfirmImpact='Medium')]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [string[]]$ComputerName,

        [Parameter(Mandatory=$true)]
        [string]$ServiceName,

        [Parameter(Mandatory=$true)]
        [string]$NewPassword
    )
    PROCESS {
        foreach ($computer in $ComputerName) {
            $svcs = Get-WmiObject -ComputerName $computer -Filter "name='$ServiceName'" -Class win32_service
            foreach ($svc in $svcs) {
                if ($PSCmdlet.ShouldProcess("$svc on $computer")) {
                    $svc.Change($null,
                                $null,
                                $null,
                                $null,
                                $null,
                                $null,
                                $null,
                                $NewPassword) | Out-Null
                    
                }
            }
        }
    }
}
Export-ModuleMember -Variable MOLErrorLogPreference, MOLConnectionString
Export-ModuleMember -Function Get-MOLSystemInfo, get-MOLservicedetails, get-MOLdiskdetails, Set-MOLInventoryInDatase, 
                              Get-MOLComputerNamesFromDatabase, Get-RemoteSMBShare, Restart-MOLComputer, Set-MOLServicePassword 
#'localhost', 'localhost', 'localhost' | get-systeminfo -Verbose
#Write-Host "-----param mode-----"
#get-systeminfo localhost  -LogErrors
#get-diskdetails localhost
#get-servicedetails localhost | ft
#Update-FormatData -PrependPath 'C:\Docs\Projects\PShell\Lunch\CustomViewA.format.ps1xml'