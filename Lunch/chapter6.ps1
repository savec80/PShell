function get-systeminfo {
    [cmdletbinding()]
    Param (
        [Parameter(
            Mandatory=$True,
            ValueFromPipeline=$True,
            HelpMessage="Computer Name or IP address")]
        [Alias('hostname')]
        [ValidateCount(1,10)]
        [string[]]$ComputerNames = "localhost",
        
        [string]$errorlog = "C:\temp\errors.log",

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
                          'SPVersion' = $data_OS.ServicePackMajorVersion;
                          'BIOSVersion' = $data_bios.Version;
                          'Manufacturer' = $data_CS.Manufacturer;
                          'Model' = $data_CS.Model;
                          'AdminPasswordStatus' = $data_CS.AdminPasswordStatus}
                Write-Verbose "wmi query complete"
                $data = New-Object -TypeName psobject -Property $props
                Write-Output -InputObject $data
            }
        } #foreach
    } #process
    end {}
}

function get-diskdetails {
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
                Write-Output $data
                }
                Remove-Variable $props
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

function get-servicedetails {
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

get-servicedetails
#Write-Host "-----pipeline mode-----"
#'localhost', 'localhost', 'localhost' | get-systeminfo -Verbose
#Write-Host "-----param mode-----"
#get-systeminfo test  -LogErrors
#get-diskdetails