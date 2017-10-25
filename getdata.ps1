function get-systeminfo {
<#
.Synopsis
   Short description
   get-systeminfo
.DESCRIPTION
   Long description
.EXAMPLE
   get-systeminfo -host localhost
   Example of how to use this cmdlet
.EXAMPLE
    "localhost" | get-systeminfo
   Another example of how to use this cmdlet
.PARAMETER LogErrors
   Switch
.PARAMETER ComputerNames
   Mandatory
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
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
            $data_CS = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $Computer `
                     | Select-Object -Property Name, Workgroup, @{Name="AdminPasswordStatus";Expression={if ($_.AdminPasswordStatus -eq 1){"Disabled"}`
                                                          elseif ($_.AdminPasswordStatus -eq 2){"Enabled"} elseif ($_.AdminPasswordStatus -eq 3){"NA"}`
                                                          else{"Unknown"}}}, Model, Manufacturer
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
        } #foreach
    } #process
    end {}
}

function get-diskdetails {
    [cmdletbinding()]
    Param ([Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [string[]]$computernames = "localhost",
        [string]$errorlog = 'c:\temp\errors.log'
    )
    begin {}
    Process {
        foreach ($computer in $computernames) {
            Write-Verbose "connecting to $computer"
            $disks = Get-CimInstance -ClassName Win32_Volume -Filter "DriveType=3"
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
        }
    }
    end {}
}

function get-servicedetails {
    [cmdletbinding()]
    Param (
        [string[]]$computernames = "localhost",
        [string]$errorlog = 'c:\temp\errors.log'
    )
    begin {}
    Process {
        foreach ($computer in $computernames) {
            $services = Get-CimInstance -ClassName win32_service -Filter 'State="Running"'
            foreach ($service in $services) {
              $proc_data = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId=$($service.ProcessId)"
              $props = @{computername = $computer;
                         ThreadCount = $proc_data.ThreadCount;
                         ProcessName = $proc_data.ProcessName;
                         Name = $service.name;
                         VMSize = $proc_data.VirtualSize;
                         PeakPageFile = $proc_data.PeakPageFileUsage;
                         DisplayName = $service.DisplayName}
              $data = New-Object -TypeName psobject -Property $props
              Write-Output $data
            }
        }
    }
    end {}
}

#get-servicedetails
#Write-Host "-----pipeline mode-----"
#'localhost', 'localhost', 'localhost' | get-systeminfo -Verbose
Write-Host "-----param mode-----"
get-systeminfo -Verbose
#get-diskdetails