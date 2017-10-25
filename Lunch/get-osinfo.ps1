    Param
    (
        # Param1 help description
        [string]
        $ComputerName = "localhost"
    )
    Get-CimInstance -ClassName Win32_OperatingSystem `
         -computername $ComputerName
