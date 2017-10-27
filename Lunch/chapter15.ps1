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

Get-RemoteSMBShare -computername localhost, test, localhost -LogError
#'localhost', 'localhost' | Get-RemoteSMBShare
#Get-RemoteSMBShare