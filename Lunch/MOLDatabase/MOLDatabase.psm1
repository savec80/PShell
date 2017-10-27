function Get-MOLDatabaseData {
    [CmdletBinding()]
    Param(
        [string]$ConnectionString,
        [string]$query,
        [switch]$isSQLServer
    )

    if ($isSQLServer) {
        Write-Verbose "in SQL Server mode"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    }
    else {
        Write-Verbose "in OleDb mode"
        $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    }
    $connection.ConnectionString = $ConnectionString
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    if ($isSQLServer) {
        $adapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter $command
    }
    else {
        $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $command
    }
    $dataset = New-Object -TypeName System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    $dataset.Tables[0] | select -ExpandProperty computername
    $connection.Close()
}

function Invoke-MOLDatabaseQuery {
        [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
    Param(
        [string]$ConnectionString,
        [string]$query,
        [switch]$isSQLServer
    )

    if ($isSQLServer) {
        Write-Verbose "in SQL Server mode"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    }
    else {
        Write-Verbose "in OleDb mode"
        $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection
    }
    $connection.ConnectionString = $ConnectionString
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    if ($PSCmdlet.ShouldProcess($query)) {
        $connection.Open()
        $command.ExecuteNonQuery()
        $connection.Close()
    }
}