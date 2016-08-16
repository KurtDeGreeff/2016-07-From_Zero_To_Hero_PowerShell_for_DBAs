[Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
[Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.ConnectionInfo')

function Get-SQLServer {

  <#
.SYNOPSIS
Gets an SMO Server object.
.DESCRIPTION
The Get-SqlServer function  gets a SMO Server object for the specified SQL Server.
.INPUTS
None
    You cannot pipe objects to Get-SqlServer 
.OUTPUTS
Microsoft.SqlServer.Management.Smo.Server
    Get-SqlServer returns a Microsoft.SqlServer.Management.Smo.Server object.
.EXAMPLE
Get-SqlServer "Z002\sql2K8"
This command gets an SMO Server object for SQL Server Z002\SQL2K8.
.EXAMPLE


Get-SqlServer "Z002\sql2K8" "sa" "Passw0rd"
This command gets a SMO Server object for SQL Server Z002\SQL2K8 using SQL authentication.
.LINK
Get-SqlServer 
#>



  param(
          [Parameter(  Position=0, 
                       Mandatory,
                       ValueFromPipeLine)]
          [string[]]$sqlserver
  )
  Begin {}

  Process {
    
    foreach ($SQL in $sqlserver) {
        try {
            $con = Get-SqlConnection $sql 
            $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
            Write-output $server 
        } catch {
            Write-Error "Error $($_.Exception.message)" 
        }       
    }

  }

  End {}
    
}



Function get-sqlconnection ([string]$sqlserver)  {
 $con = new-object ('Microsoft.SqlServer.Management.Common.ServerConnection') $sqlserver 
 $con.connect()
 Write-Output $con
}

Export-ModuleMember -Function Get-SQlServer
