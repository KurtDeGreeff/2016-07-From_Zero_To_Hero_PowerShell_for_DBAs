#0% Um Script nao é reutilizavel
# Obrigado passar uma lista de servidores
# Usando Arrays pra output - powershell faz streaming
# Nao testa o Open pra ver se ta aberto ou nao .. fazer o teste pra eles verem
# Mostrar o Write-Host nso escreve no pipeline - Usar Get-Process | format-list e Format-table
# Se eu precisar exportar prfa csv tenho que modificar o script



[Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
[Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.ConnectionInfo')
$Output = @()
Get-Content C:\temp\Servers.txt |
Foreach-object { 
  $InstanceName = $_ -replace '\\' ,'_'
  $Server=New-Object 'Microsoft.SqlServer.Management.Smo.Server' $_
  $Object = [pscustomobject][ordered]@{
    ServerName = $_
    EngineEdition   = $Server.EngineEdition
    Product = $Server.Product
    ProductLevel = $Server.ProductLevel 
    Version = $server.Version
                      
  }
  $Output += $Object
}
Write-host $Output | Format-List

###########################################################################################################################################################
#25% Dividir as funções Get-SQLConnection - Get-SQLServer
#Colocar a assembly na Get-SQLConnection

function Get-SqlConnection ([string]$sqlserver)
{
  $con = new-object ('Microsoft.SqlServer.Management.Common.ServerConnection') $sqlserver 
	
  $con.Connect()

  Write-Output $con
    
} 

#mostrar o error e dizer que depois trataremos com try-catch - nao é um terminating error
<#'deathstsar'|
    Foreach-object {
    Get-SQLConnection $_ 
    }
#>
 

function Get-SqlServer ([string]$sqlserver)
{

  $con = Get-SqlConnection $sqlserver 
  $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
     
  $Output = [pscustomobject][ordered]@{
    ServerName = $server.Name
    EngineEdition   = $server.EngineEdition
    Product = $server.Product
    ProductLevel = $server.ProductLevel 
    Version = $server.Version
                      
  }

  Write-output $Output 
    
} 



Get-SQLServer DeathStar

Get-content C:\temp\Servers.txt |
ForEach-Object {
  
  Get-SQLServer $_
}

#ta mais ai seu chefe pede pra adicoinar mais propriedades.. voce tera que novamente alterar seu script
#vamos retornar diretamente o objeto livo

function Get-SqlServer ([string]$sqlserver)
{

  $con = Get-SqlConnection $sqlserver 
  $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
     
  Write-output $server 
    
} 

Get-SQLServer DeathStar |
Select-Object name,
EngineEdition,
Product,
ProductLevel,
Version

#melhor ainda posso exportar para csv

Get-content C:\temp\Servers.txt |
ForEach-Object {
  
  $ServerInstance = $_ -replace '\\','_'
  Get-SQLServer $_ |
  Select-Object name,
                EngineEdition,
                Product,
                ProductLevel,
                Version |
  Export-Csv "c:\temp\VirtualPassPT\$($ServerInstance).csv" -NoTypeInformation -noclobber
 
  
}

#carregar numa janela separada e mostrar
#C:\temp\VIRTUALPASSPT\Convert-CSVToExcel.ps1  -inputfile (Get-ChildItem c:\temp\VirtualPassPT\*.csv)		-output   c:\temp\VirtualPassPT\Servidores.xlsx 



###########################################################################################################################################################

#Indo pra 50% parametros a passando pra advanbced function

function Get-SqlConnection  {

  param(
          [Parameter(  Position=0, 
                       Mandatory)]
          [string]$sqlserver
       )
  
  
  $con = new-object ('Microsoft.SqlServer.Management.Common.ServerConnection') $sqlserver 
	
  $con.Connect()

  Write-Output $con
    
} 

#mostrar o error e dizer que depois trataremos com try-catch - nao é um terminating error
<#'deathstsar'|
    Foreach-object {
    Get-SQLConnection $_ 
    }
#>
 

function Get-SqlServer {

  param(
          [Parameter(  Position=0, 
                       Mandatory)]
          [string]$sqlserver
       )

  $con = Get-SqlConnection $sqlserver 
  $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
     
  Write-output $server 

    
} 


#Aceitando Pipeline Input

function Get-SqlServer {

  [cmdletbinding()]

  param(
          [Parameter(  Position=0, 
                       Mandatory,
                       ValueFromPipeline)]
          [string]$sqlserver
       )
       
  Begin {}     
  Process { 
    $con = Get-SqlConnection $sqlserver 
    $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
     
    Write-output $server 
                      
  }

  End{}
    
} 

'DeathStar','DeathStar\I2012' | 
Get-SqlServer |
Select-Object name,
              EngineEdition,
              Product,
              ProductLevel,
              Version 
              
#aceitando array no sqlserver

#mostrar o get-service

function Get-SqlServer {

  [cmdletbinding()]

  param(
          [Parameter(  Position=0, 
                       Mandatory,
                       ValueFromPipeline)]
          [string[]]$sqlserver
       )
       
  Begin {}     
  Process { 
    foreach ($SQL in $sqlserver) { 
      $con = Get-SqlConnection $SQL 
      $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
      Write-output $server 
    }
                      
  }

  End{}
    
} 

Get-SqlServer 'DeathStar','DeathStar\I2012' |
Select-Object name,
              EngineEdition,
              Product,
              ProductLevel
              
Get-Content C:\temp\VIRTUALPASSPT\Servers.txt |
Get-SqlServer  |
Select-Object name,
              EngineEdition,
              Product,
              ProductLevel

###########################################################################################################################################################
 #To 75 Error Handle
 #agora temos advanced functions acesso a commom parameters

 function Get-SqlConnection  {

   param(
          [Parameter(  Position=0, 
                       Mandatory)]
          [string]$sqlserver
       )
  
   try {
  
     $con = new-object ('Microsoft.SqlServer.Management.Common.ServerConnection') $sqlserver 
	
     $con.Connect()

     Write-Output $con
   } catch {
     Write-Error "Error $($_.Exception.message)" 
        
    
   } 
 }   

 function Get-SqlServer {

     [cmdletbinding()]

     param(
          [Parameter(  Position=0, 
                       Mandatory,
                       ValueFromPipeline)]
          [string[]]$sqlserver
       )
       
     Begin {}     
     Process { 
       foreach ($SQL in $sqlserver) { 
         try { 
            $con = Get-SqlConnection $SQL 
            $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
            Write-output $server 
            
         } catch {
           Write-Error "Error $($_.Exception.message)" 
         }
                      
       }

     } 
     End{}

}
 
'iiiii','Deathstar' | 
 Get-SQLServer  |
 Select-Object name,
              EngineEdition,
              Product,
              ProductLevel
              
'iiiii','Deathstar' | 
 Get-SQLServer -ErrorAction Stop |
 Select-Object name,
              EngineEdition,
              Product,
              ProductLevel
              

#100% Help

 

 function Get-SqlServer {
 
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

     [cmdletbinding()]

     param(
          [Parameter(  Position=0, 
                       Mandatory,
                       ValueFromPipeline)]
          [string[]]$sqlserver
       )
       
     Begin {}     
     Process { 
       foreach ($SQL in $sqlserver) { 
         try { 
            $con = Get-SqlConnection $SQL 
            $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
            Write-output $server 
            
         } catch {
           Write-Error "Error $($_.Exception.message)" 
         }
                      
       }

     } 
     End{}

}
 
