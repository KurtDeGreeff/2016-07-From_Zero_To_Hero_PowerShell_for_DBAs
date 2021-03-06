﻿[Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
[Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.ConnectionInfo')


function Get-SqlConnection  {
  
  param (
  
    [parameter( position = 0,
                mandatory)]
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

  <#
      .SYNOPSIS
      Gets an SMO Server object.
      .DESCRIPTION
      The Get-SqlServer function  gets a SMO Server object for the specified SQL Server.
      .INPUTS
      None
      You can pipe objects to Get-SqlServer 
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

  param (
  
    [parameter( position = 0,
                mandatory,
                ValueFromPipeline)]
    [string[]]$sqlserver            
  
  )
  
  Begin{}
  
  Process{
  
    foreach ($sql in $sqlserver) { 
    
      try { 

        $con = Get-SqlConnection $sql 
        $server = new-object ('Microsoft.SqlServer.Management.Smo.Server') $con
     
        Write-output $server
      } catch {
      
        Write-Error "Error $($_.Exception.message)" 
      
      }
    }
  }
  
  End{}
  
    
} 

Function Convert-CSVToExcel
{
  <#   
      .SYNOPSIS  
      Converts one or more CSV files into an excel file.
     
      .DESCRIPTION  
      Converts one or more CSV files into an excel file. Each CSV file is imported into its own worksheet with the name of the
      file being the name of the worksheet.
       
      .PARAMETER inputfile
      Name of the CSV file being converted
  
      .PARAMETER output
      Name of the converted excel file
       
      .EXAMPLE  
      Get-ChildItem *.csv | ConvertCSV-ToExcel -output 'report.xlsx'
  
      .EXAMPLE  
      ConvertCSV-ToExcel -inputfile 'file.csv' -output 'report.xlsx'
    
      .EXAMPLE      
      ConvertCSV-ToExcel -inputfile @("test1.csv","test2.csv") -output 'report.xlsx'
  
      .NOTES
      Author: Boe Prox									      
      Date Created: 01SEPT210								      
      Last Modified:  
     
  #>
     
  #Requires -version 2.0  
  [CmdletBinding(
      SupportsShouldProcess = $True,
      ConfirmImpact = 'low',
      DefaultParameterSetName = 'file'
    )]
  Param (    
    [Parameter(
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        Position=0,
        Mandatory=$True,
     HelpMessage="Name of CSV/s to import")]
     [ValidateNotNullOrEmpty()]
    [string[]]$inputfile,
    [Parameter(
        ValueFromPipeline=$False,
        Position=1,
        Mandatory=$True,
     HelpMessage="Name of excel file output")]
     [ValidateNotNullOrEmpty()]
    [string]$output    
    )

  Begin {     

    Function Release-Ref ($ref) 
        {
            ([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            [System.__ComObject]$ref) -gt 0)
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers() 
        }

    #Configure regular expression to match full path of each file
    $COunter = 1
    $SName1 = @()
    [regex]$regex = "^\w\:\\"
    
    #Find the number of CSVs being imported
    $count = ($inputfile.count -1)
   
    #Create Excel Com Object
    $excel = new-object -com excel.application
    
    #Disable alerts
    $excel.DisplayAlerts = $False

    #Show Excel application
    $excel.Visible = $False

    #Add workbook
    $workbook = $excel.workbooks.Add()

    #Remove other worksheets
    #$workbook.worksheets.Item(2).delete()
    #After the first worksheet is removed,the next one takes its place
    #$workbook.worksheets.Item(2).delete()   

    #Define initial worksheet number
    $i = 1
    }

  Process {
    ForEach ($input in $inputfile) {
        #If more than one file, create another worksheet for each file
        If ($i -gt 1) {
            $workbook.worksheets.Add() | Out-Null
            }
        #Use the first worksheet in the workbook (also the newest created worksheet is always 1)
        $worksheet = $workbook.worksheets.Item(1)
        #Add name of CSV as worksheet name
      Write-Host "Processing $((GCI $input).basename)"
      $Sname = "$((GCI $input).basename)                                          "

      $Sname = $Sname.substring(0,15)
      if ($Sname -contains $SName1) {
        $Sname	= "$($Sname.substring(0,13))_$($Counter)"
        $COunter++
      }	
      $SName1	+=$Sname

		

        $worksheet.name = $Sname

        #Open the CSV file in Excel, must be converted into complete path if no already done
        If ($regex.ismatch($input)) {
            $tempcsv = $excel.Workbooks.Open($input) 
            }
        ElseIf ($regex.ismatch("$($input.fullname)")) {
            $tempcsv = $excel.Workbooks.Open("$($input.fullname)") 
            }    
        Else {    
            $tempcsv = $excel.Workbooks.Open("$($pwd)\$input")      
            }
        $tempsheet = $tempcsv.Worksheets.Item(1)
        #Copy contents of the CSV file
        $tempSheet.UsedRange.Copy() | Out-Null
        #Paste contents of CSV into existing workbook
        $worksheet.Paste()

        #Close temp workbook
        $tempcsv.close()

        #Select all used cells
        $range = $worksheet.UsedRange

        #Autofit the columns
        $range.EntireColumn.Autofit() | out-null
        $i++
        } 
    }        

  End {
    #Save spreadsheet
    $workbook.saveas("$output")

    Write-Host -Fore Green "File saved to $pwd\$output"

    #Close Excel
    $excel.quit()  

    #Release processes for Excel
    $a = Release-Ref($range)
    }
}


Export-ModuleMember -Function Convert-CSVToExcel,get-sqlserver 




