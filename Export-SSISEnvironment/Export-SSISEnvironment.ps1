<#
.SYNOPSIS
    Copies SSIS environment variables from one projec to another.
.DESCRIPTION
    This script will find all environment variables in a given folder inside a SSIS project and copy them to a target folder on the same (or any different) server.
.PARAMETER SourceServerName
    Aliases: ss
    The hostname of the SQL server you want to copy from. This is a required parameter.
.PARAMETER SourceInstanceName
    Aliases: si
    The instance name of the SQL server you want to copy from. Default value is "DEFAULT" for non-named SQL instances. This is an optional parameter.
.PARAMETER SourceFolder
    Aliases: sf
    The name of the integration services folder you want copy the environment variables from.
.PARAMETER SourceEnvironment
    Aliases: se
    The name of the environment you want to copy. This is an optional parameter; if left blank, all environments will be copied.
.PARAMETER TargetServer
    Aliases: ts
    Tha name of the server you want to copy the environment (variables) to.
.PARAMETER TargetInstance
    Aliases: ti
    The instance name of the SQL server you want to copy to. Default value is "DEFAULT" for non-named SQL instances. This is an optional parameter.
.PARAMETER TargetFolder
    Aliases: tf
    The name of the folder on the target SQL Server you want to copy the environment variables to. This is an optional parameter. If left off, it will copy the environment (variables) to same name as the source folder
.EXAMPLE
    Connect to a SQL server local instance of SQL Server and copy all the environments (and contained variables) from the folder 'ProjectFolder' to the remote server's integration services catalog in the same folder name.
    .\Export-SSISEnvironment.ps1 -SourceServer localhost -SourceFolder ProjectFolder -TargetServer remoteserver
.OUTPUTS
    None, unless -VERBOSE is specified. In fact, -VERBOSE is reccomended so you can see what is happening and when.
.NOTES
#>
param(
    [Alias("ss")] [Parameter(Mandatory=$true)]  [string]$SourceServer,
    [Alias("si")] [Parameter(Mandatory=$false)] [string]$SourceInstance = "DEFAULT",
    [Alias("sf")] [Parameter(Mandatory=$true)]  [string]$SourceFolder,
    [Alias("se")] [Parameter(Mandatory=$false)] [string]$sourceEnvironment,
    [Alias("ts")] [Parameter(Mandatory=$true)]  [string]$targetServer,
    [Alias("ti")] [Parameter(Mandatory=$false)] [string]$targetInstance = "DEFAULT",
    [Alias("tf")] [Parameter(Mandatory=$false)] [string]$targetFolder
)

$ssis = Get-ChildItem -Path ("SQLSERVER:\SSIS\" + $SourceServer) | Where-Object {$_.displayname -eq $sourceInstance}
$targetssis = Get-ChildItem -Path ("SQLSERVER:\SSIS\" + $TargetServer) | Where-Object {$_.displayname -eq $targetInstance}
$sourceCatalog = $ssis.Catalogs["SSISDB"]
$targetCatalog = $targetssis.Catalogs["SSISDB"]
if (!$targetFolder) { $targetFolder = $sourceFolder }

$environments = $ssis.Catalogs["SSISDB"].Folders[$SourceFolder].Environments # | Where-Object {$_.Name -eq $sourceEnvironment}
if ($sourceEnvironment) { $environments= $environments | Where-Object {$_.Name -eq $sourceEnvironment} }


foreach ($e in $environments)
{
    $e.Refresh()
    $variables = $e.Variables
    $variables.Refresh()

    if ($targetCatalog.Folders[$targetFolder] -eq $null)
    {
        Write-Verbose "No folder named $targetFolder exists! Creating a new one..."
	    $folder = New-Object Microsoft.SqlServer.Management.IntegrationServices.CatalogFolder ($targetCatalog, $targetFolder, $null)
	    $folder.Create()
    }
    else
    {
	    $folder = $targetCatalog.Folders[$targetFolder]
    }

    $targetEnvironment = $e.Name

    if ($targetCatalog.Folders[$targetFolder].Environments[$targetEnvironment] -eq $null)
    {
        Write-Verbose "No environment named $targetEnvironment found, creating a new one!"
	    $tenvironment = New-Object Microsoft.SqlServer.Management.IntegrationServices.EnvironmentInfo ($folder, $targetEnvironment, $null)
	    $tenvironment.Create()        
    }
    else
    {
        Write-Verbose "There is already an environment named $targetEnvironment; adding new variables only."
	    $tenvironment = $targetCatalog.Folders[$targetFolder].Environments[$targetEnvironment]
    }
    $tenvironment.Variables.Refresh()

    ForEach ($v in $variables)
    {
	    if ($tenvironment.Variables[$v.Name] -eq $null)
	    {
            Write-Verbose "Creating variable $($v.Name)..."
		    $tenvironment.Variables.Add($v.Name, $v.Type, $v.Value, $v.Sensitive, $v.Description)
	    }
    }
    $tenvironment.Alter()

}