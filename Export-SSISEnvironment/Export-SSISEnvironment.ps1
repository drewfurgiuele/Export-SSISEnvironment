<#
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