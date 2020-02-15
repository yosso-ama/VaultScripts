<#
.SYNOPSIS
Copy a numbering scheme in vault.

.DESCRIPTION
The Copy-NumSchm.ps1 script let the user select a numbering scheme in 
vault and copy it with a new name.

.PARAMETER VaultUser
Specifies the vault user name to login.

.PARAMETER VaultPassword
Specifies the password for the vault user to login.

.PARAMETER VaultServer
Specifies the vault server to connect.

.PARAMETER Vault
Specifies the vault name to login.

.PARAMETER VaultVersion
Specifies the installed vault version.

.INPUTS
None. You cannot pipe objects to Copy-NumSchm.ps1.

.OUTPUTS
None. Copy-NumSchm.ps1 does not generate any output.

.EXAMPLE
PS> .\Copy-NumSchm.ps1 -VaultUser administrator -VaultServer localhost -Vault Vault
This example login to vault with no password and assuming vault version 2019.

.EXAMPLE
PS> .\Copy-NumSchm.ps1 -VaultUser administrator -VaultPassword p@ssw0rd -VaultServer localhost -Vault Vault -VaultVersion 2019
This example login to vault with password and assuming vault version 2019.

#>

Param(
    [Parameter(Mandatory = $True)]
    [string]$VaultUser,
    [string]$VaultPassword,
    [Parameter(Mandatory = $True)]
    [string]$VaultServer,
    [Parameter(Mandatory = $True)]
    [string]$Vault,
    [string]$VaultVersion = "2019"
)

$buildNumbers = @{
    "2016" = "21.0"
    "2017" = "22.0"
    "2018" = "23.0"
    "2019" = "24.0"
}

Write-Host

if ($VaultVersion -notin $buildNumbers.Keys)
{
    Write-Host -ForegroundColor Red "Vault $VaultVersion is not supported.`n"
    return
}

$buildNumber = $buildNumbers[$VaultVersion]

$installLocation = Get-ItemPropertyValue "HKLM:\SOFTWARE\Autodesk\PLM\Autodesk Vault Professional $buildNumber\Common\" "InstallLocation" -ErrorAction SilentlyContinue
if (!$installLocation)
{
    Write-Host -ForegroundColor Red "Vault $VaultVersion is not installed.`n"
    return
}

try
{
    Add-Type -Path "${installLocation}Explorer\Autodesk.Connectivity.WebServices.dll"
}
catch
{
    Write-Host -ForegroundColor Red "Vault $VaultVersion may not be properly installed.`n"
    return
}

$mServer = New-Object Autodesk.Connectivity.WebServices.ServerIdentities
$mServer.DataServer = $VaultServer
$mServer.FileServer = $VaultServer

$login = New-Object Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials $mServer, $Vault, $VaultUser, $VaultPassword

try
{
    $serviceManager = New-Object Autodesk.Connectivity.WebServicesTools.WebServiceManager $login
}
catch
{
    Write-Host -ForegroundColor Red "Failed to log in, please check your parameters.`n"
    return
}

$documentService = $serviceManager.DocumentService

$numSchms = $documentService.GetNumberingSchemesByType("All")
if (!$numSchms)
{
    Write-Host -ForegroundColor Red "There is no numbering scheme in your Vault!`n"
    return
}

Write-Host "The numbering schemes currently in your Vault:"

$Props = @(
    @{ Label = "Id"; Expression = { $_.SchmID } }
    "Name"
    @{ Label = "Is Active"; Expression = { $_.IsAct } }
    @{ Label = "Is Default"; Expression = { $_.IsDflt } }
    @{ Label = "Is Used"; Expression = { $_.IsInUse } }
)
$numSchms | Format-Table $Props

do
{
    try
    {
        [long]$schmId = Read-Host -Prompt "Enter the id of the numbering scheme you want to copy"
    }
    catch { }
} while ($schmId -notin $numSchms.SchmID)

$numSchm = $numSchms | Where-Object SchmID -eq $schmId

Write-Host "`nThe numbering scheme you select is: " -NoNewline
Write-Host -ForegroundColor Green "$($numSchm.Name)`n"

do
{
    $newSchmName = Read-Host -Prompt "Enter the name for the new numbering scheme"
    if ($newSchmName -in $numSchms.Name)
    {
        Write-Host -ForegroundColor Yellow "`nThe name is already used.`n"
        $newSchmName = ""
    }
} while (!$newSchmName)

try
{
    $newSchm = $documentService.AddNumberingScheme($newSchmName, $numSchm.FieldArray, $numSchm.ToUpper)
    $newSchm = $documentService.ActivateNumberingSchemes($newSchm.SchmID)[0]

    Write-Host -ForegroundColor Green "`nThe numbering scheme is successfully copied.`n"
}
catch
{
    Write-Host -ForegroundColor Red "`nFailed to copy numbering scheme.`n"
}
