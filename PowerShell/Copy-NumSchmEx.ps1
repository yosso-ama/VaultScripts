<#
.SYNOPSIS
Copy or modify numbering scheme in vault

.DESCRIPTION
The Copy-NumSchmEx.ps1 script let the user select a numbering scheme in vault and copy or reset it, preserving used numbers (requires SQL Server connection). The user can modify the numbering scheme before reserving numbers.

.PARAMETER VaultUser
Specifies the vault user name to login.

.PARAMETER VaultPassword
Specifies the password for the vault user to login.

.PARAMETER VaultServer
Specifies the vault server to connect.

.PARAMETER Vault
Specifies the vault database name to login.

.PARAMETER SqlServer
Specifies the SQL Server to connect.

.PARAMETER SqlUser
Specifies the SQL user.

.PARAMETER SqlPassword
Specifies the SQL password.

.PARAMETER VaultVersion
Specifies the installed vault client version.

.EXAMPLE
PS> .\Copy-NumSchmEx.ps1 -VaultUser administrator -VaultServer localhost -Vault Vault
This example login to vault with no password, assuming SQL Server is the same as vault server, using default SQL user and password. It is equivalent to the following:

PS> .\Copy-NumSchmEx.ps1 -VaultUser administrator -VaultServer localhost -Vault Vault -SqlServer localhost -SqlUser sa -SqlPassword AutodeskVault@26200 -VaultVersion 2019

.EXAMPLE
PS> .\Copy-NumSchmEx.ps1 -VaultUser administrator -VaultPassword P@ssw0rd -VaultServer VaultSvr -Vault Vault -SqlServer SQLSvr -SqlUser sa -SqlPassword AutodeskVault@26200 -VaultVersion 2018
This example login to vault with password, assuming installed vault version is 2018. SQL Server is a different server, using default sql user and password.

#>

Param(
    [Parameter(Mandatory = $True)]  
    [string]$VaultUser,
    [string]$VaultPassword,
    [Parameter(Mandatory = $True)]  
    [string]$VaultServer,
    [Parameter(Mandatory = $True)]  
    [string]$Vault,
    [string]$SqlServer = $VaultServer,
    [string]$SqlUser = "sa",
    [string]$SqlPassword = "AutodeskVault@26200",
    [string]$VaultVersion = "2019"
)

Write-Host

$buildNumbers = @{
    "2016" = "21.0"
    "2017" = "22.0"
    "2018" = "23.0"
    "2019" = "24.0"
}

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

Add-Type -AssemblyName "System.Data"
 
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
        [long]$schmId = Read-Host -Prompt "Enter the id of the numbering scheme you want to copy or modify"
    }
    catch { }
} while ($schmId -notin $numSchms.SchmID)

$numSchm = $numSchms | Where-Object SchmID -eq $schmId

Write-Host "`nThe numbering scheme you select is: " -NoNewline
Write-Host -ForegroundColor Green "$($numSchm.Name)`n"

Write-Host "What do you want to do?`n"
Write-Host " 1. Copy the numbering scheme"
Write-Host " 2. Copy the numbering scheme, and preserve used numbers from it (require SQL Server connection)"
Write-Host " 3. Delete and recreate the numbering scheme (i.e. reset the scheme)"
Write-Host " 4. Delete and recreate the numbering scheme, and preserve all used numbers (require SQL Server connection)`n"

do
{
    try
    {
        [int]$choice = Read-Host -Prompt "Your choice"
        if ($choice -in 2, 4)
        {
            if (!$numSchm.IsInUse)
            {
                Write-Host -ForegroundColor Yellow "`nThe numbering scheme you selected is not used yet.`n"
                $choice = -1
            }
            elseif ($numSchm.FieldArray | Where-Object FieldTyp -eq "WorkgroupLabel")
            {
                Write-Host -ForegroundColor Yellow "`nOperation is not supported for numbering scheme with a WorkgroupLabel field.`n"
                $choice = -1
            }
        }
    }
    catch { }
} while ($choice -notin 1, 2, 3, 4)

Write-Host

if ($choice -in 1, 2)
{
    do
    {
        $newSchmName = Read-Host -Prompt "Enter the name for the new numbering scheme"
        if ($newSchmName -in $numSchms.Name)
        {
            Write-Host -ForegroundColor Yellow "`nThe name is already used.`n"
            $newSchmName = ""
        }
    } while (!$newSchmName)

    Write-Host
}
elseif ($choice -in 3, 4)
{
    $newSchmName = $numSchm.Name
}

if ($choice -in 2, 4)
{
    Write-Host "We are about to query the SQL Server behind the vault to get information of used numbers, please make sure you have the correct login information for SQL Server.`n"
    Write-Host -NoNewline "Press any key to continue ... "
    Read-Host

    $connString = "Server=$SqlServer\AutodeskVault;Database=$Vault;User ID=$SqlUser;Password=$SqlPassword"

    try
    {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
        try
        {
            $conn.Open()
        }
        catch
        {
            Write-Host -ForegroundColor Red "Failed to connect to SQL Server, cannot proceed.`n"
            return
        }

        $sqlCmd = $conn.CreateCommand()
        $sqlCmd.CommandText = "SELECT * FROM dbo.SchemePattern WHERE SchemeId = $schmId"
        try
        {
            $reader = $sqlCmd.ExecuteReader()
            $patterns = New-Object 'System.Collections.Generic.List[psobject]'
            while ($reader.Read())
            {
                $pattern = New-Object psobject -Property @{
                    Prefix       = $reader["Prefix"]
                    Suffix       = $reader["Suffix"]
                    SchemeId     = $reader["SchemeId"]
                    NextNum      = $reader["NextNum"]
                    AutoFieldLen = $reader["AutoFieldLen"]
                }
                $patterns.Add($pattern)
            }

            if (!$patterns)
            {
                Write-Host -ForegroundColor Red "The numbering scheme you selected is not used yet.`n"
                return
            }

            Write-Host -ForegroundColor Green "Successfully retrieved used numbers from SQL Server.`n"
        }
        finally
        {
            $reader.Close()
        }
    }
    catch
    {
        Write-Host -ForegroundColor Red "Error occurred when querying SQL Server.`n"
        return
    }
    finally
    {
        $conn.Close()
    }
}

$fields = $numSchm.FieldArray

if ($choice -in 3, 4)
{
    Write-Host "We are going to delete the selected numbering scheme from Vault, in order to recreate it.`n"
    Write-Host -ForegroundColor Yellow "Make sure nobody is using it at the moment. Do not do this during work hours.`n"
    Write-Host -NoNewline "Press any key to continue ... "
    Read-Host

    try
    {
        if ($numSchm.IsDflt)
        {
            $documentService.SetDefaultNumberingScheme(-1)
        }
        $documentService.DeleteNumberingSchemeUnconditional($numSchm.SchmId)
        Write-Host -ForegroundColor Green "The selected numbering scheme is deleted.`n"
    }
    catch
    {
        Write-Host -ForegroundColor Red "Failed to delete numbering scheme.`n"
        return
    }
}

try
{
    Write-Host "Creating new numbering scheme ...`n"

    $newSchm = $documentService.AddNumberingScheme($newSchmName, $fields, $numSchm.ToUpper)
    $newSchm = $documentService.ActivateNumberingSchemes($newSchm.SchmID)[0]

    Write-Host -ForegroundColor Green "The new numbering scheme is successfully created.`n"

    $newSchm | Format-Table $Props
}
catch
{
    Write-Host -ForegroundColor Red "Failed to create new numbering scheme.`n"
    return
}

if ($choice -in 1, 3)
{
    return
}

Write-Host "Now you can modify the new numbering scheme in Vault, for example, to add some new items to the predefined list fields in the scheme. After you have finished, return to this window and continue.`n"
Write-Host -ForegroundColor Red "DO NOT CLOSE THIS WINDOW!`n"
Write-Host -NoNewline "Press any key to continue ... "
Read-Host

Write-Host -ForegroundColor Yellow "Make sure you have finished modifying your numbering scheme!`n"
Write-Host -NoNewline "We are going to generate used numbers for the new numbering scheme, press any key to continue ... "
Read-Host


# get fields from the new scheme in case it's changed
$numSchms = $documentService.GetNumberingSchemesByType("All")
$newSchm = $numSchms | Where-Object Name -eq $newSchmName
$fields = $newSchm.FieldArray

$AutogenField = $fields | Where-Object FieldTyp -eq "Autogenerated"

$regex = '^'
foreach ($field in $fields)
{
    switch ($field.FieldTyp)
    {
        "FixedText" { $regex += $field.FixedTxtVal }
        "FreeText"
        {
            $regex += "("
            $regex += "[a-zA-Z0-9]{$($field.MinLen),$($field.MaxLen)}"
            $regex += ")"
        }
        "PredefinedList"
        {
            $regex += "("
            $regex += $field.CodeArray.Code -join "|"
            $regex += ")"
        }
        "Delimiter" { $regex += [regex]::Escape($field.DelimVal) }
        "WorkgroupLabel" { $regex += $field.Val }
        # "Autogenerated" { 
        #     if ($field.Zeropadding) { $regex += "\d{$($field.Len)}" }
        #     else { $regex += "\d{$($field.From.ToString().Length),$($field.Len)}" }
        # }
    }
}
$regex += '$'

if ($patterns)
{
    foreach ($pattern in $patterns)
    {
        $numberParts = $pattern.Prefix + $pattern.Suffix
        $numberPreview = $pattern.Prefix + [string]::new('#', $AutogenField.Len) + $pattern.Suffix
        if ($numberParts -imatch $regex)
        {
            $fieldValues = $Matches[1..($Matches.Count - 1)]
            $count = $pattern.NextNum - $AutogenField.From
            $strArray = New-Object Autodesk.Connectivity.WebServices.StringArray -Property @{ Items = $fieldValues }

            Write-Host "Generating numbers for pattern $numberPreview ..."
            $numbers = $documentService.GenerateFileNumbers(@($newSchm.SchmID) * $count, @($strArray) * $count)
            Write-Host "Total $count numbers generated, the last generated number is $($numbers[-1]).`n"
        }
        else
        {
            Write-Host -ForegroundColor Yellow "Warning: Pattern $numberPreview is not generated.`n"
        }
    }

    Write-Host -ForegroundColor Green "Successfully generated used numbers.`n"
}
