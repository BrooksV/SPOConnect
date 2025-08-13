# SPOConnect
PowerShell Class to help with login and connecting to Share Point Online (SPO)

This is a work in progress. Feel free to provide feedback on how to improve.

```powershell
<# 
Constructors:
SPOConnect() - empty object uses the Default "Documents" SPO Site Folder
SPOConnect ([string] $folderName) - Select a different SPO Site Folder

Methods:
Init([string] $DocFolder)
Init([string] $DocFolder, [string] $SiteURL)
SetSiteUrl([string] $Site)
SetSiteUrl([string] $tenantUrl, [string] $Site)
Clear() - Resets the Class Properties to the defaults
ConnectAsUser([string] $Method)
ConnectAsManagedIdentity([string] $IdentityType, [string] $Identity)
SetConnection([HashTable] $pnpSplat)
SetIdentity ([string] $UserAssignedManagedIdentity)
GetLibraryLists()
GetLibraryLists ([string] $ListTitleOrURL)
PrintLibraryLists()
SelectList()
SelectList ([string] $ListTitleOrURL)
PrintSelectedList()
GetFolders()
GetFolders ([string] $FolderNameOrUrl)
PrintFolders()

Method Parameters:
$SiteURL: URL to your SPO Tenant Site
$Method: The ConnectAsUser() connection Methods (Interactive, DeviceLogin, UseWebLogin)
$DocFolder: SPO Site Folder Path 
$Identity: The value for the IdentityType
$IdentityType: The Value for Identity can be $null, 'ObjectId', 'ClientId', 'ResourceId'
$pnpSplat: HashTable for CmdLet Connect-PnPOnline @pnpSplat
$UserAssignedManagedIdentity: IdentityType GUID Value which could be a Service Principal Object Id, ApplicationId, or ClientId
$ListTitleOrURL: List Title or the URL to the Desired List
$FolderNameOrUrl: Name of Folder ot the URL to the Desired Folder

Properties:
    [string] $TenantUrl = 'https://mysite.sharepoint.com'
    [string] $Site = '/sites/Main'
    [string] $SiteURL = "$($this.TenantUrl)/$($this.Site)"
    [string] $DocFolder = 'Documents'
    [object[]] $LibraryLists = $null
    [object[]] $DocumentLibrary = $null
    [string] $TimeStamp = (Get-Date -Format "yyyy-MMdd-HHmm")
    [Int16] $width = 999
    [string] $outputPath
    [string] $inputPath
    [bool] $outputDetails = $true # $false # 

    # PnP.PowerShell.Commands.Base.PnPConnection 
    [Object] $pnpConnection = $null

    # PnP.Framework.PnPClientContext
    [Object] $Context = $null

    [HashTable] $pnpSplat = @{}
    [string] $ConnectionMethod = [string]::Empty
    [string] $UserAssignedManagedIdentity = [guid]::Empty
    [bool] $ForceOverWrite = $false

Method Examples (not in ny specific order)
$spo = [SPO]::New('Documents/Projects/Main')
$spo.ConnectAsManagedIdentity($IdentityType, $Identity)
$spo.SetConnection($pnpSplat)
$spo.ConnectAsManagedIdentity('ObjectId', $UserAssignedManagedIdentityObjectId)
$spo.ConnectAsManagedIdentity('ClientId', $UserAssignedManagedIdentityClientId)
$spo.Clear()
$spo.ConnectAsUser('Interactive')
$spo.ConnectAsUser('DeviceLogin')
$spo.ConnectAsUser('WebLogin')
$spo.GetLibraryLists('Documents')
$spo.GetLibraryLists('DocumentLibrary')
$spo.PrintLibraryLists()
$spo.SelectList()
$spo.SelectList('Documents')
$spo.PrintSelectedList()
$spo.GetFolders()
$spo.GetFolders('Team Management')
$spo.PrintFolders()

# Non-interactive example for Pipelines
# How to determine if running from Pipeline? See AzGovViz by Julian Hayward Pipeline example
$myApplicationId = '???'
$identity = Get-AzADServicePrincipal -ApplicationId $myApplicationId
$UserAssignedManagedIdentityObjectId = $identity.Id # (Should be same as $myApplicationId)
$UserAssignedManagedIdentityClientId = $identity.AppId
$spo.ConnectAsManagedIdentity('ObjectId', $UserAssignedManagedIdentityObjectId)
$spo.ConnectAsManagedIdentity('ClientId', $UserAssignedManagedIdentityClientId)

#>
<# Usage Examples #>
$spo = [SPO]::New('Documents/Projects/Main')
$spo | FL

#------------------------------------------------------------------------------
# Different ways to Login / Connect-PnPOnline 
#------------------------------------------------------------------------------
# $spo.Clear()
$spo.ConnectAsUser('Interactive')
# $spo.ConnectAsUser('DeviceLogin')
# $spo.ConnectAsUser('WebLogin')

# $spo.GetLibraryLists('Documents')
$spo.GetLibraryLists('DocumentLibrary')
$spo.PrintLibraryLists()
$spo.LibraryLists | Select-Object BaseType, BaseTemplate, Hidden, ItemCount, Title, Id, @{L='Url';E={$_.DefaultViewUrl}} | Format-Table -AutoSize -Force -Wrap

<# SelectList Searches $spo.LibraryLists to set $spo.DocumentLibrary to the searched subset #>
# List All Doc Libraries
$spo.SelectList()
$spo.PrintSelectedList()

# List All Folders in Documents Library
$spo.GetFolders()
$spo.PrintFolders()
$spo.Folders | Format-Table -AutoSize -Force -Wrap | Select-Object BaseType, BaseTemplate, Hidden, ItemCount, Title, Id, @{L='Url';E={$_.DefaultViewUrl}} | Format-Table -AutoSize -Force -Wrap

$spo.SelectList('Documents')
$spo.PrintSelectedList()
$spo.DocumentLibrary | Select-Object BaseType, BaseTemplate, Hidden, ItemCount, Title, Id, @{L='Url';E={$_.DefaultViewUrl}} | Format-Table -AutoSize -Force -Wrap
$spo.SelectedList

# List All Folders in Documents Library
$spo.GetFolders('Team Management')
$spo.PrintFolders()
$spo.Folders | Format-Table -AutoSize -Force -Wrap | Select-Object BaseType, BaseTemplate, Hidden, ItemCount, Title, Id, @{L='Url';E={$_.DefaultViewUrl}} | Format-Table -AutoSize -Force -Wrap
$spo.Folders | Sort-Object -Property Name | Select-Object -Property Name, @{L='Type';E={$_.TypedObject.GetType().Name}}, @{L='Items/Size';E={$_.ItemCount}}, @{L='Last Modified';E={$_.TimeLastModified}},  @{L='RelativeUrl';E={$_.ServerRelativeUrl}}  | Format-Table -AutoSize -Force -Wrap | Out-String -Width $spo.width

$spo.SelectedFolder






```
