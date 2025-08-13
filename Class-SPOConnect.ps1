<#
    Class-SPOConnect.ps1 - Class for Connecting to an SPO Site

Usage Examples:

. "$Path/Class-SPOConnect.ps1"
$spo = [SPOConnect]::New('Documents/Projects/Project AzDefender')

#>

<# Connect-PnPOnline Documentation
    https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html#syntax

    DESCRIPTION: Connects to a SharePoint site or another API and creates a context that is required for the other PnP Cmdlets.
        See https://pnp.github.io/powershell/articles/connecting.html for more information on the options to connect.

    Credentials (Default)
    Connect-PnPOnline [-ReturnConnection] [-Url] <String> [-Credentials <CredentialPipeBind>] [-CurrentCredentials] [-CreateDrive]
        [-DriveName <String>] [-ClientId <String>] [-RedirectUri <String>] [-AzureEnvironment <AzureEnvironment>] [-TenantAdminUrl <String>]
        [-TransformationOnPrem] [-ValidateConnection] [-MicrosoftGraphEndPoint <string>] [-AzureADLoginEndPoint <string>] [-Connection <PnPConnection>]

    SharePoint ACS (Legacy) App Only
    Connect-PnPOnline [-ReturnConnection] [-Url] <String> [-Realm <String>] -ClientSecret <String> [-CreateDrive] [-DriveName <String>]
        -ClientId <String> [-AzureEnvironment <AzureEnvironment>] [-TenantAdminUrl <String>]
        [-ValidateConnection] [-MicrosoftGraphEndPoint <string>] [-AzureADLoginEndPoint <string>] [-Connection <PnPConnection>]

    App-Only with Azure Active Directory
    Connect-PnPOnline [-ReturnConnection] [-Url] <String> [-CreateDrive] [-DriveName <String>] -ClientId <String> -Tenant <String> 
        [-CertificatePath <String>] [-CertificateBase64Encoded <String>] [-CertificatePassword <SecureString>] 
        [-AzureEnvironment <AzureEnvironment>] [-TenantAdminUrl <String>] [-ValidateConnection] [-MicrosoftGraphEndPoint <string>]
        [-AzureADLoginEndPoint <string>] [-Connection <PnPConnection>]

    App-Only with Azure Active Directory using a certificate from the Windows Certificate Management Store by thumbprint
    Connect-PnPOnline [-ReturnConnection] [-Url] <String> [-CreateDrive] [-DriveName <String>] -ClientId <String> -Tenant <String> 
        -Thumbprint <String> [-AzureEnvironment <AzureEnvironment>] [-TenantAdminUrl <String>] [-ValidateConnection]
        [-MicrosoftGraphEndPoint <string>] [-AzureADLoginEndPoint <string>] [-Connection <PnPConnection>]

        On-premises login for page transformation from on-premises SharePoint to SharePoint Online
    Connect-PnPOnline -Url <String> -TransformationOnPrem [-CurrentCredential]

    Access Token
    Connect-PnPOnline -Url <String> -AccessToken <String> [-AzureEnvironment <AzureEnvironment>] [-MicrosoftGraphEndPoint <string>] [-AzureADLoginEndPoint <string>] [-ReturnConnection]

    Environment Variable
    Connect-PnPOnline [-ReturnConnection] [-Url] <String> [-EnvironmentVariable] [-CurrentCredentials] [-CreateDrive] [-DriveName <String>] [-RedirectUri <String>]
        [-AzureEnvironment <AzureEnvironment>] [-TenantAdminUrl <String>] [-TransformationOnPrem] [-ValidateConnection] 
        [-MicrosoftGraphEndPoint <string>] [-AzureADLoginEndPoint <string>] [-Connection <PnPConnection>]

    Azure AD Workload Identity
    Connect-PnPOnline [-ReturnConnection] [-ValidateConnection] [-Url] <String> [-AzureADWorkloadIdentity] [-Connection <PnPConnection>]

    Azure AD Workload Identity
    Connect-PnPOnline [-ReturnConnection] [-ValidateConnection] [-Url] <String> [-AzureADWorkloadIdentity] [-Connection <PnPConnection>]

#>

class SPOConnect
{
    # SPOConnect Default Settings
    [string] $TenantUrl = 'https://mysite.sharepoint.com'
    [string] $Site = '/sites/Main'
    [string] $SiteURL = "$($this.TenantUrl)/$($this.Site)"
    [string] $DocFolder = 'Documents'
    [object[]] $LibraryLists = $null
    [object[]] $DocumentLibrary = $null

    # Properties
    [string] $TimeStamp = (Get-Date -Format "yyyy-MMdd-HHmm")
    [Int16] $width = 999
    [string] $outputPath
    [string] $inputPath
    [bool] $outputDetails = $true # $false # 

    # PnP.PowerShell.Commands.Base.PnPConnection 
    [Object] $pnpConnection = $null

    # PnP.Framework.PnPClientContext
    # 
    [Object] $Context = $null

    [HashTable] $pnpSplat = @{}
    [string] $ConnectionMethod = [string]::Empty
    [string] $UserAssignedManagedIdentity = [guid]::Empty

    # Constructor -empty object uses the Default "Documents" SPO Site Folder
    SPOConnect() {
        $this.Init('Documents') 
    }
    
    # Constructor - 
    SPOConnect ([string] $folderName) { 
        $this.Init($folderName) 
        # $this.LoadModule()
    }

    # Method
    [void] LoadModule() {
        Try {
            #------------------------------------------------------
            # Import SharePoint PNP Module
            #------------------------------------------------------
            $Env:PNPPOWERSHELL_UPDATECHECK = 'Off'
            # [Environment]::GetEnvironmentVariables().GetEnumerator() | Sort-Object -Property Name
            $availableModule = Find-Module -Name PnP.PowerShell
            $loadedModule = Get-Module -Name PnP.PowerShell
            If (-not $loadedModule) {
                $loadedModule = Import-Module -Name PnP.PowerShell -SkipEditionCheck -PassThru -Force
            }
            If ($availableModule.Version -ne $loadedModule.Version.ToString()) {
                $loadedModule | Remove-Module -Force
                $availableModule | Install-Module -Force
                $loadedModule = Import-Module -Name PnP.PowerShell -SkipEditionCheck -Force # -UseWindowsPowerShell
            }
            # Get-Installedmodule -Name PnP.PowerShell
        } Catch {
            Write-Host "$_ | Out-String"
        }
    }

    # Method
    [void] Clear() {
        # Default SPOConnect Settings
        $this.TenantUrl = 'https://pwc.sharepoint.com'
        $this.Site = '/sites/GBL-IFS-gbl_cloud-security-operations'
        $this.SiteURL = "$($this.TenantUrl)/$($this.Site)"
        $this.DocFolder = 'Documents'
        $this.DocumentLibrary = $null
        $this.LibraryLists = $null
        $this.ForceOverWrite = $false
        
        # Properties
        $this.width = 999
        $this.outputPath = [String]::Empty
        $this.inputPath = [String]::Empty
        $this.outputDetails = $true
    
        $this.pnpConnection = $null
        $this.pnpSplat = @{}
        $this.ConnectionMethod = [String]::Empty
        $this.UserAssignedManagedIdentity = $null
    }

    # Hidden, chained helper methods that the constructors must call.
    hidden Init([string] $DocFolder) {
        $this.Init($DocFolder, $this.SiteURL)
    }
    hidden Init([string] $DocFolder, [string] $SiteURL) {
        If ($SiteURL -ne $this.SiteURL) {
            $this.SiteURL = $SiteURL
        }
        $this.DocFolder = $DocFolder
    }

    # Method
    [void] SetSiteUrl([string] $Site) {
        If ($Site -ne $this.Site) {
            $this.Site = $Site
        }
        SetSiteUrl($this.TenantUrl, [string] $Site)
    }

    [void] SetSiteUrl([string] $tenantUrl, [string] $Site) {
        If ($tenantUrl -ne $this.TenantUrl) {
            $this.TenantUrl = $tenantUrl
        }
        $this.Site = $Site
        $this.SiteURL = "$tenantUrl/$Site"
    }

    # Method
    <#
    PnP Management Shell / DeviceLogin
    Connect-PnPOnline [-ReturnConnection] [-Url] <String> [-CreateDrive] [-DriveName <String>] [-DeviceLogin] [-LaunchBrowser] [-ClientId <String>]
        [-AzureEnvironment <AzureEnvironment>] [-ValidateConnection] [-MicrosoftGraphEndPoint <string>] [-AzureADLoginEndPoint <string>] [-Connection <PnPConnection>]

    Web Login for Multi Factor Authentication
    Connect-PnPOnline [-ReturnConnection] [-Url] <String> [-CreateDrive] [-DriveName <String>] [-TenantAdminUrl <String>] [-UseWebLogin] [-ForceAuthentication] [-ValidateConnection]

    Interactive for Multi Factor Authentication
    Connect-PnPOnline -Interactive [-ReturnConnection] -Url <String> [-CreateDrive] [-DriveName <String>] [-LaunchBrowser] [-ClientId <String>] [-AzureEnvironment <AzureEnvironment>] [-TenantAdminUrl <String>] [-ForceAuthentication] [-ValidateConnection] [-MicrosoftGraphEndPoint <string>] [-AzureADLoginEndPoint <string>] [-Connection <PnPConnection>]

    $Method are Interactive; UseWebLogin; DeviceLogin
    #>
    [void] ConnectAsUser([string] $Method) {
        $this.ConnectionMethod = $Method
        $this.UserAssignedManagedIdentity = 'User'
        $this.pnpSplat = @{}
        $this.pnpSplat.Add('Url', $this.SiteURL)
        $this.pnpSplat.Add('ReturnConnection', $true)
        # $this.pnpSplat.Add('Verbose', $this.outputDetails)
        Switch ($Method) {
            'Interactive' {
                $this.pnpSplat.Add('Interactive', $true)
                Break}
            'DeviceLogin' {
                $this.pnpSplat.Add('DeviceLogin', $true)
                $this.pnpSplat.Add('LaunchBrowser', $true)
                Break}
            'UseWebLogin' {
                $this.pnpSplat.Add('UseWebLogin', $true)
                $this.pnpSplat.Add('ForceAuthentication', $true)
                Break}
        }
        # Set-Variable -Name pnpSplat -Value $this.pnpSplat -Force
        [void] $this.SetConnection($this.pnpSplat)
    }

    <#
    System Assigned Managed Identity
    Connect-PnPOnline [-Url <String>] -ManagedIdentity [-ReturnConnection]

    User Assigned Managed Identity by Client Id
    Connect-PnPOnline [-Url <String>] -ManagedIdentity -UserAssignedManagedIdentityClientId <String> [-ReturnConnection]

    User Assigned Managed Identity by Principal Id
    Connect-PnPOnline [-Url <String>] -ManagedIdentity -UserAssignedManagedIdentityObjectId <String> [-ReturnConnection]

    User Assigned Managed Identity by Azure Resource Id
    Connect-PnPOnline [-Url <String>] -ManagedIdentity -UserAssignedManagedIdentityAzureResourceId <String> [-ReturnConnection]
    $IdentityType is $null, 'ObjectId', 'ClientId', 'ResourceId'
    $Identity is th value for $IdentityType
    #>
    [void] ConnectAsManagedIdentity([string] $IdentityType, [string] $Identity) {
        $this.ConnectionMethod = $IdentityType
        $this.UserAssignedManagedIdentity = $Identity
        $this.pnpSplat = @{}
        $this.pnpSplat.Add('Url', $this.SiteURL)
        $this.pnpSplat.Add('ManagedIdentity', $true)
        $this.pnpSplat.Add('ReturnConnection', $true)
        Switch ($IdentityType) {
            {$null -eq $PSItem -or [string]::IsNullOrEmpty($PSItem)} {
                Break}
            'ObjectId' {
                $this.pnpSplat.Add('UserAssignedManagedIdentityObjectId', $Identity)
                Break}
            'ClientId' {
                $this.pnpSplat.Add('UserAssignedManagedIdentityClientId', $Identity)
                Break}
            'ResourceId' {
                $this.pnpSplat.Add('UserAssignedManagedIdentityAzureResourceId', $Identity)
                Break}
        }
        [void] $this.SetConnection($this.pnpSplat)
    }

    # Method
    [void] SetConnection([HashTable] $pnpSplat) {
        $this.pnpSplat = $pnpSplat
        If (-not $this.pnpConnection) {
            Write-Host "Establishing pnpConnection to: $($this.SiteURL)"
        } Else {
            Write-Host "Resetting pnpConnection to: $($this.SiteURL)"
        }
        Try {
            If ($this.outputDetails) {
                Write-Host "SPOConnect pnpConnection Parameters:`n$($pnpSplat | Out-String -Width $this.width)" -ForegroundColor Magenta
                $this.pnpConnection = Connect-PnPOnline @pnpSplat
                Write-Host ($this.pnpConnection | Out-String -Width $this.Width)
            } Else {
                $this.pnpConnection = Connect-PnPOnline @pnpSplat
            }
            #------------------------------------------------------
            # Verify if the connection works
            #------------------------------------------------------
            $this.SelectList('Documents')
            If ([string]::IsNullOrEmpty($this.DocumentLibrary)) {
                Write-Host "Error: Connection is Bad or Documents Folder ($($this.DocFolder)) Not Found" -ForegroundColor Red
            }
        } Catch {
            Write-Host ($_ | Out-String -Width $this.width) -ForegroundColor Red
        }
    }

    # Method
    [void] SetIdentity ([string] $UserAssignedManagedIdentity) {
        $this.UserAssignedManagedIdentity = $UserAssignedManagedIdentity
    }

    # Method
    [object[]] GetLibraryLists() {
        Return ($this.GetLibraryLists($null))
    }

    [object[]] GetLibraryLists ([string] $ListTitleOrURL) {
        $this.LibraryLists = Get-PnPList -Connection $this.pnpConnection
        Switch ($ListTitleOrURL) {
            {$null -eq $PSItem -or [string]::IsNullOrEmpty($PSItem)} {
                Break}
            {$PSItem.Split('/').Count -eq 1} {
                $this.LibraryLists = $this.LibraryLists | Where-Object {$_.BaseType -eq $ListTitleOrURL -or $_.Title -eq $ListTitleOrURL}
                Break}
            {$PSItem.Split('/').Count -ge 2} {
                # $this.DocumentLibrary = Get-PnPList -Connection $this.pnpConnection | Where-Object {$_.BaseType -eq $ListTitleOrURL -or $_.ServerRelativeUrl -like "*$($ListTitleOrURL)*"}
                $this.LibraryLists = $this.LibraryLists | Where-Object {$_.BaseType -eq $ListTitleOrURL -or $_.DefaultViewUrl -like "*$($ListTitleOrURL)*"}
                Break}
            Default {
                $this.LibraryLists = $this.LibraryLists | Where-Object {$_.BaseType -eq $ListTitleOrURL -or $_.Title -eq 'Documents'}
                Break}
        }
        Return ($this.GetLibraryLists)
    }

    [void] PrintLibraryLists() {
        Write-Host ($this.LibraryLists | 
            Sort-Object -Property BaseType, BaseTemplate, Title | 
            Select-Object BaseType, BaseTemplate, Hidden, ItemCount, Title, Id, @{L='Url';E={$_.DefaultViewUrl}} | 
            Format-Table -AutoSize -Force -Wrap |
            Out-String -Width $this.width)
    }

    # Method
    [void] SelectList() {
        $this.SelectList($null)
    }

    # Method
    [object[]] GetFolders() {
        Return ($this.GetFolders($null))
    }

    [object[]] GetFolders ([string] $FolderNameOrUrl) {
        $this.Folders = @()
        ForEach ($library in $this.DocumentLibrary) {
            Switch ($FolderNameOrUrl) {
                {$null -eq $PSItem -or [string]::IsNullOrEmpty($PSItem)} {
                    $this.Folders += Get-PnPFolder -Connection $this.pnpConnection -List $library
                    # | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false}
                    Break}
                {$PSItem.Split('/').Count -eq 1} {
                    $this.Folders += Get-PnPFolder -Connection $this.pnpConnection -List $library | Where-Object {$_.Name -eq $FolderNameOrUrl}
                    # | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -eq $FolderNameOrUrl}
                    Break}
                {$PSItem.Split('/').Count -ge 2} {
                    $this.Folders += Get-PnPFolder -Connection $this.pnpConnection -List $library | Where-Object {$_.ServerRelativeUrl -like "*$($FolderNameOrUrl)*"}
                    # | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.ServerRelativeUrl -like "*$($FolderNameOrUrl)*"}
                    Break}
                Default {
                    $this.Folders += Get-PnPFolder -Connection $this.pnpConnection -List $library
                    # | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -eq 'Documents'}
                    Break}
            }
        }
        $this.SelectedFolder = $this.Folders.Name -join '; '
        Return ($this.Folders | Sort-Object -Property Name)
    }

    [void] PrintFolders() {
        Write-Host ($this.Folders | 
        Sort-Object -Property Name |
            Select-Object -Property Name, @{L='Type';E={$_.TypedObject.GetType().Name}}, @{L='Items/Size';E={$_.ItemCount}}, @{L='Last Modified';E={$_.TimeLastModified}},  @{L='RelativeUrl';E={$_.ServerRelativeUrl}}  | 
            Format-Table -AutoSize -Force -Wrap |
            Out-String -Width $this.width)
    }
}
