#######################################################################
# Template: RHo HelloID SA Powershell data source
# Name: function-check-online-mailbox-exists
# Date: 05-02-2024
#######################################################################

# For basic information about powershell data sources see:
# https://docs.helloid.com/en/service-automation/dynamic-forms/data-sources/powershell-data-sources.html

# Service automation variables:
# https://docs.helloid.com/en/service-automation/service-automation-variables.html

#region init
# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# global variables (Automation --> Variable libary):
$TenantId = $EntraTenantId
$AppID = $EntraAppID
$Secret = $EntraSecret
$Organization = $EntraOrganization

# variables configured in form:
$Maildomain = $datasource.Organization.Maildomain
$Name = $datasource.name
$Alias = $datasource.alias
$PrimarySmtpAddress = $Alias.Replace(" ", "") + "@$Maildomain"

# PowerShell commands to import
$commands = @("Get-User", "Get-Mailbox")
#endregion init

try {
    #region import module
    $actionMessage = "importing $moduleName module"

    $importModuleParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = $commands
        ErrorAction = 'Stop'
    }

    Import-Module @importModuleParams
    #endregion import module

    #region create access token
    Write-Verbose "Creating Access Token"
    $actionMessage = "creating access token"
        
    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$AppID"
        client_secret = "$Secret"
        resource      = "https://outlook.office365.com"
    }

    $exchangeAccessTokenParams = @{
        Method          = 'POST'
        Uri             = "https://login.microsoftonline.com/$TenantId/oauth2/token"
        Body            = $body
        ContentType     = 'application/x-www-form-urlencoded'
        UseBasicParsing = $true
    }
        
    $accessToken = (Invoke-RestMethod @exchangeAccessTokenParams).access_token
    #endregion create access token

    #region connect to Exchange Online
    Write-Verbose "Connecting to Exchange Online"
    $actionMessage = "connecting to Exchange Online"

    $exchangeSessionParams = @{
        Organization     = $Organization
        AppID            = $AppID
        AccessToken      = $accessToken
        CommandName      = $commands
        ShowBanner       = $false
        ShowProgress     = $false
        TrackPerformance = $false
        ErrorAction      = 'Stop'
    }
    Connect-ExchangeOnline @exchangeSessionParams
        
    Write-Information "Successfully connected to Exchange Online"
    #endregion connect to Exchange Online

    #region check shared mailbox
    $actionMessage = "getting shared mailbox"

    $SharedMailboxParams = @{
        Filter               = "{Alias -eq '$Alias' -or Name -eq '$Name' -or PrimarySmtpAddress -eq '$PrimarySmtpAddress'}"
        # RecipientTypeDetails = 'SharedMailbox'
        ErrorAction          = 'Stop'        
    }
    
    $SharedMailbox = Get-Mailbox @SharedMailboxParams
   
    if ([string]::IsNullOrEmpty($SharedMailbox)) {
        Write-Information  "Shared Mailbox name [$Name] is available"
        $outputMessage = "Valid | Shared Mailbox name [$Name] is available"
        $returnObject = @{
            text = $outputMessage
        }     
    }
    else {
        Write-Information  "Shared Mailbox [$Name] exists. Please try another name" 
        $outputMessage = "Invalid | Shared Mailbox name [$Name] exists. Please try another name"
        $returnObject = @{
            text = $outputMessage
        }
    }
    #endregion check shared mailbox           
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorMessage = ($ex.ErrorDetails.Message | Convertfrom-json).error_description
    }
    else {
        $errorMessage = $($ex.Exception.message)
    }

    Write-Error "Error $actionMessage for Exchange Online shared mailbox [$Name]. Error: $errorMessage"
    
    $outputMessage = "Invalid | Error $actionMessage for Exchange Online shared mailbox [$Name]. Error: $errorMessage"
    $returnObject = @{
        text = $outputMessage
    }
}
finally {
    Write-Output $returnObject 
}
#endregion lookup
