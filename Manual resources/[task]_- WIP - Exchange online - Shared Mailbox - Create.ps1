# variables configured in form
$mailboxDisplayName = $form.displayName
$mailboxMailPrefix = $form.mailPrefix
$mailboxMailDomain = $form.mailDomain.mailDomain
$mailboxPrimarySmtpAddress = "$($mailboxMailPrefix)@$($mailboxMailDomain)"
$mailboxAlias = $form.alias

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdOrganization = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
$commands = @(
    "New-Mailbox",
    "Set-Mailbox"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

#region functions
function Get-MSEntraCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions

try {
    # Import module
    $actionMessage = "importing module [ExchangeOnlineManagement]"
        
    $importModuleSplatParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = $commands
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $null = Import-Module @importModuleSplatParams

    Write-Verbose "Imported module [ExchangeOnlineManagement]"

    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"

    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    Write-Verbose "Converted base64 certificate string to certificate object"

    # Connect to Microsoft Exchange Online
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps
    $actionMessage = "connecting to Microsoft Exchange Online"

    $createExchangeSessionSplatParams = @{
        Organization          = $EntraIdOrganization
        AppID                 = $EntraIdAppId
        Certificate           = $certificate
        CommandName           = $commands
        ShowBanner            = $false
        ShowProgress          = $false
        TrackPerformance      = $false
        SkipLoadingCmdletHelp = $true
        SkipLoadingFormatData = $true
        ErrorAction           = "Stop"
    }

    $null = Connect-ExchangeOnline @createExchangeSessionSplatParams

    # Create shared mailbox
    $action = "CreateResource"
    $actionMessage = "creating shared mailbox with displayname [$($mailboxDisplayName)] and PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]"

    $CreateMailboxParams = @{
        Shared             = $true
        Name               = $mailboxDisplayName
        DisplayName        = $mailboxDisplayName
        PrimarySmtpAddress = $mailboxPrimarySmtpAddress
        ErrorAction        = 'Stop'
    }

    # Add Alias if specified
    if (-not [string]::IsNullOrEmpty($mailboxAlias)) {
        $CreateMailboxParams["Alias"] = $mailboxAlias
    }

    $null = New-Mailbox @CreateMailboxParams

    # Send auditlog to HelloID
    $Log = @{
        Action            = $action # optional. ENUM (undefined = default) 
        System            = "ExchangeOnline" # optional (free format text) 
        Message           = "Created shared mailbox with displayname [$($mailboxDisplayName)] and PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]"  # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $mailboxDisplayName # optional (free format text) 
        TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) 
    }
    Write-Information -Tags "Audit" -MessageData $log

    # Update shared mailbox to enable MessageCopyForSendOnBehalfEnabled and MessageCopyForSentAsEnabled (can only be done after mailbox is created)
    $action = "UpdateResource"
    $actionMessage = "updating MessageCopyForSendOnBehalfEnabled and MessageCopyForSentAsEnabled to [true] for shared mailbox with PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]"

    # Wait for the mailbox to be created before updating
    Start-Sleep -Seconds 10

    $UpdateMailboxParams = @{
        Identity                          = $mailboxPrimarySmtpAddress
        MessageCopyForSendOnBehalfEnabled = $true
        MessageCopyForSentAsEnabled       = $true
        ErrorAction                       = 'Stop'
    }

    $null = Set-Mailbox @UpdateMailboxParams

    # Send auditlog to HelloID
    $Log = @{
        Action            = $action # optional. ENUM (undefined = default) 
        System            = "ExchangeOnline" # optional (free format text) 
        Message           = "Updated MessageCopyForSendOnBehalfEnabled and MessageCopyForSentAsEnabled to [true] for shared mailbox with PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]"  # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $mailboxDisplayName # optional (free format text) 
        TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) 
    }
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    $ex = $PSItem
    if (-not [string]::IsNullOrEmpty($ex.Exception.Data.RemoteException.Message)) {
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Data.RemoteException.Message)"
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Data.RemoteException.Message)"
    }
    else {
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
    }

    $Log = @{
        Action            = $action # optional. ENUM (undefined = default) 
        System            = "ExchangeOnline" # optional (free format text) 
        Message           = $auditMessage # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $mailboxDisplayName # optional (free format text) 
        TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) 
    }
    
    Write-Information -Tags "Audit" -MessageData $log
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
finally {
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps
    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }
    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
}
