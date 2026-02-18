$GroupType = "Distribution Group" # "Mail-enabled Security Group" or "Distribution Group"

# PowerShell commands to import
$commands = @("New-DistributionGroup")

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param()
    try {
        $rawCertificate = [system.convert]::FromBase64String($EntraIdCertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $EntraIdCertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

#region Import module & connect
try {    
    $actionMessage = "importing module [ExchangeOnlineManagement]"
    $importModuleSplatParams = @{
        Name        = "ExchangeOnlineManagement"
        Cmdlet      = $commands
        Verbose     = $false
        ErrorAction = "Stop"
    }
    $null = Import-Module @importModuleSplatParams

    #region Retrieving certificate
    $actionMessage = "retrieving certificate"
    $certificate = Get-MSEntraCertificate
    #endregion Retrieving certificate
    
    #region Connect to Microsoft Exchange Online
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
    Write-Information "Connected to Microsoft Exchange Online"
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
    Write-Warning $warningMessage
    Write-Error $auditMessage
}


# Create Mail-enabled Security Group
try{   
    $OwnersToAdd  = ($form.multiselectOwners.UserPrincipalName)
    $MembersToAdd = ($form.multiselectMembers.UserPrincipalName)

    $groupParams = @{
        Name                =   $form.naming.name
        DisplayName         =   $form.naming.displayName
        PrimarySmtpAddress  =   $form.naming.primarySmtpAddress
        Alias               =   $form.naming.alias
        ManagedBy           =   $OwnersToAdd
        Members             =   $MembersToAdd
        CopyOwnerToMember   =   $true
    }
    
    Switch($GroupType){
        'Mail-enabled Security Group' {
            $mailEnabledSecurityGroup = New-DistributionGroup -Type security @groupParams -ErrorAction Stop
        }

        'Distribution Group' {
            $mailEnabledSecurityGroup = New-DistributionGroup @groupParams -ErrorAction Stop
        }
    }
    
    $Log = @{
        Action            = "CreateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange Online" # optional (free format text) 
        Message           = "Created distribution group:  $($mailEnabledSecurityGroup.displayName)" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $($mailEnabledSecurityGroup.displayName) # optional (free format text) 
        TargetIdentifier  = $([string]$mailEnabledSecurityGroup.Guid)  # optional (free format text) 
    }
    #send result back  

    Write-Information -Tags "Audit" -MessageData $log
  
} catch {
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
        Action            = "CreateResource" # optional. ENUM (undefined = default) 
        System            = "Exchange Online" # optional (free format text) 
        Message           = "Error creating $GroupType [$($groupParams.Name)]. Error: $($_.Exception.Message)" # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $($groupParams.displayName) # optional (free format text) 
        
    }
    Write-Information -Tags "Audit" -MessageData $log
    Write-Warning $warningMessage
    Write-Error $auditMessage
    # exit # use when using multiple try/catch and the script must stop
}
finally {
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps
    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }
    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    Write-Information "Disconnected from Microsoft Exchange Online"
}
