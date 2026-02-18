$Mailsuffix = $ExchangeOnlineDistributionGroupDomain

# PowerShell commands to import
$commands = @("Get-DistributionGroup")

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

try {
    $iterationMax = 10
    $iterationStart = 1;
        
    for($i = $iterationStart; $i -lt $iterationMax; $i++) {
        if($i -eq $iterationStart) {
            $tempName = $datasource.name
            $DisplayName =  $tempName
    
            $Name =   $tempName.Replace(" ","")

            $PrimarySmtpAddress =   $tempName.Replace(" ","") + "@" + $Mailsuffix
            
            $Alias =   $tempName.Replace(" ","")

            $SamAccountName = $Alias
         } else {
            $tempName = $datasource.name
            $DisplayName =  $tempName + "$i"
    
            $Name =   ($tempName + "$i").Replace(" ","")

            $PrimarySmtpAddress =   ($tempName + "$i").Replace(" ","") + "@" + $Mailsuffix 
            
            $Alias =   ($tempName + "$i").Replace(" ","")

            $SamAccountName = $Alias
        }
        
        Write-Information -Message "Searching for Distribution Group Name=$Name or DisplayName=$DisplayName or EmailAddresses=$PrimarySmtpAddress or Alias=$Alias"

        $found = Get-DistributionGroup -Filter "Name -eq '$Name' -or DisplayName -eq '$DisplayName' -or EmailAddresses -eq '$PrimarySmtpAddress' -or Alias -eq '$Alias'"

        if(@($found).count -eq 0) {
            $returnObject = @{
                name=$Name;
                displayName=$DisplayName;
                primarySmtpAddress=$PrimarySmtpAddress;
                alias=$Alias;
                samAccountName=$SamAccountName
            }
            Write-Information -Message "Distribution Group Name=$Name or DisplayName=$DisplayName or EmailAddresses=$PrimarySmtpAddress or Alias=$Alias not found"
            break;
        } else {
            Write-Warning -Message "Distribution Group Name=$Name or DisplayName=$DisplayName or EmailAddresses=$PrimarySmtpAddress or Alias=$Alias found"
        }
    }
    
    Write-Output $returnObject

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
    Write-Warning $warningMessage
    Write-Error $auditMessage
    # exit # use when using multiple try/catch and the script must stop
} finally {
    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps
    $deleteExchangeSessionSplatParams = @{
        Confirm     = $false
        ErrorAction = "Stop"
    }
    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams
    Write-Information "Disconnected from Microsoft Exchange Online"
}
