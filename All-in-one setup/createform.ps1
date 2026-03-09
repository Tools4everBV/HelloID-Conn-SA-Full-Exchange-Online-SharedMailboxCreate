# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Exchange Online","mailbox Management") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> EntraIdCertificateBase64String
$tmpName = @'
EntraIdCertificateBase64String
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #2 >> EntraIDAppId
$tmpName = @'
EntraIDAppId
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> EntraIdCertificatePassword
$tmpName = @'
EntraIdCertificatePassword
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #4 >> EntraIdOrganization
$tmpName = @'
EntraIdOrganization
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #5 >> EntraIDtenantID
$tmpName = @'
EntraIDtenantID
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#make sure write-information logging is visual
$InformationPreference = "continue"

# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic}
    Write-Information "Using prefilled API credentials"
} else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key}
    Write-Information "Using manual API credentials"
}

# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
} else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}

# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  

# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if($IsCoreCLR -eq $true){
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return ,$r  # Force return value to be an array using a comma
    } else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return ,$r  # Force return value to be an array using a comma
    }
}

function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )

    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false

        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid

            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
        Write-Error "Variable '$Name', message: $_"
    }
}

function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}

        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task

            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid

            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }

    $returnObject.Value = $taskGuid
}

function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter()][String][AllowEmptyString()]$DatasourceRunInCloud,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }

    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
                runInCloud         = $DatasourceRunInCloud;
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
        Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }

    $returnObject.Value = $datasourceGuid
}

function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }

        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }

    $returnObject.Value = $formGuid
}


function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }

        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if(-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups    = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true

            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }

    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}

<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "exchange-online-shared-mailbox-create | EntraID-Check-EmailAddress-Unique" #>
$tmpPsScript = @'
# variables configured in form
$mailPrefix = $datasource.mailPrefix
$mailDomain = $datasource.mailDomain.id
$PrimarySmtpAddress = "$mailPrefix@$mailDomain"

# Build filter - Graph API uses $filter with OData syntax
# Check for mailboxes matching the displayName, mailNickname (alias), primary email or proxy addresses
# This will check ALL users (enabled and disabled), including shared/room/equipment mailboxes
$filter = "`$filter=mailNickname eq '$mailPrefix' or mail eq '$PrimarySmtpAddress' or proxyAddresses/any(x:x eq 'smtp:$PrimarySmtpAddress') or proxyAddresses/any(x:x eq 'SMTP:$PrimarySmtpAddress')"

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$propertiesToSelect = @(
    "id",
    "userPrincipalName",
    "displayName",
    "mailNickname",
    "mail",
    "proxyAddresses"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Ensure the certificate has a private key
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

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
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    # Get Microsoft Entra ID Users
    # Docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Users matching filter [$filter]"

    $getMicrosoftEntraIDUsersSplatParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/users?$filter&`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"
        Headers     = $headers
        Method      = "GET"
        Verbose     = $false
        ErrorAction = "Stop"
    }
    
    $getMicrosoftEntraIDUsersResponse = $null
    $getMicrosoftEntraIDUsersResponse = Invoke-RestMethod @getMicrosoftEntraIDUsersSplatParams

    # Select only specified properties to limit memory usage
    $microsoftEntraIDUsers = $null
    $microsoftEntraIDUsers = $getMicrosoftEntraIDUsersResponse.Value | Select-Object $propertiesToSelect
    Write-Information "Queried Microsoft Entra ID Users matching filter [$filter]. Result count: $(@($microsoftEntraIDUsers).Count)"

    # Check if value is unique and free
    if (($microsoftEntraIDUsers | Measure-Object).Count -gt 0) {
        Write-Warning "Email address is not unique. In use by object with displayName [$($microsoftEntraIDUsers.displayName)], userPrincipalName [$($microsoftEntraIDUsers.userPrincipalName)] mail [$($microsoftEntraIDUsers.mail)] and alias (mailNickName) [$($microsoftEntraIDUsers.mailNickName)]."

        # Send results to HelloID
        $actionMessage = "sending results to HelloID"
        Write-Output "Invalid: Email address is not unique. In use by object with displayName [$($microsoftEntraIDUsers.displayName)], userPrincipalName [$($microsoftEntraIDUsers.userPrincipalName)] mail [$($microsoftEntraIDUsers.mail)] and alias (mailNickName) [$($microsoftEntraIDUsers.mailNickName)]"
    }
    else {
        Write-Information "Email address is unique and free to use."

        # Send results to HelloID
        $actionMessage = "sending results to HelloID"
        Write-Output "Valid: Email address is unique and free to use." 
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
'@ 
$tmpModel = @'
[{"key":"output","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"mailPrefix","type":0,"options":1},{"description":null,"translateDescription":false,"inputFieldType":1,"key":"mailDomain","type":0,"options":1}]
'@ 
$dataSourceGuid_2 = [PSCustomObject]@{} 
$dataSourceGuid_2_Name = @'
exchange-online-shared-mailbox-create | EntraID-Check-EmailAddress-Unique
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_2_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_2) 
<# End: DataSource "exchange-online-shared-mailbox-create | EntraID-Check-EmailAddress-Unique" #>

<# Begin: DataSource "exchange-online-shared-mailbox-create | EntraID-Check-DisplayName-Unique" #>
$tmpPsScript = @'
# variables configured in form
$displayName = $datasource.displayName

# Build filter - Graph API uses $filter with OData syntax
# Check for mailboxes matching the displayName
# This will check ALL users (enabled and disabled), including shared/room/equipment mailboxes
$filter = "`$filter=displayName eq '$displayName'"

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$propertiesToSelect = @(
    "id",
    "userPrincipalName",
    "displayName",
    "mailNickname",
    "mail",
    "proxyAddresses"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Ensure the certificate has a private key
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

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
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    # Get Microsoft Entra ID Users
    # Docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Users matching filter [$filter]"
    
    $getMicrosoftEntraIDUsersSplatParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/users?$filter&`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"
        Headers     = $headers
        Method      = "GET"
        Verbose     = $false
        ErrorAction = "Stop"
    }
    
    $getMicrosoftEntraIDUsersResponse = $null
    $getMicrosoftEntraIDUsersResponse = Invoke-RestMethod @getMicrosoftEntraIDUsersSplatParams

    # Select only specified properties to limit memory usage
    $microsoftEntraIDUsers = $null
    $microsoftEntraIDUsers = $getMicrosoftEntraIDUsersResponse.Value #| Select-Object $propertiesToSelect
    Write-Information "Queried Microsoft Entra ID Users matching filter [$filter]. Result count: $(@($microsoftEntraIDUsers).Count)"

    # Check if value is unique and free
    if (($microsoftEntraIDUsers | Measure-Object).Count -gt 0) {
        Write-Warning "Display name is not unique. In use by object with displayName [$($microsoftEntraIDUsers.displayName)], userPrincipalName [$($microsoftEntraIDUsers.userPrincipalName)] mail [$($microsoftEntraIDUsers.mail)] and alias (mailNickName) [$($microsoftEntraIDUsers.mailNickName)]."

        # Send results to HelloID
        $actionMessage = "sending results to HelloID"
        Write-Output "Invalid: Display name is not unique. In use by object with displayName [$($microsoftEntraIDUsers.displayName)], userPrincipalName [$($microsoftEntraIDUsers.userPrincipalName)] mail [$($microsoftEntraIDUsers.mail)] and alias (mailNickName) [$($microsoftEntraIDUsers.mailNickName)]"
    }
    else {
        Write-Information "Display name is unique and free to use."

        # Send results to HelloID
        $actionMessage = "sending results to HelloID"
        Write-Output "Valid: Display name is unique and free to use." 
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
'@ 
$tmpModel = @'
[{"key":"output","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"displayName","type":0,"options":1}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
exchange-online-shared-mailbox-create | EntraID-Check-DisplayName-Unique
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "exchange-online-shared-mailbox-create | EntraID-Check-DisplayName-Unique" #>

<# Begin: DataSource "exchange-online-shared-mailbox-create | EntraID-Check-Alias-Unique" #>
$tmpPsScript = @'
# variables configured in form
$alias = $datasource.alias
$mailDomain = $datasource.mailDomain.id
$PrimarySmtpAddress = "$alias@$mailDomain"

# Build filter - Graph API uses $filter with OData syntax
# Check for mailboxes matching the displayName, mailNickname (alias), primary email or proxy addresses
# This will check ALL users (enabled and disabled), including shared/room/equipment mailboxes
$filter = "`$filter=mailNickname eq '$alias' or mail eq '$PrimarySmtpAddress' or proxyAddresses/any(x:x eq 'smtp:$PrimarySmtpAddress') or proxyAddresses/any(x:x eq 'SMTP:$PrimarySmtpAddress')"

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$propertiesToSelect = @(
    "id",
    "userPrincipalName",
    "displayName",
    "mailNickname",
    "mail",
    "proxyAddresses"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Ensure the certificate has a private key
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

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
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    # Get Microsoft Entra ID Users
    # Docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Users matching filter [$filter]"

    $getMicrosoftEntraIDUsersSplatParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/users?$filter&`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"
        Headers     = $headers
        Method      = "GET"
        Verbose     = $false
        ErrorAction = "Stop"
    }
    
    $getMicrosoftEntraIDUsersResponse = $null
    $getMicrosoftEntraIDUsersResponse = Invoke-RestMethod @getMicrosoftEntraIDUsersSplatParams

    # Select only specified properties to limit memory usage
    $microsoftEntraIDUsers = $null
    $microsoftEntraIDUsers = $getMicrosoftEntraIDUsersResponse.Value | Select-Object $propertiesToSelect
    Write-Information "Queried Microsoft Entra ID Users matching filter [$filter]. Result count: $(@($microsoftEntraIDUsers).Count)"

    # Check if value is unique and free
    if (($microsoftEntraIDUsers | Measure-Object).Count -gt 0) {
        Write-Warning "Alias is not unique. In use by object with displayName [$($microsoftEntraIDUsers.displayName)], userPrincipalName [$($microsoftEntraIDUsers.userPrincipalName)] mail [$($microsoftEntraIDUsers.mail)] and alias (mailNickName) [$($microsoftEntraIDUsers.mailNickName)]."

        # Send results to HelloID
        $actionMessage = "sending results to HelloID"
        Write-Output "Invalid: Alias is not unique. In use by object with displayName [$($microsoftEntraIDUsers.displayName)], userPrincipalName [$($microsoftEntraIDUsers.userPrincipalName)] mail [$($microsoftEntraIDUsers.mail)] and alias (mailNickName) [$($microsoftEntraIDUsers.mailNickName)]"
    }
    else {
        Write-Information "Alias is unique and free to use."

        # Send results to HelloID
        $actionMessage = "sending results to HelloID"
        Write-Output "Valid: Alias is unique and free to use." 
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
'@ 
$tmpModel = @'
[{"key":"output","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"alias","type":0,"options":1},{"description":null,"translateDescription":false,"inputFieldType":1,"key":"mailDomain","type":0,"options":1}]
'@ 
$dataSourceGuid_3 = [PSCustomObject]@{} 
$dataSourceGuid_3_Name = @'
exchange-online-shared-mailbox-create | EntraID-Check-Alias-Unique
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_3_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_3) 
<# End: DataSource "exchange-online-shared-mailbox-create | EntraID-Check-Alias-Unique" #>

<# Begin: DataSource "exchange-online-shared-mailbox-create | EntraID-Get-All-MailDomains" #>
$tmpPsScript = @'
# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$propertiesToSelect = @(
    "id"
    , "isVerified"
    , "supportedServices"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Ensure the certificate has a private key
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

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
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization" = "Bearer $($entraToken)"
        "Accept"        = "application/json"
        "Content-Type"  = "application/json"
    }

    # Get Microsoft Entra ID Domains
    # Docs: https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Domains"

    $getMicrosoftEntraIDDomainsSplatParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/domains"#?`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"#?$filter"
        Headers     = $headers
        Method      = "GET"
        Verbose     = $false
        ErrorAction = "Stop"
    }
    $getMicrosoftEntraIDDomainsResponse = $null
    $getMicrosoftEntraIDDomainsResponse = Invoke-RestMethod @getMicrosoftEntraIDDomainsSplatParams

    # Select only specified properties to limit memory usage
    $microsoftEntraIDDomains = $null
    $microsoftEntraIDDomains = $getMicrosoftEntraIDDomainsResponse.Value | Select-Object $propertiesToSelect
    Write-Information "Queried Microsoft Entra ID Domains. Result count: $(@($microsoftEntraIDDomains).Count)"

    # Filter for verified domains only and where Email is supported - not support by Graph API filter query
    $actionMessage = "filtering for verified domains only and where Email is supported"
    $microsoftEntraIDDomains = $microsoftEntraIDDomains | Where-Object { $_.isVerified -eq $true -and $_.supportedServices -like '*Email*' }
    Write-Information "Filter for verified domains only and where Email is supported. Result count: $(@($microsoftEntraIDDomains).Count)"

    # Send results to HelloID
    $actionMessage = "sending results to HelloID"
    $microsoftEntraIDDomains | Sort-Object -Property id | ForEach-Object {
        Write-Output $_
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}

'@ 
$tmpModel = @'
[{"key":"id","type":0},{"key":"isVerified","type":0},{"key":"supportedServices","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
exchange-online-shared-mailbox-create | EntraID-Get-All-MailDomains
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "exchange-online-shared-mailbox-create | EntraID-Get-All-MailDomains" #>

<# Begin: DataSource "exchange-online-shared-mailbox-create | EntraID-Get-All-Users" #>
$tmpPsScript = @'
$filter = "`$filter=userType eq 'Member'" # Get all users, excluding guest users

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
# Properties to select - Select only needed properties to limit memory usage and speed up processing
$propertiesToSelect = @(
    "id",
    "userPrincipalName",
    "displayName",
    "mail"
)

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Ensure the certificate has a private key
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

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
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    # Get Microsoft Entra ID Users
    # API docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Users matching filter [$filter]"
    $microsoftEntraIDUsers = [System.Collections.ArrayList]@()
    do {
        $getMicrosoftEntraIDUsersSplatParams = @{
            Uri         = "https://graph.microsoft.com/v1.0/users?$filter&`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"
            Headers     = $headers
            Method      = "GET"
            Verbose     = $false
            ErrorAction = "Stop"
        }
        if (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDUsersResponse.'@odata.nextLink')) {
            $getMicrosoftEntraIDUsersSplatParams["Uri"] = $getMicrosoftEntraIDUsersResponse.'@odata.nextLink'
        }
        
        $getMicrosoftEntraIDUsersResponse = $null
        $getMicrosoftEntraIDUsersResponse = Invoke-RestMethod @getMicrosoftEntraIDUsersSplatParams
    
        # Select only specified properties to limit memory usage
        $getMicrosoftEntraIDUsersResponse.Value = $getMicrosoftEntraIDUsersResponse.Value | Select-Object $propertiesToSelect

        if ($getMicrosoftEntraIDUsersResponse.Value -is [array]) {
            [void]$microsoftEntraIDUsers.AddRange($getMicrosoftEntraIDUsersResponse.Value)
        }
        else {
            [void]$microsoftEntraIDUsers.Add($getMicrosoftEntraIDUsersResponse.Value)
        }
    } while (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDUsersResponse.'@odata.nextLink'))
    Write-Information "Queried Microsoft Entra ID Users matching filter [$filter]. Result count: $(@($microsoftEntraIDUsers).Count)"

    # Send results to HelloID
    $actionMessage = "sending results to HelloID"
    $microsoftEntraIDUsers | Sort-Object -Property displayName | ForEach-Object {
        Write-Output $_
    }
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
  
'@ 
$tmpModel = @'
[{"key":"id","type":0},{"key":"userPrincipalName","type":0},{"key":"displayName","type":0},{"key":"mail","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_4 = [PSCustomObject]@{} 
$dataSourceGuid_4_Name = @'
exchange-online-shared-mailbox-create | EntraID-Get-All-Users
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_4_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_4) 
<# End: DataSource "exchange-online-shared-mailbox-create | EntraID-Get-All-Users" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Exchange online - Shared Mailbox - Create" #>
$tmpSchema = @"
[{"label":"Create mailbox","fields":[{"key":"displayName","templateOptions":{"label":"Display name","placeholder":"IT department","required":true,"pattern":"^[A-Za-z0-9À-ÿ .,_'-]{1,256}$","minLength":1,"maxLength":256},"validation":{"messages":{"pattern":""}},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"nameValidation","templateOptions":{"label":"Display name validation","required":true,"pattern":"^Valid.*","readonly":true,"placeholder":"Display name will be validated on uniqnueness in Entra ID","useDataSource":true,"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[{"propertyName":"displayName","otherFieldValue":{"otherFieldKey":"displayName"}}]}},"displayField":"output"},"hideExpression":"!model[\"displayName\"]","type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"templateOptions":{"title":"This address will be created as the primary SMTP address of the mailbox","titleField":"","bannerType":"Info","useBody":false},"type":"textbanner","summaryVisibility":"Show","body":"Text Banner Content","requiresTemplateOptions":false,"requiresKey":false,"requiresDataSource":false},{"key":"formRow","templateOptions":{},"fieldGroup":[{"key":"mailPrefix","templateOptions":{"label":"Email address","placeholder":"it-department","required":true,"pattern":"^[A-Za-z0-9]([A-Za-z0-9._-]*[A-Za-z0-9])?$","minLength":1,"maxLength":200},"validation":{"messages":{"pattern":""}},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"mailDomain","templateOptions":{"label":"Mail domain","required":true,"useObjects":false,"useDataSource":true,"useFilter":true,"options":[],"valueField":"id","textField":"id","dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[]}},"defaultSelectorProperty":"mailDomain"},"type":"dropdown","summaryVisibility":"Show","textOrLabel":"text","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}],"type":"formrow","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"mailValidation","templateOptions":{"label":"Email address validation","useDataSource":true,"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_2","input":{"propertyInputs":[{"propertyName":"mailPrefix","otherFieldValue":{"otherFieldKey":"mailPrefix"}},{"propertyName":"mailDomain","otherFieldValue":{"otherFieldKey":"mailDomain"}}]}},"displayField":"output","required":true,"readonly":true,"pattern":"^Valid.*","placeholder":"Email address will be validated on uniqnueness in Entra ID"},"hideExpression":"!model[\"mailPrefix\"]","type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"templateOptions":{"title":"The alias (mailNickname) is a short name used as an internal identifier for the mailbox","titleField":"","bannerType":"Info","useBody":true},"type":"textbanner","summaryVisibility":"Show","body":"**The alias is a short name used as an internal identifier for the new mailbox.**\r\n\r\nIt is not the same as the email address.  \r\nThe alias can only have **one** value, and it will be created exactly as entered.\r\n\r\nIf you do not provide an alias, the username portion of the email address will be used automatically.\r\n\r\nUse only letters, numbers, and periods (no spaces).  \r\nDo not include a domain (no \"@\").","requiresTemplateOptions":false,"requiresKey":false,"requiresDataSource":false},{"key":"alias","templateOptions":{"label":"Alias (mailNickName)","pattern":"^[A-Za-z0-9]([A-Za-z0-9._-]*[A-Za-z0-9])?$","placeholder":"it-dep"},"validation":{"messages":{"pattern":""}},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"aliasValidation","templateOptions":{"label":"Alias (mailNickName) validation","readonly":true,"pattern":"^Valid.*","useDataSource":true,"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_3","input":{"propertyInputs":[{"propertyName":"alias","otherFieldValue":{"otherFieldKey":"alias"}},{"propertyName":"mailDomain","otherFieldValue":{"otherFieldKey":"mailDomain"}}]}},"displayField":"output","required":true,"placeholder":"Alias (mailNickName) will be validated on uniqnueness in Entra ID"},"hideExpression":"!model[\"alias\"]","type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]},{"label":"Mailbox Permissions","fields":[{"key":"permission","templateOptions":{"label":"Permission","useObjects":true,"useDataSource":false,"useFilter":false,"options":[{"value":"fullaccess","text":"Full Access"},{"value":"sendas","text":"Send As"},{"value":"sendonbehalf","text":"Send on Behalf"}]},"type":"dropdown","summaryVisibility":"Show","textOrLabel":"text","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"blnIncludeSendAs","templateOptions":{"label":"Include Send As with Full Access?","useSwitch":true,"checkboxLabel":"Include Send As with Full Access"},"hideExpression":"model[\"permission\"]!=='fullaccess'","type":"boolean","defaultValue":true,"summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"usersToAdd","templateOptions":{"label":"User(s) to grant permission","grid":{"columns":[{"headerName":"User Principal Name","field":"userPrincipalName"},{"headerName":"Display Name","field":"displayName"},{"headerName":"Mail","field":"mail"}],"height":300,"rowSelection":"multiple"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_4","input":{"propertyInputs":[]}},"useFilter":true,"useDefault":false,"allowCsvDownload":false},"hideExpression":"!model[\"permission\"]","type":"multiselectgrid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true}]}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Exchange online - Shared Mailbox - Create
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if(-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)){
    foreach($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
        
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        } catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if($null -ne $delegatedFormAccessGroupGuids){
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}

$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $response = $response | Where-Object {$_.name.en -eq $category}
    
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
    
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100

        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid

        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
Exchange online - Shared Mailbox - Create
'@
$tmpTask = @'
{"name":"Exchange online - Shared Mailbox - Create","script":"# variables configured in form\n$mailboxDisplayName = $form.displayName\n$mailboxMailPrefix = $form.mailPrefix\n$mailboxMailDomain = $form.mailDomain.id\n$mailboxPrimarySmtpAddress = \"$($mailboxMailPrefix)@$($mailboxMailDomain)\"\n$mailboxAlias = $form.alias\n\n$permissions = @($form.permission)\n$blnIncludeSendAs = [System.Convert]::ToBoolean($form.blnIncludeSendAs)\nif($blnIncludeSendAs -eq $true) {\n    $permissions += 'sendas'\n}\n$usersToAdd = $form.usersToAdd\n\n# Global variables\n# Outcommented as these are set from Global Variables\n# $EntraIdOrganization = \"\"\n# $EntraIdAppId = \"\"\n# $EntraIdCertificateBase64String = \"\"\n# $EntraIdCertificatePassword = \"\"\n\n# Fixed values\n$commands = @(\n    \"New-Mailbox\",\n    \"Set-Mailbox\",\n    \"Add-MailboxPermission\",\n    \"Add-RecipientPermission\",\n    \"Set-Mailbox\",\n    \"Remove-MailboxPermission\",\n    \"Remove-RecipientPermission\"\n)\n\n# Enable TLS1.2\n[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12\n\n# Set debug logging\n$VerbosePreference = \"SilentlyContinue\"\n$InformationPreference = \"Continue\"\n$WarningPreference = \"Continue\"\n\n#region functions\nfunction Get-MSEntraCertificate {\n    [CmdletBinding()]\n    param(\n        [Parameter(Mandatory)]\n        [ValidateNotNullOrEmpty()]\n        [string]\n        $CertificateBase64String,\n        \n        [Parameter(Mandatory)]\n        [ValidateNotNullOrEmpty()]\n        [string]\n        $CertificatePassword\n    )\n    try {\n        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)\n        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)\n        Write-Output $certificate\n    }\n    catch {\n        $PSCmdlet.ThrowTerminatingError($_)\n    }\n}\n#endregion functions\n\ntry {\n    # Import module\n    $actionMessage = \"importing module [ExchangeOnlineManagement]\"\n        \n    $importModuleSplatParams = @{\n        Name        = \"ExchangeOnlineManagement\"\n        Cmdlet      = $commands\n        Verbose     = $false\n        ErrorAction = \"Stop\"\n    }\n\n    $null = Import-Module @importModuleSplatParams\n\n    Write-Verbose \"Imported module [ExchangeOnlineManagement]\"\n\n    # Convert base64 certificate string to certificate object\n    $actionMessage = \"converting base64 certificate string to certificate object\"\n\n    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword\n\n    Write-Verbose \"Converted base64 certificate string to certificate object\"\n\n    # Connect to Microsoft Exchange Online\n    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps\n    $actionMessage = \"connecting to Microsoft Exchange Online\"\n\n    $createExchangeSessionSplatParams = @{\n        Organization          = $EntraIdOrganization\n        AppID                 = $EntraIdAppId\n        Certificate           = $certificate\n        CommandName           = $commands\n        ShowBanner            = $false\n        ShowProgress          = $false\n        TrackPerformance      = $false\n        SkipLoadingCmdletHelp = $true\n        SkipLoadingFormatData = $true\n        ErrorAction           = \"Stop\"\n    }\n\n    $null = Connect-ExchangeOnline @createExchangeSessionSplatParams\n\n    # Create shared mailbox\n    $action = \"CreateResource\"\n    $actionMessage = \"creating shared mailbox with displayname [$($mailboxDisplayName)] and PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]\"\n\n    $CreateMailboxParams = @{\n        Shared             = $true\n        Name               = $mailboxDisplayName\n        DisplayName        = $mailboxDisplayName\n        PrimarySmtpAddress = $mailboxPrimarySmtpAddress\n        ErrorAction        = 'Stop'\n    }\n\n    # Add Alias if specified\n    if (-not [string]::IsNullOrEmpty($mailboxAlias)) {\n        $CreateMailboxParams[\"Alias\"] = $mailboxAlias\n    }\n\n    $null = New-Mailbox @CreateMailboxParams\n\n    # Send auditlog to HelloID\n    $Log = @{\n        Action            = $action # optional. ENUM (undefined = default) \n        System            = \"ExchangeOnline\" # optional (free format text) \n        Message           = \"Created shared mailbox with displayname [$($mailboxDisplayName)] and PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]\"  # required (free format text) \n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $mailboxDisplayName # optional (free format text) \n        TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) \n    }\n    Write-Information -Tags \"Audit\" -MessageData $log\n\n    # Update shared mailbox to enable MessageCopyForSendOnBehalfEnabled and MessageCopyForSentAsEnabled (can only be done after mailbox is created)\n    $action = \"UpdateResource\"\n    $actionMessage = \"updating MessageCopyForSendOnBehalfEnabled and MessageCopyForSentAsEnabled to [true] for shared mailbox with PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]\"\n\n    # Wait for the mailbox to be created before updating\n    Start-Sleep -Seconds 10\n\n    $UpdateMailboxParams = @{\n        Identity                          = $mailboxPrimarySmtpAddress\n        MessageCopyForSendOnBehalfEnabled = $true\n        MessageCopyForSentAsEnabled       = $true\n        ErrorAction                       = 'Stop'\n    }\n\n    $null = Set-Mailbox @UpdateMailboxParams\n\n    # Send auditlog to HelloID\n    $Log = @{\n        Action            = $action # optional. ENUM (undefined = default) \n        System            = \"ExchangeOnline\" # optional (free format text) \n        Message           = \"Updated MessageCopyForSendOnBehalfEnabled and MessageCopyForSentAsEnabled to [true] for shared mailbox with PrimarySmtpAddress [$($mailboxPrimarySmtpAddress)]\"  # required (free format text) \n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $mailboxDisplayName # optional (free format text) \n        TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) \n    }\n    Write-Information -Tags \"Audit\" -MessageData $log\n\n    if(-not [string]::IsNullOrEmpty($permissions) -and @($usersToAdd).Count -ge 1){\n        # Grant users permissions to shared mailbox\n        $actionMessage = \"granting permission [$($permissions -Join ';')] to shared mailbox to mailbox [$($mailboxDisplayName) ($($mailboxPrimarySmtpAddress))] for users\"\n        foreach ($userToAdd in $usersToAdd) {\n            foreach($permission in $permissions) {\n                switch ($permission) {\n                    \"fullaccess\" {\n                        # Grant Full Access to shared mailbox\n                        try {\n                            $actionMessage = \"granting permission [FullAccess] to mailbox [$($mailboxDisplayName) ($($mailboxPrimarySmtpAddress))] for user [$($userToAdd.userPrincipalName) ($($userToAdd.id))]\"\n\n                            $FullAccessPermissionSplatParams = @{\n                                Identity      = $mailboxPrimarySmtpAddress  # of $mailbox.UserPrincipalName\n                                User          = $userToAdd.id\n                                AccessRights  = \"FullAccess\"\n                                AutoMapping   = [bool]$AutoMapping\n                                ErrorAction   = \"Stop\"\n                                WarningAction = \"SilentlyContinue\"\n                            }\n                            $addFullAccessPermission = Add-MailboxPermission @FullAccessPermissionSplatParams\n\n                            # Send auditlog to HelloID\n                            $Log = @{\n                                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \n                                System            = \"Exchange\" # optional (free format text) \n                                Message           = \"Successfully granted permission [FullAccess] to mailbox [$($mailboxDisplayName) ($($mailboxPrimarySmtpAddress))] for user [$($userToAdd.userPrincipalName) ($($userToAdd.id))]\" # required (free format text) \n                                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n                                TargetDisplayName = $mailboxDisplayName # optional (free format text)\n                                TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text)\n                            }\n                            Write-Information -Tags \"Audit\" -MessageData $log\n                        }\n                        catch {\n                            $ex = $PSItem\n                            if (-not [string]::IsNullOrEmpty($ex.Exception.Data.RemoteException.Message)) {\n                                $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Data.RemoteException.Message)\"\n                                $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Data.RemoteException.Message)\"\n                            }\n                            else {\n                                $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\n                                $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\n                            }\n\n                            $Log = @{\n                                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \n                                System            = \"ExchangeOnline\" # optional (free format text) \n                                Message           = $auditMessage # required (free format text) \n                                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n                                TargetDisplayName = $mailboxDisplayName # optional (free format text) \n                                TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) \n                            }\n                            Write-Information -Tags \"Audit\" -MessageData $log\n                            Write-Warning $warningMessage\n                            Write-Error $auditMessage\n                        }\n                        break\n                    }\n                    \n                    \"sendas\" {\n                        # Grant Send As to shared mailbox\n                        try {\n                            $actionMessage = \"granting permission [Send As] to mailbox [$($mailboxDisplayName) ($($mailboxPrimarySmtpAddress))] for user [$($userToAdd.userPrincipalName) ($($userToAdd.id))]\"\n\n                            $sendAsPermissionSplatParams = @{\n                                Identity     = $mailboxPrimarySmtpAddress\n                                Trustee      = $userToAdd.id\n                                AccessRights = \"SendAs\"\n                                Confirm      = $false\n                                ErrorAction  = \"Stop\"\n                            } \n                            $addSendAsPermission = Add-RecipientPermission @sendAsPermissionSplatParams\n\n                            # Send auditlog to HelloID\n                            $Log = @{\n                                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \n                                System            = \"Exchange\" # optional (free format text) \n                                Message           = \"Successfully granted permission [Send As] to mailbox [$($mailboxDisplayName) ($($mailboxPrimarySmtpAddress))] for user [$($userToAdd.userPrincipalName) ($($userToAdd.id))\" # required (free format text) \n                                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n                                TargetDisplayName = $mailboxDisplayName # optional (free format text)\n                                TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text)\n                            }\n                            Write-Information -Tags \"Audit\" -MessageData $log\n                        }\n                        catch {\n                            $ex = $PSItem\n                            if (-not [string]::IsNullOrEmpty($ex.Exception.Data.RemoteException.Message)) {\n                                $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Data.RemoteException.Message)\"\n                                $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Data.RemoteException.Message)\"\n                            }\n                            else {\n                                $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\n                                $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\n                            }\n\n                            $Log = @{\n                                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \n                                System            = \"ExchangeOnline\" # optional (free format text) \n                                Message           = $auditMessage # required (free format text) \n                                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n                                TargetDisplayName = $mailboxDisplayName # optional (free format text) \n                                TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) \n                            }\n                            Write-Information -Tags \"Audit\" -MessageData $log\n                            Write-Warning $warningMessage\n                            Write-Error $auditMessage\n                        }\n                        break\n                    }\n\n                    \"sendonbehalf\" {\n                        # Grant Send on Behalf to shared mailbox\n                        try {\n                            $actionMessage = \"granting permission [Send on Behalf] to mailbox [$($mailboxDisplayName) ($($mailboxPrimarySmtpAddress))] for user [$($userToAdd.userPrincipalName) ($($userToAdd.id))]\"\n\n                            $SendonBehalfPermissionSplatParams = @{\n                                Identity            = $mailboxPrimarySmtpAddress\n                                GrantSendOnBehalfTo = @{ add = \"$($userToAdd.id)\" }\n                                Confirm             = $false\n                                ErrorAction         = \"Stop\"\n                            }\n                            Write-Warning ($SendonBehalfPermissionSplatParams | ConvertTo-Json -Depth 10)\n                            $addSendonBehalfPermission = Set-Mailbox @SendonBehalfPermissionSplatParams\n\n                            # Send auditlog to HelloID\n                            $Log = @{\n                                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \n                                System            = \"Exchange\" # optional (free format text) \n                                Message           = \"Successfully granted permission [Send on Behalf] to mailbox [$($mailboxDisplayName) ($($mailboxPrimarySmtpAddress))] for user [$($userToAdd.userPrincipalName) ($($userToAdd.id))]\" # required (free format text) \n                                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n                                TargetDisplayName = $mailboxDisplayName # optional (free format text)\n                                TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text)\n                            }\n                            Write-Information -Tags \"Audit\" -MessageData $log\n                        }\n                        catch {\n                            $ex = $PSItem\n                            if (-not [string]::IsNullOrEmpty($ex.Exception.Data.RemoteException.Message)) {\n                                $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Data.RemoteException.Message)\"\n                                $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Data.RemoteException.Message)\"\n                            }\n                            else {\n                                $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\n                                $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\n                            }\n\n                            $Log = @{\n                                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \n                                System            = \"ExchangeOnline\" # optional (free format text) \n                                Message           = $auditMessage # required (free format text) \n                                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n                                TargetDisplayName = $mailboxDisplayName # optional (free format text)\n                                TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text)\n                            }\n                            Write-Information -Tags \"Audit\" -MessageData $log\n                            Write-Warning $warningMessage\n                            Write-Error $auditMessage\n                        }\n                        break\n                    }\n                }\n            }\n        }\n    }\n}\ncatch {\n    $ex = $PSItem\n    if (-not [string]::IsNullOrEmpty($ex.Exception.Data.RemoteException.Message)) {\n        $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Data.RemoteException.Message)\"\n        $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Data.RemoteException.Message)\"\n    }\n    else {\n        $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\n        $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\n    }\n\n    $Log = @{\n        Action            = $action # optional. ENUM (undefined = default) \n        System            = \"ExchangeOnline\" # optional (free format text) \n        Message           = $auditMessage # required (free format text) \n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $mailboxDisplayName # optional (free format text) \n        TargetIdentifier  = $mailboxPrimarySmtpAddress # optional (free format text) \n    }\n    \n    Write-Information -Tags \"Audit\" -MessageData $log\n    Write-Warning $warningMessage\n    Write-Error $auditMessage\n}\nfinally {\n    # Docs: https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps\n    $deleteExchangeSessionSplatParams = @{\n        Confirm     = $false\n        ErrorAction = \"Stop\"\n    }\n    $null = Disconnect-ExchangeOnline @deleteExchangeSessionSplatParams\n}","runInCloud":false}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-inbox" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

