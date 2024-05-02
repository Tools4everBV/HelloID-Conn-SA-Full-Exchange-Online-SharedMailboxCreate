# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("mailbox Management","Office 365") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> EntraSecret
$tmpName = @'
EntraSecret
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #2 >> EntraTenantId
$tmpName = @'
EntraTenantId
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> EntraAppID
$tmpName = @'
EntraAppID
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #4 >> EntraOrganization
$tmpName = @'
EntraOrganization
'@ 
$tmpValue = @'
domain.onmicrosoft.com
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
<# Begin: DataSource "Exchange-online-shared-mailbox-generate-table-mail-domains-create" #>
$tmpStaticValue = @'
[{"Naam":"Enyoi","Maildomain":"enyoi.org"},{"Naam":"Tools4ever","Maildomain":"Tools4ever.com"}]
'@ 
$tmpModel = @'
[{"key":"Naam","type":0},{"key":"Maildomain","type":0}]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
Exchange-online-shared-mailbox-generate-table-mail-domains-create
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "2" -DatasourceStaticValue $tmpStaticValue -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "Exchange-online-shared-mailbox-generate-table-mail-domains-create" #>

<# Begin: DataSource "function-check-online-mailbox-exists" #>
$tmpPsScript = @'
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
'@ 
$tmpModel = @'
[{"key":"text","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"Name","type":0,"options":1},{"description":null,"translateDescription":false,"inputFieldType":1,"key":"Alias","type":0,"options":1},{"description":null,"translateDescription":false,"inputFieldType":1,"key":"Organization","type":0,"options":1}]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
function-check-online-mailbox-exists
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "function-check-online-mailbox-exists" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Exchange Online Shared Mailbox - Create" #>
$tmpSchema = @"
[{"templateOptions":{"title":"Retrieving this information from Exchange Online takes an average of +/- 10 seconds.","titleField":"","bannerType":"Info","useBody":true},"type":"textbanner","summaryVisibility":"Show","body":"Please wait so we can validate the input.","requiresTemplateOptions":false,"requiresKey":false,"requiresDataSource":false},{"key":"organization","templateOptions":{"label":"Organization","required":true,"useObjects":false,"useDataSource":true,"useFilter":true,"options":[],"valueField":"Naam","textField":"Naam","dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[]}},"useDefault":true,"defaultSelectorProperty":"Naam"},"type":"dropdown","summaryVisibility":"Show","textOrLabel":"text","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"name","templateOptions":{"label":"name","placeholder":"IT department","required":true,"minLength":2},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"alias","templateOptions":{"label":"Alias","placeholder":"it-department"},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"Validation","templateOptions":{"label":"Validate","readonly":true,"useDataSource":true,"pattern":"^Valid.*","dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[{"propertyName":"Name","otherFieldValue":{"otherFieldKey":"name"}},{"propertyName":"Alias","otherFieldValue":{"otherFieldKey":"alias"}},{"propertyName":"Organization","otherFieldValue":{"otherFieldKey":"organization"}}]}},"displayField":"text","required":true},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Exchange Online Shared Mailbox - Create
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
Exchange Online - Shared Mailbox - Create
'@
$tmpTask = @'
{"name":"Exchange Online - Shared Mailbox - Create","script":"#######################################################################\n# Template: RHo HelloID SA Delegated form task\n# Name: Exchange Online Shared Mailbox - Create\n# Date: 02-05-2024\n#######################################################################\n\n# For basic information about delegated form tasks see:\n# https://docs.helloid.com/en/service-automation/delegated-forms/delegated-form-powershell-scripts.html\n\n# Service automation variables:\n# https://docs.helloid.com/en/service-automation/service-automation-variables.html\n\n#region init\n\n# Set TLS to accept TLS, TLS 1.1 and TLS 1.2\n[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12\n\n$VerbosePreference = \"SilentlyContinue\"\n$InformationPreference = \"Continue\"\n$WarningPreference = \"Continue\"\n\n# global variables (Automation --> Variable libary):\n$TenantId = $EntraTenantId\n$AppID = $EntraAppID\n$Secret = $EntraSecret\n$Organization = $EntraOrganization\n\n# variables configured in form:\n$Maildomain = $form.organization.Maildomain\n$Name = $form.name\n$Alias = $form.alias\n\n# PowerShell commands to import\n$commands = @(\"Get-User\", \"New-Mailbox\", \"Set-Mailbox\")\n#endregion init\n\n#region functions\n\n#endregion functions\n\ntry {\n    #region import module\n    $actionMessage = \"importing $moduleName module\"\n\n    $importModuleParams = @{\n        Name        = \"ExchangeOnlineManagement\"\n        Cmdlet      = $commands\n        ErrorAction = 'Stop'\n    }\n\n    Import-Module @importModuleParams\n    #endregion import module\n\n    #region create access token\n    Write-Verbose \"Creating Access Token\"\n    $actionMessage = \"creating access token\"\n        \n    $body = @{\n        grant_type    = \"client_credentials\"\n        client_id     = \"$AppID\"\n        client_secret = \"$Secret\"\n        resource      = \"https://outlook.office365.com\"\n    }\n\n    $exchangeAccessTokenParams = @{\n        Method          = 'POST'\n        Uri             = \"https://login.microsoftonline.com/$TenantId/oauth2/token\"\n        Body            = $body\n        ContentType     = 'application/x-www-form-urlencoded'\n        UseBasicParsing = $true\n    }\n        \n    $accessToken = (Invoke-RestMethod @exchangeAccessTokenParams).access_token\n    #endregion create access token\n\n    #region connect to Exchange Online\n    Write-Verbose \"Connecting to Exchange Online\"\n    $actionMessage = \"connecting to Exchange Online\"\n\n    $exchangeSessionParams = @{\n        Organization     = $Organization\n        AppID            = $AppID\n        AccessToken      = $accessToken\n        CommandName      = $commands\n        ShowBanner       = $false\n        ShowProgress     = $false\n        TrackPerformance = $false\n        ErrorAction      = 'Stop'\n    }\n    Connect-ExchangeOnline @exchangeSessionParams\n        \n    Write-Information \"Successfully connected to Exchange Online\"\n    #endregion connect to Exchange Online\n\n    #region create shared mailbox\n    $actionMessage = \"creating shared mailbox\"\n    $CreateMailboxParams = @{\n        Shared             = $true\n        Name               = $Name\n        DisplayName        = $Name\n        PrimarySmtpAddress = $Alias.Replace(\" \", \"\") + \"@$Maildomain\"\n        Alias              = $Alias.Replace(\" \", \"\")\n        ErrorAction        = 'Stop'\n    }\n\n    New-Mailbox @CreateMailboxParams\n\n    Write-Information  \"Shared Mailbox [$Name] created successfully\" \n    $Log = @{\n        Action            = \"CreateResource\" # optional. ENUM (undefined = default) \n        System            = \"Exchange Online\" # optional (free format text) \n        Message           = \"Shared Mailbox [$Name] created successfully\"  # required (free format text) \n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $Name # optional (free format text) \n        TargetIdentifier  = $([string]$Alias) # optional (free format text) \n    }\n    #send result back  \n    Write-Information -Tags \"Audit\" -MessageData $log\n    #endregion create shared mailbox\n\n    #region update shared mailbox\n    $actionMessage = \"updating shared mailbox\"\n    Start-Sleep -Seconds 10\n\n    $UpdateMailboxParams = @{\n        Identity                          = \"$($CreateMailboxParams.PrimarySmtpAddress)\"\n        MessageCopyForSendOnBehalfEnabled = $true\n        MessageCopyForSentAsEnabled       = $true\n        ErrorAction                       = 'Stop'\n    }\n\n    Set-Mailbox @UpdateMailboxParams\n \n    Write-Information  \"Shared Mailbox [$Name] updated successfully\" \n    $Log = @{\n        Action            = \"CreateResource\" # optional. ENUM (undefined = default) \n        System            = \"Exchange Online\" # optional (free format text) \n        Message           = \"Shared Mailbox [$Name] updated successfully\"  # required (free format text) \n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $Name # optional (free format text) \n        TargetIdentifier  = $([string]$Alias) # optional (free format text) \n    }\n    #send result back  \n    Write-Information -Tags \"Audit\" -MessageData $log\n    #endregion update shared mailbox\n}\ncatch {\n    $ex = $PSItem\n    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or\n        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {\n        $errorMessage = ($ex.ErrorDetails.Message | Convertfrom-json).error_description\n    }\n    else {\n        $errorMessage = $($ex.Exception.message)\n    }\n\n    Write-Error \"Error $actionMessage for Exchange Online shared mailbox [$Name]. Error: $errorMessage\"\n\n    $Log = @{\n        Action            = \"CreateResource\" # optional. ENUM (undefined = default) \n        System            = \"Exchange Online\" # optional (free format text) \n        Message           = \"Error $actionMessage for Exchange Online shared mailbox [$Name]\" # required (free format text) \n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $Name # optional (free format text) \n        TargetIdentifier  = $([string]$Alias) # optional (free format text) \n    }\n    #send result back  \n    Write-Information -Tags \"Audit\" -MessageData $log\n}","runInCloud":false}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-inbox" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

