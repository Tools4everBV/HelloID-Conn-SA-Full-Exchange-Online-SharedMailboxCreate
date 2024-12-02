# HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.

<p align="center">
  <img src="https://github.com/Tools4everBV/HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate/blob/main/Logo.png?raw=true">
</p>

## Table of contents

- [HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate](#helloid-conn-sa-full-exchange-online-sharedmailboxcreate)
  - [Table of contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Remarks](#remarks)
  - [Introduction](#introduction)
      - [Description](#description)
      - [ExchangeOnlineManagement module](#exchangeonlinemanagement-module)
      - [Form Options](#form-options)
      - [Task Actions](#task-actions)
  - [Connector Setup](#connector-setup)
    - [Variable Library - User Defined Variables](#variable-library---user-defined-variables)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Requirements
1. **HelloID Environment**:
   - Set up your _HelloID_ environment.
2. **Exchange Online PowerShell V3 module**:
   - This HelloID Service Automation Delegated Form uses the [Exchange Online PowerShell V3 module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps). A HelloID agent is required to import the Exchange Online module.
3. **Entra ID Application Registration**:
   - App registration with `API permissions` select `APIs my organization uses` search for `Office 365 Exchange Online`. Select `Application permissions`:
     -  `Exchange.ManageAsApp`
   - The following information for the app registration is needed in HelloID:
     - `Application (client) ID`
     - `Directory (tenant) ID`
     - `Secret Value`
4. **Entra ID Role**:
   - The `Exchange Administrator` should provide the required permissions for any task in Exchange Online PowerShell.
     -  To assign the role(s) to your application, navigate to `Roles and administrators`.
     -  Search and select `Exchange Administrator` click `Add assignments`. Select the app registration that you created in step 3.
     -  Click `Next`, assignment type `Active`.

## Remarks
- None at this time.

## Introduction

#### Description
_HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate_ is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements. 

By using this delegated form, you can create a shared mailbox in Exchange Online. The following options are available:
 1. Select a predefined `organization`.
 2. Enter the shared mailbox `name`.
 3. Enter the shared mailbox `alias`.
 4. A data source will `validate` the input in exchange online. If this is valid you can submit the form
 5. The task will `create` the shared mailbox and wil set the properties `MessageCopyForSendOnBehalfEnabled` and `MessageCopyForSentAsEnabled`

#### ExchangeOnlineManagement module
The `ExchangeOnlineManagement` module provide a set of commands to interact with Exchange Online. The commands used are listed in the table below.

| Endpoint    | Description                                                                   |
| ----------- | ----------------------------------------------------------------------------- |
| Get-User    | Required for Get-Mailbox / New-Mailbox / Set-Mailbox                          |
| Get-Mailbox | To validate the input in Exchange Online                                      |
| New-Mailbox | To create the shared mailbox in Exchange Online                               |
| Set-Mailbox | To set attributes that are not supported in New-Mailbox on the shared mailbox |

#### Form Options
The following options are available in the form:

1. **Input and validate data**:
   - Select a `organization`, enter a `name` and enter a `alias`. The Powershell data source will `validate` the input in exchange online.

#### Task Actions
The following actions will be performed after submitting the form:

1. **Creating the shared mailbox in exchange online**:
   - The New-Mailbox command will be used to create the shared mailbox
2. **Updating the shared mailbox in exchange online**:
   - The Set-Mailbox command will be used to set the properties `MessageCopyForSendOnBehalfEnabled` and `MessageCopyForSentAsEnabled`

## Connector Setup
### Variable Library - User Defined Variables
The following user-defined variables are used by the connector. Ensure that you check and set the correct values required to connect to the API.

| Setting             | Description                                                                                |
| ------------------- | ------------------------------------------------------------------------------------------ |
| `EntraOrganization` | The name of the organization to connect to and where the Entra ID App Registration exists. |
| `EntraTenantId`     | The ID to the Tenant in Microsoft Entra ID                                                 |
| `EntraAppId`        | The ID to the App Registration in Microsoft Entra ID                                       |
| `EntraAppSecret`    | The Client Secret to the App Registration in Microsoft Entra ID                            |

## Getting help
> [!TIP]
> _For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages_.

> [!TIP]
>  _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com)_.

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/