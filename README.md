# HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate

| :information_source: Information |
| :------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as organization name, application ID, certificate, etc. You might need to coordinate with the client's application manager before implementing this connector. |

## Description

HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate is a delegated form designed for use with HelloID Service Automation (SA). It can be imported into HelloID and customized according to your requirements.

By using this delegated form, you can create a shared mailbox in Exchange Online. The following options are available:

1. Enter the display name for the shared mailbox
2. Enter the email address (prefix) and select a mail domain
   > email address is validated for uniqueness in Microsoft Entra ID
3. Optionally provide an alias for the shared mailbox
   > Alias is validated for uniqueness in Microsoft Entra ID
4. Create and configure the shared mailbox
   > A shared mailbox is created using the provided display name and email address. The mailbox is then configured to receive copies of messages that are sent on behalf of it or as the mailbox itself.

## Getting started

### Requirements

#### App Registration & Certificate Setup

Before implementing this connector, make sure to configure a Microsoft Entra ID App Registration. During the setup process, you'll create a new App Registration in the Entra portal, assign the necessary API permissions, and generate and assign a certificate.

Follow the official Microsoft documentation for creating an App Registration and setting up certificate-based authentication:

- [App-only authentication with certificate (Exchange Online)](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication)

#### HelloID-specific configuration

Once you have completed the Microsoft setup and followed their best practices, configure the following HelloID-specific requirements.

- **API Permissions** (Application permissions):
  - `User.Read.All` - To validate email uniqueness by listing users via Graph API
  - `Domain.Read.All` - To retrieve verified domains for the mail domain dropdown
  - `Exchange.ManageAsApp` - To create and manage shared mailboxes
- **Entra ID Role assignment:**
  - Assign the **Exchange Administrator** role to the App Registration
- **Certificate:**
  - Upload the public key file (.cer) in Entra ID
  - Provide the certificate as a Base64 string in HelloID. For instructions on creating the certificate and obtaining the base64 string, refer to our forum post: [Setting up a certificate for Microsoft Graph API in HelloID connectors](https://forum.helloid.com/forum/helloid-provisioning/5338-instruction-setting-up-a-certificate-for-microsoft-graph-api-in-helloid-connectors#post5338)

### Connection settings

The following global variables must be configured in HelloID when importing and configuring the delegated form.

| Setting | Description | Mandatory |
| --- | --- | --- |
| EntraIdOrganization | The Entra organization name (domain) | Yes |
| EntraIdTenantId | The Entra tenant ID (GUID) | Yes |
| EntraIdAppId | The unique identifier (ID) of the App Registration in Microsoft Entra ID | Yes |
| EntraIdCertificateBase64String | The Base64-encoded string representation of the app certificate | Yes |
| EntraIdCertificatePassword | The password associated with the app certificate | Yes |

## Remarks

### Email Address & Alias Validation via Graph API

- **Performance optimization**: Instead of using the Exchange Online cmdlet `Get-Mailbox` (which can take 30+ seconds per query), the connector uses the Microsoft Graph API to validate email address and alias uniqueness
- **Validation scope**: Checks for uniqueness across all types of objects in Entra ID (users, shared mailboxes, room mailboxes, equipment mailboxes, etc.)
- **Graph API filters**: Uses OData `$filter` queries on the following properties:
  - `mailNickname` - Mail nickname/alias
  - `mail` - Primary SMTP address
  - `proxyAddresses` - Proxy addresses (both smtp and SMTP variants)
- **Two separate data sources**:
  - `EntraID-Check-EmailAddress-Unique` - Validates the email prefix uniqueness
  - `EntraID-Check-Alias-Unique` - Validates the alias uniqueness (if an alias is provided)

### Mailbox Creation and Configuration

#### Shared Mailbox Creation Process

When the form is submitted, the following process occurs in Exchange Online:

1. **Create Shared Mailbox** (`New-Mailbox` cmdlet)
   - **Mailbox type**: Shared mailbox
   - **Display Name**: Set to the value entered in the form
   - **Mailbox Name**: Set to the value entered in the form
   - **Primary SMTP Address**: Constructed from email prefix + selected mail domain (e.g., `prefix@domain.com`)
   - **Alias**: If provided, uses the specified alias; otherwise defaults to the email prefix
   - **Recipient Type**: SharedMailbox

2. **Wait for Exchange Processing**
   - A 10-second delay is implemented to allow Exchange Online to finalize the mailbox creation
   - This ensures the mailbox object is fully available for subsequent configuration

3. **Configure Mailbox Settings** (`Set-Mailbox` cmdlet)
   - The following properties are explicitly enabled:
     - **MessageCopyForSendOnBehalfEnabled**: Set to `$true`
       - Enables the mailbox to receive a copy of messages when someone sends on behalf of the mailbox
     - **MessageCopyForSentAsEnabled**: Set to `$true`
       - Enables the mailbox to receive a copy of messages when someone sends as the mailbox
   - These settings ensure the shared mailbox owner receives copies of all messages sent using the mailbox's identity or delegated permissions

## Development resources

### API endpoints

The following Microsoft Graph API endpoints are used by the connector:

| Endpoint | Description |
| --- | --- |
| `/v1.0/domains` | Retrieve all verified domains with Email support for the mail domain dropdown |
| `/v1.0/users` | Search and retrieve users to validate email address and alias uniqueness |

### PowerShell Cmdlets

The following PowerShell cmdlets are used by the connector:

| Cmdlet | Description |
| --- | --- |
| `Connect-ExchangeOnline` | Establish session to Exchange Online using certificate-based app-only authentication |
| `New-Mailbox` | Create the shared mailbox with specified properties |
| `Set-Mailbox` | Configure mailbox delegation settings after creation |
| `Disconnect-ExchangeOnline` | Close the Exchange Online session |

### Documentation

For more information on the APIs and PowerShell cmdlets used in this connector, please refer to:

**Microsoft Graph API:**
- [Authentication with certificate](https://learn.microsoft.com/graph/auth-v2-service)
- [List domains](https://learn.microsoft.com/graph/api/domain-list)
- [List users - Advanced query capabilities](https://learn.microsoft.com/graph/aad-advanced-queries)
- [User resource reference](https://learn.microsoft.com/graph/api/user-list)

**Exchange Online PowerShell:**
- [Exchange Online PowerShell overview](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell)
- [Connect-ExchangeOnline](https://learn.microsoft.com/powershell/module/exchange/connect-exchangeonline)
- [New-Mailbox](https://learn.microsoft.com/powershell/module/exchange/new-mailbox)
- [Set-Mailbox](https://learn.microsoft.com/powershell/module/exchange/set-mailbox)
- [Disconnect-ExchangeOnline](https://learn.microsoft.com/powershell/module/exchange/disconnect-exchangeonline)

## Getting help

> 💡 **Tip:**  
> For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages.

## HelloID docs

The official HelloID documentation can be found at: [https://docs.helloid.com/](https://docs.helloid.com/)