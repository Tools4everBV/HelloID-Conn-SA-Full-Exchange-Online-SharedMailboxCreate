# HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate

| :information_source: Information |
| :------------------------------- |
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as organization name, application ID, certificate, etc. You might need to coordinate with the client's application manager before implementing this connector. |

## Description
HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements.

By using this delegated form, you can create a shared mailbox in Exchange Online. The form workflow includes:
 1. Select the organization/mail domain
 2. Enter the mailbox name and alias
 3. Validate the mailbox availability (name/alias/primary SMTP)
 4. Submit the form to create and update the mailbox in Exchange Online

Notes shown in the form:
- Retrieving and validating data typically takes ~10 seconds

## Getting started
### Requirements

#### App Registration & Certificate Setup

Before implementing this connector, make sure to configure a Microsoft Entra ID, an App Registration. During the setup process, you’ll create a new App Registration in the Entra portal, assign the necessary API permissions (such as user and group read/write), and generate and assign a certificate.

Follow the official Microsoft documentation for creating an App Registration and setting up certificate-based authentication:
- [App-only authentication with certificate (Exchange Online)](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication)

#### HelloID-specific configuration

Once you have completed the Microsoft setup and followed their best practices, configure the following HelloID-specific requirements.

- **API Permissions** (Application permissions):
  - `User.ReadWrite.All`
  - `Group.ReadWrite.All`
  - `GroupMember.ReadWrite.All`
  - `UserAuthenticationMethod.ReadWrite.All`
  - `User.EnableDisableAccount.All`
  - `User-PasswordProfile.ReadWrite.All`
  - `User-Phone.ReadWrite.All`
  - **Entra ID Role assignment:**
  - Assign the **Exchange Recipient Administrator** role to the App Registration
- **Certificate:**
  - Upload the public key file (.cer) in Entra ID
  - Provide the certificate as a Base64 string in HelloID. For instructions on creating the certificate and obtaining the base64 string, refer to our forum post: [Setting up a certificate for Microsoft Graph API in HelloID connectors](https://forum.helloid.com/forum/helloid-provisioning/5338-instruction-setting-up-a-certificate-for-microsoft-graph-api-in-helloid-connectors#post5338)

### Connection settings

The following user-defined variables are used by the connector.

| Setting     | Description                              | Mandatory |
| ----------- | ---------------------------------------- | --------- |
| EntraTenantId | Entra tenant ID                       | Yes       |
| EntraAppId    | Entra application (client) ID         | Yes       |
| EntraCertificateBase64String | Entra Certificate string      | Yes       |
| EntraCertificatePassword | Entra Certificate password      | Yes       |

## Remarks

- **Validation data source**:
  - The form includes a validation field that checks mailbox availability using Exchange Online cmdlets.
  - It verifies `Alias`, `Name`, and `PrimarySmtpAddress` using `Get-Mailbox`.
- **Alias normalization**:
  - The alias and primary SMTP address are derived from the alias with spaces removed.
- **Mailbox creation and update**:
  - After `New-Mailbox` completes, the script waits briefly and then applies mailbox settings via `Set-Mailbox` (e.g., message copy settings for Send As/Send on Behalf scenarios).
- **Performance notes**:
  - Validation/creation typically completes in ~10 seconds; actual times may vary.
- **Duplicate import**:
  - When importing a duplicate form, resource names can be suffixed automatically, as configured in the script.

## Development resources

### API endpoints

This connector uses Exchange Online PowerShell (EXO) cmdlets via the `ExchangeOnlineManagement` module:

| Cmdlet/Operation         | Description                                  |
|--------------------------|----------------------------------------------|
| Connect-ExchangeOnline   | Establish EXO session using app-only auth    |
| Get-Mailbox              | Check mailbox existence/availability         |
| New-Mailbox              | Create shared mailbox                        |
| Set-Mailbox              | Update mailbox settings after creation       |
| Disconnect-ExchangeOnline| Close EXO session                            |

### API documentation

- Exchange Online PowerShell overview: https://learn.microsoft.com/powershell/exchange/exchange-online-powershell
- Connect-ExchangeOnline: https://learn.microsoft.com/powershell/module/exchange/connect-exchangeonline
- Get-Mailbox: https://learn.microsoft.com/powershell/module/exchange/get-mailbox
- New-Mailbox: https://learn.microsoft.com/powershell/module/exchange/new-mailbox
- Set-Mailbox: https://learn.microsoft.com/powershell/module/exchange/set-mailbox
- Disconnect-ExchangeOnline: https://learn.microsoft.com/powershell/module/exchange/disconnect-exchangeonline

## Getting help
> :bulb: **Tip:**  
> For more information on Delegated Forms, please refer to our documentation pages: https://docs.helloid.com/en/service-automation/delegated-forms.html

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/