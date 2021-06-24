<!-- Description -->
## Description
This HelloID Service Automation Delegated Form provides Exchange Online (Office365) Shared mailbox functionality. The following steps will be performed:
 1. Give a name for a new shared mailbox to create
 2. Create the shared mailbox
 3. A corresponding AD group is created 
 4. The AD group is granted Full Access and Send As permissions on the mailbox
 
<!-- TABLE OF CONTENTS -->
## Table of Contents
* [Description](#description)
* [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  * [Getting started](#getting-started)
* [Post-setup configuration](#post-setup-configuration)
* [Manual resources](#manual-resources)


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

 _Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_


### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.

 
## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)
<table>
  <tr><td><strong>Variable name</strong></td><td><strong>Example value</strong></td><td><strong>Description</strong></td></tr>
  <tr><td>ExchangeOnlineAdminUsername</td><td>user@domain.com</td><td>Exchange admin account</td></tr>
  <tr><td>ExchangeOnlineAdminPassword</td><td>********</td><td>Exchange admin password</td></tr>
</table>

## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source 'Shared-mailbox-generate-table-mail-domains-create'
This Static data source the domain name for the mail address of the mailbox.

### Delegated form task 'Shared-mailbox-create'
This delegated form task will create the shared mailbox in Exchange.

# HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/
