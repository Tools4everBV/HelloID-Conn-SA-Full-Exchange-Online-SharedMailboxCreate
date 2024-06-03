<!-- Requirements -->
## Requirements
This HelloID Service Automation Delegated Form uses the [Exchange Online PowerShell V3 module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps). A HelloID agent is required to import the Exchange Online module.

<!-- Description -->
## Description
This HelloID Service Automation Delegated Form provides Exchange Online (Office365) Shared mailbox functionality. The following steps will be performed:
 1. Give a name for a new shared mailbox to create (check if unique in Exchange Online is included)
 2. Create the shared mailbox

## Versioning
| Version | Description                                             | Date       |
| ------- | ------------------------------------------------------- | ---------- |
| 2.0.0   | Rework to new logging and use exchange online module V3 | 2021/11/16 |
| 1.0.1   | Added version number and updated all-in-one script      | 2021/11/16 |
| 1.0.0   | Initial release                                         | 2021/04/29 |

<!-- TABLE OF CONTENTS -->
## Table of Contents
- [Requirements](#requirements)
- [Description](#description)
- [Versioning](#versioning)
- [Table of Contents](#table-of-contents)
- [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  - [Getting started](#getting-started)
- [Post-setup configuration (Creating the Azure AD App Registration and certificate)](#post-setup-configuration-creating-the-azure-ad-app-registration-and-certificate)
  - [Application Registration](#application-registration)
  - [Configuring App Permissions](#configuring-app-permissions)
  - [Assign Azure AD roles to the application](#assign-azure-ad-roles-to-the-application)
- [Manual resources](#manual-resources)
  - [Powershell data source 'Shared-mailbox-generate-table-mail-domains-create'](#powershell-data-source-shared-mailbox-generate-table-mail-domains-create)
  - [Delegated form task 'Shared-mailbox-create'](#delegated-form-task-shared-mailbox-create)
- [Getting help](#getting-help)
- [HelloID Docs](#helloid-docs)


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

 _Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_


### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.

 
## Post-setup configuration (Creating the Azure AD App Registration and certificate)
> _The steps below are based on the [Microsoft documentation](https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps) as of the moment of release. The Microsoft documentation should always be leading and is susceptible to change. The steps below might not reflect those changes._
> >**Please note that our steps differ from the current documentation as we use Access Token Based Authentication instead of Certificate Based Authentication**

### Application Registration
The first step is to register a new **Azure Active Directory Application**. The application is used to connect to Exchange and to manage permissions.

* Navigate to **App Registrations** in Azure, and select “New Registration” (**Azure Portal > Azure Active Directory > App Registration > New Application Registration**).
* Next, give the application a name. In this example we are using “**ExO PowerShell CBA**” as application name.
* Specify who can use this application (**Accounts in this organizational directory only**).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “**Register**” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

* To assign your application the right permissions, navigate to **Azure Portal > Azure Active Directory > App Registrations**.
* Select the application we created before, and select “**API Permissions**” or “**View API Permissions**”.
* To assign a new permission to your application, click the “**Add a permission**” button.
* From the “**Request API Permissions**” screen click “**Office 365 Exchange Online**”.
  > _The Office 365 Exchange Online might not be a selectable API. In thise case, select "APIs my organization uses" and search here for "Office 365 Exchange Online"__
* For this connector the following permissions are used as **Application permissions**:
  *	Manage Exchange As Application ***Exchange.ManageAsApp***
* To grant admin consent to our application press the “**Grant admin consent for TENANT**” button.

### Assign Azure AD roles to the application
Azure AD has more than 50 admin roles available. The **Exchange Administrator** role should provide the required permissions for any task in Exchange Online PowerShell. However, some actions may not be allowed, such as managing other admin accounts, for this the Global Administrator would be required. and Exchange Administrator roles. Please note that the required role may vary based on your configuration.
* To assign the role(s) to your application, navigate to **Azure Portal > Azure Active Directory > Roles and administrators**.
* On the Roles and administrators page that opens, find and select one of the supported roles e.g. “**Exchange Administrator**” by clicking on the name of the role (not the check box) in the results.
* On the Assignments page that opens, click the “**Add assignments**” button.
* In the Add assignments flyout that opens, **find and select the app that we created before**.
* When you're finished, click **Add**.
* Back on the Assignments page, **verify that the app has been assigned to the role**.

For more information about the permissions, please see the Microsoft docs:
* [Permissions in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/permissions-exo).
* [Find the permissions required to run any Exchange cmdlet](https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps).
* [View and assign administrator roles in Azure Active Directory](https://learn.microsoft.com/en-us/powershell/exchange/find-exchange-cmdlet-permissions?view=exchange-ps).


## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source 'Shared-mailbox-generate-table-mail-domains-create'
This Static data source the domain name for the mail address of the mailbox.

### Delegated form task 'Shared-mailbox-create'
This delegated form task will create the shared mailbox in Exchange.

## Getting help
_If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/service-automation/95-helloid-sa-exchange-online-create-shared-mailbox)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/