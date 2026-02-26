# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com), and this project adheres to [Semantic Versioning](https://semver.org).

## [3.0.0.0] - 2025-02-26

### Added
- Certificate-based authentication support for Microsoft Entra ID and Exchange Online
- Email address and alias validation using Microsoft Graph API for improved validation performance
- Dynamic domain retrieval via Microsoft Graph API to query verified domains with Email support
- Display Name field to specify the shared mailbox display name
- Email address (prefix) field separate from mail domain selection
- Optional Alias field for shared mailbox with default fallback to email prefix
- Message copy settings configuration (MessageCopyForSendOnBehalfEnabled and MessageCopyForSentAsEnabled) for shared mailbox delegation scenarios
- Two separate validation data sources for email and alias uniqueness checks
- Comprehensive README with detailed setup instructions, form workflow, and API documentation

### Changed
- **BREAKING**: Migrated authentication from secret-based to certificate-based authentication
  - Old variables: `EntraSecret`, `EntraTenantId`, `EntraAppID`, `EntraOrganization`
  - New variables: `EntraIdCertificateBase64String`, `EntraIdCertificatePassword`, `EntraIdOrganization`, `EntraIdAppId`
- **Performance**: Changed mailbox validation from `Get-Mailbox` cmdlet (30+ seconds) to Microsoft Graph API for email address and alias uniqueness checks
- **Performance**: Replaced static domain list with dynamic Graph API queries for better performance and accuracy
- Form structure completely redesigned with improved field layout and real-time validation
- Naming convention updated to use hyphens (e.g., "Exchange online - Shared Mailbox - Create")
- Data sources refactored:
  - `EntraID-Check-EmailAddress-Unique` - New Graph API-based email validation with current mailbox detection
  - `EntraID-Check-Alias-Unique` - New Graph API-based alias validation with current mailbox detection
  - `Graph-Get-MailDomains` - Dynamic Graph API-based domain retrieval (filters verified domains with Email support)
- Domain selection now automatically prioritizes the current mailbox's domain at the top of the list
- Mailbox configuration logic improved:
  - Display Name and Mailbox Name now set to user-provided value
  - Primary SMTP Address constructed dynamically from email prefix + selected domain
  - Alias passed to `New-Mailbox` if provided, otherwise defaults to email prefix
- Error handling and logging enhanced with better exception handling and context
- README updated with complete documentation and API references

### Removed
- Legacy authentication method using client secret
- Old global variable structure
- Hardcoded mailbox naming convention

### Fixed
- Improved validation performance by switching from Exchange Online cmdlets to Graph API
- Enhanced error messages with better context for troubleshooting

## [2.0.0.0] - 2024-03-06

### Changed
- Rework to new logging structure
- Migrated to Exchange Online PowerShell module V3

## [1.0.1.0] - 2021-11-16

### Added
- Added version number
- Updated all-in-one setup script

## [1.0.0.0] - 2021-04-29

### Added
- Initial release of HelloID-Conn-SA-Full-Exchange-Online-SharedMailboxCreate
- Shared mailbox creation functionality in Exchange Online
- Form-based mailbox creation workflow
- Mailbox configuration and settings management
