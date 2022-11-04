# Microsoft 365
Scripts for managing assets and operations in Microsoft 365

## Standards
- Functions should be named using the verb-noun nomenclature
- The noun portion of a function's name should be prefixed with 'M365'
  - e.g. Verb-M365Noun

## Connecting to Microsoft 365 with PowerShell
To my colleagues using the tooling in this repo:

A number of changes have been made over time to the way Microsoft has given users an API into managing their 365 platform products with PowerShell.  Among the different options you may find on the internet, the methodology going forward is going to be one which uses HTTPS and supports both basic and MFA authentication types; you'll know it's the 'right' method if the authentication pops up a web login window.  Multiple modules exists that utilize this capability, and I have listed some of the most common and applicable to what you do below:
### Azure AD
```
Connect-AzureAD
```
This cmdlet exists in the module "AzureAD" and can be obtained using the below:
```
Install-Module AzureAD -Force
Import-Module AzureAD -Force
```
Use this for general management of users, groups, permissions, and more in Azure Active Directory.
### Exchange Online
```
Connect-ExchangeOnline
```
This cmdlet exists in the module "ExchangeOnlineManagement" and can be obtained using the below:
```
Install-Module ExchangeOnlineManagement -Force
Import-Module ExchangeOnlineManagement -Force
```
Use this for anything and everything Exchange; think of this as your Exchange Management Shell for Microsoft 365 mailboxes.