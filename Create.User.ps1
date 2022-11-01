<#
- creates a new user in Microsoft 365 (AzureAD)
#>
<#
Install-Module AzureAD
Import-Module AzureAD
Connect-AzureAD
#>

# prompt user input
$FirstName = Read-Host "Enter a first name"
$LastName = Read-Host "Enter a last name"
$Password = Read-Host "Enter an initial password" -AsSecureString
$Domain = Read-Host "Enter the domain suffix for the user's UPN"

# formulating values
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $Password
$DisplayName = $FirstName + " " + $LastName
$UPN = $FirstName + "." + $LastName + "@" + $Domain
$MailNickName = $FirstName + $LastName

# user creation
$NewUserSplat = @{
    AccountEnabled = $true
    DisplayName = $DisplayName
    GivenName = $FirstName
    Surname = $LastName
    UserPrincipalName = $UPN
    PasswordProfile = $PasswordProfile
    MailNickName = $MailNickName
    UsageLocation = 'US'
}
New-AzureADUser @NewUserSplat