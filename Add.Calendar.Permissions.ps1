<#
function {
    param (
        [string]TargetUser
        [string]Permission
        [switch]AllExisting
        [string[]]OtherUser
        [switch]GrantTarget
        [switch]GrantOther
    )
}

- set up connection to 365 exchange using the module ExchangeOnlineManagement and the Connect-ExchangeOnline cmdlet

- user is prompted to choose one of the below scenarios:
    - Identity's calendar access is being granted to all users
    - Identity's calendar access is being granted to specified users
    - All users' calendar access is being granted to Identity
    - Specified users' calendar access is being granted to Identity
#>

