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

- identity is a required param
- either user(s) or all users is a required param (two param sets)
- permission is a required param (validateset)
- identity is first checked to make sure it exists
- if user param is specified, the validity of those accounts is checked as well
- user is prompted to choose one of the below scenarios:
    - (1) Identity's calendar access is being granted to all users
    - (2) Identity's calendar access is being granted to specified users
    - (3) All users' calendar access is being granted to Identity
    - (4) Specified users' calendar access is being granted to Identity
#>

