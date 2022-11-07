<#
- set up connection to 365 exchange using the module ExchangeOnlineManagement and the Connect-ExchangeOnline cmdlet

- identity is first checked to make sure it exists
- if user param is specified, the validity of those accounts is checked as well
- user is prompted to choose one of the below scenarios based on paramset:
    Paramset 1
    - (1) Identity's calendar access is being granted to specified users
    - (2) Specified users' calendar access is being granted to Identity
    Paramset 2
    - (1) Identity's calendar access is being granted to all users
    - (2) All users' calendar access is being granted to Identity
#>

function Add-M365CalendarPermission {
    [Cmdletbinding(DefaultParameterSetName = "default")]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Owner", "PublishingEditor", "Reviewer")]
        [string]$Permission,

        [Parameter(ParameterSetName = "ParamSet_User")]
        [string[]]$User,
        
        [Parameter(ParameterSetName = "ParamSet_AllUser")]
        [switch]$AllUser
    )

    # checking if $Identity exists
    try {
        $null = Get-Mailbox -Identity $Identity -ErrorAction Stop
        Write-Output "found mailbox $Identity"
    }
    catch {
        Write-Output "cound not find mailbox for $Identity"
    }
    if ($Identity) { Write-Output "Identity specified as $Identity" }
    if ($Permission) { Write-Output "Permission specified as $Permission" }
    if ($User) { Write-Output "working in the parameter set ParamSet_User" }
    if ($AllUser) { Write-Output "working in the parameter set ParamSet_AllUser" }
}