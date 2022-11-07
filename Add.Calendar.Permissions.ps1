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
        Write-Verbose "Found mailbox $Identity"
    }
    catch {
        Write-Verbose "Cound not find mailbox for $Identity"
    }

    # checking if $User[] exist(s)
    $InvalidUser = @()
    foreach ($u in $User) {
        try {
            Get-Mailbox -Identity $u -ErrorAction Stop
            Write-Verbose "Found mailbox $u"
        }
        catch {
            Write-Verbose "Could not find mailbox for $u"
            $InvalidUser += $u
        }
    }
    if ($InvalidUser.Count -gt 0) {
        Write-Output "INVALID USERS DETECTED:"`n $InvalidUser.Split(' ')
        Write-Error "The mailbox for at least one specified value of -User was not found.  No changes have been made."
        break
    }
    
    if ($Identity) { Write-Verbose "Identity specified as $Identity" }
    if ($Permission) { Write-Verbose "Permission specified as $Permission" }
    if ($User) { Write-Verbose "Working in the parameter set ParamSet_User" }
    if ($AllUser) { Write-Verbose "Working in the parameter set ParamSet_AllUser" }
}