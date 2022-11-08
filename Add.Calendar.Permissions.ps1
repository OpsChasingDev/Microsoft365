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
        Write-Error "The mailbox for $Identity was not found.  No changes haves been made."
        break
    }

    # checking if $User[] exist(s)
    $InvalidUser = @()
    foreach ($u in $User) {
        try {
            $null = Get-Mailbox -Identity $u -ErrorAction Stop
            Write-Verbose "Found mailbox $u"
        }
        catch {
            Write-Verbose "Could not find mailbox for $u"
            $InvalidUser += $u
        }
    }
    if ($InvalidUser.Count -gt 0) {
        Write-Output "INVALID USERS DETECTED:" $InvalidUser.Split(' ')
        Write-Error "The mailbox for at least one specified value of -User was not found.  No changes have been made."
        break
    }
    
    # ParamSet_User logic - deals with the Identity and specified user(s)
    if ($User) {
        Write-Verbose "Working in the parameter set ParamSet_User"
        Write-Host `n"Choose one of the below two scenarios:" -ForegroundColor 'Yellow' -BackgroundColor 'Black'
        Write-Host "  (1) Identity's calendar access is being granted to specified users."
        Write-Host "  (2) Specified users' calendar access is being granted to Identity."
        $Selection_User = Read-Host "Selection"
        Write-Verbose "Selected option $Selection_User"

        if ($Selection_User -eq '1') {
            Write-Verbose "Operation for granting specified users access to calendar of $Identity"
        }
        elseif ($Selection_User -eq '2') {
            Write-Verbose "Operation for granting $Identity access to calendar of specified users"
        }
        else {
            Write-Warning "Not a valid selection.  No changes have been made."
            break
        }
    }

    # *** MIRROR BLOCK ABOVE HERE FOR ParamSet_AllUser LOGIC ***


    if ($AllUser) { Write-Verbose "Working in the parameter set ParamSet_AllUser" }
}