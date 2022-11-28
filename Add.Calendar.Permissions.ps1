<#
- set up connection to 365 exchange using the module ExchangeOnlineManagement and the Connect-ExchangeOnline cmdlet

Next Steps:
- add ability to specify an exception along with AllUsers

#>

function Add-M365CalendarPermission {
    [Cmdletbinding(DefaultParameterSetName = "default",
        SupportsShouldProcess,
        ConfirmImpact = "Medium")]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Owner", "PublishingEditor", "Reviewer")]
        [string]$Permission,

        [Parameter(ParameterSetName = "ParamSet_User")]
        [string[]]$User,
        
        [Parameter(ParameterSetName = "ParamSet_AllUser")]
        [switch]$AllUser,

        [Parameter(ParameterSetName = "ParamSet_AllUser")]
        [string]$ExcludeUser
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
        Write-Output "`nINVALID USERS DETECTED:" $InvalidUser.Split(' ')
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
            if ($PSCmdlet.ShouldProcess("$Identity", "Granting $User access to calendar with permission $Permission...")) {
                Write-Verbose "Operation for granting specified users access to calendar of $Identity"
                foreach ($u in $User) {
                    Add-MailboxFolderPermission -Identity ($Identity + ":\calendar") -User $u -AccessRights $Permission
                }
            }
        }
        elseif ($Selection_User -eq '2') {
            Write-Verbose "Operation for granting $Identity access to calendar of specified users"
            foreach ($u in $User) {
                if ($PSCmdlet.ShouldProcess("$u", "Granting $Identity access to calendar with permission $Permission...")) {
                    Add-MailboxFolderPermission -Identity ($u + ":\calendar") -User $Identity -AccessRights $Permission
                }
            }
        }
        else {
            Write-Warning "Not a valid selection.  No changes have been made."
            break
        }
    }

    # ParamSet_AllUser logic - deals with the Identity and specified user(s)
    if ($AllUser) {
        Write-Verbose "Working in the parameter set ParamSet_AllUser"

        $Mailbox = Get-Mailbox | Where-Object {$_.WindowsEmailAddress -notlike "DiscoverySearchMailbox*" -and $_.WindowsEmailAddress -ne $Identity -and $_.WindowsEmailAddress -notcontains $ExcludeUser}
        Write-Host `n"The below mailboxes will be targeted for the operation:" -ForegroundColor 'Yellow' -BackgroundColor 'Black'
        Write-Output $Mailbox.WindowsEmailAddress | Sort-Object $_

        Write-Host `n"Choose one of the below two scenarios:" -ForegroundColor 'Yellow' -BackgroundColor 'Black'
        Write-Host "  (1) Identity's calendar access is being granted to all users."
        Write-Host "  (2) All users' calendar access is being granted to Identity."
        $Selection_AllUser = Read-Host "Selection"
        Write-Verbose "Selected option $Selection_AllUser"


        if ($Selection_AllUser -eq '1') {
            if ($PSCmdlet.ShouldProcess("$Identity", "Granting the users listed above access to calendar with permission $Permission...")) {
                Write-Verbose "Operation for granting all listed users access to calendar of $Identity"
                foreach ($m in $Mailbox) {
                    Add-MailboxFolderPermission -Identity ($Identity + ":\calendar") -User $m.WindowsEmailAddress -AccessRights $Permission
                }
            }
        }
        elseif ($Selection_AllUser -eq '2') {
            if ($PSCmdlet.ShouldProcess("ALL Users", "Granting $Identity access to calendar with permissions $Permission...")) {
                Write-Verbose "Operation for granting $Identity access to calendar of all listed users"
                foreach ($m in $Mailbox) {
                    Add-MailboxFolderPermission -Identity ($m.WindowsEmailAddress + ":\calendar") -User $Identity -AccessRights $Permission
                }
            }
        }
        else {
            Write-Warning "Not a valid selection.  No changes have been made."
            break
        }
    }
}