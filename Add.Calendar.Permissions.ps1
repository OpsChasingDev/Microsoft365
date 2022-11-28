<#
- set up connection to 365 exchange using the module ExchangeOnlineManagement and the Connect-ExchangeOnline cmdlet

Next Steps:
- write help file

#>

function Add-M365CalendarPermission {
    <#
    .SYNOPSIS
        Adds permissions to Microsoft 365 mailbox calendars.
    .DESCRIPTION
        Adds permissions to Microsoft 365 mailbox calendars.  A connection to a Microsoft 365 tenant must first be established before use.
        Users are identified by their full, primary SMTP address.
        Options exist to both grant users permissions to an Identity's calendar or to grant the Identity permissions to users' calendars.
        Both of the above options are included in two ParameterSets, one for dealing with selecting other users individually, and another for specifying all other users.
        Between all of the above options, a total of four scenarios exist for using this function:

            [ParamSet_User]
        1) Identity's calendar access is being granted to specified users.
        2) Specified users' calendar access is being granted to Identity.

            [ParamSet_AllUser]
        3) Identity's calendar access is being granted to all users.
        4) All users' calendar access is being granted to Identity.

        Entering the ParamSet_User options is done by specifying the -User parameter.
        Entering the ParamSet_AllUser options is done by specifying the -AllUser parameter.
    .NOTES
        Git repo: https://github.com/OpsChasingDev/Microsoft365
    .EXAMPLE
        Add-M365CalendarPermission -Identity robert.stapleton@savylabs.net -User api@savylabs.net -Permission "PublishingEditor" -WhatIf

Choose one of the below two scenarios:
  (1) Identity's calendar access is being granted to specified users.
  (2) Specified users' calendar access is being granted to Identity.
Selection: 1
What if: Performing the operation "Granting api@savylabs.net access to calendar with permission PublishingEditor..." on target "robert.stapleton@savylabs.net".
        
        The user api@savylabs.net is granted PublishingEditor permissions to the calendar of robert.stapleton@savylabs.net.  Note how option 1 was selected when prompted.
        Notice how the syntax of this command is identical to that of Example 2, but the resulting operation was flipped as a result of selecting option 1 instead of option 2.
    .EXAMPLE
        Add-M365CalendarPermission -Identity robert.stapleton@savylabs.net -User api@savylabs.net -Permission "PublishingEditor" -WhatIf

Choose one of the below two scenarios:
  (1) Identity's calendar access is being granted to specified users.
  (2) Specified users' calendar access is being granted to Identity.
Selection: 2
What if: Performing the operation "Granting robert.stapleton@savylabs.net access to calendar with permission PublishingEditor..." on target "api@savylabs.net".

        The user robert.stapleton@savylabs.net is granted PublishingEditor permissions to the calendar of api@savylabs.net.  Note how option 2 was selected when prompted.
        Notice how the syntax of this command is identical to that of Example 1, but the resulting operation was flipped as a result of selecting option 2 instead of option 1.
    .EXAMPLE
        Add-M365CalendarPermission -Identity api@savylabs.net -User robert.stapleton@savylabs.net, test.user2@savylabs.net -Permission "PublishingEditor" -WhatIf

Choose one of the below two scenarios:
  (1) Identity's calendar access is being granted to specified users.
  (2) Specified users' calendar access is being granted to Identity.
Selection: 1
What if: Performing the operation "Granting robert.stapleton@savylabs.net test.user2@savylabs.net access to calendar with permission PublishingEditor..." on target "api@savylabs.net".

        The users robert.stapleton@savylabs.net and test.user2@savylabs.net are being granted PublishingEditor permissions to the calendar of api@savylabs.net.
    .EXAMPLE
        Add-M365CalendarPermission -Identity test.user2@savylabs.net -AllUser -Permission "Owner" -WhatIf

The below mailboxes will be targeted for the operation:
API@savylabs.net
Robert.Stapleton@savylabs.net

Choose one of the below two scenarios:
  (1) Identity's calendar access is being granted to all users.
  (2) All users' calendar access is being granted to Identity.
Selection: 1
What if: Performing the operation "Granting the users listed above access to calendar with permission Owner..." on target "test.user2@savylabs.net".

        All non-system users in the Microsoft 365 tenant are granted Owner permissions to the calendar of test.user2@savylabs.net.
        Notice how a list of users collected in the -AllUser parameter is provided before making the option selection.
    .EXAMPLE
        Add-M365CalendarPermission -Identity test.user2@savylabs.net -AllUser -Permission "Owner" -ExcludeUser api@savylabs.net -WhatIf

The below mailboxes will be targeted for the operation:
Robert.Stapleton@savylabs.net

Choose one of the below two scenarios:
  (1) Identity's calendar access is being granted to all users.
  (2) All users' calendar access is being granted to Identity.
Selection: 1
What if: Performing the operation "Granting the users listed above access to calendar with permission Owner..." on target "test.user2@savylabs.net".

        Like in Example 4, all non-system users in the Microsoft 365 tenant are granted Owner permissions to the calendar of test.user2@savylabs.net.
        The difference in this example is we have specified the -ExcludeUser parameter because we wanted all users except the api@savylabs.net user to be granted permissions to this calendar.
        Notice how a list of users collected in the -AllUser parameter is provided before making the option selection.
    #>
    
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
        [string[]]$ExcludeUser
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

        $Mailbox = Get-Mailbox |
            Where-Object {
                $_.WindowsEmailAddress -notlike "DiscoverySearchMailbox*" -and
                $_.WindowsEmailAddress -ne $Identity -and
                $_.WindowsEmailAddress -notin $ExcludeUser
            }

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