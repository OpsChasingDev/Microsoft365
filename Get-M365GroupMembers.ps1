# Connect to Microsoft 365
Connect-MsolService

# Get all Microsoft 365 groups
$groups = Get-MsolGroup -All

# Initialize an array to hold the results
$results = @()

# Loop through each group and get its members
foreach ($group in $groups) {
    $groupMembers = Get-MsolGroupMember -GroupObjectId $group.ObjectId -All

    # Add the group and its members to the results array
    foreach ($groupMember in $groupMembers) { 
        $results += [PSCustomObject] @{
            "Group Name"   = $group.DisplayName
            "Group Email"  = $group.EmailAddress
            "Member Name"  = $groupMember.DisplayName
            "Member Email" = $groupMember.EmailAddress 
        }    
    }
}

# Export the results to a CSV file
$results | Export-Csv -Path "C:\group_members.csv" -NoTypeInformation