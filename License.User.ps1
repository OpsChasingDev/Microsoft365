<#
- takes a user to copy licensing from and applies it to a desired, unlicensed user
- this effectively gives the unlicensed user the same license assignments as the specified example user
#>

# take user input
$TargetUser = Read-Host "Enter the user (ObjectID) who you want to apply licensing to"
$ExampleUser = Read-Host "Enter the example user (ObjectID) whose licensing mirrors what you want applied to the new user"

# make sure both specified accounts exist
try { $TargetUserStored = Get-AzureADUser -ObjectId $TargetUser }
catch {
    Write-Error "Could not find a user account with an ObjectID of $TargetUser.  The script has stopped."
    break
}

try { $ExampleUserStored = Get-AzureADUser -ObjectId $ExampleUser }
catch {
    Write-Error "Could not find a user account with an ObjectID of $ExampleUser.  The script has stopped."
    break
}

# verification to proceed
$Confirm = Read-Host "`nYou are about to assign licensing to user $($TargetUserStored.DisplayName) which mirrors that of user $($ExampleUserStored.DisplayName).`n
Are you sure you want to proceed? (y/n)"
if ($Confirm -eq 'y' -or $Confirm -eq 'yes') {
    Write-Output "Assigning licensing"
}
else {
    Write-Output "No changes have been made."
    break
}

# creating licensing
$ExampleLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$ExampleLicense.SkuId = $ExampleUserStored.AssignedLicenses.SkuId
$TargetLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$TargetLicense.AddLicenses = $ExampleLicense

# assign the licensing
Set-AzureADUserLicense -ObjectId $TargetUserStored.ObjectID -AssignedLicenses $TargetLicense -ErrorAction Stop
Write-Output "Licensing assigned successfully to $($TargetUserStored.DisplayName)."