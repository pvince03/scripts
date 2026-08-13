# Assuming you're working with Active Directory users
# If not, replace "Get-ADUser" with the appropriate cmdlet for your environment

# Replace 'User1' and 'User2' with the actual usernames you want to retrieve information for
$user1 = Get-ADUser -Identity 'User1' -Properties *
$user2 = Get-ADUser -Identity 'User2' -Properties *

# Display all attributes for User1
Write-Host "Attributes for User1:"
$user1 | Format-List | Out-String

# Display all attributes for User2
Write-Host "Attributes for User2:"
$user2 | Format-List | Out-String




# Define usernames
$user1 = "username1"
$user2 = "username2"

# Get group memberships
$groups1 = Get-ADUser $user1 -Property MemberOf | Select-Object -ExpandProperty MemberOf
$groups2 = Get-ADUser $user2 -Property MemberOf | Select-Object -ExpandProperty MemberOf

# Convert Distinguished Names to readable group names
$groupNames1 = $groups1 | ForEach-Object { ($_ -split ",")[0] -replace "CN=","" }
$groupNames2 = $groups2 | ForEach-Object { ($_ -split ",")[0] -replace "CN=","" }

# Compare and show groups in User1 that aren't in User2
$missingGroups = Compare-Object -ReferenceObject $groupNames1 -DifferenceObject $groupNames2 |
    Where-Object { $_.SideIndicator -eq "<=" } |
    Select-Object -ExpandProperty InputObject

# Output the result
$missingGroups
