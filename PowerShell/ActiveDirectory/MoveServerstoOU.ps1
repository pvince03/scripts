$computers = Get-Content -Path "C:\Path\To\computers.txt"
$targetOU = "OU=Computers,OU=MyBranch,DC=domain,DC=com"

foreach ($comp in $computers) {
    $compDN = (Get-ADComputer -Identity $comp).DistinguishedName
    Move-ADObject -Identity $compDN -TargetPath $targetOU
}