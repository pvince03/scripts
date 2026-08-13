Import-Module ActiveDirectory -ErrorAction Stop
$src = 'asur'
$tgt = 'babygroot'
# Direct (non-nested) security groups of source
$srcGroups = (Get-ADUser $src -Properties memberOf).memberOf |
  ForEach-Object { Get-ADGroup $_ -Properties GroupCategory, SID, DistinguishedName, Name } |
  Where-Object { $_ -and $_.GroupCategory -eq 'Security' }
# Filter out Domain Users (RID 513)
try {
    $domainSid = (Get-ADDomain).DomainSID.Value
    $primaryGroupSid = "$domainSid-513"
    $srcGroups = $srcGroups | Where-Object { $_.SID.Value -ne $primaryGroupSid }
} catch {}
# Target current group DNs (to avoid duplicates)
$tgtDns = (Get-ADUser $tgt -Properties memberOf).memberOf
# Compute groups to add
$toAdd = $srcGroups | Where-Object { $tgtDns -notcontains $_.DistinguishedName }
if (-not $toAdd -or $toAdd.Count -eq 0) {
    Write-Host "No new security groups to add; target already matches source (direct memberships)." -ForegroundColor Yellow
} else {
    Write-Host "Adding $($toAdd.Count) group(s) to '$tgt':" -ForegroundColor Cyan
    $toAdd | Select-Object Name, DistinguishedName | Format-Table -AutoSize
    foreach ($g in $toAdd) {
        try {
            Add-ADGroupMember -Identity $g.DistinguishedName -Members $tgt -Confirm:$false -ErrorAction Stop
            Write-Host "Added: $($g.Name)" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to add '$tgt' to '$($g.Name)': $($_.Exception.Message)"
        }
    }
}