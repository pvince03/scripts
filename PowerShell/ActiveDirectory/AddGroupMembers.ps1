$Users = get-content -Path "C:\Users\picklerick\Documents\add-to-workspaces.csv"

foreach($User in $Users){Add-ADGroupMember -Identity "dlg-workspaces-usr" -Members $User}