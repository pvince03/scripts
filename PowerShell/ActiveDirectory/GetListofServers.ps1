Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' -Properties Name,Operatingsystem |
Sort-Object -Property Operatingsystem |
Select-Object -Property Name,Operatingsystem
