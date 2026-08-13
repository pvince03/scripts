Add-Computer -DomainName olympus.gaia.kosmos -OUPath "OU=App,OU=Member Servers,DC=olympus,DC=gaia,DC=kosmos" -Credential olympus\picklerick -Restart -Force

Add-Computer -DomainName hylandgov.local -OUPath "OU=App,OU=Member Servers,DC=hylandgov,DC=local" -Credential hylandgov\patrick.vince -Restart -Force

Add-Computer -DomainName perceptivecloud.com -OUPath "OU=Infrastructure,DC=perceptivecloud,DC=com" -Credential perceptivecloud\scaryterry -Restart -Force