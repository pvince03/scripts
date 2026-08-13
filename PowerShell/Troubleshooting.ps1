Set-PowerCLIConfiguration -InvalidCertificateAction ignore -confirm:$false

Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false  -confirm:$false