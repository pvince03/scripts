$Servers = @(
    "Server1","Server2","Server3"
)

$RegPath = "HKLM:\SOFTWARE\HylandCloudOps\Maintenance"
$Name = "MaintenancePool"
$Value = "Replace with the Maintenance Pool name"

Invoke-Command -ComputerName $Servers -ScriptBlock {
    param($RegPath, $Name, $Value)

    if (-not (Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }

    New-ItemProperty -Path $RegPath -Name $Name -Value $Value -PropertyType String -Force | Out-Null
} -ArgumentList $RegPath, $Name, $Value