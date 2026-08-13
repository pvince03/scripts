# Get the list of server names
$servers = Get-Content -Path "C:\software\servers.txt"

# Loop through each server
foreach ($server in $servers) {
    Write-Host "Processing $server..." -ForegroundColor Cyan

    try {
        Invoke-Command -ComputerName $server -ScriptBlock {
            $partition = Get-Partition -DriveLetter C
            $size = (Get-PartitionSupportedSize -DriveLetter C)

            if ($partition.Size -lt $size.SizeMax) {
                Resize-Partition -DriveLetter C -Size $size.SizeMax
                Write-Output "C: drive extended to $($size.SizeMax / 1GB) GB on $env:COMPUTERNAME"
            } else {
                Write-Output "C: drive on $env:COMPUTERNAME is already using all available space."
            }
        } -ErrorAction Stop
    } catch {
        Write-Warning "Failed to connect or extend C: drive on $server. Error: $_"
    }
}
