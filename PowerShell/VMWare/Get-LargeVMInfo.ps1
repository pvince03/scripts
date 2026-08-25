#|#####|#########################|###############################|###############################|###############################|#######################|
#| A01 | Gather Large VM Info    | Authored : [Your Name]        | ***** PRODUCTION-STABLE ***** |  Date Released : [Date]       |   Date Revised : [Date]       |
#region|-----------------------------------------------------------------------------------------|-------------------------------|###############################|<
#|###############################################################################################|###############################|#
#| A02 | SUMMARY                 | Priority : Very High          | Date Done Goal : [Date]       | #vSphere, #VM, #Report        |
#region------------------------------------------------------------------------------------------|###############################|<
#|^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^|A
#| -A- | Purpose                 | Assigner : [Name]             | Date Revised   : [Date]       |
#region----------------------------------------------------------|###############################|<
#\
#) - Connect to all vCenter servers and gather VM sizing information
#) - Identify VMs with more than 30 vCPUs, more than 30 GB of memory, or both
#) - Generate a comprehensive report of large VMs
#/
#endregion-------------------------------------------------------|###############################|>
#|^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^|B
#| -B- | Objectives              | Assignee : [Your Name]        | Date Revised   : [Date]       |
#region----------------------------------------------------------|###############################|<
#\
#) 1) : Connect to all applicable vCenter servers
#) 2) : Retrieve VM information (name, vCPU count, memory, power state)
#) 3) : Identify VMs with more than 30 vCPUs and/or more than 30 GB memory
#) 4) : Generate Excel and CSV reports with findings
#/
#endregion-------------------------------------------------------|###############################|>
#endregion---------------------------------------------------------------------------------------|###############################|>

param(
    [Parameter(Mandatory=$false, HelpMessage='Export report to Excel')]
    [switch]$OutExcel,
    
    [Parameter(Mandatory=$false, HelpMessage='Send report via email')]
    [switch]$Email,
    
    [Parameter(Mandatory=$false, HelpMessage='Enable debug output')]
    [switch]$Debugger
)

$ErrorActionPreference = 'Continue'

# ===========================================================================================
# CONFIGURATION VARIABLES
# ===========================================================================================

# Define your vCenter servers and data centers
# Format: @{ DataCenter = 'XX'; vCenterServer = 'xx-vmvcenter' }
$vCenterServers = @(
    @{ DataCenter = 'BWU'; vCenterServer = 'bwu-vmvcenter' }
    @{ DataCenter = 'PIT'; vCenterServer = 'pit-vmvcenter' }
    @{ DataCenter = 'IAD'; vCenterServer = 'iad-vmvcenter' }
    @{ DataCenter = 'LHR'; vCenterServer = 'lhr-vmvcenter' }
    @{ DataCenter = 'D01'; vCenterServer = 'd01-vmvcenter' }
    @{ DataCenter = 'D02'; vCenterServer = 'd02-vmvcenter' }
    @{ DataCenter = 'AKL'; vCenterServer = 'akl-vmvcenter' }
    @{ DataCenter = 'SYD'; vCenterServer = 'syd-vmvcenter' }
    @{ DataCenter = 'MEL'; vCenterServer = 'mel-vmvcenter2' }
    @{ DataCenter = 'FRA'; vCenterServer = 'fra-vmvcenter' }
    @{ DataCenter = 'MCI'; vCenterServer = 'mci-vmvcenter' }
    @{ DataCenter = 'RSE'; vCenterServer = 'rse-vmvcenter' }
    @{ DataCenter = 'WLG'; vCenterServer = 'wlg-vmvcenter' }
    # Add more data centers as needed:
    # @{ DataCenter = 'PIT'; vCenterServer = 'pit-vmvcenter' }
)

# Large VM thresholds
$CPUThreshold    = 30   # vCPU count
$MemoryThresholdGB = 30 # GB of memory

# Report output paths
$ReportPath = 'C:\Reports\Large VM Reports'
$ReportName = 'VM_LargeSize_' + (Get-Date -Format 'yyyyMMddTHHmmssZ')

# ===========================================================================================
# MAIN SCRIPT BLOCK
# ===========================================================================================

Write-Host ''
Write-Host 'Ready.. '        -ForegroundColor 'DarkYellow' -BackgroundColor 'Black' -NoNewline
Start-Sleep -Milliseconds 500
Write-Host 'Set.. '          -ForegroundColor 'Yellow'     -BackgroundColor 'Black' -NoNewline
Start-Sleep -Milliseconds 500
Write-Host 'Let''s GO!'       -ForegroundColor 'Green'      -BackgroundColor 'Black' -NoNewline
Start-Sleep -Milliseconds 500
Write-Host ' ~ ( ͡° ͜ʖ ͡°)' -ForegroundColor 'Cyan'
Start-Sleep -Milliseconds 500
Write-Host ''

# Initialize report array
$AllVMInfo = @()
$LargeVMs = @()

# Start transcript for logging
$TranscriptPath = Join-Path $ReportPath "VM_Info_$(Get-Date -Format 'yyyyMMddTHHmmssZ').txt"
if (!(Test-Path $ReportPath)) { New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null }

try {
    Start-Transcript -Path $TranscriptPath -Append | Out-Null
    Write-Host "Transcript started: $TranscriptPath"
}
catch {
    Write-Host "Warning: Could not start transcript - $($_.Exception.Message)" -ForegroundColor 'Yellow'
}

# ===========================================================================================
# PROCESS EACH vCENTER SERVER
# ===========================================================================================

foreach ($viServerInfo in $vCenterServers) {
    $DataCenter = $viServerInfo.DataCenter
    $vCenterServer = $viServerInfo.vCenterServer
    
    Write-Host ''
    Write-Host "Data Center: $DataCenter" -ForegroundColor 'Cyan'
    Write-Host "vCenter Server: $vCenterServer" -ForegroundColor 'Cyan'
    
    # Disconnect from any previously connected servers
    if ($global:DefaultVIServer -or $global:DefaultVIServers) {
        try {
            Write-Host 'Disconnecting from ALL VIServer(s)...' -ForegroundColor 'Yellow' -NoNewline
            [void](Disconnect-VIServer -Server '*' -Confirm:$false -Force -WarningAction 'SilentlyContinue' -ErrorAction 'Stop')
            Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
        }
        catch {
            Write-Host ' WARNING' -BackgroundColor 'Black' -ForegroundColor 'Yellow'
        }
    }
    
    # Connect to vCenter Server
    try {
        Write-Host "Connecting to $vCenterServer..." -ForegroundColor 'Yellow' -NoNewline
        $Connection = Connect-VIServer -Server $vCenterServer -WarningAction 'SilentlyContinue' -ErrorAction 'Stop'
        Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
    }
    catch {
        Write-Host ' FAILED' -BackgroundColor 'Black' -ForegroundColor 'Red'
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor 'Red'
        continue
    }
    
    # Retrieve VM information
    try {
        Write-Host 'Retrieving VM information...' -ForegroundColor 'Yellow' -NoNewline
        
        $VMs = Get-VM -ErrorAction 'Stop'
        Write-Host " DONE (Found $($VMs.Count) VMs)" -BackgroundColor 'Black' -ForegroundColor 'Green'
        
        # Process each VM
        foreach ($VM in $VMs) {
            $NumCPU = $VM.NumCpu
            $MemoryGB = [Math]::Round($VM.MemoryGB, 2)
            $IsLargeCPU = $NumCPU -gt $CPUThreshold
            $IsLargeMemory = $MemoryGB -gt $MemoryThresholdGB
            
            $ReasonFlags = @()
            if ($IsLargeCPU)    { $ReasonFlags += 'CPU' }
            if ($IsLargeMemory) { $ReasonFlags += 'Memory' }
            
            $VMObject = [PSCustomObject]@{
                'Data Center'      = $DataCenter
                'vCenter Server'   = $vCenterServer
                'VM Name'          = $VM.Name
                'Power State'      = $VM.PowerState
                'vCPU Count'       = $NumCPU
                'Memory (GB)'      = $MemoryGB
                'Exceeds CPU Threshold'    = $IsLargeCPU
                'Exceeds Memory Threshold' = $IsLargeMemory
                'Status'           = if ($ReasonFlags.Count -gt 0) {"WARNING - Large VM ($($ReasonFlags -join ' & '))"} else {'OK'}
            }
            
            # Add to all VM array
            $AllVMInfo += $VMObject
            
            # Add to large VM array if applicable
            if ($ReasonFlags.Count -gt 0) {
                $LargeVMs += $VMObject
                Write-Host "  ⚠ ALERT: $($VM.Name) - $NumCPU vCPU / $MemoryGB GB memory (exceeds $($ReasonFlags -join ' & ') threshold)" -ForegroundColor 'Red'
            }
        }
    }
    catch {
        Write-Host ' FAILED' -BackgroundColor 'Black' -ForegroundColor 'Red'
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor 'Red'
    }
}

# Disconnect from all vCenter servers
if ($global:DefaultVIServer -or $global:DefaultVIServers) {
    try {
        Write-Host ''
        Write-Host 'Disconnecting from ALL VIServer(s)...' -ForegroundColor 'Yellow' -NoNewline
        [void](Disconnect-VIServer -Server '*' -Confirm:$false -Force -WarningAction 'SilentlyContinue' -ErrorAction 'Stop')
        Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
    }
    catch {
        Write-Host ' WARNING' -BackgroundColor 'Black' -ForegroundColor 'Yellow'
    }
}

# ===========================================================================================
# GENERATE REPORTS
# ===========================================================================================

Write-Host ''
Write-Host '════════════════════════════════════════════════════════════════════' -ForegroundColor 'Cyan'
Write-Host 'LARGE VM SUMMARY' -ForegroundColor 'Cyan'
Write-Host '════════════════════════════════════════════════════════════════════' -ForegroundColor 'Cyan'

Write-Host "Total VMs Scanned: $($AllVMInfo.Count)"
Write-Host "VMs with >$CPUThreshold vCPU and/or >$MemoryThresholdGB GB Memory: $($LargeVMs.Count)" -ForegroundColor 'Red'
Write-Host ''

if ($LargeVMs.Count -gt 0) {
    Write-Host 'LARGE VMs DETECTED:' -ForegroundColor 'Red'
    Write-Host '════════════════════════════════════════════════════════════════════' -ForegroundColor 'Red'
    $LargeVMs | Format-Table -AutoSize
}
else {
    Write-Host 'No large VMs detected.' -ForegroundColor 'Green'
}

# Export to CSV
if ($AllVMInfo.Count -gt 0) {
    $CSVPath = Join-Path $ReportPath "$ReportName.csv"
    try {
        Write-Host ''
        Write-Host "Exporting all VMs to CSV: $CSVPath" -ForegroundColor 'Yellow' -NoNewline
        $AllVMInfo | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding 'UTF8' -Force
        Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
    }
    catch {
        Write-Host ' FAILED' -BackgroundColor 'Black' -ForegroundColor 'Red'
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor 'Red'
    }
    
    # Export large VMs to separate CSV
    $LargeVMsCSVPath = Join-Path $ReportPath "$ReportName`_LargeVMs.csv"
    try {
        Write-Host "Exporting large VMs to CSV: $LargeVMsCSVPath" -ForegroundColor 'Yellow' -NoNewline
        $LargeVMs | Export-Csv -Path $LargeVMsCSVPath -NoTypeInformation -Encoding 'UTF8' -Force
        Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
    }
    catch {
        Write-Host ' FAILED' -BackgroundColor 'Black' -ForegroundColor 'Red'
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor 'Red'
    }
    
    # Export to Excel if requested
    if ($OutExcel) {
        try {
            Write-Host ''
            Write-Host 'Attempting to export to Excel...' -ForegroundColor 'Yellow'
            
            # Check if ImportExcel module is available
            if (Get-Module -ListAvailable -Name 'ImportExcel') {
                Import-Module -Name 'ImportExcel' -ErrorAction 'Stop'
                
                $ExcelPath = Join-Path $ReportPath "$ReportName.xlsx"
                
                # Create Excel file with multiple sheets
                $ExcelParams = @{
                    Path              = $ExcelPath
                    NoTypeInformation = $true
                    AutoSize          = $true
                    FreezeTopRow      = $true
                    Show              = $false
                }
                
                Write-Host "Exporting all VMs to Excel: $ExcelPath" -ForegroundColor 'Yellow' -NoNewline
                $AllVMInfo | Export-Excel -WorksheetName 'All VMs' @ExcelParams
                Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
                
                Write-Host "Adding large VMs sheet..." -ForegroundColor 'Yellow' -NoNewline
                $LargeVMs | Export-Excel -Path $ExcelPath -WorksheetName 'Large VM Alert' -AutoSize -FreezeTopRow -Append
                Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
            }
            else {
                Write-Host 'ImportExcel module not found. Skipping Excel export.' -ForegroundColor 'Yellow'
            }
        }
        catch {
            Write-Host 'Excel export failed.' -ForegroundColor 'Yellow'
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor 'Yellow'
        }
    }
}

# Stop transcript
try {
    Stop-Transcript | Out-Null
}
catch {
    Write-Host 'Warning: Could not stop transcript' -ForegroundColor 'Yellow'
}

Write-Host ''
Write-Host 'Script execution completed.' -ForegroundColor 'Green'
Write-Host "Reports saved to: $ReportPath" -ForegroundColor 'Cyan'
Write-Host ''
