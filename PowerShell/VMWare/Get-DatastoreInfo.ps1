#|#####|#########################|###############################|###############################|###############################|#######################|
#| A01 | Gather Datastore Info   | Authored : [Your Name]        | ***** PRODUCTION-STABLE ***** |  Date Released : [Date]       |   Date Revised : [Date]       |
#region|-----------------------------------------------------------------------------------------|-------------------------------|###############################|<
#|###############################################################################################|###############################|#
#| A02 | SUMMARY                 | Priority : Very High          | Date Done Goal : [Date]       | #vSphere, #Datastore, #Report |
#region------------------------------------------------------------------------------------------|###############################|<
#|^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^|A
#| -A- | Purpose                 | Assigner : [Name]             | Date Revised   : [Date]       |
#region----------------------------------------------------------|###############################|<
#\
#) - Connect to all vCenter servers and gather datastore information
#) - Identify datastores with less than 500GB of free space
#) - Generate a comprehensive report of low-space datastores
#/
#endregion-------------------------------------------------------|###############################|>
#|^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^|B
#| -B- | Objectives              | Assignee : [Your Name]        | Date Revised   : [Date]       |
#region----------------------------------------------------------|###############################|<
#\
#) 1) : Connect to all applicable vCenter servers
#) 2) : Retrieve datastore information (name, capacity, free space, percent used)
#) 3) : Identify datastores with less than 500GB free space
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

# Datastore free space threshold (in GB)
$FreeSpaceThresholdGB = 200

# Report output paths
$ReportPath = 'C:\Reports\Datastore Reports'
$ReportName = 'Datastore_LowSpace_' + (Get-Date -Format 'yyyyMMddTHHmmssZ')

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
$AllDatastoreInfo = @()
$LowSpaceDatastores = @()

# Start transcript for logging
$TranscriptPath = Join-Path $ReportPath "Datastore_Info_$(Get-Date -Format 'yyyyMMddTHHmmssZ').txt"
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
    
    # Retrieve datastore information
    try {
        Write-Host 'Retrieving datastore information...' -ForegroundColor 'Yellow' -NoNewline
        
        $Datastores = Get-Datastore -ErrorAction 'Stop' | Where-Object {$_.State -eq 'Available'}
        Write-Host " DONE (Found $($Datastores.Count) datastores)" -BackgroundColor 'Black' -ForegroundColor 'Green'
        
        # Process each datastore
        foreach ($Datastore in $Datastores) {
            $CapacityGB = [Math]::Round(($Datastore.CapacityGB), 2)
            $FreeSpaceGB = [Math]::Round(($Datastore.FreeSpaceGB), 2)
            $UsedSpaceGB = [Math]::Round(($CapacityGB - $FreeSpaceGB), 2)
            $PercentUsed = [Math]::Round((($CapacityGB - $FreeSpaceGB) / $CapacityGB * 100), 2)
            
            $DatastoreObject = [PSCustomObject]@{
                'Data Center'      = $DataCenter
                'vCenter Server'   = $vCenterServer
                'Datastore Name'   = $Datastore.Name
                'Capacity (GB)'    = $CapacityGB
                'Free Space (GB)'  = $FreeSpaceGB
                'Used Space (GB)'  = $UsedSpaceGB
                'Percent Used (%)'  = $PercentUsed
                'Status'           = if ($FreeSpaceGB -lt $FreeSpaceThresholdGB) {'WARNING - Low Space'} else {'OK'}
                'VM Count'         = @($Datastore | Get-VM -ErrorAction 'SilentlyContinue').Count
            }
            
            # Add to all datastore array
            $AllDatastoreInfo += $DatastoreObject
            
            # Add to low space array if applicable
            if ($FreeSpaceGB -lt $FreeSpaceThresholdGB) {
                $LowSpaceDatastores += $DatastoreObject
                Write-Host "  ⚠ ALERT: $($Datastore.Name) - $FreeSpaceGB GB free (below $FreeSpaceThresholdGB GB threshold)" -ForegroundColor 'Red'
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
Write-Host 'DATASTORE SUMMARY' -ForegroundColor 'Cyan'
Write-Host '════════════════════════════════════════════════════════════════════' -ForegroundColor 'Cyan'

Write-Host "Total Datastores Scanned: $($AllDatastoreInfo.Count)"
Write-Host "Datastores with <$FreeSpaceThresholdGB GB Free: $($LowSpaceDatastores.Count)" -ForegroundColor 'Red'
Write-Host ''

if ($LowSpaceDatastores.Count -gt 0) {
    Write-Host 'DATASTORES WITH LOW FREE SPACE:' -ForegroundColor 'Red'
    Write-Host '════════════════════════════════════════════════════════════════════' -ForegroundColor 'Red'
    $LowSpaceDatastores | Format-Table -AutoSize
}
else {
    Write-Host 'No datastores with low free space detected.' -ForegroundColor 'Green'
}

# Export to CSV
if ($AllDatastoreInfo.Count -gt 0) {
    $CSVPath = Join-Path $ReportPath "$ReportName.csv"
    try {
        Write-Host ''
        Write-Host "Exporting all datastores to CSV: $CSVPath" -ForegroundColor 'Yellow' -NoNewline
        $AllDatastoreInfo | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding 'UTF8' -Force
        Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
    }
    catch {
        Write-Host ' FAILED' -BackgroundColor 'Black' -ForegroundColor 'Red'
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor 'Red'
    }
    
    # Export low-space datastores to separate CSV
    $LowSpaceCSVPath = Join-Path $ReportPath "$ReportName`_LowSpace.csv"
    try {
        Write-Host "Exporting low-space datastores to CSV: $LowSpaceCSVPath" -ForegroundColor 'Yellow' -NoNewline
        $LowSpaceDatastores | Export-Csv -Path $LowSpaceCSVPath -NoTypeInformation -Encoding 'UTF8' -Force
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
                
                Write-Host "Exporting all datastores to Excel: $ExcelPath" -ForegroundColor 'Yellow' -NoNewline
                $AllDatastoreInfo | Export-Excel -WorksheetName 'All Datastores' @ExcelParams
                Write-Host ' DONE' -BackgroundColor 'Black' -ForegroundColor 'Green'
                
                Write-Host "Adding low-space datastores sheet..." -ForegroundColor 'Yellow' -NoNewline
                $LowSpaceDatastores | Export-Excel -Path $ExcelPath -WorksheetName 'Low Space Alert' -AutoSize -FreezeTopRow -Append
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
