# ==============================
# Confluence Bulk Labeling Script
# ==============================

# Confluence Configuration
$BaseUrl  = "https://hyland.atlassian.net"
$Email    = "patrick.vince@hyland.com"
$ApiToken = "INSERT-TOKEN-HERE"

# Pages to label - one Page ID per line in this file
$PageIdsFile = "$PSScriptRoot\PageIds.txt"

if (-not (Test-Path $PageIdsFile)) {
    Write-Host "Page ID file not found: $PageIdsFile" -ForegroundColor Red
    exit 1
}

$PageIds = Get-Content -Path $PageIdsFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith("#") }

if ($PageIds.Count -eq 0) {
    Write-Host "No page IDs found in $PageIdsFile" -ForegroundColor Red
    exit 1
}

# Label to apply
$LabelName = "informational"

# Optional but useful for older Windows PowerShell sessions
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Authentication Header
$AuthString  = "$Email`:$ApiToken"
$AuthBytes   = [System.Text.Encoding]::UTF8.GetBytes($AuthString)
$EncodedAuth = [Convert]::ToBase64String($AuthBytes)

$Headers = @{
    Authorization = "Basic $EncodedAuth"
    Accept        = "application/json"
}

foreach ($PageId in $PageIds) {

    $Uri = "$BaseUrl/wiki/rest/api/content/$PageId/label"

    $Body = @(
        @{
            prefix = "global"
            name   = $LabelName
        }
    ) | ConvertTo-Json -Depth 5

    try {
        Write-Host "Adding label '$LabelName' to page $PageId..." -ForegroundColor Yellow

        $Result = Invoke-RestMethod `
            -Uri $Uri `
            -Method Post `
            -Headers $Headers `
            -ContentType "application/json" `
            -Body $Body `
            -ErrorAction Stop

        Write-Host "Success: Added '$LabelName' to page $PageId" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed: Page ${PageId}" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red

        if ($_.Exception.Response -ne $null) {
            Write-Host "HTTP Status: $($_.Exception.Response.StatusCode.value__) $($_.Exception.Response.StatusDescription)" -ForegroundColor Red
        }
    }
}