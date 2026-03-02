<#
.SYNOPSIS
    Discover your Azure AD Tenant ID

.DESCRIPTION
    This script helps you discover your Tenant ID using your app credentials
    Credentials should be stored in .env file or passed as parameters
#>

param(
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$Domain = "electricirelandsuperhomes.ie"
)

# Load credentials from .env file if not provided as parameters
if ([string]::IsNullOrEmpty($ClientId) -or [string]::IsNullOrEmpty($ClientSecret)) {
    $envFile = Join-Path $PSScriptRoot ".env"
    if (Test-Path $envFile) {
        Write-Host "Loading credentials from .env file..." -ForegroundColor Gray
        Get-Content $envFile | ForEach-Object {
            if ($_ -match '^\s*CLIENT_ID\s*=\s*(.*)$') {
                $ClientId = $matches[1].Trim()
            }
            if ($_ -match '^\s*CLIENT_SECRET\s*=\s*(.*)$') {
                $ClientSecret = $matches[1].Trim()
            }
        }
    }
}

Write-Host "`n=== Discovering Azure AD Tenant ID ===" -ForegroundColor Cyan

# Method 1: Try to get from well-known endpoint using the domain
Write-Host "`nMethod 1: Using OpenID Connect Discovery..." -ForegroundColor Green
Write-Host "Using domain: $Domain" -ForegroundColor Gray

try {
    $response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Domain/v2.0/.well-known/openid-configuration" -Method Get
    $tenantId = $response.issuer -replace 'https://login.microsoftonline.com/', '' -replace '/v2.0', ''
    
    Write-Host "[SUCCESS] Tenant ID: $tenantId" -ForegroundColor Green
    Write-Host "`nAdd this to your .env file:" -ForegroundColor Yellow
    Write-Host "TENANT_ID=$tenantId" -ForegroundColor Cyan
    
    # Update .env file automatically
    $envPath = Join-Path $PSScriptRoot ".env"
    if (Test-Path $envPath) {
        $envContent = Get-Content $envPath -Raw
        $envContent = $envContent -replace 'TENANT_ID=YOUR_TENANT_ID_HERE', "TENANT_ID=$tenantId"
        Set-Content -Path $envPath -Value $envContent -NoNewline
        Write-Host "`n✓ Updated .env file with Tenant ID" -ForegroundColor Green
    }
    
    return $tenantId
}
catch {
    Write-Host "[FAILED] Could not discover via domain" -ForegroundColor Red
}

# Method 2: Using Azure AD authorization endpoint
Write-Host "`nMethod 2: Interactive Discovery..." -ForegroundColor Green
Write-Host "This will open a browser window for authentication." -ForegroundColor Yellow
Write-Host "Please sign in with your Electric Ireland Superhomes account." -ForegroundColor Yellow

try {
    # Install required module if not present
    if (!(Get-Module -ListAvailable -Name MSAL.PS)) {
        Write-Host "Installing MSAL.PS module..." -ForegroundColor Yellow
        Install-Module -Name MSAL.PS -Scope CurrentUser -Force
    }
    
    Import-Module MSAL.PS
    
    $token = Get-MsalToken -ClientId $ClientId -TenantId "organizations" -Interactive
    
    $tenantId = $token.TenantId
    Write-Host "[SUCCESS] Tenant ID: $tenantId" -ForegroundColor Green
    
    # Update .env file
    $envPath = Join-Path $PSScriptRoot ".env"
    if (Test-Path $envPath) {
        $envContent = Get-Content $envPath -Raw
        $envContent = $envContent -replace 'TENANT_ID=YOUR_TENANT_ID_HERE', "TENANT_ID=$tenantId"
        Set-Content -Path $envPath -Value $envContent -NoNewline
        Write-Host "`n✓ Updated .env file with Tenant ID" -ForegroundColor Green
    }
}
catch {
    Write-Host "[FAILED] $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Alternative Methods ===" -ForegroundColor Cyan
Write-Host "If the above methods failed, you can find your Tenant ID by:" -ForegroundColor Yellow
Write-Host "1. Go to: https://portal.azure.com" -ForegroundColor White
Write-Host "2. Search for 'Microsoft Entra ID' or 'Azure Active Directory'" -ForegroundColor White
Write-Host "3. In the Overview page, copy the 'Tenant ID'" -ForegroundColor White
Write-Host "4. Update the .env file with: TENANT_ID=<your-tenant-id>" -ForegroundColor White
