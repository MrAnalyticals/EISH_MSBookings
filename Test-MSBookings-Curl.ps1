# Microsoft Bookings API - Quick curl Test (PowerShell)
# This script uses curl.exe to test the Microsoft Graph Bookings API

Write-Host "`n=== Microsoft Bookings API Test (using curl) ===" -ForegroundColor Cyan
Write-Host "This script will guide you through testing your Bookings instance with curl`n" -ForegroundColor Gray

# Step 1: Get an access token
Write-Host "[Step 1] Getting an access token..." -ForegroundColor Green
Write-Host "We'll use Azure CLI to get a token. Make sure you have Azure CLI installed." -ForegroundColor Gray
Write-Host ""

try {
    # Try to get token using Azure CLI
    $token = az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv 2>$null
    
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrEmpty($token)) {
        Write-Host "Azure CLI not available or not logged in." -ForegroundColor Yellow
        Write-Host "Please run: az login" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Alternatively, paste your access token manually:" -ForegroundColor Yellow
        $token = Read-Host "Access Token"
    } else {
        Write-Host "✓ Got access token using Azure CLI" -ForegroundColor Green
        Write-Host "  Token preview: $($token.Substring(0, 20))..." -ForegroundColor Gray
    }
} catch {
    Write-Host "Error getting token: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please paste your access token manually:" -ForegroundColor Yellow
    $token = Read-Host "Access Token"
}

if ([string]::IsNullOrEmpty($token)) {
    Write-Host "`n✗ No token provided. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: List all booking businesses
Write-Host "[Step 2] Listing all booking businesses..." -ForegroundColor Green
Write-Host "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" -ForegroundColor Gray
Write-Host ""

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
}

try {
    $response = curl.exe -s -X GET `
        "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" `
        -H "Authorization: Bearer $token" `
        -H "Content-Type: application/json"
    
    Write-Host "Response:" -ForegroundColor Cyan
    $response | ConvertFrom-Json | ConvertTo-Json -Depth 10
    
    # Parse response to get business IDs
    $businesses = ($response | ConvertFrom-Json).value
    
    if ($businesses.Count -gt 0) {
        Write-Host "`n✓ Found $($businesses.Count) business(es)" -ForegroundColor Green
        
        # Show list
        for ($i = 0; $i -lt $businesses.Count; $i++) {
            Write-Host "  [$i] $($businesses[$i].displayName) - $($businesses[$i].id)" -ForegroundColor Cyan
        }
        
        # Ask user which business to query
        Write-Host ""
        $selection = Read-Host "Enter number to query details (or press Enter to skip)"
        
        if (![string]::IsNullOrEmpty($selection) -and $selection -match '^\d+$') {
            $businessId = $businesses[[int]$selection].id
            
            # Step 3: Get business details
            Write-Host "`n[Step 3] Getting details for: $businessId" -ForegroundColor Green
            Write-Host "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId" -ForegroundColor Gray
            Write-Host ""
            
            $detailsResponse = curl.exe -s -X GET `
                "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId" `
                -H "Authorization: Bearer $token" `
                -H "Content-Type: application/json"
            
            Write-Host "Response:" -ForegroundColor Cyan
            $detailsResponse | ConvertFrom-Json | ConvertTo-Json -Depth 10
            
            # Step 4: List appointments
            Write-Host "`n[Step 4] Listing appointments..." -ForegroundColor Green
            Write-Host "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId/appointments" -ForegroundColor Gray
            Write-Host ""
            
            $appointmentsResponse = curl.exe -s -X GET `
                "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId/appointments" `
                -H "Authorization: Bearer $token" `
                -H "Content-Type: application/json"
            
            Write-Host "Response:" -ForegroundColor Cyan
            $appointmentsResponse | ConvertFrom-Json | ConvertTo-Json -Depth 10
            
            # Step 5: List services
            Write-Host "`n[Step 5] Listing services..." -ForegroundColor Green
            Write-Host "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId/services" -ForegroundColor Gray
            Write-Host ""
            
            $servicesResponse = curl.exe -s -X GET `
                "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId/services" `
                -H "Authorization: Bearer $token" `
                -H "Content-Type: application/json"
            
            Write-Host "Response:" -ForegroundColor Cyan
            $servicesResponse | ConvertFrom-Json | ConvertTo-Json -Depth 10
            
            # Step 6: List staff
            Write-Host "`n[Step 6] Listing staff members..." -ForegroundColor Green
            Write-Host "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId/staffMembers" -ForegroundColor Gray
            Write-Host ""
            
            $staffResponse = curl.exe -s -X GET `
                "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId/staffMembers" `
                -H "Authorization: Bearer $token" `
                -H "Content-Type: application/json"
            
            Write-Host "Response:" -ForegroundColor Cyan
            $staffResponse | ConvertFrom-Json | ConvertTo-Json -Depth 10
        }
    } else {
        Write-Host "✗ No businesses found in tenant" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Tip: You can copy these curl commands and paste them directly in your terminal" -ForegroundColor Yellow
Write-Host "Just replace the token with your actual access token." -ForegroundColor Yellow
Write-Host ""

# Show example curl command
Write-Host "Example curl command:" -ForegroundColor Cyan
Write-Host @"
curl -X GET \
  'https://graph.microsoft.com/v1.0/solutions/bookingBusinesses' \
  -H 'Authorization: Bearer YOUR_TOKEN_HERE' \
  -H 'Content-Type: application/json'
"@ -ForegroundColor Gray
