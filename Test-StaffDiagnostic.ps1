# Simple diagnostic - check raw API response
$TenantId = "YOUR_TENANT_ID_HERE"
$ClientId = "YOUR_CLIENT_ID_HERE"
$ClientSecret = "YOUR_CLIENT_SECRET_HERE"

Write-Host "Getting token..." -ForegroundColor Cyan

$tokenBody = @{
    client_id     = $ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody
$token = $tokenResponse.access_token

Write-Host "Token obtained" -ForegroundColor Green
Write-Host "Calling API..." -ForegroundColor Cyan

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

try {
    $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" -Headers $headers -Method GET
    
    Write-Host "`nFull Response:" -ForegroundColor Yellow
    $response | ConvertTo-Json -Depth 10
    
    Write-Host "`nBusinesses Count: $($response.value.Count)" -ForegroundColor Cyan
    
    if ($response.value.Count -gt 0) {
        $businessId = $response.value[0].id
        Write-Host "`nGetting appointments for first business: $businessId" -ForegroundColor Cyan
        
        $apptResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$businessId/appointments" -Headers $headers -Method GET
        
        Write-Host "`nAppointments Response:" -ForegroundColor Yellow
        $apptResponse | ConvertTo-Json -Depth 10
        
        if ($apptResponse.value.Count -gt 0) {
            Write-Host "`nFirst Appointment Properties:" -ForegroundColor Yellow
            $apptResponse.value[0].PSObject.Properties | ForEach-Object {
                Write-Host "  $($_.Name): $($_.Value)" -ForegroundColor Gray
            }
            
            if ($apptResponse.value[0].PSObject.Properties.Name -contains 'staffMemberIds') {
                Write-Host "`nstaffMemberIds field found!" -ForegroundColor Green
                Write-Host "  Value: $($apptResponse.value[0].staffMemberIds)" -ForegroundColor Gray
                Write-Host "  Count: $($apptResponse.value[0].staffMemberIds.Count)" -ForegroundColor Gray
            } else {
                Write-Host "`nstaffMemberIds field NOT found in response" -ForegroundColor Red
            }
        }
    }
} catch {
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception -ForegroundColor Red
}
