# Quick summary: count staff members per appointment
$TenantId = "YOUR_TENANT_ID_HERE"
$ClientId = "YOUR_CLIENT_ID_HERE"
$ClientSecret = "YOUR_CLIENT_SECRET_HERE"

$tokenBody = @{
    client_id     = $ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody
$token = $tokenResponse.access_token

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" -Headers $headers -Method GET

$totalAppointments = 0
$foundStaffField = $false
$staffCounts = @{}

Write-Host "Checking $($response.value.Count) businesses for staff data...`n" -ForegroundColor Cyan

foreach ($business in $response.value) {
    try {
        $apptResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$($business.id)/appointments" -Headers $headers -Method GET
        
        $count = $apptResponse.value.Count
        $totalAppointments += $count
        
        if ($count -gt 0) {
            if ($apptResponse.value[0].PSObject.Properties.Name -contains 'staffMemberIds') {
                $foundStaffField = $true
            }
            
            foreach ($appt in $apptResponse.value) {
                $staffCount = if ($appt.staffMemberIds) { $appt.staffMemberIds.Count } else { 0 }
                
                if ($staffCounts.ContainsKey($staffCount)) {
                    $staffCounts[$staffCount]++
                } else {
                    $staffCounts[$staffCount] = 1
                }
            }
        }
    } catch {
        # Skip errors
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RESULTS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total appointments: $totalAppointments" -ForegroundColor White
Write-Host "staffMemberIds field exists: $foundStaffField" -ForegroundColor White
Write-Host ""
Write-Host "Staff members per appointment:" -ForegroundColor Yellow
$staffCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
    Write-Host "  $($_.Name) staff: $($_.Value) appointments" -ForegroundColor Gray
}
Write-Host ""

$multiStaff = ($staffCounts.GetEnumerator() | Where-Object { $_.Name -gt 1 } | Measure-Object -Property Value -Sum).Sum
if ($null -eq $multiStaff) { $multiStaff = 0 }

Write-Host "========================================" -ForegroundColor Cyan
if ($multiStaff -eq 0) {
    Write-Host "CONCLUSION: SAFE TO CONCATENATE" -ForegroundColor Green
    Write-Host "All appointments have 0 or 1 staff member." -ForegroundColor Green
    Write-Host "You can use a single text field for StaffId." -ForegroundColor Gray
} else {
    Write-Host "WARNING: Multiple staff found" -ForegroundColor Yellow
    Write-Host "$multiStaff appointments have 2+ staff members" -ForegroundColor Yellow
}
Write-Host ""
