# Check all businesses for appointments with staff data
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

Write-Host "Token obtained`n" -ForegroundColor Green

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" -Headers $headers -Method GET

Write-Host "Found $($response.value.Count) businesses`n" -ForegroundColor Cyan
Write-Host "Checking each business for appointments...`n" -ForegroundColor Cyan

$totalAppointments = 0
$foundStaffField = $false
$maxStaffCount = 0
$multiStaffCount = 0

foreach ($business in $response.value) {
    Write-Host "Business: $($business.displayName)" -ForegroundColor Yellow
    
    try {
        $apptResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$($business.id)/appointments" -Headers $headers -Method GET
        
        $count = $apptResponse.value.Count
        Write-Host "  Appointments: $count" -ForegroundColor Gray
        
        if ($count -gt 0) {
            $totalAppointments += $count
            
            # Check first appointment for staff field
            $firstAppt = $apptResponse.value[0]
            
            Write-Host "  Fields in appointment:" -ForegroundColor Cyan
            $firstAppt.PSObject.Properties.Name | ForEach-Object {
                Write-Host "    - $_" -ForegroundColor Gray
            }
            
            if ($firstAppt.PSObject.Properties.Name -contains 'staffMemberIds') {
                $foundStaffField = $true
                Write-Host "  staffMemberIds EXISTS!" -ForegroundColor Green
                
                # Check all appointments for staff counts
                foreach ($appt in $apptResponse.value) {
                    if ($appt.staffMemberIds) {
                        $staffCount = $appt.staffMemberIds.Count
                        Write-Host "    Appointment $($appt.id): $staffCount staff member(s)" -ForegroundColor Cyan
                        
                        if ($staffCount -gt $maxStaffCount) {
                            $maxStaffCount = $staffCount
                        }
                        
                        if ($staffCount -gt 1) {
                            $multiStaffCount++
                            Write-Host "      MULTIPLE STAFF: $($appt.staffMemberIds -join ', ')" -ForegroundColor Yellow
                        }
                    }
                }
            } else {
                Write-Host "  staffMemberIds field NOT found" -ForegroundColor Red
            }
            
            Write-Host ""
        }
    } catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total appointments across all businesses: $totalAppointments" -ForegroundColor White
Write-Host "staffMemberIds field found: $foundStaffField" -ForegroundColor White
Write-Host "Maximum staff per appointment: $maxStaffCount" -ForegroundColor White
Write-Host "Appointments with multiple staff: $multiStaffCount" -ForegroundColor White

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RECOMMENDATION" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

if (!$foundStaffField) {
    Write-Host "staffMemberIds field NOT found in any appointments" -ForegroundColor Red
    Write-Host "This field may not be available or populated" -ForegroundColor Yellow
} elseif ($multiStaffCount -eq 0) {
    Write-Host "SAFE TO CONCATENATE: All appointments have 0 or 1 staff member" -ForegroundColor Green
    Write-Host "Use a single text field for StaffId" -ForegroundColor Gray
} else {
    Write-Host "WARNING: $multiStaffCount appointments have multiple staff members" -ForegroundColor Yellow
    Write-Host "Maximum staff count found: $maxStaffCount" -ForegroundColor Yellow
    Write-Host "Options:" -ForegroundColor Gray
    Write-Host "  1. Concatenate IDs with delimiter (comma/semicolon)" -ForegroundColor Gray
    Write-Host "  2. Create bridge table (one row per staff member)" -ForegroundColor Gray
    Write-Host "  3. Take only first staff member" -ForegroundColor Gray
}

Write-Host ""
