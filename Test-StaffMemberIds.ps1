# Test script to analyze staffMemberIds arrays in appointments
# Purpose: Determine if appointments can have multiple staff members

Write-Host "`n=== Analyzing Staff Member Assignments ===" -ForegroundColor Cyan
Write-Host "Checking if appointments have multiple staff members assigned`n" -ForegroundColor Gray

# Configuration from your M file
$TenantId = "YOUR_TENANT_ID_HERE"
$ClientId = "YOUR_CLIENT_ID_HERE"
$ClientSecret = "YOUR_CLIENT_SECRET_HERE"

# Step 1: Get Access Token
Write-Host "Step 1: Getting OAuth access token..." -ForegroundColor Green

$tokenBody = @{
    client_id     = $ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}

try {
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody
    $token = $tokenResponse.access_token
    Write-Host "✓ Token obtained successfully`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to get token: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 2: Get all booking businesses
Write-Host "Step 2: Fetching booking businesses..." -ForegroundColor Green

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

try {
    $businessesResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" -Headers $headers -Method GET
    $businesses = $businessesResponse.value
    Write-Host "✓ Found $($businesses.Count) business(es)`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to get businesses: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 3: Analyze appointments from all businesses
Write-Host "Step 3: Analyzing appointments for staff assignments..." -ForegroundColor Green
Write-Host "=" * 80 -ForegroundColor Gray

$totalAppointments = 0
$appointmentsWithStaff = 0
$appointmentsWithMultipleStaff = 0
$maxStaffCount = 0
$staffCountDistribution = @{}

foreach ($business in $businesses) {
    Write-Host "`nBusiness: $($business.displayName)" -ForegroundColor Cyan
    Write-Host "ID: $($business.id)" -ForegroundColor Gray
    
    try {
        $appointmentsUrl = "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$($business.id)/appointments"
        $appointmentsResponse = Invoke-RestMethod -Uri $appointmentsUrl -Headers $headers -Method GET
        $appointments = $appointmentsResponse.value
        
        Write-Host "Appointments found: $($appointments.Count)" -ForegroundColor Gray
        
        foreach ($appointment in $appointments) {
            $totalAppointments++
            
            # Check if staffMemberIds exists and has data
            if ($appointment.PSObject.Properties.Name -contains 'staffMemberIds') {
                $staffIds = $appointment.staffMemberIds
                $staffCount = if ($staffIds) { $staffIds.Count } else { 0 }
                
                if ($staffCount -gt 0) {
                    $appointmentsWithStaff++
                    
                    # Track distribution
                    if ($staffCountDistribution.ContainsKey($staffCount)) {
                        $staffCountDistribution[$staffCount]++
                    } else {
                        $staffCountDistribution[$staffCount] = 1
                    }
                    
                    if ($staffCount -gt 1) {
                        $appointmentsWithMultipleStaff++
                        Write-Host "  ⚠ Appointment $($appointment.id): $staffCount staff members" -ForegroundColor Yellow
                        Write-Host "    Staff IDs: $($staffIds -join ', ')" -ForegroundColor Gray
                    }
                    
                    if ($staffCount -gt $maxStaffCount) {
                        $maxStaffCount = $staffCount
                    }
                }
            }
        }
    } catch {
        Write-Host "  ✗ Error fetching appointments: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Step 4: Display results
Write-Host "`n" 
Write-Host "=" * 80 -ForegroundColor Gray
Write-Host "`n=== ANALYSIS RESULTS ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Total Appointments: $totalAppointments" -ForegroundColor White
Write-Host "Appointments with Staff Assigned: $appointmentsWithStaff" -ForegroundColor White
Write-Host "Appointments with Multiple Staff: $appointmentsWithMultipleStaff" -ForegroundColor $(if ($appointmentsWithMultipleStaff -gt 0) { "Yellow" } else { "Green" })
Write-Host "Maximum Staff per Appointment: $maxStaffCount" -ForegroundColor White

Write-Host "`nStaff Count Distribution:" -ForegroundColor Cyan
if ($staffCountDistribution.Count -gt 0) {
    $staffCountDistribution.GetEnumerator() | Sort-Object Name | ForEach-Object {
        $percentage = [math]::Round(($_.Value / $appointmentsWithStaff) * 100, 1)
        Write-Host "  $($_.Name) staff member$(if($_.Name -gt 1){'s'}): $($_.Value) appointments ($percentage%)" -ForegroundColor Gray
    }
} else {
    Write-Host "  No staff assignments found in appointments" -ForegroundColor Gray
}

# Step 5: Recommendation
Write-Host "`n=== RECOMMENDATION ===" -ForegroundColor Cyan
if ($appointmentsWithMultipleStaff -eq 0 -and $appointmentsWithStaff -gt 0) {
    Write-Host "✓ SAFE TO CONCATENATE: All appointments have 0 or 1 staff member" -ForegroundColor Green
    Write-Host "  You can safely use a single StaffId text field" -ForegroundColor Gray
} elseif ($appointmentsWithMultipleStaff -gt 0) {
    Write-Host "⚠ MULTIPLE STAFF DETECTED: $appointmentsWithMultipleStaff appointments have multiple staff" -ForegroundColor Yellow
    Write-Host "  Options:" -ForegroundColor Gray
    Write-Host "    1. Concatenate with delimiter" -ForegroundColor Gray
    Write-Host "    2. Create bridge table (expand to multiple rows)" -ForegroundColor Gray
    Write-Host "    3. Take first staff member only" -ForegroundColor Gray
} else {
    Write-Host "ℹ No staff assignments found in any appointments" -ForegroundColor Yellow
    Write-Host "  The staffMemberIds field may be optional or not populated" -ForegroundColor Gray
}

Write-Host ""
