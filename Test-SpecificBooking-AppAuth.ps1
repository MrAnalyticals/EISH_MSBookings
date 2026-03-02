<#
.SYNOPSIS
    Test Microsoft Bookings using App Registration (Service Principal)

.DESCRIPTION
    This script connects to Microsoft Bookings using app-only authentication
    and tests access to your specific booking business.
    
.NOTES
    Requires .env file with CLIENT_ID, CLIENT_SECRET, and TENANT_ID
#>

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Bookings

Write-Host "`n=== Testing Microsoft Bookings with App Authentication ===" -ForegroundColor Cyan

# Load environment variables from .env file
$envFile = Join-Path $PSScriptRoot ".env"
if (Test-Path $envFile) {
    Write-Host "Loading credentials from .env file..." -ForegroundColor Green
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]*)\s*=\s*(.*)$') {
            $name = $matches[1].Trim()
            $value = $matches[2].Trim()
            Set-Variable -Name $name -Value $value -Scope Script
            Write-Host "  [OK] Loaded: $name" -ForegroundColor Gray
        }
    }
} else {
    Write-Host "ERROR: .env file not found!" -ForegroundColor Red
    Write-Host "Please create a .env file with CLIENT_ID, CLIENT_SECRET, and TENANT_ID" -ForegroundColor Yellow
    exit 1
}

# Verify required variables
if (!$CLIENT_ID -or !$CLIENT_SECRET -or !$TENANT_ID) {
    Write-Host "ERROR: Missing required credentials in .env file" -ForegroundColor Red
    Write-Host "Required: CLIENT_ID, CLIENT_SECRET, TENANT_ID" -ForegroundColor Yellow
    exit 1
}

if (!$BOOKING_BUSINESS_ID) {
    $BOOKING_BUSINESS_ID = "EnergyReportCallWithElectricIrelandSuperhomes@electricirelandsuperhomes.ie"
}

Write-Host "`nConfiguration:" -ForegroundColor Cyan
Write-Host "  Client ID: $CLIENT_ID" -ForegroundColor Gray
Write-Host "  Tenant ID: $TENANT_ID" -ForegroundColor Gray
Write-Host "  Booking Business: $BOOKING_BUSINESS_ID" -ForegroundColor Gray

# Authenticate using app credentials (client credentials flow)
Write-Host "`nAuthenticating with Microsoft Graph..." -ForegroundColor Green
try {
    # Convert client secret to secure string
    $secureSecret = ConvertTo-SecureString $CLIENT_SECRET -AsPlainText -Force
    
    # Create credential object
    $credential = New-Object System.Management.Automation.PSCredential($CLIENT_ID, $secureSecret)
    
    # Connect to Microsoft Graph with app-only auth
    Connect-MgGraph -TenantId $TENANT_ID -ClientSecretCredential $credential -NoWelcome
    
    Write-Host "[OK] Successfully authenticated!" -ForegroundColor Green
    
    # Verify connection
    $context = Get-MgContext
    Write-Host "`nConnection Details:" -ForegroundColor Cyan
    Write-Host "  App Name: $($context.AppName)" -ForegroundColor Gray
    Write-Host "  Auth Type: $($context.AuthType)" -ForegroundColor Gray
    Write-Host "  Scopes: $($context.Scopes -join ', ')" -ForegroundColor Gray
}
catch {
    Write-Host "[ERROR] Authentication failed!" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 1: List all booking businesses
Write-Host "`n=== Test 1: List All Booking Businesses ===" -ForegroundColor Cyan
try {
    $businesses = Get-MgBookingBusiness -ErrorAction Stop
    
    if ($businesses.Count -eq 0) {
        Write-Host "  No booking businesses found" -ForegroundColor Yellow
        Write-Host "  This might indicate a permissions issue." -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($businesses.Count) booking business(es):" -ForegroundColor Green
        $businesses | ForEach-Object {
            Write-Host "    • $($_.DisplayName)" -ForegroundColor White
            Write-Host "      ID: $($_.Id)" -ForegroundColor Gray
            Write-Host "      Email: $($_.Email)" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "  [ERROR] Failed to list businesses" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: Get specific booking business
Write-Host "`n=== Test 2: Get Specific Booking Business ===" -ForegroundColor Cyan
Write-Host "Attempting to access: $BOOKING_BUSINESS_ID" -ForegroundColor Yellow
try {
    $business = Get-MgBookingBusiness -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    Write-Host "`n[SUCCESS] Connected to your booking business!" -ForegroundColor Green
    Write-Host "`nBusiness Details:" -ForegroundColor Cyan
    Write-Host "  Display Name: $($business.DisplayName)" -ForegroundColor White
    Write-Host "  Email: $($business.Email)" -ForegroundColor Gray
    Write-Host "  Phone: $($business.Phone)" -ForegroundColor Gray
    Write-Host "  Website: $($business.WebSiteUrl)" -ForegroundColor Gray
    Write-Host "  Is Published: $($business.IsPublished)" -ForegroundColor Gray
    Write-Host "  Public URL: $($business.PublicUrl)" -ForegroundColor Gray
    Write-Host "  Business Type: $($business.BusinessType)" -ForegroundColor Gray
}
catch {
    Write-Host "  [ERROR] Failed to access specific business" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`n  Possible reasons:" -ForegroundColor Yellow
    Write-Host "    1. The app doesn't have Bookings.Read.All permissions" -ForegroundColor Yellow
    Write-Host "    2. Admin consent hasn't been granted" -ForegroundColor Yellow
    Write-Host "    3. The booking business ID is incorrect" -ForegroundColor Yellow
    
    # Disconnect and exit
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Test 3: Get appointments
Write-Host "`n=== Test 3: Get Appointments ===" -ForegroundColor Cyan
try {
    $appointments = Get-MgBookingBusinessAppointment -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    if ($appointments.Count -eq 0) {
        Write-Host "  No appointments scheduled yet" -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($appointments.Count) appointment(s):" -ForegroundColor Green
        $appointments | Select-Object -First 5 | ForEach-Object {
            Write-Host "`n  Appointment:" -ForegroundColor White
            Write-Host "    ID: $($_.Id)" -ForegroundColor Gray
            Write-Host "    Customer: $($_.CustomerName)" -ForegroundColor Gray
            Write-Host "    Email: $($_.CustomerEmailAddress)" -ForegroundColor Gray
            Write-Host "    Start: $($_.Start.DateTime)" -ForegroundColor Gray
            Write-Host "    End: $($_.End.DateTime)" -ForegroundColor Gray
            Write-Host "    Service: $($_.ServiceName)" -ForegroundColor Gray
            Write-Host "    Status: $($_.OptOutOfCustomerEmail)" -ForegroundColor Gray
        }
        
        if ($appointments.Count -gt 5) {
            Write-Host "`n  ... and $($appointments.Count - 5) more" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "  [ERROR] Failed to get appointments" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 4: Get services
Write-Host "`n=== Test 4: Get Services ===" -ForegroundColor Cyan
try {
    $services = Get-MgBookingBusinessService -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    if ($services.Count -eq 0) {
        Write-Host "  No services configured" -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($services.Count) service(s):" -ForegroundColor Green
        $services | ForEach-Object {
            Write-Host "`n  Service:" -ForegroundColor White
            Write-Host "    Name: $($_.DisplayName)" -ForegroundColor Gray
            Write-Host "    ID: $($_.Id)" -ForegroundColor Gray
            Write-Host "    Duration: $($_.DefaultDuration)" -ForegroundColor Gray
            Write-Host "    Price: $($_.DefaultPrice) $($_.DefaultPriceType)" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "  [ERROR] Failed to get services" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 5: Get staff members
Write-Host "`n=== Test 5: Get Staff Members ===" -ForegroundColor Cyan
try {
    $staff = Get-MgBookingBusinessStaffMember -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    if ($staff.Count -eq 0) {
        Write-Host "  No staff members configured" -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($staff.Count) staff member(s):" -ForegroundColor Green
        $staff | ForEach-Object {
            Write-Host "    • $($_.DisplayName) - $($_.EmailAddress)" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "  [ERROR] Failed to get staff members" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Cleanup
Write-Host "`n=== Complete ===" -ForegroundColor Green
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Gray
Disconnect-MgGraph | Out-Null
Write-Host "[OK] Done!" -ForegroundColor Green
