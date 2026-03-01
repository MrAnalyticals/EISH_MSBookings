<#
.SYNOPSIS
    Test connectivity to Microsoft Bookings via Microsoft Graph API

.DESCRIPTION
    This script authenticates to Microsoft Graph and tests connectivity to your Microsoft Bookings instance.
    It lists booking businesses, retrieves details, and lists appointments.

.NOTES
    Prerequisites:
    - Microsoft.Graph PowerShell module
    - Appropriate permissions (Bookings.Read.All or higher)
    - Microsoft 365 Business Premium license

.EXAMPLE
    .\Test-MSBookingsConnection.ps1
#>

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Bookings

# Import required modules (install if needed)
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Bookings')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module: $module" -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $module -ErrorAction Stop
}

Write-Host "`n=== Microsoft Bookings Connectivity Test ===" -ForegroundColor Cyan
Write-Host "This script will test your connection to Microsoft Bookings`n" -ForegroundColor Cyan

# Connect to Microsoft Graph with required scopes
Write-Host "[1/5] Authenticating to Microsoft Graph..." -ForegroundColor Green
try {
    # Request least privilege permissions needed for reading bookings
    Connect-MgGraph -Scopes "Bookings.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "✓ Authentication successful`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Get the current user context
Write-Host "[2/5] Getting current user context..." -ForegroundColor Green
try {
    $context = Get-MgContext
    Write-Host "✓ Signed in as: $($context.Account)" -ForegroundColor Green
    Write-Host "  Tenant ID: $($context.TenantId)" -ForegroundColor Gray
    Write-Host "  Scopes: $($context.Scopes -join ', ')`n" -ForegroundColor Gray
} catch {
    Write-Host "✗ Failed to get context: $($_.Exception.Message)" -ForegroundColor Red
}

# List all booking businesses in the tenant
Write-Host "[3/5] Listing all booking businesses in tenant..." -ForegroundColor Green
try {
    $businesses = Get-MgBookingBusiness -ErrorAction Stop
    
    if ($businesses.Count -eq 0) {
        Write-Host "✗ No booking businesses found in this tenant" -ForegroundColor Yellow
        Write-Host "  Make sure you have:" -ForegroundColor Yellow
        Write-Host "  • Microsoft 365 Business Premium license" -ForegroundColor Yellow
        Write-Host "  • At least one Bookings calendar created" -ForegroundColor Yellow
    } else {
        Write-Host "✓ Found $($businesses.Count) booking business(es):`n" -ForegroundColor Green
        
        foreach ($business in $businesses) {
            Write-Host "  • Display Name: $($business.DisplayName)" -ForegroundColor Cyan
            Write-Host "    ID: $($business.Id)" -ForegroundColor Gray
            Write-Host "    Email: $($business.Email)" -ForegroundColor Gray
            Write-Host ""
        }
    }
} catch {
    Write-Host "✗ Failed to list businesses: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Error Details: $($_.Exception.Response)" -ForegroundColor Red
}

# Get details about specific business (if any found)
if ($businesses.Count -gt 0) {
    Write-Host "[4/5] Getting details for first business..." -ForegroundColor Green
    
    $businessId = $businesses[0].Id
    
    try {
        $businessDetails = Get-MgBookingBusiness -BookingBusinessId $businessId -ErrorAction Stop
        
        Write-Host "✓ Business Details:" -ForegroundColor Green
        Write-Host "  Display Name: $($businessDetails.DisplayName)" -ForegroundColor Cyan
        Write-Host "  Email: $($businessDetails.Email)" -ForegroundColor Gray
        Write-Host "  Phone: $($businessDetails.Phone)" -ForegroundColor Gray
        Write-Host "  Address: $($businessDetails.Address.Street), $($businessDetails.Address.City)" -ForegroundColor Gray
        Write-Host "  Business Hours: $($businessDetails.BusinessHours.Count) schedules defined" -ForegroundColor Gray
        Write-Host ""
    } catch {
        Write-Host "✗ Failed to get business details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # List appointments
    Write-Host "[5/5] Listing appointments..." -ForegroundColor Green
    try {
        $appointments = Get-MgBookingBusinessAppointment -BookingBusinessId $businessId -ErrorAction Stop
        
        if ($appointments.Count -eq 0) {
            Write-Host "  No appointments found" -ForegroundColor Yellow
        } else {
            Write-Host "✓ Found $($appointments.Count) appointment(s):" -ForegroundColor Green
            
            foreach ($apt in $appointments | Select-Object -First 5) {
                Write-Host "`n  Appointment ID: $($apt.Id)" -ForegroundColor Cyan
                Write-Host "  Customer: $($apt.CustomerName)" -ForegroundColor Gray
                Write-Host "  Start: $($apt.StartDateTime.DateTime)" -ForegroundColor Gray
                Write-Host "  End: $($apt.EndDateTime.DateTime)" -ForegroundColor Gray
                Write-Host "  Duration: $($apt.Duration)" -ForegroundColor Gray
            }
            
            if ($appointments.Count -gt 5) {
                Write-Host "`n  ... and $($appointments.Count - 5) more" -ForegroundColor Gray
            }
        }
        Write-Host ""
    } catch {
        Write-Host "✗ Failed to list appointments: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Try to get services
    Write-Host "[BONUS] Listing services..." -ForegroundColor Green
    try {
        $services = Get-MgBookingBusinessService -BookingBusinessId $businessId -ErrorAction Stop
        
        if ($services.Count -eq 0) {
            Write-Host "  No services found" -ForegroundColor Yellow
        } else {
            Write-Host "✓ Found $($services.Count) service(s):" -ForegroundColor Green
            foreach ($service in $services) {
                Write-Host "  • $($service.DisplayName) - $($service.DefaultDuration)" -ForegroundColor Cyan
            }
        }
        Write-Host ""
    } catch {
        Write-Host "✗ Failed to list services: $($_.Exception.Message)" -ForegroundColor Red
    }
    
} else {
    Write-Host "[4/5] Skipped - No businesses found" -ForegroundColor Yellow
    Write-Host "[5/5] Skipped - No businesses found`n" -ForegroundColor Yellow
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan
Write-Host "If you saw errors above, check:" -ForegroundColor Yellow
Write-Host "  1. You have the required Microsoft 365 license" -ForegroundColor Yellow
Write-Host "  2. You have Bookings enabled in your tenant" -ForegroundColor Yellow
Write-Host "  3. You have the appropriate Graph API permissions" -ForegroundColor Yellow
Write-Host "  4. Admin consent has been granted for the app" -ForegroundColor Yellow

# Disconnect
Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Gray
