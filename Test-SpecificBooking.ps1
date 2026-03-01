<#
.SYNOPSIS
    Quick test to connect to your specific Microsoft Bookings instance

.DESCRIPTION
    This script attempts to connect to your specific booking page:
    "Energy Report Call With Electric Ireland Superhomes"

.NOTES
    Run this after running Test-MSBookingsConnection.ps1 first
#>

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Bookings

Write-Host "`n=== Testing Your Specific Booking Instance ===" -ForegroundColor Cyan
Write-Host "Target: Energy Report Call With Electric Ireland Superhomes`n" -ForegroundColor Cyan

# Connect to Microsoft Graph
Write-Host "Authenticating to Microsoft Graph..." -ForegroundColor Green
Connect-MgGraph -Scopes "Bookings.Read.All", "BookingsAppointment.ReadWrite.All" -NoWelcome

# Your booking business identifier
# Based on your URL: https://outlook.office.com/book/EnergyReportCallWithElectricIrelandSuperhomes@electricirelandsuperhomes.ie/
$bookingBusinessId = "EnergyReportCallWithElectricIrelandSuperhomes@electricirelandsuperhomes.ie"

Write-Host "`nAttempting to connect to: $bookingBusinessId" -ForegroundColor Yellow
Write-Host ""

# Method 1: Try direct ID access
Write-Host "Test 1: Direct access by ID..." -ForegroundColor Green
try {
    $business = Get-MgBookingBusiness -BookingBusinessId $bookingBusinessId -ErrorAction Stop
    
    Write-Host "[SUCCESS] Found your booking business:" -ForegroundColor Green
    Write-Host "  Display Name: $($business.DisplayName)" -ForegroundColor Cyan
    Write-Host "  Email: $($business.Email)" -ForegroundColor Gray
    Write-Host "  Phone: $($business.Phone)" -ForegroundColor Gray
    Write-Host "  Website: $($business.WebSiteUrl)" -ForegroundColor Gray
    Write-Host "  Is Published: $($business.IsPublished)" -ForegroundColor Gray
    Write-Host "  Public URL: $($business.PublicUrl)" -ForegroundColor Gray
    Write-Host ""
    
    # If successful, try to get appointments
    Write-Host "Test 2: Fetching appointments..." -ForegroundColor Green
    try {
        $appointments = Get-MgBookingBusinessAppointment -BookingBusinessId $bookingBusinessId -ErrorAction Stop
        
        if ($appointments.Count -eq 0) {
            Write-Host "  No appointments scheduled yet" -ForegroundColor Yellow
        } else {
            Write-Host "[OK] Found $($appointments.Count) appointment(s):" -ForegroundColor Green
            
            foreach ($apt in $appointments | Select-Object -First 10) {
                Write-Host "`n  Appointment:" -ForegroundColor Cyan
                Write-Host "     Customer: $($apt.CustomerName) ($($apt.CustomerEmailAddress))" -ForegroundColor Gray
                Write-Host "     Start: $($apt.StartDateTime.DateTime)" -ForegroundColor Gray
                Write-Host "     End: $($apt.EndDateTime.DateTime)" -ForegroundColor Gray
                Write-Host "     Service ID: $($apt.ServiceId)" -ForegroundColor Gray
                Write-Host "     Duration: $($apt.Duration)" -ForegroundColor Gray
                Write-Host "     Online: $($apt.IsLocationOnline)" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "[ERROR] Could not fetch appointments: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ""
    
    # Try to get services
    Write-Host "Test 3: Fetching services offered..." -ForegroundColor Green
    try {
        $services = Get-MgBookingBusinessService -BookingBusinessId $bookingBusinessId -ErrorAction Stop
        
        if ($services.Count -eq 0) {
            Write-Host "  No services configured" -ForegroundColor Yellow
        } else {
            Write-Host "[OK] Available services:" -ForegroundColor Green
            foreach ($service in $services) {
                Write-Host "`n  Service: $($service.DisplayName)" -ForegroundColor Cyan
                Write-Host "     ID: $($service.Id)" -ForegroundColor Gray
                Write-Host "     Duration: $($service.DefaultDuration)" -ForegroundColor Gray
                Write-Host "     Price: $($service.DefaultPrice) $($service.PriceType)" -ForegroundColor Gray
                Write-Host "     Description: $($service.Description)" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "[ERROR] Could not fetch services: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ""
    
    # Try to get staff
    Write-Host "Test 4: Fetching staff members..." -ForegroundColor Green
    try {
        $staff = Get-MgBookingBusinessStaffMember -BookingBusinessId $bookingBusinessId -ErrorAction Stop
        
        if ($staff.Count -eq 0) {
            Write-Host "  No staff members configured" -ForegroundColor Yellow
        } else {
            Write-Host "[OK] Staff members:" -ForegroundColor Green
            foreach ($member in $staff) {
                Write-Host "`n  Staff: $($member.DisplayName)" -ForegroundColor Cyan
                Write-Host "     Email: $($member.EmailAddress)" -ForegroundColor Gray
                Write-Host "     Role: $($member.Role)" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "[ERROR] Could not fetch staff: $($_.Exception.Message)" -ForegroundColor Red
    }
    
} catch {
    Write-Host "[FAILED] Could not access booking business" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    
    # Method 2: Search by name
    Write-Host "Test 2 Alternative: Searching all businesses for match..." -ForegroundColor Green
    try {
        $allBusinesses = Get-MgBookingBusiness
        
        $match = $allBusinesses | Where-Object { 
            $_.DisplayName -like "*Energy*Report*" -or 
            $_.Email -like "*EnergyReport*" 
        }
        
        if ($match) {
            Write-Host "[OK] Found matching business:" -ForegroundColor Green
            Write-Host "  Display Name: $($match.DisplayName)" -ForegroundColor Cyan
            Write-Host "  ID: $($match.Id)" -ForegroundColor Cyan
            Write-Host "  Email: $($match.Email)" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Try using this ID instead: $($match.Id)" -ForegroundColor Yellow
        } else {
            Write-Host "[ERROR] No matching business found" -ForegroundColor Red
            Write-Host ""
            Write-Host "Available businesses in your tenant:" -ForegroundColor Yellow
            foreach ($biz in $allBusinesses) {
                Write-Host "  - $($biz.DisplayName) ($($biz.Id))" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Host "[ERROR] Could not search businesses: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan

# Create a summary
Write-Host "`nSummary for Power Automate:" -ForegroundColor Yellow
Write-Host "If the tests above succeeded, you can use:" -ForegroundColor Yellow
Write-Host "  - Business ID: $bookingBusinessId" -ForegroundColor Cyan
Write-Host "  - Use HTTP with Azure AD connector instead of Bookings connector" -ForegroundColor Yellow
Write-Host "  - Base URL: https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" -ForegroundColor Cyan
Write-Host "  - Required Permission: Bookings.Read.All or higher" -ForegroundColor Yellow

Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Gray
