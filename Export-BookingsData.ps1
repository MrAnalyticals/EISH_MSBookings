<#
.SYNOPSIS
    Export Microsoft Bookings data to CSV files

.DESCRIPTION
    This script connects to Microsoft Bookings using app-only authentication
    and exports all data (businesses, appointments, services, staff) to CSV files.
    
.NOTES
    Requires .env file with CLIENT_ID, CLIENT_SECRET, and TENANT_ID
    Outputs CSV files to the current directory
#>

#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Bookings

Write-Host "`n=== Exporting Microsoft Bookings Data to CSV ===" -ForegroundColor Cyan

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
}
catch {
    Write-Host "[ERROR] Authentication failed!" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Create output directory for CSV files
$outputDir = $PSScriptRoot
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

Write-Host "`n=== Exporting All Booking Businesses ===" -ForegroundColor Cyan
try {
    $businesses = Get-MgBookingBusiness -ErrorAction Stop
    
    if ($businesses.Count -eq 0) {
        Write-Host "  No booking businesses found" -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($businesses.Count) booking business(es)" -ForegroundColor Green
        
        # Export to CSV
        $businessesFile = Join-Path $outputDir "bookings_businesses_$timestamp.csv"
        $businesses | Select-Object Id, DisplayName, Email, Phone, WebSiteUrl, IsPublished, PublicUrl, BusinessType | 
            Export-Csv -Path $businessesFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  [EXPORTED] $businessesFile" -ForegroundColor Green
    }
}
catch {
    Write-Host "  [ERROR] Failed to list businesses" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Get specific booking business details
Write-Host "`n=== Exporting Specific Business Details ===" -ForegroundColor Cyan
Write-Host "Business: $BOOKING_BUSINESS_ID" -ForegroundColor Yellow
try {
    $business = Get-MgBookingBusiness -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    Write-Host "  [SUCCESS] Retrieved business details" -ForegroundColor Green
    Write-Host "  Display Name: $($business.DisplayName)" -ForegroundColor White
    
    # Export business details
    $businessDetailFile = Join-Path $outputDir "business_details_$timestamp.csv"
    $business | Select-Object Id, DisplayName, Email, Phone, WebSiteUrl, IsPublished, PublicUrl, 
        BusinessType, DefaultCurrencyIso, LanguageTag | 
        Export-Csv -Path $businessDetailFile -NoTypeInformation -Encoding UTF8
    
    Write-Host "  [EXPORTED] $businessDetailFile" -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Failed to access specific business" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    
    # Disconnect and exit
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Export appointments
Write-Host "`n=== Exporting Appointments ===" -ForegroundColor Cyan
try {
    $appointments = Get-MgBookingBusinessAppointment -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    if ($appointments.Count -eq 0) {
        Write-Host "  No appointments scheduled" -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($appointments.Count) appointment(s)" -ForegroundColor Green
        
        # Prepare appointments data for export
        $appointmentsData = $appointments | ForEach-Object {
            [PSCustomObject]@{
                AppointmentId = $_.Id
                CustomerName = $_.CustomerName
                CustomerEmail = $_.CustomerEmailAddress
                CustomerPhone = $_.CustomerPhone
                StartDateTime = $_.Start.DateTime
                StartTimeZone = $_.Start.TimeZone
                EndDateTime = $_.End.DateTime
                EndTimeZone = $_.End.TimeZone
                Duration = $_.Duration
                ServiceId = $_.ServiceId
                ServiceName = $_.ServiceName
                ServiceLocation = $_.ServiceLocation
                IsLocationOnline = $_.IsLocationOnline
                OnlineMeetingUrl = $_.OnlineMeetingUrl
                OptOutOfCustomerEmail = $_.OptOutOfCustomerEmail
                SelfServiceAppointmentId = $_.SelfServiceAppointmentId
                MaximumAttendeesCount = $_.MaximumAttendeesCount
                FilledAttendeesCount = $_.FilledAttendeesCount
                AdditionalInformation = $_.AdditionalInformation
                CreatedDateTime = $_.CreatedDateTime
            }
        }
        
        # Export to CSV
        $appointmentsFile = Join-Path $outputDir "bookings_appointments_$timestamp.csv"
        $appointmentsData | Export-Csv -Path $appointmentsFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  [EXPORTED] $appointmentsFile" -ForegroundColor Green
        Write-Host "  Showing first 5 appointments:" -ForegroundColor Gray
        
        $appointmentsData | Select-Object -First 5 | ForEach-Object {
            Write-Host "    - $($_.CustomerName) | $($_.StartDateTime) | $($_.ServiceName)" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "  [ERROR] Failed to get appointments" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Export services
Write-Host "`n=== Exporting Services ===" -ForegroundColor Cyan
try {
    $services = Get-MgBookingBusinessService -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    if ($services.Count -eq 0) {
        Write-Host "  No services configured" -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($services.Count) service(s)" -ForegroundColor Green
        
        # Prepare services data for export
        $servicesData = $services | ForEach-Object {
            [PSCustomObject]@{
                ServiceId = $_.Id
                DisplayName = $_.DisplayName
                Description = $_.Description
                DefaultDuration = $_.DefaultDuration
                DefaultPrice = $_.DefaultPrice
                DefaultPriceType = $_.DefaultPriceType
                IsHiddenFromCustomers = $_.IsHiddenFromCustomers
                IsLocationOnline = $_.IsLocationOnline
                MaximumAttendeesCount = $_.MaximumAttendeesCount
                Notes = $_.Notes
                PostBuffer = $_.PostBuffer
                PreBuffer = $_.PreBuffer
                WebUrl = $_.WebUrl
            }
        }
        
        # Export to CSV
        $servicesFile = Join-Path $outputDir "bookings_services_$timestamp.csv"
        $servicesData | Export-Csv -Path $servicesFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  [EXPORTED] $servicesFile" -ForegroundColor Green
        $servicesData | ForEach-Object {
            Write-Host "    - $($_.DisplayName) | Duration: $($_.DefaultDuration) | Price: $($_.DefaultPrice)" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "  [ERROR] Failed to get services" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Export staff members
Write-Host "`n=== Exporting Staff Members ===" -ForegroundColor Cyan
try {
    $staff = Get-MgBookingBusinessStaffMember -BookingBusinessId $BOOKING_BUSINESS_ID -ErrorAction Stop
    
    if ($staff.Count -eq 0) {
        Write-Host "  No staff members configured" -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($staff.Count) staff member(s)" -ForegroundColor Green
        
        # Prepare staff data for export
        $staffData = $staff | ForEach-Object {
            [PSCustomObject]@{
                StaffId = $_.Id
                DisplayName = $_.DisplayName
                EmailAddress = $_.EmailAddress
                Role = $_.Role
                TimeZone = $_.TimeZone
                UseBusinessHours = $_.UseBusinessHours
                IsEmailNotificationEnabled = $_.IsEmailNotificationEnabled
                ColorCode = $_.ColorCode
            }
        }
        
        # Export to CSV
        $staffFile = Join-Path $outputDir "bookings_staff_$timestamp.csv"
        $staffData | Export-Csv -Path $staffFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  [EXPORTED] $staffFile" -ForegroundColor Green
        $staffData | ForEach-Object {
            Write-Host "    - $($_.DisplayName) | $($_.EmailAddress) | Role: $($_.Role)" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Host "  [ERROR] Failed to get staff members" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Cleanup
Write-Host "`n=== Export Complete ===" -ForegroundColor Green
Write-Host "`nCSV files saved to: $outputDir" -ForegroundColor Cyan
Write-Host "Timestamp: $timestamp" -ForegroundColor Gray
Write-Host "`nFiles created:" -ForegroundColor Yellow
Get-ChildItem -Path $outputDir -Filter "*_$timestamp.csv" | ForEach-Object {
    $size = [math]::Round($_.Length / 1KB, 2)
    Write-Host "  - $($_.Name) ($size KB)" -ForegroundColor White
}

Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Gray
Disconnect-MgGraph | Out-Null
Write-Host "[OK] Done!" -ForegroundColor Green
