/* 
===============================================================================
Microsoft Bookings - Appointments Query (Fact Table)
===============================================================================
This is the MAIN query for appointments data. It should be used with the
companion queries: PowerQuery-Services.m and PowerQuery-Staff.m

THREE-QUERY ARCHITECTURE (Best Practice):
1. THIS QUERY (Appointments) = Fact table with appointment transactions
2. PowerQuery-Services.m = Dimension table for service lookup
3. PowerQuery-Staff.m = Dimension table for staff lookup

After loading all three queries in Power BI/Excel:
- Create relationship: Appointments[ServiceId] → Services[ServiceId]
- Create relationship: Appointments[BusinessId] → Services[BusinessId]
- Create relationship: Appointments[BusinessId] → Staff[BusinessId]

This follows Star Schema best practices for optimal performance and reusability.

===============================================================================
DYNAMIC DATA REFRESH:
===============================================================================
This query uses "static base URLs" (e.g., https://graph.microsoft.com) with
RelativePath parameters to comply with Power BI Service requirements.

IMPORTANT: Your data is STILL FULLY DYNAMIC!
Every time you refresh, this query:
✓ Gets a fresh list of ALL current booking businesses
✓ Retrieves current appointments for each business
✓ Automatically includes new appointments
✓ Automatically excludes deleted appointments
✓ Adapts if businesses are added or removed

The "static" part is just the domain name - everything else updates dynamically!

INSTRUCTIONS:
1. Copy this entire query
2. In Power BI/Excel: Get Data > Blank Query
3. Open Advanced Editor and paste this code
4. Click Done
5. Repeat for PowerQuery-Services.m and PowerQuery-Staff.m
6. Set up relationships in Model view

The query is fully dynamic - it will adapt if:
- Booking businesses are added/removed
- Services are added/removed from any business
- Staff members join or leave
- Appointments are created/deleted

SECURITY NOTE:
- This query contains your app credentials
- Do NOT share this query or your Power BI file publicly
- Consider using Power BI parameters for credentials in production

===============================================================================
*/

let
    // ============================================================================
    // CONFIGURATION - Your Azure AD App Registration Details
    // ============================================================================
    // These credentials authenticate your app to access Microsoft Graph API
    // Find these values in Azure Portal > App Registrations > Your App
    
    TenantId = "YOUR_TENANT_ID_HERE",        // Your organization's Azure AD tenant ID
    ClientId = "YOUR_CLIENT_ID_HERE",        // App (client) ID from app registration
    ClientSecret = "YOUR_CLIENT_SECRET_HERE", // Client secret value (not secret ID)
    
    // ============================================================================
    // STEP 1: Get OAuth Access Token (Authentication)
    // ============================================================================
    // We use the OAuth 2.0 Client Credentials flow for app-only authentication
    // This allows the app to access data without requiring user interaction
    
    // Build the request body with our credentials
    // Format: URL-encoded form data required by OAuth 2.0 spec
    TokenBody = "client_id=" & ClientId 
                & "&scope=https://graph.microsoft.com/.default"  // Request all app permissions
                & "&client_secret=" & ClientSecret
                & "&grant_type=client_credentials",              // Client credentials flow
    
    // Make HTTP POST request to get access token
    // Using static base URL with RelativePath to avoid "dynamic data source" error
    TokenResponse = Json.Document(
        Web.Contents(
            "https://login.microsoftonline.com",
            [
                RelativePath = TenantId & "/oauth2/v2.0/token",
                Headers = [#"Content-Type" = "application/x-www-form-urlencoded"],
                Content = Text.ToBinary(TokenBody)
            ]
        )
    ),
    
    // Extract the bearer token from the response
    // This token will be included in all subsequent API calls
    AccessToken = TokenResponse[access_token],
    
    // ============================================================================
    // STEP 2: Define Graph API Helper Function
    // ============================================================================
    // This reusable function makes authenticated GET requests to Microsoft Graph
    // It handles errors gracefully by returning null instead of breaking the query
    // Uses RelativePath to avoid "dynamic data source" errors in Power BI Service
    
    GetGraphData = (relativePath as text) as any =>
        let
            // Attempt to call the API and parse JSON response
            Response = try Json.Document(
                Web.Contents(
                    "https://graph.microsoft.com",
                    [
                        RelativePath = relativePath,
                        Headers = [
                            Authorization = "Bearer " & AccessToken,
                            #"Content-Type" = "application/json"
                        ]
                    ]
                )
            ) otherwise null  // If API call fails, return null instead of error
        in
            Response,
    
    // ============================================================================
    // STEP 3: Get All Booking Businesses (Dynamically)
    // ============================================================================
    // This section retrieves ALL booking calendars in your Microsoft 365 tenant
    // No hardcoded IDs - it automatically finds everything available
    
    // Call the API to get all booking businesses
    BusinessesResponse = GetGraphData("v1.0/solutions/bookingBusinesses"),
    
    // Extract the array of businesses from the response
    // The 'value' property contains the array of business objects
    // If API fails or returns no data, use empty list
    BusinessesList = 
        if BusinessesResponse = null or BusinessesResponse[value] = null 
        then {} 
        else BusinessesResponse[value],
    
    // Convert the list of business objects into a table
    // Each row represents one booking business
    BusinessesTable = Table.FromList(
        BusinessesList, 
        Splitter.SplitByNothing(),  // Don't split - each item is a complete record
        null,                        // Auto-generate column names
        null, 
        ExtraValues.Error
    ),
    
    // Expand the record column to show individual business properties
    // This creates separate columns for id, displayName, email, phone
    BusinessesExpanded = 
        if Table.RowCount(BusinessesTable) = 0 
        then #table({"BusinessId", "BusinessName", "BusinessEmail", "BusinessPhone"}, {})  // Empty table if no data
        else Table.ExpandRecordColumn(
            BusinessesTable,
            "Column1",                                                  // The column containing records
            {"id", "displayName", "email", "phone"},                   // Properties to extract
            {"BusinessId", "BusinessName", "BusinessEmail", "BusinessPhone"}  // New column names
        ),
    
    // ============================================================================
    // STEP 4: Get Appointments for Each Business (Dynamically)
    // ============================================================================
    // This function retrieves all appointments for a given booking business
    // It's designed to be called once per business in a loop
    
    GetAppointments = (businessId as text) as table =>
        let
            // Build the API relative path for this specific business's appointments
            AppointmentsPath = "v1.0/solutions/bookingBusinesses/" 
                            & businessId & "/appointments",
            
            // Call the API to get appointments
            AppointmentsResponse = GetGraphData(AppointmentsPath),
            
            // Extract the appointments array from the response
            // Returns empty list if API fails or business has no appointments
            AppointmentsList = 
                if AppointmentsResponse = null or AppointmentsResponse[value] = null 
                then {} 
                else AppointmentsResponse[value],
            
            // Convert appointments to table format
            AppointmentsTable = 
                if List.Count(AppointmentsList) = 0 
                then 
                    // No appointments - return empty table with proper column structure
                    #table(
                        {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", 
                         "ServiceId", "ServiceName", "StartDateTime", "EndDateTime", 
                         "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo"}, 
                        {}  // Empty rows
                    )
                else 
                    let
                        // Convert list of appointment records to a table
                        TempTable = Table.FromList(AppointmentsList, Splitter.SplitByNothing()),
                        
                        // Expand the nested record to get individual appointment properties
                        // Each appointment is a JSON object with multiple fields
                        ExpandedTable = Table.ExpandRecordColumn(
                            TempTable,
                            "Column1",
                            {"id", "customerName", "customerEmailAddress", "customerPhone",
                             "serviceId", "serviceName", "start", "end", "duration",
                             "isLocationOnline", "onlineMeetingUrl", "additionalInformation",
                             "staffMemberIds"},
                            {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone",
                             "ServiceId", "ServiceName", "Start", "End", "Duration",
                             "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo",
                             "StaffMemberIds"}
                        ),
                        
                        // Extract start datetime from nested record structure
                        // The API returns start/end as objects with dateTime and timeZone properties
                        ExtractStart = Table.AddColumn(
                            ExpandedTable, 
                            "StartDateTime", 
                            each try [Start][dateTime] otherwise null  // Get dateTime property or null if missing
                        ),
                        
                        // Extract end datetime from nested record structure
                        ExtractEnd = Table.AddColumn(
                            ExtractStart, 
                            "EndDateTime", 
                            each try [End][dateTime] otherwise null
                        ),
                        
                        // Remove the original nested Start and End columns
                        // We now have flattened StartDateTime and EndDateTime instead
                        RemoveTempColumns = Table.RemoveColumns(ExtractEnd, {"Start", "End"}),
                        
                        // Expand staffMemberIds array into multiple rows (Bridge Table pattern)
                        // This creates one row per staff member assigned to each appointment
                        // Appointments with multiple staff (12.7%) will become multiple rows
                        ExpandStaff = Table.ExpandListColumn(RemoveTempColumns, "StaffMemberIds"),
                        
                        // Rename the expanded column to match relationship naming convention
                        RenameStaffId = Table.RenameColumns(ExpandStaff, {{"StaffMemberIds", "StaffId"}})
                    in
                        RenameStaffId
        in
            AppointmentsTable,
    
    // ============================================================================
    // STEP 5: Get Services for Each Business (Dynamically)
    // ============================================================================
    // This function retrieves all services offered by a booking business
    // Services define what customers can book (e.g., "Energy Report Call")
    
    GetServices = (businessId as text) as table =>
        let
            // Build the API relative path for this business's services
            ServicesPath = "v1.0/solutions/bookingBusinesses/" 
                        & businessId & "/services",
            
            // Call the API to get services
            ServicesResponse = GetGraphData(ServicesPath),
            
            // Extract services array from response
            ServicesList = 
                if ServicesResponse = null or ServicesResponse[value] = null 
                then {} 
                else ServicesResponse[value],
            
            // Convert services to table
            ServicesTable = 
                if List.Count(ServicesList) = 0 
                then 
                    // No services - return empty table with proper structure
                    #table(
                        {"ServiceId", "ServiceDisplayName", "ServiceDuration", 
                         "ServicePrice", "ServiceDescription"}, 
                        {}
                    )
                else 
                    let
                        // Convert list to table
                        TempTable = Table.FromList(ServicesList, Splitter.SplitByNothing()),
                        
                        // Expand service properties
                        // Extract key service details: id, name, duration, price, description
                        ExpandedTable = Table.ExpandRecordColumn(
                            TempTable,
                            "Column1",
                            {"id", "displayName", "defaultDuration", "defaultPrice", "description"},
                            {"ServiceId", "ServiceDisplayName", "ServiceDuration", 
                             "ServicePrice", "ServiceDescription"}
                        )
                    in
                        ExpandedTable
        in
            ServicesTable,
    
    // ============================================================================
    // STEP 6: Get Staff for Each Business (Dynamically)
    // ============================================================================
    // This function retrieves all staff members who can provide services
    // Staff members are the employees who fulfill appointments
    
    GetStaff = (businessId as text) as table =>
        let
            // Build the API relative path for this business's staff
            StaffPath = "v1.0/solutions/bookingBusinesses/" 
                     & businessId & "/staffMembers",
            
            // Call the API to get staff
            StaffResponse = GetGraphData(StaffPath),
            
            // Extract staff array from response
            StaffList = 
                if StaffResponse = null or StaffResponse[value] = null 
                then {} 
                else StaffResponse[value],
            
            // Convert staff to table
            StaffTable = 
                if List.Count(StaffList) = 0 
                then 
                    // No staff - return empty table with proper structure
                    #table(
                        {"StaffId", "StaffName", "StaffEmail", "StaffRole"}, 
                        {}
                    )
                else 
                    let
                        // Convert list to table
                        TempTable = Table.FromList(StaffList, Splitter.SplitByNothing()),
                        
                        // Expand staff properties
                        // Extract staff details: id, name, email, role
                        ExpandedTable = Table.ExpandRecordColumn(
                            TempTable,
                            "Column1",
                            {"id", "displayName", "emailAddress", "role"},
                            {"StaffId", "StaffName", "StaffEmail", "StaffRole"}
                        )
                    in
                        ExpandedTable
        in
            StaffTable,
    
    // ============================================================================
    // STEP 7: Combine All Data (Main Processing Loop)
    // ============================================================================
    // This section adds nested columns to the businesses table
    // Each business row will have sub-tables for appointments, services, and staff
    
    // Add Appointments column: Call GetAppointments() for each business
    // This creates a table-within-a-table structure (nested table)
    AddAppointments = Table.AddColumn(
        BusinessesExpanded,
        "Appointments",                       // Name of the new column
        each try GetAppointments([BusinessId]) otherwise #table(  // For each row, call function with BusinessId
            {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", 
             "ServiceId", "ServiceName", "StartDateTime", "EndDateTime", 
             "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo"}, 
            {}  // If function fails, return empty table instead of error
        )
    ),
    
    // Add Services column: Call GetServices() for each business
    // Stores all services offered by each business
    AddServices = Table.AddColumn(
        AddAppointments,
        "Services",                          // Name of the new column
        each try GetServices([BusinessId]) otherwise #table(
            {"ServiceId", "ServiceDisplayName", "ServiceDuration", 
             "ServicePrice", "ServiceDescription"}, 
            {}
        )
    ),
    
    // Add Staff column: Call GetStaff() for each business
    // Stores all staff members associated with each business
    AddStaff = Table.AddColumn(
        AddServices,
        "Staff",                             // Name of the new column
        each try GetStaff([BusinessId]) otherwise #table(
            {"StaffId", "StaffName", "StaffEmail", "StaffRole"}, 
            {}
        )
    ),
    
    // ============================================================================
    // STEP 8: Expand Appointments (Fact Table Output)
    // ============================================================================
    // This is the KEY transformation - it flattens the nested structure
    // Each appointment becomes its own row, with business info repeated
    // Result: Clean fact table with one row per appointment
    // Services and Staff are separate queries for better data modeling
    
    ExpandAppointments = Table.ExpandTableColumn(
        AddAppointments,                     // Expand from appointments only (not the Staff/Services versions)
        "Appointments",                      // The nested table column to expand
        {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", 
         "ServiceId", "ServiceName", "StartDateTime", "EndDateTime", 
         "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo"},  // Columns to bring up
        {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", 
         "ServiceId", "ServiceName", "StartDateTime", "EndDateTime", 
         "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo"}   // New column names
    ),
    
    // ============================================================================
    // STEP 8.5: Convert DateTime Strings to Proper Dates
    // ============================================================================
    // The API returns dates as text strings (e.g., "2024-03-15T14:30:00")
    // We convert them to DateTime type so they work with date functions and filtering
    
    ConvertDates = Table.TransformColumns(
        ExpandAppointments,
        {
            // Transform StartDateTime: Try to parse text as datetime, use null if fails
            {"StartDateTime", each try DateTime.FromText(_) otherwise null, type nullable datetime},
            
            // Transform EndDateTime: Try to parse text as datetime, use null if fails
            {"EndDateTime", each try DateTime.FromText(_) otherwise null, type nullable datetime}
        }
    ),
    
    // ============================================================================
    // STEP 9: Add Helpful Calculated Columns
    // ============================================================================
    // These calculated columns make it easier to analyze appointments in reports
    // They extract useful information from the datetime fields
    
    // Add AppointmentDate column (Date only, without time)
    // Useful for grouping appointments by day
    AddCalculatedColumns = Table.AddColumn(
        ConvertDates,
        "AppointmentDate",                   // Column name
        each try DateTime.Date([StartDateTime]) otherwise null,  // Extract date part from datetime
        type nullable date                   // Set data type for better performance
    ),
    
    // Add DayOfWeek column (e.g., "Monday", "Tuesday")
    // Useful for analyzing which days are busiest
    AddWeekday = Table.AddColumn(
        AddCalculatedColumns,
        "DayOfWeek",
        each try Date.DayOfWeekName([AppointmentDate]) otherwise null,
        type nullable text
    ),
    
    // Add Month column (e.g., "January", "February")
    // Useful for monthly trend analysis
    AddMonth = Table.AddColumn(
        AddWeekday,
        "Month",
        each try Date.MonthName([AppointmentDate]) otherwise null,
        type nullable text
    ),
    
    // Add Year column (e.g., 2024, 2025, 2026)
    // Useful for year-over-year comparisons
    AddYear = Table.AddColumn(
        AddMonth,
        "Year",
        each try Date.Year([AppointmentDate]) otherwise null,
        type nullable number
    ),
    
    // ============================================================================
    // STEP 10: Final Column Ordering and Cleanup
    // ============================================================================
    // Reorder columns to show the most important information first
    // This creates a clean fact table optimized for Star Schema relationships
    // Join to Services and Staff queries using ServiceId and BusinessId
    
    ReorderColumns = Table.ReorderColumns(
        AddYear,
        {
            // Business context first (who owns this appointment)
            "BusinessId", "BusinessName", "BusinessEmail", "BusinessPhone",
            
            // Appointment identification
            "AppointmentId", 
            
            // Customer information
            "CustomerName", "CustomerEmail", "CustomerPhone",
            
            // Service information (what was booked)
            "ServiceId", "ServiceName", 
            
            // Timing information (when is it scheduled)
            "StartDateTime", "EndDateTime", "AppointmentDate", "DayOfWeek", "Month", "Year",
            
            // Additional details
            "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo"
        }
    ),
    
    // Set proper data types for all columns
    // This optimizes performance and enables proper filtering/sorting
    // Also helps Power BI understand what kind of data each column contains
    SetDataTypes = Table.TransformColumnTypes(
        ReorderColumns,
        {
            {"BusinessId", type text},           // Text identifier
            {"BusinessName", type text},         // Text name
            {"AppointmentId", type text},        // Text identifier
            {"CustomerName", type text},         // Text name
            {"CustomerEmail", type text},        // Text email
            {"ServiceName", type text},          // Text name
            {"Duration", type text},             // Duration as text (e.g., "01:00:00")
            {"IsLocationOnline", type logical}   // Boolean true/false
        }
    )
in
    // ============================================================================
    // FINAL OUTPUT: Appointments Fact Table
    // ============================================================================
    // This table contains:
    // - One row per appointment (flattened from all booking businesses)
    // - Complete business, customer, and appointment details
    // - ServiceId for joining to Services dimension table
    // - BusinessId for joining to Services and Staff dimension tables
    //
    // RECOMMENDED RELATIONSHIPS (Set up in Power BI Model view):
    // - Appointments[ServiceId] → Services[ServiceId] (Many-to-One)
    // - Appointments[BusinessId] → Services[BusinessId] (Many-to-One)
    // - Appointments[BusinessId] → Staff[BusinessId] (Many-to-One)
    //
    // This Star Schema design provides:
    // ✓ Optimal query performance
    // ✓ No data duplication
    // ✓ Reusable dimension tables
    // ✓ Better DAX measure efficiency
    SetDataTypes
