section Section1;
shared slvr_MSBookings_BookingsData = let
  /*
### Star Schema Diagram

```
┌─────────────────────┐
│  Bookings_Services  │ (Dimension Table)
│  ─────────────────  │
│  ServiceId (PK)     │
│  BusinessId         │
│  ServiceName        │
│  Duration           │
│  Price              │
└──────────┬──────────┘
           │
           │ Many-to-One
           ├─────────────────────────┐
           │                         │
┌──────────▼──────────────┐   ┌──────▼───────────┐
│ Bookings_Appointments   │   │  Bookings_Staff  │ (Dimension Table)
│ ─────────────────────   │   │  ──────────────  │
│ AppointmentId (PK)      │   │  StaffId (PK)    │
│ ServiceId (FK) ─────────┘   │  BusinessId      │
│ BusinessId (FK) ────────────┤  StaffName       │
│ CustomerName            │   │  StaffEmail      │
│ StartDateTime           │   │  Role            │
│ EndDateTime             │   └──────────────────┘
│ AppointmentDate         │
└─────────────────────────┘
    (Fact Table)
```

**Key:**
- **PK** = Primary Key (unique identifier)
- **FK** = Foreign Key (links to dimension table)
- **Fact Table** = Contains measurable events (appointments)
- **Dimension Tables** = Contains descriptive attributes (services, staff)

  */
  // ============================================================================
  // CONFIGURATION - Your Azure AD App Registration Details
  // ============================================================================
  // These credentials authenticate your app to access Microsoft Graph API
  // Find these values in Azure Portal > App Registrations > Your App
  TenantId = "YOUR_TENANT_ID_HERE",
  // Your organization's Azure AD tenant ID
  ClientId = "YOUR_CLIENT_ID_HERE",
  // App (client) ID from app registration
  ClientSecret = "YOUR_CLIENT_SECRET_HERE",
  // Client secret value (not secret ID)
  // ============================================================================
  // STEP 1: Get OAuth Access Token (Authentication)
  // ============================================================================
  // We use the OAuth 2.0 Client Credentials flow for app-only authentication
  // This allows the app to access data without requiring user interaction
  // Build the request body with our credentials
  // Format: URL-encoded form data required by OAuth 2.0 spec
  TokenBody = "client_id=" & ClientId & "&scope=https://graph.microsoft.com/.default" & "&client_secret=" & ClientSecret & "&grant_type=client_credentials",
  // Client credentials flow
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
  BusinessesList = if BusinessesResponse = null or BusinessesResponse[value] = null 
        then {} 
        else BusinessesResponse[value],
  // Convert the list of business objects into a table
  // Each row represents one booking business
  BusinessesTable = Table.FromList(BusinessesList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
  // Expand the record column to show individual business properties
  // This creates separate columns for id, displayName, email, phone
  BusinessesExpanded = if Table.RowCount(BusinessesTable) = 0 
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
                             "serviceId", "serviceName", "startDateTime", "endDateTime", "duration",
                             "isLocationOnline", "onlineMeetingUrl", "additionalInformation",
                             "staffMemberIds"},
                            {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone",
                             "ServiceId", "ServiceName", "StartDateTimeRecord", "EndDateTimeRecord", "Duration",
                             "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo",
                             "StaffMemberIds"}
                        ),
                        
                        // Extract start datetime from nested record structure
                        // The API returns startDateTime/endDateTime as objects with dateTime and timeZone properties
                        ExtractStart = Table.AddColumn(
                            ExpandedTable, 
                            "StartDateTime", 
                            each try [StartDateTimeRecord][dateTime] otherwise null  // Get dateTime property or null if missing
                        ),
                        
                        // Extract end datetime from nested record structure
                        ExtractEnd = Table.AddColumn(
                            ExtractStart, 
                            "EndDateTime", 
                            each try [EndDateTimeRecord][dateTime] otherwise null
                        ),
                        
                        // Remove the original nested StartDateTimeRecord and EndDateTimeRecord columns
                        // We now have flattened StartDateTime and EndDateTime instead
                        RemoveTempColumns = Table.RemoveColumns(ExtractEnd, {"StartDateTimeRecord", "EndDateTimeRecord"}),
                        
                        // Expand staffMemberIds array into multiple rows (Bridge Table pattern)
                        // This creates one row per staff member assigned to each appointment
                        // Appointments with multiple staff will become multiple rows
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
  AddAppointments = Table.AddColumn(BusinessesExpanded, "Appointments", each try GetAppointments([BusinessId]) otherwise #table(  // For each row, call function with BusinessId
            {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", 
             "ServiceId", "ServiceName", "StartDateTime", "EndDateTime", 
             "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo", "StaffId"}, 
            {}  // If function fails, return empty table instead of error
        )),
  // Add Services column: Call GetServices() for each business
  // Stores all services offered by each business
  AddServices = Table.AddColumn(AddAppointments, "Services", each try GetServices([BusinessId]) otherwise #table(
            {"ServiceId", "ServiceDisplayName", "ServiceDuration", 
             "ServicePrice", "ServiceDescription"}, 
            {}
        )),
  // Add Staff column: Call GetStaff() for each business
  // Stores all staff members associated with each business
  AddStaff = Table.AddColumn(AddServices, "Staff", each try GetStaff([BusinessId]) otherwise #table(
            {"StaffId", "StaffName", "StaffEmail", "StaffRole"}, 
            {}
        )),
  // ============================================================================
  // STEP 8: Expand Appointments (Fact Table Output)
  // ============================================================================
  // This is the KEY transformation - it flattens the nested structure
  // Each appointment becomes its own row, with business info repeated
  // Result: Clean fact table with one row per appointment
  // Services and Staff are separate queries for better data modeling
  ExpandAppointments = Table.ExpandTableColumn(AddAppointments, "Appointments", {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", "ServiceId", "ServiceName", "StartDateTime", "EndDateTime", "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo", "StaffId"}, {"AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", "ServiceId", "ServiceName", "StartDateTime", "EndDateTime", "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo", "StaffId"}),
  // ============================================================================
  // STEP 8.5: Convert DateTime Strings to Proper Dates
  // ============================================================================
  // The API returns dates as text strings (e.g., "2024-03-15T14:30:00")
  // We convert them to DateTime type so they work with date functions and filtering
  ConvertDates = Table.TransformColumns(
    ExpandAppointments, 
    {
      {"StartDateTime", each try DateTime.FromText(_) otherwise null, type nullable datetime},
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
  AddCalculatedColumns = Table.AddColumn(ConvertDates, "AppointmentDate", each try DateTime.Date([StartDateTime]) otherwise null, type nullable date),
  // Add DayOfWeek column (e.g., "Monday", "Tuesday")
  // Useful for analyzing which days are busiest
  AddWeekday = Table.AddColumn(AddCalculatedColumns, "DayOfWeek", each try Date.DayOfWeekName([AppointmentDate]) otherwise null, type nullable text),
  // Add Month column (e.g., "January", "February")
  // Useful for monthly trend analysis
  AddMonth = Table.AddColumn(AddWeekday, "Month", each try Date.MonthName([AppointmentDate]) otherwise null, type nullable text),
  // Add Year column (e.g., 2024, 2025, 2026)
  // Useful for year-over-year comparisons
  AddYear = Table.AddColumn(AddMonth, "Year", each try Date.Year([AppointmentDate]) otherwise null, type nullable number),
  // ============================================================================
  // STEP 10: Final Column Ordering and Cleanup
  // ============================================================================
  // Reorder columns to show the most important information first
  // This creates a clean fact table optimized for Star Schema relationships
  // Join to Services and Staff queries using ServiceId and BusinessId
  ReorderColumns = Table.ReorderColumns(AddYear, {"BusinessId", "BusinessName", "BusinessEmail", "BusinessPhone", "AppointmentId", "CustomerName", "CustomerEmail", "CustomerPhone", "ServiceId", "ServiceName", "StaffId", "StartDateTime", "EndDateTime", "AppointmentDate", "DayOfWeek", "Month", "Year", "Duration", "IsLocationOnline", "OnlineMeetingUrl", "AdditionalInfo"}),
  // Set proper data types for all columns
  // This optimizes performance and enables proper filtering/sorting
  // Also helps Power BI understand what kind of data each column contains
  SetDataTypes = Table.TransformColumnTypes(ReorderColumns, {{"BusinessId", type text}, {"BusinessName", type text}, {"AppointmentId", type text}, {"CustomerName", type text}, {"CustomerEmail", type text}, {"ServiceName", type text}, {"StaffId", type text}, {"Duration", type text}, {"IsLocationOnline", type logical}}),
  #"Transform columns" = Table.TransformColumnTypes(SetDataTypes, {{"BusinessEmail", type text}, {"BusinessPhone", type text}, {"CustomerPhone", type text}, {"ServiceId", type text}, {"OnlineMeetingUrl", type text}, {"AdditionalInfo", type text}}),
  //#"Profile (temporary)" = Table.Profile(#"Transform columns"),//addittionally added
  #"Replace errors" = Table.ReplaceErrorValues(#"Transform columns", {{"BusinessEmail", null}, {"BusinessPhone", null}, {"CustomerPhone", null}, {"ServiceId", null}, {"OnlineMeetingUrl", null}, {"AdditionalInfo", null}}),
  //#"Replace errors" = Table.ReplaceErrorValues(#"Profile (temporary)", {{"BusinessEmail", null}, {"BusinessPhone", null}, {"CustomerPhone", null}, {"ServiceId", null}, {"OnlineMeetingUrl", null}, {"AdditionalInfo", null}})
  // Add composite keys for relationships (BusinessId + ServiceId/StaffId)
  // These are needed because ServiceId and StaffId alone are not unique across businesses
  AddServiceKey = Table.AddColumn(#"Replace errors", "BusinessServiceKey", each [BusinessId] & "|" & [ServiceId], type text),
  AddStaffKey = Table.AddColumn(AddServiceKey, "BusinessStaffKey", each [BusinessId] & "|" & [StaffId], type text)
in
  AddStaffKey;
