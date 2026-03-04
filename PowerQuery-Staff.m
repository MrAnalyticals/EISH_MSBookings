/* 
===============================================================================
Microsoft Bookings - STAFF Table for Power Query
===============================================================================
This query retrieves all staff members from all booking businesses.
Use this as a separate table in your data model for staff lookups.

USAGE:
1. Copy this entire query
2. In Power BI/Excel: Get Data > Blank Query > Advanced Editor
3. Paste and click Done
4. Name the query "BookingStaff"

===============================================================================
DYNAMIC DATA REFRESH:
===============================================================================
This query uses "static base URLs" with RelativePath for Power BI Service.

Your data is STILL FULLY DYNAMIC! Every refresh:
✓ Gets current list of all booking businesses
✓ Retrieves current staff members for each business
✓ New staff members automatically appear
✓ Staff who leave automatically disappear
✓ Staff changes (email, role, etc.) update immediately

The "static" part is just the domain - all data updates dynamically!
===============================================================================
*/

let
    // Configuration
    TenantId = "YOUR_TENANT_ID_HERE",
    ClientId = "YOUR_CLIENT_ID_HERE",
    ClientSecret = "YOUR_CLIENT_SECRET_HERE",
    
    // Complete self-contained function that doesn't reference external steps
    GetAllStaff = () =>
        let
            // Get Access Token inline
            TokenBody = "client_id=" & ClientId 
                        & "&scope=https://graph.microsoft.com/.default"
                        & "&client_secret=" & ClientSecret
                        & "&grant_type=client_credentials",
            AccessToken = Json.Document(
                Web.Contents(
                    "https://login.microsoftonline.com",
                    [
                        RelativePath = TenantId & "/oauth2/v2.0/token",
                        Headers = [#"Content-Type" = "application/x-www-form-urlencoded"],
                        Content = Text.ToBinary(TokenBody)
                    ]
                )
            )[access_token],
            
            // Get All Booking Businesses inline
            BusinessesList = Json.Document(
                Web.Contents(
                    "https://graph.microsoft.com",
                    [
                        RelativePath = "v1.0/solutions/bookingBusinesses",
                        Headers = [Authorization = "Bearer " & AccessToken]
                    ]
                )
            )[value],
            
            // Function to get staff for a business
            GetStaff = (businessId as text, businessName as text) as table =>
                let
                    StaffPath = "v1.0/solutions/bookingBusinesses/" 
                             & businessId & "/staffMembers",
                    StaffResponse = try Json.Document(
                        Web.Contents(
                            "https://graph.microsoft.com",
                            [
                                RelativePath = StaffPath,
                                Headers = [Authorization = "Bearer " & AccessToken]
                            ]
                        )
                    ) otherwise [value = {}],
                    StaffList = StaffResponse[value],
                    StaffTable = 
                        if List.Count(StaffList) = 0 
                        then #table(
                            {"StaffId", "StaffName", "StaffEmail", "Role", "TimeZone", 
                             "UseBusinessHours", "EmailNotificationEnabled", "BusinessId", "BusinessName"}, 
                            {}
                        )
                        else 
                            let
                                TempTable = Table.FromList(StaffList, Splitter.SplitByNothing()),
                                ExpandedTable = Table.ExpandRecordColumn(
                                    TempTable,
                                    "Column1",
                                    {"id", "displayName", "emailAddress", "role", "timeZone",
                                     "useBusinessHours", "isEmailNotificationEnabled"},
                                    {"StaffId", "StaffName", "StaffEmail", "Role", "TimeZone",
                                     "UseBusinessHours", "EmailNotificationEnabled"}
                                ),
                                AddBusinessInfo = Table.AddColumn(
                                    Table.AddColumn(
                                        ExpandedTable, 
                                        "BusinessId", 
                                        each businessId
                                    ),
                                    "BusinessName",
                                    each businessName
                                )
                            in
                                AddBusinessInfo
                in
                    StaffTable,
            
            // Get staff for all businesses
            StaffFromAllBusinesses = List.Transform(
                BusinessesList,
                each GetStaff([id], [displayName])
            ),
            
            // Combine all staff
            CombinedStaff = Table.Combine(StaffFromAllBusinesses),
            
            // Reorder columns
            ReorderedColumns = Table.ReorderColumns(
                CombinedStaff,
                {"BusinessId", "BusinessName", "StaffId", "StaffName", "StaffEmail", 
                 "Role", "TimeZone", "UseBusinessHours", "EmailNotificationEnabled"}
            ),
            
            // Add composite key for relationship (BusinessId + StaffId)
            // This ensures StaffId is unique across all businesses
            Result = Table.AddColumn(ReorderedColumns, "BusinessStaffKey", each [BusinessId] & "|" & [StaffId], type text)
        in
            Result,
    
    // Execute the function and return result
    FinalResult = GetAllStaff()
in
    FinalResult
