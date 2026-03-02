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
*/

let
    // Configuration
    TenantId = "YOUR_TENANT_ID_HERE",
    ClientId = "YOUR_CLIENT_ID_HERE",
    ClientSecret = "YOUR_CLIENT_SECRET_HERE",
    
    // Get Access Token
    TokenUrl = "https://login.microsoftonline.com/" & TenantId & "/oauth2/v2.0/token",
    TokenBody = "client_id=" & ClientId 
                & "&scope=https://graph.microsoft.com/.default"
                & "&client_secret=" & ClientSecret
                & "&grant_type=client_credentials",
    TokenResponse = Json.Document(
        Web.Contents(
            TokenUrl,
            [
                Headers = [#"Content-Type" = "application/x-www-form-urlencoded"],
                Content = Text.ToBinary(TokenBody)
            ]
        )
    ),
    AccessToken = TokenResponse[access_token],
    
    // Get All Booking Businesses
    BusinessesUrl = "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses",
    BusinessesResponse = Json.Document(
        Web.Contents(
            BusinessesUrl,
            [Headers = [Authorization = "Bearer " & AccessToken]]
        )
    ),
    BusinessesList = BusinessesResponse[value],
    
    // Function to get staff for a business
    GetStaff = (businessId as text, businessName as text) as table =>
        let
            StaffUrl = "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/" 
                     & businessId & "/staffMembers",
            StaffResponse = try Json.Document(
                Web.Contents(
                    StaffUrl,
                    [Headers = [Authorization = "Bearer " & AccessToken]]
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
    ReorderColumns = Table.ReorderColumns(
        CombinedStaff,
        {"BusinessId", "BusinessName", "StaffId", "StaffName", "StaffEmail", 
         "Role", "TimeZone", "UseBusinessHours", "EmailNotificationEnabled"}
    )
in
    ReorderColumns
