/* 
===============================================================================
Microsoft Bookings - SERVICES Table for Power Query
===============================================================================
This query retrieves all services from all booking businesses.
Use this as a separate table in your data model for service lookups.

USAGE:
1. Copy this entire query
2. In Power BI/Excel: Get Data > Blank Query > Advanced Editor
3. Paste and click Done
4. Name the query "BookingServices"

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
    
    // Function to get services for a business
    GetServices = (businessId as text, businessName as text) as table =>
        let
            ServicesUrl = "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/" 
                        & businessId & "/services",
            ServicesResponse = try Json.Document(
                Web.Contents(
                    ServicesUrl,
                    [Headers = [Authorization = "Bearer " & AccessToken]]
                )
            ) otherwise [value = {}],
            ServicesList = ServicesResponse[value],
            ServicesTable = 
                if List.Count(ServicesList) = 0 
                then #table(
                    {"ServiceId", "ServiceName", "Duration", "Price", "PriceType", 
                     "Description", "IsHidden", "MaxAttendees", "BusinessId", "BusinessName"}, 
                    {}
                )
                else 
                    let
                        TempTable = Table.FromList(ServicesList, Splitter.SplitByNothing()),
                        ExpandedTable = Table.ExpandRecordColumn(
                            TempTable,
                            "Column1",
                            {"id", "displayName", "defaultDuration", "defaultPrice", 
                             "defaultPriceType", "description", "isHiddenFromCustomers", 
                             "maximumAttendeesCount"},
                            {"ServiceId", "ServiceName", "Duration", "Price", 
                             "PriceType", "Description", "IsHidden", "MaxAttendees"}
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
            ServicesTable,
    
    // Get services for all businesses
    ServicesFromAllBusinesses = List.Transform(
        BusinessesList,
        each GetServices([id], [displayName])
    ),
    
    // Combine all services
    CombinedServices = Table.Combine(ServicesFromAllBusinesses),
    
    // Reorder columns
    ReorderColumns = Table.ReorderColumns(
        CombinedServices,
        {"BusinessId", "BusinessName", "ServiceId", "ServiceName", 
         "Duration", "Price", "PriceType", "Description", "IsHidden", "MaxAttendees"}
    )
in
    ReorderColumns
