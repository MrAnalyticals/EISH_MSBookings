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
DYNAMIC DATA REFRESH:
===============================================================================
This query uses "static base URLs" with RelativePath for Power BI Service.

Your data is STILL FULLY DYNAMIC! Every refresh:
✓ Gets current list of all booking businesses
✓ Retrieves current services for each business
✓ New services automatically appear
✓ Deleted services automatically disappear
✓ Service changes (price, duration, etc.) update immediately

The "static" part is just the domain - all data updates dynamically!
===============================================================================
*/

let
    // Configuration
    TenantId = "YOUR_TENANT_ID_HERE",
    ClientId = "YOUR_CLIENT_ID_HERE",
    ClientSecret = "YOUR_CLIENT_SECRET_HERE",
    
    // Complete self-contained function that doesn't reference external steps
    GetAllServices = () =>
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
            
            // Function to get services for a business
    GetServices = (businessId as text, businessName as text) as table =>
        let
            ServicesPath = "v1.0/solutions/bookingBusinesses/" 
                        & businessId & "/services",
            ServicesResponse = try Json.Document(
                Web.Contents(
                    "https://graph.microsoft.com",
                    [
                        RelativePath = ServicesPath,
                        Headers = [Authorization = "Bearer " & AccessToken]
                    ]
                )
            ) otherwise [value = {}],
            ServicesList = ServicesResponse[value],
            ServicesTable = 
                if List.Count(ServicesList) = 0 
                then #table(
                    {"ServiceId", "ServiceDisplayName", "ServiceDuration", 
                     "ServicePrice", "ServiceDescription", "BusinessId", "BusinessName"}, 
                    {}
                )
                else 
                    let
                        TempTable = Table.FromList(ServicesList, Splitter.SplitByNothing()),
                        ExpandedTable = Table.ExpandRecordColumn(
                            TempTable,
                            "Column1",
                            {"id", "displayName", "defaultDuration", "defaultPrice", "description"},
                            {"ServiceId", "ServiceDisplayName", "ServiceDuration", 
                             "ServicePrice", "ServiceDescription"}
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
    Result = Table.ReorderColumns(
        CombinedServices,
        {"BusinessId", "BusinessName", "ServiceId", "ServiceDisplayName", 
         "ServiceDuration", "ServicePrice", "ServiceDescription"}
    )
in
    Result,

// Execute the function and return result
FinalResult = GetAllServices()
in
FinalResult
