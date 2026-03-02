# Power Query M Queries for Microsoft Bookings

This folder contains ready-to-use Power Query (M language) queries that connect to Microsoft Bookings via Microsoft Graph API and retrieve all your booking data dynamically.

## 🏗️ Architecture: Star Schema (Best Practice)

These queries follow **data warehousing best practices** with a **Star Schema** design:

- **Fact Table**: `PowerQuery-BookingsData.m` (Appointments) - Transaction records
- **Dimension Tables**: 
  - `PowerQuery-Services.m` - Service lookup data
  - `PowerQuery-Staff.m` - Staff member lookup data

### Why 3 Queries Instead of 1?

✅ **Performance**: No data duplication - Services/Staff loaded once, not repeated per appointment  
✅ **Relationships**: Proper many-to-one relationships enable better filtering and DAX measures  
✅ **Reusability**: Analyze Services or Staff independently without loading all appointments  
✅ **Scalability**: As data grows, this architecture remains fast and efficient  
✅ **Best Practice**: Industry-standard Star Schema pattern used by professional BI developers

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

## 📋 Available Queries

### 1. **PowerQuery-BookingsData.m** (Main Query - Appointments)
The primary query that retrieves all appointments from all booking businesses in a flattened table format.

**Output Columns:**
- Business details (BusinessId, BusinessName, BusinessEmail, BusinessPhone)
- Appointment details (AppointmentId, StartDateTime, EndDateTime, Duration)
- Customer information (CustomerName, CustomerEmail, CustomerPhone)
- Service information (ServiceId, ServiceName)
- Meeting details (IsLocationOnline, OnlineMeetingUrl)
- Calculated fields (AppointmentDate, DayOfWeek, Month, Year)

### 2. **PowerQuery-Services.m** (Services Lookup Table)
Retrieves all services from all booking businesses.

**Output Columns:**
- BusinessId, BusinessName
- ServiceId, ServiceName
- Duration, Price, PriceType
- Description, IsHidden, MaxAttendees

### 3. **PowerQuery-Staff.m** (Staff Lookup Table)
Retrieves all staff members from all booking businesses.

**Output Columns:**
- BusinessId, BusinessName
- StaffId, StaffName, StaffEmail
- Role, TimeZone
- UseBusinessHours, EmailNotificationEnabled

---

## 🚀 How to Use in Power BI

### Step 1: Import the Main Appointments Query

1. Open **Power BI Desktop**
2. Click **Home** > **Get Data** > **Blank Query**
3. In Power Query Editor, click **Home** > **Advanced Editor**
4. **Delete** all existing code in the editor
5. Open `PowerQuery-BookingsData.m` in a text editor
6. **Copy the entire contents** and paste into the Advanced Editor
7. Click **Done**
8. Rename the query to **"Bookings_Appointments"**
9. Click **Close & Apply**

### Step 2: Import Services and Staff (Recommended for Star Schema)

Repeat Step 1 for:
- `PowerQuery-Services.m` → Name it **"Bookings_Services"**
- `PowerQuery-Staff.m` → Name it **"Bookings_Staff"**

### Step 3: Create Relationships (Power BI - Critical for Star Schema)

1. Go to **Model View** (left sidebar icon)
2. **Drag and drop** to create relationships:
   
   **Primary Relationships:**
   - **Bookings_Appointments[ServiceId]** → **Bookings_Services[ServiceId]** (Many-to-One)
   - **Bookings_Appointments[BusinessId]** → **Bookings_Services[BusinessId]** (Many-to-One)
   - **Bookings_Appointments[BusinessId]** → **Bookings_Staff[BusinessId]** (Many-to-One)

3. **Why This Matters:**
   - Enables proper filtering: Select a service → see related appointments
   - DAX measures work correctly: `RELATED()` and `RELATEDTABLE()` functions
   - Cross-filtering: Click appointment → see which staff/service it uses
   - Performance: Power BI uses relationships for optimized queries

**Example DAX Measure (works because of relationships):**
```dax
Total Appointments = COUNTROWS(Bookings_Appointments)
Service Revenue = SUMX(Bookings_Appointments, RELATED(Bookings_Services[ServicePrice]))
```

---

## 📊 How to Use in Excel

### For Excel 365 / Excel 2016+

1. Go to **Data** > **Get Data** > **From Other Sources** > **Blank Query**
2. In Power Query Editor: **Home** > **Advanced Editor**
3. Paste the M query code
4. Click **Done** > **Close & Load**

The data will load as a table in Excel and can be refreshed anytime.

---

## 🔄 Dynamic Features

These queries are designed to be **fully dynamic**:

✅ **Automatically adapts when:**
- New booking businesses are created
- Booking businesses are deleted
- Services are added or removed
- Staff members join or leave
- Appointments are created, updated, or cancelled

✅ **Error Handling:**
- If a business has no appointments → Returns empty rows (no error)
- If API call fails → Returns null instead of breaking the query
- If data structure changes slightly → Uses `try...otherwise` logic

✅ **No Hardcoded IDs:**
- Loops through ALL booking businesses dynamically
- Doesn't rely on specific business IDs existing

---

## 🔒 Security Considerations

### ⚠️ IMPORTANT: These queries contain your app credentials

The queries include:
- **Client ID**: Your Azure AD app's client ID
- **Client Secret**: Your Azure AD app's client secret
- **Tenant ID**: Your organization's Azure AD tenant ID

**IMPORTANT:** Replace the placeholder values in each query file before using:
- `YOUR_TENANT_ID_HERE` → Your actual tenant ID
- `YOUR_CLIENT_ID_HERE` → Your actual client ID
- `YOUR_CLIENT_SECRET_HERE` → Your actual client secret

### Security Best Practices:

1. **DO NOT share your Power BI files publicly** - they contain credentials
2. **DO NOT publish to Power BI Service** without using parameters (see below)
3. **Rotate your client secret periodically** in Azure Portal
4. **Consider using Power BI parameters** for production use (see below)

### 🔐 Production Setup: Using Parameters

For production Power BI reports, use **parameters** instead of hardcoded credentials:

1. In Power Query Editor, create three parameters:
   - `TenantId` (Text)
   - `ClientId` (Text)
   - `ClientSecret` (Text)

2. In the queries, replace:
   ```m
   TenantId = "YOUR_TENANT_ID_HERE",
   ```
   with:
   ```m
   TenantId = TenantId,  // References the parameter
   ```

3. When publishing to Power BI Service, set these parameters in the dataset settings

---

## 📈 Sample Power BI Visuals You Can Create

With these queries, you can easily create:

### Appointments Dashboard
- **Card**: Total Appointments
- **Line Chart**: Appointments over time (by AppointmentDate)
- **Bar Chart**: Appointments by Service
- **Table**: Upcoming appointments
- **Map**: Customer locations (if you add address geocoding)

### Service Performance
- **Bar Chart**: Appointments per Service
- **Pie Chart**: Service distribution
- **Table**: Service details with appointment counts

### Business Overview
- **Matrix**: Appointments by Business and Month
- **KPI Cards**: Total customers, total bookings, average duration

### Customer Analysis
- **Table**: Customer list with appointment history
- **Bar Chart**: Top customers by appointment count
- **Funnel**: Customer acquisition timeline

---

## 🔄 Refreshing Data

### In Power BI Desktop:
- Click **Home** > **Refresh**
- Data refreshes from Microsoft Graph API in real-time

### In Power BI Service (after publishing):
1. Configure scheduled refresh in dataset settings
2. Set refresh frequency (e.g., daily at 8 AM)
3. Ensure gateway is configured if using on-premises data

### In Excel:
- Right-click the table > **Refresh**
- Or: **Data** > **Refresh All**

---

## 🐛 Troubleshooting

### "DataSource.Error: Access Token expired"
- This is normal - the token expires after 1 hour
- Simply refresh the query to get a new token

### "Expression.Error: The key didn't match any rows in the table"
- One of the businesses/services was deleted during query execution
- Query handles this gracefully - just refresh again

### No data returned
- Check that your app has `Bookings.Read.All` permission
- Verify admin consent was granted in Azure Portal
- Confirm you have Microsoft 365 Business Premium license

### Date fields showing as text
- The query converts dates automatically
- If still text, right-click column > **Change Type** > **Date/Time**

---

## 📞 Microsoft Graph API Endpoints Used

These queries call the following Microsoft Graph API endpoints:

```
GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses
GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{id}
GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{id}/appointments
GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{id}/services
GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/{id}/staffMembers
```

**API Documentation**: https://learn.microsoft.com/en-us/graph/api/resources/booking-api-overview

---

## 📝 Query Performance

- **Initial Load**: 5-15 seconds (depending on data volume)
- **Subsequent Refreshes**: 3-10 seconds
- **Data Volume**: Optimized for up to 10,000 appointments

For larger datasets (50,000+ appointments), consider:
- Using incremental refresh in Power BI Service
- Filtering by date range in the M query
- Archiving old appointments

---

## 🆘 Support & Updates

For issues or questions:
1. Check Azure Portal → App Registrations → Your App → API Permissions
2. Verify admin consent is granted
3. Check client secret hasn't expired
4. Review error messages in Power Query Editor

---

## ✅ Quick Checklist

Before using these queries, ensure:

- [ ] Azure AD App Registration created
- [ ] API Permissions granted (`Bookings.Read.All`)
- [ ] Admin consent granted
- [ ] Client secret generated and hasn't expired
- [ ] Credentials updated in the M queries (TenantId, ClientId, ClientSecret)
- [ ] Microsoft 365 Business Premium license active
- [ ] Power BI Desktop or Excel with Power Query installed

---

## 📅 Last Updated

These queries were generated on March 2, 2026 for Electric Ireland Superhomes.

**Credentials Expiry**: Check your client secret expiry date in Azure Portal. Typically 6 months to 2 years from creation.
