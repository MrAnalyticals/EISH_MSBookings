# Microsoft Bookings API Integration

## Power Automate Error Analysis

### The Problem: AADSTS50013

The error you're seeing in Power Automate is:

```
AADSTS50013: Assertion failed signature validation.
[Reason - Key was found, but use of the key to verify the signature failed.]
```

### What This Means

This is a **certificate validation error** that occurs when:

1. **The Power Automate connector is trying to authenticate** with Microsoft Bookings on your behalf
2. **The signing key/certificate doesn't match** what's expected
3. **Common causes:**
   - The connector's service principal has certificate/key management issues
   - Token validation is failing due to key rotation issues
   - There may be a SAML SSO configuration conflict with OAuth2

### Solutions

#### Option 1: Refresh the Connection (Quickest Fix)

1. In Power Automate, go to **Data** → **Connections**
2. Find your Microsoft Bookings connection
3. **Delete** the existing connection
4. **Recreate** the connection with fresh authentication
5. Test again

#### Option 2: Use Microsoft Graph API Directly

Instead of the Bookings connector, use the **HTTP with Azure AD** connector to call Microsoft Graph APIs directly. This gives you more control and avoids connector-specific issues.

#### Option 3: Create a Custom Connector

Build a custom Power Automate connector that uses your own app registration with proper certificates and permissions.

---

## Direct Microsoft Graph API Integration

### Prerequisites

1. **Azure AD App Registration** with these permissions:
   - `Bookings.Read.All` (least privilege for reading)
   - `BookingsAppointment.ReadWrite.All` (for appointments)
   - `Bookings.ReadWrite.All` (for full access)

2. **Microsoft 365 Business Premium** license (required for Bookings)

3. **Admin consent** granted for the permissions

### Your Bookings Instance

Based on your booking page URL:
```
https://outlook.office.com/book/EnergyReportCallWithElectricIrelandSuperhomes@electricirelandsuperhomes.ie/
```

Your bookings business ID is likely:
- **Display Name**: "Energy Report Call With Electric Ireland Superhomes"
- **Email/ID**: `EnergyReportCallWithElectricIrelandSuperhomes@electricirelandsuperhomes.ie`

---

## API Endpoints

### Base URL
```
https://graph.microsoft.com/v1.0/solutions/bookingBusinesses
```

### Key Operations

| Operation | Endpoint | Method |
|-----------|----------|--------|
| List all businesses | `/solutions/bookingBusinesses` | GET |
| Get specific business | `/solutions/bookingBusinesses/{id}` | GET |
| List appointments | `/solutions/bookingBusinesses/{id}/appointments` | GET |
| Create appointment | `/solutions/bookingBusinesses/{id}/appointments` | POST |
| List services | `/solutions/bookingBusinesses/{id}/services` | GET |
| List staff | `/solutions/bookingBusinesses/{id}/staffMembers` | GET |
| Publish booking page | `/solutions/bookingBusinesses/{id}/publish` | POST |

---

## Testing Connectivity

### Step 1: Get an Access Token

You need to authenticate and get a token. See the PowerShell script: `Test-MSBookingsConnection.ps1`

### Step 2: Test Basic Connectivity

Run the PowerShell script to:
1. Authenticate to Microsoft Graph
2. List all booking businesses in your tenant
3. Get details about your specific business
4. List appointments

---

## Common Permissions

### Delegated Permissions (User Context)
- `Bookings.Read.All` - Read all bookings
- `BookingsAppointment.ReadWrite.All` - Manage appointments
- `Bookings.ReadWrite.All` - Full read/write access
- `Bookings.Manage.All` - Full management access

### Application Permissions (App-Only Context)
- `Bookings.Read.All` - Read all bookings
- `BookingsAppointment.ReadWrite.All` - Manage appointments
- `Bookings.ReadWrite.All` - Full read/write access

---

## Next Steps

1. **Run the PowerShell script** to test connectivity
2. **Review the sample C# application** for building a full app
3. **Consider using Microsoft Graph SDK** for easier development
4. **Set up proper error handling** for production use

---

## Resources

- [Microsoft Bookings API Overview](https://learn.microsoft.com/graph/api/resources/booking-api-overview)
- [Microsoft Graph Permissions Reference](https://learn.microsoft.com/graph/permissions-reference)
- [Troubleshooting Signature Validation Errors](https://learn.microsoft.com/troubleshoot/entra/entra-id/app-integration/troubleshooting-signature-validation-errors)
