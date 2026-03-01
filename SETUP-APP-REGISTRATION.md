# Azure AD App Registration Setup Guide

This guide helps you create an Azure AD app registration for Microsoft Bookings API access.

## Prerequisites
- Admin access to Azure AD
- Microsoft 365 Business Premium license
- PowerShell or Azure CLI installed

---

## Option 1: PowerShell Setup (Recommended)

### Step 1: Install Required Module
```powershell
Install-Module -Name Az.Accounts -Scope CurrentUser
Install-Module -Name Az.Resources -Scope CurrentUser
```

### Step 2: Connect to Azure
```powershell
Connect-AzAccount
```

### Step 3: Create the App Registration
```powershell
# Set your app name
$appName = "MS Bookings API App"

# Create the app registration
$app = New-AzADApplication -DisplayName $appName

Write-Host "✓ App created: $($app.DisplayName)" -ForegroundColor Green
Write-Host "  App ID (Client ID): $($app.AppId)" -ForegroundColor Cyan
Write-Host "  Object ID: $($app.Id)" -ForegroundColor Gray

# Save these values
$clientId = $app.AppId
$tenantId = (Get-AzContext).Tenant.Id

Write-Host "`nYour Tenant ID: $tenantId" -ForegroundColor Cyan
```

### Step 4: Add Microsoft Graph Permissions
```powershell
# Get Microsoft Graph Service Principal
$graphSP = Get-AzADServicePrincipal -Filter "displayName eq 'Microsoft Graph'"

# Bookings API permissions
$permissions = @(
    @{
        ResourceAppId = $graphSP.AppId
        ResourceAccess = @(
            @{
                # Bookings.Read.All - Delegated
                Id = "33b1df99-4b29-4548-9339-7a7b83eaeebc"
                Type = "Scope"
            },
            @{
                # BookingsAppointment.ReadWrite.All - Delegated
                Id = "02a5a114-36a6-46ff-a102-954d89d9ab02"
                Type = "Scope"
            },
            @{
                # Bookings.ReadWrite.All - Delegated
                Id = "7f36b48e-542f-4d3b-9bcb-8406f0ab9fdb"
                Type = "Scope"
            },
            @{
                # Bookings.Read.All - Application
                Id = "33b1df99-4b29-4548-9339-7a7b83eaeebc"
                Type = "Role"
            },
            @{
                # BookingsAppointment.ReadWrite.All - Application
                Id = "02a5a114-36a6-46ff-a102-954d89d9ab02"
                Type = "Role"
            }
        )
    }
)

# Update the app with required permissions
Update-AzADApplication -ObjectId $app.Id -RequiredResourceAccess $permissions

Write-Host "✓ Permissions added" -ForegroundColor Green
```

### Step 5: Grant Admin Consent
```powershell
# This requires admin privileges
# Go to: https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
# Select your app → API permissions → Grant admin consent

Write-Host "`nIMPORTANT: Grant admin consent in Azure Portal" -ForegroundColor Yellow
Write-Host "1. Go to: https://portal.azure.com" -ForegroundColor Cyan
Write-Host "2. Navigate to Azure AD → App registrations → $appName" -ForegroundColor Cyan
Write-Host "3. Click 'API permissions' → 'Grant admin consent for <tenant>'" -ForegroundColor Cyan
```

### Step 6: Create a Client Secret
```powershell
# Create a client secret (expires in 2 years)
$startDate = Get-Date
$endDate = $startDate.AddYears(2)

$appSecret = New-AzADAppCredential -ObjectId $app.Id -StartDate $startDate -EndDate $endDate

Write-Host "`n✓ Client secret created" -ForegroundColor Green
Write-Host "  Secret Value: $($appSecret.SecretText)" -ForegroundColor Cyan
Write-Host "  Secret ID: $($appSecret.KeyId)" -ForegroundColor Gray
Write-Host "  Expires: $endDate" -ForegroundColor Gray
Write-Host "`n⚠️  SAVE THIS SECRET NOW - You won't be able to see it again!" -ForegroundColor Yellow
```

### Step 7: Configure Reply URLs (for delegated auth)
```powershell
# Add redirect URI for local testing
$replyUrls = @("http://localhost:5000", "http://localhost")
Update-AzADApplication -ObjectId $app.Id -Web -RedirectUri $replyUrls

Write-Host "`n✓ Redirect URIs configured: $($replyUrls -join ', ')" -ForegroundColor Green
```

### Step 8: Save Configuration
```powershell
$config = @{
    TenantId = $tenantId
    ClientId = $clientId
    ClientSecret = $appSecret.SecretText
    AppName = $appName
}

$config | ConvertTo-Json | Out-File "appsettings.json" -Encoding UTF8

Write-Host "`n✓ Configuration saved to appsettings.json" -ForegroundColor Green
Write-Host "`n=== Setup Complete ===" -ForegroundColor Cyan
```

---

## Option 2: Azure Portal Setup (Manual)

### Step 1: Register Application
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Enter:
   - **Name**: `MS Bookings API App`
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: (leave blank for now)
5. Click **Register**

### Step 2: Note Your IDs
After registration, note these values:
- **Application (client) ID**
- **Directory (tenant) ID**

### Step 3: Add API Permissions
1. In your app, go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:
   - `Bookings.Read.All`
   - `BookingsAppointment.ReadWrite.All`
   - `Bookings.ReadWrite.All`
4. For app-only scenarios, also add **Application permissions**:
   - `Bookings.Read.All`
   - `BookingsAppointment.ReadWrite.All`
5. Click **Grant admin consent for [tenant]**

### Step 4: Create Client Secret
1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Enter description: `BookingsAPISecret`
4. Select expiration: 24 months
5. Click **Add**
6. **COPY THE SECRET VALUE NOW** - you won't see it again!

### Step 5: Configure Authentication
1. Go to **Authentication**
2. Under **Platform configurations**, click **Add a platform**
3. Select **Web**
4. Add redirect URI: `http://localhost:5000`
5. Check **ID tokens** and **Access tokens**
6. Click **Configure**

---

## Option 3: Azure CLI Setup

```bash
# Login to Azure
az login

# Create the app registration
az ad app create \
  --display-name "MS Bookings API App" \
  --sign-in-audience AzureADMyOrg

# Get the app ID
APP_ID=$(az ad app list --display-name "MS Bookings API App" --query "[0].appId" -o tsv)
echo "App ID: $APP_ID"

# Get tenant ID
TENANT_ID=$(az account show --query tenantId -o tsv)
echo "Tenant ID: $TENANT_ID"

# Add Microsoft Graph permissions
az ad app permission add \
  --id $APP_ID \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions 33b1df99-4b29-4548-9339-7a7b83eaeebc=Scope \
                     02a5a114-36a6-46ff-a102-954d89d9ab02=Scope \
                     7f36b48e-542f-4d3b-9bcb-8406f0ab9fdb=Scope

# Grant admin consent (requires admin)
az ad app permission admin-consent --id $APP_ID

# Create client secret
az ad app credential reset \
  --id $APP_ID \
  --display-name "BookingsAPISecret" \
  --years 2
```

---

## Verification

After setup, verify your configuration:

```powershell
# Test authentication
Connect-MgGraph -ClientId $clientId -TenantId $tenantId

# Or use the test script
.\Test-MSBookingsConnection.ps1
```

---

## Permission Details

### Bookings.Read.All
- **Type**: Delegated / Application
- **Description**: Read all booking businesses
- **Admin Consent**: Required
- **Use Case**: List businesses, read appointments, services, staff

### BookingsAppointment.ReadWrite.All
- **Type**: Delegated / Application  
- **Description**: Manage booking appointments
- **Admin Consent**: Required
- **Use Case**: Create, update, delete appointments

### Bookings.ReadWrite.All
- **Type**: Delegated / Application
- **Description**: Full access to manage bookings
- **Admin Consent**: Required
- **Use Case**: All operations including business settings

### Bookings.Manage.All
- **Type**: Delegated / Application
- **Description**: Complete administrative access
- **Admin Consent**: Required
- **Use Case**: When you need highest privileges

---

## Security Best Practices

1. **Use Certificates in Production**
   - Instead of client secrets, use certificates
   - More secure and can be automatically rotated

2. **Principle of Least Privilege**
   - Only request permissions your app needs
   - Use `Bookings.Read.All` if you only need read access

3. **Secret Management**
   - Never commit secrets to source control
   - Use Azure Key Vault in production
   - Rotate secrets before expiration

4. **Managed Identity**
   - When running in Azure, use Managed Identity
   - No need to manage credentials

---

## Next Steps

1. Save your configuration to a secure location
2. Update the C# application with your credentials
3. Run the test scripts to verify connectivity
4. Build your integration!

## Troubleshooting

### "Insufficient privileges" error
- Make sure admin consent was granted
- Check that the correct permissions were added

### "Invalid client secret" error
- Verify you copied the secret correctly
- Check if the secret has expired
- Generate a new secret if needed

### "AADSTS50013" error
- This is the signature validation error from Power Automate
- Use direct Graph API calls instead of the connector
- Verify your app registration is configured correctly
