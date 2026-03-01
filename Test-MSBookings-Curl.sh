#!/bin/bash

#############################################
# Microsoft Bookings API - Curl Test Script
#############################################
# This script demonstrates how to call Microsoft Bookings APIs using curl
# You'll need an access token first

echo "=== Microsoft Bookings API Test (curl) ==="
echo ""

# STEP 1: Get an access token
# You can get a token using various methods:
# - Azure CLI: az account get-access-token --resource https://graph.microsoft.com
# - PowerShell: Get-MgAccessToken
# - OAuth2 flow in your app

echo "STEP 1: Get an access token"
echo "Run one of these commands to get a token:"
echo ""
echo "Option A - Using Azure CLI:"
echo "  az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv"
echo ""
echo "Option B - Using PowerShell:"
echo "  Connect-MgGraph -Scopes 'Bookings.Read.All'"
echo "  \$token = (Get-MgContext).AccessToken"
echo ""
read -p "Paste your access token here: " ACCESS_TOKEN

if [ -z "$ACCESS_TOKEN" ]; then
    echo "Error: No token provided"
    exit 1
fi

echo ""
echo "=== Testing Microsoft Bookings API ==="
echo ""

# STEP 2: List all booking businesses
echo "STEP 2: Listing all booking businesses..."
echo "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses"
echo ""

BUSINESSES_RESPONSE=$(curl -s -X GET \
  "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses" \
  -H "Authorization: Bearer $ACCESS_TOKEN" \
  -H "Content-Type: application/json")

echo "Response:"
echo "$BUSINESSES_RESPONSE" | jq '.' 2>/dev/null || echo "$BUSINESSES_RESPONSE"
echo ""

# Extract first business ID
BUSINESS_ID=$(echo "$BUSINESSES_RESPONSE" | jq -r '.value[0].id' 2>/dev/null)

if [ "$BUSINESS_ID" = "null" ] || [ -z "$BUSINESS_ID" ]; then
    echo "No booking businesses found in the response"
    exit 0
fi

echo "Found business ID: $BUSINESS_ID"
echo ""

# STEP 3: Get specific business details
echo "STEP 3: Getting details for business: $BUSINESS_ID"
echo "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID"
echo ""

BUSINESS_DETAILS=$(curl -s -X GET \
  "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID" \
  -H "Authorization: Bearer $ACCESS_TOKEN" \
  -H "Content-Type: application/json")

echo "Response:"
echo "$BUSINESS_DETAILS" | jq '.' 2>/dev/null || echo "$BUSINESS_DETAILS"
echo ""

# STEP 4: List appointments
echo "STEP 4: Listing appointments..."
echo "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID/appointments"
echo ""

APPOINTMENTS=$(curl -s -X GET \
  "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID/appointments" \
  -H "Authorization: Bearer $ACCESS_TOKEN" \
  -H "Content-Type: application/json")

echo "Response:"
echo "$APPOINTMENTS" | jq '.' 2>/dev/null || echo "$APPOINTMENTS"
echo ""

# STEP 5: List services
echo "STEP 5: Listing services..."
echo "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID/services"
echo ""

SERVICES=$(curl -s -X GET \
  "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID/services" \
  -H "Authorization: Bearer $ACCESS_TOKEN" \
  -H "Content-Type: application/json")

echo "Response:"
echo "$SERVICES" | jq '.' 2>/dev/null || echo "$SERVICES"
echo ""

# STEP 6: List staff members
echo "STEP 6: Listing staff members..."
echo "GET https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID/staffMembers"
echo ""

STAFF=$(curl -s -X GET \
  "https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/$BUSINESS_ID/staffMembers" \
  -H "Authorization: Bearer $ACCESS_TOKEN" \
  -H "Content-Type: application/json")

echo "Response:"
echo "$STAFF" | jq '.' 2>/dev/null || echo "$STAFF"
echo ""

echo "=== Test Complete ==="
