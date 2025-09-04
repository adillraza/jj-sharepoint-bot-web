# First, get an access token for Microsoft Graph
# You'll need to run this in your terminal:

# 1. Get access token (this will open browser for login)
az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv

# 2. Copy the token and use it in the curl commands below
# Replace YOUR_ACCESS_TOKEN with the actual token

# Your Managed Identity Object ID
MANAGED_IDENTITY_ID='66ac7fc1-1384-48bb-b306-8c4fc291602'

# Microsoft Graph Service Principal ID (get this first)
curl -X GET 'https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '\''00000003-0000-0000-c000-000000000000'\''' \
  -H 'Authorization: Bearer YOUR_ACCESS_TOKEN' \
  -H 'Content-Type: application/json'

# Then use the returned 'id' field in the permission assignments below
# Replace GRAPH_SERVICE_PRINCIPAL_ID with the actual ID

# Grant Sites.Read.All permission
curl -X POST 'https://graph.microsoft.com/v1.0/servicePrincipals/GRAPH_SERVICE_PRINCIPAL_ID/appRoleAssignments' \
  -H 'Authorization: Bearer YOUR_ACCESS_TOKEN' \
  -H 'Content-Type: application/json' \
  -d '{
    "principalId": "66ac7fc1-1384-48bb-b306-8c4fc291602",
    "resourceId": "GRAPH_SERVICE_PRINCIPAL_ID",
    "appRoleId": "332a536c-c7ef-4017-ab91-336970924f0d"
  }'

# Grant Files.Read.All permission  
curl -X POST 'https://graph.microsoft.com/v1.0/servicePrincipals/GRAPH_SERVICE_PRINCIPAL_ID/appRoleAssignments' \
  -H 'Authorization: Bearer YOUR_ACCESS_TOKEN' \
  -H 'Content-Type: application/json' \
  -d '{
    "principalId": "66ac7fc1-1384-48bb-b306-8c4fc291602",
    "resourceId": "GRAPH_SERVICE_PRINCIPAL_ID",
    "appRoleId": "75359482-378d-4052-8f01-80520e7db3cd"
  }'

echo 'This is complex - PowerShell is much easier!'
