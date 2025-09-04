# Azure CLI commands to grant Microsoft Graph permissions

# 1. Login to Azure CLI
az login

# 2. Install Microsoft Graph extension
az extension add --name microsoft-graph

# 3. Your Managed Identity Object ID
MANAGED_IDENTITY_ID='66ac7fc1-1384-48bb-b306-8c4fc291602'

# 4. Microsoft Graph App ID (always the same)
GRAPH_APP_ID='00000003-0000-0000-c000-000000000000'

# 5. Grant permissions
az ad app permission add --id $MANAGED_IDENTITY_ID --api $GRAPH_APP_ID --api-permissions 932f982a-5f07-4679-b5f6-3486ecbc8d97=Role  # Sites.Read.All
az ad app permission add --id $MANAGED_IDENTITY_ID --api $GRAPH_APP_ID --api-permissions df21f32d-9264-437e-a0f0-c1516c57d999=Role  # Files.Read.All  
az ad app permission add --id $MANAGED_IDENTITY_ID --api $GRAPH_APP_ID --api-permissions 14dad69e-099b-42c9-810b-d002981feec1=Role  # profile

# 6. Grant admin consent
az ad app permission admin-consent --id $MANAGED_IDENTITY_ID

echo '🎉 Permissions granted via Azure CLI!'
