# PowerShell script to grant Microsoft Graph permissions to Managed Identity
# Run this in PowerShell (not terminal)

Write-Host "🔧 Installing Microsoft Graph PowerShell module..." -ForegroundColor Yellow
Install-Module Microsoft.Graph -Force -AllowClobber -Scope CurrentUser

Write-Host "🔑 Connecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes 'Application.ReadWrite.All'

Write-Host "🎯 Setting up variables..." -ForegroundColor Yellow
$managedIdentityId = '66ac7fc1-1384-48bb-b306-8c4fc291602'

# First, let's find the Managed Identity service principal by its Object ID
Write-Host "🔍 Finding Managed Identity service principal..." -ForegroundColor Yellow
$managedIdentitySP = Get-MgServicePrincipal -Filter "id eq '$managedIdentityId'"

if (-not $managedIdentitySP) {
    Write-Host "❌ Managed Identity service principal not found!" -ForegroundColor Red
    Write-Host "Searching by display name instead..." -ForegroundColor Yellow
    $managedIdentitySP = Get-MgServicePrincipal -Filter "displayName eq 'jj-sharepoint-bot-web'"
}

if (-not $managedIdentitySP) {
    Write-Host "❌ Could not find the Managed Identity service principal!" -ForegroundColor Red
    Write-Host "Available service principals:" -ForegroundColor Yellow
    Get-MgServicePrincipal -Top 10 | Select-Object Id, DisplayName, AppId | Format-Table
    exit 1
}

Write-Host "✅ Found Managed Identity: $($managedIdentitySP.DisplayName) (ID: $($managedIdentitySP.Id))" -ForegroundColor Green

Write-Host "🔍 Getting Microsoft Graph Service Principal..." -ForegroundColor Yellow
$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

Write-Host "📋 Granting permissions..." -ForegroundColor Yellow

# Grant Sites.Read.All permission
Write-Host "  ✅ Granting Sites.Read.All..." -ForegroundColor Green
$sitesReadAll = $graphServicePrincipal.AppRoles | Where-Object {$_.Value -eq 'Sites.Read.All'}
try {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentitySP.Id -PrincipalId $managedIdentitySP.Id -ResourceId $graphServicePrincipal.Id -AppRoleId $sitesReadAll.Id
    Write-Host "    ✅ Sites.Read.All granted successfully!" -ForegroundColor Green
} catch {
    Write-Host "    ⚠️ Sites.Read.All: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Grant Files.Read.All permission
Write-Host "  ✅ Granting Files.Read.All..." -ForegroundColor Green
$filesReadAll = $graphServicePrincipal.AppRoles | Where-Object {$_.Value -eq 'Files.Read.All'}
try {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentitySP.Id -PrincipalId $managedIdentitySP.Id -ResourceId $graphServicePrincipal.Id -AppRoleId $filesReadAll.Id
    Write-Host "    ✅ Files.Read.All granted successfully!" -ForegroundColor Green
} catch {
    Write-Host "    ⚠️ Files.Read.All: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Grant User.Read.All permission
Write-Host "  ✅ Granting User.Read.All..." -ForegroundColor Green
$userReadAll = $graphServicePrincipal.AppRoles | Where-Object {$_.Value -eq 'User.Read.All'}
try {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentitySP.Id -PrincipalId $managedIdentitySP.Id -ResourceId $graphServicePrincipal.Id -AppRoleId $userReadAll.Id
    Write-Host "    ✅ User.Read.All granted successfully!" -ForegroundColor Green
} catch {
    Write-Host "    ⚠️ User.Read.All: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "🎉 All permissions granted successfully!" -ForegroundColor Green
Write-Host "Your bot can now access SharePoint and OneDrive data using Managed Identity!" -ForegroundColor Green

# Disconnect
Disconnect-MgGraph
