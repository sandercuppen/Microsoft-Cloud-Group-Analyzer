# Check for required modules
$requiredModules = @(
    "Microsoft.Graph.Applications",
    "Microsoft.Graph.Authentication"
)

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Required module '$module' is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force
    }
}

# Application configuration
$appName = "Microsoft Cloud Group Analyzer"

# Required Microsoft Graph permissions for the app
$requiredPermissions = @(
    @{
        resourceAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
        permissions = @(
            "Directory.Read.All",
            "Policy.Read.All",
            "DeviceManagementApps.Read.All",
            "DeviceManagementConfiguration.Read.All",
            "DeviceManagementManagedDevices.Read.All",
            "DeviceManagementServiceConfig.Read.All",
            "User.Read.All",
            "EntitlementManagement.Read.All"
        )
    },
    @{
        resourceAppId = "797f4846-ba00-4fd7-ba43-dac1f8f63013" # Azure Service Management
        permissions = @(
            "user_impersonation"
        )
    }
)

try {
    # Connect to Microsoft Graph with Admin privileges
    Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Directory.Read.All"
    
    # Check if app registration already exists
    $existingApp = Get-MgApplication -Filter "DisplayName eq '$appName'"
    
    if ($existingApp) {
        Write-Host "App registration '$appName' already exists." -ForegroundColor Yellow
        Write-Host "Application ID: $($existingApp.AppId)" -ForegroundColor Yellow
        
        $confirmation = Read-Host "Do you want to delete the existing app registration and create a new one? (y/n)"
        if ($confirmation -eq 'y') {
            Write-Host "Removing existing app registration..." -ForegroundColor Yellow
            Remove-MgApplication -ApplicationId $existingApp.Id
            Write-Host "Existing app registration removed." -ForegroundColor Green
        } else {
            Write-Host "Operation cancelled. Existing app registration was kept." -ForegroundColor Yellow
            return
        }
    }
    
    Write-Host "Creating app registration..." -ForegroundColor Green
    $app = New-MgApplication -DisplayName $appName -SignInAudience "AzureADMyOrg"
    
    # Create service principal
    $sp = New-MgServicePrincipal -AppId $app.AppId
    
    # Get Microsoft Graph service principal
    $graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ConsistencyLevel eventual
    $azureServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '797f4846-ba00-4fd7-ba43-dac1f8f63013'" -ConsistencyLevel eventual

    # Create required permissions
    foreach ($permission in $requiredPermissions) {
        $targetSp = if ($permission.resourceAppId -eq "00000003-0000-0000-c000-000000000000") {
            $graphServicePrincipal
        } else {
            $azureServicePrincipal
        }
        
        foreach ($p in $permission.permissions) {
            $appRole = $targetSp.AppRoles | Where-Object { $_.Value -eq $p }
            if ($appRole) {
                Write-Host "Granting permission: $p" -ForegroundColor Green
                $params = @{
                    PrincipalId = $sp.Id
                    ResourceId = $targetSp.Id
                    AppRoleId = $appRole.Id
                }
                
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter $params
            } else {
                Write-Host "Warning: Could not find role '$p' in service principal" -ForegroundColor Yellow
            }
        }
    }
    
    # Get tenant ID
    $tenantId = (Get-MgContext).TenantId
    
    # Output instructions for manual role assignment
    Write-Host @"
To assign the Reader role at the root management group level, run these Azure CLI commands:

1. Sign in to Azure CLI:
   az login

2. Assign Reader role at root management group:
   az role assignment create --assignee $($sp.Id) --role Reader --scope /providers/Microsoft.Management/managementGroups/$tenantId

Note: You need to have sufficient permissions at the root management group level to perform this operation.
If you encounter permission issues, please contact your Azure administrator.

Application Details for reference:
- Application (client) ID: $($app.AppId)
- Directory (tenant) ID: $tenantId
- Service Principal ID: $($sp.Id)
"@ -ForegroundColor Green

    # Create client secret
    $endDateTime = (Get-Date).AddYears(1)
    $passwordCred = @{
        displayName = "Auto-generated secret"
        endDateTime = $endDateTime
    }
    $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCred
    
    # Create .env file
    $envContent = @"
CLIENTID=$($app.AppId)
TENANTID=$tenantId
CLIENTSECRET=$($secret.SecretText)
"@
    
    $envPath = Join-Path $PSScriptRoot ".." ".env"
    Set-Content -Path $envPath -Value $envContent
    
    Write-Host "`nApp registration completed successfully!" -ForegroundColor Green
    Write-Host "Configuration saved to .env" -ForegroundColor Green
    Write-Host "`nApp Details:" -ForegroundColor Yellow
    Write-Host "Tenant ID: $tenantId" -ForegroundColor Yellow
    Write-Host "Client ID: $($app.AppId)" -ForegroundColor Yellow
    Write-Host "Client Secret has been saved to .env" -ForegroundColor Yellow
    
} catch {
    Write-Host "Error occurred: $_" -ForegroundColor Red
} finally {
    Disconnect-MgGraph
}
