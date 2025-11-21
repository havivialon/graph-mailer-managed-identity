#####################################################
# Check if Microsoft.Graph module is installed
#####################################################
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph module not found. Installing..."
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
} else {
    Write-Host "Microsoft.Graph module is already installed."
}

#####################################################
# Config values
#####################################################
$managedIdentityObjectId = "<ManagedIdentityObjectId>"  
$TenantId = "<TenantId>"  

#####################################################
# Connect to the MS Graph
#####################################################
Connect-MgGraph -TenantId $TenantId -Scopes 'Application.Read.All','AppRoleAssignment.ReadWrite.All'

#####################################################
# Microsoft Graph App Role Assignments
#####################################################
[string] $ServerApplicationNameGraph = "Microsoft Graph"
[string[]]$GraphAppRoleNames = @("Mail.Send", "User.Read.All")
$serverServicePrincipalGraph = (Get-MgServicePrincipal -Filter "DisplayName eq '$ServerApplicationNameGraph'")
$serverServicePrincipalGraphObjectId = $serverServicePrincipalGraph.Id

foreach ($appRoleName in $GraphAppRoleNames) {
    $appRoleId = ($serverServicePrincipalGraph.AppRoles | Where-Object { $_.Value -eq $appRoleName }).Id

    # Assign the managed identity access to the app role.
    try {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $managedIdentityObjectId `
            -PrincipalId $managedIdentityObjectId `
            -ResourceId $serverServicePrincipalGraphObjectId `
            -AppRoleId $appRoleId
    }
    catch {
        Write-Host "Failed to assign role $appRoleName. It may already be assigned."
    }
}