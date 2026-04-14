param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Setup", "Teardown")]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [string]$ObjectId,

    [Parameter(Mandatory = $true)]
    [string]$DisplayName,

    [Parameter(Mandatory = $true)]
    [string]$MailboxAddress,

    [Parameter(Mandatory = $true)]
    [string]$ScopeName,

    [switch]$RemoveServicePrincipal
)

Write-Output "===== Exchange RBAC Mail.Send - $Mode Mode ====="
Write-Output "App ID          : $AppId"
Write-Output "Object ID       : $ObjectId"
Write-Output "Display Name    : $DisplayName"
Write-Output "Mailbox         : $MailboxAddress"
Write-Output "Scope Name      : $ScopeName"
Write-Output ""

if ($Mode -eq "Setup") {

    # Step 1 - Check if service principal already exists, create if not
    Write-Output "Checking for existing Exchange service principal..."
    $existingSP = Get-ServicePrincipal | Where-Object { $_.AppId -eq $AppId }

    if ($existingSP) {
        Write-Output "Service principal already exists - skipping creation."
    }
    else {
        Write-Output "Service principal not found - creating..."
        try {
            New-ServicePrincipal -AppId $AppId -ObjectId $ObjectId -DisplayName $DisplayName -ErrorAction Stop
            Write-Output "Service principal created."
        }
        catch {
            Write-Output "Failed to create service principal: $($_.Exception.Message)"
            throw
        }
    }

    # Step 2 - Create management scope
    Write-Output ""
    Write-Output "Creating management scope '$ScopeName' for $MailboxAddress..."
    try {
        New-ManagementScope -Name $ScopeName `
            -RecipientRestrictionFilter "PrimarySmtpAddress -eq '$MailboxAddress'" `
            -ErrorAction Stop
        Write-Output "Management scope created."
    }
    catch {
        Write-Output "Failed to create management scope: $($_.Exception.Message)"
        throw
    }

    # Step 3 - Assign Mail.Send role within scope
    Write-Output ""
    Write-Output "Assigning Application Mail.Send role within scope '$ScopeName'..."
    try {
        New-ManagementRoleAssignment -App $AppId `
            -Role "Application Mail.Send" `
            -CustomResourceScope $ScopeName `
            -ErrorAction Stop
        Write-Output "Role assignment created."
    }
    catch {
        Write-Output "Failed to create role assignment: $($_.Exception.Message)"
        throw
    }

    # Step 4 - Verify
    Write-Output ""
    Write-Output "Verifying authorization..."
    try {
        $result = Test-ServicePrincipalAuthorization -Identity $AppId -Resource $MailboxAddress -ErrorAction Stop
        $result | Format-Table RoleName, GrantedPermissions, AllowedResourceScope, ScopeType, InScope -AutoSize

        if ($result.InScope -contains $true) {
            Write-Output "Verification passed - app is authorized to send from $MailboxAddress."
        }
        else {
            Write-Output "Verification returned InScope: False - review the configuration above."
        }
    }
    catch {
        Write-Output "Verification failed: $($_.Exception.Message)"
    }
}

elseif ($Mode -eq "Teardown") {

    # Step 1 - Find and remove the role assignment for this scope
    Write-Output "Looking up role assignment for scope '$ScopeName'..."
    try {
        $assignment = Get-ManagementRoleAssignment -App $AppId | Where-Object { $_.CustomResourceScope -eq $ScopeName }

        if ($assignment) {
            Write-Output "Found assignment: $($assignment.Name) - removing..."
            Remove-ManagementRoleAssignment -Identity $assignment.Name -Confirm:$false -ErrorAction Stop
            Write-Output "Role assignment removed."
        }
        else {
            Write-Output "No role assignment found for scope '$ScopeName' on this app - skipping."
        }
    }
    catch {
        Write-Output "Failed to remove role assignment: $($_.Exception.Message)"
        throw
    }

    # Step 2 - Remove the management scope
    Write-Output ""
    Write-Output "Removing management scope '$ScopeName'..."
    try {
        $scope = Get-ManagementScope -Identity $ScopeName -ErrorAction SilentlyContinue
        if ($scope) {
            Remove-ManagementScope -Identity $ScopeName -Confirm:$false -ErrorAction Stop
            Write-Output "Management scope removed."
        }
        else {
            Write-Output "Management scope '$ScopeName' not found - skipping."
        }
    }
    catch {
        Write-Output "Failed to remove management scope: $($_.Exception.Message)"
        throw
    }

    # Step 3 - Optionally remove the service principal
    if ($RemoveServicePrincipal) {
        Write-Output ""
        Write-Output "Checking for remaining role assignments on this service principal..."
        $remainingAssignments = Get-ManagementRoleAssignment -App $AppId

        if ($remainingAssignments) {
            Write-Output "The following role assignments still exist for this app - service principal will not be removed:"
            $remainingAssignments | Format-Table Name, Role, CustomResourceScope -AutoSize
            Write-Output "Remove remaining assignments first, then re-run with -RemoveServicePrincipal."
        }
        else {
            Write-Output "No remaining assignments found - removing service principal..."
            try {
                Remove-ServicePrincipal -Identity $ObjectId -Confirm:$false -ErrorAction Stop
                Write-Output "Service principal removed."
            }
            catch {
                Write-Output "Failed to remove service principal: $($_.Exception.Message)"
                throw
            }
        }
    }
    else {
        Write-Output ""
        Write-Output "Service principal retained. Use -RemoveServicePrincipal to remove it if fully decommissioning this app from Exchange."
    }

    Write-Output ""
    Write-Output "===== Teardown Complete ====="
}
