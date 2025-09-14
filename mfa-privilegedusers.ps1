<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

.SYNOPSIS
    Analyses privileged user MFA status - Fixed for single expand limitation
#>

param(
    [Parameter(Mandatory=$true)]
    [hashtable]$LegacyMfaData
)

Write-Host "`n=== CHECKING PRIVILEGED ROLE ASSIGNMENTS ===" -ForegroundColor Cyan

$privilegedRoles = @()
$additionalPrivilegedAccounts = @()
$privilegedNoMfa = 0
$privilegedWeakMfa = 0

# Define high-privilege roles
$highPrivilegeRoles = @(
    'Global Administrator',
    'Privileged Role Administrator', 
    'Security Administrator',
    'Exchange Administrator',
    'SharePoint Administrator',
    'User Administrator',
    'Helpdesk Administrator',
    'Authentication Administrator',
    'Privileged Authentication Administrator',
    'Cloud Application Administrator',
    'Application Administrator',
    'Conditional Access Administrator',
    'Azure AD Joined Device Local Administrator',
    'Directory Writers',
    'Partner Tier2 Support',
    'Company Administrator',
    'Global Reader',
    'Security Reader',
    'Compliance Administrator',
    'Password Administrator'
)

try {
    # Get role assignments with roleDefinition expansion only
    Write-Host "Retrieving privileged role assignments..." -ForegroundColor Yellow
    
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$expand=roleDefinition"
    $roleResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
    $roleAssignments = $roleResponse.value
    
    # Continue getting all pages of results
    while ($roleResponse.'@odata.nextLink') {
        $roleResponse = Invoke-MgGraphRequest -Uri $roleResponse.'@odata.nextLink' -Method GET
        $roleAssignments += $roleResponse.value
    }
    
    Write-Host "Found $($roleAssignments.Count) total role assignments" -ForegroundColor Gray
    
    # Process role assignments and collect principal IDs
    $privilegedPrincipalIds = @{}
    
    foreach ($assignment in $roleAssignments) {
        $roleName = $assignment.roleDefinition.displayName
        
        if ($roleName -in $highPrivilegeRoles) {
            $principalId = $assignment.principalId
            
            if ($principalId) {
                if (-not $privilegedPrincipalIds.ContainsKey($principalId)) {
                    $privilegedPrincipalIds[$principalId] = @{
                        Roles = @()
                        RoleTypes = @()
                    }
                }
                
                $privilegedPrincipalIds[$principalId].Roles += $roleName
                $privilegedPrincipalIds[$principalId].RoleTypes += "Direct"
            }
        }
    }
    
    # Check PIM eligible roles
    Write-Host "Checking PIM eligible roles..." -ForegroundColor Yellow
    try {
        $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules?`$expand=roleDefinition"
        $eligibleResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction SilentlyContinue
        
        if ($eligibleResponse.value) {
            $eligibleRoles = $eligibleResponse.value
            
            # Get all pages
            while ($eligibleResponse.'@odata.nextLink') {
                $eligibleResponse = Invoke-MgGraphRequest -Uri $eligibleResponse.'@odata.nextLink' -Method GET
                $eligibleRoles += $eligibleResponse.value
            }
            
            foreach ($eligibility in $eligibleRoles) {
                $roleName = $eligibility.roleDefinition.displayName
                
                if ($roleName -in $highPrivilegeRoles) {
                    $principalId = $eligibility.principalId
                    
                    if ($principalId) {
                        if (-not $privilegedPrincipalIds.ContainsKey($principalId)) {
                            $privilegedPrincipalIds[$principalId] = @{
                                Roles = @()
                                RoleTypes = @()
                            }
                        }
                        
                        $privilegedPrincipalIds[$principalId].Roles += $roleName
                        $privilegedPrincipalIds[$principalId].RoleTypes += "PIM Eligible"
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Could not check PIM roles (may not be configured)" -ForegroundColor Yellow
    }
    
    # Now resolve principal IDs to users
    Write-Host "Resolving privileged user details..." -ForegroundColor Yellow
    $privilegedUserMap = @{}
    
    foreach ($principalId in $privilegedPrincipalIds.Keys) {
        try {
            # Get user details
            $uri = "https://graph.microsoft.com/v1.0/users/$principalId"
            $userResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction SilentlyContinue
            
            if ($userResponse.userPrincipalName) {
                $upn = $userResponse.userPrincipalName
                $privilegedUserMap[$upn] = $privilegedPrincipalIds[$principalId]
            }
        }
        catch {
            # Skip if user not found or is a service principal
        }
    }
    
    # Check MFA status for privileged users
    Write-Host "Analyzing MFA status for privileged users..." -ForegroundColor Yellow
    
    foreach ($upn in $privilegedUserMap.Keys) {
        $userData = $privilegedUserMap[$upn]
        
        # Create privileged user object
        $privUser = [PSCustomObject]@{
            UserPrincipalName = $upn
            RoleName = ($userData.Roles | Select-Object -Unique) -join ", "
            RoleType = ($userData.RoleTypes | Select-Object -Unique) -join ", "
        }
        
        # Check if user is in our analyzed list
        if ($LegacyMfaData.UserMfaStatus.ContainsKey($upn)) {
            $privilegedRoles += $privUser
            
            $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
            
            # Check for users without MFA
            if ($userStatus.Status -eq "Password Only - Needs MFA" -or 
                $userStatus.Status -eq "No Authentication Methods") {
                $privilegedNoMfa++
            }
            
            # Check for users without phishing-resistant MFA
            $methods = $userStatus.Methods
            $hasFIDO2 = $methods -match "FIDO2"
            $hasCertificate = $methods -match "Certificate"
            $hasWindowsHello = $methods -match "Windows"
            
            # Exclude break-glass accounts
            $isBreakGlass = $upn -in @('CyberPerformancePack-BTG@globalmicrosolutions.onmicrosoft.com')
            
            if (-not ($hasFIDO2 -or $hasCertificate -or $hasWindowsHello) -and -not $isBreakGlass) {
                $privilegedWeakMfa++
            }
        }
        else {
            # Privileged user not in our analyzed list
            $additionalPrivilegedAccounts += [PSCustomObject]@{
                UserPrincipalName = $upn
                RoleName = $privUser.RoleName
                RoleType = $privUser.RoleType
                Note = "Not in analyzed user list"
            }
        }
    }
    
    # Display results
    if ($privilegedRoles.Count -gt 0) {
        Write-Host "`n=== PRIVILEGED USERS IN TENANT ===" -ForegroundColor Red
        Write-Host "Found $($privilegedRoles.Count) privileged users" -ForegroundColor Yellow
        
        # Group by role
        $roleGroups = $privilegedRoles | Group-Object RoleName
        foreach ($group in $roleGroups | Sort-Object Count -Descending) {
            Write-Host "  $($group.Name): $($group.Count) users" -ForegroundColor Yellow
        }
        
        # Show summary of privileged users without MFA
        if ($privilegedNoMfa -gt 0) {
            Write-Host "`n*** CRITICAL SECURITY ALERT ***" -ForegroundColor Red
            Write-Host "$privilegedNoMfa PRIVILEGED USERS HAVE NO MFA!" -ForegroundColor Red
        }
        
        if ($privilegedWeakMfa -gt 0) {
            Write-Host "`n*** CRITICAL: PRIVILEGED USERS WITHOUT PHISHING-RESISTANT MFA ***" -ForegroundColor Red
            Write-Host "$privilegedWeakMfa privileged users lack FIDO2/Certificate/Windows Hello!" -ForegroundColor Red
            Write-Host "RECOMMENDATION: Deploy FIDO2 security keys to all privileged administrators immediately!" -ForegroundColor Yellow
        }
    }
    
    if ($additionalPrivilegedAccounts.Count -gt 0) {
        Write-Host "`n=== ADDITIONAL PRIVILEGED ACCOUNTS (Not in user analysis) ===" -ForegroundColor Yellow
        Write-Host "Found $($additionalPrivilegedAccounts.Count) additional privileged accounts" -ForegroundColor Yellow
    }
    
    # Final summary
    Write-Host "`nPrivileged User Summary:" -ForegroundColor Cyan
    Write-Host "  Analyzed privileged users: $($privilegedRoles.Count)" -ForegroundColor White
    Write-Host "  Additional privileged accounts: $($additionalPrivilegedAccounts.Count)" -ForegroundColor White
    Write-Host "  Total privileged users: $($privilegedRoles.Count + $additionalPrivilegedAccounts.Count)" -ForegroundColor White
    
    return @{
        AnalyzedPrivilegedUsers = $privilegedRoles
        AdditionalPrivilegedAccounts = $additionalPrivilegedAccounts
        PrivilegedNoMfa = $privilegedNoMfa
        PrivilegedWeakMfa = $privilegedWeakMfa
    }
}
catch {
    Write-Warning "Could not check privileged roles: $_"
    return @{
        AnalyzedPrivilegedUsers = @()
        AdditionalPrivilegedAccounts = @()
        PrivilegedNoMfa = 0
        PrivilegedWeakMfa = 0
    }
}
