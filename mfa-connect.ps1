<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

.SYNOPSIS
    Connects to Microsoft Graph for MFA Assessment
#>

try {
    $context = Get-MgContext
    if ($context) {
        Write-Host "Already connected to Microsoft Graph" -ForegroundColor Green
        
        # Get tenant information
        try {
            $org = Get-MgOrganization
            $tenantName = $org.DisplayName
            Write-Host "Tenant Name: $tenantName" -ForegroundColor White
            Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor White
            Write-Host "Account: $($context.Account)" -ForegroundColor White
            
            # Store tenant info in global variable
            $global:MFAAssessmentData.TenantInfo = @{
                TenantName = $tenantName
                TenantId = $context.TenantId
                Account = $context.Account
            }
        }
        catch {
            Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor White
            Write-Host "Account: $($context.Account)" -ForegroundColor White
            
            $global:MFAAssessmentData.TenantInfo = @{
                TenantName = "Unknown"
                TenantId = $context.TenantId
                Account = $context.Account
            }
        }
    }
    else {
        Connect-MgGraph -Scopes @(
            "Policy.ReadWrite.AuthenticationMethod",
            "UserAuthenticationMethod.ReadWrite.All", 
            "Policy.Read.All", 
            "User.Read.All",
            "Reports.Read.All",
            "Directory.Read.All",
            "RoleManagement.Read.Directory",
            "RoleEligibilitySchedule.Read.Directory",
            "RoleAssignmentSchedule.Read.Directory"
        )
        
        # Get tenant information after connecting
        try {
            $context = Get-MgContext
            $org = Get-MgOrganization
            $tenantName = $org.DisplayName
            Write-Host "Connected to Tenant: $tenantName" -ForegroundColor Green
            Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor White
            
            $global:MFAAssessmentData.TenantInfo = @{
                TenantName = $tenantName
                TenantId = $context.TenantId
                Account = $context.Account
            }
        }
        catch {
            Write-Host "Connected to Tenant ID: $($context.TenantId)" -ForegroundColor Green
            
            $global:MFAAssessmentData.TenantInfo = @{
                TenantName = "Unknown"
                TenantId = $context.TenantId
                Account = $context.Account
            }
        }
    }
    
    Write-Host "OK: Connected to Microsoft Graph" -ForegroundColor Green
    return $true
} 
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    return $false
}
