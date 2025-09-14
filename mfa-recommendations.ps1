<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

.SYNOPSIS
    Generates migration recommendations based on assessment data
#>

param(
    [Parameter(Mandatory=$true)]
    [hashtable]$AssessmentData
)

Write-Host "`n=== MIGRATION RECOMMENDATIONS ===" -ForegroundColor Cyan

$recommendations = @{
    MethodsToEnable = @()
    MethodsToDisable = @()
    UsersToMigrate = @()
    SafetyChecks = @()
    Risks = @()
    CriticalActions = @()
}

$CurrentPolicy = $AssessmentData.CurrentPolicy
$LegacyMfaData = $AssessmentData.LegacyMfaData
$CaPolicies = $AssessmentData.CaPolicies

# Define authentication methods categories
$lessSecureMethods = @('Voice/Phone', 'SMS', 'Email')
$modernMethods = @('MicrosoftAuthenticator', 'Fido2', 'WindowsHello', 'X509Certificate')
$acceptableLegacyMethod = 'SoftwareOath' # Hardware tokens are acceptable

# Track users who need attention vs those who can be automatically cleaned up
$usersNeedingAssistance = @()
$usersForAutomaticCleanup = @()

# Analyze current method usage vs policy
foreach ($currentMethod in $LegacyMfaData.CurrentMethodDistribution.Keys) {
    $policyMethod = switch ($currentMethod) {
        'MicrosoftAuthenticator' { 'MicrosoftAuthenticator' }
        'Voice/Phone' { 'Voice' }
        'SMS' { 'Sms' }
        'Email' { 'Email' }
        'Fido2' { 'Fido2' }
        'SoftwareOath' { 'SoftwareOath' }
        'TemporaryAccessPass' { 'TemporaryAccessPass' }
        'X509Certificate' { 'X509Certificate' }
        'WindowsHello' { 'WindowsHello' }
        default { $null }
    }
    
    if ($policyMethod) {
        $userCount = $LegacyMfaData.CurrentMethodDistribution[$currentMethod]
        
        # Check what needs to be enabled for Phase 1 (meet deadline)
        if ($CurrentPolicy[$policyMethod] -ne "enabled" -and $userCount -gt 0) {
            $message = "Phase 1 Requirement: Enable $currentMethod for $userCount users"
            Write-Host $message -ForegroundColor Yellow
            
            $recommendations.MethodsToEnable += @{
                Method = $policyMethod
                CurrentMethod = $currentMethod
                UserCount = $userCount
                Priority = "PHASE1"
                Action = "Enable in Authentication Methods Policy before migration"
            }
        }
        else {
            $message = "OK: $currentMethod already enabled for $userCount users"
            Write-Host $message -ForegroundColor Green
        }
    }
}

# Phase 2 Analysis: Categorize users for security improvement
Write-Host "`n=== PHASE 2 PLANNING: USER CATEGORIZATION ===" -ForegroundColor Cyan

foreach ($upn in $LegacyMfaData.UserMfaStatus.Keys) {
    $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
    $methods = $userStatus.Methods -split ", "
    
    $hasLessSecure = $false
    $hasModernMethod = $false
    $userLessSecureMethods = @()
    $userModernMethods = @()
    
    foreach ($method in $methods) {
        if ($method -in @('Phone', 'SMS', 'Email')) {
            $hasLessSecure = $true
            $userLessSecureMethods += $method
        }
        elseif ($method -in @('Authenticator', 'FIDO2', 'Windows', 'Certificate', 'SoftwareOath')) {
            $hasModernMethod = $true
            $userModernMethods += $method
        }
    }
    
    # Categorize users
    if ($hasLessSecure -and $hasModernMethod) {
        # These users can have insecure methods removed automatically
        $usersForAutomaticCleanup += [PSCustomObject]@{
            UserPrincipalName = $upn
            SecureMethods = $userModernMethods -join ", "
            MethodsToRemove = $userLessSecureMethods -join ", "
            Action = "Automatic Cleanup"
        }
    }
    elseif ($hasLessSecure -and -not $hasModernMethod -and $methods.Count -gt 1) {
        # These users need assistance to register secure methods
        $usersNeedingAssistance += [PSCustomObject]@{
            UserPrincipalName = $upn
            CurrentMethods = $methods -join ", "
            LessSecureMethods = $userLessSecureMethods -join ", "
            Priority = "HIGH"
            Action = "Needs Assistance"
        }
    }
}

$totalUsersForCleanup = $usersForAutomaticCleanup.Count
$totalUsersNeedingHelp = $usersNeedingAssistance.Count

Write-Host "Phase 2 User Analysis:" -ForegroundColor Yellow
Write-Host "  - Users for automatic cleanup: $totalUsersForCleanup" -ForegroundColor Green
Write-Host "    (Have both secure and insecure methods - can remove insecure automatically)" -ForegroundColor Gray
Write-Host "  - Users needing assistance: $totalUsersNeedingHelp" -ForegroundColor Yellow
Write-Host "    (Only have insecure methods - need help registering secure alternatives)" -ForegroundColor Gray

# Store categorized users for reporting
$recommendations.UsersForAutomaticCleanup = $usersForAutomaticCleanup
$recommendations.UsersNeedingAssistance = $usersNeedingAssistance

# Analyze privileged users for Phase 2
$privilegedUsersNeedingFido = @()
if ($PrivilegedData -and $PrivilegedData.AnalyzedPrivilegedUsers) {
    foreach ($privUser in $PrivilegedData.AnalyzedPrivilegedUsers) {
        $userStatus = $LegacyMfaData.UserMfaStatus[$privUser.UserPrincipalName]
        $methods = $userStatus.Methods
        
        $hasFIDO2 = $methods -match "FIDO2"
        $isBreakGlass = $privUser.UserPrincipalName -in @('CyberPerformancePack-BTG@globalmicrosolutions.onmicrosoft.com')
        
        if (-not $hasFIDO2 -and -not $isBreakGlass) {
            $privilegedUsersNeedingFido += [PSCustomObject]@{
                UserPrincipalName = $privUser.UserPrincipalName
                RoleName = $privUser.RoleName
                CurrentMethods = $methods
                Priority = "CRITICAL"
            }
        }
    }
}

$recommendations.PrivilegedUsersNeedingFido = $privilegedUsersNeedingFido

# Summary recommendations
Write-Host "`n=== TWO-PHASE MIGRATION APPROACH ===" -ForegroundColor Cyan
Write-Host "PHASE 1 (By September 30, 2025): Meet Microsoft Deadline" -ForegroundColor Yellow
Write-Host "- Enable ALL currently used authentication methods" -ForegroundColor White
Write-Host "- Migrate to Authentication Methods Policy" -ForegroundColor White
Write-Host "- Ensure zero disruption to users" -ForegroundColor White

Write-Host "`nPHASE 2 (4-6 weeks after Phase 1): Security Enhancement" -ForegroundColor Yellow
Write-Host "- Automatically remove insecure methods for $totalUsersForCleanup users" -ForegroundColor White
Write-Host "- Assist $totalUsersNeedingHelp users with registering secure methods" -ForegroundColor White
Write-Host "- Deploy FIDO2 keys to $($privilegedUsersNeedingFido.Count) privileged users" -ForegroundColor White
Write-Host "- Disable Voice, SMS, and Email authentication methods" -ForegroundColor White

return $recommendations
