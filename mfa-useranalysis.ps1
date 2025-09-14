<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

.SYNOPSIS
    Analyses user MFA registration status using REST API with enhanced tracking fields
    Updated to use Invoke-MgGraphRequest instead of Get-MgUserAuthenticationMethod
#>

Write-Host "`n=== LEGACY MFA ANALYSIS (via Graph API) ===" -ForegroundColor Cyan

$legacyMfaData = @{
    UserMfaStatus = @{}
    MethodDistribution = @{}
    CurrentMethodDistribution = @{}
    TotalUsers = 0
    MfaEnabledUsers = 0
    MfaRegisteredUsers = 0
    CurrentMethodUsers = 0
    NoMfaUsers = @()
    ServiceAccountUsers = @()
    RegularUsersNoMfa = @()
}

# Define service account license patterns
$excludedSkuPatterns = @(
    'MTR*', 
    'PHONESYSTEM_VIRTUALUSER', 
    'MCOPSTN*', 
    'TEAMS_EXPLORATORY', 
    '*KIOSK*', 
    'INTUNE_A_VL',
    'TEAMS_AR_*',
    'TEAMS_PHONE_*',
    '*VIRTUAL*',
    'PHONESYSTEM*',
    'TEAMS_SHARED_DEVICE*',
    'Microsoft_Teams_Rooms_*',
    'TEAMS_ROOM*',
    'MCOEV*'
)

# Service account name patterns
$serviceAccountNamePatterns = @(
    'Attendant-*',
    '*@*.smtp.codetwo.online',
    'Conference*',
    'MTR-*',
    'Teams-*',
    'AutoAttendant*',
    'CallQueue*'
)

try {
    Write-Host "Analyzing user MFA registration status..." -ForegroundColor Yellow
    
    $users = Get-MgUser -All -Property UserPrincipalName,UserType,Id,AssignedLicenses,AccountEnabled | 
        Where-Object {$_.UserType -eq "Member" -and $_.UserPrincipalName -notlike "*#EXT#*" -and $_.Id}
    
    $legacyMfaData.TotalUsers = $users.Count
    Write-Host "Found $($legacyMfaData.TotalUsers) internal users to analyze..." -ForegroundColor White
    
    # Initialize progress tracking
    $processed = 0
    $startTime = Get-Date
    $lastProgressUpdate = Get-Date
    
    # Check if running in VS Code
    $isVSCode = $env:TERM_PROGRAM -eq 'vscode'
    
    foreach ($user in $users) {
        $processed++
        
        # Update progress bar
        $percentComplete = [math]::Round(($processed / $legacyMfaData.TotalUsers) * 100, 1)
        $status = "Analyzing user $processed of $($legacyMfaData.TotalUsers)"
        
        # Calculate estimated time remaining
        if ($processed -gt 10) {
            $elapsed = (Get-Date) - $startTime
            $avgTimePerUser = $elapsed.TotalSeconds / $processed
            $remainingUsers = $legacyMfaData.TotalUsers - $processed
            $estimatedRemaining = [TimeSpan]::FromSeconds($avgTimePerUser * $remainingUsers)
            $status += " - Est. time remaining: $($estimatedRemaining.ToString('mm\:ss'))"
        }
        
        # For VS Code, show console output every 10 users or every 5 seconds
        if ($isVSCode) {
            $timeSinceLastUpdate = (Get-Date) - $lastProgressUpdate
            if ($processed % 10 -eq 0 -or $processed -eq 1 -or $processed -eq $legacyMfaData.TotalUsers -or $timeSinceLastUpdate.TotalSeconds -ge 5) {
                Write-Host "`r$status ($percentComplete% complete)" -NoNewline -ForegroundColor Cyan
                $lastProgressUpdate = Get-Date
            }
        }
        else {
            # Use standard progress bar for PowerShell console/ISE
            Write-Progress -Activity "Analyzing MFA Status" -Status $status -PercentComplete $percentComplete -CurrentOperation $user.UserPrincipalName
        }
        
        try {
            if (-not $user.Id) {
                $legacyMfaData.UserMfaStatus[$user.UserPrincipalName] = @{
                    Status = "Cannot Determine"
                    Methods = "User object missing ID property"
                    Phase2Action = "Review Manually"
                    MigrationGroup = "Error"
                    Phase2Week = "Not applicable"
                    Phase2Priority = "None"
                    SecureMethodCount = 0
                    InsecureMethodCount = 0
                    MethodsToRemove = ""
                }
                continue
            }
            
            # Check if user is disabled
            if ($user.AccountEnabled -eq $false) {
                $legacyMfaData.ServiceAccountUsers += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    ServiceAccountType = "Disabled Account"
                    Status = "Disabled Account"
                    DetectionMethod = "Account disabled"
                }
                
                $legacyMfaData.UserMfaStatus[$user.UserPrincipalName] = @{
                    Status = "Disabled Account"
                    Methods = "Account is disabled"
                    Phase2Action = "No Action Required"
                    MigrationGroup = "Not Applicable"
                    Phase2Week = "Not applicable"
                    Phase2Priority = "None"
                    SecureMethodCount = 0
                    InsecureMethodCount = 0
                    MethodsToRemove = ""
                }
                continue
            }
            
            # Use REST API to get authentication methods
            $uri = "https://graph.microsoft.com/v1.0/users/$($user.Id)/authentication/methods"
            $authMethodsResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction SilentlyContinue
            $authMethods = $authMethodsResponse.value
            
            if ($authMethods -and $authMethods.Count -gt 0) {
                # Check if user only has password method
                $nonPasswordMethods = $authMethods | Where-Object { 
                    $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' 
                }
                
                # Prepare method analysis
                $methodList = @()
                $methodSet = @{}
                $secureMethodList = @()
                $insecureMethodList = @()
                
                foreach ($method in $authMethods) {
                    $type = switch ($method.'@odata.type') {
                        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' { 'Authenticator' }
                        '#microsoft.graph.phoneAuthenticationMethod' { 'Phone' }
                        '#microsoft.graph.smsAuthenticationMethod' { 'SMS' }
                        '#microsoft.graph.emailAuthenticationMethod' { 'Email' }
                        '#microsoft.graph.fido2AuthenticationMethod' { 'FIDO2' }
                        '#microsoft.graph.softwareOathAuthenticationMethod' { 'SoftwareOath' }
                        '#microsoft.graph.temporaryAccessPassAuthenticationMethod' { 'TAP' }
                        '#microsoft.graph.x509CertificateAuthenticationMethod' { 'Certificate' }
                        '#microsoft.graph.passwordAuthenticationMethod' { 'Password' }
                        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { 'Windows' }
                        default { 'Unknown' }
                    }
                    
                    # Only add if not already present
                    if (-not $methodSet.ContainsKey($type)) {
                        $methodSet[$type] = $true
                        $methodList += $type
                        
                        # Categorize as secure or insecure
                        if ($type -in @('Authenticator', 'FIDO2', 'Windows', 'Certificate', 'SoftwareOath')) {
                            $secureMethodList += $type
                        }
                        elseif ($type -in @('Phone', 'SMS', 'Email')) {
                            $insecureMethodList += $type
                        }
                    }
                }
                
                # Determine Phase 2 action and migration group
                $hasSecureMethods = $secureMethodList.Count -gt 0
                $hasInsecureMethods = $insecureMethodList.Count -gt 0
                
                $phase2Action = "No Action Required"
                $migrationGroup = "Not Applicable"
                $phase2Week = "Not applicable"
                $phase2Priority = "None"
                $methodsToRemove = ""
                
                if ($nonPasswordMethods.Count -eq 0) {
                    # Check if this is a service account
                    $isServiceAccount = $false
                    $serviceAccountType = "Unknown"
                    $detectionReason = ""
                    
                    # Check name patterns
                    foreach ($namePattern in $serviceAccountNamePatterns) {
                        if ($user.UserPrincipalName -like $namePattern) {
                            $isServiceAccount = $true
                            $serviceAccountType = "Service Account (Name Pattern)"
                            $detectionReason = "Name pattern: $namePattern"
                            break
                        }
                    }
                    
                    # Check licenses if not detected by name
                    if (-not $isServiceAccount -and $user.AssignedLicenses) {
                        try {
                            $allSkus = Get-MgSubscribedSku -ErrorAction SilentlyContinue
                            
                            foreach ($license in $user.AssignedLicenses) {
                                $sku = $allSkus | Where-Object { $_.SkuId -eq $license.SkuId }
                                
                                if ($sku) {
                                    foreach ($pattern in $excludedSkuPatterns) {
                                        if ($sku.SkuPartNumber -like $pattern) {
                                            $isServiceAccount = $true
                                            $serviceAccountType = "Service Account (License)"
                                            $detectionReason = "License: $($sku.SkuPartNumber)"
                                            break
                                        }
                                    }
                                    if ($isServiceAccount) { break }
                                }
                            }
                        }
                        catch {
                            # Continue without license detection
                        }
                    }
                    
                    if ($isServiceAccount) {
                        $legacyMfaData.ServiceAccountUsers += [PSCustomObject]@{
                            UserPrincipalName = $user.UserPrincipalName
                            ServiceAccountType = $serviceAccountType
                            Status = "Service Account (Password Only)"
                            DetectionMethod = $detectionReason
                        }
                        $migrationGroup = "Service Account"
                    }
                    else {
                        $legacyMfaData.RegularUsersNoMfa += [PSCustomObject]@{
                            UserPrincipalName = $user.UserPrincipalName
                            Status = "Password Only - Needs MFA"
                        }
                        $phase2Action = "Enable MFA"
                        $migrationGroup = "No MFA"
                        $phase2Week = "1"
                        $phase2Priority = "Critical"
                    }
                }
                else {
                    # User has authentication methods beyond password
                    $legacyMfaData.CurrentMethodUsers++
                    
                    # Count method types
                    $uniqueMethodTypes = @{}
                    
                    foreach ($method in $authMethods) {
                        $odataType = $method.'@odata.type'
                        
                        $methodType = switch ($odataType) {
                            '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' { 'MicrosoftAuthenticator' }
                            '#microsoft.graph.phoneAuthenticationMethod' { 'Voice/Phone' }
                            '#microsoft.graph.smsAuthenticationMethod' { 'SMS' }
                            '#microsoft.graph.emailAuthenticationMethod' { 'Email' }
                            '#microsoft.graph.fido2AuthenticationMethod' { 'Fido2' }
                            '#microsoft.graph.softwareOathAuthenticationMethod' { 'SoftwareOath' }
                            '#microsoft.graph.temporaryAccessPassAuthenticationMethod' { 'TemporaryAccessPass' }
                            '#microsoft.graph.x509CertificateAuthenticationMethod' { 'X509Certificate' }
                            '#microsoft.graph.passwordAuthenticationMethod' { 'Password' }
                            '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { 'WindowsHello' }
                            default { "Other" }
                        }
                        
                        if (-not $uniqueMethodTypes.ContainsKey($methodType)) {
                            $uniqueMethodTypes[$methodType] = $true
                            
                            if ($legacyMfaData.CurrentMethodDistribution.ContainsKey($methodType)) {
                                $legacyMfaData.CurrentMethodDistribution[$methodType]++
                            }
                            else {
                                $legacyMfaData.CurrentMethodDistribution[$methodType] = 1
                            }
                        }
                    }
                    
                    # Determine Phase 2 actions
                    if ($hasInsecureMethods -and $hasSecureMethods) {
                        $phase2Action = "Remove Insecure Methods"
                        $migrationGroup = "Auto-remove"
                        $phase2Week = "1-2"
                        $phase2Priority = "Medium"
                        $methodsToRemove = $insecureMethodList -join ", "
                    }
                    elseif ($hasInsecureMethods -and -not $hasSecureMethods) {
                        $phase2Action = "Register Secure Method"
                        $migrationGroup = "Needs assistance"
                        $phase2Week = "3-4"
                        $phase2Priority = "High"
                    }
                    elseif ($hasSecureMethods -and -not $hasInsecureMethods) {
                        $phase2Action = "No Action Required"
                        $migrationGroup = "Secure already"
                        $phase2Week = "Not applicable"
                        $phase2Priority = "None"
                    }
                }
                
                # Store enhanced user status
                $legacyMfaData.UserMfaStatus[$user.UserPrincipalName] = @{
                    Status = if ($nonPasswordMethods.Count -eq 0) { 
                        if ($isServiceAccount) { "Service Account" } else { "Password Only - Needs MFA" }
                    } else { "Has Current Methods" }
                    Methods = $methodList -join ", "
                    Phase2Action = $phase2Action
                    MigrationGroup = $migrationGroup
                    Phase2Week = $phase2Week
                    Phase2Priority = $phase2Priority
                    SecureMethodCount = $secureMethodList.Count
                    InsecureMethodCount = $insecureMethodList.Count
                    MethodsToRemove = $methodsToRemove
                    SecureMethods = $secureMethodList -join ", "
                    InsecureMethods = $insecureMethodList -join ", "
                }
            }
            else {
                $legacyMfaData.RegularUsersNoMfa += [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    Status = "No Authentication Methods"
                }
                
                $legacyMfaData.UserMfaStatus[$user.UserPrincipalName] = @{
                    Status = "No Authentication Methods"
                    Methods = "No methods found - critical issue"
                    Phase2Action = "Investigate Account"
                    MigrationGroup = "Error"
                    Phase2Week = "Immediate"
                    Phase2Priority = "Critical"
                    SecureMethodCount = 0
                    InsecureMethodCount = 0
                    MethodsToRemove = ""
                    SecureMethods = ""
                    InsecureMethods = ""
                }
            }
        }
        catch {
            $legacyMfaData.UserMfaStatus[$user.UserPrincipalName] = @{
                Status = "Cannot Determine"
                Methods = "API access error: $($_.Exception.Message)"
                Phase2Action = "Review Manually"
                MigrationGroup = "Error"
                Phase2Week = "Not applicable"
                Phase2Priority = "None"
                SecureMethodCount = 0
                InsecureMethodCount = 0
                MethodsToRemove = ""
                SecureMethods = ""
                InsecureMethods = ""
            }
        }
    }
    
    # Complete progress display
    if ($isVSCode) {
        Write-Host "`rAnalyzing user $processed of $($legacyMfaData.TotalUsers) (100% complete)     " -ForegroundColor Green
        Write-Host "" # New line after progress
    }
    else {
        Write-Progress -Activity "Analyzing MFA Status" -Completed
    }
    
    # Display results
    Write-Host "`nLegacy MFA Registration Summary:" -ForegroundColor Yellow
    Write-Host "Total Users Analyzed: $($legacyMfaData.TotalUsers)" -ForegroundColor White
    Write-Host "Current Authentication Methods Users: $($legacyMfaData.CurrentMethodUsers)" -ForegroundColor Green
    Write-Host "Service Account Users (No MFA Expected): $($legacyMfaData.ServiceAccountUsers.Count)" -ForegroundColor Blue
    Write-Host "Regular Users with No MFA (Attention Needed): $($legacyMfaData.RegularUsersNoMfa.Count)" -ForegroundColor Red
    
    # Show method distribution
    if ($legacyMfaData.CurrentMethodDistribution.Count -gt 0) {
        Write-Host "`n=== CURRENT AUTHENTICATION METHOD ENROLLMENTS ===" -ForegroundColor Cyan
        
        $sortedMethods = $legacyMfaData.CurrentMethodDistribution.GetEnumerator() | Sort-Object Value -Descending
        foreach ($method in $sortedMethods) {
            $percentage = [math]::Round(($method.Value / $legacyMfaData.TotalUsers) * 100, 1)
            Write-Host "  $($method.Key): $($method.Value) users ($percentage%)" -ForegroundColor White
        }
    }
    
    # Show regular users without MFA
    if ($legacyMfaData.RegularUsersNoMfa.Count -gt 0) {
        Write-Host "`n=== USERS WITHOUT MFA (ACTION REQUIRED) ===" -ForegroundColor Red
        $legacyMfaData.RegularUsersNoMfa | Select-Object -First 10 | ForEach-Object {
            Write-Host "  - $($_.UserPrincipalName)" -ForegroundColor Red
        }
        if ($legacyMfaData.RegularUsersNoMfa.Count -gt 10) {
            Write-Host "  ... and $($legacyMfaData.RegularUsersNoMfa.Count - 10) more users" -ForegroundColor Red
        }
    }
    
    # Show Phase 2 migration summary
    $migrationGroups = $legacyMfaData.UserMfaStatus.Values | Group-Object MigrationGroup
    Write-Host "`n=== PHASE 2 MIGRATION GROUPS ===" -ForegroundColor Cyan
    foreach ($group in $migrationGroups | Sort-Object Name) {
        $color = switch ($group.Name) {
            "Auto-remove" { "Green" }
            "Needs assistance" { "Yellow" }
            "No MFA" { "Red" }
            "Secure already" { "DarkGreen" }
            default { "Gray" }
        }
        Write-Host "  $($group.Name): $($group.Count) users" -ForegroundColor $color
    }
}
catch {
    Write-Warning "Could not analyze MFA registration: $_"
}

return $legacyMfaData
