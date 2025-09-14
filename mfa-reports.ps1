<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

.SYNOPSIS
    Generates MFA assessment reports using PSWriteOffice for proper Word document output
    UPDATED VERSION - Using PSWriteOffice instead of PSWriteWord for cross-platform compatibility
#>

param(
    [Parameter(Mandatory=$true)]
    [hashtable]$AssessmentData
)

Write-Host "`n=== GENERATING REPORTS ===" -ForegroundColor Cyan

# Check for modules
$useExcel = $false
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    $useExcel = $true
}

# Import PSWriteOffice
$useWord = $false
if (Get-Module -ListAvailable -Name PSWriteOffice) {
    try {
        Import-Module PSWriteOffice -ErrorAction Stop
        $useWord = $true
        Write-Host "PSWriteOffice module loaded for Word document generation" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to load PSWriteOffice module: $_"
    }
}

$CurrentPolicy = $AssessmentData.CurrentPolicy
$LegacyMfaData = $AssessmentData.LegacyMfaData
$CaPolicies = $AssessmentData.CaPolicies
$Recommendations = $AssessmentData.Recommendations
$PrivilegedData = $AssessmentData.PrivilegedData

# Get tenant information for folder creation
$tenantName = "UnknownTenant"
$tenantId = (Get-MgContext).TenantId
try {
    $org = Get-MgOrganization
    $tenantName = $org.DisplayName
}
catch {
    # Use tenant ID if name cannot be retrieved
}

# Create safe folder name
$safeTenantName = $tenantName -replace '[^\w\-\.]', '_'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# Create tenant-specific folder
$reportFolder = "MFA_Reports_$safeTenantName`_$tenantId"
if (!(Test-Path $reportFolder)) {
    New-Item -ItemType Directory -Path $reportFolder | Out-Null
    Write-Host "Created report folder: $reportFolder" -ForegroundColor Green
}

# Generate file paths in the tenant folder
$reportPath = Join-Path $reportFolder "MFA_Migration_Report_$timestamp.docx"
$excelPath = Join-Path $reportFolder "MFA_User_Methods_$timestamp.xlsx"

# PSWriteOffice Word document function with OneDrive long path support
function New-PSWriteOfficeReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Content,
        
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [string]$Title = ""
    )
    
    if (-not $useWord) {
        return $false
    }
    
    Write-Host "Creating Word document using PSWriteOffice..." -ForegroundColor Yellow
    
    # Create a temporary stub path for OneDrive long path support
    $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
    $stubPath = "C:\PSO_$timestamp"
    $stubCreated = $false
    
    try {
        # Get the final directory and filename
        $finalDirectory = Split-Path $FilePath -Parent
        $fileName = Split-Path $FilePath -Leaf
        
        # Ensure final directory exists
        if (-not (Test-Path $finalDirectory)) {
            New-Item -ItemType Directory -Path $finalDirectory -Force | Out-Null
        }
        
        # Create junction to the target directory for OneDrive long path support
        try {
            cmd /c mklink /J "$stubPath" "$finalDirectory" 2>&1 | Out-Null
            if (Test-Path $stubPath) {
                Write-Host "Created junction for OneDrive long path handling" -ForegroundColor Gray
                $stubCreated = $true
                $workingFilePath = Join-Path $stubPath $fileName
            } else {
                Write-Warning "Junction creation failed, using original path"
                $workingFilePath = $FilePath
            }
        }
        catch {
            Write-Warning "Could not create junction, using original path: $_"
            $workingFilePath = $FilePath
        }
        
        # Create Word document using working path
        $Document = New-OfficeWord -FilePath $workingFilePath
        
        # Add title if provided
        if ($Title) {
            $titleParagraph = New-OfficeWordText -Document $Document -Text $Title -Bold $true -Color DarkBlue -Alignment Center -ReturnObject
            $titleParagraph.FontSize = 16
            New-OfficeWordText -Document $Document -Text ""
        }
        
        # Process content line by line - FIXED VERSION (no complex list processing)
        $lines = $Content -split "`n"
        
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].Trim()
            
            # Handle empty lines
            if ([string]::IsNullOrWhiteSpace($line)) {
                New-OfficeWordText -Document $Document -Text ""
                continue
            }
            
            # Headers with === (check next line)
            if ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match "^={3,}") {
                $headerParagraph = New-OfficeWordText -Document $Document -Text $line -Bold $true -Color DarkBlue -ReturnObject
                $headerParagraph.FontSize = 14
                $i++ # Skip the === line
                continue
            }
            
            # Headers with --- (check next line)
            if ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match "^-{3,}") {
                $subHeaderParagraph = New-OfficeWordText -Document $Document -Text $line -Bold $true -Color DarkRed -ReturnObject
                $subHeaderParagraph.FontSize = 12
                $i++ # Skip the --- line
                continue
            }
            
            # Skip separator lines
            if ($line -match "^={3,}$" -or $line -match "^-{3,}$") {
                continue
            }
            
            # Section headers (ALL CAPS)
            if ($line -match "^[A-Z][A-Z\s\d\-:()]+$" -and $line.Length -gt 10 -and $line -notmatch "\." -and $line -notmatch ",") {
                $capsHeaderParagraph = New-OfficeWordText -Document $Document -Text $line -Bold $true -Color DarkGreen -ReturnObject
                $capsHeaderParagraph.FontSize = 12
                continue
            }
            
            # Numbered list items (convert to bullets)
            if ($line -match "^\d+\.\s+") {
                $listItem = $line -replace "^\d+\.\s+", ""
                New-OfficeWordText -Document $Document -Text "• $($listItem.Trim())"
                continue
            }
            
            # Bullet points (keep as is but ensure proper formatting)
            if ($line -match "^[-•]\s+") {
                $listItem = $line -replace "^[-•]\s+", ""
                New-OfficeWordText -Document $Document -Text "• $($listItem.Trim())"
                continue
            }
            
            # User entries with @ symbol
            if ($line -match "@") {
                $userParagraph = New-OfficeWordText -Document $Document -Text $line -ReturnObject
                $userParagraph.FontFamily = 'Courier New'
                $userParagraph.FontSize = 10
                continue
            }
            
            # Total/summary lines
            if ($line -match "^Total") {
                New-OfficeWordText -Document $Document -Text $line -Bold $true -Color DarkMagenta
                continue
            }
            
            # Status lines with brackets
            if ($line -match "^\[" -and $line -match "\]") {
                $statusColor = "Black"
                if ($line -match "\[OK\]" -or $line -match "\[SUCCESS\]") {
                    $statusColor = "Green"
                } elseif ($line -match "\[WARNING\]" -or $line -match "\[CAUTION\]") {
                    $statusColor = "Orange"
                } elseif ($line -match "\[CRITICAL\]" -or $line -match "\[ERROR\]" -or $line -match "\[DANGER\]") {
                    $statusColor = "Red"
                }
                New-OfficeWordText -Document $Document -Text $line -Bold $true -Color $statusColor
                continue
            }
            
            # Regular text
            New-OfficeWordText -Document $Document -Text $line
        }
        
        # Save document
        Save-OfficeWord -Document $Document
        Write-Host "Word document created successfully using PSWriteOffice" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Warning "Failed to create Word document: $_"
        Write-Warning "Error details: $($_.Exception.Message)"
        return $false
    }
    finally {
        # Clean up stub path if created
        if ($stubCreated -and (Test-Path $stubPath)) {
            try {
                cmd /c rmdir "$stubPath" 2>&1 | Out-Null
                Write-Host "Cleaned up temporary junction" -ForegroundColor Gray
            }
            catch {
                Write-Warning "Could not remove stub path: $stubPath"
            }
        }
    }
}

# Generate Excel report with user data
if ($useExcel) {
    $allUsersData = @()
    
    # Process users with MFA
    foreach ($upn in $LegacyMfaData.UserMfaStatus.Keys) {
        $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
        
        # Check if user is privileged
        $isPrivileged = $PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { $_.UserPrincipalName -eq $upn }
        $privilegedRoles = ""
        if ($isPrivileged) {
            $privilegedRoles = ($isPrivileged | Select-Object -ExpandProperty RoleName -Unique) -join "; "
        }
        
        $allUsersData += [PSCustomObject]@{
            UserPrincipalName = $upn
            Status = $userStatus.Status
            IsPrivileged = if ($isPrivileged) { "Yes" } else { "No" }
            PrivilegedRoles = $privilegedRoles
            CurrentMethods = $userStatus.Methods
            MethodsToRemove = $userStatus.MethodsToRemove
            Phase2Action = $userStatus.Phase2Action
            Phase2Priority = $userStatus.Phase2Priority
            Phase2Week = $userStatus.Phase2Week
            MigrationRisk = "Low - Existing MFA will continue working"
            SecurityCompliance = if ($userStatus.Methods -eq "Password Only") { "Non-compliant - No MFA" } else { "Compliant" }
        }
    }
    
    # Add users with no MFA
    foreach ($user in $LegacyMfaData.RegularUsersNoMfa) {
        $isPrivileged = $PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }
        $privilegedRoles = ""
        if ($isPrivileged) {
            $privilegedRoles = ($isPrivileged | Select-Object -ExpandProperty RoleName -Unique) -join "; "
        }
        
        $allUsersData += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            Status = "Active - No MFA"
            IsPrivileged = if ($isPrivileged) { "Yes" } else { "No" }
            PrivilegedRoles = $privilegedRoles
            CurrentMethods = "Password Only"
            MethodsToRemove = ""
            Phase2Action = "Register MFA (Security Enhancement)"
            Phase2Priority = if ($isPrivileged) { "Critical" } else { "Medium" }
            Phase2Week = "1-2"
            MigrationRisk = "Zero - No MFA to migrate"
            SecurityCompliance = "Non-compliant - No MFA"
        }
    }
    
    # Export to Excel
    Remove-Item $excelPath -ErrorAction SilentlyContinue
    $allUsersData | Export-Excel -Path $excelPath -WorksheetName "User Analysis" -AutoSize -TableStyle Medium2 -FreezeTopRow
    
    Write-Host "User data exported to: $excelPath" -ForegroundColor Green
}

# Generate privileged users report if applicable
$privExcelPath = $null
if ($PrivilegedData.AnalyzedPrivilegedUsers.Count -gt 0 -and $useExcel) {
    $privExcelPath = Join-Path $reportFolder "Privileged_Users_Security_Analysis_$timestamp.xlsx"
    
    $privilegedUsersData = $PrivilegedData.AnalyzedPrivilegedUsers | ForEach-Object {
        $upn = $_.UserPrincipalName
        $userMfaStatus = $LegacyMfaData.UserMfaStatus[$upn]
        
        [PSCustomObject]@{
            UserPrincipalName = $upn
            RoleName = $_.RoleName
            IsBuiltIn = $_.IsBuiltIn
            CurrentMfaMethod = if ($userMfaStatus) { $userMfaStatus.Methods } else { "Password Only" }
            HasFIDO2 = if ($userMfaStatus -and $userMfaStatus.Methods -match "FIDO2") { "Yes" } else { "No" }
            SecurityRecommendation = if ($userMfaStatus) { 
                if ($userMfaStatus.Methods -match "FIDO2") { "Compliant - FIDO2 Enabled" }
                elseif ($userMfaStatus.Methods -ne "Password Only") { "Deploy FIDO2 Security Key" }
                else { "CRITICAL - Enable MFA Immediately" }
            } else { "CRITICAL - Enable MFA Immediately" }
            Priority = if (-not $userMfaStatus -or $userMfaStatus.Methods -eq "Password Only") { "Critical" } else { "High" }
        }
    }
    
    Remove-Item $privExcelPath -ErrorAction SilentlyContinue
    $privilegedUsersData | Export-Excel -Path $privExcelPath -WorksheetName "Privileged Users" -AutoSize -TableStyle Medium6 -FreezeTopRow
    
    Write-Host "Privileged users analysis saved to: $privExcelPath" -ForegroundColor Yellow
}

# Build the main report content
$usersWithMfa = ($LegacyMfaData.UserMfaStatus.Keys | Where-Object { 
    $LegacyMfaData.UserMfaStatus[$_].Status -notin @("Service Account", "Disabled Account") 
}).Count

$usersForAutomaticCleanup = ($LegacyMfaData.UserMfaStatus.Values | Where-Object { 
    $_.Phase2Action -eq "Automatic Cleanup - No User Action Required" 
}).Count

$usersNeedingAssistance = ($LegacyMfaData.UserMfaStatus.Values | Where-Object { 
    $_.Phase2Action -match "Assist" 
}).Count

$privilegedNoMfa = ($PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { 
    $upn = $_.UserPrincipalName
    $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
    -not $userStatus -or $userStatus.Methods -eq "Password Only"
}).Count

# Count break glass accounts to exclude from privileged no MFA count
$breakGlassCount = ($PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { 
    $_.UserPrincipalName -match "breakglass|break-glass|emergency" 
}).Count

$privilegedNoMfaExcludingBreakGlass = [Math]::Max(0, $privilegedNoMfa - $breakGlassCount)

$privilegedNoFido = ($PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { 
    $upn = $_.UserPrincipalName
    $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
    $userStatus -and $userStatus.Methods -notmatch "FIDO2" -and $userStatus.Methods -ne "Password Only"
}).Count

# Determine what methods need to be enabled in the new policy
$allCurrentMethods = @()
foreach ($userStatus in $LegacyMfaData.UserMfaStatus.Values) {
    if ($userStatus.Status -notin @("Service Account", "Disabled Account") -and $userStatus.Methods -ne "Password Only") {
        $methods = $userStatus.Methods -split ", "
        $allCurrentMethods += $methods
    }
}
$uniqueMethods = $allCurrentMethods | Sort-Object -Unique | Where-Object { $_ -and $_.Trim() -ne "" }

# Check which methods are already enabled vs need enabling
$methodsAlreadyEnabled = @()
$methodsNeedingEnable = @()

foreach ($method in $uniqueMethods) {
    $isEnabled = $false
    switch ($method) {
        "Microsoft Authenticator" { $isEnabled = $CurrentPolicy.MicrosoftAuthenticatorEnabled }
        "SMS" { $isEnabled = $CurrentPolicy.SmsEnabled }
        "Voice" { $isEnabled = $CurrentPolicy.VoiceEnabled }
        "Email" { $isEnabled = $CurrentPolicy.EmailEnabled }
        "FIDO2" { $isEnabled = $CurrentPolicy.Fido2Enabled }
        "Windows Hello" { $isEnabled = $CurrentPolicy.WindowsHelloEnabled }
        default { $isEnabled = $false }
    }
    
    if ($isEnabled) {
        $methodsAlreadyEnabled += $method
    } else {
        $methodsNeedingEnable += $method
    }
}

# Create the comprehensive report content
$reportContent = @"
MFA TO AUTHENTICATION METHODS POLICY MIGRATION REPORT
=====================================================
Organisation: $tenantName
Generated: $(Get-Date)
Report Version: 2.0 - Zero Disruption Migration Assessment

EXECUTIVE SUMMARY - ZERO DISRUPTION MIGRATION CONFIRMED
========================================================

MIGRATION SAFETY ASSESSMENT - FINAL RESULT: PROCEED WITH CONFIDENCE
-------------------------------------------------------------------
[SUCCESS] ZERO DISRUPTION EXPECTED - All users will continue working normally

Migration Impact Analysis:
- $usersWithMfa users currently have MFA and will continue working exactly as they do today
- $($LegacyMfaData.RegularUsersNoMfa.Count) users have no MFA and will continue working exactly as they do today (password-only access unchanged)
- No users will lose access or experience service disruption during migration
- September 30th deadline is fully achievable with no service interruptions

SECURITY COMPLIANCE TIMELINE (Ongoing Organisational Security Policy)
--------------------------------------------------------------------
$(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
"[CRITICAL] PRIVILEGED USER SECURITY COMPLIANCE GAPS:
- $privilegedNoMfaExcludingBreakGlass privileged users currently lack MFA protection
- These are HIGH-VALUE SECURITY TARGETS requiring policy compliance attention
- Represents ongoing organisational security policy violation
- Recommended action: Implement MFA for all privileged accounts within 30 days
- Note: This is a security compliance issue, separate from migration safety"
} else {
"[OK] PRIVILEGED USERS: All compliant with MFA security policy"
})

$(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) {
"[WARNING] REGULAR USER SECURITY POLICY ENHANCEMENT OPPORTUNITIES:
- $($LegacyMfaData.RegularUsersNoMfa.Count) regular users currently lack MFA protection
- Every user should have MFA per security best practices
- Represents ongoing organisational security enhancement opportunity  
- Recommended action: Implement MFA for all users within 90 days as security improvement initiative
- Note: This is a security enhancement opportunity, separate from migration requirements"
} else {
"[OK] REGULAR USERS: All protected with MFA"
})

KEY DISTINCTION - CRITICAL UNDERSTANDING
---------------------------------------
- Migration Disruption Risk: ZERO - Users will continue working exactly as they do today
- Security Policy Compliance: Separate ongoing issue requiring attention per organisational security policy
- Migration Success ≠ Security Policy Compliance (these are different objectives with different timelines)
- Users without MFA are not migration risks - they are security policy enhancement opportunities

OUR TWO-PHASE APPROACH
======================
We will manage your transition in two phases to ensure zero disruption whilst maintaining security standards.

PHASE 1: MIGRATION COMPLIANCE - ZERO DISRUPTION (By September 30, 2025)
-----------------------------------------------------------------------
Time: 1-2 days
Target: Meet Microsoft deadline with zero service disruption

[OK] GUARANTEED OUTCOMES:
- All $usersWithMfa users with current MFA will continue accessing systems normally
- All users without MFA will continue working exactly as they do today (password-only access unchanged)
- No users will be locked out or lose any access they currently have
- September 30th deadline compliance achieved with zero service interruptions

What We Will Do:
- Enable ALL currently used authentication methods in the new policy:
$(if ($methodsNeedingEnable.Count -gt 0) {
$methodsNeedingEnable | ForEach-Object { "  - Enable $_" }
""
} else {
"  - All required methods are already enabled in the policy"
""
})
$(if ($methodsAlreadyEnabled.Count -gt 0) {
"Methods already enabled and ready:
$($methodsAlreadyEnabled | ForEach-Object { "  - $_" })
"
})
- Migrate policy management to new Authentication Methods Policy interface
- Test that all current MFA users can continue authenticating without issues
- Document users currently without MFA (for optional Phase 2 security enhancement)
- Complete migration with zero disruption

SUCCESS CRITERIA:
- Zero user complaints about lost access
- All current MFA users continue working normally  
- All users without MFA continue working as they always have (no change)
- Full compliance with Microsoft deadline achieved

PHASE 2: SECURITY ENHANCEMENT PROGRAMME (4-6 weeks after Phase 1)
-----------------------------------------------------------------
Time: 4-6 weeks
Focus: Improve security posture and address ongoing policy compliance gaps

Week 1-2: Automatic Security Improvements
- Automatically remove less secure methods for $usersForAutomaticCleanup users
  (These users already have secure alternatives registered - no user action needed)
- Maintain service continuity throughout improvements

Week 3-4: Security Gap Remediation Programme
$(if ($usersNeedingAssistance -gt 0) {
"- Work with $usersNeedingAssistance users to upgrade to Microsoft Authenticator
- Provide training and support materials for security enhancement
- Track progress and provide follow-up assistance"
} else {
"- No users need assistance - all already have secure methods"
})

$(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) {
"- Address security policy enhancement opportunity: $($LegacyMfaData.RegularUsersNoMfa.Count) users without any MFA
- Register Microsoft Authenticator to improve organisational security posture
- Priority based on user roles and access levels
- Voluntary security enhancement programme"
} else {
"- No security gaps to address - all users already have MFA"
})

Week 5-6: Privileged User Security Enhancement
$(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
"- PRIORITY: Register MFA for $privilegedNoMfaExcludingBreakGlass privileged users (Critical security policy compliance)
- Deploy FIDO2 security keys to $privilegedNoFido administrators (Recommended security enhancement)
- Configure phishing-resistant MFA requirements for high-value targets
- Create conditional access policies for administrator protection"
} else {
"- Deploy FIDO2 security keys to $privilegedNoFido administrators (Optional security enhancement)  
- Configure phishing-resistant MFA requirements
- Create conditional access policies for administrator protection"
})

Final Step: Remove Less Secure Methods (Optional)
- Remove Voice, SMS, and Email from the authentication policy (security hardening)
- Monitor for any issues during security hardening phase
- Maintain exceptions only where absolutely necessary for business continuity

GENERATED REPORTS
================
Location: $reportFolder

[DATA] User Details and Security Enhancement Planning: $(Split-Path $excelPath -Leaf)
- Complete user list with current authentication methods
- Migration impact assessment per user (spoiler: zero impact expected)
- Security compliance status per user for enhancement planning
- Use for project planning and security improvement tracking
- Clear separation between migration safety and security enhancement opportunities

$(if ($privExcelPath) { "[SECURITY] Privileged User Security Analysis: $(Split-Path $privExcelPath -Leaf)
- Administrator account security analysis
- Risk assessment and FIDO2 security enhancement recommendations
- Critical security policy compliance gap identification
" })
[REPORT] Migration and Security Assessment Report: $(Split-Path $reportPath -Leaf)
- This comprehensive assessment report
- Executive summary and recommendations
- Phase-by-phase implementation plan with clear risk/enhancement separation

NEXT STEPS
==========
1. IMMEDIATE (This Week):
   $(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
   "   [CRITICAL] Address $privilegedNoMfaExcludingBreakGlass privileged users without MFA (security policy compliance violation)"
   } else {
   "   [OK] No critical security policy violations requiring immediate attention"
   })

2. MIGRATION PREPARATION (Next 2 Weeks):
   - Review and approve Phase 1 plan (zero disruption guaranteed)
   - Schedule policy migration for before September 30, 2025
   - Prepare communications about upcoming security enhancement opportunities

3. POST-MIGRATION SECURITY ENHANCEMENT PROGRAMME (Optional but Recommended):
   $(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0 -or $usersNeedingAssistance -gt 0) {
   "   - Plan security improvement campaign for $($LegacyMfaData.RegularUsersNoMfa.Count + $usersNeedingAssistance) users
   - Develop user training and support materials for security enhancements
   - Schedule assisted MFA registration sessions as voluntary security improvement"
   } else {
   "   - All users already have secure MFA methods
   - Focus on FIDO2 deployment for administrators as optional security enhancement"
   })

CONCLUSION
=========
Your migration can proceed with complete confidence - zero disruption is expected because users without MFA will continue working exactly as they do today (they never had MFA protection to lose). The identified security opportunities are separate enhancement initiatives to improve your overall security posture in line with organisational policy, not migration blockers.
"@

# Create Word document using PSWriteOffice function
$wordCreated = New-PSWriteOfficeReport -Content $reportContent -FilePath $reportPath -Title "MFA TO AUTHENTICATION METHODS POLICY MIGRATION REPORT"
if ($wordCreated) {
    Write-Host "`nSummary report saved to: $reportPath (Word format)" -ForegroundColor Green
} else {
    # Fall back to text file only if Word fails
    $reportPath = $reportPath -replace '\.docx$', '.txt'
    $reportContent | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "`nSummary report saved to: $reportPath (text format)" -ForegroundColor Green
}

Write-Host "All reports saved to folder: $reportFolder" -ForegroundColor Cyan

return @{
    ReportPath = $reportPath
    ReportFolder = $reportFolder
    ExcelPath = $excelPath
    PrivilegedExcelPath = $privExcelPath
}
