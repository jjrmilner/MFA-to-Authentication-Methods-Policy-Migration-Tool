#Requires -Version 5.1
<#
.SYNOPSIS
    Generates MFA assessment reports using PSWriteWord for proper Word document output
    FINAL VERSION - Using exact working Word function with corrected content
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

# Import PSWriteWord
$useWord = $false
if (Get-Module -ListAvailable -Name PSWriteWord) {
    try {
        Import-Module PSWriteWord -ErrorAction Stop
        $useWord = $true
        Write-Host "PSWriteWord module loaded for Word document generation" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to load PSWriteWord module: $_"
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

# EXACT WORKING WORD DOCUMENT FUNCTION FROM BACKUP
function New-PSWriteWordReport {
    param(
        [string]$Content,
        [string]$FilePath,
        [string]$Title = ""
    )
    
    if (-not $useWord) {
        return $false
    }
    
    Write-Host "Creating Word document using PSWriteWord..." -ForegroundColor Yellow
    
    # Create a temporary stub path to handle long paths
    $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
    $stubPath = "C:\PSW_$timestamp"
    $stubCreated = $false
    
    try {
        # Get the final directory and filename
        $finalDirectory = Split-Path $FilePath -Parent
        $fileName = Split-Path $FilePath -Leaf
        
        # Ensure final directory exists
        if (-not (Test-Path $finalDirectory)) {
            New-Item -ItemType Directory -Path $finalDirectory -Force | Out-Null
        }
        
        # Create junction to the target directory
        try {
            cmd /c mklink /J "$stubPath" "$finalDirectory" 2>&1 | Out-Null
            if (Test-Path $stubPath) {
                Write-Host "Created junction for long path handling" -ForegroundColor Gray
                $stubCreated = $true
            }
        }
        catch {
            Write-Warning "Could not create junction, using original path"
            $stubPath = $finalDirectory
        }
        
        # Build the file path using stub
        $stubFilePath = Join-Path $stubPath $fileName
        
        # Create Word document
        $WordDocument = New-WordDocument $stubFilePath
        
        # Add title if provided
        if ($Title) {
            Add-WordText -WordDocument $WordDocument -Text $Title -FontSize 16 -Bold $true -SpacingAfter 15 -Color "DarkBlue" -Supress $true
        }
        
        # Process content line by line
        $lines = $Content -split "`n"
        $currentList = @()
        $inList = $false
        
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].Trim()
            
            # Handle empty lines
            if ([string]::IsNullOrWhiteSpace($line)) {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList -Supress $true
                    $currentList = @()
                    $inList = $false
                }
                Add-WordText -WordDocument $WordDocument -Text "" -SpacingAfter 6 -Supress $true
                continue
            }
            
            # Headers with === (check next line)
            if ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match "^={3,}") {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList -Supress $true
                    $currentList = @()
                    $inList = $false
                }
                Add-WordText -WordDocument $WordDocument -Text $line -FontSize 14 -Bold $true -SpacingBefore 12 -SpacingAfter 12 -Supress $true
                $i++ # Skip the === line
                continue
            }
            
            # Headers with --- (check next line)
            if ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match "^-{3,}") {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList -Supress $true
                    $currentList = @()
                    $inList = $false
                }
                Add-WordText -WordDocument $WordDocument -Text $line -FontSize 12 -Bold $true -SpacingBefore 10 -SpacingAfter 8 -Supress $true
                $i++ # Skip the --- line
                continue
            }
            
            # Skip separator lines
            if ($line -match "^={3,}$" -or $line -match "^-{3,}$") {
                continue
            }
            
            # Section headers (ALL CAPS)
            if ($line -match "^[A-Z][A-Z\s\d\-:()]+$" -and $line.Length -gt 10 -and $line -notmatch "\." -and $line -notmatch ",") {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList -Supress $true
                    $currentList = @()
                    $inList = $false
                }
                Add-WordText -WordDocument $WordDocument -Text $line -FontSize 12 -Bold $true -SpacingBefore 10 -SpacingAfter 8 -Supress $true
                continue
            }
            
            # Numbered list items (e.g., "1. ", "2. ")
            if ($line -match "^\d+\.\s+") {
                $inList = $true
                $listItem = $line -replace "^\d+\.\s+", ""
                $currentList += $listItem.Trim()
                continue
            }
            
            # Bullet points
            if ($line -match "^[-•]\s+") {
                $inList = $true
                $listItem = $line -replace "^[-•]\s+", ""
                $currentList += $listItem.Trim()
                continue
            }
            
            # End list if we were in one
            if ($inList -and $currentList.Count -gt 0) {
                Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList -Supress $true
                $currentList = @()
                $inList = $false
            }
            
            # Status indicators
            if ($line -match "^\[OK\]") {
                Add-WordText -WordDocument $WordDocument -Text $line -Color "Green" -SpacingAfter 6 -Supress $true
            }
            elseif ($line -match "^\[WARNING\]") {
                Add-WordText -WordDocument $WordDocument -Text $line -Color "Orange" -SpacingAfter 6 -Supress $true
            }
            elseif ($line -match "^\[CRITICAL\]") {
                Add-WordText -WordDocument $WordDocument -Text $line -Color "Red" -SpacingAfter 6 -Supress $true
            }
            elseif ($line -match "^\[X\]") {
                Add-WordText -WordDocument $WordDocument -Text $line -Color "DarkBlue" -SpacingAfter 6 -Supress $true
            }
            elseif ($line -match "^\[DATA\]" -or $line -match "^\[SECURITY\]" -or $line -match "^\[REPORT\]") {
                Add-WordText -WordDocument $WordDocument -Text $line -Color "DarkBlue" -SpacingAfter 6 -Supress $true
            }
            # Tree structure
            elseif ($line -match "^\+--" -or $line -match "^\|") {
                Add-WordText -WordDocument $WordDocument -Text $line -FontFamily "Courier New" -FontSize 10 -SpacingAfter 3 -Supress $true
            }
            # Total/summary lines
            elseif ($line -match "^Total") {
                Add-WordText -WordDocument $WordDocument -Text $line -Bold $true -SpacingBefore 8 -SpacingAfter 6 -Supress $true
            }
            # Example users with @ symbol
            elseif ($line -match "@" -and $line -match "^  - ") {
                Add-WordText -WordDocument $WordDocument -Text $line -FontFamily "Courier New" -FontSize 10 -SpacingAfter 3 -Supress $true
            }
            # "... and X more" continuation lines
            elseif ($line -match "^\.\.\. and \d+ more") {
                Add-WordText -WordDocument $WordDocument -Text $line -Italic $true -SpacingAfter 6 -Supress $true
            }
            else {
                # Regular paragraph
                Add-WordText -WordDocument $WordDocument -Text $line -SpacingAfter 6 -Supress $true
            }
        }
        
        # Handle any remaining list items
        if ($inList -and $currentList.Count -gt 0) {
            Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList -Supress $true
        }
        
        # Save document
        Save-WordDocument -WordDocument $WordDocument -Language 'en-US'
        Write-Host "Document saved successfully" -ForegroundColor Green
        
        # Verify file exists
        $actualPath = Join-Path $finalDirectory $fileName
        if (Test-Path $actualPath) {
            $fileInfo = Get-Item $actualPath
            Write-Host "Word document created: $($fileInfo.Name) - Size: $($fileInfo.Length) bytes" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "File not found at expected location" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Warning "Failed to create Word document with PSWriteWord: $_"
        return $false
    }
    finally {
        # Clean up stub path
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

# Define break-glass accounts
$breakGlassAccounts = @('CyberPerformancePack-BTG@globalmicrosolutions.onmicrosoft.com')

# Export enhanced user methods to Excel with CORRECTED dual assessment
$userReport = @()
foreach ($upn in $LegacyMfaData.UserMfaStatus.Keys) {
    $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
    $methods = $userStatus.Methods
    
    # Get user type/category
    $userCategory = "Regular"
    if ($userStatus.Status -eq "Service Account") {
        $userCategory = "Service"
    } elseif ($userStatus.Status -eq "Disabled Account") {
        $userCategory = "Disabled"
    } elseif ($PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { $_.UserPrincipalName -eq $upn }) {
        $userCategory = "Privileged"
    }
    
    # Get department and manager if possible (placeholder - would need additional API calls)
    $department = ""
    $manager = ""
    
    # CORRECTED: Determine Phase 1 action - focus on preserving EXISTING MFA users
    $phase1Action = "No migration action needed"
    $securityAction = "Compliant"
    
    # Check if user has current MFA methods that need policy support
    if ($userStatus.Status -eq "Has Current Methods") {
        $needsPhoneEnabled = ($methods -match "Phone") -and ($CurrentPolicy['Voice'] -ne 'enabled')
        $needsEmailEnabled = ($methods -match "Email") -and ($CurrentPolicy['Email'] -ne 'enabled')  
        $needsSmsEnabled = ($methods -match "SMS") -and ($CurrentPolicy['Sms'] -ne 'enabled')
        $needsWindowsEnabled = ($methods -match "Windows") -and ($CurrentPolicy['WindowsHello'] -ne 'enabled')
        
        if ($needsPhoneEnabled -or $needsEmailEnabled -or $needsSmsEnabled -or $needsWindowsEnabled) {
            $phase1Action = "Enable existing methods in policy"
        } else {
            $phase1Action = "Ready - Methods already enabled"
        }
        $securityAction = "Compliant"
    }
    elseif ($userStatus.Status -eq "Password Only - Needs MFA" -or $userStatus.Status -eq "No Authentication Methods") {
        $phase1Action = "No disruption - Continues current access"
        
        # Security compliance assessment
        if ($userCategory -eq "Privileged") {
            $securityAction = "CRITICAL - Admin requires MFA"
        } else {
            $securityAction = "WARNING - User should have MFA"
        }
    }
    
    $userReport += [PSCustomObject]@{
        UserPrincipalName = $upn
        DisplayName = $upn.Split('@')[0] -replace '\.', ' '
        Status = $userStatus.Status
        User_Category = $userCategory
        # CORRECTED: Dual assessment approach
        Migration_Impact = if ($userStatus.Status -eq "Has Current Methods") { "Protected - Will continue working" } 
                          elseif ($userStatus.Status -eq "Password Only - Needs MFA") { "Unaffected - No change in access" }
                          else { "Service Account - No action needed" }
        Phase1_Migration_Action = $phase1Action
        Security_Compliance = $securityAction
        Security_Priority = if ($userCategory -eq "Privileged" -and $userStatus.Status -match "Password Only") { "CRITICAL" }
                           elseif ($userStatus.Status -match "Password Only") { "WARNING" }
                           elseif ($userStatus.SecureMethodCount -eq 0 -and $userStatus.Status -eq "Has Current Methods") { "Insecure Methods" }
                           else { "COMPLIANT" }
        Phase2_Action = $userStatus.Phase2Action
        Phase2_Week = $userStatus.Phase2Week
        Phase2_Priority = $userStatus.Phase2Priority
        Methods_To_Remove = $userStatus.MethodsToRemove
        Methods_To_Keep = $userStatus.SecureMethods
        Secure_Methods_Count = $userStatus.SecureMethodCount
        Insecure_Methods_Count = $userStatus.InsecureMethodCount
        Migration_Status = "Not started"
        Last_Contact_Date = ""
        Contact_Attempts = 0
        Department = $department
        Manager = $manager
        HasPassword = if ($methods -match "Password") { "Yes" } else { "No" }
        HasAuthenticator = if ($methods -match "Authenticator") { "Yes" } else { "No" }
        HasPhone = if ($methods -match "Phone") { "Yes" } else { "No" }
        HasSMS = if ($methods -match "SMS") { "Yes" } else { "No" }
        HasEmail = if ($methods -match "Email") { "Yes" } else { "No" }
        HasFIDO2 = if ($methods -match "FIDO2") { "Yes" } else { "No" }
        HasWindowsHello = if ($methods -match "Windows") { "Yes" } else { "No" }
        HasSoftwareOATH = if ($methods -match "SoftwareOath") { "Yes" } else { "No" }
        HasCertificate = if ($methods -match "Certificate") { "Yes" } else { "No" }
        AllMethods = $methods
        Notes = ""
    }
}

# Export to Excel if available, otherwise CSV
if ($useExcel) {
    $userReport | Export-Excel -Path $excelPath -WorksheetName "User Methods" -AutoSize -TableName "UserMethods" -TableStyle Medium9 -FreezeTopRow
    Write-Host "Enhanced user methods Excel saved to: $excelPath" -ForegroundColor Green
} else {
    $csvPath = $excelPath -replace '\.xlsx$', '.csv'
    $userReport | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Enhanced user methods CSV saved to: $csvPath" -ForegroundColor Green
}

# Export enhanced privileged users report
$privExcelPath = $null
$privilegedNoFido = 0
$privilegedWithWeakMfa = @()
$privilegedNoMfaExcludingBreakGlass = 0

if ($PrivilegedData -and $PrivilegedData.AnalyzedPrivilegedUsers.Count -gt 0) {
    $privExcelPath = Join-Path $reportFolder "MFA_Privileged_Users_$timestamp.xlsx"
    $privReport = @()
    
    # Group by unique users
    $uniquePrivUsers = $PrivilegedData.AnalyzedPrivilegedUsers | Group-Object UserPrincipalName
    
    foreach ($userGroup in $uniquePrivUsers) {
        $privUser = $userGroup.Group[0]  # Take first entry for user details
        $allRoles = $userGroup.Group | Select-Object -ExpandProperty RoleName -Unique
        $roleList = $allRoles -join "; "
        $roleTypes = $userGroup.Group | Select-Object -ExpandProperty RoleType -Unique
        
        $userStatus = $LegacyMfaData.UserMfaStatus[$privUser.UserPrincipalName]
        $methods = $userStatus.Methods
        
        # Check for phishing-resistant methods
        $hasFIDO2 = $methods -match "FIDO2"
        $hasCertificate = $methods -match "Certificate"
        $hasWindowsHello = $methods -match "Windows"
        $hasPhishingResistant = $hasFIDO2 -or $hasCertificate -or $hasWindowsHello
        
        # Check if break-glass account
        $isBreakGlass = $privUser.UserPrincipalName -in $breakGlassAccounts
        
        # Check if user has no MFA (excluding break-glass)
        if (($userStatus.Status -eq "Password Only - Needs MFA" -or 
             $userStatus.Status -eq "No Authentication Methods") -and -not $isBreakGlass) {
            $privilegedNoMfaExcludingBreakGlass++
        }
        
        # Determine FIDO2 requirement
        $requiresFIDO2 = if ($isBreakGlass) { "No" } 
                         elseif ($hasFIDO2) { "Has already" } 
                         else { "Yes" }
        
        # Risk level based on roles and MFA status
        $riskLevel = "Medium"
        if ($roleList -match "Global Administrator") {
            $riskLevel = "Critical"
        } elseif ($roleList -match "Security Administrator|Privileged") {
            $riskLevel = "High"
        }
        
        # Determine security status
        $securityStatus = if ($isBreakGlass) {
            "Managed Break-Glass Account"
        } elseif (-not ($methods -match "Authenticator|FIDO2|Windows|Certificate|Phone|SMS|Email") -or $methods -eq "Password") {
            "NO MFA - CRITICAL RISK"
        } elseif ($hasPhishingResistant) {
            "Phishing-Resistant MFA"
        } else {
            "Basic MFA Only"
        }
        
        # FIDO2 deployment status
        $fido2Status = if ($hasFIDO2) { 
            "Complete" 
        } elseif ($isBreakGlass) {
            "Not Required - Break Glass"
        } else { 
            "Not started" 
        }
        
        # Get last sign-in if possible (placeholder)
        $lastSignIn = ""
        
        # Admin tier classification
        $adminTier = if ($roleList -match "Global Administrator|Privileged Role Administrator") {
            "Tier 0"
        } elseif ($roleList -match "Security Administrator|User Administrator|Exchange Administrator") {
            "Tier 1"
        } else {
            "Tier 2"
        }
        
        # Count for unique users without FIDO2
        if (-not $hasPhishingResistant -and -not $isBreakGlass) {
            $privilegedNoFido++
            $privilegedWithWeakMfa += [PSCustomObject]@{
                UserPrincipalName = $privUser.UserPrincipalName
                RoleName = $roleList
                CurrentMethods = $methods
            }
        }
        
        $privReport += [PSCustomObject]@{
            UserPrincipalName = $privUser.UserPrincipalName
            DisplayName = $privUser.UserPrincipalName.Split('@')[0] -replace '\.', ' '
            All_Roles = $roleList
            Role_Count = $allRoles.Count
            RoleType = $roleTypes -join ", "
            Admin_Tier = $adminTier
            Risk_Level = $riskLevel
            Security_Status = $securityStatus
            MFA_Status = if ($isBreakGlass -and ($methods -eq "Password")) { "Break-Glass - Pending Registration" } else { $userStatus.Status }
            Methods = $userStatus.Methods
            Requires_FIDO2 = $requiresFIDO2
            FIDO2_Deployment_Status = $fido2Status
            Phishing_Resistant_Status = if ($hasPhishingResistant) { "Compliant" } elseif ($isBreakGlass) { "Break-Glass Exception" } else { "Needs FIDO2" }
            Is_Break_Glass = if ($isBreakGlass) { "Yes" } else { "No" }
            Has_Phishing_Resistant_MFA = if ($hasPhishingResistant) { "Yes" } else { if ($isBreakGlass) { "PENDING - Break-Glass" } else { "NO - HIGH RISK" } }
            Has_FIDO2 = if ($hasFIDO2) { "Yes" } else { "No" }
            Has_Certificate = if ($hasCertificate) { "Yes" } else { "No" }
            Has_Windows_Hello = if ($hasWindowsHello) { "Yes" } else { "No" }
            Has_Global_Admin = if ($roleList -match "Global Administrator") { "Yes" } else { "No" }
            Last_Sign_In = $lastSignIn
            Migration_Status = "Not started"
            Notes = if ($isBreakGlass) { "Managed break-glass account - configure during migration" } else { "" }
        }
    }
    
    if ($useExcel) {
        $privReport | Export-Excel -Path $privExcelPath -WorksheetName "Privileged Users" -AutoSize -TableName "PrivilegedUsers" -TableStyle Medium10 -FreezeTopRow
        Write-Host "Enhanced privileged users Excel saved to: $privExcelPath" -ForegroundColor Yellow
    } else {
        $privCsvPath = $privExcelPath -replace '\.xlsx$', '.csv'
        $privReport | Export-Csv -Path $privCsvPath -NoTypeInformation
        Write-Host "Enhanced privileged users CSV saved to: $privCsvPath" -ForegroundColor Yellow
    }
}

# Analyze users for Phase 2 categorization
$usersForAutomaticCleanup = 0
$usersNeedingAssistance = 0
$automaticCleanupList = @()
$assistanceRequiredList = @()

foreach ($upn in $LegacyMfaData.UserMfaStatus.Keys) {
    $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
    
    if ($userStatus.MigrationGroup -eq "Auto-remove") {
        $usersForAutomaticCleanup++
        $automaticCleanupList += [PSCustomObject]@{
            UserPrincipalName = $upn
            SecureMethods = $userStatus.SecureMethods
            MethodsToRemove = $userStatus.MethodsToRemove
        }
    }
    elseif ($userStatus.MigrationGroup -eq "Needs assistance") {
        $usersNeedingAssistance++
        $assistanceRequiredList += [PSCustomObject]@{
            UserPrincipalName = $upn
            OnlyHas = $userStatus.InsecureMethods
            NeedsToRegister = "Microsoft Authenticator or FIDO2"
        }
    }
}

# Calculate migration readiness statistics
$totalUsers = $LegacyMfaData.TotalUsers
$usersWithMfa = $LegacyMfaData.CurrentMethodUsers
$usersWithSecureMethods = ($LegacyMfaData.UserMfaStatus.Values | Where-Object { $_.SecureMethodCount -gt 0 }).Count
$securityScore = if ($totalUsers -gt 0) { [math]::Round(($usersWithSecureMethods / $totalUsers) * 100, 1) } else { 0 }

# Create summary of enabled/disabled methods
$methodsNeedingEnable = @()
$methodsAlreadyEnabled = @()

foreach ($method in $LegacyMfaData.CurrentMethodDistribution.Keys) {
    $policyKey = switch ($method) {
        'MicrosoftAuthenticator' { 'MicrosoftAuthenticator' }
        'Voice/Phone' { 'Voice' }
        'SMS' { 'Sms' }
        'Email' { 'Email' }
        'Fido2' { 'Fido2' }
        'SoftwareOath' { 'SoftwareOath' }
        'WindowsHello' { 'WindowsHello' }
        'X509Certificate' { 'X509Certificate' }
        default { $null }
    }
    
    if ($policyKey) {
        if ($CurrentPolicy[$policyKey] -eq 'enabled') {
            $methodsAlreadyEnabled += "$method (preserves access for $($LegacyMfaData.CurrentMethodDistribution[$method]) users)"
        } else {
            $methodsNeedingEnable += "$method (preserves access for $($LegacyMfaData.CurrentMethodDistribution[$method]) users)"
        }
    }
}

# CORRECTED REPORT CONTENT with proper dual assessment
$reportContent = @"
MFA TO AUTHENTICATION METHODS POLICY MIGRATION REPORT
====================================================
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Tenant: $tenantName
Microsoft Deadline: September 30, 2025
Time Remaining: $((New-TimeSpan -Start (Get-Date) -End (Get-Date "2025-09-30")).Days) days

CORRECTED EXECUTIVE SUMMARY
==========================
Current MFA Coverage: $(if ($totalUsers -gt 0) { [math]::Round(($usersWithMfa / $totalUsers) * 100, 1) } else { 0 })% of users

MIGRATION READINESS ASSESSMENT:
=====================================
[OK] Phase 1 Ready: Zero disruption expected - All current MFA users will continue working
[OK] September 30th Deadline: Achievable without service interruption

SECURITY COMPLIANCE ASSESSMENT:
===============================
- Users with MFA (Compliant): $usersWithMfa
- Users without MFA: $($LegacyMfaData.RegularUsersNoMfa.Count) users $(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) { "[WARNING - Security Policy Gap]" } else { "[COMPLIANT]" })
- Privileged users without MFA: $privilegedNoMfaExcludingBreakGlass $(if ($privilegedNoMfaExcludingBreakGlass -gt 0) { "[CRITICAL - Security Violation]" } else { "[COMPLIANT]" })

DETAILED ANALYSIS
================
Total Users Analysed: $totalUsers
- Service Accounts: $($LegacyMfaData.ServiceAccountUsers.Count) (No action required)
- Regular Users with MFA: $usersWithMfa
  - Secure methods only: $($usersWithMfa - $usersForAutomaticCleanup - $usersNeedingAssistance) [COMPLIANT]
  - Can auto-remove insecure: $usersForAutomaticCleanup users [COMPLIANT - Enhancement available]
  - Need method upgrade assistance: $usersNeedingAssistance users [COMPLIANT - Enhancement recommended]

SECURITY POLICY COMPLIANCE
===========================
$(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) {
"[WARNING] Regular users without MFA: $($LegacyMfaData.RegularUsersNoMfa.Count) users
- These users currently access systems with password-only
- They will continue to have password-only access after migration
- Security policy recommends all users should have MFA protection"
} else {
"[OK] All regular users have MFA protection"
})

$(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
"[CRITICAL] Privileged users without MFA: $privilegedNoMfaExcludingBreakGlass users  
- These administrator accounts currently lack MFA protection
- This represents a critical security vulnerability
- All privileged accounts must have MFA per security policy"
} else {
"[OK] All privileged users have MFA protection"
})

Privileged User Security Details:
- Total Privileged Users: $(if ($PrivilegedData) { ($PrivilegedData.AnalyzedPrivilegedUsers | Select-Object -ExpandProperty UserPrincipalName -Unique).Count } else { 0 })
- Without Any MFA: $privilegedNoMfaExcludingBreakGlass $(if ($privilegedNoMfaExcludingBreakGlass -gt 0) { "[CRITICAL - Security Violation]" } else { "[COMPLIANT]" })
- Without FIDO2: $privilegedNoFido $(if ($privilegedNoFido -gt 0) { "[Enhancement Opportunity]" } else { "[SECURE]" })
- Managed Break-Glass: $(($PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { $_.UserPrincipalName -in $breakGlassAccounts } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count) (Will configure during migration)

DUAL TIMELINE ASSESSMENT
========================

MIGRATION TIMELINE (September 30, 2025 Deadline)
----------------------------------------------
[OK] ZERO DISRUPTION EXPECTED
- All $usersWithMfa users with current MFA will continue working normally
- $($LegacyMfaData.RegularUsersNoMfa.Count) users without MFA will continue working exactly as they do today
- No service interruptions anticipated
- September 30th deadline is achievable

SECURITY COMPLIANCE TIMELINE (Ongoing Organizational Requirement)
---------------------------------------------------------------
$(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
"[CRITICAL] PRIVILEGED USER SECURITY VIOLATIONS:
- $privilegedNoMfaExcludingBreakGlass privileged users lack MFA protection
- These are HIGH-VALUE TARGETS that must be protected
- Represents immediate organizational security risk
- Recommended action: Implement MFA for all privileged accounts within 30 days"
} else {
"[OK] PRIVILEGED USERS: All protected with MFA"
})

$(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) {
"[WARNING] SECURITY POLICY GAPS:
- $($LegacyMfaData.RegularUsersNoMfa.Count) regular users lack MFA protection
- Every user should have MFA per security best practices
- Represents moderate organizational security risk  
- Recommended action: Implement MFA for all users within 90 days"
} else {
"[OK] REGULAR USERS: All protected with MFA"
})

KEY DISTINCTION
--------------
- Migration Timeline: No urgency - zero disruption expected for September 30th deadline
- Security Compliance: Ongoing concern requiring attention per organizational security policy
- These are separate issues with different timelines and priorities

OUR TWO-PHASE MIGRATION APPROACH
================================
We will manage your migration in two phases to ensure zero disruption while maintaining security standards.

PHASE 1: MEET THE DEADLINE - PREVENT DISRUPTION (By September 30, 2025)
-----------------------------------------------------------------------
Time: 1-2 days
Target: Zero disruption for users who currently have MFA

[OK] WHAT WILL WORK:
- All $usersWithMfa users with current MFA will continue accessing systems normally
- Users without MFA will continue working exactly as they do today (password-only)
- No users will be locked out or lose access they currently have
- September 30th deadline compliance achieved

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
"Methods already enabled:
$($methodsAlreadyEnabled | ForEach-Object { "  - $_" })
"
})
- Migrate policy management to new Authentication Methods Policy interface
- Test that all current MFA users can continue authenticating
- Document any users who don't currently have MFA (for optional Phase 2 security enhancement)

SUCCESS CRITERIA:
- Zero user complaints about lost access
- All current MFA users continue working normally  
- Users without MFA continue working as they always have
- Compliance with Microsoft deadline achieved

PHASE 2: SECURITY ENHANCEMENT (4-6 weeks after Phase 1)
-------------------------------------------------------
Time: 4-6 weeks
Focus: Improve security posture and address policy gaps

Week 1-2: Automatic Security Improvements
- Automatically remove insecure methods for $usersForAutomaticCleanup users
  (These users already have secure alternatives registered)
- No user action needed for this group

Week 3-4: Security Gap Remediation
$(if ($usersNeedingAssistance -gt 0) {
"- Work with $usersNeedingAssistance users to register Microsoft Authenticator
- Provide training and support materials
- Track progress and follow up"
} else {
"- No users need assistance - all have secure methods"
})

$(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) {
"- Address security policy gap: $($LegacyMfaData.RegularUsersNoMfa.Count) users without any MFA
- Register Microsoft Authenticator for improved security
- Priority based on user roles and access levels"
} else {
"- No security gaps to address - all users have MFA"
})

Week 5-6: Privileged User Security Enhancement
$(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
"- PRIORITY: Register MFA for $privilegedNoMfaExcludingBreakGlass privileged users (Critical security gap)
- Deploy FIDO2 security keys to $privilegedNoFido administrators (Recommended)
- Configure phishing-resistant MFA requirements
- Create conditional access policies for admin protection"
} else {
"- Deploy FIDO2 security keys to $privilegedNoFido administrators (Optional enhancement)  
- Configure phishing-resistant MFA requirements
- Create conditional access policies for admin protection"
})

Final Step: Remove Insecure Methods
- Remove Voice, SMS, and Email from the authentication policy
- Monitor for any issues
- Maintain exceptions only where absolutely necessary

GENERATED REPORTS
================
Location: $reportFolder

[DATA] User Details and Phase 2 Actions: $(Split-Path $excelPath -Leaf)
- Complete user list with current methods
- Migration impact assessment per user
- Security compliance status per user  
- Use for project planning and tracking
- Enhanced with dual assessment (migration vs security)

$(if ($privExcelPath) { "[SECURITY] Privileged User Security: $(Split-Path $privExcelPath -Leaf)
- Administrator account analysis
- Risk assessment and FIDO2 recommendations
- Critical security gap identification
" })
[REPORT] Migration Report: $(Split-Path $reportPath -Leaf)
- This comprehensive assessment report
- Executive summary and recommendations
- Phase-by-phase implementation plan

NEXT STEPS
==========
1. IMMEDIATE (This Week):
   $(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
   "   [CRITICAL] Address $privilegedNoMfaExcludingBreakGlass privileged users without MFA (security violation)"
   } else {
   "   [OK] No critical security issues requiring immediate attention"
   })

2. MIGRATION PREPARATION (Next 2 Weeks):
   - Review and approve Phase 1 plan (zero disruption expected)
   - Schedule policy migration for before September 30, 2025
   - Prepare communications for Phase 2 security enhancements

3. POST-MIGRATION SECURITY ENHANCEMENT (Optional):
   $(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0 -or $usersNeedingAssistance -gt 0) {
   "   - Plan security improvement campaign for $($LegacyMfaData.RegularUsersNoMfa.Count + $usersNeedingAssistance) users
   - Develop user training and support materials
   - Schedule assisted MFA registration sessions"
   } else {
   "   - All users already have secure MFA methods
   - Focus on FIDO2 deployment for administrators"
   })

This assessment confirms that your migration can proceed safely with zero disruption while clearly identifying opportunities for security enhancement.
"@

# Create Word document using exact working function from backup
$wordCreated = New-PSWriteWordReport -Content $reportContent -FilePath $reportPath -Title "MFA TO AUTHENTICATION METHODS POLICY MIGRATION REPORT"
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