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
=====================================================
Generated: $(Get-Date -Format "dddd, MMMM dd, yyyy 'at' HH:mm:ss")
Organisation: $tenantName
Assessment Type: Migration Impact and Security Compliance

This report provides a comprehensive assessment of your organisation's readiness to migrate from Legacy MFA to Authentication Methods Policy, with clear separation between migration disruption risk and security compliance requirements.

EXECUTIVE SUMMARY
================

MIGRATION DISRUPTION ASSESSMENT: [OK] ZERO DISRUPTION EXPECTED
-------------------------------------------------------------
Your organisation can safely migrate to Authentication Methods Policy by the September 30, 2025 deadline with ZERO service disruption expected.

Key Migration Reality:
- Users with Current MFA: $usersWithMfa (Will continue working normally after migration)  
- Users without MFA: $($LegacyMfaData.RegularUsersNoMfa.Count) (Will continue working exactly as they do today - no change)
- Privileged Users: $(if ($PrivilegedData) { ($PrivilegedData.AnalyzedPrivilegedUsers | Select-Object -ExpandProperty UserPrincipalName -Unique).Count } else { 0 }) (Migration impact assessed separately)

Critical Understanding: Users without MFA are not at risk of losing access during migration because they never had MFA protection to begin with. They will continue using password-only authentication exactly as they do today.

SECURITY COMPLIANCE ASSESSMENT: $(if ($privilegedNoMfaExcludingBreakGlass -gt 0 -or $LegacyMfaData.RegularUsersNoMfa.Count -gt 0) { "[ACTION REQUIRED] Security Policy Gaps Identified" } else { "[OK] Full Security Compliance" })
------------------------------------------------------------
Separate from migration, we identified ongoing security policy compliance gaps requiring attention:

$(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
"[CRITICAL] $privilegedNoMfaExcludingBreakGlass privileged users lack MFA protection (High-value security targets)"
} else {
"[OK] All privileged users have MFA protection"
})

$(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) {
"[WARNING] $($LegacyMfaData.RegularUsersNoMfa.Count) regular users lack MFA protection (Security policy enhancement opportunity)"
} else {
"[OK] All regular users have MFA protection"
})

RECOMMENDATION
-------------
Proceed with migration immediately (zero disruption risk) whilst separately addressing security policy compliance gaps through a structured enhancement programme.

CURRENT AUTHENTICATION LANDSCAPE
===============================

MFA Usage Overview:
- Total Licensed Users: $(if ($LegacyMfaData.TotalUsers) { $LegacyMfaData.TotalUsers } else { "Data unavailable" })
- Users with MFA Registered: $usersWithMfa ($(if ($LegacyMfaData.TotalUsers -and $LegacyMfaData.TotalUsers -gt 0) { [math]::Round(($usersWithMfa / $LegacyMfaData.TotalUsers) * 100, 1) } else { "N/A" })%)
- Users without Any MFA: $($LegacyMfaData.RegularUsersNoMfa.Count) ($(if ($LegacyMfaData.TotalUsers -and $LegacyMfaData.TotalUsers -gt 0) { [math]::Round(($LegacyMfaData.RegularUsersNoMfa.Count / $LegacyMfaData.TotalUsers) * 100, 1) } else { "N/A" })%)

Current Method Distribution:
$(if ($LegacyMfaData.MethodStats) { $LegacyMfaData.MethodStats | ForEach-Object { "- $($_.Method): $($_.Count) users" } } else { "Method statistics not available" })

Authentication Policy Status:
- Current Policy: $($CurrentPolicy.PolicyType)
- Available Methods: $(if ($CurrentPolicy.EnabledMethods) { $CurrentPolicy.EnabledMethods -join ", " } else { "Not specified" })
- Tenant Default: $(if ($CurrentPolicy.IsDefault) { "Yes" } else { "No" })

Migration Impact Analysis:
$(if ($methodsNeedingEnable.Count -gt 0) {
"[ACTION REQUIRED] Enable these methods in the new policy:
$($methodsNeedingEnable | ForEach-Object { "- $_" } | Out-String)"
} else {
"[OK] All required methods are already enabled in the target policy"
})

DETAILED SECURITY COMPLIANCE ANALYSIS
====================================

Regular User Security Posture:
$(if ($LegacyMfaData.RegularUsersNoMfa.Count -gt 0) {
"[WARNING] SECURITY POLICY ENHANCEMENT OPPORTUNITY:
- $($LegacyMfaData.RegularUsersNoMfa.Count) users currently have no MFA protection
- These users rely solely on password authentication (current state - no change during migration)
- Represents ongoing security enhancement opportunity
- Every user should have MFA per security best practices
- Note: These users will continue working normally - this is a security improvement initiative, not a migration blocker"
} else {
"[OK] All regular users have MFA protection"
})

User Enhancement Opportunities:
$(if ($usersNeedingAssistance -gt 0) {
"[ENHANCEMENT] MFA Method Improvement Opportunity:
- $usersNeedingAssistance users currently have only less secure methods (SMS/Voice/Email)
- Should upgrade to Microsoft Authenticator for better security
- Current methods will continue working during and after migration with no disruption
- This is a security enhancement opportunity, not a migration requirement"
} else {
"[OK] All users with MFA have secure authentication methods"
})

Privileged User Security Analysis:
$(if ($privilegedNoMfaExcludingBreakGlass -gt 0) {
"[CRITICAL] SECURITY COMPLIANCE VIOLATION:
- $privilegedNoMfaExcludingBreakGlass privileged users lack MFA protection
- These are HIGH-VALUE TARGETS requiring immediate security attention
- Represents ongoing organisational security compliance risk
- All privileged accounts must have MFA per security policy
- Note: Migration can proceed safely, but this security gap requires urgent attention"
} else {
"[OK] All privileged users have MFA protection"
})

Privileged User Security Details:
- Total Privileged Users: $(if ($PrivilegedData) { ($PrivilegedData.AnalyzedPrivilegedUsers | Select-Object -ExpandProperty UserPrincipalName -Unique).Count } else { 0 })
- Without Any MFA: $privilegedNoMfaExcludingBreakGlass $(if ($privilegedNoMfaExcludingBreakGlass -gt 0) { "[CRITICAL - Security Compliance Violation]" } else { "[COMPLIANT]" })
- Without FIDO2: $privilegedNoFido $(if ($privilegedNoFido -gt 0) { "[Enhancement Opportunity]" } else { "[SECURE]" })
- Managed Break-Glass: $(($PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { $_.UserPrincipalName -in $breakGlassAccounts } | Select-Object -ExpandProperty UserPrincipalName -Unique).Count) (Will configure during migration)

DUAL TIMELINE ASSESSMENT
========================

MIGRATION TIMELINE (September 30, 2025 Deadline)
----------------------------------------------
[OK] ZERO DISRUPTION EXPECTED - MIGRATION CAN PROCEED SAFELY
- All $usersWithMfa users with current MFA will continue working normally after migration
- $($LegacyMfaData.RegularUsersNoMfa.Count) users without MFA will continue working exactly as they do today (password-only, no change)
- No users will lose access they currently have
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
