<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

.SYNOPSIS
    Generates Phase 2 user communication templates using PSWriteWord for proper Word output
    
.DESCRIPTION
    Creates personalised email templates in Word format for users needing assistance with MFA migration
    
.PARAMETER AssessmentData
    The assessment data containing user information
    
.PARAMETER OutputPath
    Path where communication files will be saved
#>

param(
    [Parameter(Mandatory=$true)]
    [hashtable]$AssessmentData,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "."
)

Write-Host "`n=== GENERATING USER COMMUNICATION TEMPLATES ===" -ForegroundColor Cyan

# Check for modules
$useExcel = $false
$useWord = $false

if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    $useExcel = $true
}

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

# Get data from assessment
$LegacyMfaData = $AssessmentData.LegacyMfaData
$Recommendations = $AssessmentData.Recommendations

# Get tenant name
$tenantName = try {
    (Get-MgOrganization).DisplayName
} catch {
    "Your Organization"
}

# Create output folder
$communicationFolder = Join-Path $OutputPath "MFA_User_Communications_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $communicationFolder -Force | Out-Null

# Function to create communication Word document using PSWriteWord with stub path
function New-CommunicationWordDoc {
    param(
        [string]$Content,
        [string]$FilePath,
        [string]$Title = ""
    )
    
    if (-not $useWord) {
        return $false
    }
    
    # Create a temporary stub path
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
            Add-WordText -WordDocument $WordDocument -Text $Title -FontSize 16 -Bold $true -SpacingAfter 15 -Color "DarkBlue"
        }
        
        # Process content line by line
        $lines = $Content -split "`n"
        $currentList = @()
        $inList = $false
        
        foreach ($line in $lines) {
            # Handle empty lines
            if ([string]::IsNullOrWhiteSpace($line)) {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList
                    $currentList = @()
                    $inList = $false
                }
                Add-WordText -WordDocument $WordDocument -Text "" -SpacingAfter 6
                continue
            }
            
            # Subject line
            if ($line -match "^Subject:") {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList
                    $currentList = @()
                    $inList = $false
                }
                Add-WordText -WordDocument $WordDocument -Text $line -FontSize 12 -Bold $true -SpacingAfter 10
                continue
            }
            
            # Headers (text between **)
            if ($line -match "\*\*(.*?)\*\*") {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList
                    $currentList = @()
                    $inList = $false
                }
                
                # Extract header text and remove the **
                $cleanedLine = $line -replace '\*\*(.*?)\*\*', '$1'
                Add-WordText -WordDocument $WordDocument -Text $cleanedLine -Bold $true -SpacingAfter 8
                continue
            }
            
            # Markdown headers
            if ($line -match "^#{1,3}\s(.+)") {
                if ($inList -and $currentList.Count -gt 0) {
                    Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList
                    $currentList = @()
                    $inList = $false
                }
                
                $headerLevel = ($line | Select-String -Pattern "^#{1,3}").Matches[0].Value.Length
                $headerText = $line -replace "^#{1,3}\s", ""
                
                $fontSize = switch($headerLevel) {
                    1 { 16 }
                    2 { 14 }
                    3 { 12 }
                    default { 12 }
                }
                
                Add-WordText -WordDocument $WordDocument -Text $headerText -FontSize $fontSize -Bold $true -SpacingBefore 10 -SpacingAfter 8
                continue
            }
            
            # Bullet points
            if ($line -match "^[-•]\s+" -or $line -match "^\d+\.\s+") {
                $inList = $true
                $listItem = $line -replace "^[-•]\s+", "" -replace "^\d+\.\s+", ""
                $currentList += $listItem.Trim()
                continue
            }
            
            # End list if we were in one
            if ($inList -and $currentList.Count -gt 0) {
                Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList
                $currentList = @()
                $inList = $false
            }
            
            # Questions (lines starting with Q:)
            if ($line -match "^Q:") {
                Add-WordText -WordDocument $WordDocument -Text $line -Bold $true -SpacingAfter 6
                continue
            }
            
            # Answers (lines starting with A:)
            if ($line -match "^A:") {
                Add-WordText -WordDocument $WordDocument -Text $line -SpacingAfter 8
                continue
            }
            
            # Regular paragraph
            Add-WordText -WordDocument $WordDocument -Text $line -SpacingAfter 6
        }
        
        # Handle any remaining list items
        if ($inList -and $currentList.Count -gt 0) {
            Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $currentList
        }
        
        Save-WordDocument -WordDocument $WordDocument -Language 'en-US'
        return $true
    }
    catch {
        Write-Warning "Failed to create Word document: $_"
        return $false
    }
    finally {
        # Clean up stub path
        if ($stubCreated -and (Test-Path $stubPath)) {
            try {
                cmd /c rmdir "$stubPath" 2>&1 | Out-Null
            }
            catch {
                Write-Warning "Could not remove stub path: $stubPath"
            }
        }
    }
}

# Function to generate individual user email
function Generate-UserEmail {
    param(
        [string]$UserPrincipalName,
        [string]$CurrentMethods,
        [string]$TemplateName,
        [datetime]$Deadline
    )
    
    $firstName = $UserPrincipalName.Split('@')[0].Split('.')[0]
    $firstName = (Get-Culture).TextInfo.ToTitleCase($firstName)
    
    $methodsList = if ($CurrentMethods -match "Phone") {
        "text message codes"
    } elseif ($CurrentMethods -match "Email") {
        "email codes"
    } else {
        "your current authentication method"
    }
    
    $email = @"
Subject: Action Required: Upgrade Your Account Security - Microsoft Authenticator Setup

Dear $firstName,

As part of our ongoing commitment to protecting your account and company data, we're upgrading our authentication systems to meet new Microsoft security requirements.

**What's Changing:**
Starting $(Get-Date $Deadline -Format 'MMMM dd, yyyy'), we'll be removing less secure authentication methods and transitioning all users to the Microsoft Authenticator app, which provides better protection against cyber threats.

**Why This Matters:**
- Your current method ($methodsList) can be intercepted by attackers
- Microsoft Authenticator reduces account compromise risk by 99.9%
- This change is mandatory due to Microsoft's new authentication requirements

**What You Need to Do:**
You currently use $methodsList for multi-factor authentication. You'll need to set up Microsoft Authenticator before $(Get-Date $Deadline -Format 'MMMM dd').

**We're Here to Help:**
- IT Helpdesk: support@globalmicro.co.za

**Quick Setup Steps:**
1. Download Microsoft Authenticator from your app store
2. Open the app and tap "Add account"
3. Choose "Work or school account"
4. Sign in with your email: $UserPrincipalName
5. Follow the prompts (takes about 5 minutes)

Please complete this setup by $(Get-Date $Deadline -Format 'MMMM dd') to avoid any interruption to your account access.

Thank you for helping us keep $tenantName secure.

Best regards,
IT Security Team
$tenantName
"@
    
    return $email
}

# Generate setup guide
$setupGuide = @"
# Microsoft Authenticator Setup Guide

## Quick Setup (5 Minutes)

### Step 1: Download the App
**iPhone Users:**
1. Open the App Store
2. Search for "Microsoft Authenticator"
3. Look for the blue icon with a shield
4. Tap "Get" to download

**Android Users:**
1. Open Google Play Store
2. Search for "Microsoft Authenticator"
3. Look for the blue icon with a shield
4. Tap "Install"

### Step 2: Add Your Work Account
1. Open Microsoft Authenticator
2. Tap the "+" button or "Add account"
3. Select "Work or school account"
4. Sign in with your work email address
5. Enter your current password
6. Complete your current MFA challenge (text/email code)

### Step 3: Verify Setup
1. You'll see a QR code on your computer screen
2. Tap "Scan QR code" in the app
3. Point your phone camera at the QR code
4. The app will automatically add your account

### Step 4: Test Your Setup
1. The system will send a test notification
2. Tap "Approve" on your phone
3. You're all set!

## Troubleshooting

**Can't find the app?**
- Make sure you search for "Microsoft Authenticator" (not just "Authenticator")
- Look for the publisher "Microsoft Corporation"

**QR code won't scan?**
- Make sure you've allowed camera permissions
- Try entering the code manually (there's an option below the QR code)

**Didn't receive notification?**
- Check that notifications are enabled for the app
- Make sure your phone has internet connection
- Try tapping "Refresh" in the app

## Need Help?
- Email: support@globalmicro.co.za

## Frequently Asked Questions

Q: Do I need to pay for the app?
A: No, Microsoft Authenticator is completely free.

Q: What if I get a new phone?
A: You can easily transfer your account to a new device. Contact IT for assistance.

Q: Can I still use the app without internet?
A: Yes! The app generates codes that work even offline.

Q: What if I lose my phone?
A: Contact IT immediately. We have backup methods to help you access your account.
"@

# Save setup guide as Word document
$setupGuidePath = Join-Path $communicationFolder "MFA_Setup_Guide.docx"
$wordCreated = New-CommunicationWordDoc -Content $setupGuide -FilePath $setupGuidePath -Title "Microsoft Authenticator Setup Guide"
if (-not $wordCreated) {
    $setupGuidePath = $setupGuidePath -replace '\.docx$', '.txt'
    $setupGuide | Out-File -FilePath $setupGuidePath -Encoding UTF8
}

# Process users needing assistance
$usersNeedingAssistance = $Recommendations.UsersNeedingAssistance
$phase2StartDate = (Get-Date).AddDays(30)  # Assume Phase 2 starts 30 days after assessment
$deadline = $phase2StartDate.AddDays(28)  # 4 weeks for Phase 2

Write-Host "Generating communications for $($usersNeedingAssistance.Count) users..." -ForegroundColor Yellow

# Create batch email list
$batchEmails = @()

foreach ($user in $usersNeedingAssistance) {
    # Generate individual email
    $userEmail = Generate-UserEmail `
        -UserPrincipalName $user.UserPrincipalName `
        -CurrentMethods $user.CurrentMethods `
        -TemplateName "Initial" `
        -Deadline $deadline
    
    # Save individual email as Word document
    $safeFileName = $user.UserPrincipalName -replace '@', '_at_' -replace '[^\w\-\._]', '_'
    $emailPath = Join-Path $communicationFolder "Email_$safeFileName.docx"
    
    $wordCreated = New-CommunicationWordDoc -Content $userEmail -FilePath $emailPath -Title "Email Template for $($user.UserPrincipalName)"
    if (-not $wordCreated) {
        $emailPath = $emailPath -replace '\.docx$', '.txt'
        $userEmail | Out-File -FilePath $emailPath -Encoding UTF8
    }
    
    # Add to batch list
    $batchEmails += [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        FirstName = $user.UserPrincipalName.Split('@')[0].Split('.')[0]
        CurrentMethods = $user.CurrentMethods
        EmailFile = Split-Path $emailPath -Leaf
    }
}

# Generate follow-up templates
$followUpTemplate = @"
Subject: Reminder: Complete Your Microsoft Authenticator Setup by $(Get-Date $deadline.AddDays(-7) -Format 'MMMM dd')

Hi [FirstName],

This is a friendly reminder that you still need to set up Microsoft Authenticator on your mobile device.

**Quick Setup Steps:**
1. Download Microsoft Authenticator from your app store
2. Add your work account: [UserEmail]
3. Follow the prompts (takes about 5 minutes)

**Need Help?**
- Email: support@globalmicro.co.za

**Important:** After $(Get-Date $deadline -Format 'MMMM dd'), you won't be able to sign in using [CurrentMethod].

Thanks,
IT Team
"@

$followUpPath = Join-Path $communicationFolder "Template_FollowUp.docx"
$wordCreated = New-CommunicationWordDoc -Content $followUpTemplate -FilePath $followUpPath -Title "Follow-Up Email Template"
if (-not $wordCreated) {
    $followUpPath = $followUpPath -replace '\.docx$', '.txt'
    $followUpTemplate | Out-File -FilePath $followUpPath -Encoding UTF8
}

# Generate final notice template
$finalNoticeTemplate = @"
Subject: URGENT: Complete MFA Setup to Maintain Account Access

[FirstName],

Your account still uses [CurrentMethod] authentication, which will be disabled on $(Get-Date $deadline -Format 'MMMM dd').

**Action Required Today:**
Set up Microsoft Authenticator now to avoid being locked out of your account.

**Get Immediate Help:**
- Email IT Helpdesk: support@globalmicro.co.za

This is your final reminder. Please take action today.

IT Security Team
"@

$finalNoticePath = Join-Path $communicationFolder "Template_FinalNotice.docx"
$wordCreated = New-CommunicationWordDoc -Content $finalNoticeTemplate -FilePath $finalNoticePath -Title "Final Notice Email Template"
if (-not $wordCreated) {
    $finalNoticePath = $finalNoticePath -replace '\.docx$', '.txt'
    $finalNoticeTemplate | Out-File -FilePath $finalNoticePath -Encoding UTF8
}

# Generate Excel/CSV for mail merge
if ($useExcel) {
    $xlPath = Join-Path $communicationFolder "UserCommunicationList.xlsx"
    $batchEmails | Export-Excel -Path $xlPath -WorksheetName "Mail Merge" -AutoSize -TableName "UserList" -TableStyle Medium9
} else {
    $csvPath = Join-Path $communicationFolder "UserCommunicationList.csv"
    $batchEmails | Export-Csv -Path $csvPath -NoTypeInformation
}

# Generate summary report
$summaryReport = @"
MFA Phase 2 User Communications Summary
=======================================
Generated: $(Get-Date)
Total Users: $($usersNeedingAssistance.Count)

Users Requiring Communication:
$($batchEmails | ForEach-Object { "- $($_.UserPrincipalName) (Currently using: $($_.CurrentMethods))" } | Out-String)

Communication Timeline:
- Initial Email: Send on $(Get-Date $phase2StartDate -Format 'MMMM dd, yyyy')
- Follow-up Reminder: Send on $(Get-Date $deadline.AddDays(-7) -Format 'MMMM dd, yyyy')
- Final Notice: Send on $(Get-Date $deadline.AddDays(-2) -Format 'MMMM dd, yyyy')
- Deadline: $(Get-Date $deadline -Format 'MMMM dd, yyyy')

Files Generated:
- Individual emails for each user
- Follow-up template
- Final notice template
- Setup guide
- Excel/CSV for mail merge

Next Steps:
1. Review and customize the templates as needed
2. Schedule emails in your mail system
3. Prepare support resources for help sessions
4. Track completion rates weekly
"@

$summaryPath = Join-Path $communicationFolder "Communication_Summary.docx"
$wordCreated = New-CommunicationWordDoc -Content $summaryReport -FilePath $summaryPath -Title "Communication Plan Summary"
if (-not $wordCreated) {
    $summaryPath = $summaryPath -replace '\.docx$', '.txt'
    $summaryReport | Out-File -FilePath $summaryPath -Encoding UTF8
}

Write-Host "`nUser communications generated successfully!" -ForegroundColor Green
Write-Host "Location: $communicationFolder" -ForegroundColor Cyan
Write-Host "`nFiles created:" -ForegroundColor Yellow
Write-Host "  - $($batchEmails.Count) individual user emails" -ForegroundColor White
Write-Host "  - Follow-up and final notice templates" -ForegroundColor White
Write-Host "  - Setup guide for users" -ForegroundColor White
if ($useExcel) {
    Write-Host "  - Excel file for mail merge" -ForegroundColor White
} else {
    Write-Host "  - CSV list for mail merge" -ForegroundColor White
}
Write-Host "  - Communication summary report" -ForegroundColor White

return $communicationFolder
