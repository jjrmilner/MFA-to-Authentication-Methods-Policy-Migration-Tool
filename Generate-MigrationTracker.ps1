<#
.SYNOPSIS
    Generates a consolidated migration tracker Excel workbook for managing Phase 2 migration
    Creates Word documents using PSWriteWord for proper Word output
    
.DESCRIPTION
    Creates a tracking Excel workbook that combines all users needing action in Phase 2
    with clear priorities and weekly assignments, plus a Word document for critical users
    
.PARAMETER AssessmentData
    The assessment data containing user information
    
.PARAMETER OutputPath
    Path where tracker file will be saved
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [hashtable]$AssessmentData,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "."
)

Write-Host "`n=== GENERATING MIGRATION TRACKER ===" -ForegroundColor Cyan

# Check for modules
$useExcel = $false
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    $useExcel = $true
}

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

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$trackerPath = Join-Path $OutputPath "MFA_Migration_Tracker_$timestamp.xlsx"
if (-not $useExcel) {
    $trackerPath = $trackerPath -replace '\.xlsx$', '.csv'
}

# Get data from assessment
$LegacyMfaData = $AssessmentData.LegacyMfaData
$PrivilegedData = $AssessmentData.PrivilegedData
$Recommendations = $AssessmentData.Recommendations

$tracker = @()

# Process all users and create tracking entries
foreach ($upn in $LegacyMfaData.UserMfaStatus.Keys) {
    $userStatus = $LegacyMfaData.UserMfaStatus[$upn]
    
    # Skip users who don't need action
    if ($userStatus.Phase2Action -eq "No Action Required" -or 
        $userStatus.Status -eq "Service Account" -or 
        $userStatus.Status -eq "Disabled Account") {
        continue
    }
    
    # Check if user is privileged
    $isPrivileged = $PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { $_.UserPrincipalName -eq $upn }
    $privilegedInfo = ""
    if ($isPrivileged) {
        $roles = ($isPrivileged | Select-Object -ExpandProperty RoleName -Unique) -join ", "
        $privilegedInfo = " (Admin: $roles)"
    }
    
    # Create display name from UPN
    $displayName = $upn.Split('@')[0] -replace '\.', ' '
    $displayName = (Get-Culture).TextInfo.ToTitleCase($displayName)
    
    # Determine user type
    $userType = if ($isPrivileged) { 
        if ($isPrivileged.RoleName -match "Global Administrator") { "Privileged-GlobalAdmin" }
        else { "Privileged" }
    } else { "Regular" }
    
    # Adjust priority for privileged users
    $priority = $userStatus.Phase2Priority
    if ($isPrivileged) {
        $priority = "Critical"
    }
    
    # Determine assigned team
    $assignedTo = switch ($priority) {
        "Critical" { "Security Team" }
        "High" { "Help Desk Priority" }
        "Medium" { "Help Desk Team" }
        default { "Help Desk Team" }
    }
    
    # Add to tracker
    $tracker += [PSCustomObject]@{
        UserPrincipalName = $upn
        Display_Name = $displayName
        User_Type = $userType
        Current_Methods = $userStatus.Methods
        Methods_To_Remove = $userStatus.MethodsToRemove
        Action_Required = $userStatus.Phase2Action
        Priority = $priority
        Week = $userStatus.Phase2Week
        Status = "Not started"
        Last_Updated = Get-Date -Format 'yyyy-MM-dd'
        Assigned_To = $assignedTo
        Contact_Attempts = 0
        Last_Contact_Date = ""
        Notes = $privilegedInfo
    }
}

# Add users with no MFA at all (highest priority)
foreach ($user in $LegacyMfaData.RegularUsersNoMfa) {
    $upn = $user.UserPrincipalName
    $displayName = $upn.Split('@')[0] -replace '\.', ' '
    $displayName = (Get-Culture).TextInfo.ToTitleCase($displayName)
    
    # Check if privileged
    $isPrivileged = $PrivilegedData.AnalyzedPrivilegedUsers | Where-Object { $_.UserPrincipalName -eq $upn }
    $userType = if ($isPrivileged) { "Privileged-NoMFA" } else { "Regular-NoMFA" }
    
    $tracker += [PSCustomObject]@{
        UserPrincipalName = $upn
        Display_Name = $displayName
        User_Type = $userType
        Current_Methods = "Password Only"
        Methods_To_Remove = ""
        Action_Required = "Enable MFA - Critical"
        Priority = "Critical"
        Week = "1"
        Status = "Not started"
        Last_Updated = Get-Date -Format 'yyyy-MM-dd'
        Assigned_To = if ($isPrivileged) { "Security Team - URGENT" } else { "Help Desk Priority" }
        Contact_Attempts = 0
        Last_Contact_Date = ""
        Notes = "NO MFA - Immediate action required"
    }
}

# Sort by priority and week
$tracker = $tracker | Sort-Object @{e={
    switch($_.Priority) {
        'Critical' { 1 }
        'High' { 2 }
        'Medium' { 3 }
        'Low' { 4 }
        'None' { 5 }
    }
}}, @{e={
    switch($_.Week) {
        '1' { 1 }
        '1-2' { 2 }
        '3-4' { 3 }
        '5-6' { 4 }
        default { 99 }
    }
}}, UserPrincipalName

# Calculate statistics
$stats = @{
    Total = $tracker.Count
    Critical = ($tracker | Where-Object { $_.Priority -eq "Critical" }).Count
    High = ($tracker | Where-Object { $_.Priority -eq "High" }).Count
    Medium = ($tracker | Where-Object { $_.Priority -eq "Medium" }).Count
    Week1_2 = ($tracker | Where-Object { $_.Week -in @("1", "1-2") }).Count
    Week3_4 = ($tracker | Where-Object { $_.Week -eq "3-4" }).Count
    NoMFA = ($tracker | Where-Object { $_.User_Type -like "*NoMFA" }).Count
    Privileged = ($tracker | Where-Object { $_.User_Type -like "Privileged*" }).Count
}

# Export tracker
if ($useExcel) {
    # Create Excel workbook with multiple sheets
    Remove-Item $trackerPath -ErrorAction SilentlyContinue
    
    # Main tracking sheet
    $tracker | Export-Excel -Path $trackerPath -WorksheetName "Migration Tracker" -AutoSize -TableName "MigrationTracker" -TableStyle Medium9 -FreezeTopRow
    
    # Summary sheet
    $summaryData = @(
        [PSCustomObject]@{ Metric = "Total Users to Track"; Count = $stats.Total; Category = "Overall" }
        [PSCustomObject]@{ Metric = "Critical Priority"; Count = $stats.Critical; Category = "Priority" }
        [PSCustomObject]@{ Metric = "High Priority"; Count = $stats.High; Category = "Priority" }
        [PSCustomObject]@{ Metric = "Medium Priority"; Count = $stats.Medium; Category = "Priority" }
        [PSCustomObject]@{ Metric = "Week 1-2 Actions"; Count = $stats.Week1_2; Category = "Timeline" }
        [PSCustomObject]@{ Metric = "Week 3-4 Actions"; Count = $stats.Week3_4; Category = "Timeline" }
        [PSCustomObject]@{ Metric = "Users with NO MFA"; Count = $stats.NoMFA; Category = "Risk" }
        [PSCustomObject]@{ Metric = "Privileged Users Needing Action"; Count = $stats.Privileged; Category = "Risk" }
    )
    
    $summaryData | Export-Excel -Path $trackerPath -WorksheetName "Summary" -AutoSize -TableName "Summary" -TableStyle Light15
    
    # Critical users sheet
    $criticalUsers = $tracker | Where-Object { $_.Priority -eq "Critical" }
    if ($criticalUsers.Count -gt 0) {
        $criticalUsers | Export-Excel -Path $trackerPath -WorksheetName "Critical Users" -AutoSize -TableName "CriticalUsers" -TableStyle Medium10 -FreezeTopRow
    }
    
    Write-Host "Migration tracker Excel workbook saved to: $trackerPath" -ForegroundColor Green
} else {
    $tracker | Export-Csv -Path $trackerPath -NoTypeInformation
    Write-Host "Migration tracker CSV saved to: $trackerPath" -ForegroundColor Green
}

# Generate summary statistics
Write-Host "`nMigration Tracker Summary:" -ForegroundColor Yellow
Write-Host "Total users to track: $($stats.Total)" -ForegroundColor White
Write-Host "  Critical priority: $($stats.Critical) users" -ForegroundColor Red
Write-Host "  High priority: $($stats.High) users" -ForegroundColor Yellow
Write-Host "  Medium priority: $($stats.Medium) users" -ForegroundColor Green
Write-Host "  Week 1-2 actions: $($stats.Week1_2) users" -ForegroundColor White
Write-Host "  Week 3-4 actions: $($stats.Week3_4) users" -ForegroundColor White
Write-Host "  Users with NO MFA: $($stats.NoMFA) users" -ForegroundColor Red
Write-Host "  Privileged users needing action: $($stats.Privileged) users" -ForegroundColor Yellow

# Function to create Word document using PSWriteWord with stub path
function New-TrackerWordDoc {
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
            Add-WordText -WordDocument $WordDocument -Text $Title -FontSize 16 -Bold $true -SpacingAfter 15 -Color "Red" -Supress $true
        }
        
        # Process content line by line
        $lines = $Content -split "`n"
        
        foreach ($line in $lines) {
            # Skip empty lines
            if ([string]::IsNullOrWhiteSpace($line)) {
                Add-WordText -WordDocument $WordDocument -Text "" -SpacingAfter 6 -Supress $true
                continue
            }
            
            # Section headers (lines with : at the end)
            if ($line -match "^[A-Z][A-Z\s\d\-()]+:$" -and $line.Length -gt 5) {
                Add-WordText -WordDocument $WordDocument -Text $line -FontSize 12 -Bold $true -SpacingBefore 10 -SpacingAfter 8 -Supress $true
                continue
            }
            
            # Headers with underlines (===)
            if ($lines.IndexOf($line) + 1 -lt $lines.Count -and $lines[$lines.IndexOf($line) + 1] -match "^={3,}$") {
                Add-WordText -WordDocument $WordDocument -Text $line -FontSize 14 -Bold $true -SpacingBefore 12 -SpacingAfter 12 -Supress $true
                continue
            }
            
            # Skip separator lines
            if ($line -match "^={3,}$" -or $line -match "^-{3,}$") {
                continue
            }
            
            # Bullet points or list items
            if ($line -match "^\d+\.\s+" -or $line -match "^[-•]\s+") {
                $cleanLine = $line -replace "^\d+\.\s+", "" -replace "^[-•]\s+", ""
                Add-WordText -WordDocument $WordDocument -Text "• $cleanLine" -SpacingAfter 6 -IndentationFirstLine 1 -Supress $true
                continue
            }
            
            # User entries (lines with @ symbol)
            if ($line -match "@") {
                Add-WordText -WordDocument $WordDocument -Text $line -FontFamily "Courier New" -FontSize 10 -SpacingAfter 3 -Supress $true
                continue
            }
            
            # Total/summary lines
            if ($line -match "^Total") {
                Add-WordText -WordDocument $WordDocument -Text $line -Bold $true -SpacingBefore 8 -SpacingAfter 6 -Supress $true
                continue
            }
            
            # Regular text
            Add-WordText -WordDocument $WordDocument -Text $line -SpacingAfter 6 -Supress $true
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

# Generate critical users Word document
$criticalUsersPath = Join-Path $OutputPath "Critical_Users_Immediate_Action_$timestamp.docx"
$criticalContent = @"
CRITICAL USERS REQUIRING IMMEDIATE ACTION
=========================================
Generated: $(Get-Date)

USERS WITH NO MFA AT ALL:
------------------------
$($tracker | Where-Object { $_.Current_Methods -eq "Password Only" } | ForEach-Object {
    "$($_.UserPrincipalName) - $($_.User_Type)"
} | Out-String)

PRIVILEGED USERS NEEDING SECURE MFA:
------------------------------------
$($tracker | Where-Object { $_.User_Type -like "Privileged*" -and $_.Current_Methods -ne "Password Only" } | ForEach-Object {
    "$($_.UserPrincipalName) - Current: $($_.Current_Methods)"
} | Out-String)

Total Critical Actions: $($stats.Critical)

RECOMMENDED IMMEDIATE ACTIONS:
1. Contact all users with NO MFA immediately
2. Schedule FIDO2 deployment for privileged users
3. Begin Phase 2 preparation with automatic cleanup group
"@

# Create Word document for critical users
$wordCreated = New-TrackerWordDoc -Content $criticalContent -FilePath $criticalUsersPath -Title "CRITICAL USERS REQUIRING IMMEDIATE ACTION"
if ($wordCreated) {
    Write-Host "Critical users list saved to: $criticalUsersPath" -ForegroundColor Red
} else {
    # Fall back to text file
    $criticalUsersPath = $criticalUsersPath -replace '\.docx$', '.txt'
    $criticalContent | Out-File -FilePath $criticalUsersPath -Encoding UTF8
    Write-Host "Critical users list saved to: $criticalUsersPath" -ForegroundColor Red
}

return $trackerPath