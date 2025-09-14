<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

.SYNOPSIS
    Master controller for MFA Migration Assessment Tool - Clean Version
    
.DESCRIPTION
    This script orchestrates the MFA assessment by calling individual component scripts
    Enhanced to handle long paths by creating a temporary symbolic link when needed
    
.PARAMETER FullAnalysis
    Performs complete analysis
    
.PARAMETER PolicyOnly
    Only checks Authentication Methods Policy
    
.PARAMETER GenerateReport
    Generates assessment reports with enhanced tracking
    
.PARAMETER Help
    Shows help information
#>

[CmdletBinding()]
param(
    [switch]$FullAnalysis,
    [switch]$PolicyOnly,
    [switch]$GenerateReport,
    [switch]$Help
)

# Script metadata
$script:scriptVersion = "2.2"
$script:scriptAuthor = "JJ Milner"

# Function to display help
function Show-Help {
    Write-Host ""
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "     MFA to Authentication Methods Policy Migration Tool" -ForegroundColor Cyan
    Write-Host "                    Version: $script:scriptVersion (Enhanced)" -ForegroundColor Cyan
    Write-Host "               Author: $script:scriptAuthor" -ForegroundColor Cyan
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Analyzes your Microsoft 365 tenant's MFA configuration and generates"
    Write-Host "  reports to help with the migration to Authentication Methods Policy"
    Write-Host "  before the September 30, 2025 deadline."
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  -PolicyOnly" -ForegroundColor Green
    Write-Host "      Only checks the current Authentication Methods Policy status."
    Write-Host "      Quick check to see which methods are enabled/disabled."
    Write-Host ""
    Write-Host "  -FullAnalysis" -ForegroundColor Green
    Write-Host "      Performs complete analysis including:"
    Write-Host "      • Authentication Methods Policy check"
    Write-Host "      • User MFA registration analysis (all users)"
    Write-Host "      • Privileged user security assessment"
    Write-Host "      • Conditional Access policy review"
    Write-Host "      • Migration recommendations"
    Write-Host ""
    Write-Host "  -GenerateReport" -ForegroundColor Green
    Write-Host "      Generates detailed reports and CSV files (requires -FullAnalysis):"
    Write-Host "      • Executive summary report"
    Write-Host "      • User methods CSV with migration tracking fields"
    Write-Host "      • Privileged users security report"
    Write-Host "      • Migration tracker spreadsheet"
    Write-Host "      • Automation scripts for Phase 2"
    Write-Host "      • User communication templates (optional)"
    Write-Host ""
    Write-Host "  -Help" -ForegroundColor Green
    Write-Host "      Displays this help information."
    Write-Host ""
    Write-Host "USAGE EXAMPLES:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Quick policy check:" -ForegroundColor Cyan
    Write-Host "    .\mfa-master-script.ps1 -PolicyOnly"
    Write-Host ""
    Write-Host "  Full analysis without reports:" -ForegroundColor Cyan
    Write-Host "    .\mfa-master-script.ps1 -FullAnalysis"
    Write-Host ""
    Write-Host "  Complete assessment with all reports:" -ForegroundColor Cyan
    Write-Host "    .\mfa-master-script.ps1 -FullAnalysis -GenerateReport"
    Write-Host ""
    Write-Host "OUTPUT:" -ForegroundColor Yellow
    Write-Host "  Reports are saved in a timestamped folder:"
    Write-Host "  MFA_Reports_[TenantName]_[TenantID]\"
    Write-Host ""
    Write-Host "REQUIREMENTS:" -ForegroundColor Yellow
    Write-Host "  • Microsoft Graph PowerShell SDK"
    Write-Host "  • Global Administrator or appropriate read permissions"
    Write-Host "  • All component scripts in the same directory"
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Check if no parameters provided or Help requested
if ($Help -or (-not $FullAnalysis -and -not $PolicyOnly -and -not $GenerateReport)) {
    Show-Help
    return
}

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "     MFA to Authentication Methods Policy Migration Tool" -ForegroundColor Cyan
Write-Host "                    Version: $script:scriptVersion (Enhanced)" -ForegroundColor Cyan
Write-Host "               Author: $script:scriptAuthor" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""

# Function to handle long paths
function Initialize-WorkingDirectory {
    $currentPath = Get-Location
    $pathLength = $currentPath.Path.Length
    
    # Check if path is too long or contains problematic patterns
    $needsSymlink = $false
    if ($pathLength -gt 150) {
        $needsSymlink = $true
        Write-Host "Current path is long ($pathLength chars), creating shorter working directory..." -ForegroundColor Yellow
    }
    if ($currentPath.Path -match 'OneDrive|Google Drive|Dropbox' -and $currentPath.Path -match '\s') {
        $needsSymlink = $true
        Write-Host "Detected cloud storage path with spaces, creating shorter working directory..." -ForegroundColor Yellow
    }
    
    if ($needsSymlink) {
        # Generate a unique symlink name
        $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
        $symlinkPath = "C:\MFA_$timestamp"
        
        try {
            # Try to create symbolic link (requires admin)
            New-Item -ItemType SymbolicLink -Path $symlinkPath -Target $currentPath.Path -ErrorAction Stop | Out-Null
            Write-Host "Created symbolic link: $symlinkPath -> $currentPath" -ForegroundColor Green
        }
        catch {
            # If symbolic link fails, try junction (doesn't require admin)
            try {
                cmd /c mklink /J "$symlinkPath" "$($currentPath.Path)" 2>&1 | Out-Null
                if (Test-Path $symlinkPath) {
                    Write-Host "Created junction: $symlinkPath -> $currentPath" -ForegroundColor Green
                }
                else {
                    throw "Junction creation failed"
                }
            }
            catch {
                Write-Warning "Could not create symbolic link or junction. Continuing with original path..."
                Write-Warning "If you encounter errors, try running PowerShell as Administrator or copying scripts to a shorter path"
                return @{
                    UsedSymlink = $false
                    OriginalPath = $currentPath.Path
                    WorkingPath = $currentPath.Path
                }
            }
        }
        
        # Change to the symlink directory
        Set-Location $symlinkPath
        Write-Host "Working directory changed to: $symlinkPath" -ForegroundColor Cyan
        
        return @{
            UsedSymlink = $true
            OriginalPath = $currentPath.Path
            WorkingPath = $symlinkPath
        }
    }
    
    return @{
        UsedSymlink = $false
        OriginalPath = $currentPath.Path
        WorkingPath = $currentPath.Path
    }
}

# Function to cleanup symlink
function Remove-WorkingDirectory {
    param($pathInfo)
    
    if ($pathInfo.UsedSymlink) {
        Write-Host "`n=== CLEANUP PROCESS ===" -ForegroundColor Yellow
        Write-Host "The script created a temporary working directory to handle long path names." -ForegroundColor Cyan
        Write-Host "This needs to be removed now that the assessment is complete." -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Temporary directory: $($pathInfo.WorkingPath)" -ForegroundColor Gray
        Write-Host "Your files are safely stored in: $($pathInfo.OriginalPath)" -ForegroundColor Green
        Write-Host ""
        
        # Return to original directory
        Set-Location $pathInfo.OriginalPath
        
        # Remove the symlink/junction
        if (Test-Path $pathInfo.WorkingPath) {
            try {
                Write-Host "Removing temporary directory..." -ForegroundColor Yellow
                # Use -Recurse and -Force to avoid the prompt
                Remove-Item $pathInfo.WorkingPath -Recurse -Force -ErrorAction Stop
                Write-Host "✓ Cleanup complete - temporary directory removed" -ForegroundColor Green
                Write-Host ""
            }
            catch {
                Write-Warning "Could not automatically remove temporary directory: $($pathInfo.WorkingPath)"
                Write-Warning "This is just a link - your actual files are safe in the original location"
                Write-Host "You can manually remove it later by running:" -ForegroundColor Yellow
                Write-Host "  Remove-Item '$($pathInfo.WorkingPath)' -Recurse -Force" -ForegroundColor White
                Write-Host ""
            }
        }
    }
}

# Initialize working directory
$pathInfo = Initialize-WorkingDirectory

# Check for required component scripts
$requiredScripts = @(
    "MFA-Connect.ps1",
    "MFA-PolicyCheck.ps1",
    "MFA-UserAnalysis.ps1",
    "MFA-PrivilegedUsers.ps1",
    "MFA-ConditionalAccess.ps1",
    "MFA-Recommendations.ps1",
    "MFA-Reports.ps1"
)

$optionalScripts = @(
    "Generate-UserCommunications.ps1",
    "Generate-MigrationTracker.ps1",
    "Create-AutomationScripts.ps1"
)

$missingScripts = @()
foreach ($script in $requiredScripts) {
    if (-not (Test-Path $script)) {
        $missingScripts += $script
    }
}

if ($missingScripts.Count -gt 0) {
    Write-Host "ERROR: Missing required component scripts:" -ForegroundColor Red
    foreach ($script in $missingScripts) {
        Write-Host "  - $script" -ForegroundColor Yellow
    }
    Write-Host "`nPlease ensure all component scripts are in the same directory." -ForegroundColor Yellow
    
    # Cleanup and exit
    Remove-WorkingDirectory -pathInfo $pathInfo
    return
}

# Check for optional scripts
$availableOptionalScripts = @()
foreach ($script in $optionalScripts) {
    if (Test-Path $script) {
        $availableOptionalScripts += $script
        Write-Host "Found optional script: $script" -ForegroundColor Green
    }
}

# Initialize global data storage
$global:MFAAssessmentData = @{
    TenantInfo = @{}
    CurrentPolicy = @{}
    LegacyMfaData = @{}
    PrivilegedData = @{}
    CaPolicies = @()
    Recommendations = @{}
    ReportPath = $null
    ReportFolder = $null
    TrackerPath = $null
    CommunicationFolder = $null
}

try {
    # Step 1: Connect to services
    Write-Host "Step 1: Connecting to Microsoft Graph..." -ForegroundColor Yellow
    $connected = & .\MFA-Connect.ps1
    if (-not $connected) {
        throw "Failed to connect to Microsoft Graph"
    }
    
    # Step 2: Check Authentication Methods Policy
    if ($PolicyOnly -or $FullAnalysis) {
        Write-Host "`nStep 2: Checking Authentication Methods Policy..." -ForegroundColor Yellow
        $global:MFAAssessmentData.CurrentPolicy = & .\MFA-PolicyCheck.ps1
    }
    
    # Step 3: Full analysis if requested
    if ($FullAnalysis) {
        # Analyze users
        Write-Host "`nStep 3: Analyzing user MFA status..." -ForegroundColor Yellow
        $global:MFAAssessmentData.LegacyMfaData = & .\MFA-UserAnalysis.ps1
        
        # Check Conditional Access
        Write-Host "`nStep 4: Checking Conditional Access policies..." -ForegroundColor Yellow
        $global:MFAAssessmentData.CaPolicies = & .\MFA-ConditionalAccess.ps1
        
        # Check privileged users
        Write-Host "`nStep 5: Analyzing privileged users..." -ForegroundColor Yellow
        $global:MFAAssessmentData.PrivilegedData = & .\MFA-PrivilegedUsers.ps1 -LegacyMfaData $global:MFAAssessmentData.LegacyMfaData
        
        # Generate recommendations
        Write-Host "`nStep 6: Generating recommendations..." -ForegroundColor Yellow
        $global:MFAAssessmentData.Recommendations = & .\MFA-Recommendations.ps1 -AssessmentData $global:MFAAssessmentData
    }
    
    # Step 4: Generate reports if requested
    if ($GenerateReport -and $global:MFAAssessmentData.LegacyMfaData) {
        Write-Host "`nStep 7: Generating enhanced reports..." -ForegroundColor Yellow
        $reportInfo = & .\MFA-Reports.ps1 -AssessmentData $global:MFAAssessmentData
        $global:MFAAssessmentData.ReportPath = $reportInfo.ReportPath
        $global:MFAAssessmentData.ReportFolder = $reportInfo.ReportFolder
        
        # Generate migration tracker if available
        if (Test-Path ".\Generate-MigrationTracker.ps1") {
            Write-Host "`nStep 8: Generating migration tracker..." -ForegroundColor Yellow
            $trackerPath = & .\Generate-MigrationTracker.ps1 -AssessmentData $global:MFAAssessmentData -OutputPath $global:MFAAssessmentData.ReportFolder
            $global:MFAAssessmentData.TrackerPath = $trackerPath
        }
        
        # Generate user communications if there are users needing assistance
        if ($global:MFAAssessmentData.Recommendations.UsersNeedingAssistance -and 
            $global:MFAAssessmentData.Recommendations.UsersNeedingAssistance.Count -gt 0) {
            
            if (Test-Path ".\Generate-UserCommunications.ps1") {
                Write-Host "`nStep 9: Generating user communication templates..." -ForegroundColor Yellow
                try {
                    $communicationFolder = & .\Generate-UserCommunications.ps1 -AssessmentData $global:MFAAssessmentData -OutputPath $global:MFAAssessmentData.ReportFolder
                    $global:MFAAssessmentData.CommunicationFolder = $communicationFolder
                }
                catch {
                    Write-Warning "Failed to generate user communications: $_"
                    Write-Warning "Continuing with other reports..."
                }
            }
        }
        
        # Generate PowerShell automation scripts
        if (Test-Path ".\Create-AutomationScripts.ps1") {
            Write-Host "`nStep 10: Generating automation scripts..." -ForegroundColor Yellow
            try {
                $automationPath = & .\Create-AutomationScripts.ps1 -OutputPath $global:MFAAssessmentData.ReportFolder
                Write-Host "Automation scripts saved to: $automationPath" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to generate automation scripts: $_"
            }
        }
        else {
            Write-Host "`nStep 10: Creating automation scripts folder..." -ForegroundColor Yellow
            $automationPath = Join-Path $global:MFAAssessmentData.ReportFolder "Automation_Scripts"
            New-Item -ItemType Directory -Path $automationPath -Force | Out-Null
            Write-Host "Created automation scripts folder: $automationPath" -ForegroundColor Green
        }
    }
    
    Write-Host "`n=== ASSESSMENT COMPLETE ===" -ForegroundColor Green
    
    # Display final summary
    if ($global:MFAAssessmentData.Recommendations.MethodsToEnable -and $global:MFAAssessmentData.Recommendations.MethodsToEnable.Count -gt 0) {
        Write-Host "`nWARNING: Authentication methods need to be enabled before migration!" -ForegroundColor Yellow
        Write-Host "Review the recommendations above and enable the required methods." -ForegroundColor Yellow
    }
    else {
        Write-Host "`nYour tenant appears ready for MFA migration." -ForegroundColor Green
    }
    
    # Display report location
    if ($global:MFAAssessmentData.ReportFolder) {
        Write-Host "`n=== GENERATED FILES ===" -ForegroundColor Cyan
        
        # If we used a symlink, show both paths
        if ($pathInfo.UsedSymlink) {
            Write-Host "Reports saved to:" -ForegroundColor White
            Write-Host "  Working path: $($global:MFAAssessmentData.ReportFolder)" -ForegroundColor Gray
            
            # Calculate the real path
            $realReportFolder = $global:MFAAssessmentData.ReportFolder -replace [regex]::Escape($pathInfo.WorkingPath), $pathInfo.OriginalPath
            Write-Host "  Actual path: $realReportFolder" -ForegroundColor White
        }
        else {
            Write-Host "Reports saved to: $($global:MFAAssessmentData.ReportFolder)" -ForegroundColor White
        }
        
        if ($global:MFAAssessmentData.TrackerPath) {
            Write-Host "Migration tracker: $(Split-Path $global:MFAAssessmentData.TrackerPath -Leaf)" -ForegroundColor Green
        }
        
        if ($global:MFAAssessmentData.CommunicationFolder) {
            Write-Host "User communications: $(Split-Path $global:MFAAssessmentData.CommunicationFolder -Leaf)" -ForegroundColor Green
        }
        
        Write-Host "`nKEY FILES FOR MIGRATION:" -ForegroundColor Yellow
        Write-Host "1. MFA_Migration_Tracker_*.csv - Primary tracking spreadsheet" -ForegroundColor White
        Write-Host "2. MFA_User_Methods_*.csv - Detailed user authentication data" -ForegroundColor White
        Write-Host "3. MFA_Privileged_Users_*.csv - Admin security assessment" -ForegroundColor White
        Write-Host "4. Automation_Scripts folder - PowerShell scripts for Phase 2" -ForegroundColor White
    }
}
catch {
    Write-Host "`nERROR: Assessment failed" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Gray
}
finally {
    # Disconnect from Microsoft Graph
    try {
        Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Successfully disconnected from Microsoft Graph" -ForegroundColor Green
    }
    catch {
        # Silently continue if disconnect fails
    }
    
    # Clean up global variable
    if (Test-Path variable:global:MFAAssessmentData) {
        Remove-Variable -Name MFAAssessmentData -Scope Global -ErrorAction SilentlyContinue
    }
    
    # Clean up working directory
    Remove-WorkingDirectory -pathInfo $pathInfo

}
