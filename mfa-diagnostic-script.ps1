<# SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause
# Copyright (c) 2025 Global Micro Solutions (Pty) Ltd
# All rights reserved

.WARRANTY
    Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
    either express or implied. See the Apache-2.0 WITH Commons-Clause License for the specific language
    governing permissions and limitations under the License.

# MFA Reports File Diagnostic Script
# This script checks the mfa-reports.ps1 file for encoding issues
#>

param(
    [string]$FilePath = ".\mfa-reports.ps1"
)

Write-Host "MFA Reports File Diagnostic Tool" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan

# Check if file exists
if (-not (Test-Path $FilePath)) {
    Write-Host "ERROR: File not found: $FilePath" -ForegroundColor Red
    return
}

Write-Host "`nChecking file: $FilePath" -ForegroundColor Yellow

# Get file info
$fileInfo = Get-Item $FilePath
Write-Host "File size: $($fileInfo.Length) bytes"
Write-Host "Last modified: $($fileInfo.LastWriteTime)"

# Read file content
try {
    $content = Get-Content $FilePath -Raw -Encoding UTF8
    Write-Host "Successfully read file with UTF8 encoding" -ForegroundColor Green
}
catch {
    Write-Host "ERROR reading file: $_" -ForegroundColor Red
    return
}

# Check for BOM
$bytes = [System.IO.File]::ReadAllBytes($FilePath)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    Write-Host "WARNING: File has UTF-8 BOM (Byte Order Mark)" -ForegroundColor Yellow
}
else {
    Write-Host "OK: No BOM detected" -ForegroundColor Green
}

# Find problematic Unicode characters
Write-Host "`nScanning for Unicode characters..." -ForegroundColor Yellow

$unicodeChars = @{
    '🚨' = 'Emergency Light (U+1F6A8)'
    '✓' = 'Check Mark (U+2713)'
    '✔' = 'Heavy Check Mark (U+2714)'
    '⚠️' = 'Warning Sign (U+26A0)'
    '├─' = 'Box Drawing Light Vertical and Right (U+251C U+2500)'
    '│' = 'Box Drawing Light Vertical (U+2502)'
    '└─' = 'Box Drawing Light Up and Right (U+2514 U+2500)'
    '✅' = 'White Heavy Check Mark (U+2705)'
    '📊' = 'Bar Chart (U+1F4CA)'
    '🔐' = 'Closed Lock with Key (U+1F510)'
    '📋' = 'Clipboard (U+1F4CB)'
    '→' = 'Rightwards Arrow (U+2192)'
    '⏱️' = 'Stopwatch (U+23F1)'
    'ðŸš¨' = 'Corrupted Emergency Light'
    'âœ"' = 'Corrupted Check Mark'
    'âš ï¸' = 'Corrupted Warning Sign'
    'â"œâ"€' = 'Corrupted Box Drawing'
    'â""â"€' = 'Corrupted Box Drawing'
}

$foundChars = @{}
$lineNumber = 0
$problemLines = @()

foreach ($line in $content.Split("`n")) {
    $lineNumber++
    
    foreach ($char in $unicodeChars.Keys) {
        if ($line.Contains($char)) {
            if (-not $foundChars.ContainsKey($char)) {
                $foundChars[$char] = @()
            }
            $foundChars[$char] += $lineNumber
            
            # Mark lines 240-250 and 340-350 as especially problematic
            if (($lineNumber -ge 240 -and $lineNumber -le 250) -or ($lineNumber -ge 340 -and $lineNumber -le 350)) {
                $problemLines += @{
                    LineNumber = $lineNumber
                    Content = $line.Substring(0, [Math]::Min($line.Length, 80)) + "..."
                    Character = $char
                }
            }
        }
    }
}

# Report findings
if ($foundChars.Count -eq 0) {
    Write-Host "`nNo Unicode characters found - file appears clean!" -ForegroundColor Green
}
else {
    Write-Host "`nFound Unicode characters:" -ForegroundColor Red
    foreach ($char in $foundChars.Keys) {
        $description = $unicodeChars[$char]
        $lines = $foundChars[$char] -join ", "
        Write-Host "  '$char' ($description) on lines: $lines" -ForegroundColor Yellow
    }
}

# Show problem lines
if ($problemLines.Count -gt 0) {
    Write-Host "`nProblematic lines (240-250 and 340-350):" -ForegroundColor Red
    foreach ($problem in $problemLines) {
        Write-Host "  Line $($problem.LineNumber): $($problem.Content)" -ForegroundColor Yellow
        Write-Host "    Contains: '$($problem.Character)'" -ForegroundColor Red
    }
}

# Check specific problem areas
Write-Host "`nChecking specific problem areas..." -ForegroundColor Yellow

# Check around line 243-245
$lines = $content.Split("`n")
if ($lines.Length -gt 245) {
    Write-Host "`nContent around lines 243-245:" -ForegroundColor Yellow
    for ($i = 242; $i -le 245 -and $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        if ($line.Length -gt 100) {
            $line = $line.Substring(0, 100) + "..."
        }
        Write-Host "  Line $($i+1): $line"
    }
}

# Generate fix suggestions
Write-Host "`n=== FIX SUGGESTIONS ===" -ForegroundColor Cyan
Write-Host "1. Open mfa-reports.ps1 in a text editor (Notepad++, VS Code, etc.)"
Write-Host "2. Use Find & Replace (Ctrl+H) to replace these characters:"
Write-Host ""
Write-Host "   Find: 🚨 or ðŸš¨    Replace: [CRITICAL]"
Write-Host "   Find: ✓ or âœ"      Replace: [OK]"
Write-Host "   Find: ⚠️ or âš ï¸  Replace: [WARNING]"
Write-Host "   Find: ├─ or â"œâ"€   Replace: +--"
Write-Host "   Find: │            Replace: |"
Write-Host "   Find: └─ or â""â"€   Replace: +--"
Write-Host "   Find: ✅           Replace: [X]"
Write-Host "   Find: 📊           Replace: [DATA]"
Write-Host "   Find: 🔐           Replace: [SECURITY]"
Write-Host "   Find: 📋           Replace: [REPORT]"
Write-Host "   Find: →            Replace: -->"
Write-Host "   Find: ⏱️           Replace: Time:"
Write-Host ""
Write-Host "3. Save the file as UTF-8 WITHOUT BOM"
Write-Host "4. Run this diagnostic again to verify the fix"

# Create a cleaned version
Write-Host "`nWould you like to create a cleaned version? (Y/N): " -NoNewline
$response = Read-Host

if ($response -eq 'Y' -or $response -eq 'y') {
    $cleanedContent = $content
    
    # Replace all problematic characters
    $replacements = @{
        '🚨' = '[CRITICAL]'
        'ðŸš¨' = '[CRITICAL]'
        '✓' = '[OK]'
        'âœ"' = '[OK]'
        '⚠️' = '[WARNING]'
        'âš ï¸' = '[WARNING]'
        '├─' = '+--'
        'â"œâ"€' = '+--'
        '│' = '|'
        '└─' = '+--'
        'â""â"€' = '+--'
        '✅' = '[X]'
        '📊' = '[DATA]'
        '🔐' = '[SECURITY]'
        '📋' = '[REPORT]'
        '→' = '-->'
        '⏱️' = 'Time:'
    }
    
    foreach ($find in $replacements.Keys) {
        $cleanedContent = $cleanedContent.Replace($find, $replacements[$find])
    }
    
    $cleanedPath = $FilePath.Replace('.ps1', '_cleaned.ps1')
    
    # Write without BOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($cleanedPath, $cleanedContent, $utf8NoBom)
    
    Write-Host "`nCleaned file created: $cleanedPath" -ForegroundColor Green
    Write-Host "Rename this file to mfa-reports.ps1 to use it" -ForegroundColor Yellow
}

Write-Host "`nDiagnostic complete!" -ForegroundColor Green
