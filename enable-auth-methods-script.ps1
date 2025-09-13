<#
.SYNOPSIS
    Enables required authentication methods for MFA migration
    
.DESCRIPTION
    This script enables Voice, Email, WindowsHello, and SoftwareOath authentication methods
    in the tenant's Authentication Methods Policy to prepare for migration
    
.EXAMPLE
    .\Enable-AuthMethods.ps1
#>

# Connect to Microsoft Graph if not already connected
$context = Get-MgContext
if (-not $context) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "Policy.ReadWrite.AuthenticationMethod"
}

Write-Host "`n=== ENABLING REQUIRED AUTHENTICATION METHODS ===" -ForegroundColor Cyan
Write-Host "This script will enable the following methods:" -ForegroundColor Yellow
Write-Host "  - Voice (84 users)" -ForegroundColor White
Write-Host "  - Email (62 users)" -ForegroundColor White
Write-Host "  - Windows Hello for Business (63 users)" -ForegroundColor White
Write-Host "  - Software OATH tokens (4 users)" -ForegroundColor White

# Prompt for confirmation
$confirm = Read-Host "`nDo you want to proceed? (Y/N)"
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    return
}

# Function to enable an authentication method
function Enable-AuthMethod {
    param(
        [string]$MethodId,
        [string]$DisplayName
    )
    
    try {
        Write-Host "`nEnabling $DisplayName..." -ForegroundColor Yellow
        
        # Get current configuration
        $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/$MethodId"
        $currentConfig = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        if ($currentConfig.state -eq "enabled") {
            Write-Host "  $DisplayName is already enabled" -ForegroundColor Green
            return $true
        }
        
        # Enable the method
        $body = @{
            "@odata.type" = "#microsoft.graph.$($MethodId)AuthenticationMethodConfiguration"
            "state" = "enabled"
        }
        
        $response = Invoke-MgGraphRequest -Uri $uri -Method PATCH -Body $body
        Write-Host "  $DisplayName has been enabled successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  ERROR: Failed to enable $DisplayName" -ForegroundColor Red
        Write-Host "  Details: $_" -ForegroundColor Red
        return $false
    }
}

# Track results
$results = @{
    Success = @()
    Failed = @()
}

# Enable each required method
$methods = @(
    @{Id = "voice"; Name = "Voice authentication"},
    @{Id = "email"; Name = "Email OTP authentication"},
    @{Id = "windowsHelloForBusiness"; Name = "Windows Hello for Business"},
    @{Id = "softwareOath"; Name = "Software OATH tokens"}
)

foreach ($method in $methods) {
    $success = Enable-AuthMethod -MethodId $method.Id -DisplayName $method.Name
    
    if ($success) {
        $results.Success += $method.Name
    }
    else {
        $results.Failed += $method.Name
    }
    
    # Small delay to avoid throttling
    Start-Sleep -Seconds 1
}

# Display summary
Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
if ($results.Success.Count -gt 0) {
    Write-Host "`nSuccessfully enabled:" -ForegroundColor Green
    foreach ($method in $results.Success) {
        Write-Host "  ✓ $method" -ForegroundColor Green
    }
}

if ($results.Failed.Count -gt 0) {
    Write-Host "`nFailed to enable:" -ForegroundColor Red
    foreach ($method in $results.Failed) {
        Write-Host "  ✗ $method" -ForegroundColor Red
    }
    Write-Host "`nPlease enable these methods manually in the Azure portal." -ForegroundColor Yellow
}

# Verify current status
Write-Host "`n=== VERIFYING CURRENT STATUS ===" -ForegroundColor Cyan
try {
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
    $policy = Invoke-MgGraphRequest -Uri $uri -Method GET
    
    Write-Host "Current Authentication Methods Policy Status:" -ForegroundColor Yellow
    foreach ($config in $policy.authenticationMethodConfigurations) {
        $status = if ($config.state -eq "enabled") { "ENABLED" } else { "DISABLED" }
        $color = if ($config.state -eq "enabled") { "Green" } else { "Red" }
        Write-Host "  $($config.id): $status" -ForegroundColor $color
    }
}
catch {
    Write-Host "Could not verify policy status" -ForegroundColor Yellow
}

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Cyan
Write-Host "1. Verify all methods show as ENABLED above" -ForegroundColor White
Write-Host "2. Wait 15-30 minutes for changes to propagate" -ForegroundColor White
Write-Host "3. Run the MFA assessment again to confirm readiness" -ForegroundColor White
Write-Host "4. Schedule Phase 1 migration before September 30, 2025" -ForegroundColor White

Write-Host "`nScript completed!" -ForegroundColor Green