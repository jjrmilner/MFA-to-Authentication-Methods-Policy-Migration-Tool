<#
.SYNOPSIS
    Checks current Authentication Methods Policy using REST API
#>

Write-Host "`n=== CURRENT AUTHENTICATION METHODS POLICY ===" -ForegroundColor Cyan

try {
    # Use REST API to get Authentication Methods Policy
    $uri = "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
    $policyResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
    $policy = $policyResponse
    $policyStatus = @{}
    
    Write-Host "`nAuthentication Methods Policy Status:" -ForegroundColor Yellow
    
    # Process authentication method configurations
    foreach ($config in $policy.authenticationMethodConfigurations) {
        $methodId = $config.id
        $state = $config.state
        
        $status = if ($state -eq "enabled") { "ENABLED" } else { "DISABLED" }
        $color = if ($state -eq "enabled") { "Green" } else { "Red" }
        
        Write-Host "  $($methodId): $status" -ForegroundColor $color
        $policyStatus[$methodId] = $state
    }
    
    # Also check for specific method types if they have different property names
    # Some methods might be under different property names in the policy
    if ($policy.registrationEnforcement) {
        Write-Host "`nRegistration Enforcement:" -ForegroundColor Yellow
        Write-Host "  Campaign State: $($policy.registrationEnforcement.authenticationMethodsRegistrationCampaign.state)" -ForegroundColor Gray
    }
    
    return $policyStatus
}
catch {
    Write-Host "ERROR: Cannot retrieve Authentication Methods Policy" -ForegroundColor Red
    Write-Host "Details: $_" -ForegroundColor Yellow
    
    # Try alternative endpoint
    try {
        Write-Host "`nTrying alternative endpoint..." -ForegroundColor Yellow
        $uri = "https://graph.microsoft.com/beta/policies/authenticationMethodsPolicy"
        $policyResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        $policy = $policyResponse
        $policyStatus = @{}
        
        Write-Host "Authentication Methods Policy Status (Beta API):" -ForegroundColor Yellow
        foreach ($config in $policy.authenticationMethodConfigurations) {
            $methodId = $config.id
            $state = $config.state
            
            $status = if ($state -eq "enabled") { "ENABLED" } else { "DISABLED" }
            $color = if ($state -eq "enabled") { "Green" } else { "Red" }
            
            Write-Host "  $($methodId): $status" -ForegroundColor $color
            $policyStatus[$methodId] = $state
        }
        
        return $policyStatus
    }
    catch {
        Write-Host "ERROR: Cannot retrieve policy from beta endpoint either" -ForegroundColor Red
        Write-Host "Details: $_" -ForegroundColor Yellow
        return @{}
    }
}