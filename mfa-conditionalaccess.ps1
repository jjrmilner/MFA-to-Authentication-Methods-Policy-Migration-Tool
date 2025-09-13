<#
.SYNOPSIS
    Analyzes Conditional Access policies requiring MFA using REST API
#>

Write-Host "`n=== CONDITIONAL ACCESS MFA POLICIES ===" -ForegroundColor Cyan

$mfaPolicies = @()

try {
    # Use REST API to get Conditional Access policies
    $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
    $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
    $policies = $response.value
    
    foreach ($policy in $policies) {
        # Check if policy requires MFA
        if ($policy.grantControls -and $policy.grantControls.builtInControls -contains "mfa") {
            $mfaPolicies += $policy
            
            # Determine color based on state
            $stateColor = switch ($policy.state) {
                "enabled" { "Green" }
                "enabledForReportingButNotEnforced" { "Yellow" }
                "disabled" { "Red" }
                default { "Gray" }
            }
            
            Write-Host "Policy: $($policy.displayName)" -ForegroundColor White
            Write-Host "  State: $($policy.state)" -ForegroundColor $stateColor
            
            # Show which users/groups are targeted
            if ($policy.conditions.users.includeUsers -contains "All") {
                Write-Host "  Target: All Users" -ForegroundColor Gray
            }
            elseif ($policy.conditions.users.includeGroups) {
                Write-Host "  Target: Specific Groups" -ForegroundColor Gray
            }
            elseif ($policy.conditions.users.includeUsers) {
                Write-Host "  Target: Specific Users" -ForegroundColor Gray
            }
        }
    }
    
    # Summary by state
    $enabledCount = ($mfaPolicies | Where-Object { $_.state -eq "enabled" }).Count
    $reportOnlyCount = ($mfaPolicies | Where-Object { $_.state -eq "enabledForReportingButNotEnforced" }).Count
    $disabledCount = ($mfaPolicies | Where-Object { $_.state -eq "disabled" }).Count
    
    Write-Host "`nMFA Policy Summary:" -ForegroundColor Cyan
    Write-Host "  Enabled: $enabledCount" -ForegroundColor Green
    Write-Host "  Report-Only: $reportOnlyCount" -ForegroundColor Yellow
    Write-Host "  Disabled: $disabledCount" -ForegroundColor Red
    Write-Host "  Total MFA Policies: $($mfaPolicies.Count)" -ForegroundColor White
}
catch {
    Write-Warning "Could not analyze Conditional Access policies: $_"
    
    # Check if it's a permissions issue
    if ($_.Exception.Message -like "*Insufficient privileges*" -or $_.Exception.Message -like "*Authorization*") {
        Write-Host "Note: Reading Conditional Access policies requires Policy.Read.All permission" -ForegroundColor Yellow
    }
}

return $mfaPolicies