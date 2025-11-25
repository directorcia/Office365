<#
.SYNOPSIS
    Evaluate Conditional Access policies against ASD Blueprint recommendations

.DESCRIPTION
    This script evaluates the current Conditional Access policies in an Entra ID tenant
    against the recommendations from the ASD (Australian Signals Directorate) Blueprint
    for Secure Cloud and generates an HTML report showing compliance gaps and improvements.
    
    Reference: https://blueprint.asd.gov.au/configuration/entra-id/protection/conditional-access/

.EXAMPLE
    .\asd-ca-get.ps1
    
    Connects to Microsoft Graph, evaluates all CA policies, and generates an HTML report

.NOTES
    Author: CIAOPS
    Date: 2025-11-25
    Version: 1.0
    Requires: Microsoft.Graph.Authentication PowerShell module (auto-installed if missing)
    
.LINK
    Reference - https://blueprint.asd.gov.au/configuration/entra-id/protection/conditional-access/
    Code - https://github.com/directorcia/office365-tools/tree/main/scripts/asd-ca-get.ps1
    Documentation - https://github.com/directorcia/Office365/wiki/ASD-Conditional-Access-Policy-Evaluation-Script
#>

[CmdletBinding()]
param()

# Check for required Microsoft.Graph modules
$requiredModules = @('Microsoft.Graph.Authentication')
$missingModules = @()

Write-Host "`nChecking for required PowerShell modules..." -ForegroundColor Cyan

foreach ($moduleName in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        $missingModules += $moduleName
        Write-Host "  ✗ $moduleName not found" -ForegroundColor Yellow
    }
    else {
        Write-Host "  ✓ $moduleName found" -ForegroundColor Green
    }
}

# Install missing modules
if ($missingModules.Count -gt 0) {
    Write-Host "`nInstalling missing modules..." -ForegroundColor Yellow
    foreach ($moduleName in $missingModules) {
        try {
            Write-Host "  Installing $moduleName..." -ForegroundColor Cyan
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "  ✓ Successfully installed $moduleName" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✗ Failed to install $moduleName. Error: $_" -ForegroundColor Red
            Write-Host "`nPlease install the module manually using:" -ForegroundColor Yellow
            Write-Host "  Install-Module -Name $moduleName -Scope CurrentUser -Force" -ForegroundColor White
            exit 1
        }
    }
}
else {
    Write-Host "All required modules are installed" -ForegroundColor Green
}

# Connect to Microsoft Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    $null = Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Error: $_" -ForegroundColor Red
    exit 1
}

# Get tenant information
Write-Host "`nRetrieving tenant information..." -ForegroundColor Cyan
try {
    $orgResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -OutputType PSObject
    $tenantDetails = $orgResponse.value[0]
    $tenantName = $tenantDetails.displayName
    $tenantId = $tenantDetails.id
    $tenantDomain = $tenantDetails.verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -ExpandProperty name
    if (-not $tenantDomain) {
        $tenantDomain = $tenantDetails.verifiedDomains[0].name
    }
    Write-Host "Tenant: $tenantName" -ForegroundColor Green
}
catch {
    Write-Host "Failed to retrieve tenant information. Error: $_" -ForegroundColor Red
    $tenantName = "Unknown"
    $tenantId = "Unknown"
    $tenantDomain = "Unknown"
}

Write-Host "`nEvaluating Conditional Access policies for tenant: $tenantName" -ForegroundColor Cyan

# Get all Conditional Access policies
try {
    Write-Host "Retrieving Conditional Access policies..." -ForegroundColor Cyan
    $policiesResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -OutputType PSObject
    $policies = $policiesResponse.value
    Write-Host "Found $($policies.Count) Conditional Access policies" -ForegroundColor Green
}
catch {
    Write-Host "Failed to retrieve Conditional Access policies. Error: $_" -ForegroundColor Red
    exit 1
}

# Check for available authentication contexts
Write-Host "Checking for authentication contexts..." -ForegroundColor Cyan
$authContexts = @()
try {
    $authContextResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/authenticationContextClassReferences" -OutputType PSObject -ErrorAction SilentlyContinue
    $authContexts = $authContextResponse.value
    if ($authContexts.Count -gt 0) {
        Write-Host "Found $($authContexts.Count) authentication context(s)" -ForegroundColor Green
    }
    else {
        Write-Host "No authentication contexts configured" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Authentication contexts not available (requires appropriate licensing)" -ForegroundColor Yellow
    $authContexts = @()
}

# Check for Insider Risk Management availability
Write-Host "Checking for Insider Risk Management features..." -ForegroundColor Cyan
$insiderRiskAvailable = $false
try {
    # Try to access a policy with insider risk conditions to verify availability
    # Note: This is a best-effort check as the API might not expose this directly
    $hasInsiderRiskPolicy = $policies | Where-Object { 
        $null -ne $_.conditions.insiderRiskLevels 
    }
    if ($hasInsiderRiskPolicy) {
        $insiderRiskAvailable = $true
        Write-Host "Insider Risk Management integration detected" -ForegroundColor Green
    }
    else {
        Write-Host "No policies with Insider Risk conditions found" -ForegroundColor Yellow
        Write-Host "  Note: Insider Risk Management requires Microsoft Purview and additional licensing" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Insider Risk Management features not available" -ForegroundColor Yellow
    $insiderRiskAvailable = $false
}

# Define ASD Blueprint recommended policies
$asdRecommendations = @(
    @{
        Category = "Admin Protection"
        Name = "ADM - S - Limit admin sessions"
        Description = "Limits administrative session durations"
        Key = "limit-admin-sessions"
        CheckCriteria = @{
            Users = "Directory roles (admin roles)"
            GrantControls = "Grant access"
            SessionControls = "Sign-in frequency configured"
        }
        Priority = "High"
    },
    @{
        Category = "Device Protection"
        Name = "DEV - B - Block access from unapproved devices"
        Description = "Blocks access from devices not approved"
        Key = "block-unapproved-devices"
        CheckCriteria = @{
            Users = "All users"
            GrantControls = "Block access OR Require compliant device"
            DevicePlatforms = "Specified platforms"
        }
        Priority = "High"
    },
    @{
        Category = "Device Protection"
        Name = "DEV - G - Compliant devices"
        Description = "Requires devices to be marked as compliant"
        Key = "compliant-devices"
        CheckCriteria = @{
            Users = "All users"
            Applications = "All resources"
            GrantControls = "Require device to be marked as compliant"
        }
        Priority = "High"
    },
    @{
        Category = "Device Protection"
        Name = "DEV - G - Intune enrolment with strong auth"
        Description = "Requires strong authentication for Intune enrollment"
        Key = "intune-enrolment"
        CheckCriteria = @{
            Applications = "Microsoft Intune Enrollment"
            GrantControls = "Require authentication strength (Phishing-resistant MFA)"
        }
        Priority = "Medium"
    },
    @{
        Category = "Guest Protection"
        Name = "GST - B - Block guests"
        Description = "Blocks guest access where not required"
        Key = "block-guests"
        CheckCriteria = @{
            Users = "Guest or external users"
            GrantControls = "Block access"
        }
        Priority = "Medium"
    },
    @{
        Category = "Guest Protection"
        Name = "GST - G - Guest application access with strong auth"
        Description = "Requires strong authentication for guest access to specific applications"
        Key = "guest-app-access"
        CheckCriteria = @{
            Users = "Guest or external users"
            GrantControls = "Require authentication strength"
        }
        Priority = "Medium"
    },
    @{
        Category = "Location Protection"
        Name = "LOC - B - Block access from unapproved countries"
        Description = "Blocks access from countries not on approved list"
        Key = "block-unapproved-countries"
        CheckCriteria = @{
            Users = "All users"
            Locations = "Configured with excluded locations"
            GrantControls = "Block access"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block access to PROTECTED information"
        Description = "Blocks access to information classified as PROTECTED"
        Key = "block-protected-info"
        CheckCriteria = @{
            AuthenticationContexts = "PROTECTED information"
            GrantControls = "Block access"
        }
        Priority = "Medium"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block access via legacy auth"
        Description = "Blocks legacy authentication protocols"
        Key = "block-legacy-auth"
        CheckCriteria = @{
            Users = "All users"
            ClientApps = "Exchange ActiveSync clients, Other clients"
            GrantControls = "Block access"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block high-risk sign-ins"
        Description = "Blocks sign-ins identified as high risk"
        Key = "block-high-risk-signins"
        CheckCriteria = @{
            Users = "All users"
            SignInRiskLevels = "High"
            GrantControls = "Block access"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block high-risk users"
        Description = "Blocks users identified as high risk"
        Key = "block-high-risk-users"
        CheckCriteria = @{
            Users = "All users"
            UserRiskLevels = "High"
            GrantControls = "Block access"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - B - Block users with elevated insider risk"
        Description = "Blocks users with elevated insider risk scores"
        Key = "block-insider-risk"
        CheckCriteria = @{
            InsiderRisk = "Elevated"
            GrantControls = "Block access"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Agreement to terms of use"
        Description = "Requires agreement to terms of use"
        Key = "terms-of-use"
        CheckCriteria = @{
            Users = "All users"
            GrantControls = "Require terms of use"
        }
        Priority = "Low"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Register security info with strong auth"
        Description = "Requires strong authentication when registering security information"
        Key = "register-security-info"
        CheckCriteria = @{
            Users = "All users"
            UserActions = "Register security information"
            GrantControls = "Require authentication strength"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Require strong auth"
        Description = "Requires phishing-resistant MFA for all users"
        Key = "strong-auth"
        CheckCriteria = @{
            Users = "All users"
            Applications = "All resources"
            GrantControls = "Require authentication strength (Phishing-resistant MFA)"
        }
        Priority = "Critical"
    },
    @{
        Category = "User Protection"
        Name = "USR - G - Risky sign-ins with strong auth"
        Description = "Requires strong authentication for medium/high risk sign-ins"
        Key = "risky-signins-auth"
        CheckCriteria = @{
            Users = "All users"
            SignInRiskLevels = "Medium, High"
            GrantControls = "Require MFA"
        }
        Priority = "High"
    },
    @{
        Category = "User Protection"
        Name = "USR - S - Limit user sessions"
        Description = "Limits user session durations"
        Key = "limit-user-sessions"
        CheckCriteria = @{
            Users = "All users"
            SessionControls = "Sign-in frequency configured"
        }
        Priority = "Medium"
    }
)

# Function to check if a policy matches ASD recommendation
function Test-PolicyCompliance {
    param(
        [Parameter(Mandatory)]
        $Policy,
        [Parameter(Mandatory)]
        $Recommendation
    )
    
    $findings = @()
    $complianceScore = 0
    $maxScore = 0
    
    # Check Users assignment
    if ($Recommendation.CheckCriteria.Users) {
        $maxScore += 20
        $expectedUsers = $Recommendation.CheckCriteria.Users
        
        if ($expectedUsers -match "All users" -and $Policy.conditions.users.includeUsers -contains "All") {
            $complianceScore += 20
            $findings += "✓ Applies to all users"
        }
        elseif ($expectedUsers -match "Directory roles" -and $Policy.conditions.users.includeRoles) {
            $complianceScore += 20
            $findings += "✓ Applies to directory roles"
        }
        elseif ($expectedUsers -match "Guest or external users" -and 
                ($Policy.conditions.users.includeGuestsOrExternalUsers -or 
                 $Policy.conditions.users.guestOrExternalUserTypes)) {
            $complianceScore += 20
            $findings += "✓ Applies to guest/external users"
        }
        else {
            $findings += "✗ User assignment doesn't match recommendation"
        }
    }
    
    # Check Applications
    if ($Recommendation.CheckCriteria.Applications) {
        $maxScore += 20
        $expectedApps = $Recommendation.CheckCriteria.Applications
        
        if ($expectedApps -match "All resources" -and 
            ($Policy.conditions.applications.includeApplications -contains "All" -or
             $Policy.conditions.applications.includeApplications -contains "Office365")) {
            $complianceScore += 20
            $findings += "✓ Applies to all resources"
        }
        elseif ($expectedApps -match "Intune" -and 
                ($Policy.conditions.applications.includeApplications -match "Intune" -or
                 $Policy.conditions.applications.includeApplications -contains "d4ebce55-015a-49b5-a083-c84d1797ae8c")) {
            $complianceScore += 20
            $findings += "✓ Applies to Microsoft Intune"
        }
        else {
            $findings += "✗ Application scope doesn't match recommendation"
        }
    }
    
    # Check Grant Controls
    if ($Recommendation.CheckCriteria.GrantControls) {
        $maxScore += 30
        $expectedGrant = $Recommendation.CheckCriteria.GrantControls
        
        if ($expectedGrant -match "Block access" -and $Policy.grantControls.builtInControls -contains "block") {
            $complianceScore += 30
            $findings += "✓ Blocks access as recommended"
        }
        elseif ($expectedGrant -match "Require device to be marked as compliant" -and 
                $Policy.grantControls.builtInControls -contains "compliantDevice") {
            $complianceScore += 30
            $findings += "✓ Requires compliant device"
        }
        elseif ($expectedGrant -match "Phishing-resistant MFA" -and 
                $Policy.grantControls.authenticationStrength) {
            $complianceScore += 30
            $findings += "✓ Requires authentication strength"
        }
        elseif ($expectedGrant -match "Require MFA" -and 
                ($Policy.grantControls.builtInControls -contains "mfa" -or
                 $Policy.grantControls.authenticationStrength)) {
            $complianceScore += 30
            $findings += "✓ Requires MFA"
        }
        elseif ($expectedGrant -match "Require terms of use" -and 
                $Policy.grantControls.termsOfUse) {
            $complianceScore += 30
            $findings += "✓ Requires terms of use"
        }
        else {
            $findings += "✗ Grant controls don't match recommendation"
        }
    }
    
    # Check Client Apps
    if ($Recommendation.CheckCriteria.ClientApps) {
        $maxScore += 15
        if ($Policy.conditions.clientAppTypes -and 
            ($Policy.conditions.clientAppTypes -contains "exchangeActiveSync" -or
             $Policy.conditions.clientAppTypes -contains "other")) {
            $complianceScore += 15
            $findings += "✓ Targets legacy authentication client apps"
        }
        else {
            $findings += "✗ Client app types not configured for legacy auth"
        }
    }
    
    # Check Sign-in Risk
    if ($Recommendation.CheckCriteria.SignInRiskLevels) {
        $maxScore += 15
        $expectedRisk = $Recommendation.CheckCriteria.SignInRiskLevels
        if ($Policy.conditions.signInRiskLevels) {
            if ($expectedRisk -match "High" -and $Policy.conditions.signInRiskLevels -contains "high") {
                $complianceScore += 15
                $findings += "✓ Configured for high sign-in risk"
            }
            elseif ($expectedRisk -match "Medium" -and 
                    ($Policy.conditions.signInRiskLevels -contains "medium" -or
                     $Policy.conditions.signInRiskLevels -contains "high")) {
                $complianceScore += 15
                $findings += "✓ Configured for medium/high sign-in risk"
            }
            else {
                $findings += "⚠ Sign-in risk levels partially match"
                $complianceScore += 7
            }
        }
        else {
            $findings += "✗ Sign-in risk not configured"
        }
    }
    
    # Check User Risk
    if ($Recommendation.CheckCriteria.UserRiskLevels) {
        $maxScore += 15
        if ($Policy.conditions.userRiskLevels -and 
            $Policy.conditions.userRiskLevels -contains "high") {
            $complianceScore += 15
            $findings += "✓ Configured for high user risk"
        }
        else {
            $findings += "✗ User risk not configured"
        }
    }
    
    # Check Locations
    if ($Recommendation.CheckCriteria.Locations) {
        $maxScore += 10
        if ($Policy.conditions.locations) {
            $complianceScore += 10
            $findings += "✓ Location-based conditions configured"
        }
        else {
            $findings += "✗ Location conditions not configured"
        }
    }
    
    # Check Session Controls
    if ($Recommendation.CheckCriteria.SessionControls) {
        $maxScore += 10
        if ($Policy.sessionControls.signInFrequency) {
            $complianceScore += 10
            $findings += "✓ Sign-in frequency configured"
        }
        else {
            $findings += "✗ Session sign-in frequency not configured"
        }
    }
    
    # Check if policy is enabled
    $policyEnabled = $Policy.state -eq "enabled"
    
    # Calculate final percentage
    if ($maxScore -eq 0) { $maxScore = 100 }
    $compliancePercentage = [math]::Round(($complianceScore / $maxScore) * 100, 0)
    
    return @{
        ComplianceScore = $complianceScore
        MaxScore = $maxScore
        CompliancePercentage = $compliancePercentage
        Findings = $findings
        PolicyEnabled = $policyEnabled
    }
}

# Evaluate policies
$evaluationResults = @()

foreach ($recommendation in $asdRecommendations) {
    Write-Host "`nEvaluating: $($recommendation.Name)" -ForegroundColor Yellow
    
    # Try to find matching policies based primarily on conditions
    $matchingPolicies = @()
    
    foreach ($policy in $policies) {
        # Calculate condition-based match score (0-100)
        $conditionScore = 0
        $maxConditionScore = 0
        
        # Legacy auth check (high weight: 40 points)
        if ($recommendation.Key -eq "block-legacy-auth") {
            $maxConditionScore += 40
            if ($policy.conditions.clientAppTypes -and
                ($policy.conditions.clientAppTypes -contains "exchangeActiveSync" -or
                 $policy.conditions.clientAppTypes -contains "other")) {
                $conditionScore += 40
                if ($policy.grantControls.builtInControls -contains "block") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # High-risk sign-in blocking (high weight: 40 points)
        if ($recommendation.Key -eq "block-high-risk-signins") {
            $maxConditionScore += 40
            if ($policy.conditions.signInRiskLevels -and 
                $policy.conditions.signInRiskLevels -contains "high") {
                $conditionScore += 40
                if ($policy.grantControls.builtInControls -contains "block") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # High-risk user blocking (high weight: 40 points)
        if ($recommendation.Key -eq "block-high-risk-users") {
            $maxConditionScore += 40
            if ($policy.conditions.userRiskLevels -and 
                $policy.conditions.userRiskLevels -contains "high") {
                $conditionScore += 40
                if ($policy.grantControls.builtInControls -contains "block") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Device compliance check (high weight: 40 points)
        if ($recommendation.Key -eq "compliant-devices") {
            $maxConditionScore += 40
            if ($policy.grantControls.builtInControls -contains "compliantDevice") {
                $conditionScore += 40
                if ($policy.conditions.users.includeUsers -contains "All") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Intune enrollment check (medium weight: 30 points)
        if ($recommendation.Key -eq "intune-enrolment") {
            $maxConditionScore += 30
            if ($policy.conditions.applications.includeApplications -and
                ($policy.conditions.applications.includeApplications -match "Intune" -or
                 $policy.conditions.applications.includeApplications -contains "d4ebce55-015a-49b5-a083-c84d1797ae8c")) {
                $conditionScore += 30
                if ($policy.grantControls.authenticationStrength) {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Guest blocking or authentication (medium weight: 30 points)
        if ($recommendation.Key -match "guest") {
            $maxConditionScore += 30
            if ($policy.conditions.users.includeGuestsOrExternalUsers -or
                $policy.conditions.users.guestOrExternalUserTypes) {
                $conditionScore += 30
                if ($recommendation.Key -eq "block-guests" -and 
                    $policy.grantControls.builtInControls -contains "block") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
                elseif ($recommendation.Key -eq "guest-app-access" -and 
                        $policy.grantControls.authenticationStrength) {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Location-based checks (medium weight: 30 points)
        if ($recommendation.Key -match "location" -or $recommendation.Key -match "countr") {
            $maxConditionScore += 30
            if ($policy.conditions.locations) {
                $conditionScore += 30
                if ($policy.grantControls.builtInControls -contains "block") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Admin session limiting (medium weight: 30 points)
        if ($recommendation.Key -eq "limit-admin-sessions") {
            $maxConditionScore += 30
            if ($policy.conditions.users.includeRoles -and $policy.conditions.users.includeRoles.Count -gt 0) {
                $conditionScore += 30
                if ($policy.sessionControls.signInFrequency) {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # User session limiting (medium weight: 30 points)
        if ($recommendation.Key -eq "limit-user-sessions") {
            $maxConditionScore += 30
            if ($policy.sessionControls.signInFrequency) {
                $conditionScore += 30
                if ($policy.conditions.users.includeUsers -contains "All") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Strong auth for all users (high weight: 40 points)
        if ($recommendation.Key -eq "strong-auth") {
            $maxConditionScore += 40
            if ($policy.grantControls.authenticationStrength -or
                $policy.grantControls.builtInControls -contains "mfa") {
                $conditionScore += 40
                if ($policy.conditions.users.includeUsers -contains "All" -and
                    $policy.conditions.applications.includeApplications -contains "All") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Risky sign-ins with strong auth (medium weight: 30 points)
        if ($recommendation.Key -eq "risky-signins-auth") {
            $maxConditionScore += 30
            if ($policy.conditions.signInRiskLevels -and 
                ($policy.conditions.signInRiskLevels -contains "medium" -or
                 $policy.conditions.signInRiskLevels -contains "high")) {
                $conditionScore += 30
                if ($policy.grantControls.builtInControls -contains "mfa" -or
                    $policy.grantControls.authenticationStrength) {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Register security info (medium weight: 30 points)
        if ($recommendation.Key -eq "register-security-info") {
            $maxConditionScore += 30
            if ($policy.conditions.userActions -and 
                $policy.conditions.userActions -contains "registerSecurityInfo") {
                $conditionScore += 30
                if ($policy.grantControls.authenticationStrength) {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Terms of use (low weight: 20 points)
        if ($recommendation.Key -eq "terms-of-use") {
            $maxConditionScore += 20
            if ($policy.grantControls.termsOfUse -and $policy.grantControls.termsOfUse.Count -gt 0) {
                $conditionScore += 20
            }
        }
        
        # Block unapproved devices (medium weight: 30 points)
        if ($recommendation.Key -eq "block-unapproved-devices") {
            $maxConditionScore += 30
            if ($policy.conditions.platforms -and $policy.conditions.platforms.includePlatforms) {
                $conditionScore += 30
                if ($policy.grantControls.builtInControls -contains "block" -or
                    $policy.grantControls.builtInControls -contains "compliantDevice") {
                    $conditionScore += 20
                    $maxConditionScore += 20
                }
            }
        }
        
        # Insider risk (medium weight: 30 points)
        # Note: Requires Microsoft Purview Insider Risk Management
        if ($recommendation.Key -eq "block-insider-risk") {
            if ($insiderRiskAvailable) {
                $maxConditionScore += 30
                if ($policy.conditions.insiderRiskLevels -and
                    $policy.conditions.insiderRiskLevels -contains "elevated") {
                    $conditionScore += 30
                    if ($policy.grantControls.builtInControls -contains "block") {
                        $conditionScore += 20
                        $maxConditionScore += 20
                    }
                }
            }
        }
        
        # Protected information (medium weight: 30 points)
        # Note: Requires authentication contexts to be configured
        if ($recommendation.Key -eq "block-protected-info") {
            if ($authContexts.Count -gt 0) {
                $maxConditionScore += 30
                if ($policy.conditions.authenticationContextClassReferences) {
                    $conditionScore += 30
                    if ($policy.grantControls.builtInControls -contains "block") {
                        $conditionScore += 20
                        $maxConditionScore += 20
                    }
                }
            }
        }
        
        # Calculate match percentage based on conditions
        $matchPercentage = 0
        if ($maxConditionScore -gt 0) {
            $matchPercentage = [math]::Round(($conditionScore / $maxConditionScore) * 100, 0)
        }
        
        # Include policy if it matches conditions with at least 50% confidence
        if ($matchPercentage -ge 50) {
            $matchingPolicies += @{
                Policy = $policy
                MatchScore = $matchPercentage
            }
        }
    }
    
    # Sort matching policies by match score (highest first)
    $matchingPolicies = $matchingPolicies | Sort-Object -Property MatchScore -Descending
    
    if ($matchingPolicies.Count -eq 0) {
        # No matching policy found
        $missingReason = "✗ No policy found matching this recommendation based on condition analysis"
        
        # Add context for feature availability issues
        if ($recommendation.Key -eq "block-insider-risk" -and -not $insiderRiskAvailable) {
            $missingReason = "⚠ Insider Risk Management not detected. This recommendation requires Microsoft Purview Insider Risk Management with appropriate licensing."
        }
        elseif ($recommendation.Key -eq "block-protected-info" -and $authContexts.Count -eq 0) {
            $missingReason = "⚠ No authentication contexts configured. This recommendation requires authentication contexts to be created and configured in your tenant."
        }
        
        $evaluationResults += @{
            Recommendation = $recommendation
            Status = "Missing"
            CompliancePercentage = 0
            MatchingPolicies = @()
            BestMatch = $null
            Findings = @($missingReason)
        }
        Write-Host "  Status: Missing" -ForegroundColor Red
    }
    else {
        # Evaluate each matching policy and consider both condition match score and compliance score
        $bestMatch = $null
        $bestCombinedScore = 0
        
        foreach ($matchResult in $matchingPolicies) {
            $policy = $matchResult.Policy
            $conditionMatchScore = $matchResult.MatchScore
            
            $compliance = Test-PolicyCompliance -Policy $policy -Recommendation $recommendation
            
            # Combined score: 40% condition match + 60% compliance score
            $combinedScore = ($conditionMatchScore * 0.4) + ($compliance.CompliancePercentage * 0.6)
            
            if ($combinedScore -gt $bestCombinedScore) {
                $bestCombinedScore = $combinedScore
                $bestMatch = @{
                    Policy = $policy
                    Compliance = $compliance
                    ConditionMatchScore = $conditionMatchScore
                    CombinedScore = [math]::Round($combinedScore, 0)
                }
            }
        }
        
        # Use the compliance percentage for status determination
        $bestScore = $bestMatch.Compliance.CompliancePercentage
        $status = if ($bestScore -ge 80) { "Compliant" }
                  elseif ($bestScore -ge 50) { "Partial" }
                  else { "Non-Compliant" }
        
        $evaluationResults += @{
            Recommendation = $recommendation
            Status = $status
            CompliancePercentage = $bestScore
            MatchingPolicies = $matchingPolicies
            BestMatch = $bestMatch
            Findings = $bestMatch.Compliance.Findings
        }
        
        Write-Host "  Status: $status ($bestScore%)" -ForegroundColor $(
            if ($status -eq "Compliant") { "Green" }
            elseif ($status -eq "Partial") { "Yellow" }
            else { "Red" }
        )
        Write-Host "  Best matching policy: $($bestMatch.Policy.displayName) [Condition Match: $($bestMatch.ConditionMatchScore)%]" -ForegroundColor Cyan
    }
}

# Generate HTML Report
$reportDate = Get-Date -Format "yyyyMMdd-HHmmss"
$reportPath = Join-Path (Split-Path $PSScriptRoot -Parent) "ASD-CA-Evaluation-Report-$reportDate.html"

# Calculate overall statistics
$totalRecommendations = $asdRecommendations.Count
$compliantCount = ($evaluationResults | Where-Object { $_.Status -eq "Compliant" }).Count
$partialCount = ($evaluationResults | Where-Object { $_.Status -eq "Partial" }).Count
$nonCompliantCount = ($evaluationResults | Where-Object { $_.Status -eq "Non-Compliant" }).Count
$missingCount = ($evaluationResults | Where-Object { $_.Status -eq "Missing" }).Count
$overallCompliance = [math]::Round((($evaluationResults | Measure-Object -Property CompliancePercentage -Average).Average), 0)

# Count by priority
$criticalIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "Critical" -and $_.Status -ne "Compliant" }).Count
$highIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "High" -and $_.Status -ne "Compliant" }).Count
$mediumIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "Medium" -and $_.Status -ne "Compliant" }).Count
$lowIssues = ($evaluationResults | Where-Object { $_.Recommendation.Priority -eq "Low" -and $_.Status -ne "Compliant" }).Count

# Group by category
$categories = $asdRecommendations | Select-Object -ExpandProperty Category -Unique

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ASD Conditional Access Evaluation Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background: #f5f5f5;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .tenant-info {
            background: #f8f9fa;
            padding: 20px 30px;
            border-bottom: 2px solid #e9ecef;
        }
        
        .tenant-info h2 {
            color: #495057;
            margin-bottom: 10px;
        }
        
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        
        .info-item {
            display: flex;
            align-items: center;
        }
        
        .info-label {
            font-weight: 600;
            color: #6c757d;
            margin-right: 10px;
        }
        
        .info-value {
            color: #212529;
        }
        
        .summary {
            padding: 30px;
            background: linear-gradient(to bottom, #ffffff 0%, #f8f9fa 100%);
        }
        
        .summary h2 {
            color: #495057;
            margin-bottom: 20px;
            font-size: 1.8em;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
            border-left: 4px solid;
        }
        
        .stat-card.compliant {
            border-color: #28a745;
        }
        
        .stat-card.partial {
            border-color: #ffc107;
        }
        
        .stat-card.non-compliant {
            border-color: #dc3545;
        }
        
        .stat-card.missing {
            border-color: #6c757d;
        }
        
        .stat-card.overall {
            border-color: #667eea;
        }
        
        .stat-number {
            font-size: 3em;
            font-weight: bold;
            margin: 10px 0;
        }
        
        .stat-label {
            color: #6c757d;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .priority-alerts {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        
        .priority-card {
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }
        
        .priority-card.critical {
            background: #fff5f5;
            border: 2px solid #dc3545;
        }
        
        .priority-card.high {
            background: #fff8e1;
            border: 2px solid #ff9800;
        }
        
        .priority-card.medium {
            background: #fffbf0;
            border: 2px solid #ffc107;
        }
        
        .priority-card.low {
            background: #f0f8ff;
            border: 2px solid #17a2b8;
        }
        
        .priority-number {
            font-size: 2em;
            font-weight: bold;
            margin: 5px 0;
        }
        
        .priority-label {
            font-size: 0.85em;
            font-weight: 600;
        }
        
        .content {
            padding: 30px;
        }
        
        .category-section {
            margin-bottom: 40px;
        }
        
        .category-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            font-size: 1.4em;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            transition: opacity 0.3s;
        }
        
        .category-header:hover {
            opacity: 0.9;
        }
        
        .category-header::after {
            content: '▼';
            font-size: 0.8em;
            transition: transform 0.3s;
        }
        
        .category-header.collapsed::after {
            transform: rotate(-90deg);
        }
        
        .category-content {
            max-height: 10000px;
            overflow: hidden;
            transition: max-height 0.5s ease-out;
        }
        
        .category-content.collapsed {
            max-height: 0;
        }
        
        .recommendation-card {
            background: white;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            margin-bottom: 20px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            transition: box-shadow 0.3s;
        }
        
        .recommendation-card:hover {
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }
        
        .recommendation-header {
            padding: 20px;
            background: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
            cursor: pointer;
            transition: background 0.3s;
        }
        
        .recommendation-header:hover {
            background: #e9ecef;
        }
        
        .recommendation-header::after {
            content: '▼';
            font-size: 0.7em;
            color: #6c757d;
            margin-left: 10px;
            transition: transform 0.3s;
        }
        
        .recommendation-header.collapsed::after {
            transform: rotate(-180deg);
        }
        
        .recommendation-title {
            font-size: 1.2em;
            font-weight: 600;
            color: #212529;
            flex: 1;
        }
        
        .status-badge {
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .status-badge.compliant {
            background: #d4edda;
            color: #155724;
        }
        
        .status-badge.partial {
            background: #fff3cd;
            color: #856404;
        }
        
        .status-badge.non-compliant {
            background: #f8d7da;
            color: #721c24;
        }
        
        .status-badge.missing {
            background: #e2e3e5;
            color: #383d41;
        }
        
        .priority-badge {
            padding: 6px 12px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.8em;
        }
        
        .priority-badge.critical {
            background: #dc3545;
            color: white;
        }
        
        .priority-badge.high {
            background: #ff9800;
            color: white;
        }
        
        .priority-badge.medium {
            background: #ffc107;
            color: #333;
        }
        
        .priority-badge.low {
            background: #17a2b8;
            color: white;
        }
        
        .compliance-bar {
            height: 8px;
            background: #e9ecef;
            border-radius: 4px;
            overflow: hidden;
            margin-top: 10px;
        }
        
        .compliance-fill {
            height: 100%;
            background: linear-gradient(90deg, #28a745 0%, #20c997 100%);
            transition: width 0.5s;
        }
        
        .recommendation-body {
            padding: 20px;
            max-height: 1000px;
            overflow: hidden;
            transition: max-height 0.4s ease-out, padding 0.4s ease-out;
        }
        
        .recommendation-body.collapsed {
            max-height: 0;
            padding-top: 0;
            padding-bottom: 0;
        }
        
        .description {
            color: #6c757d;
            margin-bottom: 15px;
            font-style: italic;
        }
        
        .findings {
            margin-top: 15px;
        }
        
        .findings h4 {
            color: #495057;
            margin-bottom: 10px;
            font-size: 1em;
        }
        
        .findings ul {
            list-style: none;
            padding-left: 0;
        }
        
        .findings li {
            padding: 8px 12px;
            margin-bottom: 5px;
            border-radius: 4px;
            background: #f8f9fa;
        }
        
        .findings li:before {
            margin-right: 8px;
        }
        
        .matching-policy {
            background: #e7f3ff;
            border-left: 4px solid #0066cc;
            padding: 15px;
            margin-top: 15px;
            border-radius: 4px;
        }
        
        .matching-policy h4 {
            color: #0066cc;
            margin-bottom: 8px;
        }
        
        .policy-details {
            display: grid;
            grid-template-columns: auto 1fr;
            gap: 10px;
            margin-top: 10px;
            font-size: 0.9em;
        }
        
        .policy-label {
            font-weight: 600;
            color: #495057;
        }
        
        .policy-value {
            color: #6c757d;
        }
        
        .footer {
            background: #343a40;
            color: white;
            padding: 20px;
            text-align: center;
        }
        
        .footer a {
            color: #80bdff;
            text-decoration: none;
        }
        
        .footer a:hover {
            text-decoration: underline;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
            }
            
            .recommendation-card {
                page-break-inside: avoid;
            }
            
            .category-header::after,
            .recommendation-header::after {
                display: none;
            }
            
            .category-content,
            .recommendation-body {
                max-height: none !important;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ ASD Conditional Access Evaluation Report</h1>
            <p>Australian Signals Directorate Blueprint for Secure Cloud - Compliance Assessment</p>
        </div>
        
        <div class="tenant-info">
            <h2>Tenant Information</h2>
            <div class="info-grid">
                <div class="info-item">
                    <span class="info-label">Tenant Name:</span>
                    <span class="info-value">$tenantName</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Tenant Domain:</span>
                    <span class="info-value">$tenantDomain</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Tenant ID:</span>
                    <span class="info-value">$tenantId</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Report Date:</span>
                    <span class="info-value">$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</span>
                </div>
                <div class="info-item">
                    <span class="info-label">Total Policies:</span>
                    <span class="info-value">$($policies.Count)</span>
                </div>
            </div>
            
            <h3 style="margin-top: 20px; margin-bottom: 10px; color: #495057; font-size: 1.1em;">Feature Availability</h3>
            <div class="info-grid">
                <div class="info-item">
                    <span class="info-label">Authentication Contexts:</span>
                    <span class="info-value" style="color: $(if ($authContexts.Count -gt 0) { '#28a745' } else { '#ffc107' }); font-weight: 600;">
                        $(if ($authContexts.Count -gt 0) { "$($authContexts.Count) configured" } else { "Not configured" })
                    </span>
                </div>
                <div class="info-item">
                    <span class="info-label">Insider Risk Management:</span>
                    <span class="info-value" style="color: $(if ($insiderRiskAvailable) { '#28a745' } else { '#ffc107' }); font-weight: 600;">
                        $(if ($insiderRiskAvailable) { "Available" } else { "Not detected" })
                    </span>
                </div>
            </div>
        </div>
        
        <div class="summary">
            <h2>Executive Summary</h2>
            
            <div class="stats-grid">
                <div class="stat-card overall">
                    <div class="stat-label">Overall Compliance</div>
                    <div class="stat-number">$overallCompliance%</div>
                </div>
                <div class="stat-card compliant">
                    <div class="stat-label">Compliant</div>
                    <div class="stat-number">$compliantCount</div>
                    <div class="stat-label">of $totalRecommendations</div>
                </div>
                <div class="stat-card partial">
                    <div class="stat-label">Partial</div>
                    <div class="stat-number">$partialCount</div>
                    <div class="stat-label">of $totalRecommendations</div>
                </div>
                <div class="stat-card non-compliant">
                    <div class="stat-label">Non-Compliant</div>
                    <div class="stat-number">$nonCompliantCount</div>
                    <div class="stat-label">of $totalRecommendations</div>
                </div>
                <div class="stat-card missing">
                    <div class="stat-label">Missing</div>
                    <div class="stat-number">$missingCount</div>
                    <div class="stat-label">of $totalRecommendations</div>
                </div>
            </div>
            
            <h3 style="margin-top: 30px; margin-bottom: 15px; color: #495057;">Priority Issues</h3>
            <div class="priority-alerts">
                <div class="priority-card critical">
                    <div class="priority-number">$criticalIssues</div>
                    <div class="priority-label">CRITICAL Issues</div>
                </div>
                <div class="priority-card high">
                    <div class="priority-number">$highIssues</div>
                    <div class="priority-label">HIGH Issues</div>
                </div>
                <div class="priority-card medium">
                    <div class="priority-number">$mediumIssues</div>
                    <div class="priority-label">MEDIUM Issues</div>
                </div>
                <div class="priority-card low">
                    <div class="priority-number">$lowIssues</div>
                    <div class="priority-label">LOW Issues</div>
                </div>
            </div>
        </div>
        
        <div class="content">
            <h2 style="margin-bottom: 30px; color: #495057;">Detailed Evaluation Results</h2>
"@

# Add results by category
foreach ($category in $categories) {
    $categoryResults = $evaluationResults | Where-Object { $_.Recommendation.Category -eq $category }
    
    $html += @"
            
            <div class="category-section">
                <div class="category-header">$category</div>
                <div class="category-content">
"@
    
    foreach ($result in $categoryResults) {
        $rec = $result.Recommendation
        $statusClass = $result.Status.ToLower() -replace "-", "-"
        $compliancePercent = $result.CompliancePercentage
        $priorityClass = $rec.Priority.ToLower()
        
        $html += @"
                
                <div class="recommendation-card">
                    <div class="recommendation-header">
                        <div class="recommendation-title">$($rec.Name)</div>
                        <div style="display: flex; gap: 10px; align-items: center;">
                            <span class="priority-badge $priorityClass">$($rec.Priority)</span>
                            <span class="status-badge $statusClass">$($result.Status)</span>
                        </div>
                    </div>
                    <div class="recommendation-body">
                        <div class="description">$($rec.Description)</div>
                        
                        <div style="display: flex; align-items: center; gap: 15px;">
                            <strong>Compliance Score:</strong>
                            <span style="font-size: 1.2em; font-weight: 600; color: $(
                                if ($compliancePercent -ge 80) { '#28a745' }
                                elseif ($compliancePercent -ge 50) { '#ffc107' }
                                else { '#dc3545' }
                            )">$compliancePercent%</span>
                        </div>
                        <div class="compliance-bar">
                            <div class="compliance-fill" style="width: $compliancePercent%; background: $(
                                if ($compliancePercent -ge 80) { 'linear-gradient(90deg, #28a745 0%, #20c997 100%)' }
                                elseif ($compliancePercent -ge 50) { 'linear-gradient(90deg, #ffc107 0%, #ffca2c 100%)' }
                                else { 'linear-gradient(90deg, #dc3545 0%, #e4606d 100%)' }
                            )"></div>
                        </div>
"@
        
        if ($result.Findings) {
            $html += @"
                        
                        <div class="findings">
                            <h4>Findings:</h4>
                            <ul>
"@
            foreach ($finding in $result.Findings) {
                # Add special styling for feature availability warnings
                $liStyle = ""
                if ($finding -match "⚠.*requires") {
                    $liStyle = " style='background: #fff3cd; border-left: 3px solid #ffc107; padding-left: 15px;'"
                }
                $html += "                                <li$liStyle>$finding</li>`n"
            }
            $html += @"
                            </ul>
                        </div>
"@
        }
        
        if ($result.BestMatch) {
            $policy = $result.BestMatch.Policy
            $policyState = $policy.state
            $stateColor = if ($policyState -eq "enabled") { "#28a745" } else { "#dc3545" }
            
            $html += @"
                        
                        <div class="matching-policy">
                            <h4>Best Matching Policy</h4>
                            <div class="policy-details">
                                <span class="policy-label">Name:</span>
                                <span class="policy-value">$($policy.displayName)</span>
                                
                                <span class="policy-label">State:</span>
                                <span class="policy-value" style="color: $stateColor; font-weight: 600;">$($policyState.ToUpper())</span>
                                
                                <span class="policy-label">Policy ID:</span>
                                <span class="policy-value" style="font-family: monospace; font-size: 0.85em;">$($policy.id)</span>
                            </div>
                        </div>
"@
        }
        
        $html += @"
                    </div>
                </div>
"@
    }
    
    $html += @"
                </div>
            </div>
"@
}

$html += @"
        </div>
        
        <div style="background: #f8f9fa; padding: 30px; border-top: 2px solid #e9ecef;">
            <h3 style="color: #495057; margin-bottom: 15px;">📝 Important Notes</h3>
            <div style="background: white; padding: 20px; border-radius: 8px; border-left: 4px solid #667eea;">
                <h4 style="color: #495057; margin-bottom: 10px;">Feature Requirements</h4>
                <ul style="margin-left: 20px; color: #6c757d; line-height: 1.8;">
                    <li><strong>Insider Risk Management:</strong> Requires Microsoft Purview Insider Risk Management with appropriate licensing (Microsoft 365 E5 Compliance or standalone add-on). If not available, the "Block users with elevated insider risk" recommendation cannot be evaluated.</li>
                    <li><strong>Authentication Contexts:</strong> Must be created and configured in your tenant before they can be used in Conditional Access policies. Used for the "Block access to PROTECTED information" recommendation.</li>
                    <li><strong>Identity Protection:</strong> User risk and sign-in risk features require Microsoft Entra ID P2 licensing.</li>
                    <li><strong>Terms of Use:</strong> Must be created and published in your tenant before they can be required in Conditional Access policies.</li>
                </ul>
                
                <h4 style="color: #495057; margin-top: 20px; margin-bottom: 10px;">Evaluation Methodology</h4>
                <ul style="margin-left: 20px; color: #6c757d; line-height: 1.8;">
                    <li><strong>Condition Matching:</strong> Policies are matched to recommendations based on their actual configuration (conditions, grant controls, session controls), not by policy names.</li>
                    <li><strong>Combined Scoring:</strong> Each match is scored using a weighted formula: 40% condition match + 60% compliance score.</li>
                    <li><strong>Thresholds:</strong> Compliant ≥80%, Partial 50-79%, Non-Compliant &lt;50%.</li>
                </ul>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>ASD Blueprint for Secure Cloud</strong></p>
            <p>Reference: <a href="https://blueprint.asd.gov.au/configuration/entra-id/protection/conditional-access/" target="_blank">https://blueprint.asd.gov.au/configuration/entra-id/protection/conditional-access/</a></p>
            <p style="margin-top: 10px; font-size: 0.9em;">Generated by CIAOPS - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        </div>
    </div>
    
    <script>
        // Wait for DOM to be fully loaded
        document.addEventListener('DOMContentLoaded', function() {
            console.log('Initializing collapsible sections...');
            
            // Toggle category sections
            const categoryHeaders = document.querySelectorAll('.category-header');
            console.log('Found ' + categoryHeaders.length + ' category headers');
            categoryHeaders.forEach(function(header) {
                header.addEventListener('click', function() {
                    console.log('Category clicked:', this.textContent);
                    this.classList.toggle('collapsed');
                    const content = this.nextElementSibling;
                    if (content) {
                        content.classList.toggle('collapsed');
                    }
                });
            });
            
            // Toggle recommendation cards
            const recommendationHeaders = document.querySelectorAll('.recommendation-header');
            console.log('Found ' + recommendationHeaders.length + ' recommendation headers');
            recommendationHeaders.forEach(function(header) {
                header.addEventListener('click', function(e) {
                    // Prevent event bubbling to parent elements
                    e.stopPropagation();
                    console.log('Recommendation clicked');
                    this.classList.toggle('collapsed');
                    const body = this.nextElementSibling;
                    if (body) {
                        body.classList.toggle('collapsed');
                    }
                });
            });
            
            // Add expand/collapse all buttons
            const content = document.querySelector('.content');
            if (content) {
                const buttonContainer = document.createElement('div');
                buttonContainer.style.cssText = 'margin-bottom: 20px; display: flex; gap: 10px;';
                buttonContainer.innerHTML = '<button id=\"expandAllBtn\" style=\"padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 600;\">Expand All</button>' +
                    '<button id=\"collapseAllBtn\" style=\"padding: 10px 20px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 600;\">Collapse All</button>';
                content.insertBefore(buttonContainer, content.firstChild);
                
                // Add event listeners to buttons
                document.getElementById('expandAllBtn').addEventListener('click', expandAll);
                document.getElementById('collapseAllBtn').addEventListener('click', collapseAll);
            }
            
            console.log('Initialization complete!');
        });
        
        function expandAll() {
            console.log('Expanding all...');
            document.querySelectorAll('.category-header.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
            document.querySelectorAll('.category-content.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
            document.querySelectorAll('.recommendation-header.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
            document.querySelectorAll('.recommendation-body.collapsed').forEach(function(el) { el.classList.remove('collapsed'); });
        }
        
        function collapseAll() {
            console.log('Collapsing all...');
            document.querySelectorAll('.category-header:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
            document.querySelectorAll('.category-content:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
            document.querySelectorAll('.recommendation-header:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
            document.querySelectorAll('.recommendation-body:not(.collapsed)').forEach(function(el) { el.classList.add('collapsed'); });
        }
    </script>
</body>
</html>
"@

# Save report
$html | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "`n" -NoNewline
Write-Host "============================================" -ForegroundColor Green
Write-Host "     Evaluation Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "`nReport saved to:" -ForegroundColor Cyan
Write-Host $reportPath -ForegroundColor Yellow
Write-Host "`nOverall Compliance: " -NoNewline -ForegroundColor Cyan
Write-Host "$overallCompliance%" -ForegroundColor $(
    if ($overallCompliance -ge 80) { "Green" }
    elseif ($overallCompliance -ge 50) { "Yellow" }
    else { "Red" }
)
Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  ✓ Compliant:     $compliantCount / $totalRecommendations" -ForegroundColor Green
Write-Host "  ⚠ Partial:       $partialCount / $totalRecommendations" -ForegroundColor Yellow
Write-Host "  ✗ Non-Compliant: $nonCompliantCount / $totalRecommendations" -ForegroundColor Red
Write-Host "  ⊘ Missing:       $missingCount / $totalRecommendations" -ForegroundColor Gray
Write-Host ""

# Open the report in default browser
try {
    Start-Process $reportPath
    Write-Host "Report opened in default browser" -ForegroundColor Green
}
catch {
    Write-Host "Please open the report manually: $reportPath" -ForegroundColor Yellow
}

# Disconnect from Microsoft Graph
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Cyan
}
catch {
    # Silently ignore disconnect errors
}
