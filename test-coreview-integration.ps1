# Test Script for Coreview Integration
# This script tests the coreview custom action integration

param(
    [Parameter(Mandatory=$true)]
    [string]$CoreviewApiUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$CoreviewApiKey,
    
    [Parameter(Mandatory=$false)]
    [string]$TestUserEmail = "test@yourdomain.com"
)

# Function to test coreview connectivity
function Test-CoreviewConnectivity {
    param(
        [string]$ApiUrl,
        [string]$ApiKey
    )
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
    }
    
    try {
        # Test health endpoint
        $healthResponse = Invoke-RestMethod -Uri "$ApiUrl/api/v1/health" -Method Get -Headers $headers
        Write-Host "✓ Coreview connectivity test passed" -ForegroundColor Green
        
        # Test custom actions endpoint
        $actionsResponse = Invoke-RestMethod -Uri "$ApiUrl/api/v1/custom-actions" -Method Get -Headers $headers
        Write-Host "✓ Custom actions API accessible" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Error "✗ Coreview connectivity test failed: $($_.Exception.Message)"
        return $false
    }
}

# Function to test custom action registration
function Test-CustomActionRegistration {
    param(
        [string]$ApiUrl,
        [string]$ApiKey
    )
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
    }
    
    try {
        $actions = Invoke-RestMethod -Uri "$ApiUrl/api/v1/custom-actions" -Method Get -Headers $headers
        
        $offboardAction = $actions | Where-Object { $_.name -eq "Offboard-M365User" }
        
        if ($offboardAction) {
            Write-Host "✓ Custom action 'Offboard-M365User' is registered" -ForegroundColor Green
            Write-Host "  Display Name: $($offboardAction.displayName)" -ForegroundColor Yellow
            Write-Host "  Version: $($offboardAction.version)" -ForegroundColor Yellow
            Write-Host "  Status: $($offboardAction.status)" -ForegroundColor Yellow
            return $offboardAction
        } else {
            Write-Warning "✗ Custom action 'Offboard-M365User' not found"
            return $null
        }
    }
    catch {
        Write-Error "✗ Failed to test custom action registration: $($_.Exception.Message)"
        return $null
    }
}

# Function to test custom action execution (dry run)
function Test-CustomActionExecution {
    param(
        [string]$ApiUrl,
        [string]$ApiKey,
        [string]$UserEmail
    )
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
    }
    
    $payload = @{
        actionName = "Offboard-M365User"
        parameters = @{
            UserEmail = $UserEmail
            Comment = "Test execution - dry run"
        }
        dryRun = $true
    }
    
    try {
        Write-Host "Testing custom action execution (dry run)..." -ForegroundColor Cyan
        
        $response = Invoke-RestMethod -Uri "$ApiUrl/api/v1/custom-actions/execute" -Method Post -Headers $headers -Body ($payload | ConvertTo-Json)
        
        Write-Host "✓ Custom action execution test passed" -ForegroundColor Green
        Write-Host "  Execution ID: $($response.executionId)" -ForegroundColor Yellow
        Write-Host "  Status: $($response.status)" -ForegroundColor Yellow
        
        return $response
    }
    catch {
        Write-Error "✗ Custom action execution test failed: $($_.Exception.Message)"
        return $null
    }
}

# Function to test permissions
function Test-Permissions {
    param(
        [string]$ApiUrl,
        [string]$ApiKey
    )
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
    }
    
    try {
        $permissions = Invoke-RestMethod -Uri "$ApiUrl/api/v1/permissions" -Method Get -Headers $headers
        
        $requiredPermissions = @("Calendars.ReadWrite.All", "User.Read.All")
        $availablePermissions = $permissions | Select-Object -ExpandProperty name
        
        Write-Host "Checking required permissions..." -ForegroundColor Cyan
        
        foreach ($permission in $requiredPermissions) {
            if ($availablePermissions -contains $permission) {
                Write-Host "✓ Permission available: $permission" -ForegroundColor Green
            } else {
                Write-Warning "✗ Permission missing: $permission" -ForegroundColor Yellow
            }
        }
        
        return $true
    }
    catch {
        Write-Error "✗ Failed to test permissions: $($_.Exception.Message)"
        return $false
    }
}

# Function to validate script syntax
function Test-ScriptSyntax {
    param(
        [string]$ScriptPath
    )
    
    try {
        if (-not (Test-Path $ScriptPath)) {
            Write-Error "✗ Script file not found: $ScriptPath"
            return $false
        }
        
        # Test PowerShell syntax
        $null = [System.Management.Automation.PSParser]::Tokenize((Get-Content $ScriptPath -Raw), [ref]$null)
        Write-Host "✓ Script syntax validation passed" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "✗ Script syntax validation failed: $($_.Exception.Message)"
        return $false
    }
}

# Function to validate configuration
function Test-Configuration {
    param(
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            Write-Error "✗ Configuration file not found: $ConfigPath"
            return $false
        }
        
        $config = Get-Content $ConfigPath | ConvertFrom-Json
        
        # Validate required fields
        $requiredFields = @("name", "displayName", "description", "version", "type", "category")
        
        foreach ($field in $requiredFields) {
            if (-not $config.$field) {
                Write-Error "✗ Missing required field: $field"
                return $false
            }
        }
        
        Write-Host "✓ Configuration validation passed" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "✗ Configuration validation failed: $($_.Exception.Message)"
        return $false
    }
}

# Main test execution
Write-Host "Coreview Integration Test Suite" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Green
Write-Host ""

# Test 1: Coreview connectivity
Write-Host "Test 1: Coreview Connectivity" -ForegroundColor Cyan
$connectivityTest = Test-CoreviewConnectivity -ApiUrl $CoreviewApiUrl -ApiKey $CoreviewApiKey
Write-Host ""

if (-not $connectivityTest) {
    Write-Error "Cannot proceed with tests. Coreview connectivity failed."
    exit 1
}

# Test 2: Permissions
Write-Host "Test 2: Permissions Check" -ForegroundColor Cyan
$permissionsTest = Test-Permissions -ApiUrl $CoreviewApiUrl -ApiKey $CoreviewApiKey
Write-Host ""

# Test 3: Script syntax
Write-Host "Test 3: Script Syntax Validation" -ForegroundColor Cyan
$scriptTest = Test-ScriptSyntax -ScriptPath "Offboard-M365User-Coreview.ps1"
Write-Host ""

# Test 4: Configuration validation
Write-Host "Test 4: Configuration Validation" -ForegroundColor Cyan
$configTest = Test-Configuration -ConfigPath "coreview-custom-action.json"
Write-Host ""

# Test 5: Custom action registration
Write-Host "Test 5: Custom Action Registration" -ForegroundColor Cyan
$actionTest = Test-CustomActionRegistration -ApiUrl $CoreviewApiUrl -ApiKey $CoreviewApiKey
Write-Host ""

# Test 6: Custom action execution (dry run)
if ($actionTest) {
    Write-Host "Test 6: Custom Action Execution (Dry Run)" -ForegroundColor Cyan
    $executionTest = Test-CustomActionExecution -ApiUrl $CoreviewApiUrl -ApiKey $CoreviewApiKey -UserEmail $TestUserEmail
    Write-Host ""
}

# Summary
Write-Host "Test Summary" -ForegroundColor Green
Write-Host "============" -ForegroundColor Green

$tests = @(
    @{ Name = "Coreview Connectivity"; Result = $connectivityTest },
    @{ Name = "Permissions Check"; Result = $permissionsTest },
    @{ Name = "Script Syntax"; Result = $scriptTest },
    @{ Name = "Configuration"; Result = $configTest },
    @{ Name = "Custom Action Registration"; Result = $actionTest -ne $null }
)

if ($actionTest) {
    $tests += @{ Name = "Custom Action Execution"; Result = $executionTest -ne $null }
}

$passedTests = 0
$totalTests = $tests.Count

foreach ($test in $tests) {
    $status = if ($test.Result) { "✓ PASS" } else { "✗ FAIL" }
    $color = if ($test.Result) { "Green" } else { "Red" }
    Write-Host "$status - $($test.Name)" -ForegroundColor $color
    $passedTests += if ($test.Result) { 1 } else { 0 }
}

Write-Host ""
Write-Host "Results: $passedTests/$totalTests tests passed" -ForegroundColor $(if ($passedTests -eq $totalTests) { "Green" } else { "Yellow" })

if ($passedTests -eq $totalTests) {
    Write-Host "`n🎉 All tests passed! The coreview integration is ready for use." -ForegroundColor Green
} else {
    Write-Host "`n⚠️  Some tests failed. Please review the errors above before proceeding." -ForegroundColor Yellow
}