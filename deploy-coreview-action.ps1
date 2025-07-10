# Coreview Custom Action Deployment Script
# This script deploys the Offboard-M365User custom action to coreview

param(
    [Parameter(Mandatory=$true)]
    [string]$CoreviewApiUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$CoreviewApiKey,
    
    [Parameter(Mandatory=$false)]
    [string]$ActionName = "Offboard-M365User"
)

# Function to register custom action with coreview
function Register-CoreviewCustomAction {
    param(
        [string]$ApiUrl,
        [string]$ApiKey,
        [string]$ActionName
    )
    
    $configPath = "coreview-custom-action.json"
    $scriptPath = "Offboard-M365User-Coreview.ps1"
    
    # Verify files exist
    if (-not (Test-Path $configPath)) {
        throw "Configuration file not found: $configPath"
    }
    
    if (-not (Test-Path $scriptPath)) {
        throw "Script file not found: $scriptPath"
    }
    
    # Read configuration
    $config = Get-Content $configPath | ConvertFrom-Json
    
    # Read script content
    $scriptContent = Get-Content $scriptPath -Raw
    
    # Prepare payload for coreview API
    $payload = @{
        name = $config.name
        displayName = $config.displayName
        description = $config.description
        version = $config.version
        type = $config.type
        category = $config.category
        permissions = $config.permissions
        parameters = $config.parameters
        script = @{
            content = $scriptContent
            type = $config.script.type
            timeout = $config.script.timeout
        }
        outputs = $config.outputs
        triggers = $config.triggers
        audit = $config.audit
        security = $config.security
    }
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
    }
    
    $apiEndpoint = "$ApiUrl/api/v1/custom-actions"
    
    try {
        Write-Host "Registering custom action: $($config.displayName)" -ForegroundColor Cyan
        
        $response = Invoke-RestMethod -Uri $apiEndpoint -Method Post -Headers $headers -Body ($payload | ConvertTo-Json -Depth 10)
        
        Write-Host "Custom action registered successfully!" -ForegroundColor Green
        Write-Host "Action ID: $($response.id)" -ForegroundColor Yellow
        
        return $response
    }
    catch {
        Write-Error "Failed to register custom action: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $errorResponse = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $errorBody = $reader.ReadToEnd()
            Write-Error "Error details: $errorBody"
        }
        return $null
    }
}

# Function to verify coreview connectivity
function Test-CoreviewConnection {
    param(
        [string]$ApiUrl,
        [string]$ApiKey
    )
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
    }
    
    $healthEndpoint = "$ApiUrl/api/v1/health"
    
    try {
        $response = Invoke-RestMethod -Uri $healthEndpoint -Method Get -Headers $headers
        Write-Host "Coreview connection successful" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to connect to coreview: $($_.Exception.Message)"
        return $false
    }
}

# Function to get available permissions
function Get-CoreviewPermissions {
    param(
        [string]$ApiUrl,
        [string]$ApiKey
    )
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
    }
    
    $permissionsEndpoint = "$ApiUrl/api/v1/permissions"
    
    try {
        $response = Invoke-RestMethod -Uri $permissionsEndpoint -Method Get -Headers $headers
        return $response
    }
    catch {
        Write-Error "Failed to get permissions: $($_.Exception.Message)"
        return $null
    }
}

# Main execution
Write-Host "Coreview Custom Action Deployment" -ForegroundColor Green
Write-Host "=================================" -ForegroundColor Green

# Step 1: Test connection
Write-Host "Testing coreview connection..." -ForegroundColor Cyan
$connectionTest = Test-CoreviewConnection -ApiUrl $CoreviewApiUrl -ApiKey $CoreviewApiKey

if (-not $connectionTest) {
    Write-Error "Cannot proceed without coreview connection. Exiting."
    exit 1
}

# Step 2: Verify permissions
Write-Host "Verifying required permissions..." -ForegroundColor Cyan
$permissions = Get-CoreviewPermissions -ApiUrl $CoreviewApiUrl -ApiKey $CoreviewApiKey

if ($permissions) {
    $requiredPermissions = @("Calendars.ReadWrite.All", "User.Read.All")
    $missingPermissions = @()
    
    foreach ($permission in $requiredPermissions) {
        if ($permissions -notcontains $permission) {
            $missingPermissions += $permission
        }
    }
    
    if ($missingPermissions.Count -gt 0) {
        Write-Warning "Missing permissions: $($missingPermissions -join ', ')"
        Write-Host "Please ensure these permissions are available in your coreview environment." -ForegroundColor Yellow
    } else {
        Write-Host "All required permissions are available." -ForegroundColor Green
    }
}

# Step 3: Register custom action
Write-Host "Deploying custom action..." -ForegroundColor Cyan
$result = Register-CoreviewCustomAction -ApiUrl $CoreviewApiUrl -ApiKey $CoreviewApiKey -ActionName $ActionName

if ($result) {
    Write-Host "`nDeployment completed successfully!" -ForegroundColor Green
    Write-Host "Custom action '$ActionName' is now available in coreview." -ForegroundColor Yellow
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "1. Configure approval workflows if required" -ForegroundColor White
    Write-Host "2. Test the action with a non-production user" -ForegroundColor White
    Write-Host "3. Monitor audit logs for execution details" -ForegroundColor White
} else {
    Write-Error "Deployment failed. Please check the error messages above."
    exit 1
}