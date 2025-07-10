# Microsoft 365 User Offboarding Script
# This script cancels all calendar invites for a specified user using Microsoft Graph API

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory=$true)]
    [string]$TenantId
)

# Function to get access token using client credentials flow
function Get-AccessToken {
    param(
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$TenantId
    )
    
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    
    $body = @{
        client_id = $ClientId
        client_secret = $ClientSecret
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Error "Failed to get access token: $($_.Exception.Message)"
        return $null
    }
}

# Function to get user by email
function Get-UserByEmail {
    param(
        [string]$UserEmail,
        [string]$AccessToken
    )
    
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$UserEmail"
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $graphUrl -Method Get -Headers $headers
        return $response
    }
    catch {
        Write-Error "Failed to get user: $($_.Exception.Message)"
        return $null
    }
}

# Function to get all calendar events for a user
function Get-UserCalendarEvents {
    param(
        [string]$UserId,
        [string]$AccessToken
    )
    
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$UserId/events"
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    $allEvents = @()
    $nextLink = $graphUrl
    
    try {
        do {
            $response = Invoke-RestMethod -Uri $nextLink -Method Get -Headers $headers
            $allEvents += $response.value
            
            $nextLink = $response.'@odata.nextLink'
        } while ($nextLink)
        
        return $allEvents
    }
    catch {
        Write-Error "Failed to get calendar events: $($_.Exception.Message)"
        return @()
    }
}

# Function to cancel a calendar event
function Cancel-CalendarEvent {
    param(
        [string]$UserId,
        [string]$EventId,
        [string]$AccessToken
    )
    
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$UserId/events/$EventId/cancel"
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    $body = @{
        comment = "Event cancelled as part of user offboarding process"
    } | ConvertTo-Json
    
    try {
        Invoke-RestMethod -Uri $graphUrl -Method Post -Headers $headers -Body $body
        return $true
    }
    catch {
        Write-Error "Failed to cancel event $EventId : $($_.Exception.Message)"
        return $false
    }
}

# Main execution
Write-Host "Starting Microsoft 365 user offboarding process..." -ForegroundColor Green
Write-Host "User: $UserEmail" -ForegroundColor Yellow

# Step 1: Get access token
Write-Host "Getting access token..." -ForegroundColor Cyan
$accessToken = Get-AccessToken -ClientId $ClientId -ClientSecret $ClientSecret -TenantId $TenantId

if (-not $accessToken) {
    Write-Error "Failed to obtain access token. Exiting."
    exit 1
}

Write-Host "Access token obtained successfully." -ForegroundColor Green

# Step 2: Get user details
Write-Host "Getting user details..." -ForegroundColor Cyan
$user = Get-UserByEmail -UserEmail $UserEmail -AccessToken $accessToken

if (-not $user) {
    Write-Error "Failed to get user details. Exiting."
    exit 1
}

Write-Host "User found: $($user.displayName) (ID: $($user.id))" -ForegroundColor Green

# Step 3: Get all calendar events
Write-Host "Retrieving all calendar events..." -ForegroundColor Cyan
$calendarEvents = Get-UserCalendarEvents -UserId $user.id -AccessToken $accessToken

if ($calendarEvents.Count -eq 0) {
    Write-Host "No calendar events found for the user." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($calendarEvents.Count) calendar events." -ForegroundColor Green

# Step 4: Cancel all calendar events
Write-Host "Canceling all calendar events..." -ForegroundColor Cyan
$successCount = 0
$failureCount = 0

foreach ($event in $calendarEvents) {
    Write-Host "Canceling event: $($event.subject) (ID: $($event.id))" -ForegroundColor Yellow
    
    $result = Cancel-CalendarEvent -UserId $user.id -EventId $event.id -AccessToken $accessToken
    
    if ($result) {
        $successCount++
        Write-Host "Successfully canceled event: $($event.subject)" -ForegroundColor Green
    } else {
        $failureCount++
        Write-Host "Failed to cancel event: $($event.subject)" -ForegroundColor Red
    }
}

# Summary
Write-Host "`nCalendar cancellation summary:" -ForegroundColor Cyan
Write-Host "Total events: $($calendarEvents.Count)" -ForegroundColor White
Write-Host "Successfully canceled: $successCount" -ForegroundColor Green
Write-Host "Failed to cancel: $failureCount" -ForegroundColor Red

if ($failureCount -gt 0) {
    Write-Host "`nSome events could not be canceled. Please review the errors above." -ForegroundColor Yellow
    exit 1
} else {
    Write-Host "`nAll calendar events have been successfully canceled!" -ForegroundColor Green
}