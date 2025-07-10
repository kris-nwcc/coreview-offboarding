# Microsoft 365 User Offboarding Script - Coreview Integration
# This script cancels all calendar invites for a specified user using Microsoft Graph API
# Optimized to run as a coreview custom action with coreview authentication

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$Comment = "Event cancelled as part of user offboarding process"
)

# Coreview authentication function
function Get-CoreviewAccessToken {
    param(
        [string]$TenantId
    )
    
    try {
        # Use coreview's built-in authentication method
        # This assumes coreview provides a method to get authenticated access token
        $accessToken = Get-CoreviewToken -TenantId $TenantId -Scope "https://graph.microsoft.com/.default"
        
        if (-not $accessToken) {
            throw "Failed to obtain access token from coreview"
        }
        
        return $accessToken
    }
    catch {
        Write-Error "Failed to get access token from coreview: $($_.Exception.Message)"
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

# Function to get all calendar events for a user with pagination
function Get-UserCalendarEvents {
    param(
        [string]$UserId,
        [string]$AccessToken,
        [datetime]$StartDate = (Get-Date).AddDays(-30),
        [datetime]$EndDate = (Get-Date).AddDays(365)
    )
    
    $filter = "start/dateTime ge '$($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))' and end/dateTime le '$($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))'"
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$UserId/events?`$filter=$filter&`$orderby=start/dateTime"
    
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
        [string]$AccessToken,
        [string]$Comment
    )
    
    $graphUrl = "https://graph.microsoft.com/v1.0/users/$UserId/events/$EventId/cancel"
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }
    
    $body = @{
        comment = $Comment
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

# Function to get coreview context and tenant information
function Get-CoreviewContext {
    try {
        # Get current coreview context
        $context = Get-CoreviewContext
        return $context
    }
    catch {
        Write-Error "Failed to get coreview context: $($_.Exception.Message)"
        return $null
    }
}

# Main execution
Write-Host "Starting Microsoft 365 user offboarding process with coreview integration..." -ForegroundColor Green
Write-Host "User: $UserEmail" -ForegroundColor Yellow

# Step 1: Get coreview context and determine tenant
Write-Host "Getting coreview context..." -ForegroundColor Cyan
$coreviewContext = Get-CoreviewContext

if (-not $coreviewContext) {
    Write-Error "Failed to get coreview context. Exiting."
    exit 1
}

# Use provided tenant ID or get from coreview context
if (-not $TenantId) {
    $TenantId = $coreviewContext.TenantId
    Write-Host "Using tenant ID from coreview context: $TenantId" -ForegroundColor Cyan
}

Write-Host "Using tenant ID: $TenantId" -ForegroundColor Green

# Step 2: Get access token using coreview authentication
Write-Host "Getting access token via coreview..." -ForegroundColor Cyan
$accessToken = Get-CoreviewAccessToken -TenantId $TenantId

if (-not $accessToken) {
    Write-Error "Failed to obtain access token via coreview. Exiting."
    exit 1
}

Write-Host "Access token obtained successfully via coreview." -ForegroundColor Green

# Step 3: Get user details
Write-Host "Getting user details..." -ForegroundColor Cyan
$user = Get-UserByEmail -UserEmail $UserEmail -AccessToken $accessToken

if (-not $user) {
    Write-Error "Failed to get user details. Exiting."
    exit 1
}

Write-Host "User found: $($user.displayName) (ID: $($user.id))" -ForegroundColor Green

# Step 4: Get calendar events (last 30 days to next year)
Write-Host "Retrieving calendar events..." -ForegroundColor Cyan
$startDate = (Get-Date).AddDays(-30)
$endDate = (Get-Date).AddDays(365)
$calendarEvents = Get-UserCalendarEvents -UserId $user.id -AccessToken $accessToken -StartDate $startDate -EndDate $endDate

if ($calendarEvents.Count -eq 0) {
    Write-Host "No calendar events found for the user in the specified date range." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($calendarEvents.Count) calendar events." -ForegroundColor Green

# Step 5: Cancel all calendar events
Write-Host "Canceling all calendar events..." -ForegroundColor Cyan
$successCount = 0
$failureCount = 0
$canceledEvents = @()

foreach ($event in $calendarEvents) {
    Write-Host "Canceling event: $($event.subject) (ID: $($event.id))" -ForegroundColor Yellow
    
    $result = Cancel-CalendarEvent -UserId $user.id -EventId $event.id -AccessToken $accessToken -Comment $Comment
    
    if ($result) {
        $successCount++
        $canceledEvents += @{
            Id = $event.id
            Subject = $event.subject
            StartTime = $event.start.dateTime
            EndTime = $event.end.dateTime
        }
        Write-Host "Successfully canceled event: $($event.subject)" -ForegroundColor Green
    } else {
        $failureCount++
        Write-Host "Failed to cancel event: $($event.subject)" -ForegroundColor Red
    }
}

# Step 6: Log results to coreview
try {
    $logData = @{
        UserEmail = $UserEmail
        UserId = $user.id
        UserDisplayName = $user.displayName
        TotalEvents = $calendarEvents.Count
        SuccessfullyCanceled = $successCount
        FailedToCancel = $failureCount
        CanceledEvents = $canceledEvents
        Timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        Comment = $Comment
    }
    
    # Log to coreview audit trail
    Write-CoreviewLog -Action "UserOffboarding" -Data $logData
    Write-Host "Results logged to coreview audit trail." -ForegroundColor Green
}
catch {
    Write-Warning "Failed to log results to coreview: $($_.Exception.Message)"
}

# Summary
Write-Host "`nCalendar cancellation summary:" -ForegroundColor Cyan
Write-Host "Total events: $($calendarEvents.Count)" -ForegroundColor White
Write-Host "Successfully canceled: $successCount" -ForegroundColor Green
Write-Host "Failed to cancel: $failureCount" -ForegroundColor Red
Write-Host "Date range: $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))" -ForegroundColor White

if ($failureCount -gt 0) {
    Write-Host "`nSome events could not be canceled. Please review the errors above." -ForegroundColor Yellow
    exit 1
} else {
    Write-Host "`nAll calendar events have been successfully canceled!" -ForegroundColor Green
    exit 0
}