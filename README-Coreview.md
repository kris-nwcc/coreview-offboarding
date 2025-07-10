# Microsoft 365 User Offboarding Script - Coreview Integration

This repository contains an optimized version of the Microsoft 365 user offboarding script that integrates with coreview for authentication and runs as a custom action. The script cancels all calendar invites for a specified user using Microsoft Graph API with enhanced security and audit capabilities.

## Key Improvements

### 🔐 Coreview Authentication
- **Eliminates manual credential management**: No need to manage client secrets or tenant IDs manually
- **Automatic tenant context**: Uses coreview's built-in tenant context detection
- **Enhanced security**: Leverages coreview's secure authentication mechanisms
- **Audit trail integration**: All actions are logged to coreview's audit system

### 🚀 Custom Action Integration
- **Workflow integration**: Can be called from coreview workflows and approval processes
- **Parameter validation**: Built-in validation for email addresses and other parameters
- **Role-based access control**: Configurable execution and approval roles
- **Detailed outputs**: Structured output for integration with other systems

### 📊 Enhanced Features
- **Date range filtering**: Focuses on relevant events (last 30 days to next year)
- **Better error handling**: Comprehensive error reporting and recovery
- **Progress tracking**: Real-time progress updates during execution
- **Detailed logging**: Complete audit trail of all operations

## Files Overview

| File | Description |
|------|-------------|
| `Offboard-M365User-Coreview.ps1` | Main script optimized for coreview integration |
| `coreview-custom-action.json` | Custom action configuration for coreview |
| `deploy-coreview-action.ps1` | Deployment script to register the custom action |
| `Offboard-M365User.ps1` | Original script (for reference) |

## Prerequisites

### Coreview Requirements
1. **Coreview Environment**: Access to a coreview instance with API capabilities
2. **API Access**: Valid API key with permissions to register custom actions
3. **Microsoft Graph Permissions**: The following permissions must be available in coreview:
   - `Calendars.ReadWrite.All` - To read and cancel calendar events
   - `User.Read.All` - To read user information

### PowerShell Requirements
- **PowerShell 5.1 or later**: The script uses built-in PowerShell cmdlets
- **Internet connectivity**: Required for Microsoft Graph API calls

## Installation

### 1. Deploy the Custom Action

Use the deployment script to register the custom action with coreview:

```powershell
.\deploy-coreview-action.ps1 -CoreviewApiUrl "https://your-coreview-instance.com" -CoreviewApiKey "your-api-key"
```

### 2. Verify Installation

After deployment, verify that the custom action appears in your coreview interface under the "User Management" category.

### 3. Configure Permissions

Ensure that the appropriate roles have access to execute the custom action:
- **Execution Roles**: User Administrator, Help Desk Administrator
- **Approval Roles**: User Administrator, Global Administrator

## Usage

### Via Coreview Interface

1. Navigate to the custom actions section in coreview
2. Select "Offboard Microsoft 365 User"
3. Provide the required parameters:
   - **User Email**: The email address of the user to offboard
   - **Tenant ID**: (Optional) Specific tenant ID, otherwise uses current context
   - **Cancellation Comment**: (Optional) Custom comment for canceled events
4. Submit for approval if required
5. Monitor execution progress and review results

### Via Coreview API

```powershell
# Example API call to execute the custom action
$payload = @{
    actionName = "Offboard-M365User"
    parameters = @{
        UserEmail = "user@yourdomain.com"
        Comment = "User offboarding - calendar cleanup"
    }
}

$response = Invoke-RestMethod -Uri "$coreviewApiUrl/api/v1/custom-actions/execute" -Method Post -Headers $headers -Body ($payload | ConvertTo-Json)
```

### Via Coreview Workflows

The custom action can be integrated into coreview workflows for automated offboarding processes:

```json
{
  "workflow": "User Offboarding",
  "steps": [
    {
      "name": "Cancel Calendar Events",
      "action": "Offboard-M365User",
      "parameters": {
        "UserEmail": "{{user.email}}",
        "Comment": "Automated offboarding workflow"
      }
    }
  ]
}
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `UserEmail` | string | Yes | Email address of the user to offboard |
| `TenantId` | string | No | Specific tenant ID (uses current context if not provided) |
| `Comment` | string | No | Comment to include when canceling events |

## Outputs

The custom action provides structured outputs for integration:

| Output | Type | Description |
|--------|------|-------------|
| `TotalEvents` | integer | Total number of calendar events found |
| `SuccessfullyCanceled` | integer | Number of events successfully canceled |
| `FailedToCancel` | integer | Number of events that failed to cancel |
| `CanceledEvents` | array | Detailed information about canceled events |

## Security Features

### Authentication
- **Coreview-managed tokens**: No local credential storage
- **Automatic token refresh**: Handled by coreview infrastructure
- **Tenant isolation**: Proper tenant context management

### Authorization
- **Role-based access control**: Configurable execution and approval roles
- **Approval workflows**: Optional approval requirements for sensitive operations
- **Audit logging**: Complete audit trail of all operations

### Data Protection
- **No credential exposure**: Credentials are managed by coreview
- **Secure parameter handling**: Parameters are validated and sanitized
- **Error message sanitization**: Sensitive information is not exposed in error messages

## Monitoring and Auditing

### Audit Trail
All executions are logged to coreview's audit system with:
- User who initiated the action
- Parameters provided
- Execution results
- Timestamps and duration
- Success/failure status

### Monitoring
Monitor custom action usage through:
- Coreview audit logs
- Microsoft Graph API usage reports
- Custom action execution metrics

## Troubleshooting

### Common Issues

1. **"Failed to get access token from coreview"**
   - Verify coreview connectivity
   - Check API key permissions
   - Ensure Microsoft Graph permissions are available

2. **"Failed to get user"**
   - Verify the user email exists in the tenant
   - Check if the user is accessible with current permissions

3. **"Failed to get calendar events"**
   - Ensure Calendars.ReadWrite.All permission is available
   - Check if the user has any calendar events in the date range

4. **"Custom action not found"**
   - Verify the deployment was successful
   - Check if the action is available in your coreview instance
   - Ensure you have the required permissions to execute the action

### Debug Mode

Enable detailed logging by setting the `$VerbosePreference` variable:

```powershell
$VerbosePreference = "Continue"
```

### Support

For issues related to:
- **Coreview Integration**: Contact your coreview administrator
- **Microsoft Graph API**: Check the [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)
- **Script Functionality**: Review the audit logs and error messages

## Migration from Original Script

If you're migrating from the original script:

1. **Backup existing configurations**: Save any customizations to the original script
2. **Deploy the new custom action**: Use the deployment script
3. **Update workflows**: Modify any existing automation to use the new custom action
4. **Test thoroughly**: Verify functionality with test users before production use
5. **Update documentation**: Update any internal documentation to reference the new process

## Future Enhancements

Potential improvements for future versions:
- **Batch processing**: Handle multiple users in a single execution
- **Additional offboarding tasks**: Mailbox management, license removal, etc.
- **Conditional execution**: Skip certain steps based on user properties
- **Integration with HR systems**: Automatic triggering from HR offboarding workflows
- **Enhanced reporting**: Detailed reports with visualizations