# Microsoft 365 User Offboarding Script

This PowerShell script cancels all calendar invites for a specified user in Microsoft 365 using only API calls, without requiring any external PowerShell modules.

## Prerequisites

1. **Azure App Registration**: You need to create an Azure App Registration with the following permissions:
   - `Calendars.ReadWrite.All` - To read and cancel calendar events
   - `User.Read.All` - To read user information

2. **PowerShell 5.1 or later**: The script uses built-in PowerShell cmdlets only.

## Setup Instructions

### 1. Create Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Enter a name for your app (e.g., "M365 Offboarding Script")
5. Select **Accounts in this organizational directory only**
6. Click **Register**

### 2. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Add the following permissions:
   - `Calendars.ReadWrite.All`
   - `User.Read.All`
6. Click **Grant admin consent** (requires admin privileges)

### 3. Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and select expiration
4. **Copy the secret value immediately** (you won't be able to see it again)

### 4. Get Required Information

You'll need the following information:
- **Client ID**: Found in the app registration overview
- **Client Secret**: The secret value you just created
- **Tenant ID**: Found in Azure AD overview or app registration overview

## Usage

### Basic Usage

```powershell
.\Offboard-M365User.ps1 -UserEmail "user@yourdomain.com" -ClientId "your-client-id" -ClientSecret "your-client-secret" -TenantId "your-tenant-id"
```

### Example

```powershell
.\Offboard-M365User.ps1 -UserEmail "john.doe@contoso.com" -ClientId "12345678-1234-1234-1234-123456789012" -ClientSecret "your-secret-here" -TenantId "87654321-4321-4321-4321-210987654321"
```

## What the Script Does

1. **Authentication**: Uses client credentials flow to obtain an access token
2. **User Lookup**: Retrieves the user's details by email address
3. **Calendar Events**: Gets all calendar events for the user
4. **Event Cancellation**: Cancels each calendar event with a comment indicating it's part of the offboarding process

## Output

The script provides detailed output including:
- Progress indicators for each step
- Success/failure counts for event cancellations
- Summary of all operations performed

## Error Handling

The script includes comprehensive error handling for:
- Authentication failures
- User not found
- API rate limiting
- Network connectivity issues

## Security Considerations

- Store client secrets securely (consider using Azure Key Vault)
- Use least-privilege permissions
- Regularly rotate client secrets
- Monitor script usage through Azure AD audit logs

## Troubleshooting

### Common Issues

1. **"Failed to get access token"**
   - Verify Client ID, Client Secret, and Tenant ID
   - Ensure the app has the correct permissions
   - Check if admin consent has been granted

2. **"Failed to get user"**
   - Verify the user email exists in your tenant
   - Ensure the app has User.Read.All permission

3. **"Failed to get calendar events"**
   - Ensure the app has Calendars.ReadWrite.All permission
   - Check if the user has any calendar events

### Debug Mode

To see more detailed error information, you can modify the script to include `-Verbose` parameters or add additional logging.

## Next Steps

This script handles calendar event cancellation. For a complete offboarding process, you may want to add:

1. **Mailbox Management**: Forward emails, set up auto-replies
2. **License Removal**: Remove Microsoft 365 licenses
3. **Group Membership**: Remove from all groups
4. **Account Disabling**: Disable the user account
5. **OneDrive Cleanup**: Handle OneDrive files
6. **Teams Cleanup**: Handle Teams memberships and chats

## Support

For issues related to:
- **Azure App Registration**: Contact your Azure administrator
- **Microsoft Graph API**: Check the [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)
- **Script functionality**: Review the error messages and check the prerequisites