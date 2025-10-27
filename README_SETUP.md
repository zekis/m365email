# M365 Email Integration - Setup Guide

## Installation

1. **Install the app dependencies:**
   ```bash
   cd /path/to/frappe-bench
   bench --site your-site.local pip install msal
   ```

2. **Run migrations to create DocTypes:**
   ```bash
   bench --site your-site.local migrate
   ```

3. **Restart bench:**
   ```bash
   bench restart
   ```

## Azure AD Configuration

### 1. Create Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Enter a name (e.g., "Frappe M365 Email Integration")
5. Select **Accounts in this organizational directory only**
6. Click **Register**

### 2. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph** > **Application permissions**
3. Add the following permissions:
   - `Mail.Read` - Read mail in all mailboxes
   - `Mail.ReadWrite` - Read and write mail in all mailboxes
   - `User.Read.All` - Read all users' profiles
   - `MailboxSettings.Read` - Read all mailbox settings
4. Click **Grant admin consent** for your organization

### 3. Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Enter a description and select expiration period
4. Click **Add**
5. **Copy the secret value immediately** (you won't be able to see it again)

### 4. Note Your Credentials

You'll need:
- **Tenant ID** (from Overview page)
- **Client ID** (Application ID from Overview page)
- **Client Secret** (the value you just copied)

## Frappe Configuration

### 1. Create Service Principal Settings

1. In Frappe, go to **M365 Email Service Principal Settings**
2. Click **New**
3. Fill in the details:
   - **Service Principal Name**: A unique name (e.g., "Main Tenant")
   - **Tenant ID**: Your Azure AD Tenant ID
   - **Tenant Name**: Friendly name (e.g., "Contoso Corp")
   - **Client ID**: Application (client) ID from Azure
   - **Client Secret**: The secret value you copied
   - **Enabled**: Check this box
4. Click **Save**
5. Click **Test Connection** to verify credentials

### 2. Enable Email Sync for Users

#### For User Mailboxes:

Users can enable their own mailbox sync:

1. Go to **M365 Email Account**
2. Click **New**
3. Fill in:
   - **Account Name**: Unique name (e.g., "john-doe-email")
   - **Account Type**: User Mailbox
   - **Email Address**: User's M365 email
   - **User**: Select the Frappe user
   - **Service Principal**: Select the service principal
   - **Enabled**: Check this box
4. Configure folder filters (Inbox, Sent Items, etc.)
5. Click **Save**

#### For Shared Mailboxes (System Manager only):

1. Go to **M365 Email Account**
2. Click **New**
3. Fill in:
   - **Account Name**: Unique name (e.g., "support-mailbox")
   - **Account Type**: Shared Mailbox
   - **Email Address**: Shared mailbox email (e.g., support@company.com)
   - **User**: Select yourself or another admin
   - **Service Principal**: Select the service principal
   - **Enabled**: Check this box
4. Configure folder filters
5. Click **Save**

### 3. Grant Access to Shared Mailbox Emails

For team members to access shared mailbox emails:

1. Go to **Role Permission Manager**
2. Find the **Inbox User** role
3. Assign this role to users who need access to shared mailbox emails
4. Users with Inbox User role can view all Communications (including shared mailbox emails)

## Scheduled Tasks

The following tasks run automatically:

- **Every 5 minutes**: Sync all enabled email accounts
- **Hourly**: Refresh access tokens for all service principals
- **Daily**: 
  - Cleanup old sync logs (older than 30 days)
  - Validate service principal credentials

## API Endpoints

Available whitelisted API endpoints:

- `m365email.m365email.api.enable_email_sync` - Enable email sync
- `m365email.m365email.api.disable_email_sync` - Disable email sync
- `m365email.m365email.api.trigger_manual_sync` - Manually trigger sync
- `m365email.m365email.api.get_sync_status` - Get sync status
- `m365email.m365email.api.test_service_principal_connection` - Test connection
- `m365email.m365email.api.get_available_service_principals` - List service principals
- `m365email.m365email.api.get_shared_mailboxes` - List shared mailboxes
- `m365email.m365email.api.get_available_folders` - Get mail folders
- `m365email.m365email.api.update_folder_filters` - Update folder filters

## Troubleshooting

### Sync Not Working

1. Check **M365 Email Sync Log** for errors
2. Verify service principal credentials are correct
3. Ensure Azure AD permissions are granted
4. Check that the email account is enabled
5. Verify the scheduler is running: `bench doctor`

### Token Errors

1. Go to **M365 Email Service Principal Settings**
2. Click **Test Connection**
3. If it fails, verify:
   - Client ID is correct
   - Client Secret is correct and not expired
   - Tenant ID is correct
   - Admin consent was granted for API permissions

### Emails Not Appearing

1. Check if user has **Inbox User** role (for shared mailboxes)
2. Verify folder filters are configured correctly
3. Check **Communication** list for synced emails
4. Review **M365 Email Sync Log** for sync statistics

## Security Notes

- Client secrets are encrypted by Frappe
- Token cache is encrypted before storage
- Only System Managers can configure service principals
- Only System Managers can configure shared mailboxes
- Users can only configure their own user mailboxes
- Communication permissions control who can view emails

## Multi-Tenant Support

You can configure multiple Azure AD tenants:

1. Create separate **M365 Email Service Principal Settings** for each tenant
2. Each email account links to a specific service principal
3. Tokens are managed independently per tenant

## Support

For issues or questions, please refer to:
- Feature documentation: `docs/feature/desc.md`
- Frappe Forum: https://discuss.frappe.io
- GitHub Issues: [Your repository URL]

