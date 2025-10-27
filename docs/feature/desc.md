# M365 Email Integration with Service Principal Authentication

## Feature Overview

This Frappe app enables Microsoft 365 email integration using Azure AD Service Principal (Application) authentication. It allows company administrators to configure a centralized M365 app registration, enabling users to sync their email accounts without individual OAuth consent flows.

### Key Benefits
- **Multi-Tenant Support**: Support multiple Azure AD tenants in a single Frappe instance
- **Centralized Administration**: Service principal configuration per tenant/organization
- **Shared Mailbox Support**: First-class support for shared mailboxes (support@, sales@, etc.)
- **Simplified User Experience**: Users don't need to perform OAuth consent
- **Enhanced Security**: Uses application permissions with proper scoping
- **Automated Sync**: Background email synchronization with Frappe Communication doctype
- **Token Management**: Automatic token refresh and lifecycle management

---

## Multi-Tenant Architecture

This app supports multiple Azure AD tenants in a single Frappe instance. This is useful for:

1. **Multi-Company Scenarios**: Different companies with different M365 tenants
2. **MSP/Partner Scenarios**: Service providers managing multiple client tenants
3. **Merger/Acquisition**: Organizations with multiple legacy tenants
4. **Development/Testing**: Separate tenants for dev, staging, production

### How It Works:
- Each M365 Service Principal Settings record represents one Azure AD tenant
- Email accounts link to a specific service principal (tenant)
- Tokens are managed independently per tenant
- Scheduled tasks process all tenants in sequence
- Users can have email accounts across multiple tenants

### Shared Mailbox Support

Shared mailboxes (e.g., support@company.com, sales@company.com) are first-class citizens:

1. **No Separate Credentials**: Accessed via service principal with application permissions
2. **Admin Configuration**: Only System Managers can configure shared mailbox sync
3. **Same Sync Logic**: Uses identical Graph API endpoints as user mailboxes
4. **Communication Storage**: Emails sync to Communication doctype like user mailboxes
5. **Standard Permissions**: Access controlled via Frappe's Communication permissions (Inbox User role)
6. **Communication Linking**: Emails can be linked to relevant doctypes (e.g., Support Ticket, Lead)

### Key Differences: User vs Shared Mailbox

| Feature | User Mailbox | Shared Mailbox |
|---------|--------------|----------------|
| Account Type | User Mailbox | Shared Mailbox |
| Email Address | User's personal email | Shared mailbox email (support@, sales@, etc.) |
| User Field | Required (the user) | Required (admin who configured it) |
| Who Can Configure | The user themselves | System Manager only |
| Communication Owner | The user | The admin who configured the sync |
| Access Control | User owns their Communications | Standard Communication permissions (Inbox User) |
| Use Case | Personal email sync | Team/department email sync |

---

## Architecture Components

### 1. DocTypes

#### 1.1 M365 Service Principal Settings
**Purpose**: Store M365 app registration credentials per tenant (supports multiple tenants)

**Fields**:
- `service_principal_name` (Data) - Unique name for this configuration
- `tenant_id` (Data) - Azure AD Tenant ID
- `tenant_name` (Data) - Friendly tenant name (e.g., "Contoso Corp")
- `client_id` (Data) - Application (client) ID
- `client_secret` (Password) - Client secret value
- `authority_url` (Data) - Default: `https://login.microsoftonline.com/{tenant_id}`
- `graph_api_endpoint` (Data) - Default: `https://graph.microsoft.com/v1.0`
- `enabled` (Check) - Enable/disable this service principal
- `token_cache` (Long Text) - Encrypted token cache storage
- `last_token_refresh` (Datetime) - Last successful token refresh
- `token_expires_at` (Datetime) - Token expiration time
- `scopes` (Small Text) - Default: `https://graph.microsoft.com/.default`
- `company` (Link: Company) - Associated company (optional)
- `description` (Text Editor) - Notes about this configuration

**Naming**: Auto-name with field `service_principal_name`

**Permissions**:
- Only System Manager can create/edit
- Read-only for other roles

#### 1.2 M365 Email Account
**Purpose**: Configure email sync for user mailboxes OR shared mailboxes

**Fields**:
- `account_name` (Data) - Unique identifier for this account
- `account_type` (Select) - "User Mailbox" or "Shared Mailbox"
- `email_address` (Data) - M365 email address to sync
- `user` (Link: User) - Frappe user (for User Mailbox: the user; for Shared: admin who configured it)
- `enabled` (Check) - Enable sync for this account
- `service_principal` (Link: M365 Service Principal Settings) - Which tenant/config to use
- `sync_from_date` (Date) - Start syncing emails from this date
- `last_sync_time` (Datetime) - Last successful sync timestamp
- `last_sync_status` (Select) - Success, Failed, In Progress
- `sync_error_message` (Text) - Last error message if failed
- `folder_filter` (Table: M365 Folder Filter) - Which folders to sync
- `auto_create_contact` (Check) - Auto-create contacts from email addresses
- `sync_attachments` (Check) - Download and attach email attachments
- `max_attachment_size` (Int) - Max attachment size in MB (default: 10)
- `delta_tokens` (Long Text) - JSON storage for delta tokens per folder

**Child Table: M365 Folder Filter**:
- `folder_name` (Data) - Folder name (Inbox, Sent Items, Drafts, etc.)
- `sync_enabled` (Check)
- `last_sync_time` (Datetime)
- `delta_token` (Data) - For incremental sync

**Naming**: Auto-name with field `account_name` or pattern like "M365-{email_address}"

**Permissions**:
- Users can view/edit their own email account (where user = current user AND account_type = "User Mailbox")
- System Manager can view/edit all accounts
- Only System Manager can create Shared Mailbox accounts

**Validation**:
- User field is always mandatory
- For User Mailbox: email_address should match user's email (or be one of their email accounts)
- For Shared Mailbox: email_address is the shared mailbox email
- Email address must be unique per service principal

**Notes**:
- Communications created from this account will be owned by the `user` field
- For shared mailboxes, access to synced emails is controlled by Communication permissions (Inbox User role)
- Admins configuring shared mailboxes should have access to that mailbox in M365

#### 1.3 M365 Email Sync Log
**Purpose**: Track sync operations and errors

**Fields**:
- `email_account` (Link: M365 Email Account)
- `sync_start_time` (Datetime)
- `sync_end_time` (Datetime)
- `status` (Select) - Success, Partial Success, Failed
- `emails_fetched` (Int)
- `emails_created` (Int)
- `emails_skipped` (Int)
- `error_message` (Long Text)
- `sync_type` (Select) - Manual, Scheduled, Initial
- `folder_name` (Data) - Which folder was synced

**Auto-delete**: Logs older than 30 days



---

### 2. Python Functions/Modules

#### 2.1 Authentication Module (`m365email/auth.py`)

**Functions**:
```python
def get_access_token(service_principal_name):
    """
    Get valid access token for specified service principal
    Uses MSAL library with client credentials flow
    Returns cached token if valid, otherwise requests new token
    Supports multiple tenants
    """

def refresh_token(service_principal_name):
    """
    Force refresh access token for specified service principal
    Updates token_cache in M365 Service Principal Settings
    """

def validate_service_principal(service_principal_doc):
    """
    Validate service principal credentials
    Used during initial setup and periodic validation
    Returns dict with success status and error message
    """

def get_msal_app(service_principal_name):
    """
    Initialize MSAL ConfidentialClientApplication for specified tenant
    Returns configured MSAL app instance
    """

def get_service_principal_for_email_account(email_account_name):
    """
    Get the service principal associated with an email account
    Returns M365 Service Principal Settings doc
    """
```

#### 2.2 Email Sync Module (`m365email/sync.py`)

**Functions**:
```python
def sync_email_account(email_account_name, folder_name=None):
    """
    Main sync function for an email account (user or shared mailbox)
    Fetches emails from M365 and creates Communications
    Handles both initial sync and incremental delta sync
    """

def fetch_messages_from_folder(email_account, folder_name, delta_token=None):
    """
    Fetch messages from specific folder using Graph API
    Supports delta queries for incremental sync
    Works for both user mailboxes and shared mailboxes
    """

def create_communication_from_message(message_data, email_account):
    """
    Convert M365 message to Frappe Communication doctype
    Handles attachments, recipients, threading
    Sets owner to email_account.user (the user who configured the sync)
    For shared mailboxes, can optionally link to reference doctypes
    """

def download_attachment(message_id, attachment_id, email_account):
    """
    Download email attachment from M365
    Returns file content and metadata
    Respects max_attachment_size setting
    """

def sync_all_enabled_accounts():
    """
    Sync all enabled email accounts (both user and shared mailboxes)
    Called by scheduled task
    Processes accounts across all service principals/tenants
    """

def get_delta_token(email_account, folder_name):
    """
    Get stored delta token for incremental sync from folder_filter child table
    Returns None for initial sync
    """

def save_delta_token(email_account, folder_name, delta_token):
    """
    Save delta token in folder_filter child table for next incremental sync
    """

def sync_shared_mailbox(email_account_name):
    """
    Specialized sync for shared mailboxes
    Handles permission checks and multi-user access
    """
```

#### 2.3 Graph API Module (`m365email/graph_api.py`)

**Functions**:
```python
def make_graph_request(endpoint, access_token, method='GET', data=None):
    """
    Generic Graph API request handler
    Handles pagination, rate limiting, error handling
    """

def get_user_messages(user_email, access_token, folder='inbox', top=50):
    """
    Get messages for specific user
    Endpoint: /users/{email}/messages
    """

def get_message_details(user_email, message_id, access_token):
    """
    Get full message details including body
    """

def get_message_attachments(user_email, message_id, access_token):
    """
    Get all attachments for a message
    """

def mark_message_as_read(user_email, message_id, access_token):
    """
    Mark message as read in M365
    """
```

#### 2.4 API Module (`m365email/api.py`)

**Whitelisted Functions**:
```python
@frappe.whitelist()
def enable_email_sync(email_address, service_principal):
    """
    Enable email sync for current user
    Creates M365 Email Account if not exists
    User selects which service principal/tenant to use
    """

@frappe.whitelist()
def enable_shared_mailbox_sync(email_address, service_principal):
    """
    Enable sync for a shared mailbox
    Creates M365 Email Account with type = "Shared Mailbox"
    Sets current user (admin) as the owner
    Requires System Manager role
    """

@frappe.whitelist()
def trigger_manual_sync(email_account_name=None):
    """
    Trigger manual sync for specified email account
    If not specified, syncs current user's account
    """

@frappe.whitelist()
def get_sync_status(email_account_name=None):
    """
    Get sync status and recent logs
    For current user's account or specified shared mailbox (if has access)
    """

@frappe.whitelist()
def test_service_principal_connection(service_principal_name):
    """
    Test service principal credentials
    System Manager only
    Returns connection status and token validity
    """

@frappe.whitelist()
def get_available_service_principals():
    """
    Get list of enabled service principals
    For user to select when enabling email sync
    """

@frappe.whitelist()
def get_shared_mailboxes():
    """
    Get all configured shared mailboxes
    System Manager sees all, others see none (unless they have Inbox User role for Communications)
    """
```

#### 2.5 Utilities Module (`m365email/utils.py`)

**Functions**:
```python
def get_service_principal_for_user(user):
    """
    Get appropriate service principal settings for user
    Based on company or returns list of available service principals
    """

def get_service_principal_by_name(service_principal_name):
    """
    Get service principal doc by name
    Validates it's enabled
    """

def encrypt_token_cache(token_data):
    """
    Encrypt token cache before storing
    Uses Frappe's encryption utilities
    """

def decrypt_token_cache(encrypted_data):
    """
    Decrypt token cache for use
    """

def parse_email_address(email_string):
    """
    Parse email addresses from various formats
    Handles display names and multiple addresses
    """

def should_sync_message(message, email_account):
    """
    Determine if message should be synced based on filters
    Checks date range, folder filters, etc.
    """

def is_shared_mailbox(email_address, service_principal):
    """
    Check if an email address is a shared mailbox
    Can query M365 or check mailbox type
    """

def user_can_configure_account(user, email_account):
    """
    Check if user can configure email account settings
    For user mailboxes: checks if user owns it
    For shared mailboxes: checks if user is System Manager
    """

def get_communication_reference(message_data, email_account):
    """
    Determine reference_doctype and reference_name for Communication
    Can implement logic to auto-link emails to doctypes (e.g., Support Ticket by email parsing)
    Returns tuple: (reference_doctype, reference_name) or (None, None)
    """
```

---

### 3. Scheduled Tasks

Add to `hooks.py`:

```python
scheduler_events = {
    "cron": {
        # Sync emails every 5 minutes
        "*/5 * * * *": [
            "m365email.sync.sync_all_enabled_accounts"
        ]
    },
    "hourly": [
        # Refresh tokens every hour
        "m365email.auth.refresh_all_tokens"
    ],
    "daily": [
        # Cleanup old sync logs
        "m365email.tasks.cleanup_old_sync_logs",
        # Validate all service principals
        "m365email.tasks.validate_service_principals"
    ]
}
```

**Task Functions** (`m365email/tasks.py`):
```python
def cleanup_old_sync_logs():
    """Delete sync logs older than 30 days"""

def validate_service_principals():
    """
    Check all service principals are still valid
    Tests connection for each enabled service principal
    Disables any with invalid credentials
    """

def refresh_all_tokens():
    """
    Refresh tokens for all enabled service principals
    Handles multiple tenants
    Updates token_cache and token_expires_at
    """
```

---

### 4. API Endpoints (Graph API)

**Required Microsoft Graph API Permissions**:
- `Mail.Read` (Application) - Read mail in all mailboxes (including shared)
- `Mail.ReadWrite` (Application) - Read and write mail in all mailboxes (if marking as read)
- `User.Read.All` (Application) - Read all users' profiles
- `MailboxSettings.Read` (Application) - Read mailbox settings (to identify shared mailboxes)

**Key Endpoints Used**:

*For both User and Shared Mailboxes (same endpoints):*
- `GET /users/{email}/messages` - List messages
- `GET /users/{email}/mailFolders` - List all folders
- `GET /users/{email}/mailFolders/{folder}/messages` - Messages in specific folder
- `GET /users/{email}/mailFolders/{folder}/messages/delta` - Delta query for incremental sync
- `GET /users/{email}/messages/{id}` - Get message details
- `GET /users/{email}/messages/{id}/attachments` - Get attachments
- `PATCH /users/{email}/messages/{id}` - Update message (mark as read)

*For Shared Mailbox Discovery:*
- `GET /users/{email}/mailboxSettings` - Get mailbox type and settings
- `GET /users` - List all users (to discover shared mailboxes in tenant)

**Note**: The Graph API treats shared mailboxes the same as user mailboxes. The key difference is that shared mailboxes don't have sign-in capability, but can be accessed via application permissions using their email address.

---

### 5. Communication Permissions & Access Control

#### How Communication Permissions Work

Frappe's Communication doctype has built-in permission rules:

1. **System Manager**: Full access to all Communications
2. **Inbox User**: Can read, create, delete, export, print, and email Communications
3. **All Users**: Can only read/delete/email Communications they OWN

#### Implications for M365 Email Sync

**User Mailboxes**:
- Communications are owned by the user whose mailbox is synced
- User can see their own synced emails in Communication list
- Other users cannot see these Communications (unless they have Inbox User or System Manager role)

**Shared Mailboxes**:
- Communications are owned by the admin who configured the shared mailbox sync
- The owner (admin) can see all synced emails
- To grant access to team members: Assign them the **Inbox User** role
- Inbox Users can view, manage, and respond to all Communications (including shared mailbox emails)
- This is the standard Frappe pattern for shared email access

#### Recommended Role Setup

For teams using shared mailboxes:
1. **System Manager**: Configure service principals and shared mailbox accounts
2. **Inbox User**: Grant to team members who need to access shared mailbox emails
3. **Regular Users**: Can only see their own personal email sync (if enabled)

**Note**: The Inbox User role is a standard Frappe role designed for email management. It provides appropriate access without requiring System Manager privileges.

---

### 6. Security Considerations

1. **Credential Storage**:
   - Client secret stored in Password field (encrypted by Frappe)
   - Token cache encrypted before storage
   - No tokens in logs or error messages

2. **Access Control**:
   - Service Principal Settings: System Manager only
   - User Mailbox Accounts: User can only configure their own
   - Shared Mailbox Accounts: System Manager only
   - Communication access: Standard Frappe Communication permissions (see above)
   - API endpoints validate user permissions

3. **Audit Trail**:
   - All sync operations logged
   - Failed authentication attempts logged
   - Token refresh events tracked

4. **Data Privacy**:
   - Users can disable sync anytime
   - Option to not sync attachments
   - Folder-level filtering

5. **M365 Permissions**:
   - Admin configuring shared mailbox should have access to it in M365
   - Service principal requires application-level permissions (not delegated)
   - Regular users don't need M365 access - sync happens via service principal

---

### 7. User Workflows

#### Admin Setup (One-time per Tenant):
1. **Azure AD Configuration**:
   - Create Azure AD App Registration in target tenant
   - Configure API permissions (Mail.Read, Mail.ReadWrite, User.Read.All, MailboxSettings.Read)
   - Grant admin consent for the organization
   - Create client secret (note expiration date)

2. **Frappe Configuration**:
   - Navigate to M365 Service Principal Settings
   - Create new record for each tenant
   - Enter service principal name, tenant ID, client ID, client secret
   - Test connection
   - Enable the service principal
   - Optionally link to a Company

#### User Mailbox Setup:
1. User navigates to M365 Email Account list
2. Clicks "New" or "Enable Email Sync" button
3. Selects service principal (tenant) from dropdown
4. System auto-fills email address from user profile
5. User configures:
   - Sync from date
   - Folders to sync (Inbox, Sent Items, etc.)
   - Attachment settings
6. Saves and sync starts automatically on next scheduled run
7. User can trigger manual sync immediately

#### Shared Mailbox Setup (System Manager Only):
1. System Manager navigates to M365 Email Account
2. Creates new account with:
   - Account Type = "Shared Mailbox"
   - Email address (e.g., support@company.com)
   - User = themselves (or another admin with M365 access to the shared mailbox)
   - Service principal selection
   - Sync preferences (folders, attachments, etc.)
3. Saves and sync starts automatically on next scheduled run
4. Communications created from this mailbox are owned by the configured user
5. Users with Inbox User role can view/manage these Communications
6. Optionally: Grant Inbox User role to team members who need access to shared mailbox emails

#### Ongoing Operation:
1. Scheduled task runs every 5 minutes
2. Processes all enabled accounts (user + shared mailboxes)
3. Uses delta queries for incremental sync
4. Creates Communication records
5. Downloads attachments (if enabled)
6. Updates sync status and logs
7. Tokens refreshed hourly across all tenants

---

### 8. Dependencies

**Python Packages** (add to `pyproject.toml`):
```toml
dependencies = [
    "msal>=1.24.0",  # Microsoft Authentication Library
    "requests>=2.31.0",  # HTTP requests
]
```

**Frappe Version**: 14.x or 15.x

---

### 9. Future Enhancements

1. **Send Email**: Support sending emails via Graph API
2. **Calendar Sync**: Sync M365 calendar events
3. **Contacts Sync**: Sync M365 contacts
4. **Email Rules**: Auto-assign emails to specific doctypes based on rules
5. **Two-way Sync**: Sync read status and flags back to M365
6. **Advanced Filtering**: Filter by sender, subject, categories, importance
7. **Webhook Support**: Real-time sync using Microsoft Graph webhooks (subscriptions)
8. **Email Templates**: Send emails using Frappe templates via M365
9. **Conversation Threading**: Better email thread management
10. **Search Integration**: Search M365 emails directly from Frappe

---

## Implementation Checklist

### Phase 1: Core DocTypes
- [ ] Create M365 Service Principal Settings DocType (normal, not Single)
- [ ] Create M365 Email Account DocType with account_type field
- [ ] Create M365 Email Sync Log DocType
- [ ] Create M365 Folder Filter Child Table

### Phase 2: Backend Modules
- [ ] Implement authentication module (auth.py)
  - [ ] Multi-tenant token management
  - [ ] MSAL integration
- [ ] Implement Graph API module (graph_api.py)
  - [ ] Generic request handler
  - [ ] Shared mailbox support
- [ ] Implement sync module (sync.py)
  - [ ] User mailbox sync
  - [ ] Shared mailbox sync
  - [ ] Delta query support
- [ ] Implement API endpoints (api.py)
  - [ ] User mailbox endpoints
  - [ ] Shared mailbox endpoints
- [ ] Implement utilities (utils.py)
  - [ ] Permission helpers
  - [ ] Encryption utilities
- [ ] Implement scheduled tasks (tasks.py)
  - [ ] Multi-tenant token refresh
  - [ ] Account sync across tenants

### Phase 3: Integration
- [ ] Add scheduler events to hooks.py
- [ ] Configure permissions and roles
- [ ] Add validation rules

### Phase 4: User Interface
- [ ] Create custom forms for Service Principal Settings
- [ ] Create custom forms for Email Account (with account type toggle)
- [ ] Add client-side scripts for UX
- [ ] Create dashboard for sync status
- [ ] Add shared mailbox access management UI

### Phase 5: Testing & Documentation
- [ ] Write unit tests for auth module
- [ ] Write unit tests for sync module
- [ ] Write integration tests for multi-tenant scenarios
- [ ] Test shared mailbox functionality
- [ ] Create admin setup guide (Azure + Frappe)
- [ ] Create user documentation
- [ ] Create shared mailbox setup guide

### Phase 6: Security & Performance
- [ ] Security audit of credential storage
- [ ] Performance testing with large mailboxes
- [ ] Rate limiting implementation
- [ ] Error handling and retry logic

