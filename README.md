# M365 Email Integration for Frappe

A comprehensive Frappe app that integrates Microsoft 365 email using Azure AD Service Principal authentication. This app provides **bidirectional email integration** - both syncing incoming emails and sending outgoing emails via Microsoft Graph API.

## 🌟 Features

### Email Syncing (Incoming)
- ✅ **Incremental Delta Sync** - Efficient syncing using Microsoft Graph Delta API
- ✅ **Multi-Tenant Support** - Connect multiple Azure AD tenants
- ✅ **User & Shared Mailboxes** - Sync both personal and shared mailboxes
- ✅ **Folder Filtering** - Choose which folders to sync (Inbox, Sent Items, etc.)
- ✅ **Automatic Deduplication** - Prevents duplicate emails using Message-ID
- ✅ **Attachment Support** - Downloads and stores email attachments
- ✅ **Scheduled Sync** - Automatic syncing every 5 minutes
- ✅ **Manual Sync** - Trigger sync on-demand via API

### Email Sending (Outgoing)
- ✅ **Send via M365 Graph API** - No SMTP configuration required
- ✅ **Send As Any User** - Service Principal can send as any user in the organization
- ✅ **Automatic Queue Processing** - Emails are automatically sent via M365
- ✅ **No Email Account Required** - Bypasses Frappe's default Email Account requirement
- ✅ **Seamless Integration** - Works with `frappe.sendmail()` and Communication doctype
- ✅ **HTML Email Support** - Full HTML email formatting
- ✅ **Attachment Support** - Send emails with attachments

## 🏗️ Architecture

### How It Works

#### Email Syncing Flow
```
┌─────────────────────────────────────────────────────────────────┐
│ 1. Scheduled Task (Every 5 minutes)                             │
│    └─> m365email.m365email.tasks.sync_all_email_accounts()     │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 2. For Each Enabled M365 Email Account                          │
│    └─> Get Access Token from Service Principal                  │
│    └─> Call Microsoft Graph Delta API                           │
│        /users/{email}/messages/delta                             │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 3. Process Each Email Message                                   │
│    └─> Check if already synced (by Message-ID)                  │
│    └─> Create Communication document                            │
│    └─> Download attachments                                     │
│    └─> Store delta link for next sync                           │
└─────────────────────────────────────────────────────────────────┘
```

#### Email Sending Flow
```
┌─────────────────────────────────────────────────────────────────┐
│ 1. User/System calls frappe.sendmail() or Communication.make() │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 2. Monkey-Patched EmailAccount.find_outgoing()                  │
│    └─> Check if M365 sending is available                       │
│    └─> Return dummy account to bypass validation                │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 3. Email Queue Document Created                                 │
│    └─> before_insert hook triggers                              │
│    └─> Set m365_send = 1                                        │
│    └─> Set m365_account = sending account                       │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 4. Scheduled Task (Every minute)                                │
│    └─> m365email.m365email.send.process_email_queue_m365()     │
│    └─> Find emails with m365_send = 1                           │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 5. Send via Microsoft Graph API                                 │
│    └─> Parse MIME message from Email Queue                      │
│    └─> Extract subject, body, recipients                        │
│    └─> Call /users/{sender}/sendMail                            │
│    └─> Update Email Queue status                                │
└─────────────────────────────────────────────────────────────────┘
```

## 📦 Components

### DocTypes

#### M365 Email Service Principal Settings
Stores Azure AD Service Principal credentials for authentication.

**Fields:**
- `service_principal_name` - Unique identifier
- `tenant_id` - Azure AD Tenant ID
- `tenant_name` - Friendly tenant name
- `client_id` - Azure AD Application (Client) ID
- `client_secret` - Azure AD Client Secret (encrypted)
- `enabled` - Enable/disable this service principal
- `token_cache` - Encrypted OAuth token cache

**Permissions:**
- Only System Managers can create/edit

#### M365 Email Account
Represents a mailbox to sync (user or shared mailbox).

**Fields:**
- `account_name` - Unique identifier
- `account_type` - User Mailbox or Shared Mailbox
- `email_address` - M365 email address
- `user` - Linked Frappe user
- `service_principal` - Link to Service Principal Settings
- `enabled` - Enable/disable syncing
- `use_for_sending` - Mark this account for sending emails
- `folder_filters` - Child table of folders to sync
- `delta_link` - Stores delta sync state

**Permissions:**
- Users can create their own User Mailbox accounts
- Only System Managers can create Shared Mailbox accounts

#### M365 Email Sync Log
Tracks sync operations and errors.

**Fields:**
- `email_account` - Link to M365 Email Account
- `sync_type` - Manual or Scheduled
- `status` - Success, Failed, or Partial
- `messages_synced` - Count of messages synced
- `error_message` - Error details if failed

### Core Modules

#### `auth.py` - Authentication
- `get_access_token(service_principal_name)` - Get OAuth access token
- `refresh_all_tokens()` - Refresh tokens for all service principals
- Token caching and encryption

#### `graph_api.py` - Microsoft Graph API
- `get_delta_messages(email_address, delta_link, access_token)` - Incremental sync
- `send_email_as_user(sender_email, recipients, subject, body, ...)` - Send email
- `make_graph_request(endpoint, access_token, method, data)` - Generic API wrapper

#### `sync.py` - Email Syncing
- `sync_email_account(account_name)` - Sync a single account
- `process_message(message, account, access_token)` - Process one email
- `download_attachment(attachment_id, email_address, access_token)` - Get attachments
- `parse_m365_datetime(datetime_str)` - Parse M365 datetime format

#### `send.py` - Email Sending
- `get_sending_account()` - Get the M365 account marked for sending
- `can_send_via_m365()` - Check if M365 sending is available
- `intercept_email_queue(doc, method)` - Hook to mark emails for M365
- `send_via_m365(email_queue_doc)` - Send email via Graph API
- `process_email_queue_m365()` - Scheduled task to process queue

#### `email_override.py` - Core Email Override
- Overrides `frappe.core.doctype.communication.email.make`
- Bypasses email account check when M365 is available
- Registered in `hooks.py` as `override_whitelisted_methods`

#### `__init__.py` - Monkey Patches
- `patch_email_account()` - Patches `EmailAccount.find_outgoing()`
- Returns dummy account when M365 sending is available
- Applied automatically on app load

#### `custom_fields.py` - Custom Fields
- Adds `m365_send` and `m365_account` fields to Email Queue
- Applied during installation and via migration patch

#### `tasks.py` - Scheduled Tasks
- `sync_all_email_accounts()` - Sync all enabled accounts (every 5 min)
- `refresh_all_service_principal_tokens()` - Refresh tokens (hourly)
- `cleanup_old_sync_logs()` - Delete old logs (daily)
- `validate_service_principals()` - Check credentials (daily)

#### `api.py` - Whitelisted API Endpoints
- `enable_email_sync()` - Enable sync for an account
- `disable_email_sync()` - Disable sync
- `trigger_manual_sync()` - Manual sync trigger
- `get_sync_status()` - Get sync statistics
- `test_service_principal_connection()` - Test credentials

### Hooks Configuration

**`hooks.py`** registers:

```python
# Override Frappe's core email sending
override_whitelisted_methods = {
    "frappe.core.doctype.communication.email.make": "m365email.m365email.email_override.make"
}

# Hook into Email Queue creation
doc_events = {
    "Email Queue": {
        "before_insert": "m365email.m365email.send.intercept_email_queue"
    }
}

# Scheduled tasks
scheduler_events = {
    "cron": {
        "*/5 * * * *": [  # Every 5 minutes
            "m365email.m365email.tasks.sync_all_email_accounts"
        ],
        "* * * * *": [  # Every minute
            "m365email.m365email.send.process_email_queue_m365"
        ]
    },
    "hourly": [
        "m365email.m365email.tasks.refresh_all_service_principal_tokens"
    ],
    "daily": [
        "m365email.m365email.tasks.cleanup_old_sync_logs",
        "m365email.m365email.tasks.validate_service_principals"
    ]
}
```

## 🚀 Installation

See [README_SETUP.md](README_SETUP.md) for detailed setup instructions.

**Quick Start:**

1. Install dependencies:
   ```bash
   bench --site your-site pip install msal
   ```

2. Run migrations:
   ```bash
   bench --site your-site migrate
   ```

3. Restart bench:
   ```bash
   bench restart
   ```

4. Configure Azure AD (see setup guide)

5. Create Service Principal Settings in Frappe

6. Create M365 Email Accounts and enable syncing/sending

## 📧 Email Sending Setup

See [SENDING_SETUP.md](SENDING_SETUP.md) for detailed sending configuration.

**Quick Start:**

1. Ensure Service Principal has `Mail.Send` permission in Azure AD

2. Mark one M365 Email Account for sending:
   - Open M365 Email Account
   - Check "Use for Sending"
   - Save

3. Send emails normally:
   ```python
   frappe.sendmail(
       recipients=["user@example.com"],
       sender="your-email@company.com",
       subject="Test Email",
       message="<p>Hello from M365!</p>"
   )
   ```

4. Emails are automatically sent via M365 Graph API!

## 🔒 Security

- **Client Secrets**: Encrypted by Frappe's encryption system
- **Token Cache**: Encrypted before storage in database
- **Service Principal Auth**: Uses application permissions (no user passwords)
- **Role-Based Access**: System Manager for shared mailboxes, users for personal
- **Communication Permissions**: Standard Frappe permission system

## 🧪 Testing & Debugging

### Debug Helpers

```python
from m365email.m365email.debug_helpers import *

# Check M365 sending configuration
check_m365_sending_config()

# Check Email Queue status
check_email_queue_status()

# Manually process email queue
manually_process_queue()

# Check recent errors
check_recent_errors()
```

### Manual Sync

```python
from m365email.m365email.sync import sync_email_account

# Sync a specific account
sync_email_account("Account Name")
```

## 📊 Monitoring

- **M365 Email Sync Log** - View sync history and errors
- **Email Queue** - Monitor outgoing emails
- **Communication** - View all synced emails
- **Error Log** - Check for M365-related errors

## 🤝 Contributing

Contributions are welcome! Please ensure:
- Code follows Frappe coding standards
- All new features include documentation
- Test thoroughly before submitting PR

## 📄 License

MIT License - See LICENSE file for details

## 🆘 Support

- **Setup Guide**: [README_SETUP.md](README_SETUP.md)
- **Sending Guide**: [SENDING_SETUP.md](SENDING_SETUP.md)
- **Frappe Forum**: https://discuss.frappe.io
- **GitHub Issues**: Report bugs and feature requests

---

**Built with ❤️ for the Frappe community**

