# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
API endpoints for M365 Email Integration
Whitelisted functions for frontend/client access
"""

import frappe
from frappe import _
from m365email.m365email.auth import test_connection, get_access_token
from m365email.m365email.sync import sync_email_account
from m365email.m365email.graph_api import get_mail_folders
from m365email.m365email.utils import user_can_configure_account


@frappe.whitelist()
def enable_email_sync(email_address, service_principal, account_type="User Mailbox"):
	"""
	Enable email sync for current user's mailbox
	
	Args:
		email_address: User's email address
		service_principal: Service principal name
		account_type: "User Mailbox" or "Shared Mailbox"
		
	Returns:
		dict: Created email account details
	"""
	user = frappe.session.user
	
	# For Shared Mailbox, only System Manager can create
	if account_type == "Shared Mailbox":
		if "System Manager" not in frappe.get_roles(user):
			frappe.throw(_("Only System Manager can configure shared mailboxes"))
	
	# Check if account already exists
	existing = frappe.db.exists(
		"M365 Email Account",
		{
			"email_address": email_address,
			"service_principal": service_principal
		}
	)
	
	if existing:
		frappe.throw(_("Email account already exists for this email address and service principal"))
	
	# Create email account
	account = frappe.get_doc({
		"doctype": "M365 Email Account",
		"account_name": f"M365-{email_address}",
		"account_type": account_type,
		"email_address": email_address,
		"user": user,
		"service_principal": service_principal,
		"enabled": 1,
		"folder_filter": [
			{"folder_name": "Inbox", "sync_enabled": 1},
			{"folder_name": "Sent Items", "sync_enabled": 0}
		]
	})
	account.insert()
	frappe.db.commit()
	
	return {
		"success": True,
		"account_name": account.name,
		"message": _("Email sync enabled successfully")
	}


@frappe.whitelist()
def disable_email_sync(email_account_name):
	"""
	Disable email sync for an account
	
	Args:
		email_account_name: Name of M365 Email Account
		
	Returns:
		dict: Success message
	"""
	account = frappe.get_doc("M365 Email Account", email_account_name)
	
	# Check permissions
	if not user_can_configure_account(frappe.session.user, account):
		frappe.throw(_("You don't have permission to configure this email account"))
	
	account.enabled = 0
	account.save()
	frappe.db.commit()
	
	return {
		"success": True,
		"message": _("Email sync disabled successfully")
	}


@frappe.whitelist()
def trigger_manual_sync(email_account_name):
	"""
	Manually trigger sync for an email account
	
	Args:
		email_account_name: Name of M365 Email Account
		
	Returns:
		dict: Sync results
	"""
	account = frappe.get_doc("M365 Email Account", email_account_name)
	
	# Check permissions
	if not user_can_configure_account(frappe.session.user, account):
		frappe.throw(_("You don't have permission to sync this email account"))
	
	result = sync_email_account(email_account_name)
	
	return result


@frappe.whitelist()
def get_sync_status(email_account_name=None):
	"""
	Get sync status and recent logs
	
	Args:
		email_account_name: Name of M365 Email Account (optional)
		
	Returns:
		dict: Sync status information
	"""
	user = frappe.session.user
	
	if email_account_name:
		# Get status for specific account
		account = frappe.get_doc("M365 Email Account", email_account_name)
		
		# Check permissions
		if not user_can_configure_account(user, account):
			frappe.throw(_("You don't have permission to view this email account"))
		
		# Get recent logs
		logs = frappe.get_all(
			"M365 Email Sync Log",
			filters={"email_account": email_account_name},
			fields=["name", "sync_type", "status", "start_time", "end_time", "messages_fetched", "messages_created"],
			order_by="start_time desc",
			limit=10
		)
		
		return {
			"account": account.as_dict(),
			"logs": logs
		}
	else:
		# Get all accounts for current user
		filters = {"user": user}
		if "System Manager" in frappe.get_roles(user):
			filters = {}  # System Manager can see all
		
		accounts = frappe.get_all(
			"M365 Email Account",
			filters=filters,
			fields=["name", "account_name", "account_type", "email_address", "enabled", "last_sync_time", "last_sync_status"]
		)
		
		return {
			"accounts": accounts
		}


@frappe.whitelist()
def test_service_principal_connection(service_principal_name):
	"""
	Test service principal credentials
	System Manager only
	
	Args:
		service_principal_name: Name of service principal
		
	Returns:
		dict: Connection test results
	"""
	if "System Manager" not in frappe.get_roles(frappe.session.user):
		frappe.throw(_("Only System Manager can test service principal connections"))
	
	result = test_connection(service_principal_name)
	return result


@frappe.whitelist()
def get_available_service_principals():
	"""
	Get list of enabled service principals
	For user to select when enabling email sync
	
	Returns:
		list: Available service principals
	"""
	service_principals = frappe.get_all(
		"M365 Email Service Principal Settings",
		filters={"enabled": 1},
		fields=["name", "service_principal_name", "tenant_name", "tenant_id"]
	)
	
	return service_principals


@frappe.whitelist()
def get_shared_mailboxes():
	"""
	Get all configured shared mailboxes
	System Manager sees all, others see none (unless they have Inbox User role for Communications)
	
	Returns:
		list: Shared mailboxes
	"""
	filters = {"account_type": "Shared Mailbox"}
	
	# Only System Manager can see shared mailbox configurations
	if "System Manager" not in frappe.get_roles(frappe.session.user):
		return []
	
	shared_mailboxes = frappe.get_all(
		"M365 Email Account",
		filters=filters,
		fields=["name", "account_name", "email_address", "enabled", "last_sync_time", "last_sync_status"]
	)
	
	return shared_mailboxes


@frappe.whitelist()
def get_available_folders(email_account_name):
	"""
	Get available mail folders for an email account
	
	Args:
		email_account_name: Name of M365 Email Account
		
	Returns:
		list: Available folders
	"""
	account = frappe.get_doc("M365 Email Account", email_account_name)
	
	# Check permissions
	if not user_can_configure_account(frappe.session.user, account):
		frappe.throw(_("You don't have permission to view this email account"))
	
	# Get access token
	access_token = get_access_token(account.service_principal)
	
	# Get folders from Graph API
	response = get_mail_folders(account.email_address, access_token)
	folders = response.get("value", [])
	
	# Format folder list
	folder_list = []
	for folder in folders:
		folder_list.append({
			"id": folder.get("id"),
			"displayName": folder.get("displayName"),
			"totalItemCount": folder.get("totalItemCount"),
			"unreadItemCount": folder.get("unreadItemCount")
		})
	
	return folder_list


@frappe.whitelist()
def update_folder_filters(email_account_name, folders):
	"""
	Update folder filters for an email account
	
	Args:
		email_account_name: Name of M365 Email Account
		folders: List of folder dicts with folder_name and sync_enabled
		
	Returns:
		dict: Success message
	"""
	import json
	
	account = frappe.get_doc("M365 Email Account", email_account_name)
	
	# Check permissions
	if not user_can_configure_account(frappe.session.user, account):
		frappe.throw(_("You don't have permission to configure this email account"))
	
	# Parse folders if string
	if isinstance(folders, str):
		folders = json.loads(folders)
	
	# Clear existing filters
	account.folder_filter = []
	
	# Add new filters
	for folder in folders:
		account.append("folder_filter", {
			"folder_name": folder.get("folder_name"),
			"sync_enabled": folder.get("sync_enabled", 1)
		})
	
	account.save()
	frappe.db.commit()
	
	return {
		"success": True,
		"message": _("Folder filters updated successfully")
	}

