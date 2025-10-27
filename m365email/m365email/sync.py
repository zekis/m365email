# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Email synchronization module for M365 Email Integration
Core logic for syncing emails from M365 to Frappe Communications
"""

import json
import base64
import frappe
from frappe import _
from datetime import datetime
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import (
	get_messages_delta,
	get_message_details,
	get_message_attachments,
	download_attachment
)
from m365email.m365email.utils import (
	should_sync_message,
	parse_recipients,
	format_email_body,
	create_sync_log,
	update_sync_log,
	get_or_create_contact,
	sanitize_subject,
	get_communication_reference,
	parse_m365_datetime
)


def sync_email_account(email_account_name, folder_name=None):
	"""
	Main sync function for an email account (user or shared mailbox)
	Fetches emails from M365 and creates Communications
	Handles both initial sync and incremental delta sync
	
	Args:
		email_account_name: Name of M365 Email Account
		folder_name: Specific folder to sync (None = sync all enabled folders)
		
	Returns:
		dict: Sync results
	"""
	email_account = frappe.get_doc("M365 Email Account", email_account_name)
	
	if not email_account.enabled:
		return {"success": False, "message": "Email account is not enabled"}
	
	# Create sync log
	sync_log = create_sync_log(email_account, sync_type="Delta Sync")
	
	try:
		# Get access token
		access_token = get_access_token(email_account.service_principal)
		
		# Determine which folders to sync
		folders_to_sync = []
		if folder_name:
			# Sync specific folder
			folders_to_sync = [{"folder_name": folder_name}]
		elif email_account.folder_filter:
			# Sync enabled folders from filter
			folders_to_sync = [f for f in email_account.folder_filter if f.sync_enabled]
		else:
			# Default: sync Inbox
			folders_to_sync = [{"folder_name": "Inbox"}]
		
		total_fetched = 0
		total_created = 0
		total_updated = 0
		total_failed = 0
		
		# Sync each folder
		for folder in folders_to_sync:
			folder_name = folder.get("folder_name") or folder.folder_name
			result = sync_folder(email_account, folder_name, access_token, sync_log)
			
			total_fetched += result.get("fetched", 0)
			total_created += result.get("created", 0)
			total_updated += result.get("updated", 0)
			total_failed += result.get("failed", 0)
		
		# Update email account status
		email_account.db_set("last_sync_time", datetime.now(), update_modified=False)
		email_account.db_set("last_sync_status", "Success", update_modified=False)
		email_account.db_set("sync_error_message", None, update_modified=False)
		
		# Update sync log
		update_sync_log(
			sync_log,
			status="Success",
			messages_fetched=total_fetched,
			messages_created=total_created,
			messages_updated=total_updated,
			messages_failed=total_failed
		)
		
		frappe.db.commit()
		
		return {
			"success": True,
			"fetched": total_fetched,
			"created": total_created,
			"updated": total_updated,
			"failed": total_failed
		}
		
	except Exception as e:
		error_msg = str(e)
		frappe.log_error(
			title=f"M365 Email Sync Failed: {email_account_name}",
			message=error_msg
		)
		
		# Update email account status
		email_account.db_set("last_sync_status", "Failed", update_modified=False)
		email_account.db_set("sync_error_message", error_msg, update_modified=False)
		
		# Update sync log
		update_sync_log(
			sync_log,
			status="Failed",
			error_message=error_msg
		)
		
		frappe.db.commit()
		
		return {"success": False, "message": error_msg}


def sync_folder(email_account, folder_name, access_token, sync_log):
	"""
	Sync specific folder using delta query
	
	Args:
		email_account: M365 Email Account doc
		folder_name: Folder name to sync
		access_token: Access token
		sync_log: Sync log doc
		
	Returns:
		dict: Sync results for this folder
	"""
	# Get delta token for this folder
	delta_tokens = {}
	if email_account.delta_tokens:
		try:
			delta_tokens = json.loads(email_account.delta_tokens)
		except:
			delta_tokens = {}
	
	delta_token = delta_tokens.get(folder_name)
	
	# Fetch messages using delta query
	response = get_messages_delta(
		email_account.email_address,
		access_token,
		folder=folder_name,
		delta_token=delta_token
	)
	
	messages = response.get("value", [])
	fetched = len(messages)
	created = 0
	updated = 0
	failed = 0
	
	# Process each message
	for message in messages:
		try:
			if should_sync_message(message, email_account):
				result = create_communication_from_message(message, email_account, access_token)
				if result == "created":
					created += 1
				elif result == "updated":
					updated += 1
		except Exception as e:
			failed += 1
			frappe.log_error(
				title="M365 Email Sync: Failed to sync message",
				message=f"Message ID: {message.get('id')}\n\nError: {str(e)}"
			)
	
	# Save new delta token
	delta_link = response.get("@odata.deltaLink")
	if delta_link:
		delta_tokens[folder_name] = delta_link
		email_account.db_set("delta_tokens", json.dumps(delta_tokens), update_modified=False)
	
	# Update folder filter last sync time
	if email_account.folder_filter:
		for folder in email_account.folder_filter:
			if folder.folder_name == folder_name:
				folder.db_set("last_sync_time", datetime.now(), update_modified=False)
				break
	
	return {
		"fetched": fetched,
		"created": created,
		"updated": updated,
		"failed": failed
	}


def create_communication_from_message(message_data, email_account, access_token):
	"""
	Convert M365 message to Frappe Communication doctype
	Handles attachments, recipients, threading
	Sets owner to email_account.user (the user who configured the sync)
	
	Args:
		message_data: Message dict from Graph API
		email_account: M365 Email Account doc
		access_token: Access token
		
	Returns:
		str: "created", "updated", or "skipped"
	"""
	message_id = message_data.get("id")
	
	# Check if communication already exists
	existing = frappe.db.get_value(
		"Communication",
		{"m365_message_id": message_id}
	)
	
	if existing:
		return "skipped"
	
	# Parse message data
	subject = sanitize_subject(message_data.get("subject"))
	from_email = message_data.get("from", {}).get("emailAddress", {}).get("address")
	from_name = message_data.get("from", {}).get("emailAddress", {}).get("name")
	
	to_emails = parse_recipients(message_data.get("toRecipients", []))
	cc_emails = parse_recipients(message_data.get("ccRecipients", []))
	bcc_emails = parse_recipients(message_data.get("bccRecipients", []))
	
	body_content = message_data.get("body", {}).get("content", "")
	body_type = message_data.get("body", {}).get("contentType", "html")
	
	received_datetime = message_data.get("receivedDateTime")
	sent_datetime = message_data.get("sentDateTime")
	
	# Get reference doctype/name (for auto-linking)
	reference_doctype, reference_name = get_communication_reference(message_data, email_account)
	
	# Create or get contact
	contact = None
	if email_account.auto_create_contact and from_email:
		contact = get_or_create_contact(from_email, from_name)
	
	# Create Communication
	comm = frappe.get_doc({
		"doctype": "Communication",
		"communication_type": "Communication",
		"communication_medium": "Email",
		"sent_or_received": "Received",
		"subject": subject,
		"sender": from_email,
		"sender_full_name": from_name,
		"recipients": ", ".join(to_emails) if to_emails else None,
		"cc": ", ".join(cc_emails) if cc_emails else None,
		"bcc": ", ".join(bcc_emails) if bcc_emails else None,
		"content": format_email_body(body_content, body_type),
		"communication_date": parse_m365_datetime(received_datetime or sent_datetime),
		"reference_doctype": reference_doctype,
		"reference_name": reference_name,
		"m365_message_id": message_id,
		"m365_email_account": email_account.name,
		"owner": email_account.user  # Set owner to the user who configured the sync
	})
	
	comm.insert(ignore_permissions=True)
	
	# Handle attachments
	if email_account.sync_attachments and message_data.get("hasAttachments"):
		sync_attachments(email_account, message_id, comm, access_token)
	
	frappe.db.commit()
	
	return "created"


def sync_attachments(email_account, message_id, communication, access_token):
	"""
	Download and attach email attachments to Communication
	
	Args:
		email_account: M365 Email Account doc
		message_id: M365 message ID
		communication: Communication doc
		access_token: Access token
	"""
	try:
		attachments_response = get_message_attachments(
			email_account.email_address,
			message_id,
			access_token
		)
		
		attachments = attachments_response.get("value", [])
		max_size_bytes = (email_account.max_attachment_size or 10) * 1024 * 1024
		
		for attachment in attachments:
			# Check attachment size
			size = attachment.get("size", 0)
			if size > max_size_bytes:
				continue
			
			# Only handle file attachments (not item attachments)
			if attachment.get("@odata.type") != "#microsoft.graph.fileAttachment":
				continue
			
			# Get attachment content
			content_bytes = attachment.get("contentBytes")
			if not content_bytes:
				continue
			
			# Decode base64 content
			file_content = base64.b64decode(content_bytes)
			file_name = attachment.get("name", "attachment")
			
			# Create Frappe File
			file_doc = frappe.get_doc({
				"doctype": "File",
				"file_name": file_name,
				"attached_to_doctype": "Communication",
				"attached_to_name": communication.name,
				"content": file_content,
				"is_private": 1
			})
			file_doc.save(ignore_permissions=True)
			
	except Exception as e:
		frappe.log_error(
			title="M365 Email Sync: Failed to sync attachments",
			message=f"Message ID: {message_id}\n\nError: {str(e)}"
		)

