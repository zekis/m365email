# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Utility functions for M365 Email Integration
"""

import json
import frappe
from frappe import _
from datetime import datetime
from email.utils import parseaddr
from dateutil import parser as dateutil_parser


def parse_m365_datetime(datetime_string):
	"""
	Parse M365 datetime string to naive datetime for Frappe
	M365 returns ISO 8601 format with timezone (e.g., "2025-10-24T07:22:14Z")
	Frappe/MySQL expects naive datetime without timezone

	Args:
		datetime_string: ISO 8601 datetime string from M365

	Returns:
		datetime: Naive datetime object (timezone removed)
	"""
	if not datetime_string:
		return None

	# Parse the datetime string (handles timezone)
	dt = dateutil_parser.parse(datetime_string)

	# Convert to naive datetime (remove timezone info)
	# This keeps the UTC time but removes the timezone awareness
	return dt.replace(tzinfo=None)


def parse_email_address(email_string):
	"""
	Parse email address from string (handles "Name <email@domain.com>" format)

	Args:
		email_string: Email string to parse

	Returns:
		tuple: (name, email)
	"""
	if not email_string:
		return None, None

	name, email = parseaddr(email_string)
	return name or None, email or None


def parse_recipients(recipients_list):
	"""
	Parse list of recipients from M365 message format
	
	Args:
		recipients_list: List of recipient dicts from Graph API
		
	Returns:
		list: List of email addresses
	"""
	if not recipients_list:
		return []
	
	emails = []
	for recipient in recipients_list:
		email_address = recipient.get("emailAddress", {}).get("address")
		if email_address:
			emails.append(email_address)
	
	return emails


def should_sync_message(message, email_account):
	"""
	Determine if message should be synced based on filters
	
	Args:
		message: Message dict from Graph API
		email_account: M365 Email Account doc
		
	Returns:
		bool: True if message should be synced
	"""
	# Check sync_from_date filter
	if email_account.sync_from_date:
		received_datetime = message.get("receivedDateTime")
		if received_datetime:
			received_date = frappe.utils.get_datetime(received_datetime).date()
			if received_date < email_account.sync_from_date:
				return False
	
	return True


def user_can_configure_account(user, email_account):
	"""
	Check if user can configure email account settings
	
	Args:
		user: Frappe user
		email_account: M365 Email Account doc or name
		
	Returns:
		bool: True if user can configure
	"""
	if isinstance(email_account, str):
		email_account = frappe.get_doc("M365 Email Account", email_account)
	
	# System Manager can configure all
	if "System Manager" in frappe.get_roles(user):
		return True
	
	# For User Mailbox: user can configure their own
	if email_account.account_type == "User Mailbox" and email_account.user == user:
		return True
	
	# For Shared Mailbox: only System Manager
	return False


def get_communication_reference(message_data, email_account):
	"""
	Determine reference_doctype and reference_name for Communication
	Can implement logic to auto-link emails to doctypes (e.g., Support Ticket by email parsing)
	
	Args:
		message_data: Message dict from Graph API
		email_account: M365 Email Account doc
		
	Returns:
		tuple: (reference_doctype, reference_name) or (None, None)
	"""
	# TODO: Implement auto-linking logic
	# For example:
	# - Parse subject for ticket numbers
	# - Match sender email to Contact/Lead
	# - Link to specific doctypes based on rules
	
	return None, None


def format_email_body(body_content, content_type="html"):
	"""
	Format email body content for Communication doctype
	
	Args:
		body_content: Email body content
		content_type: Content type (html or text)
		
	Returns:
		str: Formatted body content
	"""
	if not body_content:
		return ""
	
	# For HTML content, return as-is
	if content_type.lower() == "html":
		return body_content
	
	# For text content, wrap in <pre> tags to preserve formatting
	return f"<pre>{body_content}</pre>"


def create_sync_log(email_account, sync_type="Delta Sync"):
	"""
	Create a new sync log entry
	
	Args:
		email_account: M365 Email Account name or doc
		sync_type: Type of sync (Full Sync, Delta Sync, Manual Sync)
		
	Returns:
		M365 Email Sync Log doc
	"""
	if isinstance(email_account, str):
		email_account_name = email_account
	else:
		email_account_name = email_account.name
	
	log = frappe.get_doc({
		"doctype": "M365 Email Sync Log",
		"email_account": email_account_name,
		"sync_type": sync_type,
		"status": "In Progress",
		"start_time": datetime.now()
	})
	log.insert(ignore_permissions=True)
	frappe.db.commit()
	
	return log


def update_sync_log(log, status, **kwargs):
	"""
	Update sync log with results
	
	Args:
		log: M365 Email Sync Log doc or name
		status: Status (Success, Failed, Partial Success)
		**kwargs: Additional fields to update
	"""
	if isinstance(log, str):
		log = frappe.get_doc("M365 Email Sync Log", log)
	
	log.status = status
	log.end_time = datetime.now()
	
	# Calculate duration
	if log.start_time and log.end_time:
		duration = (log.end_time - log.start_time).total_seconds()
		log.duration = duration
	
	# Update additional fields
	for key, value in kwargs.items():
		if hasattr(log, key):
			setattr(log, key, value)
	
	log.save(ignore_permissions=True)
	frappe.db.commit()


def get_or_create_contact(email_address, name=None):
	"""
	Get existing contact or create new one from email address
	
	Args:
		email_address: Email address
		name: Contact name (optional)
		
	Returns:
		str: Contact name
	"""
	if not email_address:
		return None
	
	# Check if contact exists with this email
	contact = frappe.db.get_value(
		"Contact Email",
		{"email_id": email_address},
		"parent"
	)
	
	if contact:
		return contact
	
	# Create new contact
	try:
		contact_doc = frappe.get_doc({
			"doctype": "Contact",
			"first_name": name or email_address.split("@")[0],
			"email_ids": [{
				"email_id": email_address,
				"is_primary": 1
			}]
		})
		contact_doc.insert(ignore_permissions=True)
		frappe.db.commit()
		
		return contact_doc.name
	except Exception as e:
		frappe.log_error(
			title=f"Failed to create contact for {email_address}",
			message=str(e)
		)
		return None


def sanitize_subject(subject):
	"""
	Sanitize email subject for use in Communication
	
	Args:
		subject: Email subject
		
	Returns:
		str: Sanitized subject
	"""
	if not subject:
		return _("(No Subject)")
	
	# Limit length
	max_length = 140
	if len(subject) > max_length:
		subject = subject[:max_length] + "..."
	
	return subject

