# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Email sending module for M365 Email Integration
Handles sending emails via Microsoft Graph API instead of SMTP
"""

import frappe
from frappe import _
import base64
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import send_email_as_user


def get_sending_account():
	"""
	Get the M365 Email Account marked for sending
	
	Returns:
		M365EmailAccount: The account to use for sending, or None
	"""
	account_name = frappe.db.get_value(
		"M365 Email Account",
		{"use_for_sending": 1, "enabled": 1}
	)
	
	if not account_name:
		return None
	
	return frappe.get_doc("M365 Email Account", account_name)


def can_send_via_m365():
	"""
	Check if M365 sending is available
	
	Returns:
		bool: True if M365 sending is configured
	"""
	return get_sending_account() is not None


def intercept_email_queue(doc, method=None):
	"""
	Hook: Email Queue before_insert
	Check if we should send via M365 instead of SMTP
	
	Args:
		doc: Email Queue document
		method: Hook method name (unused)
	"""
	# Check if M365 sending is available
	sending_account = get_sending_account()
	
	if not sending_account:
		# No M365 sending configured, use default SMTP
		return
	
	# Mark this email for M365 sending
	doc.m365_send = 1
	doc.m365_account = sending_account.name
	
	print(f"M365 Email: Marked email '{doc.name}' for M365 sending via {sending_account.account_name}")


def send_via_m365(email_queue_doc):
	"""
	Send an email via M365 Graph API

	Args:
		email_queue_doc: Email Queue document

	Returns:
		bool: True if sent successfully
	"""
	try:
		import email
		from email import policy
		from email.utils import parseaddr

		# Get the sending account
		sending_account = frappe.get_doc("M365 Email Account", email_queue_doc.m365_account)

		# Get access token
		access_token = get_access_token(sending_account.service_principal)

		if not access_token:
			frappe.log_error(
				title="M365 Email Send Failed: No Access Token",
				message=f"Could not get access token for {sending_account.service_principal}"
			)
			return False

		# Parse the MIME message to extract subject and body
		msg = email.message_from_string(email_queue_doc.message, policy=policy.default)

		# Extract subject
		subject = msg.get('Subject', 'No Subject')

		# Extract HTML body
		body = None
		if msg.is_multipart():
			for part in msg.walk():
				content_type = part.get_content_type()
				if content_type == 'text/html':
					body = part.get_content()
					break
				elif content_type == 'text/plain' and not body:
					body = part.get_content()
		else:
			body = msg.get_content()

		if not body:
			body = "No content"

		# Append footer if configured
		if sending_account.footer:
			# Add footer with proper HTML formatting
			footer_html = f"""
<div class="m365-email-footer" style="margin-top: 20px; padding-top: 10px; border-top: 1px solid #e0e0e0;">
{sending_account.footer}
</div>
"""
			# Append footer to body
			body = body + footer_html

		# Parse recipients - Email Queue stores recipients as child table
		recipients = []
		if hasattr(email_queue_doc, 'recipients') and email_queue_doc.recipients:
			if isinstance(email_queue_doc.recipients, list):
				# It's a child table - extract email addresses
				recipients = [r.recipient for r in email_queue_doc.recipients if r.recipient]
			else:
				# It's a string - split by comma
				recipients = [r.strip() for r in email_queue_doc.recipients.split(",") if r.strip()]

		# Parse CC - stored as comma-separated string
		cc = None
		if hasattr(email_queue_doc, 'show_as_cc') and email_queue_doc.show_as_cc:
			if isinstance(email_queue_doc.show_as_cc, str):
				cc = [r.strip() for r in email_queue_doc.show_as_cc.split(",") if r.strip()]

		# BCC is not stored in Email Queue
		bcc = None
		
		# Get attachments - Email Queue stores attachments as JSON string
		attachments = None
		if hasattr(email_queue_doc, 'attachments') and email_queue_doc.attachments:
			try:
				import json
				# Parse attachments JSON
				attachments_data = email_queue_doc.attachments
				if isinstance(attachments_data, str):
					attachments_data = json.loads(attachments_data)

				if attachments_data:
					attachments = []
					for attachment in attachments_data:
						try:
							# Get file identifier - could be file_url (File name) or fid
							file_identifier = attachment.get('file_url') or attachment.get('fid')
							if not file_identifier:
								continue

							# Try to get File document - file_url might be the File name or actual URL
							file_doc = None
							try:
								# First try as File name
								file_doc = frappe.get_doc("File", file_identifier)
							except frappe.DoesNotExistError:
								# Try to find by file_url field
								file_list = frappe.get_all("File", filters={"file_url": file_identifier}, limit=1)
								if file_list:
									file_doc = frappe.get_doc("File", file_list[0].name)

							if not file_doc:
								frappe.log_error(
									title="M365 Email: Attachment Not Found",
									message=f"Email: {email_queue_doc.name}\nFile identifier: {file_identifier}\nAttachment: {attachment}"
								)
								continue

							# Get file content
							file_content = file_doc.get_content()

							# Encode to base64
							base64_content = base64.b64encode(file_content).decode('utf-8')

							attachments.append({
								"name": attachment.get('file_name') or file_doc.file_name,
								"content": base64_content
							})
						except Exception as attach_error:
							frappe.log_error(
								title="M365 Email: Single Attachment Failed",
								message=f"Email: {email_queue_doc.name}\nAttachment: {attachment}\nError: {str(attach_error)}"
							)
							# Continue with other attachments
							continue
			except Exception as e:
				frappe.log_error(
					title="M365 Email: Attachment Processing Failed",
					message=f"Email: {email_queue_doc.name}\nError: {str(e)}\nTraceback: {frappe.get_traceback()}"
				)
				# Continue without attachments rather than failing the whole email
				attachments = None

		# Determine sender email
		sender_email = getattr(email_queue_doc, 'sender', None) or frappe.session.user

		# Parse email address from "Name <email>" format
		if sender_email:
			sender_name, sender_email = parseaddr(sender_email)

		# Check if we should always use the account email as sender
		if sending_account.always_use_account_email_as_sender:
			sender_email = sending_account.email_address

		# Validate sender email
		if not sender_email or '@' not in sender_email:
			frappe.log_error(
				title="M365 Email Send Failed: Invalid Sender",
				message=f"Email: {email_queue_doc.name}\nSender: {sender_email}"
			)
			return False

		# Send email via Graph API
		result = send_email_as_user(
			sender_email=sender_email,
			recipients=recipients,
			subject=subject,
			body=body,
			access_token=access_token,
			cc=cc,
			bcc=bcc,
			attachments=attachments,
			is_html=True
		)
		
		if result.get("success"):
			print(f"M365 Email: Successfully sent email '{email_queue_doc.name}' from {sender_email}")
			
			# Update Communication if linked
			if email_queue_doc.communication:
				comm = frappe.get_doc("Communication", email_queue_doc.communication)
				comm.db_set("sent_or_received", "Sent")
				comm.db_set("delivery_status", "Sent")
			
			return True
		else:
			error_msg = result.get("message", "Unknown error")
			frappe.log_error(
				title="M365 Email Send Failed",
				message=f"Email: {email_queue_doc.name}\nSender: {sender_email}\nError: {error_msg}"
			)
			return False
			
	except Exception as e:
		frappe.log_error(
			title="M365 Email Send Exception",
			message=f"Email: {email_queue_doc.name}\n\nError: {str(e)}"
		)
		return False


def process_email_queue_m365():
	"""
	Process Email Queue items marked for M365 sending
	This can be called from a scheduled task or manually
	"""
	# Get all pending emails marked for M365 sending
	email_queue_list = frappe.get_all(
		"Email Queue",
		filters={
			"status": ["in", ["Not Sent", "Sending"]],
			"m365_send": 1
		},
		limit=100
	)
	
	sent_count = 0
	failed_count = 0
	
	for email_queue in email_queue_list:
		doc = frappe.get_doc("Email Queue", email_queue.name)
		
		# Update status to Sending
		doc.db_set("status", "Sending")
		frappe.db.commit()
		
		# Try to send
		if send_via_m365(doc):
			doc.db_set("status", "Sent")
			sent_count += 1
		else:
			doc.db_set("status", "Error")
			doc.db_set("error", "Failed to send via M365")
			failed_count += 1
		
		frappe.db.commit()
	
	if sent_count > 0 or failed_count > 0:
		print(f"M365 Email Queue: Sent {sent_count}, Failed {failed_count}")
	
	return {
		"sent": sent_count,
		"failed": failed_count
	}

