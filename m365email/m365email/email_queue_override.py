# Copyright (c) 2025, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

import frappe
from frappe.email.doctype.email_queue.email_queue import EmailQueue


class M365EmailQueue(EmailQueue):
	"""
	Override Email Queue to skip SMTP sending for M365-marked emails
	"""

	def send(self, smtp_server_instance=None, **kwargs):
		"""
		Override send method to skip SMTP for M365 emails
		M365 emails are processed by the scheduled task instead

		Args:
			smtp_server_instance: SMTP server instance (for non-M365 emails)
			**kwargs: Additional parameters (e.g., force_send) to support different Frappe versions
		"""
		# Check if this email is marked for M365 sending
		if getattr(self, 'm365_send', 0) == 1:
			print(f"M365 Email: Skipping SMTP send for '{self.name}' - queued for M365 processing")
			# Don't call parent send() - the scheduled task will handle it
			return

		# For non-M365 emails, use standard SMTP sending
		super().send(smtp_server_instance=smtp_server_instance, **kwargs)

