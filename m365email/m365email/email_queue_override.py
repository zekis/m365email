# Copyright (c) 2025, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

import frappe
from frappe.email.doctype.email_queue.email_queue import EmailQueue


class M365EmailQueue(EmailQueue):
	"""
	Override Email Queue to send M365-marked emails via Graph API instead of SMTP
	"""

	def send(self, smtp_server_instance=None, **kwargs):
		"""
		Override send method to send M365 emails via Graph API

		Args:
			smtp_server_instance: SMTP server instance (for non-M365 emails)
			**kwargs: Additional parameters (e.g., force_send) to support different Frappe versions
		"""
		# Check if this email is marked for M365 sending
		if getattr(self, 'm365_send', 0) == 1:
			print(f"M365 Email: Sending '{self.name}' via M365 Graph API")

			# Import here to avoid circular imports
			from m365email.m365email.send import send_via_m365

			# Update status to Sending
			self.db_set("status", "Sending")
			frappe.db.commit()

			# Try to send via M365
			try:
				if send_via_m365(self):
					self.db_set("status", "Sent")
					print(f"M365 Email: Successfully sent '{self.name}'")
				else:
					self.db_set("status", "Error")
					self.db_set("error", "Failed to send via M365")
					print(f"M365 Email: Failed to send '{self.name}'")
			except Exception as e:
				self.db_set("status", "Error")
				self.db_set("error", str(e))
				frappe.log_error(
					title=f"M365 Email Send Error: {self.name}",
					message=str(e)
				)
				print(f"M365 Email: Exception while sending '{self.name}': {str(e)}")

			frappe.db.commit()
			return

		# For non-M365 emails, use standard SMTP sending
		super().send(smtp_server_instance=smtp_server_instance, **kwargs)

