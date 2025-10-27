# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document


class M365EmailAccount(Document):
	"""Configure email sync for user mailboxes OR shared mailboxes"""

	def validate(self):
		"""Validate email account settings"""
		# User field is always mandatory
		if not self.user:
			frappe.throw(_("User field is required"))

		# For User Mailbox: validate email matches user
		if self.account_type == "User Mailbox":
			user_doc = frappe.get_doc("User", self.user)
			user_emails = [user_doc.email]

			# Add any additional emails from User Emails child table
			if hasattr(user_doc, "user_emails"):
				user_emails.extend([d.email_id for d in user_doc.user_emails])

			if self.email_address not in user_emails:
				frappe.msgprint(
					_("Email address {0} does not match user {1}'s email addresses").format(
						self.email_address, self.user
					),
					indicator="orange",
					alert=True
				)

		# Check for duplicate email address per service principal
		if self.service_principal and self.email_address:
			duplicate = frappe.db.exists(
				"M365 Email Account",
				{
					"name": ["!=", self.name],
					"service_principal": self.service_principal,
					"email_address": self.email_address
				}
			)
			if duplicate:
				frappe.throw(
					_("Email account for {0} already exists for this service principal").format(
						self.email_address
					)
				)

		# Ensure only one account is marked for sending
		if self.use_for_sending:
			existing_sending_account = frappe.db.get_value(
				"M365 Email Account",
				{
					"name": ["!=", self.name],
					"use_for_sending": 1
				}
			)
			if existing_sending_account:
				frappe.throw(
					_("Account {0} is already marked for sending. Only one account can be used for sending.").format(
						existing_sending_account
					)
				)


def has_permission(doc, ptype, user):
	"""Custom permission check - module level function"""
	# System Manager can access all
	if "System Manager" in frappe.get_roles(user):
		return True

	# For User Mailbox: user can only access their own
	if doc.account_type == "User Mailbox" and doc.user == user:
		return True

	# For Shared Mailbox: only System Manager can access
	return False

