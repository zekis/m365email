__version__ = "0.0.1"


def patch_email_account():
	"""
	Monkey-patch EmailAccount.find_outgoing to support M365 sending
	without requiring a default Email Account
	"""
	import frappe
	from frappe.email.doctype.email_account.email_account import EmailAccount

	# Store original method
	_original_find_outgoing = EmailAccount.find_outgoing

	@classmethod
	def patched_find_outgoing(cls, match_by_email=None, match_by_doctype=None, _raise_error=False):
		"""
		Patched version that returns a dummy account when M365 sending is available
		"""
		# Try original method first (without raising error)
		result = _original_find_outgoing.__func__(
			cls,
			match_by_email=match_by_email,
			match_by_doctype=match_by_doctype,
			_raise_error=False
		)

		if result:
			return result

		# If no email account found, check if M365 sending is available
		try:
			from m365email.m365email.send import can_send_via_m365

			if can_send_via_m365():
				# Return a minimal mock object that satisfies validation
				# This won't be used for actual sending - M365 will handle it
				class M365DummyAccount:
					def __init__(self):
						self.name = "M365 Virtual Account"
						self.email_id = match_by_email or "noreply@m365.local"
						self.footer = None
						self.enable_auto_reply = 0
						self.send_unsubscribe_message = 0
						self.track_email_status = 0
						self.enable_incoming = 0
						self.smtp_server = None
						self.use_tls = 0
						self.use_ssl = 0
						self.append_emails_to_sent_folder = 0
						self.always_use_account_email_id_as_sender = 0
						self.always_use_account_name_as_sender_name = 0
						self.add_signature = 0
						self.signature = None
						self.enable_outgoing = 1
						self.default_outgoing = 0
						self.smtp_port = None
						self.login_id = None
						self.password = None
						self.ascii_encode_password = 0
						self.always_bcc = None

					def is_exists_in_db(self):
						return False

					def get(self, key, default=None):
						return getattr(self, key, default)

				return M365DummyAccount()
		except Exception:
			pass

		# If M365 not available and _raise_error is True, raise the error
		if _raise_error:
			frappe.throw(
				frappe._("Please setup default Email Account from Settings > Email Account"),
				frappe.OutgoingEmailError,
			)

		return None

	# Apply the patch
	EmailAccount.find_outgoing = patched_find_outgoing


# Apply patches on app load
patch_email_account()
