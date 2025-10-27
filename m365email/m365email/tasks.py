# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Scheduled tasks for M365 Email Integration
"""

import frappe
from frappe import _
from m365email.m365email.sync import sync_email_account
from m365email.m365email.auth import refresh_token, test_connection


def sync_all_email_accounts():
	"""
	Sync all email accounts with incoming enabled
	Scheduled to run every 5 minutes
	"""
	print("M365 Email: Starting scheduled sync for all accounts with incoming enabled")

	# Get all accounts with incoming enabled
	accounts = frappe.get_all(
		"M365 Email Account",
		filters={"enable_incoming": 1},
		fields=["name", "account_name", "email_address", "account_type"]
	)

	if not accounts:
		print("M365 Email: No accounts with incoming enabled found")
		return

	print(f"M365 Email: Found {len(accounts)} account(s) with incoming enabled")

	success_count = 0
	failed_count = 0

	for account in accounts:
		try:
			print(f"M365 Email: Syncing {account.account_name} ({account.email_address})")
			result = sync_email_account(account.name)

			if result.get("success"):
				success_count += 1
				print(
					f"M365 Email: Successfully synced {account.account_name} - "
					f"Fetched: {result.get('fetched', 0)}, "
					f"Created: {result.get('created', 0)}, "
					f"Updated: {result.get('updated', 0)}"
				)
			else:
				failed_count += 1
				print(
					f"M365 Email: Failed to sync {account.account_name} - "
					f"Error: {result.get('message')}"
				)
		except Exception as e:
			failed_count += 1
			print(f"M365 Email: Exception syncing {account.account_name}: {str(e)}")
			frappe.log_error(
				title="M365 Email Sync Failed",
				message=f"Account: {account.account_name}\nEmail: {account.email_address}\n\nError: {str(e)}"
			)

	print(
		f"M365 Email: Scheduled sync completed - "
		f"Success: {success_count}, Failed: {failed_count}"
	)


def refresh_all_tokens():
	"""
	Refresh tokens for all enabled service principals
	Scheduled to run hourly
	"""
	print("M365 Email: Starting token refresh for all service principals")

	# Get all enabled service principals
	service_principals = frappe.get_all(
		"M365 Email Service Principal Settings",
		filters={"enabled": 1},
		fields=["name", "service_principal_name"]
	)

	if not service_principals:
		print("M365 Email: No enabled service principals found")
		return

	print(f"M365 Email: Found {len(service_principals)} enabled service principal(s)")

	success_count = 0
	failed_count = 0

	for sp in service_principals:
		try:
			print(f"M365 Email: Refreshing token for {sp.service_principal_name}")
			success = refresh_token(sp.name)

			if success:
				success_count += 1
				print(f"M365 Email: Successfully refreshed token for {sp.service_principal_name}")
			else:
				failed_count += 1
				print(f"M365 Email: Failed to refresh token for {sp.service_principal_name}")
		except Exception as e:
			failed_count += 1
			print(f"M365 Email: Exception refreshing token for {sp.service_principal_name}: {str(e)}")
			frappe.log_error(
				title="M365 Token Refresh Failed",
				message=f"Service Principal: {sp.service_principal_name}\n\nError: {str(e)}"
			)

	print(
		f"M365 Email: Token refresh completed - "
		f"Success: {success_count}, Failed: {failed_count}"
	)


def cleanup_old_logs():
	"""
	Cleanup old sync logs (older than 30 days)
	Scheduled to run daily
	"""
	print("M365 Email: Starting cleanup of old sync logs")

	from frappe.utils import add_days, now_datetime

	# Delete logs older than 30 days
	cutoff_date = add_days(now_datetime(), -30)

	old_logs = frappe.get_all(
		"M365 Email Sync Log",
		filters={
			"creation": ["<", cutoff_date]
		},
		pluck="name"
	)

	if not old_logs:
		print("M365 Email: No old logs to cleanup")
		return

	print(f"M365 Email: Deleting {len(old_logs)} old log(s)")

	for log_name in old_logs:
		try:
			frappe.delete_doc("M365 Email Sync Log", log_name, ignore_permissions=True)
		except Exception as e:
			print(f"M365 Email: Failed to delete log {log_name}: {str(e)}")

	frappe.db.commit()

	print("M365 Email: Cleanup completed")


def validate_service_principals():
	"""
	Validate all service principal credentials
	Scheduled to run daily
	"""
	print("M365 Email: Starting service principal validation")

	# Get all enabled service principals
	service_principals = frappe.get_all(
		"M365 Email Service Principal Settings",
		filters={"enabled": 1},
		fields=["name", "service_principal_name"]
	)

	if not service_principals:
		print("M365 Email: No enabled service principals to validate")
		return

	print(f"M365 Email: Validating {len(service_principals)} service principal(s)")

	valid_count = 0
	invalid_count = 0

	for sp in service_principals:
		try:
			print(f"M365 Email: Validating {sp.service_principal_name}")
			result = test_connection(sp.name)

			if result.get("success"):
				valid_count += 1
				print(f"M365 Email: {sp.service_principal_name} is valid")
			else:
				invalid_count += 1
				print(
					f"M365 Email: {sp.service_principal_name} validation failed - "
					f"Error: {result.get('message')}"
				)

				# Log error for admin attention
				frappe.log_error(
					title="M365 Service Principal Invalid",
					message=f"Service Principal: {sp.service_principal_name}\n\nError: {result.get('message')}"
				)
		except Exception as e:
			invalid_count += 1
			print(f"M365 Email: Exception validating {sp.service_principal_name}: {str(e)}")
			frappe.log_error(
				title="M365 Service Principal Validation Failed",
				message=f"Service Principal: {sp.service_principal_name}\n\nError: {str(e)}"
			)

	print(
		f"M365 Email: Validation completed - "
		f"Valid: {valid_count}, Invalid: {invalid_count}"
	)

