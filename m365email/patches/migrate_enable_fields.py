# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Migration patch to convert 'enabled' and 'use_for_sending' fields
to new 'enable_incoming', 'enable_outgoing', and 'default_outgoing' fields
"""

import frappe


def execute():
	"""
	Migrate M365 Email Account fields:
	- enabled=1 → enable_incoming=1
	- use_for_sending=1 → enable_outgoing=1, default_outgoing=1
	"""
	
	# Check if old fields exist
	if not frappe.db.has_column("M365 Email Account", "enabled"):
		print("✅ Migration already completed - old fields don't exist")
		return
	
	print("Starting M365 Email Account field migration...")
	
	# Get all M365 Email Accounts
	accounts = frappe.get_all(
		"M365 Email Account",
		fields=["name", "enabled", "use_for_sending"]
	)
	
	if not accounts:
		print("No M365 Email Accounts found")
		return
	
	# Track which account should be default_outgoing
	default_outgoing_account = None
	
	for account in accounts:
		doc = frappe.get_doc("M365 Email Account", account.name)
		
		# Migrate enabled → enable_incoming
		if account.get("enabled"):
			doc.enable_incoming = 1
			print(f"  {account.name}: enabled → enable_incoming")
		
		# Migrate use_for_sending → enable_outgoing + default_outgoing
		if account.get("use_for_sending"):
			doc.enable_outgoing = 1
			print(f"  {account.name}: use_for_sending → enable_outgoing")
			
			# First account with use_for_sending becomes default_outgoing
			if not default_outgoing_account:
				doc.default_outgoing = 1
				default_outgoing_account = account.name
				print(f"  {account.name}: Set as default_outgoing")
		
		# Save without triggering validations (in case they reference old fields)
		doc.flags.ignore_validate = True
		doc.save()
	
	frappe.db.commit()
	
	print(f"\n✅ Migrated {len(accounts)} M365 Email Account(s)")
	if default_outgoing_account:
		print(f"✅ Default outgoing account: {default_outgoing_account}")
	
	print("\nNOTE: Old fields 'enabled' and 'use_for_sending' can be removed manually if needed")

