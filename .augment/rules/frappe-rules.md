---
type: "always_apply"
---

Logging rules in frappe - must only ever use frappe.log_error(title,message) for errors. Otherwise its recommended to create a doctype especially for keeping debug, info, warning logs.
