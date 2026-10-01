---
name: kpln-revit-bridge
description: Use for reading or editing a live Revit model through KPLN_RevitMcpBridge, inspecting elements, families and parameters, or maintaining shared KPLN Revit rules. Not for ordinary C# add-in development without a live-model task.
---

# KPLN Revit bridge

Find the shared knowledge checkout through KPLN_REVIT_RULES_ROOT in the process
environment, then rules-root.txt beside this installed SKILL.md, then the User
environment variable. Read VERSION, rules/revit-policy.md and
knowledge/proposals/review-lessons.md there before model work. If the checkout
is unavailable, do not invent its engineering rules; report the missing path.
Explicitly requested data inspection can still proceed without engineering judgments.

Use list_revit_sessions and get_revit_context to identify the intended Revit
instance and document. Reuse the returned session_id/document_id in subsequent
calls. Use UniqueId for elements and parameter IDs from returned data.
get_revit_elements reports Double parameters in Revit internal units; geometry
is in millimetres. A bounding box is not exact solid geometry.

Plan parameter writes with dry_run=True using expected_value and
expected_revision from fresh reads. Apply the reviewed plan with dry_run=False
within the user's authorized scope. A type parameter affects all instances.
Re-read values after success. After a timeout, use get_revit_operation; never
blindly repeat a write. The bridge does not save or synchronize documents.

Engineering thresholds and decisions belong in the shared checkout. Keep
project-specific evidence and overrides in ignored local/ analysis records;
src/create_record.py records knowledge provenance. Edit shared sources rather
than copying policy into this skill. Do not pull or publish automatically.
