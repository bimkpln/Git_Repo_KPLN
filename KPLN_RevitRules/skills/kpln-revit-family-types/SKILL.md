---
name: kpln-revit-family-types
description: Create, populate or update Revit family types from manufacturer selections and catalogs through KPLN_RevitMcpBridge, and maintain the corresponding KPLN rules. Use for генерация и заполнение типоразмеров семейств; candidate model assessment belongs to kpln-job-candidate-model-check.
---

# KPLN Revit family types

Resolve KPLN_RevitRules from process KPLN_REVIT_RULES_ROOT, then rules-root.txt
beside this file, then the User environment variable. Read VERSION and
rules/family-types-workflow.md; follow its routing to the relevant rules.
Use the existing KPLN_RevitMcpBridge tools for model access.

Keep mappings, calculations, write/verification procedures and learned decisions
in that checkout, not in this skill. If it is unavailable, report the missing
path; requested inspection may proceed without inventing engineering rules.
