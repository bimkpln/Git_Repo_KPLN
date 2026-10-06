# KPLN job candidate model checks

Read rules/review-policy.md and rules/criteria-contract.md before model review
or changing review behavior. This layer owns criteria, evidence, evaluation
and reports for job applicants' Revit models. KPLN_RevitMcpBridge remains the
shared transport/API adapter. Family generation rules belong to KPLN_RevitRules.

Keep skills thin: change procedures here, not in installed skill copies.
Do not invent hiring thresholds or promote observations into criteria without
a source and explicit applicability. Review the submitted model read-only.
Store candidate-specific evidence, paths and identifiers in ignored local/.
Use synthetic fixtures for any executable checks; version released rule changes.
