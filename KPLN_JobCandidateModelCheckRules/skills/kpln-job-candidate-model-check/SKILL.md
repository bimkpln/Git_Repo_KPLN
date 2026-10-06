---
name: kpln-job-candidate-model-check
description: Review Revit models submitted by job candidates against the hiring test assignment and supplied criteria, or maintain those review rules. Use for проверка моделей кандидатов на работу and тестовое задание соискателя; family type generation belongs to kpln-revit-family-types.
---

# KPLN job candidate model check

Resolve KPLN_JobCandidateModelCheckRules from process
KPLN_JOB_CANDIDATE_MODEL_CHECK_RULES_ROOT, then rules-root.txt beside this file,
then the User environment variable. Read VERSION, rules/review-policy.md and
rules/criteria-contract.md. Use the existing KPLN_RevitMcpBridge for model data.

Keep criteria, scoring, evidence procedures and reports in that rules layer.
Review the submitted model read-only. If the checkout is unavailable, report
the missing path; requested inspection may proceed without invented criteria.
