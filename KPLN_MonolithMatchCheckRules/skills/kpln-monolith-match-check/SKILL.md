---
name: kpln-monolith-match-check
description: Проверка совпадения монолита АР/КР в NWD на стадиях ПД и РД, контроль исправлений, группировка и отчёты для KPLN и координатора. Также поддержка правил этой проверки; не обычная проверка пересечений инженерных систем.
---

# KPLN monolith match check

Read rules/common/local-first.md for local execution, available development
checkpoints, quiet reuse of employee changes, and network-failure recovery.
The user provides check inputs and receives reports; handle implementation
choices and provenance internally. Never hide missing coverage or invent results.

Resolve the shared KPLN_MonolithMatchCheckRules directory from process variable
KPLN_MONOLITH_MATCH_RULES_ROOT, then rules-root.txt beside this file, then
the repository root two directories above this skill. Do not use a different
project's local rules or cached history if this directory is unavailable.

Read VERSION, rules/common/process.md, rules/common/readiness.md and
rules/common/geometry.md. Select rules/pd/profile.md or rules/rd/profile.md
from the explicit documentation stage. Read rules/common/history.md for
registry selection and rules/reports/plugin.md and coordinator.md for outputs.
Before changing calculation code, reports or rules during a check, read
rules/common/contributions.md. Work in a separate copy and save the changes,
baseline, evidence and test results as a pending_review contribution. Do not
overwrite the shared release or automatically adopt colleagues' pending changes.
Integration into the shared release requires a separate explicit instruction.

Before work, obtain project, stage, building/section, NWD, output folder and
PD drawing scope if applicable. Reuse established identifiers from shared state.
For a bare request to start, show the missing fields from examples/start-request.txt
as one concise intake form. Do not ask again for values already supplied or verified.
Unfilled template placeholders are missing input, not project names or paths.
scripts/new-check.ps1 opens a fresh chat with the template in the composer;
it does not send a user message or an assistant response automatically.
Use scripts/prepare-run.py to snapshot this release and input hash; continue
from that snapshot throughout the run. Ask only for genuinely missing inputs.
The current package does not implement a portable complete NWD geometry engine.
Check its integration before promising final reports. Unverified geometry is
never a confirmed issue or a complete clean result. Do not hide unresolved work.

Follow the user's agreed scope and latest explicit project overrides, recording
them in the run. Do not change global company policy from a single project example.
Use existing Navisworks access if available; do not infer geometry from screenshots
or ordinary clash rows. Shared rule changes do not require editing installed copies.
