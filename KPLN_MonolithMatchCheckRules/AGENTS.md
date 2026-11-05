# Monolith AR/KR matching
Read rules/common/local-first.md. Work on local disk; publish verified stage
checkpoints to the shared folder. Read development/README.md for the available
research implementation. Keep versions and change selection in service logs;
ask users only for genuinely missing inputs or decisions affecting their task.
Read rules/common/process.md, the selected rules/pd or rules/rd profile,
and rules/common/geometry.md before performing checks or changing behavior.
Read rules/reports before generating outputs. Keep SKILL.md a thin entrypoint.
Use one immutable rules snapshot per run and one canonical report hierarchy.
Project data, registries, machine settings and outputs belong in ignored local/
or explicitly configured external data storage, never in reusable rules.
This release includes admission/history utilities, NOT a portable full NWD
geometry engine. See rules/common/readiness.md; do not invent successful checks.
Do not generalize project-specific IDs, material dictionaries or repairs.
Validate changes with python -B -m unittest discover -s tests -v and update VERSION/CHANGELOG.
No model edits, plugin code changes or automatic report publication are implied.
For employee adaptations, follow rules/common/contributions.md: a separate working
copy and review package, no shared release overwrite. Only explicitly commissioned
integration updates the common VERSION/CHANGELOG after review and validation.
