# Working On Clash Knowledge

Read `rules/clash-policy.md` and `knowledge/proposals/review-lessons.md` before
changing classification. This repository owns engineering decisions; the
Navisworks bridge owns property extraction, geometry and status writes.

Edit shared policy and code here, not in a user's installed skill directory.
Use the local checkout selected by KPLN_NAVIS_RULES_ROOT. Do not automatically
pull, reset, commit or push while reviewing a model. Preserve local edits.

Keep project thresholds and user overrides in each analysis record. Promote a
project decision to shared policy only when its applicability is established.
Uncertain cases remain available to skill reasoning. Record any human or agent
override separately with its evidence and scope; do not disguise it as a
deterministic classifier result.

Run `python -B -m unittest discover -s tests -v` after classifier changes.
Use anonymous synthetic cases, not production model paths or clash IDs.
Update CHANGELOG.md and VERSION when releasing shared behavior changes.
