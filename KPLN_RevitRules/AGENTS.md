# KPLN Revit knowledge

Read rules/revit-policy.md and knowledge/proposals/review-lessons.md before
changing shared behavior. This checkout owns engineering decisions; the Revit
bridge owns API extraction, normalized geometry and explicit transactional writes.
Do not add project thresholds or engineering verdicts to the bridge.

Edit shared knowledge here, not in an installed skill copy. Preserve local edits.
Do not automatically pull, reset, commit or push during model work. Keep model
paths, element IDs and project-specific overrides in ignored local/ records.
Generalize a reviewed decision only when evidence establishes its scope.

Use synthetic data for tests. Run python -B -m unittest discover -s tests -v
after changing src/. Update VERSION and CHANGELOG.md for released behavior.
