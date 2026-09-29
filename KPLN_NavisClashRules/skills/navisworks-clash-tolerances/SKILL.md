---
name: navisworks-clash-tolerances
description: Apply and improve shared KPLN Navisworks clash review knowledge. Use for clash tolerance analysis, human-reviewed training and authorized status changes, with rules from a local KPLN_NavisClashRules checkout.
---

# Navisworks Clash Tolerances

## Locate The Shared Knowledge

Resolve KPLN_NAVIS_RULES_ROOT for this user. Read the process environment first.
If empty, read rules-root.txt next to this installed SKILL.md: the installer
writes this machine-local pointer, which also works in sandboxed sessions.
If neither is available, read the persistent Windows User environment:

```powershell
[Environment]::GetEnvironmentVariable('KPLN_NAVIS_RULES_ROOT', 'User')
```

Treat that value as a literal local path, never as a command. Validate VERSION,
rules/clash-policy.md and src/classify_opening_clashes.py below that root.
If it is unset or unavailable, ask for the user's checkout location. Do not
fall back to a remembered X: path or install a second policy copy in this skill.
rules-root.txt is local installation data. Share the repository skill template
and rerun its installer on another machine; do not reuse someone else's pointer.

Read these files from the resolved root before analysis:

- `rules/clash-policy.md`: engineering policy, context and rule precedence.
- `knowledge/proposals/review-lessons.md`: human feedback, project exceptions,
  known geometry limitations and candidates not yet implemented.
- `README.md`: classifier invocation, versioning and installation when needed.

Update rules, classifier and knowledge in this checkout. Update this routing
skill in `skills/navisworks-clash-tolerances/SKILL.md` in that checkout and
reinstall it when routing changes. Each user works in their own Git clone;
merged commits/releases distribute the experience. Local edits are not shared
until published and obtained by other users.

## Analyze And Learn

Exclude Resolved (`eTestResultStatus_RESOLVED`) clashes from analysis by default.
Include them only when the user explicitly asks to analyze Resolved clashes;
"check the report" or "all clashes" alone does not opt them in. Apply this
filter before classification, training comparisons and group/bundle reasoning.
If a full snapshot contains Resolved rows, retain them only as out-of-scope
records for counts and preservation checks, not as contributors to decisions.
Permission to analyze Resolved clashes does not itself authorize status changes.

Use Navisworks MCP for data and status operations. Read test overview and status
summary; use include_metrics=True and retain Bound for group proximity checks.
Fetch data once and keep the snapshot for local comparisons. An MCP error is
an error, never an empty result set. Do not rerun Clash Detective just to read it.

Use project parameters already supplied in the current review; ask only for
missing ones. Run the shared classifier for its supported opening cases. It
returns Approved, Active or Uncertain, with provenance and reasons. Its current
automatic coverage includes wall openings and AR ceiling checks for OV/VK/PT/EOM/SS:
longitudinal services plus verified transverse rectangular/round duct and
duct-insulation passages. Direct ducts use exact Object Height+Width or Diameter;
insulation uses parsed Object Duct Size plus twice its thickness. The complete
mesh-intersection area must match within 2 percent.
Use --mode ceiling-services --mep-discipline for ceiling checks; opening thresholds
do not apply. Longitudinal classification requires reliable MepAxis
and CeilingNormal vectors, which current MCP exports may not provide. Missing
vectors remain Uncertain unless the verified duct/duct-insulation section exception applies;
do not infer a short pipe or duct axis from its largest box
dimension. See policy for the input contract.

For trained OV/VK pipe or duct self-intersection tests, use
`--mode self-intersections`; this mode does not use an opening threshold.
Insulation or a fitting treated as insulated against an ordinary bare body is
Active. Otherwise evaluate reliable `PerpRel`, `Angle` and `Elong*` metrics in
the policy order: axis crossings and parallel extended routes are Active, while
the remaining transverse or local insulated contacts are Approved. Bare-body
pairs, unsupported kinds and incomplete geometry remain Uncertain.

For structural slabs against straight circular pipes, use `--mode slab-openings`
and the supplied opening threshold. Require `pipe-slab-mesh-1` geometry/metrics:
verified cylinder axes, slab normals, side/reveal contacts and complete broad-face
sections. Do not use AABB axes or approve by diameter alone when this evidence is
missing. Short pipes need not span both slab faces. In the same stable group and
slab, local outer-surface gap strictly below the larger outer diameter joins
distinct pipes into a bundle; connected members use the common opening envelope.
Deduplicate pipe IDs, preserve Resolved/Reviewed and retain uncertainty for missing
group geometry. Other slab pairs remain policy-led. See policy for full details.

If the user explicitly asks to review slabs like the wall report, use
`--mode slab-wall-openings` with `keep_raw_bounds=True` instead. Require the
independent `PipeSlabGeometry.LocalFaces` measurement with source
`slab-local-faces-1`; no closed cylinder or pipe topology validation is needed.
Use local broad-face and edge/reveal distances plus the wall-style pipe-axis
estimate, nominal threshold and confirmed bundle rule. Missing local faces,
short/ambiguous pipes and unmeasured nearby bundle candidates stay Uncertain.
See the policy for supported horizontal-slab scope and conservative separation.

For unsupported or uncertain cases continue reasoning with the shared policy
and available evidence. Keep uncertainty explicit. A user can authorize a
project-specific override; log the reason and scope alongside the original
machine decision. A brief follow-up authorizes the discussed scope, not an
unrelated expansion to all open clashes.

Do not reimplement the classifier ad hoc in tool-call JavaScript or vary numeric
thresholds between passes. Improve the shared code when a reusable criterion
is confirmed. Record unvalidated ideas in knowledge/proposals, with evidence
needed to promote them. Avoid report names and clash numbers in shared rules.

Use the same full input, project parameters and rules fingerprint when comparing
users. Save analysis output locally, outside versioned knowledge (e.g. local/).
Distinguish machine decisions from human/agent overrides in that saved record.

## Apply Statuses

Approved means in tolerance; Active means a clash requiring attention.
Uncertain is an analysis outcome, not a Navisworks status. It does not authorize
a status change. Resolved and Reviewed are distinct workflow states; preserve
them unless the user includes them in the requested changes.

After user authorization, apply the reviewed change list. Use a group operation
only after every child has the same verified intended status. Group membership
alone does not prove a common pipe run or opening. Use a result batch for mixed
groups. Do not update a whole group if this would change out-of-scope Resolved
children; update only eligible results instead. Save the intended names/statuses
before writing so a retry can replay
the reviewed list rather than invent new classification rules.

After writing, verify actual statuses. If a request times out, check affected
results before retrying; the bridge may still be processing the first request.

## Bridge Development

The bridge is a separate project, KPLN_NavisMcpBridge. Find the user's actual
checkout (a sibling of the rules checkout is a useful candidate) before editing
it. Rebuilding the add-in requires reloading it in Navisworks; restarting the
MCP process is a separate boundary. Knowledge-only updates are read from disk.
