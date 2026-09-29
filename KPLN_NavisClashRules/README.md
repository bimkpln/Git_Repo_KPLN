# KPLN NavisClashRules

Shared engineering knowledge for Navisworks Clash Detective. Python 3.10+ and
Git are sufficient for offline classification. The Navisworks 2020 add-in and
MCP server remain in the separate KPLN_NavisMcpBridge project.

## Install From Your Own Clone

Obtain this repository locally from your team's Git remote. The remote URL is
configured by your team; this package does not create or publish a remote.
From the root of your clone run:

```powershell
& .\scripts\install-skill.ps1
```

The installer sets the User environment variable KPLN_NAVIS_RULES_ROOT to this
clone and installs the thin skill in CODEX_HOME/skills (default ~/.codex/skills).
It also writes a machine-local rules-root.txt beside the installed skill so
existing/sandboxed sessions can locate this checkout without inherited variables.
It backs up an existing skill, preserves existing agents/openai.yaml metadata,
and removes the old local policy/script copies only after backup. The process
environment takes priority, followed by the local pointer and persistent User
variable. The pointer is installation data, not a file to share between users.
Already-loaded skill instructions may persist in an existing conversation;
start a new task or explicitly reload the installed SKILL.md after migration.
Navisworks and the bridge need no restart for knowledge-only updates.

For a custom location use `-SkillDirectory <path>`. For a disposable installer
test use both a test SkillDirectory and `-NoEnvironment`.

## Analyze Offline

Save the full MCP result array (or {"result": [...]}) as UTF-8 JSON. Keep Bound
for group proximity and metrics for classification. Inputs/outputs can live in
the ignored local/ directory. Supply the current project's threshold explicitly:

```powershell
python -B .\src\classify_opening_clashes.py .\local\input.json --opening-min-edge-mm 150 --compare-status --include-names
python -B .\src\classify_opening_clashes.py .\local\ceilings.json --mode ceiling-services --mep-discipline PT
python -B .\src\classify_opening_clashes.py .\local\slab-pipes.json --mode slab-openings --opening-min-edge-mm 150
python -B .\src\classify_opening_clashes.py .\local\slab-pipes.json --mode slab-wall-openings --opening-min-edge-mm 150
python -B .\src\classify_opening_clashes.py .\local\self-intersections.json --mode self-intersections --compare-status
python -B -m unittest discover -s tests -v
```

150 is an example parameter, not a repository-wide default. The CLI reads data
and prints JSON; it never connects to Navisworks or changes statuses. Each record
contains its decision, reason, current status and missing evidence. Uncertain
remains available to skill reasoning and explicit project overrides.

When the user asks for the slab report to be reviewed like the wall report,
use `slab-wall-openings` with raw item bounds retained (`keep_raw_bounds=True`).
It needs `PipeSlabGeometry.LocalFaces.Source=slab-local-faces-1` from the add-in,
but does not require a verified cylindrical mesh or closed surface topology.
It reuses the wall-style pipe direction estimate for sufficiently elongated
pipes and actual point-to-triangle distances to broad and edge/reveal faces.
Nearby bundle candidates, short pipes and unsupported directions stay Uncertain.

Current automatic coverage is wall openings for pipes, insulation, ducts and
recognized duct accessories/fittings, plus duct/finish opening cases. The
ceiling-services mode implements longitudinal checks for AR ceilings against
OV/VK/PT/EOM/SS and approves verified transverse rectangular/round duct and
duct-insulation passages whose complete ceiling intersection matches the exact
section properties and area.
The `slab-openings` mode supports verified straight circular pipe/slab mesh
geometry and the same-group gap < larger outer diameter bundle rule. Other slab
pairs, unknown pairs and incomplete data return Uncertain. Load the add-in and
MCP with `SlabGeometryVersion/SlabMetricsVersion=pipe-slab-mesh-1` before reviewing.
The `self-intersections` mode implements the trained OV/VK pipe/duct rules for
insulation, fittings treated as insulated, bare bodies, axis crossings and
parallel extended routes. It does not use an opening-size threshold; unsupported
bare-body pairs and incomplete geometry remain Uncertain.
Existing wall heuristics remain engineering approximations; synthetic tests
verify implementation behavior, not universal geometric correctness.

Read `knowledge/proposals/review-lessons.md` before extending coverage. Earlier
slab actions from human review remain project decisions; the new pipe/slab
classifier requires fresh mesh evidence and does not validate those actions retrospectively.
Equal-size duct/valve groups
also retain a documented connectivity limitation.

Ceiling input must include MepAxis and CeilingNormal, each an [x,y,z] vector or
an {X,Y,Z} object in the same coordinate system. Selecting ceiling-services
declares that the input test is AR ceilings against the supplied discipline.
These vectors are a rules-engine contract, not fields guaranteed by the current
bridge. Use extracted/verified axes, never guess a short cylinder's axis from
its longest bounding-box dimension. Missing vectors yield Uncertain. For example,
MepAxis=[1,0,0] and CeilingNormal=[0,0,1] are longitudinal and Approved; a vertical
axis with that normal is outside the longitudinal scope and causes no write.
Verified straight duct and duct-insulation passages are the explicit exception: a true
DuctSectionFootprintMatch with source object-section-properties+duct-ceiling-mesh-section
returns Approved without requiring a guessed longitudinal axis. Missing or
nonmatching footprint data does not prove a transverse passage.

## Share And Improve

The installed skill always reads this checkout. Edit rules/, src/ and knowledge/
here. Edit skills/ only when the router needs changing, then rerun the installer.
Use normal local Git branches and review/merge to exchange changes. Only the
owner writes the original X: checkout; other users work in their own clones.
Pulling is an explicit development action, not part of analyzing a live model.

Output identifies VERSION, Git HEAD (null before the first commit), dirty state,
a SHA-256 fingerprint of code/policy/knowledge/skill and a hash of the input.
Fingerprinting normalizes text line endings for Windows clones. Compare these
values and project parameters to reproduce machine decisions. Human/agent
overrides must be recorded separately with evidence and scope; flexible reasoning
does not promise identical decisions between users.

clash_rules.json is not required. Project parameters may be saved with each
analysis; the repository does not reinstate a JSON rule language.

## Rollback

The installer prints the backup location. Restore the backed-up skill directory
to its original path to return to the former self-contained skill. The backup
also preserves the original classifier and policy. KPLN_NAVIS_RULES_ROOT can be
removed from the User environment if no longer needed. Do not overwrite later
user changes without reviewing the backup first.
