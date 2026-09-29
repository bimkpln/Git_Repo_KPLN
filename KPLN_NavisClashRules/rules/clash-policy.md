# KPLN Clash Review Policy

This reference condenses reusable rules from the Claude-origin skill while removing file-specific conflict numbers, old report ids, and one-off examples. Use it to classify new Navisworks Clash Detective results and to decide whether a newly observed pattern is suitable for training.

## Architecture Boundary

- The C# add-in should expose raw COM/managed Navisworks data: paths, statuses, geometry, groups, properties, and comments.
- The Python MCP server should expose Navisworks data and normalized metrics: feet to millimeters, axes, angles, elongation, perpendicular relation, element class/kind, size parsing, centers, bounds, and status-write commands.
- The skill supplies review policy: thresholds, rule order, project-specific questions, and how to treat untrained or unreliable cases.

Keep engineering policy out of both the C# extraction layer and the MCP server. This repository owns policy and classifier code; the installed skill routes to this local checkout.

## Shared Execution And Coverage

Resolved (`eTestResultStatus_RESOLVED`) clashes are outside the default analysis
scope. Include them only on an explicit user request to analyze Resolved
clashes; a general request for a report or "all clashes" does not override this
exclusion. Apply the scope filter before classification, training comparisons,
group/bundle aggregation and decision counts. A full read-only snapshot may
retain these rows for total-status counts and preservation checks, but they do
not contribute to classification of other clashes. Report them as excluded,
not as analyzed, Approved or Uncertain.

An explicit request to analyze Resolved clashes is not permission to change
their statuses. Without explicit authorization to modify them, preserve them
also inside groups: use individual-result updates when a group operation
would affect out-of-scope Resolved children.

Use `src/classify_opening_clashes.py` for supported wall-opening pairs. Its
Uncertain result represents missing evidence or unsupported scope, not a
Navisworks status. Earlier guidance to keep unreliable cases Active means
withhold automatic approval; it does not prove that a real clash was established.

Straight circular pipe/slab reviews use the `slab-openings` mode described
below. Other slab pairs and material/workset decisions require policy-led
analysis. Read `knowledge/proposals/review-lessons.md` before generalizing.
Do not apply size-only fallback to unsupported or geometrically uncertain pairs.

For reproducibility, retain the full input, project parameters, classifier
output (including hashes), and separately reasoned human/agent overrides.
Matching Git commits alone is insufficient if local files, inputs, project
parameters or overrides differ. Free-form analysis remains flexible and is not
guaranteed to be identical between agents.

## Training Discipline

When the user provides a human-reviewed report:

1. Retrieve all in-scope results, not only Active or only Approved. Exclude Resolved by default unless the user explicitly included them; if a full snapshot is fetched, filter before comparison.
2. Run the current rules and compare them to the human statuses.
3. Group mismatches by measurable cause: class pair, size, group behavior, axis relation, project parameter, construction type, or missing classifier.
4. Ask about mismatches when no reusable rule explains them.
5. Update this skill and its scripts only with general rules.

Do not encode the following in a skill or ruleset:

- individual clash numbers;
- saved viewpoint numbers;
- old file names as decision criteria;
- "approve this exact named result" lists;
- exceptions that are not explainable by geometry, category, project parameters, or KPLN policy.

A report where every result is Active may be unreviewed rather than "all real clashes"; confirm before learning from it.

## Numerical Tolerances

The current KPLN regulation does not provide a universal clash tolerance matrix. For a Clash Detective test, read `Tolerance`, `ToleranceType`, and `TestType` from `list_clash_tests`.

- If a test has a meaningful configured tolerance, use that as the numeric source of truth.
- If tolerance is zero or absent, ask the user which project rule applies.
- Do not transfer a tolerance from a different test or project.
- The stable policy number for duplicate checks is 25 mm.
- Raw COM distances and bounds are in feet; multiply by 304.8 for millimeters. Metrics returned by the Python server are already in millimeters.

Many tests only contain results beyond their configured tolerance because Navisworks has already filtered smaller ones. In those tests, tolerance alone may not decide the status.

## Metrics To Prefer

Use `get_clash_test_results` with `include_metrics=True` for classification. Prefer these computed fields over manual interpretation:

- `Kind1`, `Kind2`, `PairKind`: pipe/duct self-intersection taxonomy such as insulation, fitting, bare pipe.
- `Class1`, `Class2`, `PairClass`: building-service taxonomy such as wall, slab, finish, duct, duct fitting, valve, pipe, insulation.
- `Angle`, `PerpRel`, `Dia*`, `Len*`, `Elong*`: axis geometry for pipe/duct self-intersections.
- `Size1`, `Size2`, `SectionMin`, `SectionMax`, `SectionSource`: Revit size string and parsed nominal section.
- `Group`: Navisworks clash group name when exported by the bridge.
- `MepMinEdge`, `StructThick`, `Through`: minimum MEP bounding edge, structural thickness, and through/intersection indicators.
- `DistMm`, `GeomOk`, `GeomNote`: distance and metric reliability.

Request raw bounds only when the metrics are insufficient, using the bridge option intended for raw bounds.

Avoid these known traps:

- Do not infer element direction from the clash intersection `Bound`; it can invert the engineering conclusion.
- Do not use bounding-box section sizes when `SectionMax` is available from the Revit size property.
- Do not use pipe-axis metrics for structural elements.
- Do not classify by searching the whole path string. Prefer the Revit category segment of the path.
- Do not reconstruct bundles only from distance. Use Navisworks grouping and geometric proximity together when bundle logic is needed.

## Element Classification

Classify `Class*` from the Revit category segment, not from arbitrary words in type names. Type names may contain misleading words such as inserts, fittings, or system abbreviations.

For `Kind*` in insulated pipe/duct self-intersection tests:

- fittings and duct/pipe accessories may need to be treated as insulated even when insulation is not exported to NWC;
- ordinary pipe/duct insulation may be exported as a separate body;
- if exported insulation touches another element's bare body, that is usually more severe than insulation-to-insulation contact.

When a family is misclassified, do not patch around one result. Add a reusable category/family pattern to the server's data classifier and then re-run the evaluation. The classifier may identify what an element is, but it must not decide whether the clash is acceptable.

## Ruleset A: OV/VK Self-Intersections

Applies to trained pipe/duct self-intersection tests where the purpose is to distinguish tolerable insulation contact from real intersections.

Evaluate in order; the first matching rule wins:

1. Insulation against an ordinary bare pipe/duct body, excluding fittings/accessories treated as insulated, is a real clash regardless of angle.
2. `PerpRel <= 0.15` is a real clash because the second element passes through the axis rather than touching tangentially.
3. `Angle <= 5` and at least one participant has `Elong >= 5` is a real clash: parallel extended routes overlap.
4. Otherwise, transverse or local insulation contact is in tolerance.

Bare-body against bare-body cases are not covered by this ruleset unless a trained project rule says otherwise. Ask before classifying them automatically.

Treat possible end-of-pipe contacts and unusually elongated fittings as candidates for future training, not as automatic rules until the user confirms the reusable criterion.

## Ruleset B: Structure Against Ducts Or Services Requiring Openings

This ruleset depends on the project parameter `opening_min_edge_mm`: the minimum nominal edge or equivalent diameter for which an opening is modeled. Ask for it at the start of a review and do not cache it across projects.

Reusable policy:

- Geometry comes before nominal size for wall penetrations. A diameter below
  the opening threshold is only eligible for tolerance when the service makes
  a normal passage through the broad face of the wall.
- For diagonal walls, never infer wall direction from an axis-aligned bounding
  box. Use the wall plan axis recovered from its mesh geometry (`mesh-pca`) and
  compare the service axis with the wall normal.
- A pipe entering the end of a wall is a real clash regardless of diameter.
  This includes external wall ends and internal end/reveal faces at doors,
  niches, and other openings. Detect internal faces from nearby mesh triangle
  normals: a face normal aligned with the wall axis is an end/reveal face,
  while a normal perpendicular to the wall axis is a broad wall face. If the
  service axis lies within one service radius of an end/reveal face, keep the
  clash Active. For an external wall end, use the larger of wall thickness and
  half the service diameter as the end zone.
- A pipe whose component along the wall plane is greater than its component
  through the wall normal (`WallNormalAngle > 45°`) is a longitudinal wall
  intersection, not an opening, and remains Active regardless of diameter.
- If wall orientation or end distance is missing or unreliable, do not approve
  a wall-pipe clash from diameter alone; keep it Active for manual review.
- For pipe-insulation clashes, the Revit `Размер` value may describe the
  insulation layer rather than the complete insulated assembly. Use the larger
  of nominal size and geometric outer diameter (`MepMinEdge`) after the
  geometry checks above.
- Finish layers such as plaster, insulation, mesh, or lining usually do not require modeled openings by themselves. Do not treat this as an absolute approval: when an oversized duct passes through a finish layer, keep the result Active because the opening issue is still visible in that layer.
- For a single straight duct, use nominal `SectionMax` from the Revit size property first. Fall back to `MepMinEdge` only when the size property is missing, and treat that as a lower-confidence estimate.
- If the relevant opening size is strictly greater than `opening_min_edge_mm`, the result is normally a real clash: an element that requires an opening must be inside that opening. For shallow, non-through edge cuts in a near-coincident mixed-size group, do not rely on the single size alone; classify by the common opening evidence.
- If the relevant size is equal to or below the threshold, it is in tolerance, unless another trained rule for the element kind or bundle overrides it.
- Use strict `>` comparison at the threshold unless the user defines a different project convention.

Duct insulation against walls follows the duct opening rules using its own
outer section. Read Object -> Duct Size (rectangular or diameter notation)
and add twice the insulation thickness to both edges or to the diameter.
The bridge supplies SectionMin/SectionMax in millimeters with SectionSource
Object/Duct Size and InsulationThicknessMm. A missing exact property or thickness
is Uncertain; the old generic Size or narrow AABB edge cannot replace it.
Direct ducts use exact Object Height+Width or Diameter when exported.
Pipe insulation retains its existing outer-size and wall-geometry checks.
Do not count an insulation shell and its duct as two bundle members. For
multiple insulation members require unique host associations based on section
and containing item bounds before applying a common-opening decision. Use each
matched outer section once when calculating the common opening.

Accessories and fittings are not one class for approval:

- Rectangular fire dampers and other rectangular duct accessories use their nominal rectangular connection/body size. When that size is above the opening threshold, keep the result Active.
- Round fire dampers must not be rejected only because their bounding box is above the threshold: the box often includes casing, flange, or actuator rather than the nominal opening. If a round damper's exported box is above the threshold but the clash is a shallow wall cut rather than a near-through pass, keep it Active; if it is near-through and no other rule triggers, keep it in tolerance.
- A round damper in the same physical opening cluster as a directly Active duct or bundle follows that Active decision, because the clash belongs to the common opening issue.
- Duct elbows, transition fittings, and other fittings against a wall are not automatic tolerance cases. When the fitting body itself intersects the wall outside a modeled opening, keep it Active unless a more specific trained rule proves that the nominal opening is compliant.

Bundle policy:

- A standalone wall-pipe result or a group containing exactly one pipe is a single-member bundle. Use that pipe's own opening section; do not require multi-member bundle geometry. Available evidence that the pipe enters an end/reveal or runs longitudinally in the wall still takes precedence and remains Active. If those decisive conditions are absent, missing wall geometry alone does not block a size-based decision for the verified single pipe.
- Do not sum every Navisworks group blindly. A group is only a bundle candidate when it has at least two contributing MEP members and their clash-zone centers are close enough to represent one common opening rather than separate penetrations in the same wall.
- A group containing exactly one straight duct and one inline valve/accessory with the same nominal `SectionMin` and `SectionMax` represents one continuous passage, not two bundle members. Evaluate each component against the opening threshold, but do not sum their equal sections.
- Contributing members are straight ducts, duct fittings, and rectangular accessories. Ignore round dampers when computing bundle size because their bounding boxes are unreliable for opening size; use their clash centers as supporting geometry when deciding whether the group is one common opening, and let them inherit the decision from an active bundle.
- Use both Navisworks `Group` and geometry. The middle spread of clash-zone centers should be greater than half of the largest member section, to avoid duplicate coincident clashes, and no more than about three summed member sections, to avoid unrelated openings.
- Compare the effective common-opening size, such as `max(sum(member_sections), largest_member_section + middle_center_spread)`, against `opening_min_edge_mm` with strict `>`.
- When a bundle exceeds the threshold only by rounding noise, such as about 1 mm, treat it as a boundary case. Require either a spacing-based effective opening above the threshold or clearer evidence that the members form one required opening.
- If a two-member group already contains one member directly above threshold, do not automatically propagate Active to the smaller member; require a larger contributing cluster or clearer common-opening evidence.

Do not transfer a depth rule from a wall-pipe test to a duct-opening test. Ask about special cases such as ducts cutting the top of a partition when size alone would approve them.

For rectangular ducts, the user has confirmed the equivalent diameter formula for equal area: `d = sqrt(4ab/pi)`.

Use [../src/classify_opening_clashes.py](../src/classify_opening_clashes.py) for supported wall-opening cases. It expects JSON from `get_clash_test_results(include_metrics=True)`; include `Bound` in the data when bundle rules need clash-center spread. The classifier does not write Navisworks statuses.

## General KPLN Category Rules

### Structural Slabs Against Straight Circular Pipes

When the user requests the same practical review as the wall/pipe report,
use `--mode slab-wall-openings`. This is an explicit alternative to the strict
mesh-section mode below; do not make cylinder topology a prerequisite for it.
The add-in exports `PipeSlabGeometry.LocalFaces` with source
`slab-local-faces-1`: independently measured slab normal, nearest broad-face
distance and nearest edge/reveal-face distance, using the wall workflow's
point-to-triangle distance. This includes internal opening reveals. Failure
to fit the pipe mesh must not suppress these slab measurements.

The wall-style mode uses the same item-bound axis estimate as wall review for
sufficiently elongated bare pipes. It is currently scoped to horizontal slabs;
short pipes (including discs whose height is below the nominal diameter),
inclined slabs without a signed pipe axis and missing local face distances
remain Uncertain. Contact with an edge/reveal within half the nominal section
is Active; an axis angle above 45 degrees to the slab normal is Active.
Otherwise a normal broad-face passage uses nominal section strictly above the
project threshold as Active and at/below it as Approved. Preserve Reviewed and
Resolved. Never manufacture missing local edge measurements from slab AABBs.

Keep the user-confirmed bundle condition below. For vertical members on the
same slab, separate penetrations can be proved conservatively when the XY
distance between complete item bounds is at least the larger transverse box
width: actual surface distance is no smaller and actual circular diameter is
no greater. Deduplicate pipe IDs and split by slab ID. This test only rules
out a bundle; a remaining nearby candidate requires measured local sections
and stays Uncertain until those are available. Longitudinal members retain
their own Active decision and do not enlarge a transverse opening solely
because the user grouped the results together.

The original `slab-openings` mode remains available for snapshots containing
the full cylinder/section evidence. Its stronger requirements are not a
prerequisite for an explicitly requested wall-style review.

Use `--mode slab-openings --opening-min-edge-mm <project threshold>`.
Require `SlabMetricsVersion=pipe-slab-mesh-1`, a verified cylinder axis,
closed slab mesh with two parallel broad planes, actual section contours,
and stable pipe/slab/group identities. The normal comes from mesh surfaces;
neither pipe direction nor slab direction may be guessed from AABB dimensions.
All coordinates must share the model/world coordinate system.

Evaluate geometry before the project's opening threshold:

- A pipe envelope contacting an external slab side or an internal hole/reveal
  face is Active, regardless of diameter. The bridge clips each finite side
  triangle to the pipe's actual axial interval and tests the whole outer radius.
- A pipe with a component along the slab greater than its component through
  the slab normal (angle >45 degrees) is Active. Exactly 45 degrees is Uncertain.
- For a transverse passage, require its full outer slice to be covered by each
  intersected broad face. Require at least one such face; a short pipe does not
  need to span both slab faces. This does not approve pipes wholly embedded in
  a slab without a broad-face passage. Missing/partial evidence is Uncertain.
- A verified single passage uses the nominal pipe section: strictly above the
  project threshold is Active, otherwise Approved. Do not silently add clearance
  to the explicitly supplied section threshold.
- Resolved and Reviewed are excluded from this mode and from its bundle graph;
  preserve their workflow statuses. This mode does not classify fittings.

User-confirmed pipe bundle rule: two distinct pipes belong to a bundle when
they have the same actual Navisworks group ID and their surface-to-surface
gap at a common slab surface is strictly smaller than the larger outer pipe
diameter. Evaluate in the same slab; neither center spacing, clash Bound,
group names nor total pipe length replaces the measured local gap. Connected
pairs form a bundle. Deduplicate repeated results for the same pipe instance.

Compute the common opening from the convex hull of the member section contours
on the slab planes, projected into the slab plane, using the larger side of
the minimum-area enclosing rectangle. Include spacing; do not simply add all
diameters. A bundle opening above the threshold makes its otherwise approved
members Active. It never overrides an edge/longitudinal clash or uncertainty.
Missing group geometry prevents automatic approval of affected neighbors.
The existing duct/wall-specific bundle rules remain scoped to those pairs.

Extraction limitations are explicit: noncylindrical, multi-axis, disconnected,
incomplete or heavily tessellated unsupported pipe meshes, nonwatertight,
stepped or ambiguous slab meshes remain Uncertain. The bridge currently
validates straight cylinder endpoint rings and supports open pipe caps, hollow
pipes, short segments and inclined slabs; it does not infer connectivity from
another member's axis. Numerical geometry tolerances are not engineering
opening thresholds; the 2-percent duct/ceiling allowance does not apply here.

### AR Ceilings Against OV, VK, PT, EOM And SS

User-confirmed base scope: analyze longitudinal intersections for these
five disciplines against AR ceilings. The `ceiling-services` classifier mode
implements this scope. A longitudinal service has a greater component in the
ceiling plane than along its normal, analogous to WallNormalAngle > 45 degrees.
Confirmed longitudinal results are Approved. Except for the verified duct and
duct-insulation passages below, transverse and 45-degree boundary
cases are outside this approval scope; preserve their statuses. Missing direction
or ceiling-plane data is Uncertain. Do not apply opening-size rules in this mode.

For straight ducts, read dedicated Object properties: Height+Width for a
rectangular duct and Diameter for a round duct. Do not depend on the formatted
Size string. For duct insulation, parse Object -> Duct Size as either `150x500`
or `ø100`, then add twice the insulation thickness to each outside dimension.
Verify both rectangular sides or the round diameter and the corresponding area
against the actual intersection footprint on the ceiling mesh surface. A matching section means a transverse passage
and takes precedence over a longitudinal estimate from a short duct's shape.
Equal area alone is insufficient: different rectangles can have the same area.
Use the bridge's DuctSectionFootprintMatch with source
object-section-properties+duct-ceiling-mesh-section. Its numerical comparison uses max(2 mm,
2 percent) for each side/diameter and 2 percent for area and full-slice coverage; retain measured deviations in
the review record. This is a geometry comparison, not an opening-size rule.
The footprint must be clipped against the ceiling surface, including holes
and partial overlap; the clash Bound or its projection is not that footprint.
Missing size/footprint is Uncertain. A nonmatching footprint alone does not
prove longitudinal direction: that still requires reliable axis/normal evidence.
User clarification: a confirmed perpendicular rectangular or round duct passage, including its insulation, through
an AR ceiling is Approved. A true DuctSectionFootprintMatch from the documented
mesh/size source is sufficient; an axis guessed from the short duct's longest
dimension must not overturn it. The intersection must cover the complete duct
slice (FullDuctSectionCovered in current bridge metrics), not a clipped portion
of a larger longitudinal slice which happens to match the nominal dimensions.
This exception is specific to verified straight ducts and duct insulation. It does not approve
unknown geometry, fittings or all transverse service clashes.
Existing longitudinal rules for other services remain unchanged. The classifier
only produces decisions; writing statuses is a separate authorized operation.

Supply reliable MepAxis and CeilingNormal vectors in the same coordinate system
for each result. They are normalized input fields for the rules engine, not
currently guaranteed MCP fields. Do not fabricate vectors from the longest AABB
dimension of a short element. An explicit user-reviewed axis/plane can be used
with its source recorded in the analysis. Otherwise request geometry extraction
or retain Uncertain. Sloped ceilings require their actual normal.

### Door And Revision Hatch Opening Zones

Mutual overlap between a door opening zone and a revision/access hatch opening
zone is Approved under the user-confirmed allowance for sequential opening.
Both actual clash participants must be identified as opening-zone geometry;
establish the door and hatch roles from their categories/properties and the
relevant parent family, not a report name, clash number or arbitrary path text.
The allowance does not require assuming that the zone solids are physical
panels. Missing or ambiguous participant identification remains Uncertain.

This rule approves only zone-to-zone overlap. It does not approve door-to-door
or window-to-window cases, zones against physical bodies, or intersections of
panels, frames or hardware. It does not establish equipment accessibility,
independent opening with the neighbor closed, or fire/egress compliance; review
any evidence of those problems separately. Neither penetration depth nor
different associated level labels is an approval criterion. Preserve Resolved
and Reviewed unless the user explicitly authorizes changing them.

### Other Category Cues

These are reusable policy cues for other section pairs. Apply only when the current test and project stage match the condition.

- OV/VK insulation: use the dedicated self-intersection rules where applicable.
- Architecture/structure against engineering insulation: generally not corrected except insulation in vertical shafts or intersections of parallel elements.
- For AR ceiling checks of OV/VK/PT/EOM/SS use the scope above, including verified straight duct and duct-insulation passages.
- Mullions against curtain walls: stage 4 may be unconditional; stage 5 depends on the absence of through intersections.
- AR/KR walls, beams, slabs against AR/KR: may be non-corrected when the total structural volume error is within 3 percent and there are no through clashes.
- Piles against foundation slabs, thickenings, or pits: may be non-corrected when penetration depth is less than 300 mm.
- Structural slabs against plumbing floor drains (трапы) are in tolerance. Approve only when one participant is classified as a slab and the other belongs to the sanitary-fixtures category with a family or type that explicitly identifies a floor drain; do not extend this rule to generic plumbing equipment.
- Maintenance/access zones: often informational, not corrective.
- Valve actuators: may be non-corrected when no technical access requirement is violated.
- EOM incoming cable against IOS: generally non-corrected.
- Cross-linked polyethylene pipe in screed or conduit in wall: may be non-corrected when the project intentionally models it that way.
- Stage 4 main IOS-vs-IOS networks: check only mains at or above these section sizes unless project rules override them: ducts edge 300 mm, pipelines outside diameter 50 mm, trays edge 200 mm. Do not apply this stage 4 rule automatically at stage 5.

## Openings For Engineering Systems

Use project-specific opening rules first. When the regulation is the only available source:

- vertical constructions: clearance between engineering element section, excluding insulation, and opening is at least 49 mm;
- horizontal constructions: clearance is at least 99 mm for OV2 ducts and 49 mm for pipes and trays;
- stage 4 opening threshold: opening edge/diameter below 500 mm;
- stage 5 thresholds: KR walls/slabs/roof below 200 mm, ordinary AR walls below 300 mm, gypsum board below 500 mm;
- metal mesh, insulation, and finish layers do not receive modeled openings.

Confirm project convention when these thresholds conflict with the Navisworks test setup or BEP/project instructions.

## Exclusions And Coordination

Excluded from duplicate checks in all stages: evacuation routes and topography. At early stages, special shaft families may also be excluded.

Often excluded from collision checks:

- duplicate elements from adjacent models;
- additional KR structures;
- stage 4 and 5 sensors, sockets, lights, SS devices, and EOM devices against everything;
- self-intersections of cross-linked polyethylene pipes in screed;
- flexible pipes and flexible ducts against everything.

Intersections with beams, columns, capitals, and pylons must be coordinated with KR.

Do not apply old responsibility matrices unless the user confirms they are current for the project.

## Review Workflow

For a trained test:

1. Use `list_clash_tests` to inspect name, tolerance, and result counts.
2. Choose the applicable skill ruleset from this reference by the engineering pair, not by a project-specific filename.
3. Ask for missing project parameters.
4. Retrieve results with `include_metrics=True`; for bundle analysis keep `Bound` available and process the data with the skill script instead of reviewing chunks by eye.
5. Classify by ruleset and metric reliability.
6. Summarize counts by decision rule and list unresolved or unreliable groups.
7. Use status-setting MCP tools only after the decision basis is clear and the user has authorized bulk changes.

Timeouts during large status updates do not necessarily mean failure. Wait briefly and verify with a read-only query.

After rerunning a Clash Detective test or reopening a model without saving, statuses may reset to Active.

## Updating The Bridge

When changing the bridge itself, locate the user's separate KPLN_NavisMcpBridge checkout. Engineering rules are edited in the current KPLN_NavisClashRules checkout.

Important restart boundary:

- rebuilding the Navisworks add-in does not restart the Python MCP server;
- restarting the MCP client does not reload the Navisworks add-in already loaded in Navisworks;
- if a new tool argument appears ignored, check the MCP tool schema first because unsupported arguments may be dropped before reaching the add-in.

Known useful COM API details:

- result group: `InwOclTestResult2.GroupPath` with a safe cast from `InwOclTestResult`;
- properties: `ComApiBridge.ToModelItem(path)` then `ModelItem.PropertyCategories`, `PropertyCategory.Properties`, and `DataProperty.Value.ToDisplayString()`.
