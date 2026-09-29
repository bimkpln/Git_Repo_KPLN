# Review Lessons And Promotion Candidates

These observations preserve experience from human review. They are not blanket
approvals or executable rules. Project overrides belong to the analysis record.

## Slab Penetrations And Short Pipe Segments

A user subsequently requested a wall-equivalent slab review instead of making
ideal cylinder topology a gate for every result. The explicit
`slab-wall-openings` mode now separates local slab-face measurements from pipe
mesh fitting. Synthetic cylinder tests passing does not prove that a live
export contains weld noise, differently tessellated fragments or oblique ends;
do not report any of those hypotheses as an established cause from a generic
failure reason. Read actual evidence before changing extraction tolerances.

The user has now authorized the geometry-first slab/straight-pipe implementation
and the same-group local gap < larger outer diameter bundle criterion. These
are implemented in `slab-openings`; see the policy for exact scope and limits.
Synthetic coverage includes short/long/hollow pipes, inclined slabs, edges,
holes, group IDs, strict gap boundaries and workflow exclusions. Real model
verification still requires the matching add-in and MCP versions to be loaded.

A user accepted short vertical penetrations below the project opening threshold.
This supports treating shortness alone as insufficient reason to reject a
penetration. It does not validate a universal AABB aspect-ratio cutoff.

Previous exploratory passes used Z/XY ratios of 2, 1.05 and 1, and sometimes
inherited orientation from another group member. These produced different
counts and were not saved in the old classifier. Do not reproduce those changing
heuristics as established shared behavior. A short vertical cylinder can have
X = Y > Z; the largest box dimension is not necessarily its longitudinal axis.
Likewise, a vertical pipe can hit a slab edge rather than pass normally through it.

Before automating: obtain a reliable service axis and local slab-face relation,
distinguish pipe bodies from fittings, and test short cylinders, long cylinders,
inclined services, slab edges, voids and multi-service groups. A group member's
axis does not establish the axis of its neighbors.

## Ceiling Services

The user promoted longitudinal-only review of OV/VK/PT/EOM/SS against AR ceilings
to shared policy and the ceiling-services classifier. That scope is implemented,
not a pending proposal. The earlier blanket status overrides below remain
project-specific history, not the behavior of this mode.

Longitudinal means along the actual ceiling plane, which must be established.
Long vertical risers are transverse to a horizontal ceiling. Later explicit
user instructions approved all open results in reviewed ceiling checks, including
transverse cases. Preserve that as a project override, not a universal rule
that all ceiling-service clashes are acceptable. Unknown orientation remains
uncertain; do not infer approval from the word finish in a category alone.

Subsequent human review specifically approved perpendicular rectangular and round
duct passages, including duct insulation, when the actual ceiling footprint
matches the dedicated section properties and area. Direct ducts use Height+Width
or Diameter. Insulation uses parsed Duct Size plus twice its thickness. Area and
complete-slice comparisons use a 2 percent tolerance. This is implemented as a scoped exception,
not a blanket approval of all transverse service clashes. Short duct length
may be smaller than either section side. Require the complete slice to overlap
the ceiling so a partial longitudinal intersection cannot mimic a normal passage.

## Complex Walls And Local Ends

Human review confirmed that small pipes entering ends or reveals can remain
real clashes despite meeting the nominal opening threshold. Mesh-PCA and local
face distances help, but a single global axis can misdescribe complex walls.
Nearest-point face distance is not proof that the entire pipe avoids an edge.
Do not overturn an end-intersection decision from one nonzero distance alone.
Keep ambiguous complex-wall cases available to further geometric/human review.

Human clarification established a separate bundle fact: one wall-pipe result
means the opening bundle contains that one pipe, so its own section is sufficient
for the bundle-size calculation. This does not erase positive evidence of a
longitudinal run or contact with an end/reveal; those remain Active. The shared
classifier now uses the single-pipe size when such decisive geometry is absent,
instead of demanding evidence for nonexistent neighboring pipes.

## Inline Components And Groups

The shared classifier preserves a same-section duct/valve heuristic from earlier
reviews. Equal sections alone do not prove inline connectivity. A future geometry
check should establish alignment and continuity before this is generalized.
Group status changes are an execution shortcut after all children are reviewed.

## Materials And Worksets

Human-approved review used steel materials containing C245 or C255 (the model
may use Cyrillic C) against concrete, with a separate workflow for reinforcement
worksets. These are project-dependent material/workset criteria. Retrieve the
steel participant's workset from Object properties and ancestors; do not test
the other participant's workset by accident. Define allowed spellings or a
reviewed normalization when the user says similar. Missing values are unknown.

## Door And Revision Hatch Opening Zones

Human review explicitly permits overlap of a door opening zone with a
revision/access hatch opening zone under sequential use. This is promoted to
the scoped policy in rules/clash-policy.md, not inferred from one previously
Approved result. Identify both zone geometries and their respective parent
roles before applying it. Do not generalize to physical-body intersections,
other opening-zone pairs or compliance/accessibility conclusions. Differing
level labels are not evidence of vertical separation. This remains policy-led
review; the pipe/duct self-intersection classifier does not cover these objects.

## Promoting A Lesson

Record the proposed condition, applicability, supporting human feedback,
counterexamples and missing evidence. Once applicability is established, update
rules/clash-policy.md and add a synthetic behavioral test if implemented in code.
No production file paths or specific result IDs should be needed to explain it.
