# Changelog

## Unreleased

- Add geometry-first slab/straight-pipe opening review and same-group local
  surface gap < larger outer diameter bundles. Preserve short normal passages,
  reject edge/longitudinal clashes, and keep missing geometry uncertain.
- Require stable group/pipe/slab IDs and measured common-plane contours; exclude
  Resolved/Reviewed from this mode and prevent duplicate pipe contributions.

- Support duct insulation against walls with the same opening threshold as
  ducts, using exact Duct Size plus twice its thickness. Preserve uncertainty
  for missing properties and unresolved insulation hosts in bundles.

- Approve verified perpendicular rectangular and round duct or duct-insulation
  passages through AR ceilings when the nominal outer section and the complete
  mesh-intersection area match.
- Preserve uncertainty for missing geometry and partial intersections; retain
  the existing rules for other ceiling services.
- Support round ducts and rectangular/round duct insulation. Read direct duct
  Height+Width or Diameter; parse insulation Duct Size and add its thickness.
  Use 2 percent area and complete-slice tolerance.

## 0.1.0

- Migrate existing policy and opening classifier to a shared local Git checkout.
- Install a portable skill router via KPLN_NAVIS_RULES_ROOT.
- Include version, Git revision, knowledge fingerprint and input hash in analysis.
- Return Uncertain for unsupported pairs and missing decision data instead of
  silently approving them; keep established supported wall rules.
- Document project overrides and unvalidated slab/ceiling heuristics separately.
- Add anonymous regression tests and a backup-preserving installer.
- Implement user-confirmed longitudinal-only AR ceiling checks for OV/VK/PT/EOM/SS,
  using service-axis and ceiling-normal vectors without opening-size rules.
