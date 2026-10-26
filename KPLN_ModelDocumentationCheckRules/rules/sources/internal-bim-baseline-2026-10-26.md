# KPLN Internal BIM Baseline

Source context:

- `СО 2.4. План реализации проектов.docx`, document date 21.01.2021, last changed 07.07.2025.
- `KPLN_Чек-лист ключевых позиций при моделировании в Revit.docx`.

Use this as KPLN's internal baseline when comparing customer EIR/BEP requirements.

## Purpose And Governance

- BEP is the project BIM implementation plan. It is normally prepared on the basis of EIR when EIR exists.
- KPLN designers must use a unified project structure, naming rules, Revit modeling rules, and clash-check procedures.
- Changes to internal standards are handled by BIM department.
- Exceptions to BIM/modeling rules must be explicitly agreed with BIM department or fixed in BEP.

## Software And Environment

- Base software in the contractor-transfer section: Revit 2023, AutoCAD 2020, Navisworks Manage 2020, Russian UI.
- If customer source data, templates, or previous-stage models are in another version, software version must be separately agreed with KPLN BIM department.
- Revit Server use depends on customer environment. If the project initially works on customer Revit Server, periodic model upload rules may be canceled.
- BIM360/Review is preferred when BIM360 is used, because it tracks document versions and review/issue history.

## Model Creation And File Lifecycle

- Project Revit files are created only by BIM department from current KPLN templates by request of GIP or project manager.
- When project stage changes, BIM department copies/moves files to the correct stage folder.
- Do not delete from the model: previous-stage sheets/views, view templates/filters used for previous-stage documentation, and continuing elements from previous stages.
- When correcting a previous stage from a later-stage model, modify the preserved previous-stage sheets/model data rather than recreating everything.
- Auxiliary files include general plan model, federated model, coordination file, and grid/level setting-out file.
- The coordination file is the single base for absolute/relative coordinates and true north.
- The setting-out file contains grids and levels; all discipline files load it and copy/monitor grids and levels.

## Naming And Organization

- Revit/AutoCAD/Civil project file names use fields: project code, building/section/fragment/stage, project stage, discipline code, subdivision code, software version, optional addition.
- Model/project file separator is underscore `_`; spaces outside separators are forbidden.
- Exported DWG/DWF/PDF names use hyphen `-`.
- File and parameter naming generally uses Cyrillic unless a project composition already contains Latin characters.
- Project code is assigned by BIM manager; do not invent codes without notifying BIM department.
- Model division is decided by BIM department and communicated to project participants.
- Parameter names use `КП` author code for KPLN common/project parameters. Examples: `КП_О_Марка`, `КП_Р_Ширина`, `КП_О_Секция`.
- Only BIM department may edit the shared-parameter file and project parameters.
- Loaded family names and type names follow KPLN coding. Descriptive parts must be unambiguous, not unexplained abbreviations.

## General Revit Model Requirements

- Maximum model weight: 350 MB for AR/KR; 200 MB for all IOS disciplines.
- Model only elements that exist in project documentation.
- Tabular data shown in documentation must be calculated automatically from modeled elements; exceptions require BIM department approval.
- Grids and levels must be copied from the setting-out file and monitored.
- Link instances, grids, and levels must be pinned.
- Every level must have a correct `На этаж выше`; default is allowed only for top levels.
- Model at true size 1:1 in metric units.
- Elements must belong to the level where they are actually located; level offset greater than floor height is not allowed.
- Volumes/sheets are оформлены in the same model where the documentation elements are created.
- Room area parameters in room schedules use one decimal place with calculated values.
- Do not use Revit Mirror tools for side-dependent elements such as windows and doors.
- Title block families must follow KPLN documentation standard.
- Designers must not edit families directly; family creation/changes go through BIM department with concise task and source drawings/links.
- Revit Phases (`Стадия возведения`, `Стадия сноса`) and Design Options are prohibited unless separately developed and approved for a project.
- CAD Import is prohibited. CAD Link is allowed; source DWG links must be delivered with the model.
- Inaccurate modeling followed by rounded dimensions is prohibited. Accuracy: AR/IOS to 1.0 mm, KR to 0.1 mm, grid distances to 0.00001 mm.

## Required Geometry Subcategories

Use strict subcategory names for additional model geometry:

- `Зона обслуживания`: equipment access/service zones.
- `Зона действия`: operating/action zones.
- `Зона открывания`: opening swing/clearance.
- `Зона открывания_Обратная сторона`: access zone behind opened element.
- `Атрибут`: extra model attributes such as soil, glue, fasteners.
- `Раструб/Фланец/Муфта`: socket/flange/coupling geometry where unavoidable collision exists.
- `Крепление`: brackets, boxes, and mounting elements.
- `Привод/Ручка`: drives/handles for MEP elements.

## Model Division

- Stages 1-4: AR/KR division priority is standalone building, corpus/stylobate, then sections when there are more than four sections.
- Stages 1-2: IOS may be one model for all IOS sections.
- From stage 5: model division follows discipline-specific rules.
- AR: do not model more than one corpus in one file; facade may be separate; upper/lower parts are distinguished around first residential floor.
- KR: generally `1 KJ/KM volume = 1 model`, with stated exceptions; each KJ file concrete volume should not exceed 1200 m3.
- VK/PT, heating, ventilation, power, low-current systems are divided by corpus where applicable; one file must not contain more than one corpus unless specifically allowed.
- SS is always divided into two files: SS and AV.

## Stage And LOD Logic

Internal stages include: architectural concept, fire concept, engineering concept, project documentation, BIM model 1, BIM model 2, working documentation.

- Stage 4 is typically used to issue preliminary quantities and adds calculation parameters.
- LOD300 appears as the target for many stage 5/6 and working-documentation elements.
- Exact LOD requirements must be checked by element and discipline, not assumed globally.

## Architecture (AR)

- Vertical constructions must be split by floors.
- Multilayer walls are prohibited: each wall layer should be a separate type/element, except early stages 1-2, gypsum-board frame walls, and finishes.
- Walls from stage 1 must use system Wall family, not in-place forms.
- Elements from different corpuses/sections or fire compartments/fire sections must not be merged into one in-place element.
- Finishes are modeled depending on stage and requirement. At stage 3, wall finishes are modeled if customer requires areas including finishes. At stage 4, finishes are modeled fully.
- Window/door elements are modeled to exact overall dimensions. Profiles are conditional; 3D hardware is absent. Profile data may be issued in details if customer requires it.
- Window families must include `Зона открывания`; door families must include `Зона открывания` and may include `Зона открывания_Обратная сторона`.
- Door/window Mirror tools are categorically prohibited.
- Do not edit slab boundary to model openings; use special shaft/elevator/pit families.
- Facades are not cut to ground at stages 2-3; at stage 4 they consider vertical planning.
- Areas approved after stage 3 are fixed and used in later schedules/drawings through special room tags with prefix `П_`.

## Structures (KR)

- KR levels may be offset relative to AR to avoid constant offsets for slabs/beams.
- Reinforcement is modeled conditionally through nodes/areas, not every single bar in all cases.
- Reinforcement lengths may be issued in linear meters without individual specification positions for every bar length, with compromise wording if needed.
- Additional structures such as equipment foundations, metal stairs, and metal railings may be excluded from clash analysis.
- KM modeling scope must be explicitly described. Anything outside modeled scope is counted by PTO from issued details.
- KR model division should be highlighted early; the checklist flags `1 том = 1 модель` as a likely customer discussion point.

## MEP And Low-Current/Electrical

- Engineering calculations are not performed in Revit; tools such as Danfoss, Dialux, Excel, etc. remain valid.
- Supports for ducts, pipes, trays, boxes, and similar elements are generally not modeled; quantities are calculated from length/specifications.
- Slopes are modeled only for sewerage and drainage pipes, only for diameter above 25 mm and slope above 2%.
- Heating pipes of cross-linked polyethylene in floor screed are not checked for mutual intersections.
- Any intersection with beams, columns, capitals, or pylons must be agreed with KR regardless of pipe/duct size.
- For SS/EOM, distribution boxes, shield internal 3D content, boxes/tubes/corrugations may be non-modeled or calculated manually depending on item.
- Sensors, buttons, sockets, switches, phone devices, call/notification devices, communication and alarm devices generally do not participate in clash checks against adjacent sections.

## Openings

- Checklist baseline: PD openings are generally 500x500 and larger; RD openings are 300x300 in AR (or 200x200 in brick), 200x200 in KR. Sleeves are issued at all sizes. Exceptions must be agreed before design starts.
- For AR vertical openings on stages 3-4: openings 500x500 mm and larger.
- For AR vertical openings on stage 5: openings 200x200 mm and larger.
- Walls where openings are made on site, such as gypsum board or metal mesh, are not processed/modelled as openings.
- Stage 4 acceptable AR/KR vs IOS intersections: openings with one side/diameter under 500 mm may be acceptable.
- Stage 5 acceptable openings: KR walls/slabs/AR slabs/AR roof vs IOS under 200x200 or diameter under 200; AR walls vs IOS under 300x300 or diameter under 300; gypsum-board partitions vs IOS under 500x500 or diameter under 500.
- Opening clearance rule: vertical AR/KR structures require minimum 49 mm difference between IOS element section and opening section, excluding insulation.
- Horizontal AR/KR structures require minimum 99 mm for OV2 ducts and 49 mm for pipes/trays, excluding insulation.

## Clash Checks

- Stage 3 clash analysis is for all objects and focuses on project decisions; clashes that do not affect expert-review documentation may remain.
- Stage 4 clash analysis applies when preliminary quantities are delivered to customer and all error types must be resolved, except stated exclusions.
- Stage 5 clash analysis is for all objects and requires full coordination, with first/second-priority checks in the matrix.
- Typical Navisworks clash tolerance is 50 mm by matrix; duplication tolerance is 25 mm.
- Stage 3 checklist headline: PD 50 mm for all elements.
- RD checklist headline: 25 mm for all except finishes; finishes 50 mm.
- Introduce a non-reducible remainder if needed: 10 clashes per 5,000 m2.
- Duplicate elements from adjacent models are not corrected in duplicate copies; source elements are corrected in the model where they are developed.
- Valid/allowed clashes must be explicitly classified, not silently ignored.

## Responsibility For Clash Acceptance

- Collisions with escape routes, parking spaces, rooms, shafts, and openings must be agreed with AR.
- Door opening-zone clashes are accepted/defined by AR.
- Clashes with monolithic pylons, columns, beams, capitals are agreed with KR regardless of requested opening size.
- Non-pressure sewerage pipes VK have priority over all IOS sections during clash resolution; consult VK.
- Engineering equipment service/access zone intersections are decided by the engineers responsible for those elements.

## Monolith Matching

- Stages 3-4: check AR/KR monolithic structures that affect project decisions and documentation.
- Stage 5: check full AR/KR monolith correspondence except specified tolerances.
- Basic tolerance for monolith matching: 15 mm.
- Stair flight height mismatch tolerance: 20 mm.
- Stages 3-4 opening tolerance: 100 mm.
- Stage 5 fully mounted monolithic structure tolerance within one building: 50 mm.
- Stated acceptable mismatches include pit outer dimensions, stair/landing thickness when AR includes finish, missing installation openings, missing deformation joints, missing wall/stair-to-slab abutments, non-impacting AR self-intersections, monolithic thresholds in door families, and non-through shaft openings.

## Deliverables And Transfer

- PDF documentation should be generated from BIM-supporting CAD/BIM software; PDF version not lower than 1.7.
- RVT model division must match internal section 5.3 and be agreed with customer in BEP Appendix 1 / model list.
- Final RVT upload requires purging unused families, views, sheets, and schedules; model must not contain critical errors preventing future editing.
- Check common coordinates before transfer.
- Keep drawing views, sheets, sections, schemes, schedules, and working views placed on documentation sheets.
- Configure delivery 3D view: required elements visible, worksets checked, 3D section box off, high detail, DWG underlays off.
- NWF is for intermediate Navisworks exchange and should include search sets and clash tests according to BEP.
- NWD intermediate uploads may be without configured clash checks.
- Final stage NWD must contain saved and processed clash results with acceptable clashes confirmed according to chapter 6.
- Intermediate model upload: first upload no later than four calendar weeks from design start; subsequent uploads every two working weeks, unless a project platform rule replaces this.
- Final model acceptance and documentation issue dates should differ by no more than five working days.
