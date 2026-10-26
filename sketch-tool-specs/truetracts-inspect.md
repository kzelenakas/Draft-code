# TrueTracts Inspect (iPad sketch) — Behavior Spec
Source: 2 screen recordings ("ScreenRecording_10-26-2026 09-39-43_1", two versions: Clip A 0:45, Clip B 0:51), 640×448, 96 frames sampled at 1 fps + change detection · Extracted: 2026-10-26
Not analyzed: the two "09-33-16" recordings (6.5 MB and 7.7 MB) — too large for the Drive connector to transfer.

## Summary
A touch-first iPad app that pairs a floor-plan canvas with a right-side inspector. Area numbers are "calculated from sketch" and rolled up at four levels: property → structure → unit → level. Each level is split into above/below grade and standard/non-standard/unfinished. Both clips explore an already-drawn sketch. **No drawing, editing, or export is shown**, so this spec covers the data model, area rollups and navigation, not the drawing tools.

## Screen Map
| Region | Contents | Evidence |
|---|---|---|
| Top-left | Back `<`, green `+ Add ▾` menu | [seen A 0:00] |
| Top-right | Camera, eye/view menu `▾`, crosshair, sidebar toggle | [seen A 0:00] |
| Canvas | Grid; green-filled areas labeled "Structure · Level"; dimension lines on every side (whole feet); garage drawn in gray with a car icon and its sq ft | [seen A 0:00] |
| Bottom-center | Undo (active), Redo (disabled) | [seen A 0:00] |
| Right inspector | With nothing selected: "Property summary". With something selected: 3 tabs (Structure 🏠, Unit #, Level ≋) | [seen A 0:00, A 0:09] |

## Tool Inventory
| Tool / Command | Location | Behavior | Evidence |
|---|---|---|---|
| Add → Structure › | Add menu | Has a submenu (not opened) | [seen B 0:16] [not demonstrated] |
| Add → Level | Add menu | — | [not demonstrated] |
| Add → Property Unit | Add menu | — | [not demonstrated] |
| Add → Rooms › | Add menu | Has a submenu | [not demonstrated] |
| Add → Internal Spaces › | Add menu | Has a submenu | [not demonstrated] |
| Add → Garage / Carport › | Add menu | Has a submenu | [not demonstrated] |
| Add → Amenity › | Add menu | Has a submenu | [not demonstrated] |
| Add → Noncontinuous Area | Add menu | **Disabled**, with inline reason: "Only for a dwelling with one living unit (excluding ADUs)" | [seen B 0:16] |
| Add → Stamps › | Add menu | Has a submenu | [not demonstrated] |
| Camera | Top bar | Probably attaches photos | [inferred] [not demonstrated] |
| Eye ▾ | Top bar | Probably view/visibility options | [inferred] [not demonstrated] |
| Crosshair | Top bar | Probably recenters/fits the canvas | [inferred] [not demonstrated] |
| Sidebar toggle | Top bar | Probably shows/hides the inspector | [inferred] [not demonstrated] |
| Undo / Redo | Bottom | Redo is grayed out when nothing can be redone | [seen A 0:00] |
| Edit pencil | Each inspector header | Probably renames the structure, unit or level | [inferred] [not demonstrated] |

## Interaction Model
### Drawing
[not demonstrated] Everything was already drawn. The canvas shows whole-foot dimensions on every exterior segment, including notches (e.g. 10', 9', 20', 29', 30'). [seen A 0:00]

### Selection & navigation
- Selecting a structure gives it a dark green outline and switches the inspector from Property summary to the Structure tab. [seen A 0:08→0:09]. The input was probably a tap on the area. [inferred]
- The tabs switch between the Structure, Unit and Level views of the selected item. [seen A 0:35, 0:37, 0:41]
- One-finger drag pans the canvas, shown by a touch indicator. The whole sketch moves and the inspector stays put. [seen B 0:24–0:25]
- The Add menu is a popover that opens and closes on tap. It was opened about 8 times with no item picked. [seen B 0:16–0:50]

### Editing
- **Structure type** is a picker with Dwelling / Outbuilding / Clear. A ✓ marks the current value. [seen A 0:15]. It was closed without changing the value. [seen A 0:17]
- Construction method and Dwelling style are pickers showing "—" (empty). [seen A 0:09]
- On the Level tab, a **Finished / Unfinished** segmented toggle sets the whole level, with the note: "Rooms follow the level unless marked otherwise." [seen A 0:35]
- On the Unit tab, an **ADU No/Yes** toggle. [seen A 0:37]
- Moving, resizing and vertex editing are [not demonstrated].

### Snapping & precision
Dimensions display in whole feet and areas in whole sq ft. [seen] Input precision (inches / 0.1 ft) is [not demonstrated].

## Area & Calculation Model
### Hierarchy
Property → Structure(s) → Unit(s) → Level(s) → Rooms/spaces [seen A 0:00, 0:35–0:37; rooms inferred from the Add menu]. The summary subtitle reads "2 structures · 2 units · 4 levels". [seen A 0:00]

### Property summary (values on screen)
| Item | GBA | Finished | Unfinished |
|---|---|---|---|
| Main Building (1 unit · 3 levels) | 4,180 | 4,180 | 0 |
| — Main Level | | 1,400 | 0 |
| — Level 2 | | 1,400 | 0 |
| — Level 3 | | 1,380 | 0 |
| Structure 2 (1 unit · 1 level) | 1,400 | 1,000 | 0 |
| Car storage: Garage | | 400 sq ft | |

Headline: **Total finished area 5,180**, split as Above grade 4,180 · Below grade 1,000 · Unfinished 0. [seen A 0:00]
Arithmetic check: 1,400 + 1,400 + 1,380 = 4,180 ✓; 4,180 + 1,000 = 5,180 ✓; 1,400 − 400 = 1,000 ✓ [verified arithmetic]

### Structure-level measurements ("Calculated from sketch")
Gross building area 1,400 · Finished area 1,000 ("Excl. vehicle storage & ADU") · Unfinished area 0 ("Excl. vehicle storage & ADU") · Non-contiguous finished: No. [seen A 0:09]

### Level-level measurements
Level: "Level B1" · Grade level: "Below Grade" · Perimeter area 1,400 ("Not part of the finished total") · Finished 1,000 · Unfinished 0. [seen A 0:35]
So the level's perimeter area is gross. Finished area is derived from what's inside the perimeter: here, perimeter minus the garage. [inferred]

### Unit-level measurements
Above grade: Standard finished / Non-standard finished / Unfinished. Below grade: the same three. Then Total finished. Structure 2's unit: below-grade standard finished 1,000, everything else 0, Total finished 1,000. [seen A 0:37]

### Output / export
[not demonstrated]

## State & Feedback
| State | Visual cue | Evidence |
|---|---|---|
| Selected structure | Dark green outline on the canvas | [seen A 0:09] |
| Field filled | Green left bar on the row (Structure type, Level, Grade level) | [seen A 0:09, 0:35] |
| Field incomplete | Amber left bar (Construction method, Dwelling style, ADU) | [seen A 0:09, 0:37]. Meaning [inferred] |
| Tab has incomplete items | Amber dot on the tab icon (Structure, Unit) | [seen A 0:09]. Meaning [inferred] |
| Photo slots | "Front" / "Rear" tiles marked "Suggested" | [seen A 0:09] |
| Unavailable action | Grayed item with an inline reason | [seen B 0:16] |
| Garage | Gray fill + car icon + sq ft inside the shape | [seen A 0:00] |

## Workflows
### 1. Review area totals
1. Open the sketch. With nothing selected, the inspector shows the Property summary. [seen A 0:00]
2. Expand a structure row (chevron) to see per-level finished area. [seen A 0:00, Main Building already expanded]
3. Car storage is listed separately below the structures. [seen A 0:00]

### 2. Inspect and classify a structure
1. Select the structure → Structure tab. [seen A 0:09]
2. Set Structure type (Dwelling/Outbuilding). [seen A 0:15]
3. Set Construction method and Dwelling style. [not demonstrated]
4. Add front and rear photos. [not demonstrated]
5. Level tab: confirm the level name, grade level, and the finished toggle. [seen A 0:35]
6. Unit tab: set ADU No/Yes and review the standard/non-standard split. [seen A 0:37]

### 3. Add an element
1. Tap `+ Add` and pick a category. [seen B 0:16]. Completing an add is [not demonstrated].

Steps 1–2 and 5–8 of a full sketch workflow (new sketch, draw and close the perimeter, add levels, label, export) are **[not demonstrated]**.

## Action Log (condensed)
| Time | Input | Result |
|---|---|---|
| A 0:00–0:08 | — | Property summary shown |
| A 0:09 | Select Structure 2 | Outline + Structure tab |
| A 0:15 | Tap Structure type | Picker: Dwelling ✓ / Outbuilding / Clear |
| A 0:17 | Dismiss | Value unchanged |
| A 0:35 | Level tab | Level B1, Below Grade, Finished |
| A 0:37 | Unit tab | ADU toggle, grade/standard split |
| A 0:41 | Structure tab | Back to Structure |
| B 0:16 | Tap + Add | Add menu (Noncontinuous Area disabled) |
| B 0:24–0:25 | One-finger drag | Canvas pans |
| B 0:29–0:50 | Tap + Add repeatedly | Menu opens and closes, nothing picked |

## ANSI Z765 Observations
| Requirement | Status | Evidence |
|---|---|---|
| Precision to nearest inch / 0.1 ft | Not shown (display is whole feet) | [seen A 0:00] |
| Totals to whole sq ft | Supported | [seen A 0:00] |
| Below grade separated from above grade | Supported | Summary + Unit tab [seen A 0:00, 0:37] |
| Garage excluded from finished | Supported | "Excl. vehicle storage", 1,400 → 1,000 [seen A 0:09] |
| Non-standard finished tracked separately | Supported | Unit tab rows [seen A 0:37] |
| Ceiling height (7 ft / 5 ft) | Not shown | — |
| Openings to below / stairwells | Not shown | — |
| Exception explanation field | Not shown | — |
| Partially below grade (UAD 3.6) | Not shown. Only Above and Below grade values were visible | [seen A 0:00, 0:35] |

## Gaps & Open Questions
- **Drawing tools are the biggest gap.** The two 09-33-16 recordings probably cover them, but they need a different transfer route (see below).
- **Headline "Total finished area" (5,180) combines above and below grade.** It isn't GLA. Above grade (4,180) is the GLA candidate.
- **Multiple structures:** Structure 2 is a separate dwelling (ADU = No) whose finished area is added to the property total. Would an above-grade second structure add to the property's "Above grade" number? If so, a form filler that reads property-level above grade would overstate the main dwelling's GLA.
- **Data oddity in the sample:** Structure 2's only level is named "Level 1" but set as "Level B1 / Below Grade". Is that test data or a naming gap?
- Submenus for Structure, Rooms, Internal Spaces, Garage/Carport, Amenity and Stamps were never opened.
- No ceiling-height input, room counts (bed/bath), or export screen was seen.
- **UAD 3.6 partially below grade:** the only grade values seen were Above grade and Below Grade. UAD 3.6 requires partially below grade area to be reported separately, so check whether the Grade level picker offers a Partially Below option (the picker was never opened). If it doesn't, that's a gap to close.

## Implementation Notes (form filler)
**Data structure implied:** `Property{structures[]} → Structure{type, constructionMethod, dwellingStyle, photos[], units[]} → Unit{isADU, levels[]} → Level{name, gradeLevel, finishedDefault, perimeterArea, rooms[]}`, plus car storage as its own collection.

**Recommended field mapping (to be confirmed against the export data, not just the screen):**
| Form field (1004/UAD) | Source | Note |
|---|---|---|
| GLA | Main dwelling structure → unit (ADU = No) → above-grade standard finished | **Not** the property headline total |
| Non-standard finished above grade | Unit → above-grade non-standard | Required UAD 3.6 reporting. Also needs an ANSI explanation when > 0 |
| Partially below grade area | Level → Grade level = Partially Below (if offered) | Required UAD 3.6 reporting. Not seen in the demo |
| Basement area / finished | Below-grade level perimeter area / below-grade finished | |
| Garage / carport | Car storage list (type + sq ft) | Car count not seen |
| Design / style | Structure → Dwelling style | Empty in demo |
| ADU | Unit → ADU flag; separate structure or unit | Keep out of main GLA |
| Outbuildings | Structures with type = Outbuilding | |

**Validation rules worth copying:** block or flag empty required fields (the amber indicators), and disable impossible actions with an explanation instead of hiding them (the Noncontinuous Area pattern).

**Training notes:**
1. Read GLA from above grade, not "Total finished area."
2. Set the grade level on every level. A wrong grade level moves area between GLA and basement.
3. Fill in Construction method and Dwelling style before export (both empty in the demo).
4. Set the ADU flag correctly, since it removes area from finished totals.
