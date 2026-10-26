# TOTAL iPad Sketcher — Behavior Spec
Source: 7 screen recordings, "ScreenRecording_10-25-2026 …" (17-56-09, 17-56-47, 17-57-46, 17-58-21, 17-59-08, 20-59-08, 21-01-11). Clips 17-59-08 and 20-59-08 are always written in full. 640×448, 2 min 20 s total, 280 frames sampled at 2 fps · Extracted: 2026-10-26
Clip IDs below use the recording time, e.g. `[56-09 0:07]` = clip 17-56-09 at 0:07.

## Summary
A **keyboard-first** sketcher. You type a length on an on-screen keypad and tap one of 8 direction arrows, and a wall is drawn from the current point. **Enter** closes the shape. A Properties popover then opens to assign the area name, area class, grade level and finished status. Its standout behavior is **area splitting**: drawing a line inside an existing area and pressing Enter carves out a new area and automatically reduces the parent. Pages are used as levels. Export and the GLA summary/calculator are not shown.

## Screen Map
| Region | Contents | Evidence |
|---|---|---|
| Top-left toolbar | 5 mode icons: draw (square), move (4-way arrows), select (hand), target/center, segment (square with corner handles) | [seen 56-09 0:00] |
| Top-right toolbar | Page/document, calculator, tag, stamp, T (text), ?, check/sign, a curved-arrow icon, expand | [seen 56-09 0:00]. Purposes [not demonstrated] |
| Left rail (blue) | List, sketch (active), image, pin, folder, page, mic, T, save | [seen 56-09 0:00] |
| Canvas | Purple grid; start point shown as a crosshair | [seen 56-09 0:00] |
| Keypad (right) | Entry field, 0–9, `.`, delete, 8 direction arrows, angle, enter, fit, undo, redo | [seen 56-09 0:00] |
| Offset indicator (top-right of canvas) | Distance and direction from the current point back to the start, e.g. `25' ↓  40' ←` | [seen 56-09 0:07] |
| Area toolbar (when an area is selected) | Copy, Define, Area Labels ON/OFF, Line Labels ON/OFF, Info, rotate ⟲ [45°] ⟳ | [seen 56-09 0:22] |
| Line toolbar (when a wall is selected) | Label ON/OFF, Angle ⟲ [270] ⟳, Length ◀ [20'] ▶ | [seen 20-59-08 0:04.5] |

## Tool Inventory
| Tool / Command | Behavior | Evidence |
|---|---|---|
| Number + arrow | Draws a wall of that length in that direction (orthogonal or 45° diagonal) from the current point | [seen 56-09 0:03.5–0:04, 25 ▲] |
| Arrow with empty entry | Draws the remaining offset to the start point or the opposite wall in that direction | [seen 01-11 0:10.5 (▲ closes 24'), 0:20.5 (▼ to opposite wall)]. Rule [inferred] |
| Number + arrow while on an existing wall | Moves the cursor along the wall without drawing a duplicate wall | [seen 01-11 0:18–0:19.5 (20 ►)]. Rule [inferred] |
| Enter | Closes the shape back to the start, or completes a split along existing walls. Then opens Properties | [seen 56-09 0:21; 56-47 0:07.5] |
| angle | Angle entry | [not demonstrated] |
| fit | Fits the drawing to the screen | [not demonstrated] |
| undo / redo | Undo becomes active after the first wall; redo stays grayed | [seen 56-09 0:04] |
| delete | Backspace on the entry field | [inferred] |
| Copy → Area / Outline | Two choices: "Area" copies area, walls, symbols and labels. "Outline" copies the perimeter only. Banner: "Area Copied, Long Press To Paste" | [seen 57-46 0:06.5–0:07] |
| Long press on canvas → Item Options → Paste Area | Pastes the copied shape | [seen 57-46 0:12] |
| Pages popover (Edit, +) | Lists Page 1…n with a ✓ on the current one. + adds a page | [seen 57-46 0:09–0:10.5; 17-59-08 0:02–0:05.5] |
| Define, Info, Area/Line Labels, Rotate 45° | Visible on the area toolbar | [not demonstrated] |
| Line Length ◀ ▶ / Angle ⟲ ⟳ | Edit a selected wall | [seen 20-59-08 0:04.5]. Use [inferred] |

## Interaction Model
### Drawing
- **Keyboard grammar:** `<length> <arrow>` per wall, finished with `enter`. Lengths are whole feet in every clip. A decimal key exists. [seen]
- Example, Level 1 [56-09]: 25 ▲, 40 ►, 32 ▼, 20 ◄, 7 ▲, then enter (closes the last 20 ◄). [seen 0:03–0:21]
- **Live preview:** while open, the shape shows a gray translucent fill from the last point back to the start. The active wall is green with green endpoints. A red right-angle marker appears at corners. [seen 56-09 0:07]
- **Offset indicator** always shows what's left to close, which acts as a built-in error check. [seen]

### Area splitting / carving (key behavior)
- Start on an existing vertex or wall, draw inside the area, and press Enter. The app closes the new shape along existing walls and **automatically subtracts it from the parent**:
  - Garage carved from Level 1: 13 ▲, 20 ►, enter → "2 Car Built-in 400 Sq ft". Level 1 goes from 1140 to 740. [seen 56-47 0:05–0:08]
  - Single line splitting the basement: 12 ▼, enter → new area 240. Basement 1 goes from 740 to 500. [seen 58-21 0:03.5–0:04.5]
  - Single line splitting Level 3: move 20 along the top wall, then ▼, enter → 2 Car Built-in 480 + Level 3 720. [seen 01-11 0:18–0:22]
- The new area's default name was "2 Car Built-in" or "Concrete Patio" before the user touched anything. [seen 56-47 0:08; 58-21 0:04.5]. Where the default comes from is [not demonstrated]: it may be the last used or the first name in the list.

### Editing
- **Wall edit:** select a wall to get the line toolbar (Label, Angle 270, Length 20'). Typing a number plus an arrow then **moved the wall**, and the adjoining walls stretched: Basement 2 went from 1040 to 640 sq ft. [seen 20-59-08 0:01–0:07]. The exact input sequence is [inferred].
- **Copy outline to another page:** select Level 1 (yellow fill = selected), Copy → Outline, Pages → + Page 2, long press → Paste Area, then rename it Basement 1. [seen 57-46 0:03–0:21]
- Move, rotate and vertex delete are [not demonstrated].

### Snapping & precision
Walls snap to 8 directions through the arrow keys. The cursor snaps to existing vertices and walls. [inferred] Precision finer than whole feet is [not demonstrated].

## Area & Calculation Model
### Properties popover (opens after every close)
- **Current Assignment:** Dwelling [seen]
- **Current Area Name:** pick list. Seen values: Level 1, Level 2, Level 3, Basement 1, 2 Car Built-in, Concrete Patio, Outbuilding (list heading only) [seen]
- **Common Levels** list for the area
- **Area class:** Standard / Non-Standard / Subtract From Living / No-Calc Area
- **Grade:** Above Grade / Partially Below / Below Grade
- **Finish:** Finished / Unfinished
- For garage and patio areas, only the class row was visible. The popover scrolls, so grade and finish rows may sit below it. [seen 56-47 0:08; 58-21 0:05]

### Auto-classification
Picking "Basement 1" automatically switched **Standard → Non-Standard** and **Above Grade → Below Grade**. [seen 57-46 0:17→0:18]. Garage and patio were already Non-Standard. [seen]
**Vocabulary warning:** TOTAL applied "Non-Standard" to the garage, the patio and the basement. That looks like "not counted in GLA," not the ANSI Z765 / UAD 3.6 meaning of non-standard *finished* area (finished space that doesn't meet the standard, such as low ceilings). UAD 3.6 requires reporting non-standard finished area as its own figure, so the form filler must not map TOTAL's "Non-Standard" class to that field without confirming what TOTAL exports. [inferred from the areas it was applied to]

### Areas created (all values read off screen; arithmetic checked)
| Area | Page | Class / Grade / Finish | Sq ft | Check |
|---|---|---|---|---|
| Level 1 (before garage) | 1 | Standard / Above / Finished | 1140 | 40×25 + 20×7 = 1140 ✓ |
| 2 Car Built-in | 1 | Non-Standard | 400 | 20×20 ✓ |
| Level 1 (after garage) | 1 | Standard / Above / Finished | 740 | 1140 − 400 ✓ |
| Basement 1 | 2 | Non-Standard / Below / Finished | 500 | 740 − 240 ✓ |
| Basement 1 (split) | 2 | Non-Standard / Below / **Unfinished** | 240 | 20×12 ✓ |
| Level 2 | 3 | Standard / Above / Finished | 1000 | 40×25 ✓ |
| Basement 2 | (separate sketch) | — | 1040 → 640 after wall edit | 40×20+20×12 ✓; 20×20+20×12 ✓ |
| Level 3 | (separate sketch) | Standard / Above / Finished | 1200 → 720 | 50×24 ✓; 30×24 ✓ |
| 2 Car Built-in | (separate sketch) | Non-Standard | 480 | 20×24 ✓ |

[verified arithmetic], no mismatches.

### Multi-level handling
Each level was drawn on its own **page** (Page 1 = Level 1 + garage, Page 2 = basement, Page 3 = Level 2). [seen 57-46, 17-59-08] Whether pages stack or align vertically is [not demonstrated].

### Rounding, totals, export
The area label shows whole sq ft. [seen] The GLA summary, calculator and form transfer are [not demonstrated].

## State & Feedback
| State | Visual cue | Evidence |
|---|---|---|
| Shape open | Gray preview fill to the start point + offset indicator | [seen 56-09 0:07] |
| Active wall / cursor | Green line, green endpoint, crosshair | [seen] |
| Pressed key | Blue highlight on the keypad key | [seen 56-09 0:04, 0:21] |
| Area selected | Yellow fill + area toolbar | [seen 56-09 0:22] |
| Wall selected | Green wall + line toolbar | [seen 20-59-08 0:04.5] |
| Clipboard full | Top banner "Area Copied, Long Press To Paste" | [seen 57-46 0:07] |
| Vertex / wall midpoint | Red dot / blue dot | [seen 57-46 0:03] |

## Workflows
### 1. Draw the main level
1. On a blank page, type a length, tap an arrow, and repeat for each wall. [seen 56-09]
2. Watch the offset indicator. Press enter to close. [seen 0:21]
3. In Properties, set the name to Level 1, Standard / Above Grade / Finished, then Ok. [seen 0:21.5–0:23.5]

### 2. Carve a garage or patio from a level
1. Start at a vertex on the level, draw the walls inside it, and press enter. [seen 56-47]
2. Name it, e.g. "2 Car Built-in" (Non-Standard). The parent level is reduced automatically. [seen]

### 3. Add a basement from the main level's outline
1. Select Level 1 → Copy → Outline. [seen 57-46 0:06.5]
2. Pages → + → Page 2. Long press → Paste Area. [seen 0:09–0:13]
3. Name it Basement 1. Non-Standard / Below Grade are set automatically. Set Finished. [seen 0:17–0:20]
4. Optionally split it with a line and mark one part Unfinished. [seen 58-21]

### 4. Add an upper level
Pages → + → Page 3, draw, and name it Level 2 (Standard / Above / Finished). [seen 17-59-08]

### 5. Fix a wall
Select the wall, then change its length or move it. The area recalculates. [seen 20-59-08]

Not demonstrated: new-file setup, symbols (doors, windows, stairs), the Define/Info dialogs, GLA summary, and transfer to the form.

## ANSI Z765 Observations
| Requirement | Status | Evidence |
|---|---|---|
| Precision to inch / 0.1 ft | Not shown (decimal key exists, only whole feet used) | — |
| Totals to whole sq ft | Supported | Area labels |
| Below grade separated | Supported (Below Grade auto-set for basement) | [57-46 0:18] |
| Garage excluded from GLA | Supported (Non-Standard, subtracted from level) | [56-47 0:08] |
| "Partially Below" grade option | Present. ANSI doesn't separate partially below grade from below grade, but UAD 3.6 requires it to be reported, so this should map straight through to UAD 3.6 (and stay out of above-grade GLA) | [56-09 0:22] |
| Non-standard finished (ANSI / UAD 3.6 sense) | Not shown. TOTAL's "Non-Standard" class appears to be a different concept (see warning above). UAD 3.6 requires this figure, so its source has to be confirmed | — |
| Ceiling height / stairwells / exception explanation | Not shown | — |

## Gaps & Open Questions
- **GLA total and form transfer:** we need footage of the calculator/summary and of what lands on the 1004 page. This decides the form-filler mapping.
- **What "Subtract From Living" and "No-Calc Area" do to totals** wasn't demonstrated.
- **Default naming of carved areas:** is it the last used name or a smart guess?
- **Partially Below and non-standard finished in the export:** both are required UAD 3.6 reporting. Does TOTAL's export carry Partially Below as its own value, and does it export ANSI-sense non-standard finished area anywhere?
- **Angle entry** and non-orthogonal walls aren't shown.

## Implementation Notes — what to bring into our form filler / sketch
Compared with TrueTracts Inspect (see `truetracts-inspect.md`), TOTAL's strengths are in **drawing speed**:
1. **Keypad + 8-way arrow grammar** (`length + direction`), with an empty arrow meaning "draw to the end." It works one-handed and makes exterior measurement entry fast. **Recommend adopting it.**
2. **Live "offset to close" indicator.** It's a free error check while measuring in the field. **Adopt.**
3. **Carve/split with automatic parent subtraction.** Garage, patio and finished/unfinished splits each take 2–3 keystrokes and the totals stay consistent. **Adopt.** This is also the biggest error-reducer for GLA.
4. **Copy Outline → paste to another level** for stacked floors and basements. **Adopt.** TrueTracts' level stack could do this as "duplicate perimeter to new level."
5. **Name-driven auto-classification** (Basement → Below Grade). **Adopt, but** keep the ANSI / UAD 3.6 vocabulary TrueTracts already uses, and never reuse TOTAL's "Non-Standard" label for non-GLA areas.

Where TrueTracts is already ahead (keep it): explicit property → structure → unit → level hierarchy, ADU handling, ANSI standard vs non-standard finished split, and completeness indicators.

**Data model for import/compatibility:** `Page{areas[]}`, `Area{name, assignment, levelRef, class: Standard|NonStandard|SubtractFromLiving|NoCalc, grade: Above|PartiallyBelow|Below, finish, polygon[], parentArea?}`.
**Mapping rule (proposed):**
- GLA = Σ Standard + Above Grade + Finished.
- Below-grade area = Below Grade. Partially Below is reported separately as its own UAD 3.6 value, and neither counts toward above-grade GLA.
- Non-standard finished area (UAD 3.6) has to come from a source that means ANSI non-standard, not from TOTAL's "Non-Standard" class.
- Garage, patio and porch come from Non-Standard areas, keyed by name.

**Training notes:**
1. Watch the offset indicator before pressing enter. If it isn't 0'/0', the shape isn't closed the way you think.
2. After carving a garage, confirm the parent level shrank.
3. After pasting an outline, rename it. It keeps the old label ("Level 1") until you do. [seen 57-46 0:13–0:17]
4. Check the grade and finish on basement splits. The unfinished portion has to be set by hand.
