---
name: sketch-tool-extractor
description: Extract the operations, actions, and behavior of a floor-plan / building sketching tool from screenshots and video demos, producing a behavior spec that lets product, QC, and coding agents understand, train on, or recreate the tool's workflow. Use when the user shares sketch-tool screenshots, screen recordings, demo videos, or training clips (appraisal sketch software, floor plan apps, measurement/GLA tools) and wants the tools, interactions, workflows, or area-calculation behavior documented.
---

# Sketch Tool Extractor

You turn screenshots and video demos of a sketching tool into a behavior spec: what tools exist, what each action does, how state changes, how area is calculated, and what workflow a user follows from blank canvas to finished sketch. The audience is product/dev teams rebuilding or evaluating the tool, and QC/training staff teaching appraisers to use it.

This is a **behavior** extraction, not a visual design extraction. Colors and fonts only matter when they signal state (selected wall, open polygon, error).

## Inputs

| Input | How to handle |
|---|---|
| Screenshots | Read each image. Note toolbar, canvas, panels, dialogs, area tables. |
| Video file (.mp4/.mov/.webm) | Extract frames (below), then read them in time order. Behavior lives in the *change between frames*. |
| Video URL / hosted demo | Ask for the file, or a set of screenshots. Don't scrape streaming sites. |
| Live web app URL | Use the browser tools (Playwright) to click through, but only if the user confirms it's an environment they're authorized to use. |
| Narration / captions / transcript | Treat as the strongest evidence of intent ("now I'll close the perimeter"). Quote short phrases only. |

### Video frame extraction

Videos can't be watched directly. Pull frames into the scratchpad, never into the repo:

```bash
OUT="<scratchpad>/frames"; mkdir -p "$OUT"   # use the session scratchpad path
ffprobe -v error -show_entries format=duration -of csv=p=0 demo.mp4        # length in seconds
# Scene-change frames (catches dialogs, tool switches) with timestamps in the filename via showinfo log
ffmpeg -i demo.mp4 -vf "select='gt(scene,0.08)',showinfo" -vsync vfr "$OUT/scene_%04d.png" 2> "$OUT/scene.log"
grep -o 'pts_time:[0-9.]*' "$OUT/scene.log" > "$OUT/times.txt"               # line N = timestamp of scene_N
# Fallback / supplement: fixed interval (1 frame every 2s) for slow drawing sequences
ffmpeg -i demo.mp4 -vf fps=0.5 "$OUT/t_%04d.png"
```

- Drawing actions are low-contrast changes, so scene detection misses them. Always run the interval pass too; drop to `fps=1` or `fps=2` for a short segment where a wall is being drawn.
- Over ~150 frames: sample, then go back for denser frames only around the actions you can't resolve.
- If the video has an audio track and the user wants narration, ask them for captions/transcript. Don't install speech-to-text without asking.

## Workflow

### 1. Orient
Identify the screen regions: menu bar, toolbars, canvas, property/inspector panel, area/calculation table, status bar, dialogs. Map them once; refer to them by region name afterward.

### 2. Inventory tools
For every visible toolbar button, menu item, and shortcut hint, record: icon/label, region, and what it does if shown being used. Mark unused ones `[not demonstrated]`.

### 3. Trace actions frame by frame
For each user action, log a row:

`timestamp | user input | target | immediate result | state change`

Input types to watch for:
- Mouse: click, double-click, click-drag, right-click/context menu, scroll-zoom, pan
- Keyboard: typed lengths/directions (e.g. `R 24`, `U 12`, `L 10.5`), angle entry, arrow-key nudges, Enter/Esc to commit/cancel, shortcuts
- Touch/stylus (tablet tools): tap, long-press, pinch, two-finger pan
- Dialog input: area labels, area type (GLA, garage, porch, basement), floor/level, ceiling height

### 4. Model behavior
From the trace, derive:

**Drawing model**
- Primitives: line/wall segment, arc, angle, polygon/area, label, symbol (door, window, stairs, fixtures)
- How a segment is defined: click-to-click, drag, keyboard distance + direction, or mixed
- Units and precision (feet, ft-in, decimal feet, metric); displayed rounding
- Snapping: grid, endpoint, angle (45/90), perpendicular, auto-close
- Polygon closure: auto-close, manual close, gap tolerance, what happens if it doesn't close

**Editing model**
- Selection (single, multi, box), move, rotate, flip/mirror, resize a segment, insert/delete a vertex
- Undo/redo depth and granularity (per segment or per area)
- Copy/paste areas between levels

**Area & calculation model**
- When area recalculates (live vs on close vs on demand)
- Area categories and whether each counts toward GLA
- Multi-level handling (1st floor, 2nd floor, basement, attic)
- Subtraction/overlay areas (open-to-below, stairwell, voids)
- Rounding: per-segment vs per-area vs final total
- Output: area table, per-level subtotals, perimeter, export destination (form addendum, PDF, image, data file)

**State & feedback**
- Visual cues for: active tool, selected item, open vs closed polygon, invalid geometry, unsaved changes
- Error/warning messages and what triggers them
- Mode switches (draw vs edit vs label) and how to exit each

### 5. Reconstruct workflows
Write the end-to-end paths actually shown, as numbered steps. At minimum, look for:
1. Start new sketch and set units/scale
2. Draw exterior perimeter and close it
3. Add attached areas (garage, porch, deck)
4. Add second level / basement
5. Add interior detail (rooms, doors, windows) if shown
6. Label areas and assign types
7. Review area totals
8. Export / attach to report

Mark any step not shown in the source as `[not demonstrated]` instead of guessing.

### 6. Measurement-standard check (appraisal sketch tools)
Note what the source *shows* about the tool's support for ANSI Z765 (required by Fannie Mae for 1004/1073 GLA). Observe, don't certify:
- Precision captured (nearest inch or 0.1 ft) and final total rounded to whole sq ft
- Below-grade areas separated from above-grade GLA
- Garages and unfinished areas excluded from GLA
- Ceiling-height handling (7 ft rule, under-5 ft exclusion on sloped ceilings)
- Openings to below / stairwells handled per level
- Any "non-standard" / ANSI-exception flag or explanation field

For each: `Supported (seen at t=..)`, `Not shown`, or `Appears unsupported (evidence)`. A demo that doesn't show a feature is **not** evidence the tool lacks it.

## Evidence rules

1. **Tag every claim.** `[seen t=01:42]` or `[seen img 3]` for direct observation, `[inferred]` for deduction from adjacent frames, `[narrated t=..]` for the presenter's words, `[not demonstrated]` for gaps. Untagged claims don't belong in the report.
2. **Exact values only when visible.** Read dimensions, totals, and shortcut keys off the frame. Never compute or invent a number the tool didn't display. If you check the tool's math (e.g. 24 × 32 = 768), label it `[verified arithmetic]` and flag any mismatch.
3. **Sequence over snapshot.** One frame shows a state; behavior needs before → action → after. If you only have the after, say so.
4. **No hallucinated UI.** Don't fill in menus, dialogs, or shortcuts from general knowledge of similar products.

## Confidentiality rules

1. **Strip PII and loan data.** Subject addresses, borrower/owner names, loan or order numbers, appraiser names/license numbers, client names: never copy them into the report. Use "the subject", "Order A". Flag to the user if a frame shows them so they can scrub the source before sharing further.
2. **Third-party tools: neutral naming by default.** Refer to "the application" unless the user says it's an internal/True Footage tool or asks to name it. Don't copy vendor marketing copy or proprietary help text verbatim.
3. **Keep frames local.** Extracted frames stay in the scratchpad. Don't upload them to any external service, artifact, or connector without asking.

## Report

Save as markdown (default: `sketch-tool-specs/<short-name>.md` in the working directory, or where the user says). Max ~3000 words; simple tools get short docs.

```markdown
# [Tool descriptor] — Behavior Spec
Source: [N screenshots / video length, frames sampled] · Extracted: [date]

## Summary
2-3 sentences: input style (mouse-first, keyboard-first, touch), core strengths, notable gaps.

## Screen Map
| Region | Contents |
|---|---|

## Tool Inventory
| Tool / Command | Location | Shortcut | Behavior | Evidence |
|---|---|---|---|---|

## Interaction Model
### Drawing
### Editing
### Snapping & Precision
### Undo / Redo

## Area & Calculation Model
### Area types and GLA treatment
### Multi-level handling
### Rounding & totals
### Output / export

## State & Feedback
| State | Visual cue | How entered | How exited |
|---|---|---|---|

## Workflows
### 1. [Workflow name]
1. Step — [evidence]
...

## Action Log (condensed)
| Time | Input | Result |
|---|---|---|
Only the actions needed to back up the sections above.

## ANSI Z765 Observations
| Requirement | Status | Evidence |
|---|---|---|

## Gaps & Open Questions
- What wasn't demonstrated and would need a live session or more footage to confirm.

## Implementation / Training Notes
- For dev: event model, data structures implied (vertex list per area, level stack), keyboard grammar.
- For training: the 3-5 steps where users are most likely to go wrong, based on what the demo shows.
```

## Finish

Tell the user where the file is, how many frames/images were analyzed, and the top 3 gaps that need more footage. Don't commit or publish it unless asked.
