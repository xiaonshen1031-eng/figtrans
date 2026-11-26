# Native scene and PPTX refinement helpers

`scripts/render_scene.ps1` is an optional Windows PowerShell 5.1 + desktop Office executor. The agent reads the reference and supplies explicit objects. The helper saves and reopens separate native PPTX/VSDX files. For each requested attached connector it compares both endpoints with the specified target shapes and records their names and IDs in `verified_connector_targets`. Inspect the exported previews for routing, arrowheads and spacing.

```powershell
powershell -NoProfile -File scripts/render_scene.ps1 `
  -ScenePath examples/native-smoke.json -Target both `
  -OutputDir D:\figures\example
```

`-Target` accepts `pptx`, `visio`, `both`. `-ValidateOnly` checks the scene without starting Office; `-ExportEmf` also requests EMF previews. Existing outputs are refused. Use a new task output directory. With `both`, a reported backend error is recorded and the other target is attempted. Inspect each target's status in the receipt, retain successful artifacts, and recover only failed or unattempted targets into a fresh directory. An unreturned COM call remains pending; inspect the recorded stage and live process before deciding whether to continue waiting or recover. Do not rerun successful targets or assume a stale running receipt proves a live process.

## Visio startup and recovery

The helper first tries `GetActiveObject` to reuse a running Visio instance. Only the COM result meaning that no running instance is available triggers a cold startup. In that case [start_visio_host.ps1](../scripts/start_visio_host.ps1) runs in a separate hidden Windows PowerShell 5.1 process and creates the application and a uniquely titled task-owned blank guard. After the worker exits, the drawing client verifies the Windows process ID and the exact guard ID and title. Both window handles are recorded so a change in the startup window can be distinguished from a change of process.

The startup-worker wait is bounded to 20 seconds. The main `<name>_render_receipt.json` records the worker PID and receipt path; `<name>_visio_host.json` records startup stages, worker and Visio PIDs, HWND, guard identity and any error. On timeout the worker may still be running: inspect its process and both receipts. The helper does not automatically kill or restart either process. This bound does not interrupt subsequent COM calls that never return.

In observed tests, same-client Visio creation followed by save-close-reopen stalled with both `Documents.Open` and `Documents.OpenEx`; the internal cause remains unknown. The separate-initializer/reattachment strategy completed saving, reopening and PNG/EMF export in the tested sequence. The helper now automates that strategy, but each run still needs its own completion evidence. It currently reopens only its own newly generated macro-free VSDX through `Documents.Open`; this is not an OpenEx-only diagnosis or a reason to weaken controls for arbitrary input files. See [the Visio route](visio.md) for verification and document-ownership rules.

## Scene fields

```json
{
  "name": "figure",
  "canvas": {"width": 720, "height": 420},
  "objects": [
    {"id": "input", "type": "rect", "x": 40, "y": 80, "w": 160, "h": 60,
     "text": "Input", "fill": "#EDF4FA", "stroke": "#557994"},
    {"id": "output", "type": "rect", "x": 400, "y": 80, "w": 160, "h": 60,
     "text": "Output", "fill": "#F2F0FA", "stroke": "#78659B"},
    {"id": "flow", "type": "connector", "from": "input", "to": "output",
     "from_side": "right", "to_side": "left", "route": "straight",
     "stroke": "#374F68", "stroke_width": 3, "head": "triangle", "head_size": "small"}
  ]
}
```

Coordinates are **points**, top-left origin. Canvas dimensions are 72–4032 pt; select a physical aspect ratio appropriate to the reference. `name` is a simple ASCII filename stem, IDs are unique ASCII semantic names starting with a letter. `sft_canvas_background` is reserved for the helper's Visio background and cannot be a scene object ID. Text supports Unicode.

| Field / type | Meaning |
| --- | --- |
| `rect`, `roundRect`, `ellipse`, `text` | `x,y,w,h`; optional text |
| `line` | `points: [[x1,y1],[x2,y2]]` |
| `polyline` | Two or more points; one continuous editable path, fixed geometry |
| `connector` | `from,to` node IDs; `from_side,to_side`: top/bottom/left/right; `route`: straight/elbow |
| `fill`, `stroke` | `#RRGGBB` or `none` |
| `stroke_width` | Positive point width; not tied to arrowhead size |
| `head` | `none` or `triangle`; terminal end only |
| `head_size` | `small`, `medium`, `large` |
| `head_width`, `head_length` | Independent PowerPoint small/medium/large settings; default to head_size |
| `text_color`, `font`, `font_size`, `bold` | Text style; default Times New Roman, 16 pt |
| `align`, `text_margin` | left/center/right and margin in points |

Visio's supplied adapter uses native overall `EndArrowSize`; independent width/length fields cannot be transferred directly and produce a warning. If the required head proportions need custom native geometry, use a task-specific Visio adapter and inspect its attachment and joins rather than silently claiming parity.

In Visio, the helper first creates a native white, opaque, borderless full-canvas rectangle named `sft_canvas_background`, then creates scene nodes and paths. The background preserves the canvas and outer margins in drawing-bounds exports; it is not a picture. PowerPoint uses its slide background and adds no corresponding shape. Thus a 20-object scene has 20 native PowerPoint objects and 21 native Visio objects. The receipt distinguishes `scene_objects`, `canvas_background_objects` and `native_objects`.

Within the scene, the helper orders nodes first and paths afterward. It intentionally lacks grouping, imported icons, charts, formula OLE, clipping masks and arbitrary z-order. Use the actual application's APIs or extend a **task-local** script for those needs. Unknown object types are errors, not silently dropped or rasterized. A successful smoke example verifies basic primitives, not every kind of scientific figure.

The MLP/KAN conversion used a task-local adapter for italic text, dashed rules, explicit z-order and native groups, then replaced fragmented Visio formulas with character-formatted text. Those extensions are not automatically available in this basic helper. Do not pass `groups`, `z_index`, `italic`, `dash` or `runs` expecting an effect here; validate required fields in the chosen adapter and reject unsupported formatting. Follow [the native Visio formula route](equations.md#native-visio-formulas) and [group/attachment checks](visio.md#groups-and-attachment) for those operations. Keep scene specifications and task scripts consistent with the final native file after repairs.

## Refine an existing native PPTX

Use [refine_pptx.ps1](../scripts/refine_pptx.ps1) with desktop PowerPoint on Windows for targeted work on a native PPTX. Start with the current accepted deck. Prepare an explicit refinement JSON from the reference and the current render. The [base scene](../examples/pptx-fidelity-base.json) and [refinement spec](../examples/pptx-fidelity.json) provide a small example. Its formula shift illustrates how to enter a measurement. Run `-ValidateOnly` to check the spec against the source slide and planned outputs before opening PowerPoint. Omit the switch to apply the edits to a copy in a fresh output directory.

```powershell
powershell -NoProfile -File scripts/render_scene.ps1 `
  -ScenePath examples/pptx-fidelity-base.json -Target pptx -OutputDir D:\figures\base
powershell -NoProfile -File scripts/refine_pptx.ps1 `
  -SourceDeck D:\figures\base\pptx_fidelity_base.pptx `
  -SpecPath examples/pptx-fidelity.json `
  -OutputDir D:\figures\refined -ValidateOnly
```

| Field | Purpose |
| --- | --- |
| `name`, `slide` | Output stem and 1-based slide number. |
| `receiver_groups` | `[{"id": "receiver_group", "members": ["receiver", "formula"]}]`. Name each new group and list distinct existing top-level native members. Attached connector shafts stay outside the receiving group. |
| `arrowheads` | Specify an existing, two-ended native `connector`, its `receiver_group`, a cardinal `direction` (`right`, `left`, `down`, `up`), absolute page `tip` and local `rear_offset`. The group must contain the original destination. The helper attaches the shaft to the new rear port before grouping the port and filled head with the receiver. |
| `path` in an arrowhead | Initial `start` and 2–64 ordered `segments`. Each segment has `type` (`line` or `cubic`) and `to`. Cubic segments also have `control1` and `control2`. Include `[0, 0]` at the tip and close the outline by returning to `start`. Optional arrowhead fields `fill` and `port_size` default to black and 0.25 pt. |
| `max_head_extent_pt` | Optional path extent limit. Defaults to one third of the shorter slide side and can be raised to that side's length. Choose a scale that fits the intended placement. |
| `formula_placements` | Each entry names an existing top-level native text shape by `id`. For a vertical correction, supply both `observed_baseline_pt` and `target_baseline_pt`. Use `delta_x_pt` for horizontal spacing, independently or with the baseline pair. Record the glyph measurement or layout reference in `evidence`. The helper retains the native text. Confirm the resulting glyph position in the export. |
| `edit_check` | Optional `move` group ID and `delta: [dx, dy]`. Set `text: {"id": "receiver", "value": "Test label"}` to also test an edit to a child label. The helper performs these edits on a separate copy. |

Coordinates use points. A `tip` is measured in the slide's top-left coordinate system. In the arrowhead-local frame use the tip as `[0, 0]` and positive x as the forward direction. Specify the outline and rear port offset in that same frame. Match the rear overlap, approach direction and clearance to the source. See [PowerPoint custom heads](powerpoint.md#custom-heads-and-semantic-receiver-groups) for choosing and verifying this geometry, and [formula placement](equations.md#fitting-editable-powerpoint-math-to-a-reference) for measuring baseline corrections.

The result includes `<name>.pptx`, the reopened PowerPoint preview `<name>_pptx.png` and `<name>_refine_receipt.json`. With `edit_check`, the script also saves `<name>_edited.pptx` and `<name>_edited_pptx.png`. Check for `native_saved_and_reopened` in the receipt. Inspect the result in PowerPoint and compare full and detail crops against the reference. Compare the main and edit-check PPTX files to test the movement recorded in the receipt.

```powershell
python scripts/check_pptx_roundtrip.py `
  D:\figures\refined\pptx_fidelity.pptx D:\figures\refined\pptx_fidelity_edited.pptx `
  --moved-name receiver_unit --expected-delta-pt 12 0 `
  --expect-text "receiver_label=Edited receiver" --output D:\figures\refined\proof.json
```

The independent comparison checks the saved package's group/member identities, connector targets and connection sites, and endpoint displacement in slide coordinates. `--expected-delta-pt` checks the planned translation and `--expect-text` checks the edited label. The report gives `CHECKS_PASS_REVIEW_REQUIRED` for successful numeric checks and `FAIL` for a mismatch. Use the reopened render to check the actual arrow join and formula appearance. Use the [Visio route](visio.md) for VSDX, with its own grouping, ShapeSheet connections and export checks.

## Audit

```powershell
python scripts/audit_editability.py D:\figures\example\figure.pptx --require-native --output D:\figures\example\pptx-audit.json
python scripts/audit_editability.py D:\figures\example\figure.vsdx --require-native --output D:\figures\example\visio-audit.json
```

The audit reports parts, text, native shapes, groups, connection references, pictures, OLE markers and relationship failures. The reserved blank Visio canvas backing does not supply evidence of editable reconstruction. Duplicate object IDs, duplicate endpoint declarations and attachments to the connector itself are structural failures. Fixed native paths and loops whose two ends attach to the same node are valid. `--require-native` checks for native evidence per page. Finish by inspecting the saved output in the target application and comparing it against the source.
