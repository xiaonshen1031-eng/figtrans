# Basic native scene helper

`scripts/render_scene.ps1` is an optional Windows PowerShell 5.1 + desktop Office executor. It does not recognize an image. The agent analyzes the reference and supplies explicit objects. It creates separate native PPTX/VSDX documents, saves and reopens them, checks object names and connector attachment, and exports previews. Visual review remains required.

```powershell
powershell -NoProfile -File scripts/render_scene.ps1 `
  -ScenePath examples/native-smoke.json -Target both `
  -OutputDir D:\figures\example
```

`-Target` accepts `pptx`, `visio`, `both`. `-ValidateOnly` checks the scene without starting Office; `-ExportEmf` also requests EMF previews. Existing outputs are refused. Use a new task output directory. With `both`, a reported backend error is recorded and the other target is attempted. Inspect each target's status in the receipt, retain successful artifacts, and recover only failed or unattempted targets into a fresh directory. An unreturned COM call remains pending; inspect the recorded stage and live process before deciding whether to continue waiting or recover. Do not rerun successful targets or assume a stale running receipt proves a live process.

## Visio startup and recovery

The helper first tries `GetActiveObject` to reuse a running Visio instance. Only the COM result meaning that no running instance is available triggers a cold startup. For that case, [start_visio_host.ps1](../scripts/start_visio_host.ps1) runs in a separate hidden Windows PowerShell 5.1 process and creates only the application and a uniquely titled task-owned blank guard. The drawing client waits for the worker to exit, then reattaches and verifies the recorded HWND and guard ID/title before drawing. Existing user documents are not closed.

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

## Audit

```powershell
python scripts/audit_editability.py D:\figures\example\figure.pptx --require-native --output D:\figures\example\pptx-audit.json
python scripts/audit_editability.py D:\figures\example\figure.vsdx --require-native --output D:\figures\example\visio-audit.json
```

The audit reports parts, text, native shapes, groups, connection references, pictures, OLE markers and relationship failures. `--require-native` rejects obvious image-only reconstruction; it cannot prove complete correspondence or detect every flattened panel. `STRUCTURE_PASS_REVIEW_REQUIRED` is intentionally not a visual or MathType certificate. Inspect the saved output in the target application and compare it against the source.
