# Native Visio route

Use for VSDX or both. Detect installed Windows Visio and COM registration before promising a `.vsdx`. Use its native object model where available. Creating an OOXML-like ZIP or renaming another format is not a valid Visio conversion. If Visio is unavailable, retain the shared reconstruction specification and state which requested format remains incomplete.

## Geometry

Visio page coordinates use inches and a bottom-left origin. Keep task coordinates in top-left points, then convert exactly once:

```text
x_visio = x_pt / 72
y_visio = (canvas_height_pt - y_pt) / 72
rectangle = [x/72, (H-y-h)/72, (x+w)/72, (H-y)/72]
```

Use panel-local geometry before the page conversion. Set PageWidth and PageHeight explicitly. Keep page size and print/export extents consistent. Shape cells should use explicit unit strings (`pt`, `in`) to avoid locale-dependent scale.

For newly built scenes, the helper adds a native white, opaque, borderless rectangle named `sft_canvas_background` behind all scene objects. Its bounds cover the full canvas so drawing-bounds PNG/EMF export retains the intended outer margins. This is an editable background shape, not a flattened figure. The name is reserved: a 20-object scene has 21 native Visio shapes, including this background. Keep `scene_objects` and `canvas_background_objects` separate when comparing the receipt with the saved file.

## Native content

- Draw containers and symbols using native rectangle/oval/freeform/geometry shapes. Text belongs in native shapes with explicit margins, alignment and font lookup; font ID 0 is not a portable guarantee of Times New Roman.
- For meaningful dynamic relations, glue begin/end connector cells to explicit connection points or Shape ports. Inspect the resulting `Connects` relationships. A DrawLine with matching coordinates alone is not attached.
- For a required fixed multi-bend route, use a single geometry path with consistent stroke and terminal arrow. It can be editable geometry without being a dynamic connector; state that distinction and re-route after node movement.
- Set `LineWeight`, `BeginArrow` / `EndArrow`, and available arrow-size cells separately. Validate the renderer: Visio arrow-size enums are not numerically equivalent to PowerPoint's.
- Use semantic NameU identifiers and meaningful groups. Keep independent text/legend objects editable inside groups. Check nested bounds and clipped children after grouping.
- Import only atomic icons/photos or verified equation objects. A page-sized imported SVG/EMF remains imported artwork, even if it appears sharp.

For formulas, read [equations.md](equations.md). A Visio ForeignData element may contain an OLE object or a picture; count and inspect the actual relationship and payload rather than treating every ForeignData as MathType.

## Groups and attachment

For repeated edge-function plots, make the frame and native curve a logical group. The verified construction attached network connectors to each frame first, restored the intended z-order, then grouped each frame with its curve. Grouping kept both objects independently editable and moving the complete group preserved the connector's attachment to its inner frame. Verify this behavior for the actual host and grouping operation; replacing a frame or ungrouping/regrouping can change IDs or routing.

Use stable `NameU` identifiers and retain actual endpoint identities. Visio `BeginX` / `EndX` cells can be glued with `GlueToPos(target,u,v)`; normalized coordinates use a bottom-left origin, so top-center is `(0.5,1)` and bottom-center is `(0.5,0)`. These coordinates are not PowerPoint connection-site numbers. A circle center is not automatically an available PowerPoint connection site; use supported sites or a verified custom port instead of leaving a coordinate-only endpoint and reporting it as attached.

Check attachment with both structure and interaction:

- Recurse through native groups (`Shape.Type=2`, `Shape.Shapes`). Top-level `Shapes.Count` no longer counts all leaf objects. Record groups, leaves, background shapes and connector identities separately; derive expectations from the current specification rather than a fixed example count.
- Inspect each relationship's `FromSheet` / `FromCell` and `ToSheet` / `ToCell`, including references to grouped child shapes. Confirm both `BeginX` and `EndX` point to the intended objects. Two records or a matching total alone cannot establish correct targets.
- On a disposable copy, move a representative node and compare the attached endpoint coordinates before/after. Also move an entire function-frame group and verify the endpoint follows its child frame while the curve moves with the group. Save and reopen the copy; check relationships and a representative text edit persist. Keep the final source/layout unchanged.
- Keep edges behind node dots and plot frames. An endpoint that merely touches a node in the render still needs the above attachment checks. A fixed decorative bracket or curve can remain coordinate geometry; do not classify it as a dynamic network relation.

The MLP/KAN example preserved 76 attached connector segments in each format through native grouping, save/reopen and representative node/group movement. That is a verified example, not a required count for another figure. See [the cross-host check](powerpoint.md#save-and-inspect) for PowerPoint.

## Native miniature curves

Preserve each distinct schematic function curve as editable geometry, rather than substituting one generic curve throughout the network. Tracing the supplied schematic stroke is acceptable for visual reconstruction; retain the source/spec and do not present the trace as measured experimental data. Group the curve with its frame after checking for missing branches, clipped endpoints and incorrect z-order.

A Visio native path may pack many vertices into a `PolylineTo` or NURBS formula. Geometry-row count is not vertex count: a short `MoveTo` + `PolylineTo` representation can still contain the complete editable curve. Inspect the path row types and their coordinate/formula payload, along with the rendered result; do not reject it merely for having fewer rows than the sampled points. A stalled indexed COM geometry query is not evidence that the curve became a picture. Inspect its saved geometry or use a verified typed accessor before repeating the query or reconstructing the drawing.

## Implementation and failures

The supplied scene helper creates basic native objects and uses explicit errors. For a complex task, copy/adapt task-specific code in the workspace; do not edit the installed helper for a single paper. Do not silently ignore COM cell errors, attach failures or unsupported types.

A targeted package edit may be useful for a known existing shape property, but avoid speculative whole-file XML construction. Preserve required part relationships and XML order. Close only the task-owned document before package-level edits, then reopen and render in Visio.

If reopening stops responding, distinguish successful creation/save/render from the pending reopen stage. Inspect the live process and any visible dialog. A still-running call is not a failed file; a valid ZIP alone is not a successful Visio reopen. When even a minimal saved document stalls, do not repeatedly regenerate the same drawing or launch more instances. Keep the saved artifacts and stage log, resolve the interactive application state, then retry only the missing check. Do not change licensing, security or global Office settings as a drawing workaround. The documented [OpenEx flags](https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex) can select read-only/hidden/no-workspace behavior; they do not guarantee that an application call cannot block.

In observed host tests, creating Visio and performing save-close-reopen from the same Windows PowerShell 5.1 client could stall with both `Documents.Open` and `Documents.OpenEx`, including when a guard document remained open. The internal cause has not been established; do not attribute the problem only to OpenEx or treat changing its flags as a proven fix.

The verified strategy was to let one Windows PowerShell 5.1 process create only the Visio application and a uniquely identified task-owned blank guard, then exit. A separate drawing client reattached with `GetActiveObject`, and completed saving, closing, reopening and PNG/EMF export. The helper automates that startup sequence through [start_visio_host.ps1](../scripts/start_visio_host.ps1) when no running instance is available. It waits at most 20 seconds for the startup worker to exit, records the worker PID, Visio PID/HWND and guard identity, and verifies the returned HWND and guard ID/title before using the new instance. A startup timeout leaves evidence for inspection; it does not kill or restart the worker or Visio. The 20-second limit covers the startup-worker wait, not every later COM call. Verify the receipt and saved outputs for each run rather than assuming this observed strategy proves success in every host state.

When Visio is already running, the helper reuses that instance without closing its existing documents. It keeps its verified startup guard, or creates a task-owned guard when needed before closing the new drawing, and closes only task-owned documents during cleanup. For its own newly constructed macro-free VSDX it currently uses `Documents.Open`. This is not permission to open arbitrary macro-bearing input with weaker controls. Keep existing user files protected and do not alter global security settings to obtain a preview.

With both targets selected, a reported target failure is recorded independently and the other target is attempted. Inspect each target's receipt entry, preserve successful artifacts, and recover only failed or unattempted targets into a fresh directory. A COM call that has not returned is still pending; inspect its stage and live process before launching a recovery attempt.

## Save, exports and verification

Save the `.vsdx`, reopen the saved version and export the actual page as a preview. Inspect native shapes/text, groups, geometry, Connects and media/OLE entries. On a temporary copy move a representative node and verify attached lines follow; inspect the complete tiny motifs and joins in the render.

EMF/SVG/PDF/PNG are optional exports of the verified page. For an editable PPTX request, construct native PowerPoint objects from the shared spec. Do not insert the whole Visio render into PowerPoint and label it converted.

Only manage task-owned documents. Reuse an existing Visio session cautiously, never close unrelated files, and never globally terminate VISIO to recover a drawing task. A timeout needs state inspection, not an assumption that work failed or stopped.
