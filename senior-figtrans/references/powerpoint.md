# Native PowerPoint route

Use for PPTX or both. Use an existing connected PowerPoint toolset when available; discover tools by their actual names and inspect current capabilities. A typical Scientific Illustrator installation exposes `powerpoint_status`, `powerpoint_get_capabilities`, `powerpoint_inspect`, shape/text/connector APIs, exact alignment, grouping, export and save. Do not assume a tool supports an object because PowerPoint itself does.

## Choose the implementation

- Windows PowerPoint: native COM or installed PowerPoint MCP tools. Preserve foreground focus by default; use a new task-owned presentation, or a copy of the user's specified deck.
- macOS PowerPoint: a connected Office.js task pane can edit live. Inspect supported APIs; some chart and connector implementations are editable composites rather than native chart/attachment objects. Report that distinction.
- File-backed OOXML generation: allowed when native live automation is unavailable. Use real DrawingML shapes, text, connectors, charts and tables. State the renderer used; a generated file or a successful open dispatch does not prove the target app opened it correctly.
- WPS: only when requested or deliberately selected as an available PPTX host. Verify actual open/save state; do not describe file refresh as a live PowerPoint COM connection.

The bundled `render_scene.ps1` supports basic Office COM reconstruction without the plugin. Advanced charts, formula OLE, imported icons and complex paths need targeted native API operations in a task script; the helper deliberately rejects unknown object types.

## Mapping

| Reference content | Desired PPTX object |
| --- | --- |
| Labels / captions | TextFrame text, with explicit font and margins |
| Panels / blocks / symbols | AutoShapes, freeform paths, or editable member groups |
| Relationships | Native connector with explicit begin/end site attachment where supported |
| Deliberate fixed route | Native line/freeform polyline, one semantic ID; re-route after moving nodes |
| Regular table | Native table or editable grid/text composite |
| Data chart | Native chart with embedded workbook when suitable; otherwise editable primitives with retained data/spec |
| MathType formula | Verified MathType OLE, not a picture of a formula |
| Photo / approved illustration icon | Tightly bounded independent picture, with editable overlays separate |

One full-slide SVG is still one picture. Ungrouping an imported EMF can outline text, lose attachment and fragment strokes; inspect and repair those objects before making an editability claim. Prefer direct native construction.

## Coordinates and styles

PowerPoint uses points and a top-left origin. Preserve the reference aspect ratio; do not default a tall paper figure to 16:9. Use panel-local coordinates transformed once into page coordinates. Give semantic names to objects and group members. Apply consistent text margins and control auto-fit; verify long labels with the actual renderer.

Set line weight independently of arrowhead width/length. When using a native elbow connector, inspect its automatic route after attachment; after moving a node, re-evaluate its route. For a deliberately fixed multi-bend path use a freeform and clearly identify that it is not automatically attached. Do not superimpose arrow shapes on a separate stem to hide a bad join.

For native frame/curve groups, attach connectors to their intended member before grouping, then inspect saved relationships recursively through `GroupItems` (`Shape.Type=6`). Test movement of both a node and a complete group on a copy. PowerPoint's `BeginConnect` / `EndConnect` use shape-specific site numbers, not Visio's normalized coordinates. Preserve text/group members and endpoint identities after any regrouping. Native arrowhead and dash presets have different painted sizes across hosts; compare actual exports rather than sharing enum numbers or assuming equal preset names yield equal geometry.

## Save and inspect

Save to the requested `.pptx`, reopen that file, and export the slide with PowerPoint when available. Run the package audit and review the render against the source. On a disposable copy check editable labels, group membership and connector attachment. If a fallback renderer was used, label it; layout equivalence in PowerPoint remains unverified until checked there.

Do not overwrite unrelated decks, close user presentations or quit an application containing user documents. Only close task-owned saved working documents when finished.
