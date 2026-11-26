# Native PowerPoint route

Use for PPTX or both. Use an existing connected PowerPoint toolset when available; discover tools by their actual names and inspect current capabilities. A typical Scientific Illustrator installation exposes `powerpoint_status`, `powerpoint_get_capabilities`, `powerpoint_inspect`, shape/text/connector APIs, exact alignment, grouping, export and save. Do not assume a tool supports an object because PowerPoint itself does.

## Choose the implementation

- Windows PowerPoint: native COM or installed PowerPoint MCP tools. Preserve foreground focus by default; use a new task-owned presentation, or a copy of the user's specified deck.
- macOS PowerPoint: a connected Office.js task pane can edit live. Inspect supported APIs; some chart and connector implementations are editable composites rather than native chart/attachment objects. Report that distinction.
- File-backed OOXML generation: allowed when native live automation is unavailable. Use real DrawingML shapes, text, connectors, charts and tables. State the renderer used; a generated file or a successful open dispatch does not prove the target app opened it correctly.
- WPS: only when requested or deliberately selected as an available PPTX host. Verify actual open/save state; do not describe file refresh as a live PowerPoint COM connection.

The bundled `render_scene.ps1` builds basic native Office COM objects without a plugin. Use the [PPTX refinement helper](scene-helper.md#refine-an-existing-native-pptx) on an existing deck for the custom head and formula placement operations described below. Advanced charts, formula OLE, imported icons and other complex paths use targeted native API operations in a task script.

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

Set line weight independently of arrowhead width and length. Inspect native straight or elbow presets when they meet the reference. After attaching or moving a node, recheck routed bends. For a deliberately fixed multi-bend path, use a freeform and record its coordinate-based routing in the specification.

For native frame/curve groups, attach connectors to their intended member before grouping, then inspect saved relationships recursively through `GroupItems` (`Shape.Type=6`). PowerPoint's `BeginConnect` and `EndConnect` use shape-specific site numbers. Resolve the actual site for the shape and verify its identity again after grouping. Compare actual exports for the appearance of arrowheads and dashes across hosts.

## Custom heads and semantic receiver groups

Use this approach when a known arrow silhouette, tip position or shaft-to-head join needs more control than PowerPoint's presets provide. The PID example uses native connector shafts and filled PowerPoint freeforms with a curved rear edge. For ordinary arrows use the native preset that fits. For custom heads, group by the receiving node so the incoming arrowhead and its attachment point follow the node during edits.

1. Measure the source at the target slide scale. Record the tip, incoming direction, head length and width, rear outline and shaft weight separately. A filled vector path can inherit a stroke-width property even when its stroke is disabled. Read the shaft weight from the shaft itself. Keep the head profile relative to its tip and map it into slide coordinates using the direction of the incoming stem.
2. Build a closed editable head with the matching native straight and cubic Bézier segments (`BuildFreeform`, `AddNodes`, `ConvertToShape`). Use the head fill and keep the outline off if that matches the source. Turn off the connector's native end arrow for this link. Keep enough terminal run for the head and verify the painted tip reaches its insertion point.
3. Give the connector an exact landing point in the rear of the head. Create a small native port shape with invisible fill and stroke, then align the selected site to the measured shaft entry. Where the outline curves, a small tested overlap into the fill keeps the join closed at export. Attach the connector end to that site. Its begin remains attached to the intended source member and site. Verify the shaft follows the intended incoming direction after attachment.
4. Group the receiver's visible body, associated labels and markings, and each incoming head with its port. Use semantic names for these objects and give each incoming link its own head and port. Keep connector shafts outside the group so other nodes remain independently movable. Attach before grouping and recheck the saved native relationship to the inner port. Outgoing connectors still use the relevant source member of the receiver. Inspect the group's layering after grouping, including whether each head covers its shaft end without obscuring adjacent labels.
5. For a text-only edit, select the existing text-bearing child within the group, change its native text and preserve its identity. Edit the intended range in a styled label so adjacent runs keep their formatting. Reassess any line breaks, auto-fit or bounds that the edit affects. Replacing the child or regrouping requires a fresh attachment check. Test custom heads with representative moves along the allowed routing lanes. If a receiver move causes the connector to approach from a different direction, reorient or reroute the head and recheck the join. Group rotation or arbitrary routing changes require their own verification.

This construction is verified in desktop PowerPoint. Visio uses its own ShapeSheet geometry and connection points. For two requested formats, implement and check the Visio path independently when custom head geometry is required. See [Visio groups and attachment](visio.md#groups-and-attachment).

## Save and inspect

Save to the requested `.pptx`, reopen that file, and export the slide with PowerPoint when available. Run the package audit and compare the actual render with the source. Report the renderer used if PowerPoint rendering is unavailable.

For movable relationships, edit an existing child label or move a representative node on a disposable copy, as applicable. Move a complete receiver group when present, then save and reopen the copy. Compare connector and endpoint shape IDs, names, connection-site numbers and global endpoint coordinates in the two saved PPTX files. Expect the endpoints on the moved group to follow its translation and each stationary endpoint to stay at its prior position. Check that the head and port move with their receiver, that the head keeps its shape, and that edited text persists on the same child. Nested group coordinates may need transformation into slide space. A connector's overall bounding-box change cannot show which endpoint moved. Inspect render crops of the moved junctions to check approach direction, overlap and gaps.

In the PID edit check, desktop PowerPoint moved the plant group 12 pt horizontally. Reopening preserved the edited plant label on the same child, the moved incoming actuation end and the moved outgoing feedback start. All other endpoints stayed in place. These tests illustrate which geometry to measure for a node that both receives and sends flows. Base object counts, expected identities and displacement on each figure's own specification.

Do not overwrite unrelated decks, close user presentations or quit an application containing user documents. Only close task-owned saved working documents when finished.
