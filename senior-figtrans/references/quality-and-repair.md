# Geometry and visual repair

Use for reconstruction and bounded edits. The current user-approved source is authoritative; preserve unrelated accepted details.

## Before drawing

Separate the task's meaning from its decoration: a system mechanism, an algorithm flow and a quantitative result need different evidence. Reproduce a supplied design unless redesign is requested. Compactness comes from a shared grid, balanced occupied areas and deliberate routing lanes; do not add meaningless modules or remove necessary detail to make the diagram easier.

Use the accepted figure boundary as a common coordinate frame for the source and output. Record the intended placement width when the user supplies one. Otherwise use the deliverable's page or slide size and make that assumption visible in the review. Mark the regions that carry meaning and give their objects stable names. A region can be a panel, an inset, a group of connected modules or a route crossing panel boundaries. Include small motifs, miniature plots and the objects named by each callout.

For each region, record the visible evidence that distinguishes a faithful result. For text, capture exact wording, case, mathematical symbols and logical line breaks. For repeated components, capture count, order, spacing and alignment. For connectors, capture source and destination shape or member, port, direction, intended junctions and whether the relation must stay attached when a node moves. Record important curves and grid patterns at the level the input supports. Mark an unreadable character or ambiguous junction at its specific location and continue with the legible regions. A cropped or low-resolution source may support visual tracing of a pattern while leaving its underlying data unknown.

Build one representative repeated module at its final geometry, then check its label, ports and smallest internal details against the reference before reusing it. This catches systematic drift before it spreads across rows. Keep the source region and the target object names in the reconstruction spec or temporary review notes so a correction can reach the responsible objects directly.

## Connectors: diagnose before changing size

Measure at the intended output scale. A source stroke of 3 pt scaled to half width appears about 1.5 pt. The user's approximately 3 pt preference applies when requested to that figure; it is not a universal width for all diagrams.

- Use one continuous route for an elbow or polyline. When an API requires segments, share exact endpoint coordinates, apply an arrow only at the terminal segment, and check painted joins. Do not cover a tail with a white patch.
- Keep the stem and the connected route at the same width within one logical connection. Major flows, secondary flows, panel borders and miniature graph edges may have distinct weights.
- Set arrowhead length and width separately from stroke weight. A wide head can be ugly even when the stem is correct; shrinking every arrow can instead make the figure unreadable.
- Leave enough straight terminal run for the selected head. Move the bend, reserve a lane, or adjust the terminal length before making a head microscopic. Check that a short final segment does not fold its head into the corner.
- Test the painted bounds, not just the centerline. A connector nominally within the canvas can still be clipped by a group bound, shape mask or page edge.
- Use native connector attachment where movement should preserve relations. Explicitly distinguish fixed coordinate paths from attached relationships.
- Check the intended member and port at both ends. A route can touch the correct group border while landing on the wrong child or on a generic center port. Where branches share a stem, check the actual shared junction, direction and destination at each terminal; a visual crossing alone cannot establish a junction.
- Check edge-to-arrow joins, T junctions, self loops, coincident duplicate lines and unintentional overshoot. An intentional open circuit, switch gap or nonjoining crossing must remain open.

For a geometric check, use one documented tolerance appropriate to the physical size (often at most 0.5 pt for intended coincident endpoints), then inspect the rendered result. Tolerance is a diagnostic aid, not permission for visible defects.

## Symmetry and repeated modules

Make column centers, row baselines and equal margins explicit. Distinguish the bounds of the picture frame from the nontransparent visible content. Normalize apparent size without distorting aspect ratio. Align icon centers, label baselines and outer frames independently.

For four symmetric stages or a two-row model family, use shared dimensions and anchor arrays. Long labels need room; don't shrink only one label unnoticed. Check the bottom row as carefully as the first row. Make each repeated cell a reusable group only after its contents align.

## Detail and style

- Match the accepted palette and typography; a local arrow repair does not authorize recoloring the figure.
- Transparent automotive/building icons should share perspective, outline/fill style and visual weight. Inspect for residual rectangular backgrounds and unwanted internal backing shapes.
- A tree is complete only when every intended node and branch is present. Check tiny motifs at enlarged scale as well as print size.
- Every legend sample must encode an element that is actually used; every numbered callout must identify a visible area or relation. Do not add unsupported mechanisms just to justify a leftover legend.
- Font substitution can change line breaks and apparent alignment. Reflow within the intended box or adjust layout before indiscriminately shrinking fonts.
- Read the smallest labels from the export at the intended placement size. In the enlarged crop, check that `I`, `l` and `1`, Greek symbols, subscripts and superscripts still match the source. Trace full text through its rendered baseline and compare nearby borders, ports and arrowheads for enough clearance. An intact text object can still render clipped or sit on a connector.
- Review complete formulas in each host. If fragmented Visio text drifts off the common baseline, use [one native text object with range formatting](equations.md#native-visio-formulas) when that equation technology is appropriate. Do not conceal wrapped indices or operators by cropping the preview.
- Remove unwanted notes or bottom-row prose when requested; do not replace it with implementation details or credits.

## Quantitative panels

Data retain their values, units, case selection, method order and scale during visual repairs. Additional curves or uncertainty bands need actual supporting arrays and meaningful interpretation. Preserve neutral and adverse outcomes alongside improvements. More panels do not establish innovation; merge redundant views according to their scientific roles, not a fixed panel quota. A schematic curve without source data is an illustration, not a new measured result.

## Review and acceptance

Export after saving and reopening the deliverable in its target host. Tie the export to that native file and page. With both PPTX and VSDX requested, each file needs its own target-host check. Use three views of that fresh export for distinct questions.

| View | Evidence to inspect |
| --- | --- |
| Whole canvas | Figure boundary, aspect ratio, reading order, panel alignment, major routes and relationships between regions. Follow routes that cross region boundaries from start to finish. |
| Intended placement size | Smallest text, formula legibility, stroke and arrowhead balance, contrast and clearance between nearby elements. Scale the whole figure to the intended slide, page or insertion width. A zoomed application window cannot answer this check. |
| Paired detail | Enlarged crops of the same source and output regions, including their nearest meaningful anchors. Use the figure boundary to map coordinates, keep both crops at a comparable scale and retain a shared anchor when a layout change moves a region. Inspect dense text, tiny branches, repeated marks, connector ports, joins and formula baselines. |

Use the region inventory to compare the observed output with the accepted source. For repeated motifs, compare one exemplar closely and scan the full set for count, missing members, drift and outliers. Include each distinct topology and any clipped or crowded instance in paired detail. For a bounded repair, focus paired checks on the edited regions, their neighbors and the ends of affected routes. Retain the annotated crop or coordinates long enough to verify repairs. A preview that cannot be tied to the reopened deliverable and an unavailable target host both leave target rendering unverified.

Run the native structure audit for each deliverable using `python scripts/audit_editability.py FILE --require-native --output audit.json`. Inspect the relevant native object names, text, group members, chart or formula objects and expected connector endpoints. `STRUCTURE_PASS_REVIEW_REQUIRED` confirms the automated structural gate only. Confirm representative editability on a temporary copy. Edit text in a dense region, move a node whose connections must follow, and where applicable move a content group, ungroup a repeated module or activate an equation object. Check actual endpoint identities, intended routes and changed coordinates, then save and reopen the test copy. Select the exercises from what the user expects to edit.

Use `region / object / source or expected behavior / observed defect / fix / evidence to recheck` for short repair notes. Fix the smallest set of responsible objects, then check neighboring content and the whole canvas again. Prioritize incorrect content and topology, then missing editability where required, then rendering, typography and aesthetic alignment. Batch related changes before re-exporting. If the same symptom returns, inspect the owning shape, group, font or backend before changing visual parameters again. Report source ambiguities or unavailable verification with the affected object and specific consequence.

If inserting the figure into a manuscript is requested, check the actual inserted page at its final placement size too. Recheck font rendering, clipping and stroke weight after insertion.

## Worked review for the iTransformer demo

The released reference and preview both contain four callout panels around the central `TrmBlock`. The overall composition is recognizable. The following region checks expose changes that the whole-canvas view and object counts can miss.

| Region | Source and checks that drive repairs |
| --- | --- |
| (a) Embedding | Match the three colored raw-series traces, their local turns and their one-to-one row alignment through the embedding block. The demo preview redraws several turns on the lower traces. Align the traces to the source crop near the axes before checking the pink arrows and three token rows at intended size. |
| (b) Multivariate Attention | Follow each colored variate into Query, Key and Value, then trace `MatMul` through `Scale` and the correlation map to the second `MatMul`. Check all three output terminals. The demo preview preserves the map's visual role with editable colored cells and changes its internal pattern. For a faithful visual trace, compare major color bands and their positions in paired crops. Exact numerical claims require the source values. Estimates based on a readable scale and cells should be identified as estimates. |
| (c) Feed-forward | Compare the faint Dense connection fans and their endpoints row by row. The demo preview makes the fans visually heavier. Preserve legible endpoints and line order when adjusting weight. Check the three activation symbols and the `Dense`, `Act & Drop`, `Dense` labels together at placement size. |
| (d) Temporal LayerNorm | Read the two statistics, hat, fraction line and denominator in one crop. Check arrow direction, grouping of the three variates and that no arrow or token obscures a formula symbol. |
| Central `TrmBlock` | Inspect both residual loops, their add junctions and the Q/K/V split. Test moving a block when these routes are expected to keep following it. The released native files attach `main_emb_to_attention`, and store `main_residual1`, `main_residual2`, `main_q_branch` and `main_v_branch` without native endpoint records. Those routes need a movement check and repair if the requested edit makes them dynamic. |
