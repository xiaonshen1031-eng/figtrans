# Geometry and visual repair

Use for reconstruction and bounded edits. The current user-approved source is authoritative; preserve unrelated accepted details.

## Before drawing

Separate the task's meaning from its decoration: a system mechanism, an algorithm flow and a quantitative result need different evidence. Reproduce a supplied design unless redesign is requested. Compactness comes from a shared grid, balanced occupied areas and deliberate routing lanes; do not add meaningless modules or remove necessary detail to make the diagram easier.

Record a short object inventory with stable names. Cover tiny search trees, distribution curves, mini grids and branching motifs, not just their enclosing panels. If source text is unreadable, isolate that ambiguity while continuing legible areas. Do not invent model names or equations.

## Connectors: diagnose before changing size

Measure at the intended output scale. A source stroke of 3 pt scaled to half width appears about 1.5 pt. The user's approximately 3 pt preference applies when requested to that figure; it is not a universal width for all diagrams.

- Use one continuous route for an elbow or polyline. When an API requires segments, share exact endpoint coordinates, apply an arrow only at the terminal segment, and check painted joins. Do not cover a tail with a white patch.
- Keep the stem and the connected route at the same width within one logical connection. Major flows, secondary flows, panel borders and miniature graph edges may have distinct weights.
- Set arrowhead length and width separately from stroke weight. A wide head can be ugly even when the stem is correct; shrinking every arrow can instead make the figure unreadable.
- Leave enough straight terminal run for the selected head. Move the bend, reserve a lane, or adjust the terminal length before making a head microscopic. Check that a short final segment does not fold its head into the corner.
- Test the painted bounds, not just the centerline. A connector nominally within the canvas can still be clipped by a group bound, shape mask or page edge.
- Use native connector attachment where movement should preserve relations. Explicitly distinguish fixed coordinate paths from attached relationships.
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
- Review complete formulas in each host. If fragmented Visio text drifts off the common baseline, use [one native text object with range formatting](equations.md#native-visio-formulas) when that equation technology is appropriate. Do not conceal wrapped indices or operators by cropping the preview.
- Remove unwanted notes or bottom-row prose when requested; do not replace it with implementation details or credits.

## Quantitative panels

Data retain their values, units, case selection, method order and scale during visual repairs. Additional curves or uncertainty bands need actual supporting arrays and meaningful interpretation. Preserve neutral and adverse outcomes alongside improvements. More panels do not establish innovation; merge redundant views according to their scientific roles, not a fixed panel quota. A schematic curve without source data is an illustration, not a new measured result.

## Review and acceptance

Use these evidence channels together:

1. **Content correspondence:** compare the reference inventory to actual object names, text and rendered motifs. Include tiny branches and all legend/callout targets.
2. **Structure:** inspect native text/shapes, group members, connection endpoints, data-backed charts and formula type. A ZIP with 100 shapes is not evidence that all labels remain editable.
3. **Renderer:** review a whole-page export at intended reading size plus crops of dense details, formula baselines, bends and junctions. Check actual target software after saving and reopening.
4. **Interaction:** on a temporary copy, edit representative text, move a connected node, move a complete frame/curve group, ungroup a representative module and, when requested, activate a formula object. Confirm intended endpoint identities and coordinate changes, then save/reopen the copy. Do not disturb the final layout during this check.

Record concrete failures as `object / defect / correction / recheck`. Fix the smallest responsible set, then inspect affected neighbors and the whole figure. Batch related corrections; do not force a separate render per primitive. After a repeated failure with unchanged symptoms, inspect the cause or backend rather than retry the identical command. Stop at verified completion; report a remaining source ambiguity or unavailable capability precisely, without inventing a confidence score.

If inserting into a manuscript is requested, preview the actual inserted page too. Native source correctness does not prove Word's EMF rendering, fonts or clipping are correct.
