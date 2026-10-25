# Editable equations and MathType

Read when reference content contains formulas or the user asks for MathType. Track each formula's source text, intended mathematical meaning, target object ID and actual equation technology. Superscripts in a text box, native Office mathematics (OMML), a vector picture and MathType OLE are distinct choices.

## Choose truthfully

When the user explicitly requests MathType, build and verify real MathType objects. Otherwise preserve an existing accepted formula type or choose a suitable editable representation; simple variable labels can remain formatted text. Do not convert every label into OLE just because the figure contains one equation.

Detect MathType and the destination application's OLE capabilities. Installed MathType in Word alone is not proof that a PowerPoint or Visio insertion path works. If a required MathType route is unavailable, continue the rest of the figure, retain formula source text and clearly report the formula requirement as incomplete. Do not silently replace it with an EMF.

## Native Visio formulas

Use this route for simple linear formulas when native formatted text is acceptable. Fractions, matrices, nested limits or an explicit MathType request need the corresponding equation route. Native character formatting is editable text, not MathType or OMML.

In the MLP/KAN reconstruction, separate text boxes for equation fragments rendered acceptably in PowerPoint but wrapped indices and shifted operator baselines in Visio. Enlarging `TxtWidth` and anchoring the text at the left edge reduced wrapping without fixing the shared baseline. Replacing each complete formula with one native Visio text shape, then formatting its character ranges, solved both symptoms in the saved and reopened output. The internal cause of the earlier fragmentation is unproven; do not blame a particular space character, font or COM operation without isolating it.

For this class of formula:

1. Assign the complete, final Unicode string to `Shape.Text`. Retain that string and an explicit list of symbol/index ranges. Set the base font, size, style and normal character position before applying local formatting.
2. Set the intended text block width/height, margins, alignment and text pin consistently. Typical formulas use `TxtWidth=Width`, `TxtHeight=Height`, `TxtPinX=Width*0.5`, `TxtPinY=Height*0.5`, and matching local pins. Size the shape for a complete line; split a long formula at meaningful operators only when the design needs multiple lines. An oversized text block can expand export bounds and is not a baseline fix.
3. Use `Characters.Begin` and `End` for each range, with a zero-based start and exclusive end. Include the entire index, such as `L−1`, rather than lowering only its final digit. Compute ranges from the final assigned string; recompute after normalization or text edits.
4. Apply real character-position formatting for superscripts/subscripts. Avoid manually offset text boxes or Unicode subscript substitutions as the default for a complete formula. Keep upright and italic symbols intentional; in the example the `Φ` characters were upright while the surrounding mathematical variables remained italic.
5. Resolve fonts in the destination document. If Latin/Greek symbols use an unintended Asian fallback, inspect both `Char.Font` and `Char.AsianFont`; the verified example explicitly set both to its selected Times New Roman font. Do not impose a Latin font on actual Chinese text.

The installed Visio PIA verified these constants; prefer its named enums when adapting to another installation:

| Purpose | PIA enum | Verified value |
| --- | --- | --- |
| Character style cell | `VisCellIndices.visCharacterStyle` | 2 |
| Character position cell | `VisCellIndices.visCharacterPos` | 4 |
| Character size cell | `VisCellIndices.visCharacterSize` | 7 |
| Normal / superscript / subscript | `VisCellVals.visPosNormal` / `visPosSuper` / `visPosSub` | 0 / 1 / 2 |
| Character section | `VisSectionIndices.visSectionCharacter` | 3 |

`visCharacterPos` is **4**, not the character-size index. Setting subscript position in the tested host provided the native size/offset behavior; do not also shrink every index blindly. Superscript uses the same API with `visPosSuper`; check its actual appearance in the target figure rather than claiming the subscript example tested every position/style combination.

For Windows PowerShell's ambiguous indexed COM setters, a small task-local C# bridge can call the typed PIA directly. Load the PIA from the detected installation, not a hardcoded user path:

```csharp
// Compile with a reference to the installed Microsoft.Office.Interop.Visio PIA.
using Microsoft.Office.Interop.Visio;

public static class SftVisioCharacters {
    public static void SetPosition(IVShape shape, int start, int length, short position) {
        IVCharacters chars = shape.Characters;
        chars.Begin = start;
        chars.End = start + length;
        chars.set_CharProps((short)VisCellIndices.visCharacterPos, position);
    }

    // For a uniform character-formatting run after saving and reopening:
    public static short ReadPosition(IVShape shape, int start, int length) {
        IVCharacters chars = shape.Characters;
        chars.Begin = start;
        chars.End = start + length;
        short row = chars.get_CharPropsRow(0);
        if (row < 0) throw new System.InvalidOperationException("Mixed character run");
        return (short)shape.get_CellsSRC(
            (short)VisSectionIndices.visSectionCharacter, row,
            (short)VisCellIndices.visCharacterPos).ResultIU;
    }
}
```

`CharPropsRow` returns a **row index**, not the position/style value; its argument is a left/right bias, not a cell index. Read the appropriate character cell. Verify every intended uniform run; if a selection spans different runs, inspect its constituent runs instead of assuming the first row certifies the entire range. For PowerShell 5 task scripts containing literal Greek/minus/ellipsis characters, preserve an explicit compatible encoding (UTF-8 with BOM worked in this conversion), or construct the intended Unicode characters explicitly.

After saving and reopening, compare the full formula text and all intended sub/superscript ranges, then render the actual page. Check shared baselines, Greek letters, the minus sign, composition operators, ellipses, clipping and canvas proportions. An unexpected increase in preview height is a useful signal for overflow or wrapping, not a reason to crop the preview and hide the defect. A formula replacement changes object counts: update the expected inventory while preserving unrelated network objects and connections.

## Construction options

1. Reuse a user-supplied, verified MathType object when its formula is exactly the intended one. Copy its native object with its preview; never put a new visual preview on an unrelated formula payload.
2. Create an equation through the installed MathType editor/integration, then insert or paste it as an **embedded** OLE object into the target host. Check the actual paste/insertion type. A clipboard operation can select a picture format unless explicitly controlled.
3. When available, the local `word-mathtype` skill can create a small carrier Word document from declared TeX and verify its native payload. Read that skill only for this optional Word conversion subtask. It does not itself guarantee PowerPoint/Visio embedding. Transfer the verified object into the selected host and verify again there. Avoid converting an already-converted document through Toggle TeX a second time.

Do not manufacture Equation.DSMT4 metadata around a PNG or a dummy binary. Keep formula-source strings and the actual native payload associated by object ID; multiple formulas require distinct, correctly matched payloads.

## Verification in the saved target file

- Inspect native embedded OLE relationships and actual object type. `Equation.DSMT4` is a useful identifier, not sufficient proof by itself.
- Inspect the compound object for a valid Equation Native/MTEF stream using an available parser (e.g. olefile plus a verified MathType audit). Compare the intended symbols/content with the known source or independent valid conversion where possible.
- Check rendered preview, font sizing, baselines, sub/superscripts, vector accents, fractions and bounding-box clipping. Preserve mathematical meaning during size adjustment.
- Save, close and reopen the target file. On a copy, activate a representative equation in MathType and verify it is editable. For all objects check that embedded payloads persist and match their intended formulas.
- If activation cannot be verified, state the narrower verified status: embedded MathType payload verified, activation in target application not verified. Do not declare full editability proven.

The package audit supplied with Senior-FigTrans reports equation markers and payload relationships; it is not a full MTEF semantic parser and cannot certify MathType on its own.

The final `.pptx` / `.vsdx` must retain the declared equation representation. A publication EMF exported from the figure is a vector picture and need not itself contain editable MathType. Keep those claims separate.
