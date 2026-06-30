# Label Redesign - Notes

A polished, clinic-grade redesign of the DK-1202 medication label. Drop-in
replacement for the existing print workflow - nothing about parsing, selection,
logging, or the Brother print path changed.

## What changed in code

Two routines were rewritten and four small helpers added in `MedParser.bas`:

- `BuildLabelPreviewLayout` - new grid, typography, section rules, and logo.
- `UpdateLabelPreviewFromSelection` - writes the new cell map + adaptive sizing.
- `SetMiniValue` - renders a small field label + larger bold value in one cell
  (used for the EXP / LOT footer).
- `MedFontSize` / `NameFontSize` / `SigFontSize` - step the font size down for
  long medication names, patient names, and directions so they stay readable
  and on the label (Excel cannot "shrink to fit" inside merged cells, so this is
  done in code).

`PrintLabel` and `PrintCheckedLabels` print areas were updated `A1:H10 -> A1:H15`
to match the taller label.

## Layout (printed area = A1:H15, landscape)

```
SATURDAY CLINIC FOR THE UNINSURED                 [ SC emblem, top-right ]
1121 E. North Ave, Milwaukee WI  .  (414) 588-2865
------------------------------------------------------------ (header rule)
PATIENT                                           DOB   03/14/1985
Doe, Jonathan                                     Rx    06/29/2026

Metronidazole 500 mg            <- HERO line, largest, bold
Tablet  .  Qty 56

DIRECTIONS
Take 1 tablet by mouth twice daily with food.
------------------------------------------------------------ (footer rule)
EXP  05/2027                      LOT  ABC1234
```

Cell map (printed rows 1-15; rows 16-18 hold off-label helper text that does not
print):

| Field | Cell(s) |
|---|---|
| Clinic identity | A2:H2 |
| Contact line | A3:H3 |
| "PATIENT" mini-label | A5:E5 |
| DOB | F5:H5 (right) |
| Patient name | A6:E6 |
| Rx date | F6:H6 (right) |
| Medication + strength (hero) | A8:H8 |
| Dosage form + qty | A9:H9 |
| "DIRECTIONS" mini-label | A11:H11 |
| SIG / directions | A12:H12 |
| EXP | A14:D14 |
| LOT | E14:H14 (right) |

## Design choices

- **Hierarchy by size/weight, not color.** Thermal printing is monochrome, so the
  design carries hierarchy through type size, weight, and small uppercase
  labels - never color. The medication name is the clear focal point; patient
  name and directions are next; clinic identity and DOB/Rx are supporting.
- **Clean structure, no heavy box.** Two thin hairline rules (under the header,
  above the EXP/LOT footer) organize the label without a "spreadsheet" look.
- **Emphasized footer.** EXP and LOT sit in their own footer with a rule above
  them and bold values, so they are always easy to find for safe dispensing.
- **Generous whitespace and consistent left margin** for a calm, professional feel.

## Typography / font fallback

- Font is **Arial** throughout. Helvetica is not reliably installed on Windows
  Excel; if you set "Helvetica" it silently substitutes anyway. Arial is the
  closest reliable Helvetica-equivalent, is on every Windows machine, and prints
  crisply on the Brother. (Segoe UI is a fine alternative if you ever want a more
  modern feel - change the `.Name = "Arial"` lines in `FmtLbl` / `SetMiniValue`.)
- Excel cannot apply per-cell letter-spacing, so the on-screen mockup's subtle
  tracking on the clinic name is approximated by uppercasing only. The printed
  result is still clean.

## Branding / logo

- The SC emblem was extracted from the full clinic logo (just the circular mark,
  no wordmark) and converted to a crisp **all-black** PNG for clean thermal
  output: `scu_emblem.png` (kept next to the workbook).
- It is placed small in the **top-right** of the header and is inserted with
  `SaveWithDocument = msoTrue`, so once it loads it is **embedded in the workbook**
  and travels with the file even if the PNG is later moved or OneDrive-dehydrated.
- The insert is wrapped in error handling: a missing or unreadable logo can never
  break setup or printing - the label simply prints without it.

## Print + Excel compatibility (unchanged, offline)

- Target: **Brother QL-1100C**, **DK-1202 62 x 100 mm**, **landscape**. Confirmed.
- **No `.PaperSize`** in VBA - the Brother driver controls the media size.
- Excel controls only print area, orientation, margins (0.04"), and `Zoom = 100`
  (no fit-to-page, so nothing is squished).
- Printed content height = **170 pt = 2.361"**, which exactly matches the DK-1202
  printable height (62 mm minus the 0.08" margins). Width is unchanged from the
  previously-working layout, so horizontal fit is preserved.
- 100% local / offline. No fonts, images, or services are fetched from the network.

## Edge-case handling

- **Long medication names** step down 14 -> 12.5 -> 11 -> 10 pt and can wrap.
- **Long patient names** step down 13 -> 11 -> 9.5 pt and can wrap.
- **Long directions** step down 10.5 -> 9.5 -> 8.5 pt; the SIG block is 36 pt tall
  (~3 lines) so multi-line instructions (insulin, tapers) wrap cleanly.
- **Short/simple labels** keep the EXP/LOT footer pinned to the same position, so
  the critical fields are always in the same place regardless of content length.

## Remaining limitations / tradeoffs

- Exact on-paper kerning/spacing will differ slightly from the on-screen mockup
  (Excel text metrics differ from a browser), but the structure and hierarchy match.
- Extremely long medication names plus very long directions on the same label is
  the only stress point; both shrink, and the directions should be kept concise.
- The one thing not yet verified is a **physical test print** - geometry is exact
  on paper, but confirm one real label once the printer is available.
