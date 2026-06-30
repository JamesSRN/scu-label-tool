# Label redesign (2026-06-30)

Reference for the current printed label layout in `MedParser.bas` (`BuildLabelPreviewLayout`, `PrintLabel`, `PrintAllLabels`).

## Physical target

- **Printer:** Brother QL-1100c
- **Media:** DK-1202 continuous die-cut labels, **62 x 100 mm**, **landscape**
- **Print area:** `A1:H15` on the `Label Preview` sheet (hidden from normal view)

## Layout (rows 1–15)

| Rows | Content |
|------|---------|
| 2–3 | Clinic header — Helvetica (Arial fallback), shrink-to-fit |
| 5–6 | Patient name (A:E), Rx / DOB (F) |
| 7 | Medication name |
| 8 | Form / quantity |
| 10–12 | SIG (directions) |
| 13 | EXP / LOT |
| Top-right | SC emblem (`InsertLabelLogo`) |

## Typography

| Element | Font | Notes |
|---------|------|--------|
| Header (clinic name, address) | Helvetica → Arial | `LabelHeaderFont()` |
| Patient name | Helvetica → Arial | `PatientNameFontSize()` — at least med line + 1 pt, steps 17→8 |
| Medication, SIG, footer | Helvetica → Arial | `NameFontSize()` — steps 14→8 |
| Rx / DOB / EXP / LOT | Helvetica → Arial | Fixed 8 pt |

All body/header text uses shrink-to-fit where applicable.

## Print width

- **`LABEL_WIDTH_PT = 228`** — scales columns A:H so content fits **one** label (not two).
- **`ApplyLabelContentWidth`** runs from layout and print paths.
- Tune this constant if Brother output is still too wide or narrow on physical media.

## Logo / emblem

| Item | Value |
|------|--------|
| Source crop (manual) | `cropped_Black SCU Logo + Transparent Background - Copy.png` |
| Workbook file | `scu_emblem.png` (copy via `tools/Build-ScuEmblem.ps1`) |
| Aspect ratio | `LOGO_ASPECT = 1.488` (6150×4133 px) |
| On-label height | **28 pt** in `InsertLabelLogo` |
| Insertion | `AddPicture2` with compress=`0`; fallback `AddPicture` |
| Resolution | `LogoFilePath()` prefers local `scu_emblem.png`; else embedded `LogoB64()` temp file |

### Logo workflow

1. Edit the manual crop PNG if the emblem artwork changes (do **not** re-enable auto-crop without approval).
2. Run `tools/Build-ScuEmblem.ps1` → refreshes `scu_emblem.png`.
3. Run `Build-Release.vbs` → `SetupWorkbook` re-imports logo on `Label Preview`.

### Known issue — logo sizing

Build compiles and prints, but the emblem may still appear **too large or small** on the physical label. Next tuning targets:

- `InsertLabelLogo` height (currently 28 pt)
- Anchor cell / top-right placement in `BuildLabelPreviewLayout`
- `LOGO_ASPECT` if the crop dimensions change
- Brother driver scale / “fit to page” settings (should match 100 mm label width)

## Build / release

See `tools/BUILD_RELEASE_NOTES.md` and `HANDOFF.md` section 4.

- **`Build-Release.vbs`** — import `MedParser.bas`, run `'MedicationDispensing.xlsm'!SetupWorkbook`, save.
- **`tools/Run-BuildRelease.ps1`** — optional PowerShell wrapper with MsgBox auto-dismiss.
- Compile fixes applied: UTF-8 BOM stripped from `MedParser.bas`, `FONT_LABEL_HDR_FB` restored, optional default literals, `msoPictureCompressNone` → `0`, VBScript untyped `Dim`.

## Related constants (MedParser.bas)

```
LABEL_WIDTH_PT = 228
LOGO_ASPECT = 1.488
LOGO_EMBLEM_FILE = "scu_emblem.png"
LOGO_EMBLEM_SOURCE = "cropped_Black SCU Logo + Transparent Background - Copy.png"
FONT_LABEL_HDR = "Helvetica"
FONT_LABEL_HDR_FB = "Arial"
FONT_LABEL_BODY = "Helvetica"
```
