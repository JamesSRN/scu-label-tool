# Label redesign (2026-06-30)

Reference for the current printed label layout in `MedParser.bas` (`BuildLabelPreviewLayout`, `PrintLabel`, `PrintAllLabels`).

## Physical target

- **Printer:** Brother QL-1100c
- **Media:** DK-1202 continuous die-cut labels, **62 x 100 mm**, **landscape**
- **Print area:** `A1:H15` on the `Label Preview` sheet (hidden from normal view)

## Layout (rows 1–15)

| Rows | Content |
|------|---------|
| 2 | Clinic name — **A2:F2**, bold, bottom-aligned, 14 pt |
| 3 | Address / phone — **A3:F3**, top-aligned, 7 pt |
| 2–3 (right) | SC emblem — **G2:H3** slot, 30 pt height |
| 5–6 | Patient name (A:E), Rx / DOB (F) |
| 7 | Medication name (hero) |
| 8 | Form / quantity |
| 10–12 | SIG (directions) |
| 13 | EXP / LOT |

## Header cell map

| Field | Cells |
|-------|-------|
| Clinic name | **A2:F2** (merged) |
| Contact line | **A3:F3** (merged) |
| Logo slot (no fill) | **G2:H3** (merged) |

**Fixed gotcha:** do not merge A2:H2 — cell fill hides the emblem (only a sliver showed). Text and logo must stay in separate merge regions.

## Typography

| Element | Font | Notes |
|---------|------|--------|
| Clinic name | Helvetica → Arial | `FmtLblClinicName` — **14 pt print / 12 pt gallery**, always **bold**, bottom-aligned |
| Address | Helvetica → Arial | `FmtLblHeader` — **7 pt**, top-aligned, `ShrinkToFit`, no wrap |
| Patient name | Helvetica → Arial | `LabelNameFontSize()` — steps **12→8 pt** by length so long names fit the half-width name cell (both preview + gallery). Qty line reads **form · qty · source · refills**. |
| Medication, SIG, footer | Helvetica → Arial | `NameFontSize()` — steps 14→8 |
| Rx / DOB / EXP / LOT | Helvetica → Arial | Fixed 8 pt |

Header row heights: **20 pt** (row 2) + **10 pt** (row 3).

**Compile note:** use `Font.Bold = True` only — do not use `Font.Weight` or `xlBold` (breaks compile in this project).

## Print width

- **`LABEL_WIDTH_PT = 242`** — scales columns A:H so content uses more of the 100 mm die-cut.
- **228 pt** was the prior no-bleed sweet spot; **218 pt** was too narrow.
- **`ApplyLabelContentWidth`** runs from layout and print paths.
- Re-test on the clinic Brother after any width change.

## Logo / emblem

| Item | Value |
|------|--------|
| Source crop (manual) | `cropped_Black SCU Logo + Transparent Background - Copy.png` |
| Workbook file | `scu_emblem.png` (via `tools/Build-ScuEmblem.ps1` — pure black pixels) |
| Aspect ratio | `LOGO_ASPECT = 1.488` (6150×4133 px) |
| Print height | **`LOGO_HEIGHT_PRINT = 30`** |
| Gallery height | **`LOGO_HEIGHT_GALLERY = 28`** |
| Print slot inset | **`LOGO_SLOT_INSET_PT = 4`** |
| Gallery inset | **`LOGO_GALLERY_INSET_PT = 2`** |
| Right pad | **`LOGO_RIGHT_PAD_PT = 0`** |
| Insertion | Natural aspect (not `maxW × height`); `LockAspectRatio`; `ZOrder msoBringToFront`; vertical center in band |
| Print refresh | `RefreshPrintLabelLogo` on every preview update and before print |
| Resolution | **`LogoFilePath()` — local `scu_emblem.png` only** (embedded fallback removed) |

### Logo workflow

1. Edit the manual crop PNG if the emblem artwork changes (do **not** re-enable auto-crop without approval).
2. Run `tools/Build-ScuEmblem.ps1` → refreshes `scu_emblem.png` (forces solid black for thermal).
3. Run `Build-Release.vbs` → `SetupWorkbook` rebuilds layout and logos.

### Problems solved (logo/header)

| Issue | Fix |
|-------|-----|
| Emblem sliver in gallery/print | Separate merges: text A2:F2 / A3:F3, logo G2:H3; shape to front |
| Gallery OK, print wrong | `RefreshPrintLabelLogo` before print |
| Logo squished wide | Insert at natural aspect, not max-width box |
| Gray snake on thermal | `Build-ScuEmblem.ps1` pure black conversion |
| Clinic name clipped | Taller rows + bottom/top vertical align |
| Clinic name too small | Wider merge A2:F2; 14 pt font constant |
| Weird partial logo after rebuild | Removed `LogoB64()` fallback from `LogoFilePath` |

### Still verify

- **242 pt width** on physical Brother (one label per die-cut, no bleed).
- **Bold clinic title** on thermal — may look subtle; `ShrinkToFit` can cap apparent size.

## Print pipeline

- `ApplyLabelPageSetup` — `A1:H15`, **fit-to-one-page** (`.Zoom = False`, `.FitToPagesWide = 1`, `.FitToPagesTall = 1`), no `.PaperSize`, landscape. (V2: fit-to-page replaced `zoom 100 / no-fit` so the bottom EXP/LOT row can't clip on a printer with a shorter printable area.)
- `PrepareLabelSheetForPrint` — only `scuLogo` may print as a shape; preview buttons excluded.
- `PrintLabelSurfaceSafe` — **`PrintOut From:=1, To:=1`** (fixes blank 2nd label).

## Build / release

See `tools/BUILD_RELEASE_NOTES.md` and `HANDOFF.md` section 5.

- **`Build-Release.vbs`** — import `MedParser.bas`, run `'MedicationDispensing.xlsm'!SetupWorkbook`, save.
- Bootstrap from `Broken_PrettyPrint_…` **only if** `MedicationDispensing.xlsm` is missing (warning MsgBox).
- **`tools/Run-BuildRelease.ps1`** — optional PowerShell wrapper with MsgBox auto-dismiss.

## Related constants (MedParser.bas)

```
LABEL_WIDTH_PT = 242
LOGO_ASPECT = 1.488
LOGO_HEIGHT_PRINT = 30
LOGO_HEIGHT_GALLERY = 28
LOGO_HDR_ROW1_PT = 20
LOGO_HDR_ROW2_PT = 10
LOGO_SLOT_INSET_PT = 4
LOGO_GALLERY_INSET_PT = 2
LOGO_RIGHT_PAD_PT = 0
CLINIC_NAME_FONT_PRINT = 14
CLINIC_NAME_FONT_GALLERY = 12
CLINIC_ADDR_FONT_PRINT = 7
LOGO_EMBLEM_FILE = "scu_emblem.png"
LOGO_EMBLEM_SOURCE = "cropped_Black SCU Logo + Transparent Background - Copy.png"
FONT_LABEL_HDR = "Helvetica"
FONT_LABEL_HDR_FB = "Arial"
FONT_LABEL_BODY = "Arial"
```
