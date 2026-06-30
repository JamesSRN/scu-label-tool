# SCU Label Printing Tool - Handoff

Last updated: 2026-06-30 (end of label/logo/print tuning session)

An Excel + VBA tool for the **Saturday Clinic for the Uninsured (SCU)** free
pharmacy. It turns pasted prescription text into a validated medication table and
prints professional labels on a **Brother QL-1100C** (DK-1202, 62 x 100 mm,
landscape). It runs **100% offline** and is built to be robust for high-turnover
volunteers.

---

## 1. Where everything lives (IMPORTANT - read first)

Everything is in **one folder** (code, docs, workbook, backups, and Git):

```
C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing
```

This folder is **outside OneDrive** (Documents was moved off OneDrive sync in
June 2026). It contains the code, docs, the SC emblem, the workbook
(`MedicationDispensing.xlsm`), and `_backups\`. It is also the local Git repo
(remote: `JamesSRN/scu-label-tool`). The `.xlsm` and `_backups\` are git-ignored,
so patient data can never be committed.

**Two open setup items:**

1. **Trusted Location.** The old `Desktop\SCU Label Printing` folder (the prior
   Excel Trusted Location) was removed. The new folder is **not trusted yet**, so
   macros will be blocked until you add it:
   `Excel > File > Options > Trust Center > Trust Center Settings > Trusted
   Locations > Add new location > browse to the folder above > OK`.
2. **Easy access (optional).** Right-click the folder > Send to > Desktop
   (create shortcut). It can be renamed in place to drop the `GIT_VERSION_`
   prefix (then re-point GitHub Desktop with Locate).

Code/docs are version-controlled; the `.xlsm` (PHI) is local-only by design.

---

## 2. Workbook structure (5 visible tabs + 1 hidden)

- **Patient & Input** - patient name / DOB / Rx date + paste box. Buttons:
  `PARSE MEDICATIONS`, `Clear Paste Area`, `Reset Session`, `Start NEW Patient`.
- **Medications** - the parsed table. Columns A-N: `#`, Name, Strength, Dosage
  Form, SIG, Quantity, Refills, **Expiration**, **Lot**, Date of Rx, Confidence,
  Warnings, Raw Text, **Printed?**; then **O = "# of Prints"** (auto-increments,
  protected from manual edits) and **P = "Print?"** (double-click toggles a checkmark).
  Buttons (col Q): `+ Add Medication`, `- Remove Selected`, `Review & Validate`,
  `Preview ALL Labels`, `Print Checked Labels`.
- **Label Previews** - auto-generated gallery, one card per medication, each with
  Print / Edit / Remove. Rebuilds on tab activate.
- **Log** - running dispense log.
- **Setup & Help** - instructions.
- **Label Preview** (HIDDEN) - the internal print surface. Every label renders
  here and prints from here. Do NOT delete it.

Shortcuts: `Ctrl+Shift+P` parse, `Ctrl+Shift+R` reset, `Ctrl+Shift+L` refresh
Label Previews.

---

## 3. Feature set

### Parsing & validation
- Block splitter starts a new medication per drug-header line, even without blank
  lines between drugs.
- Per-drug **Confidence** (High / Medium / Low / Manual) from how cleanly it parsed.
- **Validation** flags missing strength/SIG/quantity/Exp/Lot, bad Exp format, and
  possible duplicates -> Warnings column.
- Add / Remove / Renumber medications manually. **Remove Selected removes every
  checked (Print?) row**, not just the clicked one.
- Exp/Lot forced to text.

### Selection, color, printing (checkbox-driven)
- **Double-click a Print? (col P) cell to toggle a checkmark.** A checkmark = "selected".
- Row background by state: **Selected = green** (matches the gallery tint),
  **Validated = blue**, **Non-validated = gray**. Confidence cell colored by its
  own triad. (No printed/lavender state - printing no longer recolors rows.)
- **Print Checked Labels** prints every checked med in sequence (skips any checked
  row missing Exp/Lot). Print confirmations list exactly what will print (single =
  full detail block; batch = numbered list flagging skips).
- **Brother auto-select** (`SelectBrotherPrinter`) finds the QL-1100C via WMI and
  confirms before printing.

### Logging & multi-patient
- **Every print is logged**: timestamp, patient, DOB, medication, strength, SIG,
  quantity, refills, Exp, Lot, Rx date, dosage form, print #, and volunteer
  initials. Single AND batch prints; initials asked once per batch.
- **# of Prints** auto-increments on each print and is protected from manual edits.
- **Start NEW Patient** clears patient + meds but keeps the Log (back-to-back
  patients). **Reset Session** is a full wipe (patient + meds + Log) with confirmation.

---

## 4. The redesigned label (current)

Polished, clinic-grade DK-1202 label. Printed area = **A1:H15**, landscape.

```
SATURDAY CLINIC FOR THE UNINSURED                 [ SC emblem, top-right ]
1121 E. North Ave, Milwaukee WI  .  (414) 588-2865
-----------------------------------------------------------
Doe, Jonathan                            Rx    06/29/2026
                                         DOB   03/14/1985

Metronidazole 500 mg          <- HERO line (largest, bold)
Tablet  .  Qty 56

DIRECTIONS
Take 1 tablet by mouth twice daily with food.
-----------------------------------------------------------
EXP  05/2027                   LOT  ABC1234
```

Cell map (rows 16–18 hold off-label helper text that does not print):

| Field | Cell(s) |
|---|---|
| Clinic name (bold, bottom-aligned) | **A2:F2** (merged) |
| Contact line (top-aligned) | **A3:F3** (merged) |
| Logo slot (no fill) | **G2:H3** (merged) |
| Patient name | A5:E6 (merged) |
| Rx date | F5:H5 (right) |
| DOB | F6:H6 (right) |
| Medication + strength (hero) | A7:H7 |
| Dosage form + qty | A8:H8 |
| "DIRECTIONS" mini-label | A9:H9 |
| SIG / directions | A10:H12 (white on black) |
| EXP | A13:D13 |
| LOT | E13:H13 (right) |

**Design / type:** Hierarchy is carried by size + weight + small uppercase labels
(never color - thermal is monochrome). Fonts use **Helvetica** when installed,
with **Arial** fallback (`LabelHeaderFont()` for header lines; body via `FmtLbl`).
Two thin hairline rules (under the header, above the footer); no heavy box. The
medication name is the focal point; EXP/LOT are pinned in a footer.

**Adaptive sizing:**

| Element | Logic |
|---|---|
| Patient name | `PatientNameFontSize()` — steps 17→8 pt, always **≥ med line + 1 pt** |
| Medication / SIG | `NameFontSize()` / `SigFontSize()` — steps down for long text |
| Clinic name | `FmtLblClinicName()` — **14 pt print / 12 pt gallery**, always **bold**, bottom-aligned in row 2 |
| Clinic address | `FmtLblHeader()` — **7 pt print**, top-aligned in row 3; one line, `ShrinkToFit`, no wrap |
| Header row heights | **20 pt** (row 2) + **10 pt** (row 3) — tight pairing with emblem band |

**Print width:** `LABEL_WIDTH_PT = 242` scales columns A:H via
`ApplyLabelContentWidth` so content uses more of the 100 mm die-cut. **228 pt** was
the prior sweet spot (no bleed); **218 pt** was too narrow. Re-test on the clinic
Brother after any width change.

**Logo / emblem:**

| Item | Detail |
|---|---|
| Manual source crop | `cropped_Black SCU Logo + Transparent Background - Copy.png` |
| Workbook file | `scu_emblem.png` (built by `tools/Build-ScuEmblem.ps1`) |
| Build script | Copies manual crop and forces every visible pixel to **solid black** (`#000000`) for crisp thermal output (gray snake anti-aliasing prints poorly) |
| Aspect ratio | `LOGO_ASPECT = 1.488` (6150×4133 px) |
| Print height | **`LOGO_HEIGHT_PRINT = 30 pt`**, vertically centered in rows 2–3 |
| Gallery height | **`LOGO_HEIGHT_GALLERY = 28 pt`**, column F slot |
| Slot inset | **`LOGO_SLOT_INSET_PT = 4`** (print), **`LOGO_GALLERY_INSET_PT = 2`** (gallery) |
| Right pad | **`LOGO_RIGHT_PAD_PT = 0`** — emblem flush to print-area right edge |
| Insertion | `InsertLabelLogo` — insert at natural aspect (never `maxW × height` box); `LockAspectRatio`; `ZOrder msoBringToFront`; `xlFreeFloating`; centered vertically in band after final size |
| Print refresh | `RefreshPrintLabelLogo` runs on every preview update and before each print (gallery rebuilds logos on tab activate; print sheet does not) |
| Resolution | **`LogoFilePath()` uses only local `scu_emblem.png` beside the workbook** — embedded `LogoB64()` fallback **removed** (caused partial/weird logos after rebuild) |

Do **not** re-enable automatic logo cropping without explicit approval — the
manual crop is the source of truth.

**Logo gotcha (fixed):** merging A2:H2 hid the emblem under cell fill (only a
sliver showed). Header merges now leave **G2:H3** open for the emblem; text uses
**A2:F2** / **A3:F3**; logo shape is brought to front.

**Print geometry:** `ApplyLabelPageSetup` sets `A1:H15`, `Zoom = 100`, no
fit-to-page, **no `.PaperSize`** (Brother driver controls media), 0.04 in margins,
landscape, no gridlines/headings.

**Blank label after each print (fixed):** Excel was sometimes sending **page 2**
(empty). `PrintLabelSurfaceSafe` now calls `PrepareLabelSheetForPrint` (only
`scuLogo` may print as a shape; preview buttons excluded; `ResetAllPageBreaks`)
and **`PrintOut From:=1, To:=1`**.

**SetupWorkbook fix:** header `PasteSpecial` runtime 1004 resolved via
`MatchHeaderFormat` (safe format copy for new Medications/Log columns).
`SetupWorkbook` also calls **`PreviewAllLabels`** after layout rebuild so gallery
and print sheet stay in sync.

See also `LABEL_REDESIGN.md` and `CHANGELOG.md` (Unreleased).

---

## 5. Applying code changes (release build)

`MedParser.bas` must stay **pure ASCII with Windows CRLF** line endings or Excel
rejects the import. **No UTF-8 BOM** on line 1. (`.gitattributes` enforces CRLF
for `.bas` / `.vbs`.)

### Quick path (recommended)

1. Close `MedicationDispensing.xlsm`.
2. Run `tools/Build-ScuEmblem.ps1` when the manual crop changes → rebuilds
   `scu_emblem.png` (pure black for thermal). Re-run after any crop edit.
3. Double-click **`Build-Release.vbs`** (needs **Trust access to the VBA project
   object model** + folder as **Trusted Location**).
4. Click OK on **Setup complete!** and **Release build complete**.

`Build-Release.vbs` opens `MedicationDispensing.xlsm` in place (bootstraps from
`Broken_PrettyPrint_MedicationDispensing.xlsm` only if missing), removes the old
`MedParser` module, imports `MedParser.bas`, runs
`'MedicationDispensing.xlsm'!SetupWorkbook`, saves.

**Bootstrap warning:** if `MedicationDispensing.xlsm` is missing, the script copies
from the bootstrap template and shows a **MsgBox** that patient data is **not**
restored — check `_backups\` or Windows File History.

Optional: `tools/Run-BuildRelease.ps1` — same via Excel COM with MsgBox auto-dismiss.

See `tools/BUILD_RELEASE_NOTES.md` for full detail.

### Manual path

1. VBA editor (Alt+F11): remove the old `MedParser` module, **Import
   `MedParser.bas`**, run **`SetupWorkbook`**, save. (`SetupWorkbook` running is
   the compile check.)
2. **One-time:** the two worksheet event handlers are PRE-installed in the sheet
   modules (SetupWorkbook does not modify the VBProject). If setting up a fresh
   workbook, paste them once:
   - **Medications** sheet module: `Worksheet_SelectionChange` +
     `Worksheet_Change` (protect "# of Prints") + `Worksheet_BeforeDoubleClick`
     (toggle Print? on col 16, block col 15).
   - **Label Previews** sheet module: `Worksheet_Activate` -> `PreviewAllLabels`.
3. Keep `scu_emblem.png` beside the workbook so the logo embeds on SetupWorkbook.

### Problems solved (2026-06-30 session)

| Issue | Fix |
|---|---|
| Runtime 1004 on header PasteSpecial | `MatchHeaderFormat` |
| Labels bleed to 2nd die-cut | `LABEL_WIDTH_PT` tuning (228 safe; now 242 for wider template — re-verify) |
| Blank label after each print | `PrintOut From:=1, To:=1`; `PrepareLabelSheetForPrint`; only `scuLogo` prints as shape |
| Emblem sliver (merged cells hid logo) | Logo slot **G2:H3**; text **A2:F2** / **A3:F3**; `ZOrder msoBringToFront` |
| Gallery OK, print sheet wrong logo | `RefreshPrintLabelLogo` on preview update + before each print |
| Logo squished wide | `InsertLabelLogo` inserts at natural aspect, not `maxW × height` box |
| Gray snake on thermal | `Build-ScuEmblem.ps1` forces solid black pixels |
| Clinic name clipped top/bottom | `FmtLblHeader` (no wrap) + row heights 20 + 10 pt; bottom/top vertical align |
| Clinic name too small | Wider text merge **A2:F2** (was A2:E2); `CLINIC_NAME_FONT_PRINT = 14` |
| Clinic name not bold enough | `FmtLblClinicName` — always bold via `Font.Bold` only |
| Compile: `xlBold` / `Font.Weight` | Removed — use **`Font.Bold = True`** only (Weight/xlBold break compile) |
| Weird partial logos after rebuild | Removed embedded `LogoB64()` fallback from `LogoFilePath`; reduced inset; in-band vertical centering |
| Build replaced workbook / lost data | Bootstrap only when `.xlsm` missing + warning MsgBox |
| VBScript / VBA compile | See table below |

### Compile / build fixes (2026-06-30)

| Issue | Fix |
|---|---|
| VBScript "Expected end of statement" | No typed `Dim` in `Build-Release.vbs` |
| VBA import / compile failures | Restore `FONT_LABEL_HDR_FB`; optional defaults must be literals; strip UTF-8 BOM; `msoPictureCompressNone` → `0` |
| `xlBold` / `Font.Weight` in VBA | Do not use — `Font.Bold` only |

---

## 6. TASKS STILL TO COMPLETE / VERIFY

1. **Add the consolidated folder as an Excel Trusted Location** (section 1) so
   macros run without prompts. Without this, content is blocked.
2. **(Optional) Desktop access:** make a Desktop shortcut to the folder; rename
   to drop `GIT_VERSION_` and re-point GitHub Desktop.
3. **`LABEL_WIDTH_PT = 242`** — user wanted a wider template; confirm on the clinic
   Brother that labels still fit **one** die-cut (228 pt was the prior no-bleed
   sweet spot).
4. **Bold clinic title on thermal** — bold may still look subtle on the Brother;
   `ShrinkToFit` can limit apparent size on long clinic name text.
5. **Physical print regression** — spot-check emblem size/position and header
   after the latest logo/header changes (user confirmed screen + print path OK
   at end of session).
6. **Bootstrap / data loss** — if `MedicationDispensing.xlsm` was ever missing
   during a build, the bootstrap copy overwrote it; recover from `_backups\` if
   needed.
7. **Dead code cleanup (optional):** `LogoB64()` remains in `MedParser.bas` but is
   unused after fallback removal — safe to delete later to shrink the module.

**Done / verified (2026-06-30):** release build (`Build-Release.vbs`); label layout
and typography; emblem (aspect, pure black, print + gallery sync); one label per
print (no blank follow-on label); logo/header layout on screen and print path
(user confirmed working at end of session).

---

## 7. Challenges & gotchas (for the next dev)

- **CRLF + ASCII required; no UTF-8 BOM.** Editing `.bas` on Linux/Mac (LF endings
  or BOM) makes the VBA importer reject the file. Always save CRLF + pure ASCII.
- **`MatchHeaderFormat` for new columns.** Do not raw `PasteSpecial` header formats
  onto new Medications/Log columns — use the helper (fixes runtime 1004).
- **`Shapes.AddPicture` / logo** — use local `scu_emblem.png` only; run
  `Build-ScuEmblem.ps1` after crop edits. Never insert at `maxWidth × height`
  (squashes aspect). Reserve **G2:H3** for the emblem; text in **A2:F2** /
  **A3:F3**; bring shape to front. Do **not** re-enable `LogoB64()` fallback.
- **Gallery vs print logo** — gallery rebuilds on tab activate; print sheet needs
  `RefreshPrintLabelLogo` on every `UpdateLabelPreviewForMedRow` / print.
- **Blank label between prints** — if it returns, check Excel print preview page
  count; ensure `PrintOut From:=1, To:=1` and preview buttons have
  `PrintObject = False`.
- **`Font.Bold` only for bold headers.** Do not use `Font.Weight` or `xlBold` —
  they caused compile errors in this project.
- **No runtime VBProject injection.** Event handlers are preinstalled in the sheet
  modules; `SetupWorkbook` does not touch the VBProject (more robust, no trust
  setting needed for that).
- **`PaperSize = xlPaperUser` throws 1004** on this Brother driver - never set it;
  the driver's DK-1202 default is correct.
- **OneDrive dehydration:** OneDrive can mark files "cloud-only" (cloud icon).
  They re-download on open. Set the folder to "Always keep on this device" to stop
  it (helps automation/editing).
- **OneDrive + Git:** never run non-Windows Git/automation against a `.git` folder
  inside OneDrive (lock errors, can corrupt `.git/config`). And OneDrive blocks
  *moving* a folder that contains `.git` (error 0x80004005) - use GitHub Desktop on
  native Windows, and don't try to relocate the repo folder via Explorer.
- **Macros blocked by Mark-of-the-Web** when the file lives in OneDrive - fixed via
  Trusted Location.

---

## 8. Key VBA routines (`MedParser.bas`)

| Routine | Purpose |
|---|---|
| `SetupWorkbook` | Run after each import: rebuilds buttons, headers, colors, label layout; calls `PreviewAllLabels`. Uses `MatchHeaderFormat` for new columns. |
| `ParseMedications` / `SplitMedBlocks` / `IsMedHeaderLine` / `ParseOneBlock` | Parse pasted text -> one row per med. |
| `ValidateMedications` / `ReviewMedications` | Flag issues; summary. |
| `AddMedicationRow` / `RemoveSelectedMedication` / `RenumberMeds` | List editing. Remove Selected removes all checked rows. |
| `ApplyRowState` / `ApplyAllRowStates` / `IsRowSelected` / `ToggleRowSelect` | Row color by state; confidence triad; selection checkmark. |
| `PrintCheckedLabels` | Batch-print every checked med; logs each. |
| `BuildLabelPreviewLayout` / `ApplyLabelContentWidth` / `ApplyLabelPageSetup` / `UpdateLabelPreviewForMedRow` / `FmtLbl` / `FmtLblHeader` / `FmtLblClinicName` | Label layout, width scaling (`LABEL_WIDTH_PT = 242`), page setup, value writing, bold clinic name. |
| `InsertLabelLogo` / `RefreshPrintLabelLogo` / `EnsurePrintLabelHeaderLayout` / `LogoFilePath` | Emblem insert (30 pt print / 28 pt gallery), refresh before print, header merges A2:F2 / G2:H3. |
| `PrepareLabelSheetForPrint` / `PrintLabelSurfaceSafe` | Suppress extra shapes; single-page `PrintOut From:=1, To:=1`. |
| `LabelHeaderFont` / `SetMiniValue` / `MedFontSize` / `NameFontSize` / `PatientNameFontSize` / `SigFontSize` | Fonts and adaptive sizing. |
| `MatchHeaderFormat` | Safe header format copy (PasteSpecial 1004 fix). |
| `PrintLabel` / `PrintCheckedLabels` / `SelectBrotherPrinter` / `MarkPrinted` | Auto-select Brother, confirm, print, bump # of Prints. |
| `LogPrint(medRow, vol)` / `AskInitials` | Full per-print Log row (single + batch). |
| `StartNewPatient` | Clear patient + meds, keep the Log. |
| `PreviewAllLabels` / `BuildAllLabelsPreview` / `EnsureAllLabelsSheet` | Label Previews gallery. |
| `ResetSession` / `ClearPasteArea` | Full reset (incl. Log) / clear paste box. |

---

## 9. HIPAA / PHI

Patient name/DOB/meds are PHI. Keep the `.xlsm` and the dispense Log **local** -
never commit them to GitHub (the `.gitignore` excludes `*.xlsm`/`*.xlsx`/`*.csv`
and `_backups/`) and never route PHI through an external AI/API without a signed
BAA. The repo is code + docs + no-PHI samples only.

---

**Bottom line:** parsing, validation, checkbox-driven selection/printing/removal,
the dispense log, multi-patient flow, and the **redesigned DK-1202 label** (emblem,
header typography, width, single-page print) are built and **verified on screen
and print path** at end of the 2026-06-30 session. Everything lives in one
folder; code is backed up to GitHub. Remaining: **Trusted Location**, optional
**242 pt width re-test** on physical Brother, and optional cleanup of unused
`LogoB64()`. Detail: `LABEL_REDESIGN.md`, `CHANGELOG.md`,
`tools/BUILD_RELEASE_NOTES.md`.
