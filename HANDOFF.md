# SCU Label Printing Tool - Handoff

Last updated: 2026-06-30

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
| Clinic identity | A2:H2 |
| Contact line | A3:H3 |
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
| Header (rows 2–3) | Shrink-to-fit; Helvetica → Arial |

**Print width:** `LABEL_WIDTH_PT = 228` scales columns A:H via
`ApplyLabelContentWidth` so content fits **one** DK-1202 die-cut (fixes bleed
onto a second label). Tune this constant if physical output is still too wide
or narrow.

**Logo / emblem:**

| Item | Detail |
|---|---|
| Manual source crop | `cropped_Black SCU Logo + Transparent Background - Copy.png` |
| Workbook file | `scu_emblem.png` (copy via `tools/Build-ScuEmblem.ps1`) |
| Aspect ratio | `LOGO_ASPECT = 1.488` (6150×4133 px) |
| On-label height | **28 pt** in `InsertLabelLogo` (top-right anchor) |
| Insertion | `AddPicture2` with compress=`0`; fallback `AddPicture` |
| Resolution | `LogoFilePath()` prefers local `scu_emblem.png`; else embedded `LogoB64()` temp file |

Do **not** re-enable automatic logo cropping without explicit approval — the
manual crop is the source of truth.

**Known issue:** build compiles and prints, but the emblem may still appear
**too large or small** on the physical label. Next tuning: `InsertLabelLogo`
height, anchor placement, `LOGO_ASPECT`, or Brother driver scale.

**Print geometry:** `Zoom = 100`, no fit-to-page, **no `.PaperSize`** (Brother
driver controls media), 0.04 in margins, landscape. Print areas in `PrintLabel`
and batch print paths use `A1:H15`.

**SetupWorkbook fix:** header `PasteSpecial` runtime 1004 resolved via
`MatchHeaderFormat` (safe format copy for new Medications/Log columns).

See also `LABEL_REDESIGN.md` and `CHANGELOG.md` (Unreleased).

---

## 5. Applying code changes (release build)

`MedParser.bas` must stay **pure ASCII with Windows CRLF** line endings or Excel
rejects the import. **No UTF-8 BOM** on line 1. (`.gitattributes` enforces CRLF
for `.bas` / `.vbs`.)

### Quick path (recommended)

1. Close `MedicationDispensing.xlsm`.
2. (If emblem changed) run `tools/Build-ScuEmblem.ps1` → refreshes `scu_emblem.png`.
3. Double-click **`Build-Release.vbs`** (needs **Trust access to the VBA project
   object model** + folder as **Trusted Location**).
4. Click OK on **Setup complete!** and **Release build complete**.

`Build-Release.vbs` opens `MedicationDispensing.xlsm` in place (bootstraps from
`Broken_PrettyPrint_MedicationDispensing.xlsm` only if missing), removes the old
`MedParser` module, imports `MedParser.bas`, runs
`'MedicationDispensing.xlsm'!SetupWorkbook`, saves.

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

### Recent compile / build fixes (2026-06-30)

| Issue | Fix |
|---|---|
| Runtime 1004 on header PasteSpecial | `MatchHeaderFormat` |
| VBScript "Expected end of statement" | No typed `Dim` in `Build-Release.vbs` |
| VBA import / compile failures | Restore `FONT_LABEL_HDR_FB`; optional defaults must be literals; strip UTF-8 BOM; `msoPictureCompressNone` → `0` |

---

## 6. TASKS STILL TO COMPLETE

1. **Logo sizing on physical label** — build compiles; emblem still not sized
   appropriately on Brother output. Tune `InsertLabelLogo` (28 pt height), anchor,
   `LOGO_ASPECT`, or driver scale after a test print.
2. **Add the consolidated folder as an Excel Trusted Location** (section 1) so
   macros run without prompts. Without this, content is blocked.
3. **Physical test print** on the Brother QL-1100C — confirm one label per
   die-cut, landscape, readable, not clipped. Adjust `LABEL_WIDTH_PT` (228) or
   margins in `BuildLabelPreviewLayout` if width is off.
4. **(Optional) Desktop access:** make a Desktop shortcut to the folder; rename
   to drop `GIT_VERSION_` and re-point GitHub Desktop.

**Done since last handoff:** `Build-Release.vbs` runs successfully; label width
tuning, header/patient typography, manual emblem workflow, and compile fixes are
in `MedParser.bas` (see `CHANGELOG.md` Unreleased).

---

## 7. Challenges & gotchas (for the next dev)

- **CRLF + ASCII required; no UTF-8 BOM.** Editing `.bas` on Linux/Mac (LF endings
  or BOM) makes the VBA importer reject the file. Always save CRLF + pure ASCII.
- **`MatchHeaderFormat` for new columns.** Do not raw `PasteSpecial` header formats
  onto new Medications/Log columns — use the helper (fixes runtime 1004).
- **`Shapes.AddPicture` / logo** — wrapped in error handling; prefer local
  `scu_emblem.png`. Manual crop is authoritative; auto-crop scripts are disabled.
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
| `SetupWorkbook` | Run after each import: rebuilds buttons, headers, colors, label layout. Uses `MatchHeaderFormat` for new columns. |
| `ParseMedications` / `SplitMedBlocks` / `IsMedHeaderLine` / `ParseOneBlock` | Parse pasted text -> one row per med. |
| `ValidateMedications` / `ReviewMedications` | Flag issues; summary. |
| `AddMedicationRow` / `RemoveSelectedMedication` / `RenumberMeds` | List editing. Remove Selected removes all checked rows. |
| `ApplyRowState` / `ApplyAllRowStates` / `IsRowSelected` / `ToggleRowSelect` | Row color by state; confidence triad; selection checkmark. |
| `PrintCheckedLabels` | Batch-print every checked med; logs each. |
| `BuildLabelPreviewLayout` / `ApplyLabelContentWidth` / `UpdateLabelPreviewFromSelection` / `FmtLbl` | Label layout, width scaling (`LABEL_WIDTH_PT`), value writing. |
| `InsertLabelLogo` / `LogoFilePath` / `LogoB64` | Emblem insert (28 pt), local PNG or embedded fallback. |
| `LabelHeaderFont` / `SetMiniValue` / `MedFontSize` / `NameFontSize` / `PatientNameFontSize` / `SigFontSize` | Fonts and adaptive sizing. |
| `MatchHeaderFormat` | Safe header format copy (PasteSpecial 1004 fix). |
| `PrintLabel` / `SelectBrotherPrinter` / `MarkPrinted` | Auto-select Brother, confirm, print, bump # of Prints. |
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
the dispense log, multi-patient flow, label width tuning, typography updates,
and the **release build script** are in place and compile. Everything lives in
one folder; code is backed up to GitHub. Remaining: **emblem sizing on physical
print**, Trusted Location, and Brother test print (section 6). Detail:
`LABEL_REDESIGN.md`, `CHANGELOG.md`, `tools/BUILD_RELEASE_NOTES.md`.
