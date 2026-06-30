# SCU Label Printing Tool - Handoff

Last updated: 2026-06-30

An Excel + VBA tool for the **Saturday Clinic for the Uninsured (SCU)** free
pharmacy. It turns pasted prescription text into a validated medication table and
prints professional labels on a **Brother QL-1100C** (DK-1202, 62 x 100 mm,
landscape). It runs **100% offline** and is built to be robust for high-turnover
volunteers.

---

## 1. Where everything lives (IMPORTANT - read first)

Everything was consolidated into **one folder**:

```
C:\Users\ringo\OneDrive\Documents\GitHub\GIT_VERSION_SCU Label Printing
```

This single folder contains the code, docs, the SC emblem, the workbook
(`MedicationDispensing.xlsm`), and `_backups\`. It is also the local Git repo
(remote: `JamesSRN/scu-label-tool`). The `.xlsm` and `_backups\` are git-ignored,
so patient data can never be committed.

**Two open setup items because of the consolidation:**

1. **Trusted Location.** The old `Desktop\SCU Label Printing` folder (the prior
   Excel Trusted Location) was removed. The new folder is **not trusted yet**, so
   macros will be blocked until you add it:
   `Excel > File > Options > Trust Center > Trust Center Settings > Trusted
   Locations > Add new location > browse to the folder above > OK`.
2. **Location.** It sits in `Documents\GitHub` rather than the Desktop because
   OneDrive refuses to move a folder that contains a `.git` directory (error
   0x80004005). For easy access: right-click the folder > Send to > Desktop
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
PATIENT                                  DOB   03/14/1985
Doe, Jonathan                            Rx    06/29/2026

Metronidazole 500 mg          <- HERO line (largest, bold)
Tablet  .  Qty 56

DIRECTIONS
Take 1 tablet by mouth twice daily with food.
-----------------------------------------------------------
EXP  05/2027                   LOT  ABC1234
```

Cell map (rows 16-18 hold off-label helper text that does not print):

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

**Design / type:** Hierarchy is carried by size + weight + small uppercase labels
(never color - thermal is monochrome). Font is **Arial** (the reliable
Helvetica substitute on Windows; switch the `.Name = "Arial"` lines in `FmtLbl` /
`SetMiniValue` to Segoe UI for a more modern feel). Two thin hairline rules (under
the header, above the footer); no heavy box. The medication name is the focal
point; EXP/LOT are pinned in a footer so they're always in the same place.

**Logo:** the SC emblem was extracted from the full logo (mark only) and made a
crisp all-black PNG, `scu_emblem.png`, kept in the folder. It is inserted with
`SaveWithDocument = msoTrue` (embedded in the workbook, so it survives even if the
PNG is later moved/dehydrated) and wrapped in error handling (a missing logo can
never break setup or printing).

**Print geometry:** content height = 170 pt = 2.361 in, which exactly matches the
DK-1202 printable height (62 mm minus 0.08 in margins). Width unchanged from the
prior working layout. `Zoom = 100`, no fit-to-page, **no `.PaperSize`** (the
Brother driver controls media), 0.04 in margins, landscape. Print areas in
`PrintLabel` and `PrintCheckedLabels` were updated to `A1:H15`.

**Edge cases:** long medication names step 14 -> 12.5 -> 11 -> 10 pt; long patient
names 13 -> 11 -> 9.5 pt; long directions 10.5 -> 9.5 -> 8.5 pt with a 36 pt SIG
block (multi-line insulin/tapers wrap cleanly).

---

## 5. Applying code changes (release build)

`MedParser.bas` must stay **pure ASCII with Windows CRLF** line endings or Excel
rejects the import. (`.gitattributes` enforces CRLF for `.bas`.) To apply:

1. VBA editor (Alt+F11): remove the old `MedParser` module, **Import
   `MedParser.bas`**, run **`SetupWorkbook`**, save. (`SetupWorkbook` running is
   the compile check.)
2. **One-time:** the two worksheet event handlers are PRE-installed in the sheet
   modules (SetupWorkbook no longer modifies the VBProject). If setting up a fresh
   workbook, paste them once:
   - **Medications** sheet module: `Worksheet_SelectionChange` +
     `Worksheet_Change` (protect "# of Prints") + `Worksheet_BeforeDoubleClick`
     (toggle Print? on col 16, block col 15).
   - **Label Previews** sheet module: `Worksheet_Activate` -> `PreviewAllLabels`.
3. Make sure `scu_emblem.png` is in the same folder as the workbook (it is) so the
   logo embeds on first SetupWorkbook run.

`Build-Release.vbs` automates import + SetupWorkbook + save (needs "Trust access
to the VBA project object model" and the workbook closed).

---

## 6. TASKS STILL TO COMPLETE

1. **Apply the redesign to the live workbook:** import the new `MedParser.bas`,
   run `SetupWorkbook`, save. Then preview a med to see the new label.
2. **Add the consolidated folder as an Excel Trusted Location** (section 1) so
   macros run. Without this, content is blocked.
3. **(Optional) Desktop access:** make a Desktop shortcut to the folder; rename it
   to drop `GIT_VERSION_` and re-point GitHub Desktop.
4. **Physical test print** on the Brother QL-1100C - the one thing never verified.
   Confirm the label fills DK-1202, landscape, readable, not clipped or squished.
   If off, adjust margins/centering in `BuildLabelPreviewLayout` + the print subs.
5. **Commit this updated HANDOFF.md** to git (and the workbook stays git-ignored).

---

## 7. Challenges & gotchas (for the next dev)

- **CRLF + ASCII required.** Editing `.bas` on Linux/Mac (LF endings) makes the VBA
  importer reject it with "unable to import file." Always save CRLF + pure ASCII.
- **`Shapes.AddPicture` can abort setup** with a bad/oversized/cloud-only image -
  always wrap it in `On Error Resume Next` and use a small local PNG.
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
| `SetupWorkbook` | Run after each import: rebuilds buttons, headers, colors, the label layout. No VBProject access. |
| `ParseMedications` / `SplitMedBlocks` / `IsMedHeaderLine` / `ParseOneBlock` | Parse pasted text -> one row per med. |
| `ValidateMedications` / `ReviewMedications` | Flag issues; summary. |
| `AddMedicationRow` / `RemoveSelectedMedication` / `RenumberMeds` | List editing. Remove Selected removes all checked rows. |
| `ApplyRowState` / `ApplyAllRowStates` / `IsRowSelected` / `ToggleRowSelect` | Row color by state; confidence triad; selection checkmark. |
| `PrintCheckedLabels` | Batch-print every checked med; logs each. |
| `BuildLabelPreviewLayout` / `UpdateLabelPreviewFromSelection` / `FmtLbl` | **Redesigned** label layout + value writing. |
| `SetMiniValue` / `MedFontSize` / `NameFontSize` / `SigFontSize` | Label helpers: footer mini-label+value; adaptive sizing. |
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
the dispense log, multi-patient flow, and the **redesigned professional label**
are all built and (except the physical print) verified. Everything is in one
folder and the code is backed up to GitHub. Remaining: apply the redesign to the
live workbook, add the Trusted Location, and do the physical test print (section 6).
