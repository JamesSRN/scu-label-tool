# SCU Medication Dispensing & Label Printing — Handoff

_Last updated: 2026-06-28 (evening)_

## 1. What this project is & the goal

A medication dispensing + label-printing tool for the **Saturday Clinic for the Uninsured (SCU)** pharmacy, built so that **high-turnover volunteers** can do it reliably with built-in error-catching. Core workflow:

```
Tebra medication order  →  copy text  →  paste into Excel  →  parser extracts
structured fields  →  validate + fill Expiration/Lot  →  select meds  →
print labels on the Brother QL-1100c (DK-1202, 62 x 100 mm)
```

**Design goals:** robust for non-technical volunteers; catch and surface errors (missing/uncertain fields); one label per medication; print one, several, or all meds; everything runs **100% locally** in Excel + VBA (no PHI leaves the machine — see §9 HIPAA).

`MedParser.bas` is the **single source of truth** for the code. The workbook (`.xlsm`) is built from it (see §6).

---

## 2. Files (in `Desktop\SCU Label Printing\`)

| File | Role |
|---|---|
| **`MedicationDispensing.xlsm`** | **The working file.** Macro-enabled workbook. |
| **`MedParser.bas`** | VBA source module — the source of truth. ~101 KB. |
| **`Build-Release.vbs`** | One-double-click script: re-imports `MedParser.bas` → runs `SetupWorkbook` → saves. See §6. |
| `MedicationDispensing.xlsx` | Old base file, superseded. Ignore. |
| `_backups\` | Timestamped `MedParser.bas` backups at each milestone (e.g. `_pre_select`, `_pre_landscape2`, `_pre_autorefresh`). |
| `label_preview_landscape.png` | To-scale label mockup. |
| `HANDOFF.md` | This document. |

Also on the Desktop: the clinic's **old DYMO** templates (`dispensary label template DYMO.dymo`, `lab label template DYMO.dymo`) — reference only; we moved to Brother.

---

## 3. Workbook structure (6 tabs)

- **Patient & Input** — patient name / DOB / Rx date, the paste box; buttons: `PARSE MEDICATIONS`, `Clear Paste Area`, `Reset Session`, **`Start NEW Patient`** (clears the current patient + med list but **keeps the Log** — for seeing patients back-to-back).
- **Medications** — the parsed table. Columns A–N: `#`, Name, Strength, Dosage Form, SIG, Quantity, Refills, **Expiration**, **Lot**, Date of Rx, Confidence, Warnings, Raw Text, **Printed?**; then **column O = "Prints"** (how many times that label has printed, auto-increments) and **column P = "Print?"** (double-click toggles a ✓ to select for batch print). Buttons: `+ Add Medication`, `- Remove Selected`, `Review & Validate`, `Print Selected Label`, `Preview ALL Labels`, `Print Checked Labels`.
- **Label Preview** — renders one label in the **Brother LANDSCAPE** layout (≈3.9″ × 2.4″); buttons `Update Label Preview`, `Print This Label`.
- **All Labels** — auto-generated gallery: one card per medication, each with **Print this label / Edit this med / Remove this med** buttons. **Rebuilds itself every time you click onto the tab.**
- **Log** — running dispense log.
- **Setup & Help** — instructions.

Keyboard shortcuts: `Ctrl+Shift+P` parse, `Ctrl+Shift+R` reset, `Ctrl+Shift+L` update label preview.

---

## 4. Feature set — what the tool does now

### Parsing (done, tested)
- **Block splitter** (`IsMedHeaderLine`) starts a new medication on any drug-header line (name+strength, a form-headed name like "Basaglar KwikPen", or a bare capitalized name) **even with no blank line between drugs**. A real paste that used to collapse to 6 rows now correctly yields **12**.
- **Confidence** is set per drug from how cleanly it parsed: starts High, collects a "warning" for each uncertain/missing field → **0 warnings = High, 1–2 = Medium, 3+ = Low** (manually-added = "Manual").
- **Validation** (`ValidateMedications`) flags missing strength/SIG/quantity/Exp/Lot, bad Exp format, and possible duplicates → Warnings column.
- **Review** (`ReviewMedications`) summarizes every med + its issues.
- **Add / Remove / Renumber** medications manually.
- Exp/Lot forced to **text** so `05/2027` stays literal.

### Labels & printing
- **Label Preview** = single label, **landscape** DK-1202 (final orientation choice).
- **Print Selected Label** (Medications tab) prints the clicked row.
- **Preview ALL Labels** = gallery of every label with per-card **Print / Edit / Remove**; auto-refreshes on tab activate.
- **Print? selection column (O)** — **double-click a cell to toggle a ✓**. **Print Checked Labels** prints every checked med in sequence (skips any checked row missing Exp/Lot), one confirmation up front.
- **Brother auto-select** (`SelectBrotherPrinter`): WMI finds the Brother QL-1100c and sets it active (tries name+port, then `Ne00:`–`Ne99:`, then common names); single-print confirms before committing a label; graceful "not found" fallback.
- **Clinic logo** embedded **top-right of every printed label**; a clean full-width rule sits under the clinic name (replaced the old gappy line).
- **Print count** — the `Prints` column auto-increments on every print (single, per-card, or batch).

### Dispense Log & multi-patient
- **Every print is logged** to the Log sheet: timestamp, patient, DOB, medication, strength, SIG, quantity, refills, Exp, Lot, Rx date, **dosage form**, **print #**, and **volunteer initials**. *(Fixed a real bug where the medication details were never actually recorded; batch prints are now logged too; initials are asked once per batch.)*
- **`Start NEW Patient`** clears the current patient + med list but **never touches the Log**, so the dispense record accumulates across patients all session.

### Color scheme (latest design)
- **Confidence cell only** is colored by the triad: High=green, Medium=yellow, Low=red, Manual/blank=gray.
- **Row background = workflow STATE** (priority top→bottom): **Selected** (checked) = blue · **Printed** = lavender · **Validated** (no warnings) = green · **Non-validated** (has warnings) = gray. Updates in real time when you toggle a row's ✓ (via an injected `Worksheet_BeforeDoubleClick`).
- **Exp/Lot cells:** **red when missing**, otherwise **blend into the row color** (the old "orange when filled" was removed per request).
- **All Labels gallery:** a **selected (checked) med's preview card fills light green** each time the gallery rebuilds (reads the live ✓ state).

---

## 5. CURRENT STATE — read this first

- ✅ **Brother QL-1100c driver is installed.** It shows in Excel's printer list. The driver's default label is **DK-1202 (2.4″ × 3.9″ = 62 × 100 mm)** — the 4×6 default was corrected.
- ✅ **Macros run.** They were being **blocked** (see §8); fixed by adding the **`SCU Label Printing` folder as an Excel Trusted Location**, and "Trust access to the VBA project object model" is enabled.
- ✅ **Live & working** in the `.xlsm`: full parser, validation, add/remove, landscape Label Preview, Print Selected, Preview ALL Labels gallery + auto-refresh.
- 🟡 **The newest `MedParser.bas` is on disk but NOT yet re-imported.** One release build (§6) makes all of these live: **clinic logo (top-right) + fixed header rule**, **`Prints` print-count column**, **`Start NEW Patient`** button/workflow, and the **fully-fixed dispense Log** (now actually records the medication on every print, logs batch prints, captures dosage form + print # + initials). Earlier pending items from before — Print? selection column, state-based coloring, Print Checked Labels, Exp/Lot "no orange when filled" — also apply in the same run. If the Medications sheet has no `Prints`/`Print?` columns or the Patient tab has no `Start NEW Patient` button, run §6.
- ⏳ **Physical test print on the Brother: NOT yet verified.** This is the main open item (§7).

---

## 6. Applying code changes (release build)

Whenever `MedParser.bas` changes, re-sync the workbook. **Acceptance:** newest `.bas` imported → project compiles → `SetupWorkbook` run → saved as `.xlsm`.

**Fastest — `Build-Release.vbs`:** close the workbook, then double-click the script (the Trusted-Location + VBA-project-access prerequisites are already set). Click OK on "Setup complete".

**Manual:** Enable Content → **Alt+F11** → Project Explorer → Modules → right-click **MedParser → Remove → No** → **File ▸ Import File ▸ `MedParser.bas`** → **Debug ▸ Compile** → **Macros ▸ SetupWorkbook ▸ Run** → **Ctrl+S**.

Running `SetupWorkbook` successfully *is* the compile check (VBA won't run a sub if the project has compile errors). `SetupWorkbook` also injects the two worksheet event handlers (`Worksheet_Activate` on All Labels; `Worksheet_BeforeDoubleClick` on Medications) — this needs "Trust access to the VBA project object model" (already on).

> **Status line:** as of the last release build, `MedicationDispensing.xlsm` reflects `MedParser.bas`. If the Print?/colors changes aren't visible, `SetupWorkbook` hasn't been run against the latest import yet — do §6.

---

## 7. NEXT STEPS

1. **Run `SetupWorkbook`** against the latest import (§6) so the Print? column, "Print Checked Labels" button, double-click selection, and new colors appear. Save.
2. **Physical test print** *(the key remaining validation)*: Medications tab → click a med row → **Print Selected Label** (or a card's **Print this label**) → confirm the Brother → print. Check the label **fills the DK-1202 stock, landscape, readable, not clipped**. If off, adjust margins (`0.04"`) / centering in `PrintLabel` + `BuildLabelPreviewLayout`, re-import, retest.
3. **Test batch print**: check a few rows' **Print?** cells (double-click) → **Print Checked Labels** → confirm sequence prints, printed rows turn lavender.
4. **(Optional)** literal clickable checkbox controls instead of the ✓ toggle, if preferred (toggle was chosen for reliability).

---

## 8. Challenges encountered & how they were solved (for the next dev)

- **The parser had never actually run** before this work. Fixing it surfaced a chain of runtime bugs, each only appearing after the previous was fixed:
  - **Error 13 (type mismatch):** `Array()` assigned to `Dim ...() As String` → changed those (`verbs`, `sigWords`, `pharmas`) to `As Variant`.
  - **Error 5017 (regex):** `VBScript.RegExp` has **no look-behind** `(?<!...)`; rewrote 4 patterns with `\b`.
  - **Latent compile error:** `Dim input As String` (`Input` is reserved) → `inText`.
  - **Mojibake buttons:** `.bas` was UTF-8 but VBA imports ANSI → converted the whole file to **pure ASCII** (keep it ASCII-only).
- **Macros blocked by "Mark of the Web."** Because the file lives in OneDrive, Office **blocked all macros** (Info bar showed *"active content blocked"* with only a Trust Center link, no Enable button), compounded by a strict "disable except digitally signed" macro policy. **Fix:** add the folder as a **Trusted Location** (Trust Center) and reopen. (Unblocking the single file via Explorer → Properties → Unblock also works.)
- **`PaperSize = xlPaperUser` threw run-time 1004** ("Unable to set the PaperSize property") on this Brother driver. **Fix:** removed all `.PaperSize = xlPaperUser` lines — the driver's DK-1202 default is already correct, so Excel just inherits it. (If you re-add paper handling, wrap it in `On Error Resume Next`.)
- **Orientation flip-flop:** built portrait → switched to landscape → back to portrait → finally **landscape** (volunteers wanted the wider layout, matching the old DYMO). The layout now lives in code (`BuildLabelPreviewLayout`) so it's reproducible.
- **OneDrive AutoSave** pops "Syncing workbook…" dialogs and can momentarily interfere; just let it finish or Cancel the dialog. Excel also **caches the printer list at launch**, so a newly installed printer needs an Excel restart to appear.
- **Gallery vs. checkboxes:** the All Labels gallery auto-rebuilds (would wipe checkboxes), so selection lives on the **Medications** sheet instead.

---

## 9. Known limitations / things to watch

- **Parser field accuracy** on unusual entries isn't perfect (e.g., insulin SIG, a "chewable tablet" prefix). Acceptable because rows are **flagged and fully editable** before printing — validation is the safety net.
- **Windows-only:** uses `VBScript.RegExp` (no Mac Excel). Clinic is Windows, so fine.
- **Date-of-Rx placeholder** on Label Preview can show a date serial when idle; real prints use the macro-written formatted date.
- **HIPAA:** patient name/DOB/meds are PHI. Keep this **local**. Do **not** route PHI through an external AI API without a signed **BAA** (enterprise contract). The existing Teams/Power Automate flow only creates per-patient chats — don't add external-AI parsing without a BAA.

---

## 10. Key VBA routines (in `MedParser.bas`)

| Routine | Purpose |
|---|---|
| `SetupWorkbook` | Run after every import: (re)creates all buttons, the Print? header, injects the two worksheet events, recolors rows, rebuilds the label layout. |
| `ParseMedications` · `SplitMedBlocks` · `IsMedHeaderLine` · `ParseOneBlock` | Parse pasted text → one row per medication. |
| `ValidateMedications` · `ReviewMedications` | Flag issues; summary. |
| `AddMedicationRow` · `RemoveSelectedMedication` · `RenumberMeds` | Manual list editing. |
| `ApplyRowState` · `ApplyAllRowStates` · `IsRowSelected` · `ToggleRowSelect` | **New color engine** — row background by state, confidence-cell triad, Exp/Lot highlight, selection ✓. |
| `InstallMedSheetEvents` | Injects `Worksheet_BeforeDoubleClick` (toggle Print? ✓) into the Medications sheet module. |
| `PrintCheckedLabels` | Batch-prints every checked (✓) medication in sequence. |
| `UpdateLabelPreviewFromSelection` · `BuildLabelPreviewLayout` · `FmtLbl` | Single label preview + **landscape** layout (in code). |
| `PrintLabel` · `SelectBrotherPrinter` · `MarkPrinted` | Auto-select Brother, confirm, print, mark printed + bump the `Prints` count. |
| `LogPrint(medRow, vol)` · `AskInitials` | Write one full Log row per print (timestamp, patient, DOB, med, strength, SIG, qty, refills, Exp, Lot, Rx date, dosage form, print #, initials). Called by both single and batch printing; takes the med row explicitly so details are always captured. |
| `StartNewPatient` | Clear current patient + med list (with confirm), reset Exp/Lot to text — **but leave the Log intact** for the next patient. |
| `PreviewAllLabels` · `BuildAllLabelsPreview` · `EnsureAllLabelsSheet` · `InstallAutoRefresh` | All Labels gallery + injects `Worksheet_Activate` so it auto-rebuilds on tab click. |
| `RowPrint` · `RowEdit` · `RowRemove` · `CallerRow` | Per-card gallery handlers (`Application.Caller`). |
| `ResetSession` · `ClearPasteArea` | Clear between patients. |

---

## 11. Version control (git / GitHub) — code-only, PHI excluded

**Rule #1 — never commit PHI.** The `.xlsm`/`.xlsx` workbooks hold patient names, DOB, and meds. GitHub is **not** BAA-covered, so those files must **never** be pushed. A `.gitignore` is already in the folder excluding `*.xlsm`, `*.xlsx`, `*.xls`, `*.csv`, `~$*`, and `_backups/`. The repo holds **only**: `MedParser.bas`, `Build-Release.vbs`, `HANDOFF.md`, the logo PNG, and `.gitignore`.

**Rule #2 — don't git-init inside OneDrive.** OneDrive and git fight over the `.git` folder (this was tried and failed). Put the repo **outside** the synced folder.

**Setup with GitHub Desktop (now installed):**
1. **Delete the leftover broken `.git` folder** in `SCU Label Printing`: File Explorer → View → Show → **Hidden items**, then delete `.git` (or in a Command Prompt there: `rmdir /s /q .git`).
2. Make a new folder outside OneDrive, e.g. `C:\Users\<you>\source\scu-label-tool`. Copy in `MedParser.bas`, `Build-Release.vbs`, `HANDOFF.md`, the logo PNG, and `.gitignore`.
3. GitHub Desktop → **File ▸ Add local repository** → pick that folder → "create a repository" → name it, keep `.gitignore`, **Publish** as **Private**.
4. To sync going forward: edit `MedParser.bas` in that folder (or copy the updated one in), then in GitHub Desktop **Commit** + **Push**. Keep developing the `.xlsm` in OneDrive; only the **code** lives in git.

---

**Bottom line:** parsing, validation, editing, the landscape single label, the all-labels gallery, selection + per/batch printing, and the new color scheme are all built and (except the new selection/colors awaiting one `SetupWorkbook` run, and the physical print) verified. The driver is installed and macros run. The one true unknown left is **how a label physically prints on the Brother** — do §7 step 2.
