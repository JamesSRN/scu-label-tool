# Changelog

## Unreleased

### V2 build

Work toward the V2 build, applied in blocks (see `IMPROVEMENT_REPORT_2026-07-09.md`).

**Developer Test panel (2026-07-09)**
- New **Developer Test** sheet (last tab, rebuilt by `SetupWorkbook`) with testing buttons: **Generate Test Patient** (random name, DOB, and 3-5 realistic meds into the paste box - Exp/Lot/Source left blank - ready to Parse), **Fill Random Exp/Lot/Source** (fast path to a printable state), **Run Parser Self-Tests**, and **Reset Session**.

**Workflow simplification: one print path (2026-07-09)**
- **Print is now done only from the gallery.** Removed **Print Checked Labels** from the Medications tab and **Print This Label** from the preview surface; the gallery's **Print Checked Labels** is the single print action (Preview ALL Labels -> Print Checked Labels).
- Removed the **Reprint Last Batch** button (every print already does 2 copies).
- Renamed the **Print?** column to **Check Med** (structure self-check + Start Here guide + prompts updated to match).
- **Edit -> reprint prints only CHECKED meds.** Reopening an encounter now pre-checks every med; uncheck any you won't reprint before saving, and the reprint (2 copies each) prints just the checked ones.

**Edit a past Encounter (2026-07-09)**
- New **Edit Past Encounter** / **Save Edited Encounter** buttons on the Medications tab. Each print now saves a **full snapshot** of the patient's med list under its Encounter # (hidden `EncounterData` sheet). Reopening an encounter (pick from a list by number) restores the patient + every med exactly (all fields, incl. lot/exp/source), where you can add, remove, or fix meds.
- **Save replaces** that encounter's Log rows and snapshot in place (same number, no duplicates) and **rebuilds the Tebra note** from the corrected Log. It then **asks whether to reprint** the corrected labels (2 copies each, no re-logging). Edited Log rows are marked "(edited)" in the Initials column; the dated CSV keeps history (append-only).
- Snapshots share the Log's lifecycle: kept across **Start NEW Patient**, wiped on **full reset / on close**. New Log constant map (`LG_*`) and snapshot map (`ES_*`) drive all Log/snapshot reads and writes.

**Encounter column moved next to Timestamp (2026-07-09)**
- In the Log (and the dated CSV), **Encounter** moved from the far right to **column 2, right after Timestamp**; all other columns shifted right one. The whole Log now runs through an `LG_*` constant map (header row, LogPrint, Tebra, CSV, NextEncounter), so future Log reorders are a one-line change.

**Block 1 - source pre-check (2026-07-09)**
- Added `tools/check-encoding.ps1`: validates `MedParser.bas` and `Build-Release.vbs` are pure ASCII, have no BOM, use Windows CRLF (no lone LF / doubled CR), and have balanced block keywords. Encoding problems exit 1 (build-breaking); balance is advisory (exit 3) since the VBA compile is the authoritative structural check.
- `Build-Release.vbs` now runs that check first and **aborts the build only on an encoding failure** (fail-safe: if PowerShell or the script is unavailable, or only a balance advisory is returned, the build proceeds).

**Block 2 - print workflow UX (2026-07-09)**
- **Skipped labels are now named.** The Print Complete dialog lists each medication skipped for a missing Expiration/Lot (by name), instead of only a count, and uses a warning icon when anything was skipped (`PrintCheckedLabels`).
- **Reprint Last Batch.** New `ReprintLastBatch` action + gallery button reprints the last successfully-printed batch (same rows + initials, logged as a reprint) for paper-jam/misfeed recovery. The remembered batch is **cleared on any session reset or Start New Patient**, so it can never reprint a previous patient's labels. Session-only (module vars `gLastBatchRows`, `gLastBatchVol`).
- Dev note: `MedParser.bas` defines a custom String-typed `IIf`; new code must use an explicit `If` for numeric/flag values rather than `IIf`.

**Block 3 - data integrity (2026-07-09)**
- **Dated CSV dispense-log archive.** Every print now also appends its Log row to `dispense-log/YYYY-MM-DD.csv` next to the workbook, so the day's dispensing record survives the on-close Log wipe. The folder is PHI and **git-ignored** (`dispense-log/` added to `.gitignore`; `*.csv` was already ignored) and stays on the local machine. Best-effort and RFC-4180 quoted: any write failure is swallowed so it can never interrupt printing. Toggle with `DISPENSE_CSV_ENABLED`. New `ArchiveDispenseRow` + `CsvField`, called from `LogPrint`.
- **Amber flag for out-of-format expirations.** A filled Expiration that fails the `MM/YYYY` check now shows an amber cell on the Medications tab (in addition to the existing WARNINGS text), so it's catchable at a glance during review. Added to `ApplyRowState`.
- **Expiration standardization + multiple bottles.** Expirations are now normalized (`NormalizeExp`): `.`, `-` and spaces become `/`, 2-digit years expand to 4 digits, and month/day are zero-padded, keeping `MM/YYYY` (or `MM/DD/YYYY` when a day is given). Both Expiration and Lot accept **multiple comma-separated values** (e.g. a med split across two bottles); each expiration is validated/standardized independently, lot numbers are left exactly as typed. Applied on prompt entry and on Validate; the popup labels now say "comma-separate multiple bottles." The label footer's shrink-to-fit keeps multi-value Exp/Lot on one line.

**Block 4 - reliability (2026-07-09)**
- **Version stamp.** `SetupWorkbook` writes `SCU Label Tool v2.0 - built YYYY-MM-DD` onto the Patient & Input sheet and shows the version in the Setup-complete dialog, so the loaded build is identifiable (`APP_VERSION = "2.0"`).
- **Structure self-check on open.** New `CheckWorkbookStructure` verifies the five required sheets and the `Print?` column marker exist, and alerts (only on a problem) so a deleted/renamed sheet or shifted column produces a clear message instead of a cryptic runtime error. Wired into `Workbook_Open`.
- **Printer cache.** `SelectBrotherPrinter` now caches the resolved Brother name for the session and reuses it (skipping the slow WMI lookup) on repeat prints; it self-heals by re-detecting if the cached printer can no longer be selected. Detection body moved to `DetectBrotherPrinter`.
- **Debug toggle.** `DEBUG_MODE` + `Dbg()` write a timestamped trace to the VBE Immediate window when enabled; a no-op in production. Traces added at the printer lookup, structure check, and print completion.
- **Gallery button layout fix.** The gallery header is now a two-row band so **Print Checked Labels** and **Reprint Last Batch** stack one above the other (they previously overlapped), with **Refresh Previews** beside Print Checked. Cards start one row lower to make room.
- **Faster batch printing.** The per-label build was re-inserting the logo PNG from disk (and re-merging the header) for **every** label - the main "constructing the labels" delay. `UpdateLabelPreviewForMedRow` gained a `refreshChrome` flag; batch printing (`PrintCheckedLabels`, `ReprintLastBatch`) now sets the logo/header/column-widths **once up-front** and skips it per label, since the logo shape persists across prints. Single-label preview is unchanged.

**Block 5 - UX polish (2026-07-09)**
- **Exp field focused on open.** The Exp/Lot popup now puts the cursor in the Expiration field automatically (`frmExpLot` `UserForm_Activate`).
- **Confirm + fix for gallery remove.** The gallery "Remove this med" now removes *that specific medication* (previously it routed through the checkbox selection, so it could act on the wrong row) and asks a named confirmation first. The Medications-tab "Remove Selected" already confirmed. `RowRemove`.
- **Bigger primary buttons.** Enlarged Start NEW Patient (230x28), the Medications-tab Print Checked Labels (190x30), and the gallery Print Checked Labels / Reprint Last Batch (176x30) for easier targeting; the gallery header band grew to fit and the version stamp moved clear of the buttons.
- **Gallery logo alignment fix.** The gallery header row heights are now set *before* the card loop, so each card's SC emblem is positioned against the correct row tops and lines up with the (shifted-down) cards instead of sitting too high.
- **Land on the Log after printing.** After Print Checked Labels or Reprint Last Batch, the workbook now switches to the dispense Log (scrolled to the newest entry) so the volunteer sees the record. New `ShowLogSheet`.

**Block 6 - targeted constants (2026-07-09)**
- Centralized the medication-name **wrap threshold and wrap font** (`MED_WRAP_MAXLEN = 38`, `MED_WRAP_FONT = 11`) that were duplicated in both the print label and the gallery, so the two can't drift out of sync. Values unchanged. The single-use, physically-calibrated row-height table was intentionally left as-is (extracting it would risk the tuned print geometry for no benefit).

**Block 7 - folder reorg (light-touch, 2026-07-09)**
- `.gitignore` now ignores throwaway `tools\_*` scratch images and `*.generated.txt`, and documents an optional exception for a tracked clean template.
- `Build-Release.vbs` resolves the bootstrap template from `templates\Template_MedicationDispensing.xlsm` first, falling back to the original root filename - so the build works before *or* after the reorg (non-breaking).
- Added `tools/reorg-v2.ps1`: a preview-by-default migration script (run locally; `-Execute` to apply) that deletes the redundant ~10 MB nested repo clone under `_backups\`, removes stale Excel locks, untracks the `tools\` temp files, and moves reference docs into `docs\` and logo-source art into `assets\logo-source\`. Build-critical files (`MedParser.bas`, `Build-Release.vbs`, `scu_emblem.png`, `MedicationDispensing.xlsm`) stay at the root. The reorg itself is run by the user (the sandbox can't operate on the real repo).

**Block 8 - volunteer quick-start card (2026-07-09)**
- Added `docs/SCU_QuickStart_Card.pdf`: a one-page, print-ready reference (4 steps: enter/paste, review + Exp/Lot, check, print) with a "Good to know" strip covering Reprint Last Batch, Start New Patient, close-clears-and-archives behavior, and the printer/roll. For laminating and taping by the clinic workstation.
- **In-app "Start Here" guide sheet.** The workbook now builds a first sheet ("Start Here") mirroring the card - the 4 color-coded steps, a "Good to know" strip, and a "Go to Patient & Input" button - and **opens to it** (activated in `Workbook_Open`; the build ends on it too). New `BuildQuickStartSheet` / `GuideStep` / `GoToInput`; `SH_GUIDE` constant.

**Log: Encounters - numbered + green-banded (2026-07-09)**
- Each print action (a batch print, a reprint, or a single label) is now logged as one **Encounter**, numbered 1, 2, 3, ... in a new **Encounter** column (Log col 16, also added to the dated CSV). Every med row in the same print shares the number, derived from the Log itself (`NextEncounter`) so it survives across prints and macro resets.
- Each encounter's Log rows are **shaded** in one of **3 cycling greens**, so each patient's print reads as one clearly-bounded block. Replaced the old plain even/odd gray banding.

**Add Medication prompts for Exp/Lot (2026-07-09)**
- The manual **+ Add Medication** flow now offers the same **Enter Expiration and Lot** popup the Parse flow uses (`PromptExpLotPair`, with the MM/YYYY format check), right after the name/strength/instructions prompts. Choosing No leaves the highlighted cells to fill later.

**Missing Quantity + Source now highlight yellow (2026-07-09)**
- Empty **Quantity** and empty **Source** cells highlight **yellow** (like Exp/Lot highlight red when missing) until filled; validation still lists both as warnings.

**Table grid + Refills default (2026-07-09)**
- The Medications data area now gets a **light-gray box grid** across the full width (all columns through Print?), so the table reads as outlined boxes even when empty. The grid survives clearing (borders aren't wiped with contents).
- **Refills now defaults to 0** when parsing (or a manual add) finds no refill count, instead of leaving the cell blank.

**Fix: Reset/New Patient could leave medication rows behind (2026-07-09)**
- All three clear paths (Reset Session, Start NEW Patient, and the silent on-open/close reset) previously found the last row to clear by scanning the **Name** column. If leftover data sat in other columns but the Name cell was blank (a partial row, or old data left after the column reorder), the scan found "nothing" and skipped it - so the sheet still showed stale rows. Replaced with a single `ClearMedArea` helper that wipes a **fixed range** (rows 3-503, all table columns) unconditionally, re-applies the Exp/Lot text format and the Source dropdown. `SetupWorkbook` now also calls it, so a freshly built workbook always opens with a clean Medications tab.

**Fix: Exp/Lot missing on print on some printers (2026-07-09)**
- The label print area is A1:H15 and the **EXP/LOT fields sit on row 15** (the very bottom). With `Zoom=100` and no fit-to-page, a printer whose printable height is slightly shorter than the design would push row 15 onto a never-printed page 2 - so Exp/Lot vanished on print while still showing in the on-screen gallery. This is driver-dependent, which is why identical code printed fine on one PC and dropped the bottom row on another. Label page setup now uses **FitToPagesWide=1 / FitToPagesTall=1 (Zoom=False)**, scaling the whole label to the driver's real printable area on any machine so the bottom row can't fall off.

**Medications header spruce-up + resilient banner (2026-07-09)**
- The **blue title banner and header row now span the full table (A:Q)** so nothing hangs off the right after the column reorder. Both are written/formatted **in code** at build time (title text preserved, defaults to "MEDICATIONS"), so they no longer depend on the template's original merge width - resilient to future column changes.
- Header row restyled uniformly (bold, wrapped, centered, divider underneath) and the two **internal columns (Raw text, Printed?) are hidden** to keep the volunteer view clean; Confidence/Warnings widths tidied.

**Source now required + mirrored in the Log; tidy columns (2026-07-09)**
- **Source defaults blank and is required.** New med rows no longer auto-fill "IN HOUSE" - the Source cell starts empty and shows **red** until a value is picked, and **Review & Validate** now flags `SOURCE missing` (alongside Expiration/Lot). Passing rows (which now include a chosen Source) still auto-check for printing.
- **Log mirrors the Medications order.** In the dispense Log (and the dated CSV), **Source moved to sit right after Lot** (Log col 11), with Rx Date, Initials, Dosage Form, and Print # shifting one column right. Tebra note generation and the CSV header were updated to the new order.
- **Tidy table.** Short columns (#, Qty, Exp, Lot, Source, Rx Date, Refills, # of Prints, Print?) are centered; name/strength/form/directions stay left; all vertically centered. The Source dropdown is now applied as the **only** data-validation in the grid (any stray dropdown elsewhere is cleared first).

**Source column moved right of Lot # (2026-07-09)**
- The **Source** column moved from the far right to sit **immediately right of Lot #** (new Medications order: ... Expiration, Lot #, Source, Date of Rx, Refills, Confidence, Warnings, ...). Done by renumbering the `C_*` constants (Source = 9; Date/Refills/Confidence/Warnings/Raw/Printed/# of Prints/Print? each shift right one). Because this shift moves the "# of Prints" and "Print?" columns, `SetupWorkbook` now **rewrites all the reordered headers** (code-authoritative) and **re-installs the double-click Print? handler with the current column numbers** (`InstallMedSheetEvents`, using constants; runs at build time only). Row-state formatting and the row-delete ranges were re-pointed to the new rightmost column. Log/CSV column order is unchanged.

**Workflow: 2 copies + auto-check on validate (2026-07-09)**
- **Every label now prints 2 copies.** Print Checked Labels, Reprint Last Batch, and the single Print This Label all print `LABEL_COPIES` (= 2) of each label. The print/reprint confirmations and completion messages state the copies and total label count. (Change `LABEL_COPIES` to adjust.) Uses the printer driver's copies count; verify on the Brother that 2 die-cut labels feed per medication.
- **Passing meds auto-check for printing.** After **Review & Validate**, every medication that passes (WARNINGS = OK) is automatically checked in the Print? column, so the volunteer doesn't have to check each one. Flagged rows stay unchecked; the summary notes "uncheck any you don't want."

**Tebra Template sheet (2026-07-09)**
- New **"TEBRA TEMPLATE"** sheet at the end of the workbook: **one full pasteable note per patient for the whole session**, built from the dispense **Log**. A teal session header, then for each distinct patient in the Log a block with the patient's **Name + DOB (Name left, DOB right)** and the reconciliation / counseling template, with their meds listed **one line each** grouped into the three sections **by Source** (IN HOUSE / blank -> Dispensed in Clinic; DOH/Other -> DOH & Outside Pharmacy; RxAPS -> RxAPs). Each line is `Name Strength Form - Directions (Qty, Refills, Exp, Lot)` with blank fields omitted (DOH/outside meds show no Lot/Exp), and duplicate lines (reprints) are de-duplicated. A **"Refresh from session log"** button rebuilds it. `FillTebraTemplate` / `EnsureTebraSheet` / `TebraPatientBlock` / `TebraPatientHeader` / `TebraLogSection` / `TebraLogMedLine` / `TebraLine` / `TebraAppend`; `SH_TEBRA`.

**Label tweaks - Name/DOB + directions (2026-07-09)**
- Patient **Name and DOB** are a bit smaller (12 pt); the freed space is given to the **directions** block, which grows and shifts up (rows 10-12 to 15 pt each). The **directions font is larger** (11/10/9 pt by length, was 9/8/7). The A1:H15 print-area height is unchanged (name/DOB rows shrank by the same amount the directions box grew), so it still fits the DK-1202. Mirrored in the gallery.

**Source column + Refills reorder (2026-07-09)**
- **Medication Source dropdown.** New "Source" column on the Medications tab with an in-cell dropdown (**DOH / IN HOUSE / RxAPS / Other**), defaulting to **IN HOUSE** for parsed/added meds. Recorded in the **Log** (new Source column) and the dated CSV archive. It is not printed on the label (Log-only). New `C_SRC` (col 17) + `ApplySourceValidation`; the side buttons moved to column 19; row-delete ranges and Log/CSV writers extended to include it.
- **Refills moved right of Date of Rx.** The Medications columns were reordered so Refills sits just to the right of the Rx Date column (new order: Qty, Expiration, Lot, Date of Rx, Refills, ...). Done by renumbering the `C_*` constants (`C_EXP=7, C_LOT=8, C_DATE=9, C_REF=10`) and having `SetupWorkbook` write those headers authoritatively - the whole codebase addresses med fields via the constants, so the logic follows automatically. The dispense Log column order is unchanged.

**Release packaging (2026-07-09)**
- Added `docs/RELEASE.md` (download/install steps for clinics + how a maintainer cuts a GitHub Release) and `tools/make-release-zip.ps1`, which packages `dist\SCU-Label-Printing-v<version>.zip` (workbook + emblem + quick-start card + `INSTALL.txt`) with a PHI-free confirmation prompt. `dist\` is git-ignored; the ZIP is attached to a GitHub Release rather than committed. README now links the Releases page.

### Print progress popup + EXP/LOT shrink-to-fit (2026-07-09)

- **Progress popup during Print Checked Labels.** After the confirm dialog, a small `frmBusy` popup ("Locating the Brother QL-1100c..." -> "Preparing the label page...") covers the delay while the printer is located and the page is set up, then hides just before the initials prompt. New `BusyShow(pct, msg)` / `BusyHide` helpers in `MedParser.bas` (module var `busyFrm`); the form is built by `Build-Release.vbs`. Falls back to the status bar if `frmBusy` isn't present.
- **EXP/LOT shrink-to-fit.** The footer Expiration/Lot value now steps its font down for long values (14/9/7.5/6.5 pt by length) so long lot numbers stay on one line. Excel's real ShrinkToFit is ignored on merged cells, so this is done by length in `SetMiniValue` (applies to print + gallery).
- **Long medication names wrap to two lines.** When a medication + strength line is longer than 38 characters it now wraps to two lines at a readable 11 pt (instead of shrinking to a cramped single 10 pt line), filling the white space in the med band. Row 7 grows to 28 pt and the spacer rows 13/14 shrink to 1 pt each, so the `A1:H15` print-area height is unchanged. Short names are unchanged. Mirrored in the gallery previews. `UpdateLabelPreviewForMedRow`, `BuildAllLabelsPreview`.
- **`frmExpLot` popup enlarged.** Widened/heightened (292 x 258) with a bigger bottom margin so the OK button and lower fields no longer clip on the bottom-left.
- **Forms are now self-healing (`EnsureForm`).** `Build-Release.vbs` rebuilds `frmExpLot`, `frmMedEdit`, and `frmBusy` **every run** instead of "create-if-missing" — so an old, cropped, or wrongly-captioned form (e.g. a stale "UserForm1" that the old builder kept skipping) is resized and rebuilt automatically. **No manual delete needed anymore.** `EnsureForm` clears the form's controls + code and resizes it but never calls `VBComponents.Remove` (which previously caused an "Unknown runtime error" before `xl.Run`).
- **Bigger medication name on `frmExpLot`.** The medication + strength line is now 16 pt bold (was ~8 pt) with word-wrap, so it's easy to read at a glance while entering Exp/Lot; the form grew to fit (286 tall).
- **Old-form controls reconciled by name (`NewCtl` + `HideExtras`).** Removing controls from an existing form via automation proved unreliable on the clinic PC — the stale small-font `lblMed` and a leftover header label survived, and re-adding `lblMed` silently failed as a duplicate (so the 16 pt size never applied). The builder now **reuses each control by name and sets its properties** (which does work) and **hides any leftover controls**, so an old form is corrected in place instead of duplicated.
- **Form caption + size set at runtime (`UserForm_Initialize`).** Each form now sets its own `Me.Caption`/`Me.Width`/`Me.Height` when it loads, fixing forms that opened titled "UserForm1" and cropped at the bottom (the design-time `frm.Properties` sizing didn't apply reliably, likely under Windows display scaling). Heights bumped for margin: frmExpLot 270, frmMedEdit 350, frmBusy 170.
- **Build-Release progress meter.** `Build-Release.vbs` now shows a filled-block progress bar in Excel's status bar for each phase (Starting Excel -> Opening workbook -> Importing MedParser -> Building forms -> Installing handlers -> Compiling + layout -> Saving -> Done), so the previously "silent" build startup gives live feedback. New `Prog(pct, msg)` helper; the status bar is handed back to Excel when the build finishes.
- **`frmBusy` popup enlarged.** Grown to 288 x 152 (same fix as `frmExpLot`) so the percent label no longer clips at the bottom. Picked up automatically via `EnsureForm` (no manual delete).

### Input forms, exp/lot prompt, session auto-reset (2026-07-02)

**Expiration / Lot prompt**
- After parsing, Exp + Lot are asked **per medication** (paired), not all-Exp-then-all-Lot.
- Expiration is format-checked (`MM/YYYY`, also `MM/YY`, `MM-YYYY`): if it looks wrong the volunteer is prompted to re-enter, but can keep it as typed (never blocked). `PromptExpiration`, `IsBadExpFormat`.

**Two-field UserForms (built by `Build-Release.vbs`, baked into the workbook)**
- `frmExpLot` — one popup per med with **Expiration** and **Lot** fields together (`PromptExpLotPair`).
- `frmMedEdit` — "Edit this med" now opens a prefilled editor: **Medication + Strength bold at the top**, then Dosage form, Quantity, Directions (multi-line), Expiration, Lot. Refills intentionally excluded (edit on the sheet if ever needed). `RowEdit` / `EditMedWithForm`.
- Forms are referenced **late-bound** in `MedParser.bas` (no compile dependency) and both routines **fall back to plain input boxes** if a form isn't present, so editing/prompting always works.
- Forms are **create-if-missing**. To change an existing form's design, delete it once in the VBA editor, then rebuild (automation can't safely `Remove`+recreate a form in the same run — doing so caused an "Unknown runtime error").

**Session auto-reset (PHI hygiene)**
- `ThisWorkbook.Workbook_Open` clears patient + meds + paste box (keeps the Log) so it always opens empty; `Workbook_BeforeClose` does a **full wipe including the Log** (`ClearSessionSilent` + `ClearLogSilent`) and **saves**, so no patient data persists on disk. Handlers installed by `Build-Release.vbs`, which **replaces** the ThisWorkbook module each build (so a stale/legacy `Workbook_Open` can't block the clear-on-open).

**Build-Release**
- Now also builds `frmExpLot` + `frmMedEdit` and installs the auto-reset handlers. On a compile/runtime failure it now **leaves Excel open** so the exact line is inspectable (Alt+F11 -> Debug -> Compile); other failures still clean up.
- Fixed a builder bug where a duplicated `btnCancel_Click` was injected into `frmMedEdit` (regex stopped at the first `End Sub`) -> "Ambiguous name" compile error.

### Header redesign, refills, gallery, emblem fix (2026-07-02)

**Three-zone header (print + gallery)**
- Header split into three zones: clinic name (two lines, left) | SC emblem (centered) | phone-over-address (right).
- Cell map now: name **A2:C2** / **A3:C3**, emblem **D2:E3** (merged, centered), phone **F2:H2**, address **F3:H3**. Gallery cards mirror this (name A:B, emblem C:D, phone/addr E:F).
- Clinic name font changed to **Century Gothic** (`FONT_LABEL_HDR = "Century Gothic"`, Arial fallback). Phone/address in Arial via new `FmtLblContactRight`; name line 2 via new `FmtLblNameSub`.
- New constants: `CLINIC_NAMESUB_FONT_PRINT = 7`, `CLINIC_PHONE_FONT_PRINT = 11`; `CLINIC_ADDR_FONT_PRINT` 7 → **8.5**; header rows `LOGO_HDR_ROW1_PT/ROW2_PT` 20/10 → **18/12**.
- Emblem now **centered** in its slot on both surfaces — `InsertLabelLogo` gained an optional `centerHoriz` parameter.
- Patient name vertically centered on the Rx/DOB block. In the gallery, Rx and DOB now **stack** (Rx over DOB) to match the print.

**Bottom of label**
- Directions (SIG) block enlarged to **3 lines** (A10:H12).
- **EXP/LOT pinned to the bottom row** (A15:D15 / E15:H15).
- **Refills** added after quantity on the form/qty line (print label + gallery).

**Gallery buttons**
- `Print Checked Labels` and `Refresh Previews` are now **code-created at the top-right** of the gallery; each rebuild sweeps stray manually-placed autoshapes.
- Per-card `Print this label` → state-aware **`Check this label` / `Uncheck this label`** (`RowCheck` → `ToggleRowSelect`, which toggles the med's `Print?` selection on the Medications tab).

**Emblem-blank bug (critical fix)**
- The logo had vanished from print AND gallery because **`scu_emblem.png` was a fully transparent (blank) image** — `Build-ScuEmblem.ps1` had zeroed the alpha channel. Regenerated `scu_emblem.png` from the source crop (ink forced to solid black, **alpha preserved**). **`Build-ScuEmblem.ps1` still needs repair — do not run it until fixed.**

### Label layout, logo, print (2026-06-30)

**Layout & typography**
- Header cell map: clinic name **A2:F2** (bold, bottom-aligned), address **A3:F3** (top-aligned), emblem slot **G2:H3** (no fill).
- Clinic name fonts: **14 pt print / 12 pt gallery** via `FmtLblClinicName` (always bold; `Font.Bold` only — no `Font.Weight` / `xlBold`).
- Header row heights: **20 pt + 10 pt** (`LOGO_HDR_ROW1_PT`, `LOGO_HDR_ROW2_PT`).
- Address line: **7 pt** (`CLINIC_ADDR_FONT_PRINT`); `FmtLblHeader` with `ShrinkToFit`, no wrap.
- Patient name typography steps down from **17 pt** and stays **at least 1 pt larger than the medication line** (`PatientNameFontSize`).
- Clinic header lines use **Helvetica** when available, with Arial fallback (`LabelHeaderFont`).

**Print width**
- `LABEL_WIDTH_PT` tuned: **228** stopped bleed onto a second die-cut; **242** uses more of the 100 mm width (re-verify on clinic Brother).

**Logo / emblem**
- Manual crop: `cropped_Black SCU Logo + Transparent Background - Copy.png` → `scu_emblem.png` via `tools/Build-ScuEmblem.ps1` (forces **pure black** pixels for thermal).
- Aspect ratio `LOGO_ASPECT = 1.488`; print height **30 pt**, gallery **28 pt**.
- `InsertLabelLogo`: natural aspect insert (not `maxW × height` box), `LockAspectRatio`, vertical centering in band, `ZOrder msoBringToFront`.
- `RefreshPrintLabelLogo` on preview update and before each print (fixes gallery vs print mismatch).
- `LogoFilePath()` uses **local `scu_emblem.png` only** — removed embedded `LogoB64()` fallback (caused partial/weird logos after rebuild).
- Slot insets: **4 pt print**, **2 pt gallery**; right pad **0 pt**.

**Print fixes**
- Blank label after each print: `PrintOut From:=1, To:=1`; `PrepareLabelSheetForPrint`; only `scuLogo` prints as a shape.
- `SetupWorkbook` calls `PreviewAllLabels` after layout rebuild.

**Build & compile**
- Fixed `SetupWorkbook` runtime 1004 on header `PasteSpecial` (`MatchHeaderFormat`).
- Fixed `Build-Release.vbs` (VBScript syntax, qualified `'MedicationDispensing.xlsm'!SetupWorkbook` macro call).
- Bootstrap warning MsgBox when `MedicationDispensing.xlsm` is missing (patient data not restored).
- Fixed `MedParser.bas` import/compile issues: restore `FONT_LABEL_HDR_FB`, no typed optional defaults in VBA, strip UTF-8 BOM, replace undefined `msoPictureCompressNone` with `0`.
- `Build-Release.vbs` opens `MedicationDispensing.xlsm` in place (bootstrap from `Broken_PrettyPrint_…` only if missing). Added `tools/Run-BuildRelease.ps1` for optional automation.

### Docs & repo path

- Relocated the repo off OneDrive to `C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing` (Documents no longer synced to OneDrive). Updated `HANDOFF.md` and `README.md` to reflect the single-folder layout.

Repository/documentation hardening:

- Clarified that the GitHub repo is code-only and must not contain live workbooks or PHI.
- Added de novo setup instructions for Windows, Excel, Brother QL-1100C, workbook template, VBA import, and sheet event handlers.
- Added branching/release guidance: `main` stable, `dev` active integration, `feature/*` targeted changes.
- Added troubleshooting guidance for run-time error 9, `.bas` import failures, Brother paper-size issues, squished labels, macro blocking, and Git/OneDrive issues.
- Added `.gitattributes` to preserve CRLF line endings for `.bas` and `.vbs` files.
- Added no-PHI sample Tebra-like medication text for parser testing.

## 2026-06-29

Current workbook/source status from project handoff:

- Added checkbox-driven selection and `Print Checked Labels` workflow.
- Added `# of Prints` column and print-count logging.
- Added `Start NEW Patient` workflow that keeps the dispense log.
- Hid the internal `Label Preview` sheet and moved volunteer-facing previews to `Label Previews`.
- Removed lavender printed row state; row color now reflects selected/validated/non-validated state.
- Disabled logo import because `Shapes.AddPicture` could abort setup.
- Removed `PageSetup.PaperSize` handling because Brother DK labels are controlled by printer defaults.
- Changed print scaling to `Zoom = 100` with no fit-to-page behavior.
- Removed runtime VBProject injection in favor of preinstalled worksheet event handlers.
