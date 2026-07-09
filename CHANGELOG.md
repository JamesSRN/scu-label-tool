# Changelog

## Unreleased

### V2 build

Work toward the V2 build, applied in blocks (see `docs/IMPROVEMENT_REPORT_2026-07-09.md`).

**Block 1 - source pre-check (2026-07-09)**
- Added `tools/check-encoding.ps1`: validates `MedParser.bas` and `Build-Release.vbs` are pure ASCII, have no BOM, use Windows CRLF (no lone LF / doubled CR), and have balanced block keywords. Encoding problems exit 1 (build-breaking); balance is advisory (exit 3) since the VBA compile is the authoritative structural check.
- `Build-Release.vbs` now runs that check first and **aborts the build only on an encoding failure** (fail-safe: if PowerShell or the script is unavailable, or only a balance advisory is returned, the build proceeds).

**Block 2 - print workflow UX (2026-07-09)**
- **Skipped labels are now named.** The Print Complete dialog lists each medication skipped for a missing Expiration/Lot (by name), instead of only a count, and uses a warning icon when anything was skipped (`PrintCheckedLabels`).
- **Reprint Last Batch.** New `ReprintLastBatch` action + gallery button reprints the last successfully-printed batch (same rows + initials, logged as a reprint) for paper-jam/misfeed recovery. The remembered batch is **cleared on any session reset or Start New Patient**, so it can never reprint a previous patient's labels. Session-only (module vars `gLastBatchRows`, `gLastBatchVol`).
- Dev note: `MedParser.bas` defines a custom String-typed `IIf`; new code must use an explicit `If` for numeric/flag values rather than `IIf`.

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
