# Changelog

## Unreleased

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
