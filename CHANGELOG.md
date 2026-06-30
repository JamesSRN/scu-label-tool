# Changelog

## Unreleased

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
