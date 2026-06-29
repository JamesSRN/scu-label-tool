# SCU Label Printing tool

Excel + VBA medication dispensing / label-printing tool for the **Saturday Clinic
for the Uninsured (SCU)** free pharmacy. Parses pasted prescription text into a
medication table, validates it, and prints labels on a **Brother QL-1100c**
(DK-1202, 62 x 100 mm, landscape).

## IMPORTANT - HIPAA / PHI
The Excel workbook (`MedicationDispensing.xlsm`) contains patient PHI (names, DOB,
medications) and **must never be committed**. `.gitignore` excludes all `*.xlsm` /
`*.xlsx` / `*.csv` files. This repo holds **code and docs only**. Do not add the
workbook or any patient data to version control.

## Contents
- `MedParser.bas` - the VBA source of truth (parser, validation, label layout, printing, logging).
- `Build-Release.vbs` - one-click: imports `MedParser.bas`, runs `SetupWorkbook`, saves the `.xlsm`.
- `HANDOFF.md` - full handoff: architecture, feature set, build steps, known issues, routine map.
- `Black SCU Logo + Transparent Background.png` - clinic logo used on the label.

## Applying code changes
Close the workbook, then run `Build-Release.vbs` (requires Excel "Trust access to the
VBA project object model"). Or in the VBA editor: remove the `MedParser` module,
Import `MedParser.bas`, run `SetupWorkbook`, save. See `HANDOFF.md` section 6.
