# Build Release Notes

`Build-Release.vbs` (repo root) re-imports `MedParser.bas`, runs `SetupWorkbook`, and saves **`MedicationDispensing.xlsm`**.

## Before running

1. Close `MedicationDispensing.xlsm` in Excel (file must not be locked).
2. One-time in Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → check **Trust access to the VBA project object model**.
3. Register the repo folder as an **Excel Trusted Location** so macros are not blocked on open.

## What the script does

1. Opens `MedicationDispensing.xlsm` (or copies `Broken_PrettyPrint_MedicationDispensing.xlsm` if the target is missing).
2. Removes the existing `MedParser` module and imports `MedParser.bas` from the repo root.
3. Runs `'MedicationDispensing.xlsm'!SetupWorkbook` (compile check + rebuild buttons/label layout).
4. Saves the workbook.

Click **OK** on the **Setup complete!** dialog when it appears, then on the final **Release build complete** message.

## Optional automation

`tools/Run-BuildRelease.ps1` does the same steps via Excel COM (includes a timer to dismiss the setup MsgBox). Use only on a developer PC, not during clinic hours.

## Emblem asset

After updating the manual crop file, run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/Build-ScuEmblem.ps1
```

That copies `cropped_Black SCU Logo + Transparent Background - Copy.png` → `scu_emblem.png`. Then run `Build-Release.vbs` so the workbook picks up the new image.

## `MedParser.bas` requirements

- **Pure ASCII**, **Windows CRLF** line endings (enforced by `.gitattributes`).
- **No UTF-8 BOM** on the first line (`Attribute VB_Name` must be byte 1).
- Do not use VBA-only syntax in `Build-Release.vbs` (e.g. `Dim x As String` is invalid in VBScript).

## Separation of concerns

```text
GitHub repo        source, docs, scu_emblem.png, no-PHI samples
Build-Release.vbs  developer convenience only
MedicationDispensing.xlsm   local clinic workbook (git-ignored, may contain PHI)
```

Volunteers should open the built `.xlsm` only; they do not run the build script.
