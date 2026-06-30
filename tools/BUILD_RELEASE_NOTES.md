# Build Release Notes

`Build-Release.vbs` (repo root) re-imports `MedParser.bas`, runs `SetupWorkbook`, and saves **`MedicationDispensing.xlsm`**.

## Before running

1. Close `MedicationDispensing.xlsm` in Excel (file must not be locked).
2. One-time in Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings** → check **Trust access to the VBA project object model**.
3. Register the repo folder as an **Excel Trusted Location** so macros are not blocked on open.

## What the script does

1. Opens `MedicationDispensing.xlsm` (or copies `Broken_PrettyPrint_MedicationDispensing.xlsm` if the target is missing).
2. If bootstrapping: shows a **warning MsgBox** that patient data from a previous `.xlsm` is **not** restored — check `_backups\` or Windows File History.
3. Removes the existing `MedParser` module and imports `MedParser.bas` from the repo root.
4. Runs `'MedicationDispensing.xlsm'!SetupWorkbook` (compile check + rebuild buttons/label layout + `PreviewAllLabels`).
5. Saves the workbook.

Click **OK** on the **Setup complete!** dialog when it appears, then on the final **Release build complete** message.

## Optional automation

`tools/Run-BuildRelease.ps1` does the same steps via Excel COM (includes a timer to dismiss the setup MsgBox). Use only on a developer PC, not during clinic hours.

## Emblem asset

After updating the manual crop file, run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/Build-ScuEmblem.ps1
```

That copies `cropped_Black SCU Logo + Transparent Background - Copy.png` → `scu_emblem.png` and forces every visible pixel to **solid black** (`#000000`) for crisp thermal output. Then run `Build-Release.vbs` so the workbook picks up the new image.

**Important:** `LogoFilePath()` uses **only** the local `scu_emblem.png` beside the workbook. The old embedded `LogoB64()` fallback was removed — if the PNG is missing, the logo will not appear until you run the emblem script.

## `MedParser.bas` requirements

- **Pure ASCII**, **Windows CRLF** line endings (enforced by `.gitattributes`).
- **No UTF-8 BOM** on the first line (`Attribute VB_Name` must be byte 1).
- Do not use VBA-only syntax in `Build-Release.vbs` (e.g. `Dim x As String` is invalid in VBScript).
- Do not use `Font.Weight` or `xlBold` for bold text — use `Font.Bold = True` only.

## Separation of concerns

```text
GitHub repo        source, docs, scu_emblem.png, no-PHI samples
Build-Release.vbs  developer convenience only
MedicationDispensing.xlsm   local clinic workbook (git-ignored, may contain PHI)
```

Volunteers should open the built `.xlsm` only; they do not run the build script.

## Session fixes reflected in current build (2026-06-30)

| Area | Change |
|------|--------|
| Header layout | A2:F2 / A3:F3 text; G2:H3 emblem slot |
| Logo | 30 pt print / 28 pt gallery; natural aspect; pure black PNG |
| Print | Single page `PrintOut From:=1, To:=1`; logo refresh before print |
| Width | `LABEL_WIDTH_PT = 242` (re-test on Brother) |
| Bootstrap | Warning when `.xlsm` was missing |

See `HANDOFF.md` and `CHANGELOG.md` for the full problem/solution list.
