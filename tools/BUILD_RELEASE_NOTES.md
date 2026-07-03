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
5. **Saves AND closes** the workbook, then quits its Excel instance.

Click **OK** on the **Setup complete!** dialog when it appears, then on the final **Release build complete** message. On success the workbook is now closed — reopen `MedicationDispensing.xlsm` to use it.

## Robustness — never leaves a stray Excel process

The script cleans up on **every** exit (added 2026-07-02):

- If the workbook opens **read-only** (already open in another window, or a stray Excel is holding it, or the file is marked read-only), it shows a plain-English message and quits **before changing anything**.
- Every failure path (open failed, read-only, VBA project not accessible, `SetupWorkbook` compile error, save failed) shows a clear message **and quits Excel** via a shared cleanup — so no background `EXCEL.EXE` is left locking the file.
- On success it saves, closes the workbook, and quits Excel.

If an **older** run already left stray Excel processes, end them once via **Task Manager** (any `Microsoft Excel` / `EXCEL.EXE`) so the `~$MedicationDispensing.xlsm` lock clears; after that the script keeps itself clean.

## Optional automation

`tools/Run-BuildRelease.ps1` does the same steps via Excel COM (includes a timer to dismiss the setup MsgBox). Use only on a developer PC, not during clinic hours.

## Emblem asset

**WARNING (2026-07-02): `tools/Build-ScuEmblem.ps1` is currently BROKEN — do NOT run it.**
It regenerated `scu_emblem.png` as a fully **transparent (blank)** PNG (it zeroed the
alpha channel), which made the logo invisible on every label even though the file
existed and loaded. `scu_emblem.png` has been rebuilt by hand (ink forced to solid
black, **alpha preserved**) and is correct — leave it as-is until the script is fixed.

The script's *intended* job: copy `cropped_Black SCU Logo + Transparent Background - Copy.png`
→ `scu_emblem.png`, forcing ink to solid black `#000000` for thermal **while keeping the
alpha channel** (or emitting black-on-white fully opaque). Repair that (do not zero the
alpha), then it is safe to run again, followed by `Build-Release.vbs`.

**Note:** `LogoFilePath()` uses **only** the local `scu_emblem.png` beside the workbook
(the embedded `LogoB64()` fallback was removed) — if the PNG is missing or blank, the
logo will not appear.

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

## Session updates (2026-07-02)

| Area | Change |
|------|--------|
| Header | Three-zone Century Gothic: name **A2:C2 / A3:C3**, emblem centered **D2:E3**, phone **F2:H2** / address **F3:H3** |
| Emblem | Centered (`centerHoriz`); `scu_emblem.png` rebuilt after the blank-PNG bug (see Emblem asset warning) |
| Bottom | Directions **3 lines**; **EXP/LOT on bottom row 15**; **Refills** on the qty line |
| Gallery | Cards mirror the header; top-right `Print Checked Labels` / `Refresh Previews`; per-card `Check` / `Uncheck`; full shape-clear each rebuild; Rx over DOB |
| Build-Release | Robust cleanup — never leaves a stray Excel; read-only guard; **saves and closes** on success |

See `HANDOFF.md` and `CHANGELOG.md` for the full problem/solution list.
