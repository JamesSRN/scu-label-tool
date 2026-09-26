# SCU Dispensary Label Tool — Download & Releases

Repository: <https://github.com/JamesSRN/scu-label-tool>
Latest releases: <https://github.com/JamesSRN/scu-label-tool/releases>

---

## Download & install (clinics / volunteers)

This is the **Dispensary** tool (medication labels on Hermione). Lab / specimen labels are a different repo: [SCU Lab Label Tool](https://github.com/JamesSRN/lab-label-printer).

1. Open the **[latest release](https://github.com/JamesSRN/scu-label-tool/releases/latest)** and, under **Assets**, download **`SCU-Label-Printing-vX.Y.zip`**.
2. **Move the ZIP to the Desktop** (out of Downloads / OneDrive). **Right-click → Extract All…** — do not open files from inside the ZIP window. Keep **`MedicationDispensing.xlsm`**, **`scu_emblem.png`**, and **`OPEN LABEL TOOL (double-click me).cmd`** together in that extracted folder.
3. One-time in Excel (**File → Options → Trust Center → Trust Center Settings…**):
   - **Macro Settings** — check **Trust access to the VBA project object model** (the click-me file rebuilds the workbook on every open).
   - **Trusted Locations → Add new location…** — add the **Desktop** itself, and check **Subfolders of this location are also trusted**, so every version extracted to the Desktop is trusted without re-adding it each release.
4. **Every time:** close the workbook, then double-click **`OPEN LABEL TOOL (double-click me).cmd`** (the click-me file). If that file doesn't run, double-click **`Build-Release.vbs`** instead. Don't open the `.xlsm` by hand unless both of those fail.
5. Make sure the **Brother QL-1100c (Hermione)** is connected by **USB cable** (not Bluetooth), the driver is installed, and the **DK-1202 (62 × 100 mm)** roll is loaded.
6. The workbook opens to the **Start Here** guide — follow the numbered tabs. A printable copy, **`SCU_QuickStart_Card.pdf`**, is included in the ZIP.

**Privacy:** patient info and the on-screen Log clear when you close the file; a dated CSV copy of each day's dispensing is saved locally in a `dispense-log` folder next to the workbook and never leaves the PC.

---

## Cut a new release (maintainer)

1. **Build & verify:** run `Build-Release.vbs`, then do a real test print.
2. **Get a PHI-free workbook** to ship — a freshly built copy that has never had a patient entered (empty Log, no name/DOB). *Never ship a workbook that has held patient data.*
3. **Build the download ZIP:**
   ```
   powershell -ExecutionPolicy Bypass -File tools\make-release-zip.ps1 -Version 2.4
   ```
   That script still builds the **older slim zip** (`MedicationDispensing.xlsm`, `scu_emblem.png`, `SCU_QuickStart_Card.pdf`, `INSTALL.txt`). The zip attached to the **v2.4** GitHub release is the **full working folder** (source, click-me launcher, docs, tools, emblem, and the wiped workbook), flat at the ZIP root. Ship that full folder unless you have deliberately switched back to the slim script.
4. **Publish on GitHub:** *Releases → Draft a new release*.
   - **Choose a tag:** `v2.4` (select "Create new tag on publish").
   - **Title:** `v2.4`.
   - **Description:** paste the current version section from `CHANGELOG.md`. The published v2.4 page still tells people to download `SCU-Label-Printing-v2.3.zip`. The asset file is `SCU-Label-Printing-v2.4.zip`. The next release text should name the zip that is actually attached.
   - **Attach** the zip.
   - Click **Publish release**.
5. Done — the Download section above now points people to it.

> `dist\` is git-ignored. `MedicationDispensing.xlsm` is tracked only after a close, which wipes patient data and the Log. Never attach a workbook that still has a patient in it.

---

## Release notes

Copy the top section of [CHANGELOG.md](CHANGELOG.md). Current version is **v2.4** (2026-08-27): Log row Print / Edit / Remove, and cancelling the initials prompt prints and logs nothing. Review does not auto-check. Do not paste the old v2.0 blurb (it still described auto-check, Reprint Last Batch, and a sheet named TEBRA TEMPLATE).
