# SCU Label Printing — Download & Releases

Repository: <https://github.com/JamesSRN/scu-label-tool>
Latest releases: <https://github.com/JamesSRN/scu-label-tool/releases>

---

## Download & install (clinics / volunteers)

1. Open the **[Releases page](https://github.com/JamesSRN/scu-label-tool/releases)** and, under the newest release, download **`SCU-Label-Printing-vX.Y.zip`** (under *Assets*).
2. **Unzip** it to a folder on the clinic PC. Keep **`MedicationDispensing.xlsm`** and **`scu_emblem.png`** together in that folder.
3. Open **`MedicationDispensing.xlsm`** in Microsoft Excel (Windows).
4. One-time Excel setup:
   - **Enable macros** when prompted (the tool is macro-driven).
   - Recommended: add the folder as a **Trusted Location** (*File → Options → Trust Center → Trust Center Settings → Trusted Locations*) so macros aren't blocked each time.
5. Make sure the **Brother QL-1100c** driver is installed and the **DK-1202 (62 × 100 mm)** roll is loaded.
6. The workbook opens to the **Start Here** guide — follow the 4 steps. A printable copy, **`SCU_QuickStart_Card.pdf`**, is included in the ZIP.

**Privacy:** patient info and the on-screen Log clear when you close the file; a dated CSV copy of each day's dispensing is saved locally in a `dispense-log` folder next to the workbook and never leaves the PC.

---

## Cut a new release (maintainer)

1. **Build & verify:** run `Build-Release.vbs`, then do a real test print.
2. **Get a PHI-free workbook** to ship — a freshly built copy that has never had a patient entered (empty Log, no name/DOB). *Never ship a workbook that has held patient data.*
3. **Build the download ZIP:**
   ```
   powershell -ExecutionPolicy Bypass -File tools\make-release-zip.ps1 -Version 2.0
   ```
   This creates `dist\SCU-Label-Printing-v2.0.zip` (workbook + emblem + quick-start card + `INSTALL.txt`). It asks you to confirm the workbook is PHI-free first. To package a specific clean file: add `-Workbook "C:\path\clean.xlsm"`.
4. **Publish on GitHub:** *Releases → Draft a new release*.
   - **Choose a tag:** `v2.0` (select "Create new tag on publish").
   - **Title:** `v2.0`.
   - **Description:** paste the notes below (or from `CHANGELOG.md`).
   - **Attach** `dist\SCU-Label-Printing-v2.0.zip` (drag it into the *Attach binaries* area).
   - Click **Publish release**.
5. Done — the Download section above now points people to it.

> The built `.xlsm` is git-ignored (PHI policy) and the ZIP lives in the git-ignored `dist\` folder. The workbook reaches users **only** through the release asset you attach — it is never committed to the repo.

---

## Release notes — v2.0

Offline Excel + VBA tool for printing medication labels on a Brother QL-1100c (DK-1202 62 × 100 mm) for the Saturday Clinic for the Uninsured.

**New in v2.0**
- **Start Here guide** — the workbook opens to an in-app quick-start, plus a printable `SCU_QuickStart_Card.pdf`.
- **Reprint Last Batch** — recover from a jam/misfeed without re-selecting.
- **Named skips** — the print summary lists any label skipped for a missing Exp/Lot.
- **Multiple bottles** — Expiration and Lot accept comma-separated values; dates are standardized (`.`/`-`/spaces → `/`, 2-digit years expanded); out-of-format expirations are flagged amber.
- **Dispense-log archive** — each print is also saved to a dated local CSV so the day's record survives the on-close wipe (PHI, stays local).
- **Faster batch printing** — the logo is placed once per batch instead of per label.
- **Reliability** — version stamp, on-open structure check, session printer cache, and a debug toggle.
- **Polish** — Exp field focused on open, bigger primary buttons, per-med remove confirmation, and it lands on the Log after printing.
- **Safer builds** — a source encoding/structure pre-check runs before every build.

See `CHANGELOG.md` for the full list.
