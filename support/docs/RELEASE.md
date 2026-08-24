# SCU Dispensary Label Tool — Download & Releases

Repository: <https://github.com/JamesSRN/scu-label-tool>
Latest releases: <https://github.com/JamesSRN/scu-label-tool/releases>

---

## Download & install (clinics / volunteers)

This is the **Dispensary** tool (medication labels). Lab / specimen labels are a different repo: [lab-label-printer](https://github.com/JamesSRN/lab-label-printer).

1. Open the **[latest release](https://github.com/JamesSRN/scu-label-tool/releases/latest)** and, under **Assets**, download **`SCU-Label-Printing-vX.Y.zip`**.
2. **Move the ZIP to the Desktop** (out of Downloads / OneDrive). **Right-click → Extract All…** — do not open files from inside the ZIP window. Keep **`MedicationDispensing.xlsm`**, **`scu_emblem.png`**, and **`OPEN LABEL TOOL (double-click me).cmd`** together in that extracted folder.
3. One-time in Excel (**File → Options → Trust Center → Trust Center Settings…**):
   - **Macro Settings** — check **Trust access to the VBA project object model** (the click-me file rebuilds the workbook on every open).
   - **Trusted Locations → Add new location…** — the extracted Desktop folder. Check **Subfolders of this location are also trusted**.
4. **Every time:** close the workbook, then double-click **`OPEN LABEL TOOL (double-click me).cmd`** (the click-me file). If that file doesn't run, double-click **`Build-Release.vbs`** instead. Don't open the `.xlsm` by hand unless both of those fail.
5. Make sure the **Brother QL-1100c (Hermione)** driver is installed and the **DK-1202 (62 × 100 mm)** roll is loaded.
6. The workbook opens to the **Start Here** guide — follow the numbered tabs. A printable copy, **`SCU_QuickStart_Card.pdf`**, is included in the ZIP.

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

Offline Excel + VBA tool for printing medication labels on a Brother QL-1100c (Hermione) — DK-1202 62 × 100 mm — for the Saturday Clinic for the Uninsured.

**New in v2.0**
- **2 copies per label** — Print Checked Labels (and Reprint / single Print) print two of each label; the Log still records one row per med.
- **Source column** — a required DOH / IN HOUSE / RxAPS / Other dropdown, right of Lot #, mirrored into the Log and CSV.
- **Auto-check on validate** — meds that pass Review & Validate are automatically checked for printing.
- **Encounters** — each print is logged as one numbered Encounter, and its Log rows are shaded in one of three cycling greens.
- **Add Medication prompts for Exp/Lot** — the manual add offers the same Expiration/Lot popup as parsing.
- **TEBRA TEMPLATE sheet** — paste-ready session notes, one block per patient, grouped by source, name/DOB on the right.
- **Clearer Medications tab** — full-width blue banner, boxed grid, centered columns, internal columns hidden, Refills default 0, and red (Exp/Lot) / yellow (Quantity/Source) missing-field highlights.
- **Print reliability** — the label now fits to one page so the bottom Exp/Lot row can't clip on printers with a shorter printable area.
- **Bullet-proof reset** — Reset / New Patient / on-open reset always clear the whole Medications area.
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
