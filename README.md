# SCU Label Printing tool

An Excel + VBA tool for the **Saturday Clinic for the Uninsured (SCU)** free
pharmacy that turns a block of pasted prescription text into clean, validated
medication labels printed on a **Brother QL-1100c** (DK-1202, 62 x 100 mm,
landscape) - and keeps a running dispense log.

> **Download & install:** grab the latest `SCU-Label-Printing-vX.Y.zip` from the
> **[Releases page](https://github.com/JamesSRN/scu-label-tool/releases)**, unzip, and
> open `MedicationDispensing.xlsm`. Full steps in **[docs/RELEASE.md](docs/RELEASE.md)**.

It is built to be **robust for high-turnover volunteers**: the tool parses,
checks, and flags entries so that someone using it for the first time is guided
toward correct labels and caught before mistakes reach a patient.

## The workflow it's built around

1. **Paste** the prescription text for a patient into the *Patient & Input* tab
   and enter the patient name + DOB.
2. **Parse** (one button / `Ctrl+Shift+P`): the text is split into one row per
   medication on the *Medications* tab - name, strength, form, directions (SIG),
   quantity, refills, expiration, and lot.
3. **Review & fix**: every drug gets a **confidence** rating, and a validation
   pass flags anything missing or suspicious (no strength/SIG/quantity, bad
   expiration format, possible duplicates). Rows are color-coded and fully
   editable; volunteers can add or remove medications by hand.
4. **Select what to print** with the *Print?* checkboxes - one med or many.
   Selected rows turn green, matching the preview gallery.
5. **Preview** every label on the *Label Previews* tab (auto-refreshes), then
   **Print Checked Labels**. A confirmation lists exactly what's about to print
   and flags any rows skipped for missing expiration/lot.
6. **Log**: every printed label is recorded (timestamp, patient, DOB,
   medication, strength, SIG, quantity, refills, expiration, lot, dosage form,
   print count, and volunteer initials).
7. **Next patient**: *Start NEW Patient* clears the patient and medication list
   but keeps the Log, so the clinic can run patient after patient. *Reset
   Session* wipes everything (including the Log) for a fresh day.

## Key features

- One-paste-to-labels parsing with per-drug confidence scoring.
- Validation that catches missing or malformed fields before printing.
- Checkbox-driven selection that drives printing, row color, and removal alike.
- Auto-refreshing gallery of all label previews.
- Single and batch printing with detailed confirmation popups.
- Auto-incrementing, edit-protected print count per medication.
- Complete dispense log for single and batch prints.
- Brother QL-1100c auto-select with a load-the-roll confirmation.

## Contents

- `MedParser.bas` - the VBA source of truth (parser, validation, label layout, printing, logging).
- `MedicationDispensing.xlsm` - clinic workbook (local, git-ignored; rebuilt by the script below).
- `Build-Release.vbs` - one-click: imports `MedParser.bas`, runs `SetupWorkbook`, saves the `.xlsm`.
- `scu_emblem.png` - SC emblem for the label header (from manual crop of the clinic logo).
- `cropped_Black SCU Logo + Transparent Background - Copy.png` - source crop for the emblem.
- `tools/Build-ScuEmblem.ps1` - copies the manual crop to `scu_emblem.png` (pure black for thermal).
- `HANDOFF.md` - full handoff: architecture, feature set, build steps, known issues, routine map.
- `LABEL_REDESIGN.md` - printed label layout, typography, logo workflow, tuning notes.
- `CHANGELOG.md` - release history (see Unreleased for 2026-06-30 label/build fixes).
- `tools/BUILD_RELEASE_NOTES.md` - `Build-Release.vbs` prerequisites and steps.
- `Black SCU Logo + Transparent Background.png` - full clinic logo source art.

## Folder on the clinic PC

Everything lives in **one folder** (outside OneDrive):

```
C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing
```

This folder holds the Git repo (pushed to GitHub), the VBA source, docs, and
the live workbook (`MedicationDispensing.xlsm`). Code and docs are version-
controlled; the `.xlsm` and `_backups\` are git-ignored so patient data stays
local. Register this folder as an Excel Trusted Location so macros run.

## Applying code changes

Close the workbook, then double-click **`Build-Release.vbs`** (requires Excel
**Trust access to the VBA project object model** and this folder as a **Trusted
Location**). Click OK on **Setup complete!** when prompted.

Or in the VBA editor: remove the `MedParser` module, Import `MedParser.bas`, run
`SetupWorkbook`, save.

After updating the emblem crop, run `tools/Build-ScuEmblem.ps1` first, then
`Build-Release.vbs`.

The two worksheet event handlers are pasted once into the sheet modules - see
`HANDOFF.md` section 5.

## Known open items

- **Trusted Location** — register the folder in Excel Trust Center so macros are not blocked.
- **Print width** — `LABEL_WIDTH_PT = 242` uses more of the 100 mm die-cut; re-test on the clinic Brother that labels still fit one die-cut (228 pt was the prior no-bleed value).
- **Bold clinic title on thermal** — may still look subtle; `ShrinkToFit` can limit apparent size.
- **Physical regression test** — spot-check emblem and header after the 2026-06-30 logo/header fixes (screen + print path verified; Brother thermal spot-check recommended).

## Recently fixed (2026-06-30)

Logo/header/print tuning session — full detail in `HANDOFF.md` and `CHANGELOG.md`:

- Emblem sliver, squished aspect, gray thermal output, gallery vs print mismatch
- Blank label after each print (`PrintOut From:=1, To:=1`)
- Clinic name clipped / too small / bold (`FmtLblClinicName`, A2:F2 merge)
- Embedded logo fallback removed (weird partial logos after rebuild)
- Bootstrap warning when `MedicationDispensing.xlsm` is missing during build

---

*Note on patient data:* this repo holds **code and docs only**. The working
`.xlsm` contains patient information and is kept local (excluded via
`.gitignore`); it is never committed.
