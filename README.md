# SCU Label Printing tool

An Excel + VBA tool for the **Saturday Clinic for the Uninsured (SCU)** free
pharmacy that turns a block of pasted prescription text into clean, validated
medication labels printed on a **Brother QL-1100c** (DK-1202, 62 x 100 mm,
landscape) - and keeps a running dispense log.

> **Download & install:** grab the latest `SCU-Label-Printing-vX.Y.zip` from the
> **[Releases page](https://github.com/JamesSRN/scu-label-tool/releases)**, unzip, and
> open `MedicationDispensing.xlsm`. Full steps in **[support/docs/RELEASE.md](support/docs/RELEASE.md)**.

## Which file do I click?

Only two files in this folder are meant to be clicked:

- **`MedicationDispensing.xlsm`** — **open this to use the tool** (parse, print labels, log).
- **`APPLY UPDATES (double-click me).cmd`** — **double-click this to apply code updates**
  after `MedParser.bas` changes. It reminds you to close the workbook first, then runs the
  build. (It just launches `Build-Release.vbs`, which still works if you prefer to run it directly.)

Everything else is support material — leave it in place:

- **`support/`** — reference material grouped together: `docs/`, `test-data/`, `assets/`, `_backups/`.
- **`tools/`** — build scripts the build runs (stays in the main folder).
- **`dispense-log/`** — the daily CSV archive the tool writes (stays next to the workbook).
- The `.bas` source, the emblem PNG, and the bootstrap workbook are also needed by the build.

It is built to be **robust for high-turnover volunteers**: the tool parses,
checks, and flags entries so that someone using it for the first time is guided
toward correct labels and caught before mistakes reach a patient.

## The workflow it's built around

1. **Paste** the prescription text for a patient into the *Patient & Input* tab
   and enter the patient name + DOB.
2. **Parse** (one button / `Ctrl+Shift+P`): the text is split into one row per
   medication on the *Medications* tab - name, strength, form, directions (SIG),
   quantity, refills, expiration, and lot. **Refills default to 0** and a
   **Source** dropdown (DOH / IN HOUSE / RxAPS / Other) starts blank. After
   parsing (and from **+ Add Medication**) the tool offers to collect Expiration
   and Lot right away.
3. **Review & fix**: every drug gets a **confidence** rating, and a validation
   pass flags anything missing or suspicious. Cells are color-coded: **red** for
   missing Expiration/Lot, **yellow** for missing Quantity/Source, plus bad-format
   and duplicate warnings. Rows are fully editable; volunteers can add or remove
   medications by hand. Meds that pass validation are **auto-checked** for printing.
4. **Select what to print** with the *Print?* checkboxes (double-click to toggle) -
   one med or many. Selected rows turn green, matching the preview gallery.
5. **Preview** every label on the *Label Previews* tab (auto-refreshes), then
   **Print Checked Labels** - **2 copies of each label**. A confirmation lists
   exactly what's about to print and flags any rows skipped for missing
   expiration/lot. **Reprint Last Batch** recovers from a jam without re-selecting.
6. **Log**: every printed label is recorded (timestamp, patient, DOB, medication,
   strength, SIG, quantity, refills, expiration, lot, **source**, Rx date, initials,
   dosage form, print count, and **Encounter #**). Each print is one **Encounter**,
   numbered and shaded in one of three cycling greens so each patient's print reads
   as one block. A dated CSV copy is also saved locally.
7. **Tebra notes**: the *TEBRA TEMPLATE* sheet builds paste-ready session notes -
   one block per patient, grouped by source, with name/DOB on the right.
8. **Next patient**: *Start NEW Patient* clears the patient and medication list
   but keeps the Log, so the clinic can run patient after patient. *Reset
   Session* wipes everything (including the Log) for a fresh day.

## Key features

- One-paste-to-labels parsing with per-drug confidence scoring.
- Validation that catches missing or malformed fields before printing, with
  color-coded cells (red = Exp/Lot, yellow = Quantity/Source) and auto-check of
  passing meds.
- **Source** dropdown (DOH / IN HOUSE / RxAPS / Other), required and logged.
- Checkbox-driven selection that drives printing, row color, and removal alike.
- Auto-refreshing gallery of all label previews.
- Single and batch printing with detailed confirmation popups - **2 copies per
  label** - plus Reprint Last Batch, and fit-to-page so the bottom row never clips.
- Auto-incrementing, edit-protected print count per medication.
- Complete dispense log with per-print **Encounters** (numbered + green-banded)
  and a dated local CSV archive.
- **TEBRA TEMPLATE** sheet: paste-ready session notes grouped by source per patient.
- Brother QL-1100c auto-select with a load-the-roll confirmation.

## Contents

**In the main folder:**

- `MedicationDispensing.xlsm` - clinic workbook (local, git-ignored; open this to use the tool).
- `APPLY UPDATES (double-click me).cmd` - friendly launcher that runs the build.
- `Build-Release.vbs` - imports `MedParser.bas`, runs `SetupWorkbook`, saves the `.xlsm`.
- `MedParser.bas` - the VBA source of truth (parser, validation, label layout, printing, logging).
- `scu_emblem.png` - SC emblem for the label header (loaded at print time; must stay here).
- `README.md` - this file.

**In `support/docs/`:** `HANDOFF.md` (architecture / routine map), `LABEL_REDESIGN.md`, `CHANGELOG.md`
(release history), `RELEASE.md`, `SETUP_INSTRUCTIONS.md`, `TROUBLESHOOTING.md`, `PRIVACY_AND_PHI.md`.

**In `support/`:** `assets/logo-source/` (full + cropped clinic logo art), `test-data/` (no-PHI samples),
`_backups/` (local snapshots, git-ignored).

**In `tools/`:** `check-encoding.ps1` (pre-build ASCII/CRLF check), `make-release-zip.ps1`,
`Build-ScuEmblem.ps1`, `BUILD_RELEASE_NOTES.md`.

**`dispense-log/`:** dated CSV archive the tool writes each clinic day (git-ignored, PHI).

## Folder on the clinic PC

Everything lives in **one folder** (outside OneDrive):

```
C:\Users\ringo\Documents\GitHub\GIT_VERSION_SCU Label Printing
```

This folder holds the Git repo (pushed to GitHub), the VBA source, docs, and
the live workbook (`MedicationDispensing.xlsm`). Code and docs are version-
controlled; the `.xlsm`, `dispense-log\`, and `support\_backups\` are git-ignored so patient data stays
local. Register this folder as an Excel Trusted Location so macros run.

## Applying code changes

Close the workbook, then double-click **`APPLY UPDATES (double-click me).cmd`**
(it launches `Build-Release.vbs` for you). Click OK on **Setup complete!** when
prompted. Requires Excel **Trust access to the VBA project object model** and this
folder as a **Trusted Location**.

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
