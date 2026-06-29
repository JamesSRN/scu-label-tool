# SCU Label Printing tool

An Excel + VBA tool for the **Saturday Clinic for the Uninsured (SCU)** free
pharmacy that turns a block of pasted prescription text into clean, validated
medication labels printed on a **Brother QL-1100c** (DK-1202, 62 x 100 mm,
landscape) — and keeps a running dispense log.

It is built to be **robust for high-turnover volunteers**: the tool parses,
checks, and flags entries so that someone using it for the first time is guided
toward correct labels and caught before mistakes reach a patient.

## The workflow it's built around

1. **Paste** the prescription text for a patient into the *Patient & Input* tab
   and enter the patient name + DOB.
2. **Parse** (one button / `Ctrl+Shift+P`): the text is split into one row per
   medication on the *Medications* tab — name, strength, form, directions (SIG),
   quantity, refills, expiration, and lot.
3. **Review & fix**: every drug gets a **confidence** rating, and a validation
   pass flags anything missing or suspicious (no strength/SIG/quantity, bad
   expiration format, possible duplicates). Rows are color-coded and fully
   editable; volunteers can add or remove medications by hand.
4. **Select what to print** with the *Print?* checkboxes — one med or many.
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
- `Build-Release.vbs` - one-click: imports `MedParser.bas`, runs `SetupWorkbook`, saves the `.xlsm`.
- `HANDOFF.md` - full handoff: architecture, feature set, build steps, known issues, routine map.
- `Black SCU Logo + Transparent Background.png` - clinic logo.

## Applying code changes

Close the workbook, then run `Build-Release.vbs` (requires Excel "Trust access to
the VBA project object model"). Or in the VBA editor: remove the `MedParser`
module, Import `MedParser.bas`, run `SetupWorkbook`, save. The two worksheet
event handlers are pasted once into the sheet modules — see `HANDOFF.md` section 6.

---

*Note on patient data:* this repo holds **code and docs only**. The working
`.xlsm` contains patient information and is kept local (excluded via
`.gitignore`); it is never committed.
