# Privacy and PHI Guidance

## Core rule

Do not commit patient data to GitHub.

This includes:

- Patient names
- DOBs
- Medication lists linked to a patient
- Lot/expiration records tied to a patient dispense
- A dispense log that still has rows
- Clinic screenshots containing patient data
- CSV exports from a clinic day

## What Git is allowed to hold

`MedicationDispensing.xlsm` is tracked **because close wipes it**. `Workbook_BeforeClose` clears the patient, the medication list, and the on-screen Log, then saves. Commit that workbook only after Excel has closed it.

Also allowed:

```text
MedParser.bas
Build-Release.vbs and the click-me launcher
Setup and troubleshooting docs
support/test-data/sample_tebra_pastes_no_phi.txt
Screenshots of the Developer Test random patient
scu_emblem.png
```

Not allowed:

```text
A workbook that still has a patient on screen
dispense-log/*.csv
support/_backups/
Tebra exports
Screenshots with real patients
```

`.gitignore` ignores `*.csv`, `*.xlsx`, `*.xls`, `dispense-log/`, and `support/_backups/`. It does **not** ignore `*.xlsm`.

## Local/offline design

The tool runs in Excel + VBA on the clinic PC. The parser does not call an external API. Local Windows pieces (`VBScript.RegExp`, WMI printer lookup) are part of the design.

## AI and Teams caution

Do not paste real patient medication data into an external AI tool unless SCU leadership has approved that workflow and the vendor is covered by the right agreement.

## Before every commit

- The workbook was closed in Excel first (so the wipe-and-save ran).
- `dispense-log/` and `support/_backups/` are not staged.
- Screenshots are the random test patient, or have no patient data.
- `git diff --cached` has no names, DOBs, or lot numbers from a real visit.
