# Privacy and PHI Guidance

## Core rule

Do not commit PHI to GitHub.

This includes:

- Patient names
- DOBs
- Medication lists linked to a patient
- Lot/expiration records tied to a patient dispense
- Dispense logs
- Clinic screenshots containing patient data
- CSV/XLSX/XLSM exports from clinic workflows

## Repository scope

This repository is for source code and documentation only. It is not the operational patient-data store.

Allowed in Git:

```text
MedParser.bas
Build/release notes
Setup documentation
Troubleshooting notes
No-PHI sample paste text
No-PHI diagrams or screenshots
```

Not allowed in Git:

```text
MedicationDispensing.xlsm with real patients
Excel logs
Tebra exports
Screenshots with PHI
Ad hoc backups from clinic days
```

## Local/offline design

The current tool is intended to run locally in Excel + VBA. The parser should not call external AI APIs or web services. Local Windows components such as `VBScript.RegExp` and WMI printer lookup are acceptable for the current design.

## AI and Teams caution

Do not paste real patient medication data into consumer ChatGPT, Claude, or any external AI parser unless SCU leadership has approved the workflow and the vendor relationship is covered by the appropriate privacy/security agreement.

Teams and Power Automate may be part of the future workflow, but any PHI-routing design should be reviewed before deployment.

## No-PHI sample data

Use fictional patients and no-PHI medication examples for testing. Provider/pharmacy names should also be fictionalized in test files unless there is a specific approved reason to include real operational references.

## Before every commit

Check:

- No workbook files are staged.
- No CSV/log exports are staged.
- No screenshots with patient data are staged.
- `git diff --cached` contains only source/docs/no-PHI samples.
