# Branching and Releases

## What this repo actually uses

The GitHub branch is **`master`**. There is no `main` or `dev` branch in the current history. Tags in use: `v2.1`, `v2.2`, `v2.3`, `v2.4`.

Source of truth:

```text
MedParser.bas          repo root (not src/MedParser.bas)
Build-Release.vbs
OPEN LABEL TOOL (double-click me).cmd
README.md
HANDOFF.md
support/docs/
```

`MedicationDispensing.xlsm` is tracked. Commit it only after a close, which wipes patient data and the Log.

Do not commit:

```text
dispense-log/
*.csv
support/_backups/
real patient screenshots
dist/ release zips (attach those to a GitHub Release)
```

## How to change the app

1. Edit `MedParser.bas`. Keep it ASCII + CRLF (`tools/check-encoding.ps1`).
2. Leave `Build-Release.vbs` unchanged unless the build itself must change. A new hash can be flagged by Windows Defender on fresh downloads.
3. Close the workbook, then run `OPEN LABEL TOOL (double-click me).cmd` so `SetupWorkbook` compiles.
4. Smoke-test with `support/test-data/sample_tebra_pastes_no_phi.txt` or Developer Test. No real patients.
5. Update `CHANGELOG.md` and `README.md` if volunteer-visible behavior changed.
6. Commit on `master` (or a short-lived branch merged back to `master`) and tag a release when clinic PCs should pick it up.

The published release ZIP is the **full folder**, flat at the ZIP root. `tools/make-release-zip.ps1` still builds an older slim zip and is not what v2.4 shipped. See [RELEASE.md](RELEASE.md).
