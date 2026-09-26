# SCU Dispensary Label Tool — current state

_Last updated: 2026-09-26. App version **2.4**. GitHub `JamesSRN/scu-label-tool`, branch `master`. The long routine map is [support/docs/HANDOFF.md](support/docs/HANDOFF.md). Volunteer steps are [README.md](README.md)._

This is the medication-label tool (Brother QL-1100c, **Hermione**, USB, DK-1202). Lab labels are a different repo: [lab-label-printer](https://github.com/JamesSRN/lab-label-printer) (Brother QL-820NWB, **Harry**).

---

## How a change has to land

- **Edit `MedParser.bas`.** It is the source of truth (about 7,600 lines). The workbook is rebuilt from it every launch.
- **Keep `MedParser.bas` pure ASCII + Windows CRLF.** `tools/check-encoding.ps1` checks that file only, and `Build-Release.vbs` aborts the build if it fails. Typographic dashes and quotes break the VBA import.
- **Leave `Build-Release.vbs` alone unless you must change the build.** It opens Excel and injects VBA. Windows Defender trusts the current file by reputation. Editing even one character makes a new hash that fresh downloads can flag as a virus.
- **Open the tool with `OPEN LABEL TOOL (double-click me).cmd`**, which runs `Build-Release.vbs`. That imports `MedParser.bas`, rebuilds the four UserForms and the `ThisWorkbook` open/close handlers, runs `SetupWorkbook` (full compile), saves, and **leaves Excel open** on Start Here. On failure it quits Excel so no stray process locks the file.
- **Close the workbook before a rebuild.** If it is already open, answer **No** to Excel's replace prompt, or the open copy overwrites the fresh build.

## What the current code does

Five numbered tabs: **1. Patient & Input**, **2. Medications**, **3. Print Labels**, **4. Log**, **5. Tebra Notes**. Hidden sheets: **Label Preview** (print surface), **EncounterData** (snapshots). **Developer Test** and **Setup & Help** sit after the workflow.

- Parse clears the previous medication list first (keeps name, DOB, and the Log) and asks before wiping a list that already has rows.
- A check is a cell, not a control. Column 2 (`C_SEL`) holds a checkmark. `IsRowSelected` tests that the cell is non-empty.
- **Review does not auto-check.** It validates rows to blue. The volunteer checks what prints (green). Double-click the Check Med header to check or uncheck all.
- Missing Quantity, Expiration, Lot, and Source cells are **yellow**. A filled expiration in the wrong format is **amber**. There is no red missing-field highlight.
- **Print Checked Labels** prints `LABEL_COPIES` (2) of each checked med, logs each row, and lands on the Log. Cancelling the initials prompt prints nothing and logs nothing.
- Gallery cards have Check, Edit, Remove, and **Print extra (no log)** (1 copy, not logged).
- Each Log row has **Print** (1 reprint, not logged again), **Edit** (writes the row and that day's CSV), **Add med** (inserts a new medication for that same patient on the next row and appends it to the CSV), and **Remove**. Those buttons stay on the Log. Typing in a Log cell, or adding a row by hand, writes that day's CSV as well. Closing the workbook does not erase the CSV. Use **Remove** to drop a row from the CSV; deleting the Excel row by hand does not.
- Tebra notes are built from the Log when that tab is activated. Edit Encounter reloads from the Log.
- Re-saving an edited encounter stamps Log rows `1`, `1 (v2)`, `1 (v3)`, …
- On open, patient and meds clear and the Log is kept. On close, patient, meds, **and the Log** clear, then the workbook saves. The day's CSV in `dispense-log/` is the record that survives.

## What not to break

- Logged printing goes through `PrintCheckedLabels` → `LogPrint`. `RowPrintExtra` and `PrintLogRow` must not call `LogPrint` or `MarkPrinted`.
- `PrintLogRow` renders from the Log row (`RenderLabelSurfaceFromLog`), not from the patient currently on the Input sheet.
- Edit and Remove on the Log update that row's line in `dispense-log/YYYY-MM-DD.csv`. A hand-edit or a hand-added row does the same (`LogSheetChanged`). The match is timestamp + encounter + medication + lot. If that line is not in the file yet, the row is appended. Do not rewrite the whole CSV. The on-close wipe does not delete the CSV.
- Anything that writes a Medications cell and needs the result to stick (`ValidateMedications`, `ToggleRowSelect`, `ClearMedArea`) runs with `EnableEvents = False`. `AppReady` turns events back on at the start of the main buttons.
- The custom `IIf` in this module returns a **String**. Use a real `If` for numbers and flags.
- Do not set `PaperSize`. Fit-to-page (`FitToPagesWide/Tall = 1`) is what keeps Exp/Lot on the label.
- Sheet event code is injected at **build** time (`InstallMedSheetEvents`, `InstallAutoRefresh`, `InstallTebraAutoRefresh`). Do not paste the old column-15/16 handlers from the setup docs.

## Still open in the code

- If Hermione is not found, `RowPrintExtra` and `PrintLogRow` show the Windows print dialog and then still call `PrintLabelSurfaceSafe`. That can print twice. `PrintLabel` returns after the dialog.
- `LABEL_WIDTH_PT = 242` has not been re-checked on the clinic Brother (228 was the earlier no-bleed width).
- `tools/Build-ScuEmblem.ps1` blanks the emblem. Do not run it.
- The published **v2.4** release page still tells people to download `SCU-Label-Printing-v2.3.zip`. The attached asset is `SCU-Label-Printing-v2.4.zip` (the full folder). `tools/make-release-zip.ps1` still builds an older slim zip (workbook + emblem + card + `INSTALL.txt`) and is not what that release shipped.
- `support/_backups/` holds old `MedParser.bas` snapshots. They are not the current source.

## Git and PHI

`MedicationDispensing.xlsm` **is tracked**. Close wipes patient data before the save, and the rule is to commit only that wiped workbook. `*.csv`, `dispense-log/`, and `support/_backups/` are git-ignored. `dist/` is git-ignored. Release ZIPs are attached to GitHub Releases.
