# Setup Instructions

These instructions set up the **SCU Dispensary Label Tool** (v2.4) on a Windows PC. Source is `MedParser.bas` at the repo root. There is no `src/` folder.

## 1. What this setup creates

```text
Tebra medication text -> Excel paste box -> VBA parser -> reviewed rows -> Brother QL-1100c (Hermione) label printing
```

The tool runs locally in Excel. Dated dispense CSVs stay on the PC.

## 2. Requirements

- Windows PC
- Microsoft Excel desktop
- Brother QL-1100c (Hermione), USB cable (not Bluetooth)
- Brother DK-1202 labels, 62 x 100 mm. In the Brother driver this size is named **Shipping Label** (about 2.44 in x 3.93 in). There is no menu entry literally called "2.4 x 3.9".
- The release ZIP, or a clone of `JamesSRN/scu-label-tool`

## 3. Get the files

Clinic PC (usual path):

1. Download `SCU-Label-Printing-vX.Y.zip` from the [latest release](https://github.com/JamesSRN/scu-label-tool/releases/latest).
2. Move the ZIP to the Desktop. Right-click → **Extract All**. Do not run files from inside the ZIP window.

Maintainer clone:

```bash
git clone https://github.com/JamesSRN/scu-label-tool.git
```

Keep the folder outside OneDrive. Use GitHub Desktop or Windows Git, not a OneDrive-backed `.git` folder from WSL.

## 4. Brother printer

1. Plug in the USB cable and power on Hermione.
2. Install the Brother driver if Windows does not set it up.
3. Set the media to DK-1202 / Shipping Label in both **Printing preferences** and **Printer properties → Advanced → Printing Defaults**.
4. Changing Printing Defaults needs admin rights. Without them, Apply can silently revert.

Fully close and reopen Excel after changing printer defaults. Excel caches media settings.

## 5. One-time Excel trust

The click-me file rebuilds the workbook from `MedParser.bas` every time it opens, so Excel has to trust the VBA project and the Desktop.

1. Excel → **File → Options → Trust Center → Trust Center Settings**.
2. **Macro Settings** — check **Trust access to the VBA project object model**.
3. **Trusted Locations → Add new location** — the Desktop, with **Subfolders of this location are also trusted**.
4. Close Excel.

## 6. Every open

1. Close `MedicationDispensing.xlsm` if it is open.
2. Double-click **`OPEN LABEL TOOL (double-click me).cmd`**.
3. If that does nothing, double-click **`Build-Release.vbs`**.
4. The workbook opens on **Start Here**. Click **Enable Content** if Excel asks.

Do not start from a blank workbook. If `MedicationDispensing.xlsm` is missing, Build-Release copies `support/Broken_PrettyPrint_MedicationDispensing.xlsm` and warns that previous patient data is not restored.

## 7. Manual import (only if the click-me file cannot run)

1. Alt+F11. Remove the old `MedParser` module. Do not export it.
2. **File → Import** `MedParser.bas` from the repo root (not `src/MedParser.bas`).
3. Run `SetupWorkbook`. That compiles the project and re-installs the sheet handlers from the current column constants.
4. Save.

`MedParser.bas` must stay ASCII with Windows CRLF. `tools/check-encoding.ps1` checks that file.

Do not paste an older sheet module. Check Med is column 2. `# of Prints` is column 14. The build injects `Worksheet_BeforeDoubleClick` and `Worksheet_Change` on Medications, `Worksheet_Activate` on Print Labels, and `Worksheet_Activate` on Tebra Notes.

## 8. Smoke test (no real patients)

1. Developer Test → **Generate Test Patient**, or paste `support/test-data/sample_tebra_pastes_no_phi.txt`.
2. **PARSE MEDICATIONS**.
3. Fill Expiration, Lot, and Source. Yellow cells should turn white as you type.
4. **Review**. Complete rows turn blue. They are not checked for you.
5. Double-click Check Med (or the Check Med header) so the rows you want turn green.
6. **Print Checked Labels**. Two copies of each. Cancelling the initials prompt prints nothing.
7. You should land on the Log. Open Tebra Notes and confirm the note matches the Log.

## 9. Before a release

- `MedParser.bas` passes `tools/check-encoding.ps1`.
- `SetupWorkbook` compiles.
- Parser smoke test above, with fictional patients only.
- Physical DK-1202 label is one die-cut, landscape, Exp/Lot visible.
- `CHANGELOG.md` and `README.md` match the behavior you just shipped.
- Commit `MedicationDispensing.xlsm` only after Excel has closed it (close wipes the patient and the Log). Never commit `dispense-log/` or `support/_backups/`.
