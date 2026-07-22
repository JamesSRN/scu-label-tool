# Setup Instructions

These instructions are for setting up the SCU Label Tool on a new Windows computer.

## 1. What this setup creates

This setup creates a local Excel/VBA medication label workflow:

```text
Tebra medication text -> Excel paste box -> VBA parser -> reviewed rows -> Brother QL-1100C label printing
```

Everything should run locally in Excel/VBA. No PHI should be committed to GitHub.

## 2. Requirements

- Windows PC
- Microsoft Excel desktop app
- Brother QL-1100C label printer
- Brother DK-1202 labels, 62 x 100 mm / approximately 2.4 in x 3.9 in
- GitHub Desktop recommended
- Clean no-PHI `MedicationDispensing.xlsm` workbook template from the project owner

This workflow is Windows-only. The parser uses Windows/Excel components such as `VBScript.RegExp`, and the print workflow targets the Windows Brother driver.

## 3. Clone the repository

Using GitHub Desktop:

1. Open GitHub Desktop.
2. Clone `JamesSRN/scu-label-tool`.
3. Choose a local folder outside OneDrive if possible, for example:

```text
C:\Users\<you>\source\scu-label-tool
```

Using command line:

```bash
git clone https://github.com/JamesSRN/scu-label-tool.git
cd scu-label-tool
```

Use native Windows Git/GitHub Desktop for this repo. Avoid managing a OneDrive-backed `.git` folder from WSL or automation.

## 4. Install and configure the Brother printer

1. Plug in and power on the Brother QL-1100C.
2. Install the Brother full driver/software package if Windows does not configure it correctly.
3. Confirm the printer appears in Windows printer settings.
4. Set the printer default media to DK-1202 / 62 x 100 mm. In the Brother driver this
   size is named **"Shipping Label"** (Width 2.44" x Length 3.93") - there is no entry
   literally called "2.4 x 3.9". Setting *Printing Defaults* (below) needs **admin
   rights**; without them the change silently reverts on Apply.

Check both places if Windows exposes both:

```text
Control Panel
-> Devices and Printers
-> Brother QL-1100C
-> Printing preferences
-> DK-1202 / 62 x 100 mm
```

and:

```text
Printer properties
-> Advanced
-> Printing Defaults
-> DK-1202 / 62 x 100 mm
```

Fully close and reopen Excel after changing printer defaults. Excel can cache printer/media settings.

## 5. Prepare the workbook

Do not import the VBA into a blank workbook unless you are intentionally testing bootstrap behavior.

Start from the clean no-PHI `MedicationDispensing.xlsm` workbook template. The workbook should already contain these sheets:

```text
Patient & Input
Medications
Label Preview
Log
Label Previews
```

`Label Preview` is the hidden/internal print surface. Do not delete it.

## 6. Trust the local workbook folder

If macros are blocked:

1. In Excel: `File -> Options -> Trust Center -> Trust Center Settings`.
2. Add the local SCU workbook folder as a Trusted Location.
3. Close and reopen Excel.

Only do this for the known local project folder. Do not broadly trust Downloads or random folders.

## 7. Import the VBA source

Manual method:

1. Open the clean no-PHI workbook in Excel.
2. Click `Enable Content` if appropriate.
3. Press `Alt + F11`.
4. In Project Explorer, expand the workbook.
5. Under `Modules`, remove the old `MedParser` module if present. Choose `No` when asked to export.
6. Import `src/MedParser.bas`.
7. Run `Debug -> Compile VBAProject`.
8. Save as `.xlsm`.

Important: `src/MedParser.bas` must remain ASCII with Windows CRLF line endings. This repo includes `.gitattributes` to help preserve CRLF for `.bas` files.

## 8. Worksheet event handlers

**V2 note:** `SetupWorkbook` now **re-installs the Medications `Worksheet_BeforeDoubleClick`
handler automatically at build time** (`InstallMedSheetEvents`), built from the `C_*`
column constants so it tracks the current columns (# of Prints = 16, Print? = 17) after
the V2 reorder. You normally do **not** need to paste it by hand. The **Label Previews**
`Worksheet_Activate` handler is preinstalled in that sheet module. The steps below are the
manual fallback if you are hand-building a fresh workbook.

### Medications sheet module

In the VBA editor, double-click the sheet module for **Medications** under `Microsoft Excel Objects`. Replace anything there with:

```vb
Private mP15Addr As String
Private mP15Val As Variant

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count = 1 And Target.Column = 15 And Target.Row > 3 Then
        mP15Addr = Target.Address
        mP15Val = Target.Value
    Else
        mP15Addr = ""
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Protect the auto-managed "# of Prints" column from manual edits
    If mP15Addr = "" Then Exit Sub
    Dim c As Range
    For Each c In Target.Cells
        If c.Address = mP15Addr Then
            Application.EnableEvents = False
            c.Value = mP15Val
            Application.EnableEvents = True
            MsgBox "The '# of Prints' column updates automatically and cannot be edited by hand.", _
                   vbInformation, "Protected column"
            Exit Sub
        End If
    Next c
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column = 16 Then        ' # of Prints - read-only
        Cancel = True
        Exit Sub
    End If
    If Target.Column = 17 And Target.Row > 3 Then   ' Print? - toggle the check
        If Trim(Me.Cells(Target.Row, 2).Value) <> "" Then
            Cancel = True
            ToggleRowSelect Target.Row
        End If
    End If
End Sub
```

Current column layout (V2 - see HANDOFF §2 for the full map):

```text
14 = Raw text (hidden)   15 = Printed? (hidden)
16 = # of Prints         17 = Print?
```

(The `Worksheet_SelectionChange` / `Worksheet_Change` guard above still references
col 15; in V2 the "# of Prints" column is 16, so update those two `15` values to `16`
if you hand-install. The auto-installed double-click handler already blocks editing 16.)

### Label Previews sheet module

In the VBA editor, double-click the sheet module for **Label Previews**. If the workbook has not yet migrated the name, it may still appear as **All Labels**. Paste:

```vb
Private Sub Worksheet_Activate()
    On Error Resume Next
    Application.EnableEvents = False
    PreviewAllLabels
    Application.EnableEvents = True
End Sub
```

Then save the workbook.

## 9. Run SetupWorkbook

In Excel:

1. Press `Alt + F8`.
2. Select `SetupWorkbook`.
3. Click `Run`.
4. Save the workbook.

Expected result:

- Buttons are created/refreshed.
- `Print?` and `# of Prints` columns exist.
- `Label Preview` internal print surface is built and hidden.
- `Label Previews` gallery is available.
- Row coloring and validation behaviors are ready.

## 10. Test with no-PHI sample text

Use:

```text
test-data/sample_tebra_pastes_no_phi.txt
```

Minimum test:

1. Enter fake patient data.
2. Paste sample medication text.
3. Click `PARSE MEDICATIONS`.
4. Review/correct rows; pick a **Source** for each (yellow until set) and fill any
   red Exp/Lot cells.
5. Enter fake Expiration and Lot (the prompt after parse, or via **+ Add Medication**).
6. Click `Review & Validate` - passing rows auto-check for printing.
7. Double-click the `Print?` cell to toggle any row manually.
8. Click `Print Checked Labels` (prints 2 copies of each).
9. Confirm a physical test label prints correctly, **including the bottom Exp/Lot row**,
   and that the Log shows one green-banded **Encounter** for the batch.

Do not test with real patient data until setup and printing have been validated locally.

## 11. Known setup failure: run-time error 9

If `SetupWorkbook` fails on:

```vb
Set ws1 = ThisWorkbook.Sheets(SH_INPUT)
```

then the workbook is missing the expected sheet name. Use the clean template workbook or manually create/rename the required sheets exactly. The current source is not yet a complete blank-workbook bootstrapper.

## 12. Release checklist

Before merging to `main` or publishing a release:

- No PHI in committed files.
- `src/MedParser.bas` is ASCII and CRLF.
- `Debug -> Compile VBAProject` passes.
- `SetupWorkbook` runs cleanly.
- Sheet handlers are installed.
- Parser works with no-PHI samples.
- Brother printer defaults are DK-1202 / 62 x 100 mm.
- Physical label print is readable, landscape, not clipped, and not squished.
- `HANDOFF.md`, `CHANGELOG.md`, and setup docs are updated.
