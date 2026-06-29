# Troubleshooting

## Run-time error 9: Subscript out of range

Typical failing line:

```vb
Set ws1 = ThisWorkbook.Sheets(SH_INPUT)
```

Cause: the workbook does not contain a required sheet with the exact expected name.

Required sheets:

```text
Patient & Input
Medications
Label Preview
Log
Label Previews
```

Fix:

- Start from the clean no-PHI workbook template, not a blank workbook.
- Or manually create/rename the missing sheets exactly.
- Future enhancement: add `EnsureWorksheet()` bootstrap logic to create missing sheets automatically.

## Unable to import file when importing MedParser.bas

Known causes:

1. `MedParser.bas` has Unix LF line endings instead of Windows CRLF.
2. The file contains non-ASCII characters.
3. The file was edited by a tool that changed encoding.

Fix:

- Save `MedParser.bas` as ASCII or ANSI-compatible text.
- Save with Windows CRLF line endings.
- The repo includes `.gitattributes` to force CRLF for `.bas` and `.vbs` files.

## Unable to import file during SetupWorkbook

If the error occurs while running `SetupWorkbook`, it may not be a `.bas` import problem. Excel can also show this error when `Shapes.AddPicture` fails.

Known cause: large, corrupt, or OneDrive-cloud-only PNG logo file.

Current fix: logo import is disabled. If re-enabled, use a small local PNG and wrap `AddPicture` in error handling so a bad image never breaks setup.

## Macros are blocked

Cause: Office Mark-of-the-Web or Trust Center policy.

Fix:

- Store the workbook in a known local SCU project folder.
- Add that folder as an Excel Trusted Location.
- Reopen the workbook.

Do not broadly trust Downloads or unknown folders.

## Run-time error 1004: Unable to set PaperSize

Cause: Excel/VBA cannot reliably set custom Brother DK label paper size using `PageSetup.PaperSize`.

Fix:

- Do not use `.PaperSize = xlPaperUser` in VBA.
- Set DK-1202 / 62 x 100 mm in the Brother driver defaults.
- Let Excel control only print area, orientation, margins, and zoom.

## Printer sends 4.07 x 6.4 instead of DK-1202

Cause: Brother printer defaults still expose the larger shipping-label media size.

Fix:

Check both:

```text
Printing preferences
Printer properties -> Advanced -> Printing Defaults
```

Set both to DK-1202 / 62 x 100 mm, then fully restart Excel.

## Label prints squished

Cause: Excel fit-to-page scaling.

Fix: VBA should use:

```vb
.Zoom = 100
.FitToPagesWide = False
.FitToPagesTall = False
```

Do not use:

```vb
.Zoom = False
.FitToPagesWide = 1
.FitToPagesTall = 1
```

## Double-click Print? does not toggle

Check the `Medications` sheet module. It must include the `Worksheet_BeforeDoubleClick` handler.

Current column layout:

```text
O / 15 = # of Prints
P / 16 = Print?
```

The double-click handler should toggle column 16.

## Label Previews does not refresh

Check the `Label Previews` sheet module. It must include:

```vb
Private Sub Worksheet_Activate()
    On Error Resume Next
    Application.EnableEvents = False
    PreviewAllLabels
    Application.EnableEvents = True
End Sub
```

Also make sure macros are enabled.

## Git/OneDrive problems

Avoid running non-Windows Git/automation against a `.git` directory inside OneDrive. Use GitHub Desktop on native Windows and keep the source repo outside OneDrive when possible.

If a bad `git init` created a broken local repo, the safest fix is usually:

1. Do not delete the working files.
2. Clone the GitHub repo into a clean folder.
3. Copy source/docs over, excluding the old `.git` folder.
4. Commit from the clean clone.
