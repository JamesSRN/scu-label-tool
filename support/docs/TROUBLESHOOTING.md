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

- Keep the extracted tool folder on your **Desktop** (not OneDrive or Downloads).
- In Excel's Trust Center, add the **Desktop** as a Trusted Location with **Subfolders of this location are also trusted** — every version extracted to the Desktop is then trusted.
- Reopen the workbook.

Trust the Desktop on the dedicated clinic PC; don't trust Downloads or synced/cloud folders.

## Run-time error 1004: Unable to set PaperSize

Cause: Excel/VBA cannot reliably set custom Brother DK label paper size using `PageSetup.PaperSize`.

Fix:

- Do not use `.PaperSize = xlPaperUser` in VBA.
- Set DK-1202 / 62 x 100 mm in the Brother driver defaults.
- Let Excel control only print area, orientation, margins, and zoom.

## The Brother does not list "2.4 x 3.9" / DK-1202 as a paper size

This is usually **not** a missing size - the Brother driver lists the DK-1202 die-cut
label under the name **"Shipping Label"** (dimensions shown beneath it as
**Width 2.44" x Length 3.93"** = 62 x 100 mm). Select **Shipping Label**. There is no
entry literally called "2.4 x 3.9"; the dropdown names labels, not dimensions.

Also confirm the physically-loaded roll is the DK-1202 - the driver only offers die-cut
sizes that match the detected roll.

## Printer sends 4.07 x 6.4 (or 4x6) instead of DK-1202

Cause: the Brother QL-1100c (Hermione) supports both DK-1202 (62 x 100) and larger 4x6 media, and its
**default** is sitting on the wrong one.

Fix - set it in the place applications actually inherit from:

```text
Control Panel -> Devices and Printers -> right-click Brother QL-1100c (Hermione)
  -> Printer properties -> Advanced tab -> Printing Defaults... -> Paper Size = Shipping Label (62 x 100 mm)
```

Note: *Printing Defaults* (under Advanced) is different from the per-user *Printing
Preferences*. Changing Printing Defaults needs **admin rights** - on a locked-down clinic
laptop, without admin the change silently reverts when you click Apply (that revert is the
tell-tale sign). The tool's label sheet also specifies 62 x 100 in its own Page Setup, so
prints are usually correct even when the Windows default looks wrong.

## Labels print but the bottom row (Expiration / Lot) is missing

Cause: the EXP/LOT fields sit on the last row of the label print area. If the printer's
printable height is a hair shorter than the design, that row spills onto a never-printed
"page 2" and disappears - while still showing in the on-screen gallery (no page limit).
This is driver-dependent, so identical code can print fine on one PC and drop the row on
another.

Fix (already in the code): the label page setup fits the whole label to one page, so the
bottom row can't fall off any printer:

```vb
.Zoom = False
.FitToPagesWide = 1
.FitToPagesTall = 1
```

If a label ever looks slightly small, it is because fit-to-page scaled it to that driver's
printable area - trim the label's spacer rows rather than turning fit-to-page off.

## Reset Session / Start New Patient leaves rows in the Medications tab

Fixed in code: all clear paths (Reset Session, Start New Patient, on-open auto-reset) now
call `ClearMedArea`, which wipes a **fixed range** (rows 3-503, all table columns) instead
of scanning the Name column for the last used row. The old scan skipped rows that had data
in other columns but a blank Name. If you still see leftovers on an old build, select the
row numbers and Delete, then rebuild.

## Double-click Print? does not toggle

Check the `Medications` sheet module - it must include the `Worksheet_BeforeDoubleClick`
handler. As of V2 the handler is **re-installed automatically at build time** by
`SetupWorkbook` (`InstallMedSheetEvents`) using the column constants, so it stays correct
after column reorders.

Current column layout (after the V2 reorder):

```text
P / 16 = # of Prints
Q / 17 = Print?
```

The handler toggles the Print? column (17) and blocks editing # of Prints (16).

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


## Build aborts: "Source pre-check FAILED - encoding or structure problem"

`Build-Release.vbs` runs `tools\check-encoding.ps1` before building and aborts if
`MedParser.bas` is not **pure ASCII with CRLF line endings**. The usual cause is a
non-ASCII typographic character that slipped into `MedParser.bas` - an em-dash (—), a
curly quote (’ “ ”), an ellipsis (…), an arrow (→), or × - often pasted from a word
processor, a browser, or an AI/code assistant.

Fix:

- Run the checker directly to see the offending file and line:
  `powershell -ExecutionPolicy Bypass -File tools\check-encoding.ps1`
- Replace the character with its ASCII equivalent (`-`, `'`, `"`, `...`, `->`, `x`) and rebuild.

Only `MedParser.bas` is enforced - it is imported into the VBA project, where non-ASCII
breaks the import. `Build-Release.vbs` runs under Windows Script Host and may keep
non-ASCII in its comments, so it is intentionally **not** ASCII-checked (that also keeps
its bytes stable for antivirus reputation - see below).

## Windows Defender or Chrome flags the download / the tool as a "virus"

This is a **false positive**, not real malware. `Build-Release.vbs` opens Excel and writes
code into the workbook's VBA project on every run - a legitimate part of how the tool
rebuilds itself from source, but the same behavior antivirus heuristics use to spot
script-based malware. Defender/SmartScreen and Chrome also lean on **file reputation**: the
released launcher is a known, trusted file, but a freshly-edited copy (even a one-character
change) becomes an unknown script and gets flagged until it earns reputation. That is why an
identical-behaving tool can suddenly be blocked right after a source edit.

What to do (do **not** disable antivirus wholesale - a folder exclusion is enough):

- Prefer the **official release ZIP**, and don't re-edit `Build-Release.vbs` - keeping its
  bytes stable keeps it a trusted file. Only `MedParser.bas` ever needs editing.
- **Chrome refused the download:** open `chrome://downloads`, find the blocked item, open its
  menu, and choose **Keep**.
- **Right-click the extracted `.vbs`/`.cmd` -> Properties -> Unblock** if Windows added a
  "downloaded from the internet" mark.
- **Add the tool's folder (or the Desktop) as a Defender exclusion:** Windows Security ->
  Virus & threat protection -> Manage settings -> Exclusions -> Add or remove exclusions ->
  Add a folder.
- **Report the false positive to Microsoft** at <https://www.microsoft.com/wdsi/filesubmission>
  so it stops flagging on clinic PCs.
