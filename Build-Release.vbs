' ============================================================
'  Build-Release.vbs
'  Produces the definitive MedicationDispensing.xlsm:
'    1. imports the newest MedParser.bas
'    2. runs SetupWorkbook (this compiles the project + builds buttons)
'    3. saves the workbook as .xlsm
'
'  HOW TO RUN:
'    a) Close MedicationDispensing.xlsm in Excel first (so it isn't locked).
'    b) One-time: in Excel, File > Options > Trust Center > Trust Center
'       Settings > Macro Settings > check "Trust access to the VBA project
'       object model" > OK.  (Required for any script to import VBA code.)
'    c) Double-click this file.  Click OK on the "Setup complete" dialog
'       when it appears.  A final message confirms success.
' ============================================================
Option Explicit

Dim fso, scriptDir, xlsmPath, basPath, xl, wb, proj
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlsmPath  = fso.BuildPath(scriptDir, "MedicationDispensing.xlsm")
basPath   = fso.BuildPath(scriptDir, "MedParser.bas")

If Not fso.FileExists(xlsmPath) Then
    MsgBox "Cannot find:" & vbCrLf & xlsmPath, vbCritical, "Build-Release"
    WScript.Quit 1
End If
If Not fso.FileExists(basPath) Then
    MsgBox "Cannot find:" & vbCrLf & basPath, vbCritical, "Build-Release"
    WScript.Quit 1
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
On Error Resume Next
xl.AutomationSecurity = 1   ' msoAutomationSecurityLow - allow the workbook's macros to run
On Error GoTo 0

Set wb = xl.Workbooks.Open(xlsmPath)

' --- Verify programmatic access to the VBA project ---
On Error Resume Next
Set proj = wb.VBProject
If Err.Number <> 0 Then
    MsgBox "Cannot access the VBA project." & vbCrLf & vbCrLf & _
        "Enable it once in Excel:" & vbCrLf & _
        "File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
        "Macro Settings > check 'Trust access to the VBA project object model'," & vbCrLf & _
        "click OK, then run this script again.", vbExclamation, "Build-Release"
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If
On Error GoTo 0

' --- Remove any existing MedParser module, then import the newest one ---
On Error Resume Next
proj.VBComponents.Remove proj.VBComponents("MedParser")
On Error GoTo 0
proj.VBComponents.Import basPath

' --- Run SetupWorkbook (forces a full compile of its call tree + builds buttons) ---
On Error Resume Next
xl.Run "SetupWorkbook"
If Err.Number <> 0 Then
    MsgBox "SetupWorkbook failed - this usually means a compile error in the module:" & _
        vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
        "The workbook was left open and NOT saved so you can inspect it.", _
        vbCritical, "Build-Release"
    WScript.Quit 1
End If
On Error GoTo 0

' --- Save as macro-enabled workbook (.xlsm) ---
wb.Save

MsgBox "Release build complete." & vbCrLf & vbCrLf & _
    "- MedParser.bas imported" & vbCrLf & _
    "- SetupWorkbook ran successfully (project compiles)" & vbCrLf & _
    "- MedicationDispensing.xlsm saved" & vbCrLf & vbCrLf & _
    "This .xlsm now contains the current code. You can close Excel.", _
    vbInformation, "Build-Release - Done"
