' ============================================================
'  Build-Release.vbs  (robust: never leaves a stray Excel process)
'  Produces the definitive MedicationDispensing.xlsm:
'    1. opens MedicationDispensing.xlsm (bootstraps from
'       Broken_PrettyPrint_MedicationDispensing.xlsm only if missing)
'    2. imports the newest MedParser.bas
'    3. runs SetupWorkbook (compiles the project + builds buttons)
'    4. saves AND closes the workbook
'  On ANY failure it shows a message and quits Excel cleanly, so no
'  background Excel process is left holding the file.
'
'  RUN: close MedicationDispensing.xlsm first, then double-click this.
'  One-time: Excel > File > Options > Trust Center > Trust Center
'  Settings > Macro Settings > check "Trust access to the VBA project
'  object model".
' ============================================================
Option Explicit

Dim fso, scriptDir, bootstrapPath, xlsmPath, basPath
Dim xl, wb, proj, macroName, setupErr, saveErr

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
bootstrapPath = fso.BuildPath(scriptDir, "Broken_PrettyPrint_MedicationDispensing.xlsm")
xlsmPath = fso.BuildPath(scriptDir, "MedicationDispensing.xlsm")
basPath = fso.BuildPath(scriptDir, "MedParser.bas")

If Not fso.FileExists(xlsmPath) Then
    If Not fso.FileExists(bootstrapPath) Then
        MsgBox "Cannot find MedicationDispensing.xlsm or bootstrap workbook:" & vbCrLf & _
            xlsmPath & vbCrLf & bootstrapPath, vbCritical, "Build-Release"
        WScript.Quit 1
    End If
    MsgBox "MedicationDispensing.xlsm was missing - copying from the bootstrap template." & vbCrLf & vbCrLf & _
        "Patient data in a previous .xlsm is NOT restored. Check _backups folder if you have one.", _
        vbExclamation, "Build-Release"
    fso.CopyFile bootstrapPath, xlsmPath, True
End If
If Not fso.FileExists(basPath) Then
    MsgBox "Cannot find:" & vbCrLf & basPath, vbCritical, "Build-Release"
    WScript.Quit 1
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
xl.AutomationSecurity = 1
xl.EnableEvents = True

On Error Resume Next
Set wb = xl.Workbooks.Open(xlsmPath, 0, False)
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not open MedicationDispensing.xlsm (it may be open elsewhere or locked)." & vbCrLf & vbCrLf & _
        "Close it and end any stray Microsoft Excel in Task Manager, then run again.", _
        vbExclamation, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If
On Error GoTo 0
wb.Activate

' Guard: if it opened read-only, it is locked/read-only elsewhere. Abort cleanly.
If wb.ReadOnly Then
    MsgBox "MedicationDispensing.xlsm opened READ-ONLY, so it cannot be saved." & vbCrLf & vbCrLf & _
        "It is most likely still open in another Excel window, or a stray Excel " & _
        "process is holding it (end them in Task Manager), or the file is marked " & _
        "read-only (right-click > Properties > uncheck Read-only)." & vbCrLf & vbCrLf & _
        "Nothing was changed. Excel has been closed. Fix that and run again.", _
        vbExclamation, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If

' Verify programmatic access to the VBA project.
On Error Resume Next
Set proj = wb.VBProject
If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    MsgBox "Cannot access the VBA project." & vbCrLf & vbCrLf & _
        "Enable it once in Excel:" & vbCrLf & _
        "File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
        "Macro Settings > check 'Trust access to the VBA project object model'," & vbCrLf & _
        "click OK, then run this script again.", vbExclamation, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If
On Error GoTo 0

' Remove any existing MedParser module(s), then import the newest one.
On Error Resume Next
proj.VBComponents.Remove proj.VBComponents("MedParser")
proj.VBComponents.Remove proj.VBComponents("MedParser1")
Err.Clear
On Error GoTo 0
proj.VBComponents.Import basPath

' Run SetupWorkbook (forces a full compile + builds buttons).
macroName = "'" & wb.Name & "'!SetupWorkbook"
On Error Resume Next
xl.Run macroName
If Err.Number <> 0 Then
    setupErr = Err.Description
    Err.Clear
    On Error GoTo 0
    MsgBox "SetupWorkbook failed:" & vbCrLf & vbCrLf & setupErr & vbCrLf & vbCrLf & _
        "This is usually a compile error. Import MedParser.bas manually (Alt+F11) and " & _
        "use Debug > Compile VBAProject to see the exact line. Nothing was saved.", _
        vbCritical, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If
On Error GoTo 0

' Save; on failure clean up so no Excel process is left behind.
On Error Resume Next
wb.Save
If Err.Number <> 0 Then
    saveErr = Err.Description
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not save MedicationDispensing.xlsm:" & vbCrLf & vbCrLf & saveErr & vbCrLf & vbCrLf & _
        "The file is open elsewhere or read-only. Close it (and end stray Excel in " & _
        "Task Manager), then run again. Nothing was saved; Excel has been closed cleanly.", _
        vbCritical, "Build-Release"
    CleanupQuit
    WScript.Quit 1
End If
On Error GoTo 0

' Success: close cleanly so nothing is left holding the file.
wb.Close False
xl.Quit
Set wb = Nothing
Set proj = Nothing
Set xl = Nothing

MsgBox "Release build complete." & vbCrLf & vbCrLf & _
    "- MedParser.bas imported" & vbCrLf & _
    "- SetupWorkbook ran (project compiles)" & vbCrLf & _
    "- MedicationDispensing.xlsm saved and closed" & vbCrLf & vbCrLf & _
    "Open MedicationDispensing.xlsm to use it.", vbInformation, "Build-Release - Done"

Sub CleanupQuit
    On Error Resume Next
    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set proj = Nothing
    Set xl = Nothing
    On Error GoTo 0
End Sub
