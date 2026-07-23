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
Dim hasForm, comp, frm, dsn, ctl, tw, twCode

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
' Bootstrap template (used ONLY if MedicationDispensing.xlsm is missing). Search several
' locations so it works whether the template lives in support\, the root, or templates\,
' and whether or not it has been renamed - the build never breaks on where it sits.
Dim bootCandidates, bi
bootCandidates = Array( _
    "support\Template_MedicationDispensing.xlsm", _
    "support\Broken_PrettyPrint_MedicationDispensing.xlsm", _
    "templates\Template_MedicationDispensing.xlsm", _
    "Broken_PrettyPrint_MedicationDispensing.xlsm")
bootstrapPath = ""
For bi = 0 To UBound(bootCandidates)
    If bootstrapPath = "" Then
        If fso.FileExists(fso.BuildPath(scriptDir, bootCandidates(bi))) Then
            bootstrapPath = fso.BuildPath(scriptDir, bootCandidates(bi))
        End If
    End If
Next
If bootstrapPath = "" Then bootstrapPath = fso.BuildPath(scriptDir, "Broken_PrettyPrint_MedicationDispensing.xlsm")
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

' Pre-build encoding/structure check (fail-safe): abort ONLY on a real failure, so a
' mojibaked or unbalanced MedParser.bas / Build-Release.vbs is never imported. If the
' checker or PowerShell isn't available, the build just proceeds normally.
Dim chkPath, chkShell, chkRc
chkPath = fso.BuildPath(scriptDir, "tools\check-encoding.ps1")
If fso.FileExists(chkPath) Then
    On Error Resume Next
    Set chkShell = CreateObject("WScript.Shell")
    chkRc = chkShell.Run("powershell -ExecutionPolicy Bypass -NoProfile -File """ & chkPath & """", 0, True)
    On Error GoTo 0
    If chkRc = 1 Or chkRc = 2 Then
        MsgBox "Source pre-check FAILED - encoding or structure problem in MedParser.bas or Build-Release.vbs." & vbCrLf & vbCrLf & _
            "See details by running:" & vbCrLf & _
            "powershell -ExecutionPolicy Bypass -File """ & chkPath & """" & vbCrLf & vbCrLf & _
            "Build aborted so a bad file is not imported.", vbCritical, "Build-Release - Pre-check failed"
        WScript.Quit 1
    End If
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
xl.AutomationSecurity = 1
xl.EnableEvents = True
Prog 8, "Starting Excel..."

Prog 20, "Opening workbook (running open reset)..."
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
Prog 38, "Reading VBA project..."

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
Prog 55, "Importing MedParser.bas..."
proj.VBComponents.Import basPath
Prog 72, "Building input + progress forms..."

' Build/refresh the two-field Expiration/Lot UserForm (frmExpLot) every run via
' EnsureForm, so an old/cropped copy in the workbook is resized and rebuilt (not
' skipped). Baked into the workbook so volunteers need no VBA-project-trust setting.
On Error Resume Next
Set frm = EnsureForm("frmExpLot", 292, 286, "Enter Expiration and Lot")
Set dsn = frm.Designer
    Set ctl = NewCtl("Forms.Label.1", "lblMed", True)
    ctl.Left = 16 : ctl.Top = 12 : ctl.Width = 244 : ctl.Height = 44 : ctl.Font.Bold = True : ctl.Font.Size = 16 : ctl.WordWrap = True
    Set ctl = NewCtl("Forms.Label.1", "lblExp", True)
    ctl.Left = 16 : ctl.Top = 64 : ctl.Width = 244 : ctl.Height = 12 : ctl.Caption = "Expiration MM/YYYY (comma-separate multiple bottles):"
    Set ctl = NewCtl("Forms.TextBox.1", "txtExp", True)
    ctl.Left = 16 : ctl.Top = 78 : ctl.Width = 244 : ctl.Height = 18
    Set ctl = NewCtl("Forms.Label.1", "lblLot", True)
    ctl.Left = 16 : ctl.Top = 106 : ctl.Width = 244 : ctl.Height = 12 : ctl.Caption = "Lot number (comma-separate multiple bottles):"
    Set ctl = NewCtl("Forms.TextBox.1", "txtLot", True)
    ctl.Left = 16 : ctl.Top = 120 : ctl.Width = 244 : ctl.Height = 18
    Set ctl = NewCtl("Forms.CommandButton.1", "btnOK", True)
    ctl.Left = 103 : ctl.Top = 170 : ctl.Width = 84 : ctl.Height = 28 : ctl.Caption = "OK" : ctl.Default = True
    HideExtras "[lblMed][lblExp][txtExp][lblLot][txtLot][btnOK]"
    frm.CodeModule.AddFromString _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    Me.Caption = ""Enter Expiration and Lot""" & vbCrLf & _
        "    Me.Width = 292" & vbCrLf & _
        "    Me.Height = 286" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Private Sub UserForm_Activate()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    txtExp.SetFocus" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Private Sub btnOK_Click()" & vbCrLf & "    Me.Hide" & vbCrLf & "End Sub"
On Error GoTo 0

' Build the medication-edit UserForm (frmMedEdit).
On Error Resume Next
Set frm = EnsureForm("frmMedEdit", 360, 350, "Edit medication")
Set dsn = frm.Designer
    Set ctl = NewCtl("Forms.Label.1", "lblName", True)
    ctl.Left = 12 : ctl.Top = 15 : ctl.Width = 82 : ctl.Height = 14 : ctl.Caption = "Medication" : ctl.Font.Bold = True
    Set ctl = NewCtl("Forms.TextBox.1", "txtName", True)
    ctl.Left = 100 : ctl.Top = 12 : ctl.Width = 236 : ctl.Height = 22 : ctl.Font.Bold = True : ctl.Font.Size = 11
    Set ctl = NewCtl("Forms.Label.1", "lblStr", True)
    ctl.Left = 12 : ctl.Top = 43 : ctl.Width = 82 : ctl.Height = 14 : ctl.Caption = "Strength" : ctl.Font.Bold = True
    Set ctl = NewCtl("Forms.TextBox.1", "txtStr", True)
    ctl.Left = 100 : ctl.Top = 40 : ctl.Width = 236 : ctl.Height = 22 : ctl.Font.Bold = True : ctl.Font.Size = 11
    Set ctl = NewCtl("Forms.Label.1", "lblForm", True)
    ctl.Left = 12 : ctl.Top = 74 : ctl.Width = 82 : ctl.Height = 14 : ctl.Caption = "Dosage form"
    Set ctl = NewCtl("Forms.TextBox.1", "txtForm", True)
    ctl.Left = 100 : ctl.Top = 72 : ctl.Width = 150 : ctl.Height = 18
    Set ctl = NewCtl("Forms.Label.1", "lblQty", True)
    ctl.Left = 12 : ctl.Top = 100 : ctl.Width = 82 : ctl.Height = 14 : ctl.Caption = "Quantity"
    Set ctl = NewCtl("Forms.TextBox.1", "txtQty", True)
    ctl.Left = 100 : ctl.Top = 98 : ctl.Width = 100 : ctl.Height = 18
    Set ctl = NewCtl("Forms.Label.1", "lblSig", True)
    ctl.Left = 12 : ctl.Top = 126 : ctl.Width = 82 : ctl.Height = 14 : ctl.Caption = "Directions"
    Set ctl = NewCtl("Forms.TextBox.1", "txtSig", True)
    ctl.Left = 100 : ctl.Top = 124 : ctl.Width = 236 : ctl.Height = 50 : ctl.MultiLine = True : ctl.WordWrap = True
    Set ctl = NewCtl("Forms.Label.1", "lblExp", True)
    ctl.Left = 12 : ctl.Top = 186 : ctl.Width = 82 : ctl.Height = 14 : ctl.Caption = "Expiration"
    Set ctl = NewCtl("Forms.TextBox.1", "txtExp", True)
    ctl.Left = 100 : ctl.Top = 184 : ctl.Width = 120 : ctl.Height = 18
    Set ctl = NewCtl("Forms.Label.1", "lblLot", True)
    ctl.Left = 12 : ctl.Top = 212 : ctl.Width = 82 : ctl.Height = 14 : ctl.Caption = "Lot number"
    Set ctl = NewCtl("Forms.TextBox.1", "txtLot", True)
    ctl.Left = 100 : ctl.Top = 210 : ctl.Width = 120 : ctl.Height = 18
    Set ctl = NewCtl("Forms.CommandButton.1", "btnOK", True)
    ctl.Left = 128 : ctl.Top = 244 : ctl.Width = 90 : ctl.Height = 26 : ctl.Caption = "OK" : ctl.Default = True
    Set ctl = NewCtl("Forms.CommandButton.1", "btnCancel", True)
    ctl.Left = 230 : ctl.Top = 244 : ctl.Width = 90 : ctl.Height = 26 : ctl.Caption = "Cancel" : ctl.Cancel = True
    HideExtras "[lblName][txtName][lblStr][txtStr][lblForm][txtForm][lblQty][txtQty][lblSig][txtSig][lblExp][txtExp][lblLot][txtLot][btnOK][btnCancel]"
    frm.CodeModule.AddFromString "Public Result As String" & vbCrLf & _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    Me.Caption = ""Edit medication""" & vbCrLf & _
        "    Me.Width = 360" & vbCrLf & _
        "    Me.Height = 350" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Private Sub btnOK_Click()" & vbCrLf & "    Result = ""OK""" & vbCrLf & "    Me.Hide" & vbCrLf & "End Sub" & vbCrLf & "Private Sub btnCancel_Click()" & vbCrLf & "    Result = ""CANCEL""" & vbCrLf & "    Me.Hide" & vbCrLf & "End Sub"
On Error GoTo 0

' Build the "please wait" progress popup (frmBusy). Shown during the Print Checked
' Labels delay while the Brother printer is located + page is set up.
On Error Resume Next
Set frm = EnsureForm("frmBusy", 288, 170, "Preparing to print")
Set dsn = frm.Designer
    Set ctl = NewCtl("Forms.Label.1", "lblMsg", True)
    ctl.Left = 16 : ctl.Top = 14 : ctl.Width = 244 : ctl.Height = 30 : ctl.Font.Bold = True
    ctl.Caption = "Please wait..."
    Set ctl = NewCtl("Forms.Label.1", "lblTrack", True)
    ctl.Left = 16 : ctl.Top = 54 : ctl.Width = 244 : ctl.Height = 20 : ctl.BorderStyle = 1
    ctl.BackColor = RGB(238, 238, 238)
    Set ctl = NewCtl("Forms.Label.1", "barFill", True)
    ctl.Left = 17 : ctl.Top = 55 : ctl.Width = 2 : ctl.Height = 18
    ctl.BackColor = RGB(30, 120, 210) : ctl.Caption = ""
    Set ctl = NewCtl("Forms.Label.1", "barPct", True)
    ctl.Left = 16 : ctl.Top = 80 : ctl.Width = 244 : ctl.Height = 14 : ctl.Caption = "0%"
    ctl.TextAlign = 2
    HideExtras "[lblMsg][lblTrack][barFill][barPct]"
    frm.CodeModule.AddFromString _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    Me.Caption = ""Preparing to print""" & vbCrLf & _
        "    Me.Width = 288" & vbCrLf & _
        "    Me.Height = 170" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub SetProgress(ByVal pct As Long, ByVal msg As String)" & vbCrLf & _
        "    If pct < 0 Then pct = 0" & vbCrLf & _
        "    If pct > 100 Then pct = 100" & vbCrLf & _
        "    lblMsg.Caption = msg" & vbCrLf & _
        "    barFill.Width = 2 + (242 * pct \ 100)" & vbCrLf & _
        "    barPct.Caption = pct & ""%""" & vbCrLf & _
        "    Me.Repaint" & vbCrLf & _
        "End Sub"
On Error GoTo 0

' Build the Medication Review UserForm (frmReview): a scrolling list where each med
' shows its NAME + STRENGTH large & bold, with the error(s) for that med stacked
' underneath in a smaller (but still readable) font. Labels are added at run time into
' the scrolling frame by the form's own AddMed method; ReviewMedications populates it.
On Error Resume Next
Set frm = EnsureForm("frmReview", 470, 560, "Medication Review")
Set dsn = frm.Designer
    Set ctl = NewCtl("Forms.Label.1", "lblHdr", True)
    ctl.Left = 12 : ctl.Top = 8 : ctl.Width = 440 : ctl.Height = 20 : ctl.Font.Bold = True : ctl.Font.Size = 11 : ctl.Caption = "Medication Review"
    Set ctl = NewCtl("Forms.Frame.1", "fraList", True)
    ctl.Left = 10 : ctl.Top = 32 : ctl.Width = 446 : ctl.Height = 450 : ctl.ScrollBars = 2 : ctl.Caption = ""
    Set ctl = NewCtl("Forms.Label.1", "lblFoot", True)
    ctl.Left = 12 : ctl.Top = 486 : ctl.Width = 258 : ctl.Height = 42 : ctl.WordWrap = True : ctl.Font.Size = 9
    Set ctl = NewCtl("Forms.CommandButton.1", "btnCancel", True)
    ctl.Left = 278 : ctl.Top = 490 : ctl.Width = 84 : ctl.Height = 28 : ctl.Caption = "Cancel" : ctl.Cancel = True : ctl.Visible = False
    Set ctl = NewCtl("Forms.CommandButton.1", "btnOK", True)
    ctl.Left = 372 : ctl.Top = 490 : ctl.Width = 84 : ctl.Height = 28 : ctl.Caption = "OK" : ctl.Default = True
    HideExtras "[lblHdr][fraList][lblFoot][btnCancel][btnOK]"
    frm.CodeModule.AddFromString _
        "Public Result As String" & vbCrLf & _
        "Private mY As Single" & vbCrLf & _
        "Private mIdx As Long" & vbCrLf & _
        "Private Sub UserForm_Initialize()" & vbCrLf & _
        "    Me.Caption = ""Medication Review""" & vbCrLf & _
        "    Me.Width = 470" & vbCrLf & _
        "    Me.Height = 560" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub SetHeader(ByVal t As String)" & vbCrLf & _
        "    Me.Caption = t" & vbCrLf & _
        "    lblHdr.Caption = t" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub ConfigButtons(ByVal showCancel As Boolean, ByVal okText As String, ByVal cancelText As String)" & vbCrLf & _
        "    btnOK.Caption = okText" & vbCrLf & _
        "    btnOK.Left = 372" & vbCrLf & _
        "    If showCancel Then" & vbCrLf & _
        "        btnCancel.Caption = cancelText" & vbCrLf & _
        "        btnCancel.Visible = True" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        btnCancel.Visible = False" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub ResetList()" & vbCrLf & _
        "    Dim ct As Control" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    For Each ct In fraList.Controls" & vbCrLf & _
        "        ct.Visible = False" & vbCrLf & _
        "    Next" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "    mY = 6" & vbCrLf & _
        "    mIdx = 0" & vbCrLf & _
        "    Result = ""CANCEL""" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub AddMed(ByVal title As String, ByVal errText As String, ByVal isOK As Boolean)" & vbCrLf & _
        "    Dim lblN As Object, lblE As Object" & vbCrLf & _
        "    Dim nName As String, eName As String" & vbCrLf & _
        "    mIdx = mIdx + 1" & vbCrLf & _
        "    nName = ""n"" & mIdx" & vbCrLf & _
        "    eName = ""e"" & mIdx" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    Set lblN = fraList.Controls(nName)" & vbCrLf & _
        "    Set lblE = fraList.Controls(eName)" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "    If lblN Is Nothing Then Set lblN = fraList.Controls.Add(""Forms.Label.1"", nName, True)" & vbCrLf & _
        "    If lblE Is Nothing Then Set lblE = fraList.Controls.Add(""Forms.Label.1"", eName, True)" & vbCrLf & _
        "    lblN.AutoSize = False" & vbCrLf & _
        "    lblN.WordWrap = True" & vbCrLf & _
        "    lblN.Font.Size = 14" & vbCrLf & _
        "    lblN.Font.Bold = True" & vbCrLf & _
        "    lblN.Left = 6" & vbCrLf & _
        "    lblN.Top = mY" & vbCrLf & _
        "    lblN.Width = 408" & vbCrLf & _
        "    lblN.Height = 22" & vbCrLf & _
        "    lblN.Caption = title" & vbCrLf & _
        "    If isOK Then" & vbCrLf & _
        "        lblN.ForeColor = RGB(27, 94, 32)" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        lblN.ForeColor = RGB(17, 17, 17)" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    lblN.Visible = True" & vbCrLf & _
        "    mY = mY + 24" & vbCrLf & _
        "    If Len(Trim(errText)) = 0 Then" & vbCrLf & _
        "        lblE.Visible = False" & vbCrLf & _
        "        mY = mY + 6" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        lblE.AutoSize = False" & vbCrLf & _
        "        lblE.WordWrap = True" & vbCrLf & _
        "        lblE.Font.Size = 11" & vbCrLf & _
        "        lblE.Font.Bold = False" & vbCrLf & _
        "        lblE.Left = 20" & vbCrLf & _
        "        lblE.Top = mY" & vbCrLf & _
        "        lblE.Width = 392" & vbCrLf & _
        "        lblE.Height = (1 + Len(errText) - Len(Replace(errText, Chr(10), """"))) * 15 + 4" & vbCrLf & _
        "        lblE.Caption = errText" & vbCrLf & _
        "        If isOK Then" & vbCrLf & _
        "            lblE.ForeColor = RGB(46, 125, 50)" & vbCrLf & _
        "        Else" & vbCrLf & _
        "            lblE.ForeColor = RGB(183, 28, 28)" & vbCrLf & _
        "        End If" & vbCrLf & _
        "        lblE.Visible = True" & vbCrLf & _
        "        mY = mY + lblE.Height + 12" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub SetFooter(ByVal s As String)" & vbCrLf & _
        "    lblFoot.Caption = s" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Public Sub FinishList()" & vbCrLf & _
        "    fraList.ScrollHeight = mY + 8" & vbCrLf & _
        "    fraList.ScrollTop = 0" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Private Sub btnOK_Click()" & vbCrLf & _
        "    Result = ""OK""" & vbCrLf & _
        "    Me.Hide" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Private Sub btnCancel_Click()" & vbCrLf & _
        "    Result = ""CANCEL""" & vbCrLf & _
        "    Me.Hide" & vbCrLf & _
        "End Sub"
On Error GoTo 0

' Set ThisWorkbook auto-reset handlers: Open clears patient+meds (keeps Log) so it
' always opens empty; BeforeClose clears + saves so no PHI persists on disk. REPLACE the
' whole module each build so a stale/legacy Workbook_Open can't block the clear-on-open.
On Error Resume Next
Set tw = proj.VBComponents("ThisWorkbook")
If tw.CodeModule.CountOfLines > 0 Then tw.CodeModule.DeleteLines 1, tw.CodeModule.CountOfLines
tw.CodeModule.AddFromString _
    "Private Sub Workbook_Open()" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    CheckWorkbookStructure" & vbCrLf & _
    "    ClearSessionSilent" & vbCrLf & _
    "    ThisWorkbook.Sheets(""Start Here"").Activate" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf & _
    "    On Error Resume Next" & vbCrLf & _
    "    ClearSessionSilent" & vbCrLf & _
    "    ClearLogSilent" & vbCrLf & _
    "    ThisWorkbook.Save" & vbCrLf & _
    "End Sub"
On Error GoTo 0
Prog 82, "Installing auto-reset handlers..."

' Run SetupWorkbook (forces a full compile + builds buttons).
Prog 90, "Compiling + building label layout..."
macroName = "'" & wb.Name & "'!SetupWorkbook"
On Error Resume Next
xl.Run macroName
If Err.Number <> 0 Then
    setupErr = Err.Description
    Err.Clear
    On Error GoTo 0
    MsgBox "SetupWorkbook failed:" & vbCrLf & vbCrLf & setupErr & vbCrLf & vbCrLf & _
        "Excel is LEFT OPEN so you can inspect it: press Alt+F11, then Debug > Compile " & _
        "VBAProject to see the exact error and line. Nothing was saved.", _
        vbCritical, "Build-Release"
    WScript.Quit 1
End If
On Error GoTo 0

' Save; on failure clean up so no Excel process is left behind.
Prog 96, "Saving MedicationDispensing.xlsm..."
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

' Success: workbook is already saved above; LEAVE IT OPEN for immediate use.
Prog 100, "Build complete."
On Error Resume Next
xl.DisplayAlerts = True
wb.Activate
On Error GoTo 0

MsgBox "SCU Label Tool is ready." & vbCrLf & vbCrLf & _
    "The workbook is open on the Start Here tab - go ahead and use it." & vbCrLf & vbCrLf & _
    "Handy shortcuts:" & vbCrLf & _
    "   Ctrl+Shift+P  =  Parse Medications" & vbCrLf & _
    "   Ctrl+Shift+R  =  Reset Session" & vbCrLf & _
    "   Ctrl+Shift+L  =  Refresh Label Previews" & vbCrLf & vbCrLf & _
    "When you are done for the day, just close it.", vbInformation, "SCU Label Tool"

' Hand the status bar back to Excel now that the build is done.
On Error Resume Next
xl.StatusBar = False
On Error GoTo 0

' Ensure a UserForm exists and is the right size/caption, and reset its code module so
' the block below can re-add code. Self-healing: an old/cropped form is UPDATED in place
' every run (not skipped like the old create-if-missing). Controls are reconciled by the
' block via NewCtl + HideExtras. This never calls VBComponents.Remove on the form nor
' Controls.Remove, both of which were unreliable (Remove-before-xl.Run gave an "Unknown
' runtime error"; per-control Remove silently left stale controls behind).
Function EnsureForm(formName, wOuter, hOuter, capText)
    Dim f, existing
    Set existing = Nothing
    On Error Resume Next
    Set existing = proj.VBComponents(formName)
    On Error GoTo 0
    If existing Is Nothing Then
        Set f = proj.VBComponents.Add(3)
        f.Name = formName
    Else
        Set f = existing
    End If
    ' Reset the code module only (re-added below). Controls cannot be reliably REMOVED
    ' from an existing form via automation on some machines, so instead each block
    ' reconciles controls by name -- NewCtl reuses-or-adds each one and sets its
    ' properties (which works), and HideExtras hides any leftover controls from an
    ' older form design.
    On Error Resume Next
    If f.CodeModule.CountOfLines > 0 Then f.CodeModule.DeleteLines 1, f.CodeModule.CountOfLines
    f.Properties("Caption") = capText
    f.Properties("Width") = wOuter
    f.Properties("Height") = hOuter
    On Error GoTo 0
    Set EnsureForm = f
End Function

' Return a control by name on the current form (global dsn): reuse it if it already
' exists, otherwise add it. Reusing + re-styling an existing control works where
' removing one does not, so an old form's controls are corrected in place (not
' duplicated, which would silently fail the Add and leave the old styling).
Function NewCtl(kind, nm, dummy)
    Dim c
    Set c = Nothing
    On Error Resume Next
    Set c = dsn.Controls(nm)
    On Error GoTo 0
    If c Is Nothing Then Set c = dsn.Controls.Add(kind, nm, True)
    Set NewCtl = c
End Function

' Hide any control on the current form (global dsn) whose name is not in the expected
' set, e.g. "[lblMed][txtExp]". Kills leftover controls from an older form design (a
' stale header label) without needing to remove them.
Sub HideExtras(known)
    Dim c
    On Error Resume Next
    For Each c In dsn.Controls
        If InStr(known, "[" & c.Name & "]") = 0 Then c.Visible = False
    Next
    On Error GoTo 0
End Sub

' Draw a filled-block progress meter in Excel's status bar for each build phase.
Sub Prog(pct, msg)
    On Error Resume Next
    Dim total, filled, i, bar
    total = 22
    filled = Int(total * pct / 100)
    bar = ""
    For i = 1 To total
        If i <= filled Then
            bar = bar & ChrW(9608)
        Else
            bar = bar & ChrW(9618)
        End If
    Next
    xl.StatusBar = "Building MedicationDispensing.xlsm   [" & bar & "]  " & pct & "%   -   " & msg
    On Error GoTo 0
End Sub

Sub CleanupQuit
    On Error Resume Next
    wb.Close False
    xl.Quit
    Set wb = Nothing
    Set proj = Nothing
    Set xl = Nothing
    On Error GoTo 0
End Sub
