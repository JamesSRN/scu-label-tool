Attribute VB_Name = "MedParser"
' ============================================================
'  Saturday Clinic for the Uninsured
'  Medication Parser  -  MedParser.bas  -  Version 1.0
'
'  INSTALL:  Developer -> Visual Basic -> File -> Import File
'  Save workbook as .xlsm after importing.
'
'  Keyboard shortcuts (registered by SetupWorkbook):
'    Ctrl+Shift+P  - Parse Medications
'    Ctrl+Shift+R  - Reset Session
'    Ctrl+Shift+L  - Update Label Preview from selected row
' ============================================================
Option Explicit

' -- Sheet names ---------------------------------------------
Private Const SH_INPUT  As String = "Patient & Input"
Private Const SH_MEDS   As String = "Medications"
Private Const SH_LABEL  As String = "Label Preview"
Private Const SH_LOG    As String = "Log"
Private Const SH_ALL    As String = "Label Previews"

' -- Medications sheet column indices ------------------------
Private Const C_NUM     As Integer = 1
Private Const C_NAME    As Integer = 2
Private Const C_STR     As Integer = 3
Private Const C_FORM    As Integer = 4
Private Const C_SIG     As Integer = 5
Private Const C_QTY     As Integer = 6
Private Const C_REF     As Integer = 7
Private Const C_EXP     As Integer = 8
Private Const C_LOT     As Integer = 9
Private Const C_DATE    As Integer = 10
Private Const C_CONF    As Integer = 11
Private Const C_WARN    As Integer = 12
Private Const C_RAW     As Integer = 13
Private Const C_PRTD    As Integer = 14
Private Const C_CNT     As Integer = 15
Private Const C_SEL     As Integer = 16

Private Const MEDS_HDR_ROWS As Integer = 3   ' rows before data begins
Private Const LOG_HDR_ROWS  As Integer = 2

' -- Colours (hex, no #) -------------------------------------
Private Const CLR_HIGH   As Long = 12780748   ' &HC8E6C  pastel green
Private Const CLR_MED    As Long = 16775620   ' &HFFF9C4 pastel yellow
Private Const CLR_LOW    As Long = 16764106   ' &HFFCDD2 pastel red
Private Const CLR_MANUAL As Long = 16773600   ' &HFFF3E0 orange-tint
Private Const CLR_WARN   As Long = 16775936   ' &HFFF800 warn yellow
Private Const CLR_WHITE  As Long = 16777215

' ============================================================
'  DATA TYPE
' ============================================================
Private Type MedRecord
    MedName    As String
    Strength   As String
    DosageForm As String
    SIG        As String
    Quantity   As String
    Refills    As String
    Expiration As String   ' entered manually; pre-filled if found
    LotNumber  As String   ' entered manually
    Confidence As String   ' "High" | "Medium" | "Low"
    Warnings   As String
    RawText    As String
End Type

' ============================================================
'  WORKBOOK SETUP  (run once after importing this module)
' ============================================================
Public Sub SetupWorkbook()
    Application.OnKey "^+P", "ParseMedications"
    Application.OnKey "^+R", "ResetSession"
    Application.OnKey "^+L", "PreviewAllLabels"

    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Set ws1 = ThisWorkbook.Sheets(SH_INPUT)
    Set ws2 = ThisWorkbook.Sheets(SH_MEDS)
    Set ws3 = ThisWorkbook.Sheets(SH_LABEL)
    ws3.Visible = xlSheetVisible   ' show while we rebuild it; re-hidden at the end

    Call AddButtonToSheet(ws1, "btnParse",  "PARSE MEDICATIONS",  "ParseMedications",   48, 2, 240, 26, RGB(21, 101, 192))
    Call AddButtonToSheet(ws1, "btnClear",  "Clear Paste Area",    "ClearPasteArea",     50, 2, 150, 20, RGB(84, 110, 122))
    Call AddButtonToSheet(ws1, "btnReset",  "Reset Session",       "ResetSession",       50, 3, 150, 20, RGB(191, 54, 12))
    Call AddButtonToSheet(ws1, "btnNewPt",  "Start NEW Patient",   "StartNewPatient",    53, 2, 200, 24, RGB(0, 121, 107))
    Call BuildLabelPreviewLayout(ws3)
    Call AddButtonToSheet(ws3, "btnUpd",    "Update Label Preview", "UpdateLabelPreviewFromSelection", 12, 1, 220, 24, RGB(21, 101, 192))
    Call AddButtonToSheet(ws3, "btnPrint",  "Print This Label",    "PrintLabel",         12, 6, 220, 24, RGB(46, 125, 50))
    Call AddButtonToSheet(ws2, "btnAddMed", "+ Add Medication",   "AddMedicationRow",         1, 17, 150, 22, RGB(46, 125, 50))
    Call AddButtonToSheet(ws2, "btnRemMed", "- Remove Selected",  "RemoveSelectedMedication", 3, 17, 150, 22, RGB(191, 54, 12))
    Call AddButtonToSheet(ws2, "btnRevMed", "Review & Validate",  "ReviewMedications",        5, 17, 150, 22, RGB(21, 101, 192))
    Call AddButtonToSheet(ws2, "btnPrvAll", "Preview ALL Labels",   "PreviewAllLabels",         7, 17, 150, 22, RGB(0, 121, 107))
    Call AddButtonToSheet(ws2, "btnPrnChk", "Print Checked Labels", "PrintCheckedLabels",       9, 17, 150, 22, RGB(216, 67, 21))
    ' Removed the single "Print Selected Label" button - printing is now via Print Checked Labels
    On Error Resume Next
    ws2.Shapes("btnPrnMed").Delete
    On Error GoTo 0
    ' Print Count + Print? selection column headers
    ws2.Cells(2, C_CNT).Value = "# of Prints"
    ws2.Cells(2, C_SEL).Value = "Print?"
    ws2.Cells(2, C_PRTD).Copy
    ws2.Cells(2, C_CNT).PasteSpecial xlPasteFormats
    ws2.Cells(2, C_SEL).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws2.Columns(C_CNT).ColumnWidth = 7
    ws2.Columns(C_SEL).ColumnWidth = 8
    ' Worksheet event handlers are PREINSTALLED in the sheet code modules now
    ' (no runtime VBProject modification - more robust for production).
    ' One-time paste documented in HANDOFF section 6.
    ' (was: Call InstallMedSheetEvents(ws2))
    Call ApplyAllRowStates(ws2)

    ' Extend the dispense Log header with Dosage Form + Print #
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)
    wsLog.Cells(2, 13).Value = "Dosage Form"
    wsLog.Cells(2, 14).Value = "Print #"
    wsLog.Cells(2, 12).Copy
    wsLog.Range(wsLog.Cells(2, 13), wsLog.Cells(2, 14)).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    ' Default Date of Rx to today
    If Trim(ws1.Range("C7").Value) = "" Then
        ws1.Range("C7").Value = Format(Now(), "MM/DD/YYYY")
    End If

    ' Consolidate previews: migrate/rename the gallery and hide the internal print sheet
    Dim wsGallery As Worksheet
    Set wsGallery = EnsureAllLabelsSheet()
    ws1.Activate
    On Error Resume Next
    ws3.Visible = xlSheetHidden
    On Error GoTo 0

    MsgBox "Setup complete!" & vbCrLf & _
           "Keyboard shortcuts registered:" & vbCrLf & _
           "  Ctrl+Shift+P  =  Parse Medications" & vbCrLf & _
           "  Ctrl+Shift+R  =  Reset Session" & vbCrLf & _
           "  Ctrl+Shift+L  =  Refresh Label Previews", _
           vbInformation, "Saturday Clinic - Setup Complete"
End Sub

' ============================================================
'  LABEL PREVIEW LAYOUT  (portrait 62 x 100 mm, code-defined)
'  Rebuilt by SetupWorkbook so the layout lives in this module.
' ============================================================
Private Sub BuildLabelPreviewLayout(ws As Worksheet)
    On Error Resume Next
    Application.ScreenUpdating = False

    ws.Range("A1:H20").UnMerge

    ' LANDSCAPE: content spans A:H (~3.8" wide), label is 100mm wide x 62mm tall
    ws.Columns("A:H").ColumnWidth = 5.8

    ' Wide LANDSCAPE stack (points) - total ~2.42"
    ws.Rows(1).RowHeight = 4
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 3
    ws.Rows(4).RowHeight = 22
    ws.Rows(5).RowHeight = 17
    ws.Rows(6).RowHeight = 17
    ws.Rows(7).RowHeight = 22
    ws.Rows(8).RowHeight = 30
    ws.Rows(9).RowHeight = 20
    ws.Rows(10).RowHeight = 5
    ws.Rows(11).RowHeight = 10
    ws.Rows(12).RowHeight = 30
    ws.Rows(13).RowHeight = 8
    ws.Rows(14).RowHeight = 16
    ws.Rows(15).RowHeight = 30

    Dim r As Variant
    For Each r In Array(2, 4, 5, 6, 7, 8, 9)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Merge
    Next r
    ws.Range("A14:H14").Merge
    ws.Range("A15:H15").Merge

    ws.Cells(2, 1).Value = "Saturday Clinic for the Uninsured" & Chr(10) & _
                           "1121 E. North Ave, Milwaukee WI" & Chr(10) & "(414) 588-2865"
    ws.Cells(4, 1).Formula = "=IF('Patient & Input'!C5<>"""",'Patient & Input'!C5,""[Patient Name]"")"
    ws.Cells(5, 1).Formula = "=IF('Patient & Input'!C6<>"""",""DOB: ""&'Patient & Input'!C6,""DOB: [date]"")"
    ws.Cells(6, 1).Formula = "=IF('Patient & Input'!C7<>"""",""Date of Rx: ""&TEXT('Patient & Input'!C7,""mm/dd/yyyy""),""Date of Rx: [date]"")"
    If Trim(ws.Cells(7, 1).Value) = "" Then _
        ws.Cells(7, 1).Value = "[Select a medication row, then click 'Update Label Preview']"
    ws.Cells(14, 1).Value = "<- Select a medication row, then click 'Update Label Preview'"
    ws.Cells(15, 1).Value = "Print: File > Print > Brother QL-1100c > Label: DK-1202 62 x 100 mm, Landscape"

    Call FmtLbl(ws.Cells(2, 1), 8.5, True, "C", "C")
    Call FmtLbl(ws.Cells(4, 1), 13, True, "L", "C")
    Call FmtLbl(ws.Cells(5, 1), 10.5, False, "L", "C")
    Call FmtLbl(ws.Cells(6, 1), 10.5, False, "L", "C")
    Call FmtLbl(ws.Cells(7, 1), 11.5, True, "L", "C")
    Call FmtLbl(ws.Cells(8, 1), 10.5, False, "L", "T")
    Call FmtLbl(ws.Cells(9, 1), 10.5, True, "L", "C")
    Call FmtLbl(ws.Cells(14, 1), 8.5, False, "C", "C")
    Call FmtLbl(ws.Cells(15, 1), 8.5, False, "L", "C")
    ws.Cells(14, 1).Font.Color = RGB(150, 150, 150)

    ' Clear any leftover inner borders, then draw a clean outer box
    ws.Range(ws.Cells(1, 1), ws.Cells(16, 8)).Borders.LineStyle = xlNone
    With ws.Range("A1:H10")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(136, 136, 136)
        .Borders(xlEdgeRight).Color = RGB(136, 136, 136)
        .Borders(xlEdgeTop).Color = RGB(136, 136, 136)
        .Borders(xlEdgeBottom).Color = RGB(136, 136, 136)
    End With
    ' Clean full-width rule under the clinic name (replaces the old gappy line)
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, 8)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With

    ' Clinic logo - TEMPORARILY DISABLED.
    ' Shapes.AddPicture was aborting SetupWorkbook with "unable to import file"
    ' when the PNG could not be read (OneDrive cloud-only file / oversized image).
    ' To re-enable: use a small, locally-present PNG and wrap AddPicture in
    ' On Error Resume Next so a bad logo never breaks setup. See HANDOFF section 4.
    On Error Resume Next
    ws.Shapes("scuLogo").Delete
    On Error GoTo 0

    With ws.PageSetup
        .PrintArea = ws.Range("A1:H10").Address
        .LeftMargin = Application.InchesToPoints(0.04)
        .RightMargin = Application.InchesToPoints(0.04)
        .TopMargin = Application.InchesToPoints(0.04)
        .BottomMargin = Application.InchesToPoints(0.04)
        .HeaderMargin = 0
        .FooterMargin = 0
        .FitToPagesWide = False
        .FitToPagesTall = False
        .Zoom = 100
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .CenterVertically = False
    End With

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub FmtLbl(rng As Range, sz As Single, bld As Boolean, h As String, v As String)
    With rng.Font
        .Name = "Arial"
        .Size = sz
        .Bold = bld
        .Color = RGB(0, 0, 0)
    End With
    Select Case h
        Case "C": rng.HorizontalAlignment = xlCenter
        Case "R": rng.HorizontalAlignment = xlRight
        Case Else: rng.HorizontalAlignment = xlLeft
    End Select
    Select Case v
        Case "T": rng.VerticalAlignment = xlTop
        Case "B": rng.VerticalAlignment = xlBottom
        Case Else: rng.VerticalAlignment = xlCenter
    End Select
    rng.WrapText = True
End Sub

' ============================================================
'  PREVIEW ALL LABELS  (generated gallery + per-label actions)
' ============================================================
Public Sub PreviewAllLabels()
    Call BuildAllLabelsPreview
End Sub

Private Function EnsureAllLabelsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_ALL)
    On Error GoTo 0
    If ws Is Nothing Then
        ' Migrate the old "All Labels" tab to the new name, or create it
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("All Labels")
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Name = SH_ALL
        Else
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(SH_LABEL))
            ws.Name = SH_ALL
        End If
    End If
    ' (was: Call InstallAutoRefresh(ws) - now a preinstalled Worksheet_Activate handler)
    Set EnsureAllLabelsSheet = ws
End Function

Private Sub InstallAutoRefresh(ws As Worksheet)
    ' Inject a Worksheet_Activate handler so the gallery rebuilds whenever the
    ' user clicks onto the All Labels tab. Needs "Trust access to the VBA
    ' project object model" (Trust Center). Fails silently if not allowed.
    On Error Resume Next
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule
    If cm Is Nothing Then Exit Sub
    Dim existing As String
    existing = ""
    If cm.CountOfLines > 0 Then existing = cm.Lines(1, cm.CountOfLines)
    If InStr(existing, "Worksheet_Activate") = 0 Then
        cm.AddFromString _
            "Private Sub Worksheet_Activate()" & vbCrLf & _
            "    On Error Resume Next" & vbCrLf & _
            "    Application.EnableEvents = False" & vbCrLf & _
            "    PreviewAllLabels" & vbCrLf & _
            "    Application.EnableEvents = True" & vbCrLf & _
            "    On Error GoTo 0" & vbCrLf & _
            "End Sub"
    End If
    On Error GoTo 0
End Sub

Private Sub BuildAllLabelsPreview()
    Dim wsM As Worksheet, wsI As Worksheet, ws As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    Set ws = EnsureAllLabelsSheet()

    Application.ScreenUpdating = False

    ' Clear previous tiles + our dynamic buttons
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 3) = "al_" Then shp.Delete
    Next shp
    ws.Cells.Clear
    ws.Rows.RowHeight = 15

    ws.Columns("A:F").ColumnWidth = 9
    ws.Columns("G").ColumnWidth = 2
    ws.Columns("H:J").ColumnWidth = 16

    Dim patName As String, dob As String
    patName = Trim(wsI.Range("C5").Value)
    dob = Trim(wsI.Range("C6").Value)

    ws.Cells(1, 1).Value = "ALL MEDICATION LABELS" & IIf(patName <> "", "   -   " & patName, "")
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Name = "Arial"

    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, C_NAME).End(xlUp).Row

    Dim n As Integer, base As Long, r As Long
    n = 0
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(wsM.Cells(r, C_NAME).Value) = "" Then GoTo NextR
        n = n + 1
        base = 3 + (n - 1) * 9

        Dim medName As String, strength As String, qty As String
        Dim sig As String, expv As String, lotv As String, dateRx As String, warn As String
        medName = Trim(wsM.Cells(r, C_NAME).Value)
        strength = Trim(wsM.Cells(r, C_STR).Value)
        qty = Trim(wsM.Cells(r, C_QTY).Value)
        sig = Trim(wsM.Cells(r, C_SIG).Value)
        expv = Trim(wsM.Cells(r, C_EXP).Value)
        lotv = Trim(wsM.Cells(r, C_LOT).Value)
        dateRx = Trim(wsM.Cells(r, C_DATE).Value)
        warn = Trim(wsM.Cells(r, C_WARN).Value)

        Dim medLine As String
        medLine = medName
        If strength <> "" Then medLine = medLine & " " & strength
        If qty <> "" Then medLine = medLine & "   (QTY: " & qty & ")"

        ws.Range(ws.Cells(base, 1), ws.Cells(base, 6)).Merge
        ws.Cells(base, 1).Value = "Saturday Clinic for the Uninsured  -  1121 E. North Ave, Milwaukee WI"
        Call FmtLbl(ws.Cells(base, 1), 7.5, False, "C", "C")

        ws.Range(ws.Cells(base + 1, 1), ws.Cells(base + 1, 6)).Merge
        ws.Cells(base + 1, 1).Value = IIf(patName <> "", patName, "[Patient Name]")
        Call FmtLbl(ws.Cells(base + 1, 1), 13, True, "L", "C")

        ws.Range(ws.Cells(base + 2, 1), ws.Cells(base + 2, 6)).Merge
        ws.Cells(base + 2, 1).Value = "DOB: " & IIf(dob <> "", dob, "________") & _
                                      "      Date of Rx: " & IIf(dateRx <> "", dateRx, "________")
        Call FmtLbl(ws.Cells(base + 2, 1), 9.5, False, "L", "C")

        ws.Range(ws.Cells(base + 3, 1), ws.Cells(base + 3, 6)).Merge
        ws.Cells(base + 3, 1).Value = medLine
        Call FmtLbl(ws.Cells(base + 3, 1), 11, True, "L", "C")

        ws.Range(ws.Cells(base + 4, 1), ws.Cells(base + 5, 6)).Merge
        ws.Cells(base + 4, 1).Value = IIf(sig <> "", sig, "[Instructions not found - enter manually]")
        Call FmtLbl(ws.Cells(base + 4, 1), 9.5, False, "L", "T")

        ws.Range(ws.Cells(base + 6, 1), ws.Cells(base + 6, 6)).Merge
        ws.Cells(base + 6, 1).Value = "Exp: " & IIf(expv <> "", expv, "______") & _
                                      "       Lot: " & IIf(lotv <> "", lotv, "______")
        Call FmtLbl(ws.Cells(base + 6, 1), 9.5, True, "L", "C")

        ws.Rows(base).RowHeight = 16
        ws.Rows(base + 1).RowHeight = 22
        ws.Rows(base + 2).RowHeight = 16
        ws.Rows(base + 3).RowHeight = 18
        ws.Rows(base + 4).RowHeight = 16
        ws.Rows(base + 5).RowHeight = 16
        ws.Rows(base + 6).RowHeight = 18

        With ws.Range(ws.Cells(base, 1), ws.Cells(base + 6, 6))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            If IsRowSelected(wsM, r) Then .Interior.Color = RGB(212, 237, 218)
        End With

        If warn <> "" And UCase(warn) <> "OK" Then
            ws.Cells(base + 7, 1).Value = "! " & warn
            Call FmtLbl(ws.Cells(base + 7, 1), 8, False, "L", "C")
            ws.Cells(base + 7, 1).Font.Color = RGB(191, 54, 12)
        End If

        Call AddRowButton(ws, "al_print_" & r, "Print this label", "RowPrint", base, RGB(123, 31, 162))
        Call AddRowButton(ws, "al_edit_" & r, "Edit this med", "RowEdit", base + 2, RGB(21, 101, 192))
        Call AddRowButton(ws, "al_remove_" & r, "Remove this med", "RowRemove", base + 4, RGB(191, 54, 12))
NextR:
    Next r

    If n = 0 Then
        ws.Cells(3, 1).Value = "No medications to preview yet. Parse or add medications first."
    End If

    ws.Activate
    ws.Cells(1, 1).Select
    Application.ScreenUpdating = True
End Sub

Private Sub AddRowButton(ws As Worksheet, nm As String, caption As String, macro As String, atRow As Long, bg As Long)
    Dim leftPos As Double, topPos As Double
    leftPos = ws.Columns("H").Left + 2
    topPos = ws.Rows(atRow).Top + 1
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, 150, 20)
    shp.Name = nm
    shp.OnAction = macro
    shp.Fill.ForeColor.RGB = bg
    shp.Line.Visible = msoFalse
    With shp.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Function CallerRow() As Long
    On Error Resume Next
    Dim nm As String, pos As Integer
    nm = CStr(Application.Caller)
    pos = InStrRev(nm, "_")
    If pos > 0 Then CallerRow = CLng(Mid(nm, pos + 1))
    On Error GoTo 0
End Function

Public Sub RowPrint()
    Dim r As Long
    r = CallerRow()
    If r <= MEDS_HDR_ROWS Then Exit Sub
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    wsM.Activate
    wsM.Cells(r, C_NAME).Select
    Call PrintLabel
    On Error Resume Next
    ThisWorkbook.Sheets(SH_ALL).Activate
    On Error GoTo 0
End Sub

Public Sub RowRemove()
    Dim r As Long
    r = CallerRow()
    If r <= MEDS_HDR_ROWS Then Exit Sub
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    wsM.Activate
    wsM.Cells(r, C_NAME).Select
    Call RemoveSelectedMedication
    Call BuildAllLabelsPreview
End Sub

Public Sub RowEdit()
    Dim r As Long
    r = CallerRow()
    If r <= MEDS_HDR_ROWS Then Exit Sub
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)

    Dim t As String
    t = InputBox("Medication NAME:", "Edit medication", CStr(wsM.Cells(r, C_NAME).Value))
    If StrPtr(t) <> 0 Then wsM.Cells(r, C_NAME).Value = Trim(t)
    t = InputBox("STRENGTH (e.g. 10 mg):", "Edit medication", CStr(wsM.Cells(r, C_STR).Value))
    If StrPtr(t) <> 0 Then wsM.Cells(r, C_STR).Value = Trim(t)
    t = InputBox("INSTRUCTIONS (SIG):", "Edit medication", CStr(wsM.Cells(r, C_SIG).Value))
    If StrPtr(t) <> 0 Then wsM.Cells(r, C_SIG).Value = Trim(t)
    t = InputBox("QUANTITY:", "Edit medication", CStr(wsM.Cells(r, C_QTY).Value))
    If StrPtr(t) <> 0 Then wsM.Cells(r, C_QTY).Value = Trim(t)

    Call ValidateMedications(False)
    Call BuildAllLabelsPreview
End Sub

Private Sub AddButtonToSheet(ws As Worksheet, btnName As String, _
                              caption As String, macro As String, _
                              dataRow As Long, dataCol As Long, _
                              w As Long, h As Long, bgColor As Long)
    On Error Resume Next
    ws.Shapes(btnName).Delete
    On Error GoTo 0

    Dim topPos As Double, leftPos As Double
    topPos  = ws.Rows(dataRow).Top + 2
    leftPos = ws.Columns(dataCol).Left + 2

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, w, h)
    With shp
        .Name = btnName
        .OnAction = macro
        .Fill.ForeColor.RGB = bgColor
        .Line.Visible = msoFalse
        With .TextFrame2
            .TextRange.Text = caption
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

' ============================================================
'  PUBLIC ENTRY POINTS
' ============================================================

Public Sub ParseMedications()
    Dim wsIn  As Worksheet
    Dim wsMed As Worksheet
    Set wsIn  = ThisWorkbook.Sheets(SH_INPUT)
    Set wsMed = ThisWorkbook.Sheets(SH_MEDS)

    ' -- Get raw pasted text ----------------------------------
    Dim rawText As String
    rawText = Trim(wsIn.Range("C12").Value)   ' merged paste cell starts B12; VBA reads first cell of merge

    If rawText = "" Then
        ' Try the legacy position B12 (in case sheet was re-saved differently)
        rawText = Trim(wsIn.Range("B12").Value)
    End If

    If rawText = "" Then
        MsgBox "Please paste medication text in the box on the Patient & Input tab first.", _
               vbExclamation, "No Medication Text"
        Exit Sub
    End If

    ' -- Validate patient info ---------------------------------
    Dim patName As String, dob As String, dateRx As String
    patName = Trim(wsIn.Range("C5").Value)
    dob     = Trim(wsIn.Range("C6").Value)
    dateRx  = Trim(wsIn.Range("C7").Value)

    If patName = "" Then
        patName = InputBox("Patient name is missing." & vbCrLf & _
                           "Enter patient name (from the Teams chat):", _
                           "Patient Name Required", "")
        If Trim(patName) = "" Then
            MsgBox "Patient name is required.", vbExclamation, "Required"
            Exit Sub
        End If
        wsIn.Range("C5").Value = patName
    End If

    If dob = "" Then
        dob = InputBox("Date of Birth is missing." & vbCrLf & _
                       "Enter DOB (MM/DD/YYYY):", "DOB Required", "")
        If Trim(dob) = "" Then
            MsgBox "Date of Birth is required.", vbExclamation, "Required"
            Exit Sub
        End If
        wsIn.Range("C6").Value = dob
    End If

    If dateRx = "" Then
        dateRx = Format(Now(), "MM/DD/YYYY")
        wsIn.Range("C7").Value = dateRx
    End If

    ' -- Split into blocks -------------------------------------
    Dim blocks() As String
    blocks = SplitMedBlocks(rawText)

    If UBound(blocks) < 0 Then
        MsgBox "Could not identify any medication blocks in the pasted text.", _
               vbExclamation, "Parse Error"
        Exit Sub
    End If

    ' -- Find first empty data row -----------------------------
    Dim nextRow As Long
    nextRow = FirstEmptyRow(wsMed)

    ' -- Parse each block --------------------------------------
    Application.ScreenUpdating = False
    Dim added As Integer
    added = 0

    Dim i As Integer
    For i = 0 To UBound(blocks)
        Dim blk As String
        blk = Trim(blocks(i))
        If Len(blk) > 5 Then
            Dim rec As MedRecord
            rec = ParseOneBlock(blk)
            If rec.MedName <> "" Then
                rec = PromptMissingRequired(rec)
                WriteMedRow wsMed, nextRow + added, rec, patName, dob, dateRx, added + 1
                added = added + 1
            End If
        End If
    Next i

    Application.ScreenUpdating = True

    If added = 0 Then
        MsgBox "Could not extract any medications. Please check the pasted text.", _
               vbExclamation, "Parse Failed"
    Else
        ThisWorkbook.Sheets(SH_MEDS).Activate

        ' Offer to collect Expiration + Lot now, or let the volunteer fill them in later
        Dim doExpLot As Integer
        doExpLot = MsgBox(added & " medication(s) parsed." & vbCrLf & vbCrLf & _
                  "Enter EXPIRATION and LOT for each one now?" & vbCrLf & _
                  "(No = fill the orange cells on the sheet later;" & vbCrLf & _
                  " those rows stay flagged until they are filled.)", _
                  vbYesNo + vbQuestion, "Enter Exp / Lot?")
        If doExpLot = vbYes Then
            Dim r As Long
            For r = nextRow To nextRow + added - 1
                If Trim(wsMed.Cells(r, C_NAME).Value) <> "" Then
                    Dim medLabel As String
                    medLabel = Trim(wsMed.Cells(r, C_NAME).Value & " " & wsMed.Cells(r, C_STR).Value)
                    wsMed.Cells(r, C_EXP).NumberFormat = "@"
                    wsMed.Cells(r, C_LOT).NumberFormat = "@"
                    If Trim(wsMed.Cells(r, C_EXP).Value) = "" Then
                        wsMed.Cells(r, C_EXP).Value = Trim(InputBox( _
                            "Enter EXPIRATION DATE for:" & vbCrLf & medLabel & vbCrLf & _
                            "(format: MM/YYYY  -  check the bottle)", _
                            "Expiration Required", ""))
                    End If
                    If Trim(wsMed.Cells(r, C_LOT).Value) = "" Then
                        wsMed.Cells(r, C_LOT).Value = Trim(InputBox( _
                            "Enter LOT NUMBER for:" & vbCrLf & medLabel & vbCrLf & _
                            "(check the bottle or package)", _
                            "Lot Number Required", ""))
                    End If
                End If
            Next r
        End If

        ' Validate + show the full review of every medication
        Call ReviewMedications
    End If
End Sub

Public Sub ClearPasteArea()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_INPUT)
    ws.Range("B12").Value = ""
    ws.Range("C12").Value = ""
    ws.Activate
    ws.Range("B12").Select
End Sub

Public Sub ResetSession()
    If MsgBox("FULL RESET: clear ALL patient data, the medication list, AND the" & vbCrLf & _
              "entire dispense Log, and start completely fresh?" & vbCrLf & vbCrLf & _
              "(This cannot be undone. To keep the Log, use 'Start NEW Patient' instead.)", _
              vbYesNo + vbExclamation, "Reset Session") = vbNo Then
        Exit Sub
    End If

    Dim wsIn  As Worksheet
    Dim wsMed As Worksheet
    Set wsIn  = ThisWorkbook.Sheets(SH_INPUT)
    Set wsMed = ThisWorkbook.Sheets(SH_MEDS)

    ' Clear patient fields
    wsIn.Range("C5").Value = ""
    wsIn.Range("C6").Value = ""
    wsIn.Range("C7").Value = Format(Now(), "MM/DD/YYYY")
    wsIn.Range("B12").Value = ""
    wsIn.Range("C12").Value = ""

    ' Clear medication rows
    Dim lastRow As Long
    lastRow = wsMed.Cells(wsMed.Rows.Count, C_NAME).End(xlUp).Row
    If lastRow > MEDS_HDR_ROWS Then
        wsMed.Rows(MEDS_HDR_ROWS + 1 & ":" & lastRow).ClearContents
        wsMed.Range(wsMed.Cells(MEDS_HDR_ROWS + 1, 1), wsMed.Cells(lastRow, C_SEL)).Interior.ColorIndex = xlNone
        Dim r As Long
        For r = MEDS_HDR_ROWS + 1 To lastRow
            wsMed.Cells(r, C_EXP).NumberFormat = "@"
            wsMed.Cells(r, C_LOT).NumberFormat = "@"
        Next r
    End If

    ' Clear the entire dispense Log (full reset only; Start NEW Patient keeps it)
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)
    Dim lastLog As Long
    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastLog > LOG_HDR_ROWS Then
        With wsLog.Range(wsLog.Cells(LOG_HDR_ROWS + 1, 1), wsLog.Cells(lastLog, 14))
            .ClearContents
            .Interior.ColorIndex = xlNone
        End With
    End If

    ' Clear label preview dynamic cells
    Dim wsLbl As Worksheet
    Set wsLbl = ThisWorkbook.Sheets(SH_LABEL)
    wsLbl.Cells(7, 1).Value = "[Select a row in Medications tab and click Update Label Preview]"
    wsLbl.Cells(8, 1).Value = ""
    wsLbl.Cells(9, 1).Value = ""

    wsIn.Activate
    wsIn.Range("C5").Select
    MsgBox "Full reset complete - patient, medications, and Log cleared." & vbCrLf & _
           "Ready for a fresh session.", vbInformation, "Reset Complete"
End Sub

' ============================================================
'  START NEW PATIENT  (clears current patient + meds, KEEPS the Log)
' ============================================================
Public Sub StartNewPatient()
    If MsgBox("Start a NEW patient?" & vbCrLf & vbCrLf & _
              "This clears the current patient name, DOB, pasted text, and the" & vbCrLf & _
              "medication list so you can enter the next patient." & vbCrLf & vbCrLf & _
              "The dispense LOG is kept (not cleared).", _
              vbYesNo + vbQuestion, "Start New Patient") = vbNo Then Exit Sub

    Dim wsIn As Worksheet, wsMed As Worksheet
    Set wsIn = ThisWorkbook.Sheets(SH_INPUT)
    Set wsMed = ThisWorkbook.Sheets(SH_MEDS)

    wsIn.Range("C5").Value = ""
    wsIn.Range("C6").Value = ""
    wsIn.Range("C7").Value = Format(Now(), "MM/DD/YYYY")
    wsIn.Range("B12").Value = ""
    wsIn.Range("C12").Value = ""

    Dim lastRow As Long, r As Long
    lastRow = wsMed.Cells(wsMed.Rows.Count, C_NAME).End(xlUp).Row
    If lastRow > MEDS_HDR_ROWS Then
        wsMed.Rows(MEDS_HDR_ROWS + 1 & ":" & lastRow).ClearContents
        wsMed.Range(wsMed.Cells(MEDS_HDR_ROWS + 1, 1), wsMed.Cells(lastRow, C_SEL)).Interior.ColorIndex = xlNone
        For r = MEDS_HDR_ROWS + 1 To lastRow
            wsMed.Cells(r, C_EXP).NumberFormat = "@"
            wsMed.Cells(r, C_LOT).NumberFormat = "@"
        Next r
    End If

    wsIn.Activate
    wsIn.Range("C5").Select
    MsgBox "Ready for the next patient." & vbCrLf & "The dispense Log was kept.", _
           vbInformation, "New Patient"
End Sub

' ============================================================
'  BLOCK SPLITTER
'  Strategy:
'    1. Remove "See All" lines
'    2. Detect new-medication boundaries using sentinel lines
'       and drug-name patterns
'    3. Return array of text blocks, one per medication
' ============================================================
Private Function SplitMedBlocks(rawText As String) As String()
    ' Normalise line endings
    rawText = Replace(rawText, vbCrLf, vbLf)
    rawText = Replace(rawText, vbCr, vbLf)

    Dim lines() As String
    lines = Split(rawText, vbLf)

    ' Pass 1: strip noise lines that always signal end-of-block or are junk
    Dim cleaned() As String
    ReDim cleaned(UBound(lines))
    Dim nc As Integer
    nc = 0
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = Trim(lines(i))
        ' Skip "See All" - Tebra UI artifact
        If LCase(ln) = "see all" Then
            cleaned(nc) = "[[SEE_ALL]]"
            nc = nc + 1
        ElseIf IsPharmacyLine(ln) Or IsProviderLine(ln) Then
            cleaned(nc) = "[[SKIP]]"
            nc = nc + 1
        Else
            cleaned(nc) = ln
            nc = nc + 1
        End If
    Next i
    ReDim Preserve cleaned(nc - 1)

    ' Pass 2: identify block boundaries
    ' A new block starts when:
    '   a) we hit [[SEE_ALL]]
    '   b) we see a drug-name line AFTER we've already seen content
    Dim blocks() As String
    ReDim blocks(0)
    Dim blockCount As Integer
    blockCount = 0
    Dim currentBlock As String
    currentBlock = ""
    Dim inBlock As Boolean
    inBlock = False
    Dim linesInBlock As Integer
    linesInBlock = 0

    For i = 0 To UBound(cleaned)
        Dim cl As String
        cl = cleaned(i)

        If cl = "[[SEE_ALL]]" Then
            ' Flush current block
            If Len(Trim(currentBlock)) > 5 Then
                ReDim Preserve blocks(blockCount)
                blocks(blockCount) = Trim(currentBlock)
                blockCount = blockCount + 1
            End If
            currentBlock = ""
            linesInBlock = 0
            inBlock = False

        ElseIf cl = "[[SKIP]]" Then
            ' Ignore pharmacy / provider lines - they don't split blocks

        ElseIf cl = "" Then
            ' Blank line - possible block separator if we've seen enough content
            If linesInBlock >= 2 Then
                ' Look ahead: if next non-blank looks like a new drug name -> split
                Dim j As Integer
                For j = i + 1 To UBound(cleaned)
                    If Trim(cleaned(j)) <> "" And cleaned(j) <> "[[SEE_ALL]]" And _
                       cleaned(j) <> "[[SKIP]]" Then
                        If IsMedHeaderLine(Trim(cleaned(j))) Then
                            ' Split here
                            If Len(Trim(currentBlock)) > 5 Then
                                ReDim Preserve blocks(blockCount)
                                blocks(blockCount) = Trim(currentBlock)
                                blockCount = blockCount + 1
                            End If
                            currentBlock = ""
                            linesInBlock = 0
                            inBlock = False
                        End If
                        Exit For
                    End If
                Next j
            End If

        Else
            ' Content line
            If Not inBlock Then
                ' Starting a new block
                inBlock = True
                linesInBlock = 0
            End If
            ' Start a new medication as soon as a new drug-header line appears
            If linesInBlock >= 1 And IsMedHeaderLine(cl) Then
                If Len(Trim(currentBlock)) > 5 Then
                    ReDim Preserve blocks(blockCount)
                    blocks(blockCount) = Trim(currentBlock)
                    blockCount = blockCount + 1
                End If
                currentBlock = cl & vbLf
                linesInBlock = 1
            Else
                currentBlock = currentBlock & cl & vbLf
                linesInBlock = linesInBlock + 1
            End If
        End If
    Next i

    ' Flush last block
    If Len(Trim(currentBlock)) > 5 Then
        ReDim Preserve blocks(blockCount)
        blocks(blockCount) = Trim(currentBlock)
        blockCount = blockCount + 1
    End If

    If blockCount = 0 Then
        ' Fall back: treat entire text as one block
        ReDim blocks(0)
        blocks(0) = Trim(rawText)
    End If

    SplitMedBlocks = blocks
End Function

' ============================================================
'  SINGLE BLOCK PARSER
' ============================================================
Private Function ParseOneBlock(blk As String) As MedRecord
    Dim rec As MedRecord
    rec.RawText = blk
    rec.Confidence = "High"
    rec.Warnings = ""

    ' Normalise
    blk = Replace(blk, vbCrLf, vbLf)
    blk = Replace(blk, vbCr, vbLf)

    Dim lines() As String
    lines = Split(blk, vbLf)

    ' -- Strip noise lines first ------------------------------
    Dim cleanLines() As String
    ReDim cleanLines(UBound(lines))
    Dim nc As Integer
    nc = 0
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = Trim(lines(i))
        If ln <> "" And LCase(ln) <> "see all" And _
           Not IsPharmacyLine(ln) And Not IsProviderLine(ln) Then
            cleanLines(nc) = ln
            nc = nc + 1
        End If
    Next i
    If nc = 0 Then
        rec.MedName = ""
        ParseOneBlock = rec
        Exit Function
    End If
    ReDim Preserve cleanLines(nc - 1)

    ' -- Find first alpha line (skip garbage like "!!! URGENT !!!") --
    Dim nameIdx As Integer
    nameIdx = 0
    For i = 0 To UBound(cleanLines)
        If Len(cleanLines(i)) > 0 Then
            Dim firstChar As String
            firstChar = Left(cleanLines(i), 1)
            If (firstChar >= "A" And firstChar <= "Z") Or _
               (firstChar >= "a" And firstChar <= "z") Then
                nameIdx = i
                Exit For
            End If
        End If
    Next i

    ' -- Parse name / strength / form from the name line -------
    Dim firstLine As String
    firstLine = cleanLines(nameIdx)

    rec.Strength   = ExtractStrength(firstLine)
    rec.DosageForm = ExtractDosageForm(firstLine)
    rec.MedName    = ExtractMedName(firstLine)

    ' -- Check for inline SIG on the same line as the drug ----
    '    e.g. "metFORMIN 1,000 mg tablet, 1 tab(s) orally 2 times a day"
    Dim sigFound As Boolean
    Dim qtyFound As Boolean
    Dim refFound As Boolean
    sigFound = False
    qtyFound = False
    refFound = False

    If IsSIGLine(firstLine) Then
        Dim inlineSIG As String
        inlineSIG = ExtractInlineSIG(firstLine)
        If inlineSIG <> "" Then
            rec.SIG = NormaliseSIG(inlineSIG)
            sigFound = True
        End If
    End If

    For i = nameIdx + 1 To UBound(cleanLines)
        ln = Trim(cleanLines(i))
        If ln = "" Then GoTo NextLine

        ' Strength (if not already on line 0)
        If rec.Strength = "" And ContainsStrengthPattern(ln) And Not IsSIGLine(ln) Then
            rec.Strength = ExtractStrength(ln)
            If rec.DosageForm = "" Then rec.DosageForm = ExtractDosageForm(ln)
            GoTo NextLine
        End If

        ' SIG line
        If Not sigFound And IsSIGLine(ln) Then
            rec.SIG = NormaliseSIG(ln)
            sigFound = True
            ' Also try to grab form from SIG if not found yet
            If rec.DosageForm = "" Then rec.DosageForm = ExtractDosageForm(ln)
            GoTo NextLine
        End If

        ' Quantity + refills (often on the same line: "30 tablets, 0 refills")
        If Not qtyFound Or Not refFound Then
            Dim qtyTmp As String, refTmp As String
            qtyTmp = ""
            refTmp = ""
            ExtractQtyRefill ln, qtyTmp, refTmp
            If qtyTmp <> "" And Not qtyFound Then
                rec.Quantity = qtyTmp
                qtyFound = True
            End If
            If refTmp <> "" And Not refFound Then
                rec.Refills = refTmp
                refFound = True
            End If
        End If

        ' Standalone refills line
        If Not refFound And IsRefillLine(ln) Then
            rec.Refills = ExtractRefills(ln)
            refFound = True
        End If

NextLine:
    Next i

    ' -- If SIG still not found, check first line more carefully -
    If Not sigFound And IsSIGLine(firstLine) Then
        rec.SIG = NormaliseSIG(firstLine)
        sigFound = True
    End If

    ' -- Dosage form fallback from name ----------------------
    If rec.DosageForm = "" Then
        rec.DosageForm = ExtractDosageForm(rec.MedName)
        rec.MedName = StripFormFromName(rec.MedName)
    End If

    ' -- Clean up medication name -----------------------------
    rec.MedName = NormaliseMedName(rec.MedName)

    ' -- Confidence & warnings --------------------------------
    Dim warns As String
    warns = ""
    Dim misses As Integer
    misses = 0

    If rec.MedName = "" Then
        warns = warns & "Medication name not found. "
        misses = misses + 2
    End If
    If rec.Strength = "" Then
        warns = warns & "Strength not found. "
        misses = misses + 1
    End If
    If rec.SIG = "" Then
        warns = warns & "SIG / instructions not found. "
        misses = misses + 1
    End If
    If rec.Quantity = "" Then
        warns = warns & "Quantity not found. "
        misses = misses + 1
    End If
    If rec.Refills = "" Then
        warns = warns & "Refills not found. "
        misses = misses + 1
    End If

    rec.Warnings = Trim(warns)

    Select Case misses
        Case 0:      rec.Confidence = "High"
        Case 1, 2:   rec.Confidence = "Medium"
        Case Else:   rec.Confidence = "Low"
    End Select

    ParseOneBlock = rec
End Function

' ============================================================
'  FIELD EXTRACTORS
' ============================================================

Private Function ExtractMedName(line As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    Dim result As String
    result = line

    ' Anchor on first strength pattern - everything before it is the name
    re.Pattern = "\d[\d,\.]*(?:-\d[\d\.]*)?\s*" & _
                 "(mg|mcg|g(?![a-zA-Z])|mL(?![a-zA-Z])|units?(?:\s*/\s*mL)?|%|IU|mmol|mcg/actuation|mcg/inh)"
    re.IgnoreCase = True
    re.Global = False

    If re.Test(result) Then
        Dim m As Object
        Set m = re.Execute(result)(0)
        result = Left(result, m.FirstIndex)
    Else
        ' Fallback: cut at first non-digit comma (avoids cutting inside "1,000")
        re.Pattern = "(?:^|[^\d]),"
        re.Global = False
        If re.Test(result) Then
            Set m = re.Execute(result)(0)
            Dim commaPos As Long
            commaPos = m.FirstIndex + m.Length - 1
            If commaPos > 3 Then
                result = Left(result, commaPos)
            End If
        End If
    End If

    ' Remove trailing dosage form words (word-boundary safe)
    re.Pattern = "\b(tablet[s]?|capsule[s]?|solution|suspension|syrup|" & _
                 "cream|ointment|gel|lotion|patch(?:es)?|inhaler|aerosol|nasal\s*spray|" & _
                 "spray|eye\s*drops?|ear\s*drops?|drops?|kwikpen|flexpen|solostar|" & _
                 "autoinjector|pen[s]?|syringe[s]?|vial[s]?|powder|packet|lozenge[s]?|" & _
                 "suppositorie[s]?|enema|injection|injectable|topical)(?![a-zA-Z]).*$"
    re.IgnoreCase = True
    re.Global = False
    result = Trim(re.Replace(result, ""))

    ' Remove trailing parenthetical, commas, hyphens
    re.Pattern = "\s*\(.*$"
    result = Trim(re.Replace(result, ""))
    re.Pattern = "[,\-\s]+$"
    result = Trim(re.Replace(result, ""))

    ExtractMedName = Trim(result)
End Function

Private Function NormaliseMedName(name As String) As String
    If name = "" Then
        NormaliseMedName = name
        Exit Function
    End If

    ' Detect TALL MAN lettering (mixed case in single word like "metFORMIN")
    ' Strategy: if a word has both lower-then-upper transitions mid-word,
    ' convert to Title Case
    Dim words() As String
    words = Split(name, " ")
    Dim i As Integer
    For i = 0 To UBound(words)
        Dim w As String
        w = words(i)
        If HasTallManPattern(w) Then
            ' Convert to Title Case
            words(i) = UCase(Left(w, 1)) & LCase(Mid(w, 2))
        ElseIf w = UCase(w) And Len(w) > 2 Then
            ' All-caps word that isn't an abbreviation - leave as-is (brand)
            words(i) = w
        ElseIf w = LCase(w) And Len(w) > 1 Then
            ' All-lowercase -> Title Case
            words(i) = UCase(Left(w, 1)) & Mid(w, 2)
        End If
        ' Mixed case like "KwikPen" -> leave as-is
    Next i
    NormaliseMedName = Join(words, " ")
End Function

Private Function HasTallManPattern(w As String) As Boolean
    ' Returns True if word looks like TALL MAN lettering
    ' e.g. "metFORMIN" - starts lower, then goes upper mid-word
    If Len(w) < 4 Then
        HasTallManPattern = False
        Exit Function
    End If
    Dim i As Integer
    Dim seenLower As Boolean
    Dim seenUpperAfterLower As Boolean
    seenLower = False
    seenUpperAfterLower = False
    For i = 1 To Len(w)
        Dim ch As String
        ch = Mid(w, i, 1)
        If ch >= "a" And ch <= "z" Then
            seenLower = True
        ElseIf ch >= "A" And ch <= "Z" Then
            If seenLower Then seenUpperAfterLower = True
        End If
    Next i
    HasTallManPattern = seenUpperAfterLower
End Function

Private Function ExtractStrength(line As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    ' Allow hyphenated combos like "5-10 mg" for combination products
    re.Pattern = "\d[\d,\.]*(?:-\d[\d\.]*)?\s*" & _
                 "(mg|mcg|g(?![a-zA-Z])|mL(?![a-zA-Z])|unit[s]?(?:\s*/\s*mL)?|%|IU|mmol|mcg/actuation|mcg/inh)" & _
                 "(?:\s*\(\s*\d[\d\.]*\s*mL\s*\))?"
    re.IgnoreCase = True
    re.Global = False

    If re.Test(line) Then
        ExtractStrength = Trim(re.Execute(line)(0))
    Else
        ExtractStrength = ""
    End If
End Function

Private Function ContainsStrengthPattern(line As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d[\d,\.]*(?:-\d[\d\.]*)?\s*(mg|mcg|g(?![a-zA-Z])|mL(?![a-zA-Z])|unit[s]?|%|IU|mmol)"
    re.IgnoreCase = True
    ContainsStrengthPattern = re.Test(line)
End Function

Private Function ExtractDosageForm(line As String) As String
    ' Use word-boundary pattern to avoid "pen" inside "dispensing", etc.
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False

    ' Ordered: longer/more specific forms first
    Dim forms(33) As String
    Dim canonical(33) As String
    forms(0)  = "kwikpen"         : canonical(0)  = "KwikPen"
    forms(1)  = "flexpen"         : canonical(1)  = "FlexPen"
    forms(2)  = "solostar"        : canonical(2)  = "SoloStar"
    forms(3)  = "autoinjector"    : canonical(3)  = "Autoinjector"
    forms(4)  = "nasal spray"     : canonical(4)  = "Nasal Spray"
    forms(5)  = "eye drops"       : canonical(5)  = "Eye Drops"
    forms(6)  = "ear drops"       : canonical(6)  = "Ear Drops"
    forms(7)  = "tablet"          : canonical(7)  = "Tablet"
    forms(8)  = "capsule"         : canonical(8)  = "Capsule"
    forms(9)  = "solution"        : canonical(9)  = "Solution"
    forms(10) = "suspension"      : canonical(10) = "Suspension"
    forms(11) = "syrup"           : canonical(11) = "Syrup"
    forms(12) = "cream"           : canonical(12) = "Cream"
    forms(13) = "ointment"        : canonical(13) = "Ointment"
    forms(14) = "lotion"          : canonical(14) = "Lotion"
    forms(15) = "aerosol"         : canonical(15) = "Aerosol"
    forms(16) = "inhaler"         : canonical(16) = "Inhaler"
    forms(17) = "injectable"      : canonical(17) = "Injectable"
    forms(18) = "injection"       : canonical(18) = "Injection"
    forms(19) = "syringe"         : canonical(19) = "Syringe"
    forms(20) = "suppository"     : canonical(20) = "Suppository"
    forms(21) = "suppositorie"    : canonical(21) = "Suppository"
    forms(22) = "lozenge"         : canonical(22) = "Lozenge"
    forms(23) = "topical"         : canonical(23) = "Topical"
    forms(24) = "vial"            : canonical(24) = "Vial"
    forms(25) = "powder"          : canonical(25) = "Powder"
    forms(26) = "packet"          : canonical(26) = "Packet"
    forms(27) = "enema"           : canonical(27) = "Enema"
    forms(28) = "patch"           : canonical(28) = "Patch"
    forms(29) = "spray"           : canonical(29) = "Spray"
    forms(30) = "drops"           : canonical(30) = "Drops"
    forms(31) = "drop"            : canonical(31) = "Drops"
    forms(32) = "gel"             : canonical(32) = "Gel"
    forms(33) = "pen"             : canonical(33) = "Pen"

    Dim i As Integer
    For i = 0 To 33
        re.Pattern = "\b" & forms(i) & "(?![a-zA-Z])"
        If re.Test(line) Then
            ExtractDosageForm = canonical(i)
            Exit Function
        End If
    Next i
    ExtractDosageForm = ""
End Function

Private Function ExtractInlineSIG(line As String) As String
    ' Given "metFORMIN 1,000 mg tablet, 1 tab(s) orally twice daily"
    ' returns the SIG portion: "1 tab(s) orally twice daily"
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False

    ' Find strength position
    re.Pattern = "\d[\d,\.]*(?:-\d[\d\.]*)?\s*" & _
                 "(mg|mcg|g(?![a-zA-Z])|mL(?![a-zA-Z])|unit[s]?(?:\s*/\s*mL)?|%|IU|mmol|mcg/actuation)"
    If Not re.Test(line) Then
        ExtractInlineSIG = ""
        Exit Function
    End If

    Dim strMatch As Object
    Set strMatch = re.Execute(line)(0)
    Dim afterStrength As String
    afterStrength = Mid(line, strMatch.FirstIndex + strMatch.Length + 1)

    ' Remove dosage form word at start
    re.Pattern = "^[\s,]*\b(tablet[s]?|capsule[s]?|solution|suspension|" & _
                 "syrup|cream|ointment|gel|lotion|patch(?:es)?|inhaler|aerosol|" & _
                 "spray|drops?|pen[s]?|syringe[s]?|vial[s]?|powder)(?![a-zA-Z])[\s,]*"
    Dim remainder As String
    remainder = Trim(re.Replace(afterStrength, ""))

    ' Strip leading comma/space
    re.Pattern = "^[\s,]+"
    remainder = Trim(re.Replace(remainder, ""))

    If IsSIGLine(remainder) Then
        ExtractInlineSIG = remainder
    Else
        ExtractInlineSIG = ""
    End If
End Function

Private Function StripFormFromName(name As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\s+\b(tablet[s]?|capsule[s]?|solution|suspension|syrup|cream|ointment|" & _
                 "gel|lotion|patch|inhaler|aerosol|spray|drops?|pen|syringe|vial|" & _
                 "powder|packet|lozenge|suppositorie[s]?|enema|injection|injectable|topical)\b.*$"
    re.IgnoreCase = True
    StripFormFromName = Trim(re.Replace(name, ""))
End Function

Private Function IsSIGLine(line As String) As Boolean
    Dim lln As String
    lln = LCase(Trim(line))

    ' Starts with action verb
    Dim verbs As Variant
    verbs = Array("take ", "apply ", "inject ", "instill ", "use ", "inhale ", _
                  "spray ", "place ", "insert ", "swallow ", "dissolve ", "chew ", _
                  "administer ")
    Dim i As Integer
    For i = 0 To UBound(verbs)
        If Left(lln, Len(verbs(i))) = verbs(i) Then
            IsSIGLine = True
            Exit Function
        End If
    Next i

    ' Contains route / frequency
    Dim sigWords As Variant
    sigWords = Array("orally", "by mouth", "subcutaneously", "subcut", "topically", _
                     "vaginally", "rectally", "intramuscularly", "transdermally", _
                     "every morning", "every evening", "every night", "at bedtime", _
                     "once daily", "twice daily", "three times", "four times", _
                     "times a day", "times daily", "times per day", _
                     "as needed", "prn", "as directed", "with food", "with meal", _
                     "tab(s)", "capsule(s)", "unit(s)", "puff(s)", "spray(s)", _
                     "1 tab", "2 tab", "one tablet", "two tablet", "one capsule", _
                     "one unit", "20 unit", "as directed by", _
                     "nightly", "daily", "weekly", "every other day")
    For i = 0 To UBound(sigWords)
        If InStr(lln, sigWords(i)) > 0 Then
            IsSIGLine = True
            Exit Function
        End If
    Next i

    IsSIGLine = False
End Function

Private Function NormaliseSIG(sig As String) As String
    ' Expand common abbreviations
    Dim result As String
    result = sig

    Dim abbrevs(13, 1) As String
    abbrevs(0, 0) = "tab(s)"       : abbrevs(0, 1) = "tablet(s)"
    abbrevs(1, 0) = "cap(s)"       : abbrevs(1, 1) = "capsule(s)"
    abbrevs(2, 0) = "q.d."         : abbrevs(2, 1) = "once daily"
    abbrevs(3, 0) = "b.i.d."       : abbrevs(3, 1) = "twice daily"
    abbrevs(4, 0) = "t.i.d."       : abbrevs(4, 1) = "three times daily"
    abbrevs(5, 0) = "q.i.d."       : abbrevs(5, 1) = "four times daily"
    abbrevs(6, 0) = "p.r.n."       : abbrevs(6, 1) = "as needed"
    abbrevs(7, 0) = "p.o."         : abbrevs(7, 1) = "by mouth"
    abbrevs(8, 0) = "subcut."      : abbrevs(8, 1) = "subcutaneously"
    abbrevs(9, 0) = "subq."        : abbrevs(9, 1) = "subcutaneously"
    abbrevs(10, 0) = "sq."         : abbrevs(10, 1) = "subcutaneously"
    abbrevs(11, 0) = "sq "         : abbrevs(11, 1) = "subcutaneously "
    abbrevs(12, 0) = "2 times a day": abbrevs(12, 1) = "twice daily"
    abbrevs(13, 0) = "3 times a day": abbrevs(13, 1) = "three times daily"

    Dim i As Integer
    For i = 0 To 13
        result = ReplaceCI(result, abbrevs(i, 0), abbrevs(i, 1))
    Next i

    NormaliseSIG = Trim(result)
End Function

Private Sub ExtractQtyRefill(line As String, ByRef qty As String, ByRef ref As String)
    Dim lln As String
    lln = LCase(Trim(line))
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Quantity: "30 tablets" / "10 (3 mL syringe)" / "90 capsules" / "#30"
    re.Pattern = "(?:^|[\s,])(\d+)\s*(tablet[s]?|capsule[s]?|cap[s]?|tab[s]?|" & _
                 "ml|mL|pen[s]?|syringe[s]?|vial[s]?|patch(?:es)?|spray[s]?|" & _
                 "drop[s]?|suppositorie[s]?|lozenge[s]?|unit[s]?|inhaler[s]?)"
    re.IgnoreCase = True
    If re.Test(line) Then
        Dim m As Object
        Set m = re.Execute(line)(0)
        qty = Trim(m.Value)
        ' Canonical: "30 tablets" -> "30"
        re.Pattern = "(\d+)"
        qty = re.Execute(qty)(0).Value
    End If

    ' Refills: "1 refill" / "0 refills" / "3 refills"
    re.Pattern = "(\d+)\s+refill[s]?"
    re.IgnoreCase = True
    If re.Test(line) Then
        ref = re.Execute(line)(0).SubMatches(0)
    End If
End Sub

Private Function IsRefillLine(line As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d+\s+refill[s]?$"
    re.IgnoreCase = True
    IsRefillLine = re.Test(Trim(line))
End Function

Private Function ExtractRefills(line As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d+)\s+refill[s]?"
    re.IgnoreCase = True
    If re.Test(line) Then
        ExtractRefills = re.Execute(line)(0).SubMatches(0)
    Else
        ExtractRefills = ""
    End If
End Function

' -- Line classifiers -----------------------------------------

Private Function IsPharmacyLine(line As String) As Boolean
    Dim lln As String
    lln = LCase(Trim(line))
    Dim pharmas As Variant
    pharmas = Array("walgreens", "cvs", "rite aid", "kroger", "walmart", _
                    "meijer", "ascension rx", "aurora pharmacy", "froedtert", _
                    "pick n save", "target pharmacy", "costco pharmacy", _
                    "express scripts", "mail order", "caremark", "pharmacy")
    Dim i As Integer
    For i = 0 To UBound(pharmas)
        If InStr(lln, pharmas(i)) > 0 Then
            IsPharmacyLine = True
            Exit Function
        End If
    Next i
    IsPharmacyLine = False
End Function

Private Function IsProviderLine(line As String) As Boolean
    Dim lln As String
    lln = LCase(Trim(line))
    ' "Started on MM/DD/YYYY HH:MM PM by Lastname, Firstname"
    If InStr(lln, "started on") > 0 And InStr(lln, " by ") > 0 Then
        IsProviderLine = True
        Exit Function
    End If
    If Left(lln, 10) = "started on" Then
        IsProviderLine = True
        Exit Function
    End If
    If InStr(lln, "entered by") > 0 Or InStr(lln, "prescribed by") > 0 Then
        IsProviderLine = True
        Exit Function
    End If
    IsProviderLine = False
End Function

Private Function IsDateLine(line As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(started on|entered on|prescribed on|\d{1,2}/\d{1,2}/\d{4})"
    re.IgnoreCase = True
    IsDateLine = re.Test(Trim(line))
End Function

Private Function IsDrugNameLine(line As String) As Boolean
    ' A line is likely a drug name if:
    ' 1. It starts with a letter (not a digit)
    ' 2. AND (contains a strength unit OR ends with a form word OR first word is a drug-name pattern)
    Dim lln As String
    lln = Trim(line)
    If Len(lln) = 0 Then
        IsDrugNameLine = False
        Exit Function
    End If
    If Not (Left(lln, 1) >= "A" And Left(lln, 1) <= "Z") And _
       Not (Left(lln, 1) >= "a" And Left(lln, 1) <= "z") Then
        IsDrugNameLine = False
        Exit Function
    End If
    If ContainsStrengthPattern(lln) Then
        IsDrugNameLine = True
        Exit Function
    End If
    If ExtractDosageForm(lln) <> "" Then
        IsDrugNameLine = True
        Exit Function
    End If
    IsDrugNameLine = False
End Function

Private Function StartsWithLetter(s As String) As Boolean
    Dim c As String
    If Len(s) = 0 Then Exit Function
    c = Left(s, 1)
    StartsWithLetter = ((c >= "A" And c <= "Z") Or (c >= "a" And c <= "z"))
End Function

Private Function StartsWithUpper(s As String) As Boolean
    Dim c As String
    If Len(s) = 0 Then Exit Function
    c = Left(s, 1)
    StartsWithUpper = (c >= "A" And c <= "Z")
End Function

Private Function HasDigit(s As String) As Boolean
    Dim i As Integer, c As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c >= "0" And c <= "9" Then
            HasDigit = True
            Exit Function
        End If
    Next i
    HasDigit = False
End Function

Private Function WordCount(s As String) As Integer
    Dim t As String
    t = Trim(s)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    If t = "" Then
        WordCount = 0
    Else
        WordCount = UBound(Split(t, " ")) + 1
    End If
End Function

Private Function StartsWithVerb(line As String) As Boolean
    Dim lln As String
    lln = LCase(Trim(line))
    Dim verbs As Variant
    verbs = Array("take ", "apply ", "inject ", "instill ", "use ", "inhale ", _
                  "spray ", "place ", "insert ", "swallow ", "dissolve ", "chew ", _
                  "administer ", "mix ", "rub ", "shake ", "give ", "put ", _
                  "massage ", "drink ", "swish ", "wash ")
    Dim i As Integer
    For i = 0 To UBound(verbs)
        If Left(lln, Len(verbs(i))) = verbs(i) Then
            StartsWithVerb = True
            Exit Function
        End If
    Next i
    StartsWithVerb = False
End Function

Private Function IsBareNameCandidate(line As String) As Boolean
    ' A short, capitalised, alphabetic line with no strength/SIG -> a bare drug name
    Dim s As String
    s = Trim(line)
    If Len(s) < 3 Then Exit Function
    If Not StartsWithUpper(s) Then Exit Function
    If HasDigit(s) Then Exit Function
    If StartsWithVerb(s) Then Exit Function
    If IsSIGLine(s) Then Exit Function
    If IsDateLine(s) Or IsProviderLine(s) Or IsPharmacyLine(s) Then Exit Function
    If WordCount(s) > 5 Then Exit Function
    IsBareNameCandidate = True
End Function

Private Function IsMedHeaderLine(line As String) As Boolean
    ' TRUE when a line begins a NEW medication (used to split blocks)
    Dim s As String
    s = Trim(line)
    If Len(s) = 0 Then Exit Function
    If Not StartsWithLetter(s) Then Exit Function    ' strength-only lines start with a digit
    If StartsWithVerb(s) Then Exit Function           ' that is a SIG instruction
    If IsDateLine(s) Or IsProviderLine(s) Or IsPharmacyLine(s) Then Exit Function
    ' 1) drug name + strength on the same line (strongest signal)
    If ContainsStrengthPattern(s) Then
        IsMedHeaderLine = True
        Exit Function
    End If
    ' 2) name headed by a dosage form, e.g. "Basaglar KwikPen U-100 Insulin"
    If ExtractDosageForm(s) <> "" And Not IsSIGLine(s) Then
        IsMedHeaderLine = True
        Exit Function
    End If
    ' 3) a bare drug name on its own line, e.g. "Amlodipine"
    If IsBareNameCandidate(s) Then
        IsMedHeaderLine = True
        Exit Function
    End If
    IsMedHeaderLine = False
End Function

' ============================================================
'  MISSING FIELD PROMPTS
' ============================================================
Private Function PromptMissingRequired(rec As MedRecord) As MedRecord
    ' Only prompt for fields that affect patient safety if blank
    If rec.MedName = "" Then
        Dim nm As String
        nm = InputBox("Could not identify medication name from:" & vbCrLf & _
                      Left(rec.RawText, 200) & vbCrLf & vbCrLf & _
                      "Please enter the medication name:", _
                      "Medication Name Needed", "")
        rec.MedName = Trim(nm)
    End If

    If rec.Strength = "" Then
        Dim ans As Integer
        ans = MsgBox("Strength not found for:  " & rec.MedName & vbCrLf & _
                     "Raw text: " & vbCrLf & Left(rec.RawText, 150) & vbCrLf & vbCrLf & _
                     "Enter strength manually?", _
                     vbYesNo + vbQuestion, "Strength Not Found")
        If ans = vbYes Then
            Dim st As String
            st = InputBox("Enter strength (e.g. 500 mg, 10 mg, 100 unit/mL):", _
                          "Enter Strength - " & rec.MedName, "")
            rec.Strength = Trim(st)
        End If
    End If

    If rec.SIG = "" Then
        Dim ans2 As Integer
        ans2 = MsgBox("Instructions (SIG) not found for:  " & rec.MedName & vbCrLf & _
                      "Enter instructions manually?", _
                      vbYesNo + vbQuestion, "SIG Not Found")
        If ans2 = vbYes Then
            Dim sg As String
            sg = InputBox("Enter instructions (e.g. Take one tablet by mouth daily):", _
                          "Enter SIG - " & rec.MedName, "")
            rec.SIG = Trim(sg)
        End If
    End If

    PromptMissingRequired = rec
End Function

' ============================================================
'  WRITE ROW TO MEDICATIONS SHEET
' ============================================================
Private Sub WriteMedRow(ws As Worksheet, r As Long, rec As MedRecord, _
                         patName As String, dob As String, dateRx As String, _
                         rowNum As Integer)
    ws.Cells(r, C_NUM).Value   = rowNum
    ws.Cells(r, C_NAME).Value  = rec.MedName
    ws.Cells(r, C_STR).Value   = rec.Strength
    ws.Cells(r, C_FORM).Value  = rec.DosageForm
    ws.Cells(r, C_SIG).Value   = rec.SIG
    ws.Cells(r, C_QTY).Value   = rec.Quantity
    ws.Cells(r, C_REF).Value   = rec.Refills
    ws.Cells(r, C_EXP).Value   = rec.Expiration
    ws.Cells(r, C_LOT).Value   = rec.LotNumber
    ws.Cells(r, C_DATE).Value  = dateRx
    ws.Cells(r, C_CONF).Value  = rec.Confidence
    ws.Cells(r, C_WARN).Value  = rec.Warnings
    ws.Cells(r, C_RAW).Value   = rec.RawText
    ws.Cells(r, C_PRTD).Value  = "No"

    ' Apply row formatting
    Call ApplyRowState(ws, r)
End Sub

Private Sub ApplyRowState(ws As Worksheet, r As Long)
    If Trim(ws.Cells(r, C_NAME).Value) = "" Then Exit Sub
    Dim warn As String, printed As String, bg As Long
    warn = Trim(ws.Cells(r, C_WARN).Value)
    printed = LCase(Trim(ws.Cells(r, C_PRTD).Value))

    ' Row background state - priority: Selected > Validated > Non-validated
    If IsRowSelected(ws, r) Then
        bg = RGB(212, 237, 218)        ' Selected  - green (matches Label Previews tint)
    ElseIf UCase(warn) = "OK" Then
        bg = RGB(187, 222, 251)        ' Validated  - blue
    Else
        bg = RGB(239, 239, 239)        ' Non-validated - gray
    End If

    Dim c As Integer
    For c = 1 To C_SEL
        With ws.Cells(r, c)
            .Font.Name = "Arial"
            .Font.Size = 10
            If c = C_EXP Or c = C_LOT Then
                .NumberFormat = "@"
                If Trim(.Value) = "" Then
                    .Interior.Color = RGB(255, 205, 210)    ' missing - red
                Else
                    .Interior.Color = bg                     ' filled - blends with row (no orange)
                End If
            ElseIf c = C_CONF Then
                Select Case Trim(ws.Cells(r, C_CONF).Value)
                    Case "High":   .Interior.Color = RGB(200, 230, 201)
                    Case "Medium": .Interior.Color = RGB(255, 249, 196)
                    Case "Low":    .Interior.Color = RGB(255, 205, 210)
                    Case Else:     .Interior.Color = RGB(236, 239, 241)
                End Select
                .Font.Bold = True
            Else
                .Interior.Color = bg
            End If
        End With
    Next c

    With ws.Cells(r, C_SEL)
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(27, 94, 32)
    End With
    ws.Cells(r, C_CNT).HorizontalAlignment = xlCenter

    If warn <> "" And UCase(warn) <> "OK" Then
        ws.Cells(r, C_WARN).Font.Color = RGB(191, 54, 12)
        ws.Cells(r, C_WARN).Font.Bold = True
    Else
        ws.Cells(r, C_WARN).Font.Color = RGB(46, 125, 50)
        ws.Cells(r, C_WARN).Font.Bold = False
    End If
End Sub

Private Function IsRowSelected(ws As Worksheet, r As Long) As Boolean
    IsRowSelected = (Trim(ws.Cells(r, C_SEL).Value) <> "")
End Function

Public Sub ToggleRowSelect(ByVal r As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)
    If r <= MEDS_HDR_ROWS Then Exit Sub
    If Trim(ws.Cells(r, C_NAME).Value) = "" Then Exit Sub
    If Trim(ws.Cells(r, C_SEL).Value) = "" Then
        ws.Cells(r, C_SEL).Value = ChrW(10003)
    Else
        ws.Cells(r, C_SEL).Value = ""
    End If
    Call ApplyRowState(ws, r)
End Sub

Private Sub ApplyAllRowStates(ws As Worksheet)
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" Then Call ApplyRowState(ws, r)
    Next r
End Sub

Private Sub InstallMedSheetEvents(ws As Worksheet)
    ' Inject Worksheet_BeforeDoubleClick so double-clicking the Print? cell
    ' toggles selection. Needs trust access to the VBA project object model.
    On Error Resume Next
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents(ws.CodeName).CodeModule
    If cm Is Nothing Then Exit Sub
    ' Always refresh the handler so it matches the current columns
    Dim startLine As Long, numLines As Long
    startLine = cm.ProcStartLine("Worksheet_BeforeDoubleClick", 0)
    If startLine > 0 Then
        numLines = cm.ProcCountLines("Worksheet_BeforeDoubleClick", 0)
        cm.DeleteLines startLine, numLines
    End If
    cm.AddFromString _
        "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
        "    ' Prints (#) column is auto-managed - block editing it" & vbCrLf & _
        "    If Target.Column = 15 Then" & vbCrLf & _
        "        Cancel = True" & vbCrLf & _
        "        Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    ' Print? column - toggle the selection check" & vbCrLf & _
        "    If Target.Column = 16 And Target.Row > 3 Then" & vbCrLf & _
        "        If Trim(Me.Cells(Target.Row, 2).Value) <> """" Then" & vbCrLf & _
        "            Cancel = True" & vbCrLf & _
        "            ToggleRowSelect Target.Row" & vbCrLf & _
        "        End If" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub"
    On Error GoTo 0
End Sub

Public Sub PrintCheckedLabels()
    Dim ws As Worksheet, wsL As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)
    Set wsL = ThisWorkbook.Sheets(SH_LABEL)

    Dim lastRow As Long, r As Long, cnt As Integer
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    cnt = 0
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" And IsRowSelected(ws, r) Then cnt = cnt + 1
    Next r
    If cnt = 0 Then
        MsgBox "No medications are checked." & vbCrLf & _
               "Double-click the 'Print?' cell next to each med you want to print.", _
               vbExclamation, "Nothing Selected"
        Exit Sub
    End If
    Dim listMsg As String, idx As Integer
    listMsg = ""
    idx = 0
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" And IsRowSelected(ws, r) Then
            idx = idx + 1
            listMsg = listMsg & "   " & idx & ".  " & MedConfirmLine(ws, r)
            If Trim(ws.Cells(r, C_EXP).Value) = "" Or Trim(ws.Cells(r, C_LOT).Value) = "" Then
                listMsg = listMsg & "   [SKIPPED - missing Exp/Lot]"
            End If
            listMsg = listMsg & vbCrLf
        End If
    Next r
    If MsgBox("About to print these " & cnt & " label(s) on the Brother QL-1100c:" & vbCrLf & vbCrLf & _
              listMsg & vbCrLf & _
              "Make sure the DK-1202 (62 x 100 mm) roll is loaded." & vbCrLf & vbCrLf & _
              "YES = print all       NO = cancel", _
              vbYesNo + vbQuestion, "Print Checked Labels") = vbNo Then Exit Sub

    Dim brother As String
    brother = SelectBrotherPrinter()
    If brother = "" Then
        MsgBox "Brother QL-1100c not found. Make sure it is plugged in, powered on," & vbCrLf & _
               "and loaded with the DK-1202 roll, then try again.", _
               vbExclamation, "Printer Not Found"
        Exit Sub
    End If

    With wsL.PageSetup
        .PrintArea = wsL.Range("A1:H10").Address
        .LeftMargin = Application.InchesToPoints(0.04)
        .RightMargin = Application.InchesToPoints(0.04)
        .TopMargin = Application.InchesToPoints(0.04)
        .BottomMargin = Application.InchesToPoints(0.04)
        .HeaderMargin = 0
        .FooterMargin = 0
        .FitToPagesWide = False
        .FitToPagesTall = False
        .Zoom = 100
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .CenterVertically = False
    End With

    Dim batchVol As String
    batchVol = AskInitials()
    Dim done As Integer, skipped As Integer
    done = 0
    skipped = 0
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" And IsRowSelected(ws, r) Then
            If Trim(ws.Cells(r, C_EXP).Value) = "" Or Trim(ws.Cells(r, C_LOT).Value) = "" Then
                skipped = skipped + 1
            Else
                ws.Activate
                ws.Cells(r, C_NAME).Select
                Call UpdateLabelPreviewFromSelection
                On Error Resume Next
                wsL.PrintOut Copies:=1, Collate:=True
                On Error GoTo 0
                Call MarkPrinted(r)
                Call LogPrint(r, batchVol)
                Call ApplyRowState(ws, r)
                done = done + 1
            End If
        End If
    Next r

    Dim msg As String
    msg = done & " label(s) sent to the Brother."
    If skipped > 0 Then msg = msg & vbCrLf & skipped & " skipped (missing Expiration or Lot)."
    MsgBox msg, vbInformation, "Print Complete"
End Sub

Private Function FirstEmptyRow(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    If lastRow <= MEDS_HDR_ROWS Then
        FirstEmptyRow = MEDS_HDR_ROWS + 1
    Else
        FirstEmptyRow = lastRow + 1
    End If
End Function

' ============================================================
'  VALIDATION  /  REVIEW  /  ADD  /  REMOVE
' ============================================================
Private Function IsBadExpFormat(s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    ' Accept MM/YYYY, MM/YY, MM-YYYY, MM-YY
    re.Pattern = "^\s*(0?[1-9]|1[0-2])[/\-](\d{2}|\d{4})\s*$"
    IsBadExpFormat = Not re.Test(s)
End Function

' Scan every medication row, flag problems in the WARNINGS column.
Public Sub ValidateMedications(ByVal showSummary As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    If lastRow <= MEDS_HDR_ROWS Then
        If showSummary Then MsgBox "No medications to validate yet.", vbInformation, "Validate"
        Exit Sub
    End If

    Dim rowsWithIssues As Long
    rowsWithIssues = 0

    Dim r As Long
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) = "" Then GoTo NextR
        Dim w As String
        w = ""
        If Trim(ws.Cells(r, C_STR).Value) = "" Then w = w & "Strength missing. "
        If Trim(ws.Cells(r, C_SIG).Value) = "" Then w = w & "Instructions missing. "
        If Trim(ws.Cells(r, C_QTY).Value) = "" Then w = w & "Quantity missing. "
        If Trim(ws.Cells(r, C_EXP).Value) = "" Then w = w & "EXPIRATION missing. "
        If Trim(ws.Cells(r, C_LOT).Value) = "" Then w = w & "LOT missing. "
        If Trim(ws.Cells(r, C_EXP).Value) <> "" Then
            If IsBadExpFormat(CStr(ws.Cells(r, C_EXP).Value)) Then _
                w = w & "Check expiration format (use MM/YYYY). "
        End If

        ' Duplicate check (same name + strength as an earlier row)
        Dim r2 As Long
        For r2 = MEDS_HDR_ROWS + 1 To r - 1
            If Trim(ws.Cells(r2, C_NAME).Value) <> "" Then
                If LCase(Trim(ws.Cells(r2, C_NAME).Value)) = LCase(Trim(ws.Cells(r, C_NAME).Value)) And _
                   LCase(Trim(ws.Cells(r2, C_STR).Value)) = LCase(Trim(ws.Cells(r, C_STR).Value)) Then
                    w = w & "Possible duplicate of row " & ws.Cells(r2, C_NUM).Value & ". "
                    Exit For
                End If
            End If
        Next r2

        If w = "" Then
            ws.Cells(r, C_WARN).Value = "OK"
        Else
            ws.Cells(r, C_WARN).Value = Trim(w)
            rowsWithIssues = rowsWithIssues + 1
        End If
        Call ApplyRowState(ws, r)
NextR:
    Next r

    If showSummary Then
        If rowsWithIssues = 0 Then
            MsgBox "All medications look complete. No issues found.", _
                   vbInformation, "Validation Passed"
        Else
            MsgBox rowsWithIssues & " medication row(s) need attention." & vbCrLf & _
                   "Look at the orange WARNINGS column and fix those before printing.", _
                   vbExclamation, "Validation - Issues Found"
        End If
    End If
End Sub

' Button wrapper for validation alone
Public Sub RunValidation()
    Call ValidateMedications(True)
End Sub

' Review the whole medication list with a summary + how to add/remove.
Public Sub ReviewMedications()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)
    ws.Activate
    Call ValidateMedications(False)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    If lastRow <= MEDS_HDR_ROWS Then
        MsgBox "No medications on the list yet." & vbCrLf & _
               "Paste medication text and click PARSE, or click 'Add Medication'.", _
               vbInformation, "Medication Review"
        Exit Sub
    End If

    Dim summary As String, n As Integer, flagged As Integer
    summary = "" : n = 0 : flagged = 0
    Dim r As Long
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" Then
            n = n + 1
            Dim mark As String
            If Trim(ws.Cells(r, C_WARN).Value) = "OK" Then
                mark = "   [OK]"
            Else
                mark = "   << " & ws.Cells(r, C_WARN).Value
                flagged = flagged + 1
            End If
            summary = summary & n & ". " & Trim(ws.Cells(r, C_NAME).Value & " " & _
                      ws.Cells(r, C_STR).Value) & mark & vbCrLf
        End If
    Next r

    Dim head As String
    head = n & " medication(s) on the list:" & vbCrLf & vbCrLf
    Dim foot As String
    If flagged = 0 Then
        foot = vbCrLf & "All rows look complete - ready to print."
    Else
        foot = vbCrLf & flagged & " row(s) still need attention (red / orange cells)."
    End If
    foot = foot & vbCrLf & vbCrLf & _
        "ADD a medication:  click 'Add Medication'." & vbCrLf & _
        "REMOVE one:  click its row, then 'Remove Selected'." & vbCrLf & _
        "Re-check anytime:  click 'Review & Validate'."

    MsgBox head & summary & foot, vbInformation, "Medication Review"
End Sub

' Manually add a medication row.
Public Sub AddMedicationRow()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)

    Dim nm As String
    nm = Trim(InputBox("Enter the MEDICATION NAME to add:", "Add Medication", ""))
    If nm = "" Then Exit Sub

    Dim rec As MedRecord
    rec.MedName = nm
    rec.Strength = Trim(InputBox("Strength for " & nm & "  (e.g. 10 mg)" & vbCrLf & _
                    "Leave blank if unknown:", "Add Medication - Strength", ""))
    rec.DosageForm = ""
    rec.SIG = Trim(InputBox("Instructions (SIG) for " & nm & vbCrLf & _
                    "Leave blank if unknown:", "Add Medication - Instructions", ""))
    rec.Quantity = ""
    rec.Refills = ""
    rec.Expiration = ""
    rec.LotNumber = ""
    rec.Confidence = "Manual"
    rec.Warnings = ""
    rec.RawText = "[manually added]"

    Dim wsI As Worksheet
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    Dim rr As Long
    rr = FirstEmptyRow(ws)
    WriteMedRow ws, rr, rec, Trim(wsI.Range("C5").Value), Trim(wsI.Range("C6").Value), _
                Trim(wsI.Range("C7").Value), 0
    Call RenumberMeds
    Call ValidateMedications(False)
    ws.Activate
    ws.Cells(rr, C_NAME).Select
    MsgBox nm & " added." & vbCrLf & _
           "Fill in any missing details - they are highlighted.", _
           vbInformation, "Medication Added"
End Sub

' Remove the medication on the currently selected row.
Public Sub RemoveSelectedMedication()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)

    Dim lastRow As Long, r As Long, cnt As Integer, names As String
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    cnt = 0
    names = ""
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" And IsRowSelected(ws, r) Then
            cnt = cnt + 1
            names = names & "   - " & Trim(ws.Cells(r, C_NAME).Value & " " & _
                    ws.Cells(r, C_STR).Value) & vbCrLf
        End If
    Next r

    If cnt = 0 Then
        MsgBox "No rows are checked." & vbCrLf & _
               "Check the 'Print?' box on each medication you want to remove, then try again.", _
               vbExclamation, "Remove Selected"
        Exit Sub
    End If

    If MsgBox("Remove these " & cnt & " checked medication(s)?" & vbCrLf & vbCrLf & _
              names & vbCrLf & "(This cannot be undone.)", _
              vbYesNo + vbExclamation, "Remove Selected") = vbNo Then Exit Sub

    ' Delete bottom-to-top so row indexes stay valid. Delete only the table
    ' columns (1..C_SEL) and shift up, so the side buttons in col 17 do not move.
    Application.EnableEvents = False
    For r = lastRow To MEDS_HDR_ROWS + 1 Step -1
        If Trim(ws.Cells(r, C_NAME).Value) <> "" And IsRowSelected(ws, r) Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, C_SEL)).Delete Shift:=xlUp
        End If
    Next r
    Application.EnableEvents = True

    Call RenumberMeds
    Call ValidateMedications(False)
    Call ApplyAllRowStates(ws)
    MsgBox "Removed " & cnt & " medication(s).", vbInformation, "Remove Selected"
End Sub

Private Sub RenumberMeds()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    Dim r As Long, n As Integer
    n = 0
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" Then
            n = n + 1
            ws.Cells(r, C_NUM).Value = n
        End If
    Next r
End Sub

' ============================================================
'  LABEL PREVIEW
' ============================================================
Public Sub UpdateLabelPreviewFromSelection()
    Dim wsM As Worksheet
    Dim wsL As Worksheet
    Dim wsI As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsL = ThisWorkbook.Sheets(SH_LABEL)
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)

    ' Find selected row in Medications sheet
    Dim selRow As Long
    If ActiveSheet.Name = SH_MEDS Then
        selRow = ActiveCell.Row
    Else
        ' Find first row with data
        selRow = MEDS_HDR_ROWS + 1
    End If

    If selRow <= MEDS_HDR_ROWS Then
        MsgBox "Please click on a medication row in the Medications tab first.", _
               vbExclamation, "No Row Selected"
        Exit Sub
    End If

    ' Check if row has data
    If Trim(wsM.Cells(selRow, C_NAME).Value) = "" Then
        MsgBox "Selected row appears to be empty. Click on a medication name cell first.", _
               vbExclamation, "Empty Row"
        Exit Sub
    End If

    ' Build label fields
    Dim patName  As String : patName  = Trim(wsI.Range("C5").Value)
    Dim dob      As String : dob      = Trim(wsI.Range("C6").Value)
    Dim dateRx   As String : dateRx   = Trim(wsM.Cells(selRow, C_DATE).Value)
    Dim medName  As String : medName  = Trim(wsM.Cells(selRow, C_NAME).Value)
    Dim strength As String : strength = Trim(wsM.Cells(selRow, C_STR).Value)
    Dim qty      As String : qty      = Trim(wsM.Cells(selRow, C_QTY).Value)
    Dim sig      As String : sig      = Trim(wsM.Cells(selRow, C_SIG).Value)
    Dim expDate  As String : expDate  = Trim(wsM.Cells(selRow, C_EXP).Value)
    Dim lotNum   As String : lotNum   = Trim(wsM.Cells(selRow, C_LOT).Value)

    ' Validate required label fields
    Dim missing As String
    missing = ""
    If patName = "" Then missing = missing & "Patient Name  "
    If dob = "" Then missing = missing & "DOB  "
    If expDate = "" Then missing = missing & "Expiration  "
    If lotNum = "" Then missing = missing & "Lot Number  "

    If missing <> "" Then
        If MsgBox("The following fields are still empty:" & vbCrLf & missing & vbCrLf & vbCrLf & _
                  "Print anyway?", vbYesNo + vbExclamation, "Missing Label Fields") = vbNo Then
            Exit Sub
        End If
    End If

    ' Build med line:  Medication Strength (QTY: ##)
    Dim medLine As String
    If strength <> "" Then
        medLine = medName & " " & strength
    Else
        medLine = medName
    End If
    If qty <> "" Then medLine = medLine & " (QTY: " & qty & ")"

    ' Build Exp / Lot line
    Dim expLotLine As String
    expLotLine = "Exp: " & IIf(expDate <> "", expDate, "______") & _
                 "     Lot: " & IIf(lotNum <> "", lotNum, "______")

    ' Write to label preview sheet
    ' Row 2: clinic header (static - already formatted)
    ' Row 4: patient name
    wsL.Cells(4, 1).Value = IIf(patName <> "", patName, "[Patient Name]")
    ' Row 5: DOB
    wsL.Cells(5, 1).Value = "DOB: " & IIf(dob <> "", dob, "00/00/0000")
    ' Row 6: Date of Rx
    wsL.Cells(6, 1).Value = "Date of Rx: " & IIf(dateRx <> "", dateRx, "00/00/0000")
    ' Row 7: Medication + Qty
    wsL.Cells(7, 1).Value = medLine
    ' Row 8: SIG
    wsL.Cells(8, 1).Value = IIf(sig <> "", sig, "[Instructions not found - enter manually]")
    ' Row 9: Exp + Lot
    wsL.Cells(9, 1).Value = expLotLine

    ' Label Preview sheet is hidden - it is only the internal print surface.
    ' Do not activate it; the visible previews live on the "Label Previews" tab.
End Sub

' ============================================================
'  PRINT LABEL
' ============================================================
Private Function MedConfirmBlock(ws As Worksheet, ByVal r As Long) As String
    Dim s As String
    s = "    Medication:  " & Trim(ws.Cells(r, C_NAME).Value & " " & _
            ws.Cells(r, C_STR).Value & " " & ws.Cells(r, C_FORM).Value) & vbCrLf
    s = s & "    Directions:  " & ws.Cells(r, C_SIG).Value & vbCrLf
    s = s & "    Quantity:    " & ws.Cells(r, C_QTY).Value & _
            "      Refills:  " & ws.Cells(r, C_REF).Value & vbCrLf
    s = s & "    Expiration:  " & ws.Cells(r, C_EXP).Value & _
            "      Lot:  " & ws.Cells(r, C_LOT).Value
    MedConfirmBlock = s
End Function

Private Function MedConfirmLine(ws As Worksheet, ByVal r As Long) As String
    MedConfirmLine = Trim(ws.Cells(r, C_NAME).Value & " " & ws.Cells(r, C_STR).Value) & _
        "   (Qty " & ws.Cells(r, C_QTY).Value & ", Exp " & ws.Cells(r, C_EXP).Value & _
        ", Lot " & ws.Cells(r, C_LOT).Value & ")"
End Function

Public Sub PrintLabel()
    ' Remember which medication row is selected BEFORE the preview steals focus
    Dim medRowToMark As Long
    medRowToMark = 0
    If ActiveSheet.Name = SH_MEDS Then
        If ActiveCell.Row > MEDS_HDR_ROWS Then medRowToMark = ActiveCell.Row
    End If

    ' Ensure label preview is current
    Call UpdateLabelPreviewFromSelection

    Dim wsL As Worksheet
    Set wsL = ThisWorkbook.Sheets(SH_LABEL)

    ' Set up page for Brother QL-1100c  DK-1202  62mm x 100mm die-cut label
    ' LANDSCAPE: 100mm wide x 62mm tall  =  approx. 3.9" x 2.4"
    ' The label content block (A1:H10) is proportioned to print at ~100% scale.
    With wsL.PageSetup
        .PrintArea       = wsL.Range("A1:H10").Address
        .LeftMargin      = Application.InchesToPoints(0.04)
        .RightMargin     = Application.InchesToPoints(0.04)
        .TopMargin       = Application.InchesToPoints(0.04)
        .BottomMargin    = Application.InchesToPoints(0.04)
        .HeaderMargin    = 0
        .FooterMargin    = 0
        .FitToPagesWide  = False
        .FitToPagesTall  = False
        .Zoom            = 100
        .Orientation     = xlLandscape
        .CenterHorizontally = True
        .CenterVertically   = False
    End With

    ' Auto-select the Brother QL-1100c label printer
    Dim brotherName As String
    brotherName = SelectBrotherPrinter()

    If brotherName = "" Then
        Dim resp As Integer
        resp = MsgBox("The Brother QL-1100c label printer was not found." & vbCrLf & vbCrLf & _
            "Please check that it is:" & vbCrLf & _
            "   - plugged in and powered on" & vbCrLf & _
            "   - loaded with the DK-1202 (62 x 100 mm) label roll" & vbCrLf & _
            "   - installed in Windows (Settings > Bluetooth & devices > Printers)" & vbCrLf & vbCrLf & _
            "Click OK to choose a printer manually, or Cancel to stop.", _
            vbOKCancel + vbExclamation, "Brother Printer Not Found")
        If resp = vbCancel Then Exit Sub
        On Error Resume Next
        Application.Dialogs(xlDialogPrint).Show
        On Error GoTo 0
        Call LogPrint(medRowToMark, AskInitials())
        Call MarkPrinted(medRowToMark)
        Exit Sub
    End If

    ' Confirm with the volunteer before committing a label - show what will print
    Dim detail As String
    If medRowToMark > MEDS_HDR_ROWS Then
        detail = MedConfirmBlock(ThisWorkbook.Sheets(SH_MEDS), medRowToMark)
    Else
        detail = "    (the label currently shown on the Label Preview tab)"
    End If
    If MsgBox("About to print this label:" & vbCrLf & vbCrLf & _
        detail & vbCrLf & vbCrLf & _
        "Printer:  " & brotherName & vbCrLf & _
        "Is the DK-1202 (62 x 100 mm) label roll loaded and ready?" & vbCrLf & vbCrLf & _
        "YES = print now       NO = cancel", _
        vbYesNo + vbQuestion, "Confirm Label Print") = vbNo Then Exit Sub

    Call LogPrint(medRowToMark, AskInitials())

    On Error Resume Next
    wsL.PrintOut Copies:=1, Collate:=True
    If Err.Number <> 0 Then
        MsgBox "Printing failed:  " & Err.Description & vbCrLf & vbCrLf & _
               "Check the printer connection and that a label roll is loaded, then try again.", _
               vbExclamation, "Print Error"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Call MarkPrinted(medRowToMark)
End Sub

Private Sub MarkPrinted(ByVal medRow As Long)
    If medRow > MEDS_HDR_ROWS Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets(SH_MEDS)
        ws.Cells(medRow, C_PRTD).Value = "Yes"
        Application.EnableEvents = False
        ws.Cells(medRow, C_CNT).Value = Val(ws.Cells(medRow, C_CNT).Value) + 1
        Application.EnableEvents = True
        ws.Cells(medRow, C_CNT).HorizontalAlignment = xlCenter
    End If
End Sub

' Find the Brother QL-1100c and make it Excel's active printer.
' Returns the printer name on success, "" if not found / not selectable.
Private Function SelectBrotherPrinter() As String
    Dim foundName As String, foundPort As String
    foundName = ""
    foundPort = ""

    ' 1) Ask Windows (WMI) for the exact installed printer name + port
    On Error Resume Next
    Dim svc As Object, col As Object, prn As Object
    Set svc = GetObject("winmgmts:\\.\root\cimv2")
    If Not svc Is Nothing Then
        Set col = svc.ExecQuery("SELECT Name, PortName FROM Win32_Printer")
        If Not col Is Nothing Then
            For Each prn In col
                Dim nm As String
                nm = LCase(prn.Name)
                If InStr(nm, "ql-1100") > 0 Or InStr(nm, "ql_1100") > 0 Or _
                   (InStr(nm, "brother") > 0 And InStr(nm, "ql") > 0 And InStr(nm, "1100") > 0) Then
                    foundName = prn.Name
                    foundPort = prn.PortName
                    Exit For
                End If
            Next prn
        End If
    End If
    On Error GoTo 0

    ' 2) If we know the exact name+port, try that combination first
    If foundName <> "" And foundPort <> "" Then
        On Error Resume Next
        Err.Clear
        Application.ActivePrinter = foundName & " on " & foundPort & ":"
        If Err.Number = 0 Then
            On Error GoTo 0
            SelectBrotherPrinter = foundName
            Exit Function
        End If
        Err.Clear
        Application.ActivePrinter = foundName & " on " & foundPort
        If Err.Number = 0 Then
            On Error GoTo 0
            SelectBrotherPrinter = foundName
            Exit Function
        End If
        On Error GoTo 0
    End If

    ' 3) Probe the Ne port suffixes for the found name or common names
    Dim candidates() As String
    If foundName <> "" Then
        ReDim candidates(0)
        candidates(0) = foundName
    Else
        ReDim candidates(2)
        candidates(0) = "Brother QL-1100c"
        candidates(1) = "Brother QL-1100"
        candidates(2) = "Brother QL-1100c LE"
    End If

    Dim i As Integer
    For i = 0 To UBound(candidates)
        If SetActivePrinterByName(candidates(i)) Then
            SelectBrotherPrinter = candidates(i)
            Exit Function
        End If
    Next i

    SelectBrotherPrinter = ""
End Function

' Set Application.ActivePrinter by base name, probing the Ne00:..Ne99: port aliases.
Private Function SetActivePrinterByName(ByVal printerName As String) As Boolean
    Dim i As Integer
    On Error Resume Next
    For i = 0 To 99
        Err.Clear
        Application.ActivePrinter = printerName & " on Ne" & Format(i, "00") & ":"
        If Err.Number = 0 Then
            On Error GoTo 0
            SetActivePrinterByName = True
            Exit Function
        End If
    Next i
    On Error GoTo 0
    SetActivePrinterByName = False
End Function

' ============================================================
'  DISPENSE LOG
' ============================================================
Private Function AskInitials() As String
    AskInitials = Trim(InputBox("Enter your initials (for the dispense log):", _
                  "Volunteer Initials", ""))
End Function

Private Sub LogPrint(ByVal medRow As Long, ByVal vol As String)
    On Error Resume Next
    Dim wsI As Worksheet, wsM As Worksheet, wsLg As Worksheet
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsLg = ThisWorkbook.Sheets(SH_LOG)

    Dim nextLog As Long
    nextLog = wsLg.Cells(wsLg.Rows.Count, 1).End(xlUp).Row + 1
    If nextLog <= LOG_HDR_ROWS Then nextLog = LOG_HDR_ROWS + 1

    wsLg.Cells(nextLog, 1).Value = Format(Now(), "MM/DD/YYYY HH:MM:SS")
    wsLg.Cells(nextLog, 2).Value = Trim(wsI.Range("C5").Value)
    wsLg.Cells(nextLog, 3).Value = Trim(wsI.Range("C6").Value)
    If medRow > MEDS_HDR_ROWS Then
        wsLg.Cells(nextLog, 4).Value = wsM.Cells(medRow, C_NAME).Value
        wsLg.Cells(nextLog, 5).Value = wsM.Cells(medRow, C_STR).Value
        wsLg.Cells(nextLog, 6).Value = wsM.Cells(medRow, C_SIG).Value
        wsLg.Cells(nextLog, 7).Value = wsM.Cells(medRow, C_QTY).Value
        wsLg.Cells(nextLog, 8).Value = wsM.Cells(medRow, C_REF).Value
        wsLg.Cells(nextLog, 9).NumberFormat = "@"
        wsLg.Cells(nextLog, 9).Value = wsM.Cells(medRow, C_EXP).Value
        wsLg.Cells(nextLog, 10).NumberFormat = "@"
        wsLg.Cells(nextLog, 10).Value = wsM.Cells(medRow, C_LOT).Value
        wsLg.Cells(nextLog, 11).Value = wsM.Cells(medRow, C_DATE).Value
        wsLg.Cells(nextLog, 13).Value = wsM.Cells(medRow, C_FORM).Value
        wsLg.Cells(nextLog, 14).Value = wsM.Cells(medRow, C_CNT).Value
    End If
    wsLg.Cells(nextLog, 12).Value = vol

    Dim c As Integer
    For c = 1 To 14
        With wsLg.Cells(nextLog, c)
            .Font.Name = "Arial"
            .Font.Size = 9
            If nextLog Mod 2 = 0 Then .Interior.Color = RGB(245, 245, 245)
        End With
    Next c
End Sub

' ============================================================
'  SELF-TEST HARNESS
'  Tools -> Macros -> RunParserTests
' ============================================================
Public Sub RunParserTests()
    Dim passed As Integer
    Dim failed As Integer
    Dim log As String
    passed = 0
    failed = 0
    log = ""

    Dim cases As Variant
    cases = GetTestCases()

    Dim i As Integer
    For i = 0 To UBound(cases, 1)
        Dim inText As String
        Dim expectName As String
        Dim expectStrength As String
        Dim expectCount As Integer

        inText        = cases(i, 0)
        expectName   = cases(i, 1)
        expectStrength = cases(i, 2)
        expectCount  = CInt(cases(i, 3))
        Dim desc     As String
        desc = cases(i, 4)

        ' Test block count
        Dim blocks() As String
        blocks = SplitMedBlocks(inText)
        Dim gotCount As Integer
        gotCount = UBound(blocks) + 1

        If gotCount <> expectCount Then
            failed = failed + 1
            log = log & "FAIL [" & desc & "] Expected " & expectCount & _
                  " block(s), got " & gotCount & vbCrLf
        Else
            ' Test first block parse
            Dim rec As MedRecord
            rec = ParseOneBlock(blocks(0))
            Dim nameOK As Boolean
            Dim strOK  As Boolean
            nameOK = (InStr(LCase(rec.MedName), LCase(expectName)) > 0)
            strOK  = (expectStrength = "" Or InStr(LCase(rec.Strength), LCase(expectStrength)) > 0)

            If nameOK And strOK Then
                passed = passed + 1
                log = log & "PASS [" & desc & "]" & vbCrLf
            Else
                failed = failed + 1
                log = log & "FAIL [" & desc & "]" & vbCrLf
                If Not nameOK Then
                    log = log & "  Name: expected '" & expectName & "', got '" & rec.MedName & "'" & vbCrLf
                End If
                If Not strOK Then
                    log = log & "  Strength: expected '" & expectStrength & "', got '" & rec.Strength & "'" & vbCrLf
                End If
            End If
        End If
    Next i

    Dim summary As String
    summary = "Parser Test Results" & vbCrLf & String(40, "-") & vbCrLf & _
              "Passed: " & passed & vbCrLf & _
              "Failed: " & failed & vbCrLf & _
              "Total:  " & (passed + failed) & vbCrLf & String(40, "-") & vbCrLf & log

    ' Write results to a new sheet
    Dim wsTest As Worksheet
    On Error Resume Next
    Set wsTest = ThisWorkbook.Sheets("Test Results")
    If wsTest Is Nothing Then
        Set wsTest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTest.Name = "Test Results"
    End If
    On Error GoTo 0

    wsTest.Cells.ClearContents
    wsTest.Range("A1").Value = "Parser Test Results - " & Format(Now(), "MM/DD/YYYY HH:MM")
    wsTest.Range("A2").Value = "Passed: " & passed & "  /  Failed: " & failed & _
                               "  /  Total: " & (passed + failed)

    Dim logLines() As String
    logLines = Split(log, vbCrLf)
    Dim lr As Integer
    For lr = 0 To UBound(logLines)
        wsTest.Cells(lr + 4, 1).Value = logLines(lr)
        If Left(Trim(logLines(lr)), 4) = "FAIL" Then
            wsTest.Cells(lr + 4, 1).Font.Color = RGB(198, 40, 40)
        ElseIf Left(Trim(logLines(lr)), 4) = "PASS" Then
            wsTest.Cells(lr + 4, 1).Font.Color = RGB(46, 125, 50)
        End If
    Next lr

    wsTest.Columns("A").AutoFit
    wsTest.Activate

    MsgBox "Tests complete: " & passed & " passed, " & failed & " failed." & vbCrLf & _
           "See 'Test Results' sheet for details.", _
           IIf(failed = 0, vbInformation, vbExclamation), "Parser Tests"
End Sub

Private Function GetTestCases() As Variant
    ' Array columns: input, expectedName, expectedStrength, expectedBlockCount, description
    Dim t(29, 4) As Variant

    ' -- Single medication, basic ------------------------------
    t(0, 0) = "metFORMIN 1,000 mg tablet, 1 tab(s) orally 2 times a day with meals" & vbLf & _
               "1 refill" & vbLf & _
               "Started on 02/04/2026 09:15 PM by Sobieski, James" & vbLf & _
               "Ascension Rx 1112" & vbLf & "See All"
    t(0, 1) = "Metformin"
    t(0, 2) = "1,000 mg"
    t(0, 3) = "1"
    t(0, 4) = "Metformin basic"

    t(1, 0) = "empagliflozin 10 mg tablet," & vbLf & _
               "Take one tablet by mouth every morning" & vbLf & _
               "30 tablets, 0 refills"
    t(1, 1) = "Empagliflozin"
    t(1, 2) = "10 mg"
    t(1, 3) = "1"
    t(1, 4) = "Empagliflozin basic"

    t(2, 0) = "Basaglar KwikPen U-100 Insulin" & vbLf & _
               "100 unit/mL (3 mL)" & vbLf & _
               "20 unit(s) subcutaneously every evening" & vbLf & _
               "10 (3 mL syringe. Total quantity: 30 mLs)"
    t(2, 1) = "Basaglar"
    t(2, 2) = "100 unit"
    t(2, 3) = "1"
    t(2, 4) = "Insulin pen"

    ' -- Multiple medications ----------------------------------
    t(3, 0) = "lisinopril 10 mg tablet" & vbLf & _
               "Take one tablet by mouth once daily" & vbLf & _
               "30 tablets, 3 refills" & vbLf & _
               "Started on 01/01/2026 by Smith, John" & vbLf & "See All" & vbLf & _
               "atorvastatin 40 mg tablet" & vbLf & _
               "Take one tablet by mouth at bedtime" & vbLf & _
               "30 tablets, 3 refills" & vbLf & "See All"
    t(3, 1) = "Lisinopril"
    t(3, 2) = "10 mg"
    t(3, 3) = "2"
    t(3, 4) = "Two medications with See All"

    t(4, 0) = "amLODIPine 5 mg tablet, 1 tab(s) orally once a day" & vbLf & _
               "30 tablets, 6 refills" & vbLf & "See All" & vbLf & _
               "amITriptyline 25 mg tablet" & vbLf & _
               "Take one tablet by mouth at bedtime as needed for pain" & vbLf & _
               "30 tablets, 0 refills" & vbLf & "See All" & vbLf & _
               "atorvaSTATin 40 mg tablet" & vbLf & _
               "Take one tablet by mouth at bedtime" & vbLf & _
               "30 tablets, 3 refills" & vbLf & "See All"
    t(4, 1) = "Amlodipine"
    t(4, 2) = "5 mg"
    t(4, 3) = "3"
    t(4, 4) = "Three medications, TALL MAN lettering"

    ' -- Missing quantity --------------------------------------
    t(5, 0) = "sertraline 50 mg tablet" & vbLf & _
               "Take one tablet by mouth every morning" & vbLf & _
               "2 refills"
    t(5, 1) = "Sertraline"
    t(5, 2) = "50 mg"
    t(5, 3) = "1"
    t(5, 4) = "Missing quantity"

    ' -- Missing refills ---------------------------------------
    t(6, 0) = "gabapentin 300 mg capsule" & vbLf & _
               "Take two capsules by mouth three times daily" & vbLf & _
               "180 capsules"
    t(6, 1) = "Gabapentin"
    t(6, 2) = "300 mg"
    t(6, 3) = "1"
    t(6, 4) = "Missing refills"

    ' -- All-lowercase drug name -------------------------------
    t(7, 0) = "omeprazole 20 mg capsule" & vbLf & _
               "Take one capsule by mouth once daily before breakfast" & vbLf & _
               "30 capsules, 5 refills"
    t(7, 1) = "Omeprazole"
    t(7, 2) = "20 mg"
    t(7, 3) = "1"
    t(7, 4) = "All lowercase drug name"

    ' -- Cream -------------------------------------------------
    t(8, 0) = "triamcinolone acetonide 0.1% cream" & vbLf & _
               "Apply a thin layer to affected area twice daily" & vbLf & _
               "1 tube, 2 refills"
    t(8, 1) = "Triamcinolone Acetonide"
    t(8, 2) = "0.1%"
    t(8, 3) = "1"
    t(8, 4) = "Topical cream"

    ' -- Eye drops --------------------------------------------
    t(9, 0) = "latanoprost 0.005% eye drops" & vbLf & _
               "Instill 1 drop in affected eye(s) every evening" & vbLf & _
               "2.5 mL, 3 refills"
    t(9, 1) = "Latanoprost"
    t(9, 2) = "0.005%"
    t(9, 3) = "1"
    t(9, 4) = "Eye drops"

    ' -- Inhaler ----------------------------------------------
    t(10, 0) = "albuterol sulfate 90 mcg/actuation inhaler" & vbLf & _
                "Inhale 2 puff(s) by mouth every 4 hours as needed for shortness of breath" & vbLf & _
                "1 inhaler, 3 refills"
    t(10, 1) = "Albuterol Sulfate"
    t(10, 2) = "90 mcg"
    t(10, 3) = "1"
    t(10, 4) = "Inhaler"

    ' -- Liquid suspension ------------------------------------
    t(11, 0) = "amoxicillin 250 mg/5 mL oral suspension" & vbLf & _
                "Take 5 mL by mouth three times daily for 10 days" & vbLf & _
                "150 mL, 0 refills"
    t(11, 1) = "Amoxicillin"
    t(11, 2) = "250 mg"
    t(11, 3) = "1"
    t(11, 4) = "Liquid suspension"

    ' -- PRN medication ---------------------------------------
    t(12, 0) = "hydrOXYzine 25 mg tablet" & vbLf & _
                "Take 1 tablet by mouth every 6 hours as needed for anxiety" & vbLf & _
                "30 tablets, 0 refills"
    t(12, 1) = "Hydroxyzine"
    t(12, 2) = "25 mg"
    t(12, 3) = "1"
    t(12, 4) = "PRN with TALL MAN"

    ' -- Steroid taper (complex SIG) --------------------------
    t(13, 0) = "prednisone 10 mg tablet" & vbLf & _
                "Take 4 tablets daily for 3 days, then 3 tablets daily for 3 days, " & _
                "then 2 tablets daily for 3 days, then 1 tablet daily for 3 days, then stop" & vbLf & _
                "30 tablets, 0 refills"
    t(13, 1) = "Prednisone"
    t(13, 2) = "10 mg"
    t(13, 3) = "1"
    t(13, 4) = "Steroid taper"

    ' -- Long drug name (combination) -------------------------
    t(14, 0) = "amlodipine-benazepril 5-10 mg capsule" & vbLf & _
                "Take one capsule by mouth once daily" & vbLf & _
                "30 capsules, 3 refills"
    t(14, 1) = "Amlodipine-Benazepril"
    t(14, 2) = "5"
    t(14, 3) = "1"
    t(14, 4) = "Combination drug"

    ' -- Patch ------------------------------------------------
    t(15, 0) = "nicotine 21 mg/24 hr transdermal patch" & vbLf & _
                "Apply one patch to skin daily, rotate sites" & vbLf & _
                "14 patches, 0 refills"
    t(15, 1) = "Nicotine"
    t(15, 2) = "21 mg"
    t(15, 3) = "1"
    t(15, 4) = "Transdermal patch"

    ' -- Extra spaces & blank lines ---------------------------
    t(16, 0) = "  metoprolol succinate   50 mg   tablet  " & vbLf & vbLf & _
                "   Take one tablet by mouth daily   " & vbLf & vbLf & _
                "  30 tablets, 1 refill  "
    t(16, 1) = "Metoprolol Succinate"
    t(16, 2) = "50 mg"
    t(16, 3) = "1"
    t(16, 4) = "Extra spaces and blank lines"

    ' -- Duplicate medication names (should produce 2 blocks) -
    t(17, 0) = "metformin 500 mg tablet" & vbLf & _
                "Take one tablet twice daily" & vbLf & _
                "60 tablets, 3 refills" & vbLf & "See All" & vbLf & _
                "metformin 1000 mg tablet" & vbLf & _
                "Take one tablet twice daily with evening meal" & vbLf & _
                "30 tablets, 3 refills" & vbLf & "See All"
    t(17, 1) = "Metformin"
    t(17, 2) = "500 mg"
    t(17, 3) = "2"
    t(17, 4) = "Duplicate medication different strength"

    ' -- Five medications at once ------------------------------
    t(18, 0) = "Tylenol 325 mg tablet" & vbLf & "Take 2 tablets every 6 hours as needed for pain" & vbLf & "30 tablets, 2 refills" & vbLf & "See All" & vbLf & _
                "ibuprofen 400 mg tablet" & vbLf & "Take one tablet three times daily with food" & vbLf & "90 tablets, 0 refills" & vbLf & "See All" & vbLf & _
                "amlodipine 5 mg tablet" & vbLf & "Take one tablet daily" & vbLf & "30 tablets, 6 refills" & vbLf & "See All" & vbLf & _
                "amitriptyline 25 mg tablet" & vbLf & "Take one tablet at bedtime" & vbLf & "30 tablets, 0 refills" & vbLf & "See All" & vbLf & _
                "atorvastatin 40 mg tablet" & vbLf & "Take one tablet at bedtime" & vbLf & "30 tablets, 3 refills" & vbLf & "See All"
    t(18, 1) = "Tylenol"
    t(18, 2) = "325 mg"
    t(18, 3) = "5"
    t(18, 4) = "Five medications"

    ' -- Brand-name insulin (preserve brand) ------------------
    t(19, 0) = "NovoLOG FlexPen 100 unit/mL insulin aspart" & vbLf & _
                "Inject 10 units subcutaneously before each meal" & vbLf & _
                "5 pens, 2 refills"
    t(19, 1) = "Novolog Flexpen"
    t(19, 2) = "100 unit"
    t(19, 3) = "1"
    t(19, 4) = "Brand insulin pen"

    ' -- Low-dose aspirin (unusual presentation) ---------------
    t(20, 0) = "aspirin 81 mg tablet" & vbLf & _
                "Take one tablet by mouth daily"
    t(20, 1) = "Aspirin"
    t(20, 2) = "81 mg"
    t(20, 3) = "1"
    t(20, 4) = "Low-dose aspirin, no refills line"

    ' -- Nasal spray ------------------------------------------
    t(21, 0) = "fluticasone propionate 50 mcg/spray nasal spray" & vbLf & _
                "Spray 2 sprays in each nostril once daily" & vbLf & _
                "1 bottle, 5 refills"
    t(21, 1) = "Fluticasone Propionate"
    t(21, 2) = "50 mcg"
    t(21, 3) = "1"
    t(21, 4) = "Nasal spray"

    ' -- Suppository ------------------------------------------
    t(22, 0) = "promethazine 25 mg suppository" & vbLf & _
                "Insert 1 suppository rectally every 4 to 6 hours as needed for nausea" & vbLf & _
                "12 suppositories, 0 refills"
    t(22, 1) = "Promethazine"
    t(22, 2) = "25 mg"
    t(22, 3) = "1"
    t(22, 4) = "Suppository"

    ' -- Unexpected text / garbage lines ----------------------
    t(23, 0) = "!!! URGENT !!!" & vbLf & _
                "metformin 500 mg tablet" & vbLf & _
                "Take one tablet twice daily" & vbLf & _
                "60 tablets, 3 refills" & vbLf & _
                "PLEASE VERIFY BEFORE DISPENSING"
    t(23, 1) = "Metformin"
    t(23, 2) = "500 mg"
    t(23, 3) = "1"
    t(23, 4) = "Unexpected text garbage"

    ' -- Medication with no SIG (just strength and refills) ---
    t(24, 0) = "levothyroxine 50 mcg tablet" & vbLf & _
                "30 tablets, 11 refills"
    t(24, 1) = "Levothyroxine"
    t(24, 2) = "50 mcg"
    t(24, 3) = "1"
    t(24, 4) = "No SIG line"

    ' -- High-dose vitamin D -----------------------------------
    t(25, 0) = "cholecalciferol (Vitamin D3) 50,000 unit capsule" & vbLf & _
                "Take one capsule by mouth once weekly" & vbLf & _
                "12 capsules, 3 refills"
    t(25, 1) = "Cholecalciferol"
    t(25, 2) = "50,000 unit"
    t(25, 3) = "1"
    t(25, 4) = "High-dose Vitamin D"

    ' -- Complex: SIG on same line as drug --------------------
    t(26, 0) = "metformin 500 mg tablet, Take 1 tab by mouth twice daily with meals" & vbLf & _
                "60 tablets, 3 refills"
    t(26, 1) = "Metformin"
    t(26, 2) = "500 mg"
    t(26, 3) = "1"
    t(26, 4) = "SIG on same line as drug"

    ' -- Generic with pharmacy number artifact ----------------
    t(27, 0) = "rosuvastatin 10 mg tablet" & vbLf & _
                "Take one tablet by mouth at bedtime" & vbLf & _
                "30 tablets, 3 refills" & vbLf & _
                "Started on 03/15/2026 by Jones, Mary" & vbLf & _
                "CVS Pharmacy 4421" & vbLf & "See All"
    t(27, 1) = "Rosuvastatin"
    t(27, 2) = "10 mg"
    t(27, 3) = "1"
    t(27, 4) = "With pharmacy and provider lines"

    ' -- g (gram) strength ------------------------------------
    t(28, 0) = "metronidazole 0.75% vaginal gel" & vbLf & _
                "Apply one applicatorful vaginally once daily at bedtime for 5 days" & vbLf & _
                "1 tube, 0 refills"
    t(28, 1) = "Metronidazole"
    t(28, 2) = "0.75%"
    t(28, 3) = "1"
    t(28, 4) = "Vaginal gel"

    ' -- Z-pack / complex dosing days -------------------------
    t(29, 0) = "azithromycin 250 mg tablet" & vbLf & _
                "Take 2 tablets on Day 1, then 1 tablet daily on Days 2-5" & vbLf & _
                "6 tablets, 0 refills"
    t(29, 1) = "Azithromycin"
    t(29, 2) = "250 mg"
    t(29, 3) = "1"
    t(29, 4) = "Z-pack complex dosing"

    GetTestCases = t
End Function

' ============================================================
'  UTILITY
' ============================================================
Private Function ReplaceCI(s As String, find As String, repl As String) As String
    Dim pos As Integer
    Dim result As String
    result = s
    pos = 1
    Do
        Dim found As Integer
        found = InStr(pos, result, find, vbTextCompare)
        If found = 0 Then Exit Do
        result = Left(result, found - 1) & repl & Mid(result, found + Len(find))
        pos = found + Len(repl)
    Loop
    ReplaceCI = result
End Function

Private Function IIf(condition As Boolean, trueVal As String, falseVal As String) As String
    If condition Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function

' ============================================================
'  AUTO-RUN on workbook open
' ===========================================