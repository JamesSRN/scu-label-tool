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
Private Const SH_GUIDE  As String = "Start Here"
Private Const SH_TEBRA  As String = "TEBRA TEMPLATE"

' -- Medications sheet column indices ------------------------
Private Const C_NUM     As Integer = 1
Private Const C_NAME    As Integer = 2
Private Const C_STR     As Integer = 3
Private Const C_FORM    As Integer = 4
Private Const C_SIG     As Integer = 5
Private Const C_QTY     As Integer = 6
Private Const C_EXP     As Integer = 7
Private Const C_LOT     As Integer = 8
Private Const C_SRC     As Integer = 9    ' Source dropdown (DOH / IN HOUSE / RxAPS / Other), logged; right of Lot #
Private Const C_DATE    As Integer = 10
Private Const C_REF     As Integer = 11   ' Refills, right of Date of Rx
Private Const C_CONF    As Integer = 12
Private Const C_WARN    As Integer = 13
Private Const C_RAW     As Integer = 14
Private Const C_PRTD    As Integer = 15
Private Const C_CNT     As Integer = 16
Private Const C_SEL     As Integer = 17

' Dispense Log column map (1-based). Encounter sits right after Timestamp; the rest follow
' in reading order. Every Log read/write and the header row use these, so reordering the
' Log is just a matter of changing these numbers.
Private Const LG_TIME As Long = 1     ' Timestamp
Private Const LG_ENC  As Long = 2     ' Encounter #
Private Const LG_PT   As Long = 3     ' Patient
Private Const LG_DOB  As Long = 4     ' DOB
Private Const LG_NAME As Long = 5     ' Medication
Private Const LG_STR  As Long = 6     ' Strength
Private Const LG_SIG  As Long = 7     ' Directions
Private Const LG_QTY  As Long = 8     ' Quantity
Private Const LG_REF  As Long = 9     ' Refills
Private Const LG_EXP  As Long = 10    ' Expiration
Private Const LG_LOT  As Long = 11    ' Lot
Private Const LG_SRC  As Long = 12    ' Source
Private Const LG_DATE As Long = 13    ' Rx Date
Private Const LG_INIT As Long = 14    ' Initials
Private Const LG_FORM As Long = 15    ' Dosage Form
Private Const LG_CNT  As Long = 16    ' Print #
Private Const LG_LAST As Long = 16    ' rightmost Log column

' Encounter snapshot store (hidden sheet). One row per (encounter, med) capturing the FULL
' medication detail so a past encounter can be reopened and edited exactly as it was.
Private Const SH_ENC As String = "EncounterData"
Private Const ES_ENC    As Long = 1     ' Encounter #
Private Const ES_PT     As Long = 2     ' Patient
Private Const ES_DOB    As Long = 3     ' DOB
Private Const ES_RXDATE As Long = 4     ' Patient Rx date (Input C7)
Private Const ES_NAME   As Long = 5
Private Const ES_STR    As Long = 6
Private Const ES_FORM   As Long = 7
Private Const ES_SIG    As Long = 8
Private Const ES_QTY    As Long = 9
Private Const ES_EXP    As Long = 10
Private Const ES_LOT    As Long = 11
Private Const ES_SRC    As Long = 12
Private Const ES_DATE   As Long = 13
Private Const ES_REF    As Long = 14
Private Const ES_LAST   As Long = 14

Private Const MEDS_HDR_ROWS As Integer = 3   ' rows before data begins
Private Const LOG_HDR_ROWS  As Integer = 2

' -- Colours (hex, no #) -------------------------------------
Private Const CLR_HIGH   As Long = 12780748   ' &HC8E6C  pastel green
Private Const CLR_MED    As Long = 16775620   ' &HFFF9C4 pastel yellow
Private Const CLR_LOW    As Long = 16764106   ' &HFFCDD2 pastel red
Private Const CLR_MANUAL As Long = 16773600   ' &HFFF3E0 orange-tint
Private Const CLR_WARN   As Long = 16775936   ' &HFFF800 warn yellow
Private Const CLR_WHITE  As Long = 16777215

' DK-1202 die-cut label: 62 mm x 100 mm. Landscape print uses the 100 mm edge
' as page width (~283 pt). 228 = safe minimum; 242 uses more of the 100 mm die-cut.
Private Const LABEL_WIDTH_PT As Double = 242
Private Const LABEL_COPIES   As Long = 2   ' every label prints this many copies

Private Const FONT_LABEL_BODY As String = "Arial"
Private Const FONT_LABEL_HDR As String = "Century Gothic"
Private Const FONT_LABEL_HDR_FB As String = "Arial"   ' fallback when Helvetica is unavailable
Private Const LOGO_ASPECT As Double = 1.488   ' manual crop width / height (6150 / 4133)
Private Const LOGO_EMBLEM_FILE As String = "scu_emblem.png"
Private Const LOGO_EMBLEM_SOURCE As String = "cropped_Black SCU Logo + Transparent Background - Copy.png"
Private Const LOGO_HEIGHT_PRINT As Single = 30    ' emblem height on printed label (pt)
Private Const LOGO_HEIGHT_GALLERY As Single = 28  ' emblem height in Label Previews gallery
Private Const LOGO_HDR_ROW1_PT As Single = 18     ' clinic name row (paired with row 3 for tight header)
Private Const LOGO_HDR_ROW2_PT As Single = 12     ' address row
Private Const LOGO_RIGHT_PAD_PT As Single = 0     ' flush emblem to right edge of print area
Private Const LOGO_SLOT_INSET_PT As Single = 4    ' inset from slot left - keeps emblem right without over-shrinking
Private Const LOGO_GALLERY_INSET_PT As Single = 2  ' single-column gallery slot
Private Const CLINIC_NAME_FONT_PRINT As Single = 14
Private Const CLINIC_NAME_FONT_GALLERY As Single = 12
Private Const CLINIC_ADDR_FONT_PRINT As Single = 8.5
Private Const CLINIC_NAMESUB_FONT_PRINT As Single = 7
Private Const CLINIC_PHONE_FONT_PRINT As Single = 11

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

' Holds the "please wait" popup while Print Checked Labels locates the printer.
Private busyFrm As Object

' Last successfully-printed batch, for "Reprint Last Batch" (session-only; cleared on
' close). gLastBatchRows is a comma-separated list of Medications-tab row numbers.
Private gLastBatchRows As String
Private gLastBatchVol As String
Private gEditingEncounter As Long   ' >0 while a past encounter is loaded for editing; 0 otherwise

' V2: append each print to a dated local CSV archive (PHI - stays on this machine,
' git-ignored) so the day's dispensing record survives the on-close Log wipe.
' Set to False to disable the archive entirely.
Private Const DISPENSE_CSV_ENABLED As Boolean = True

' V2 app version - stamped into the workbook by SetupWorkbook so the loaded build is visible.
Public Const APP_VERSION As String = "2.0"

' V2 debug switch: when True, Dbg() writes a timestamped trace to the VBE Immediate
' window (Ctrl+G). Leave False for production (Dbg is then a no-op).
Public Const DEBUG_MODE As Boolean = False

' Session cache for the resolved Brother printer name (skips the slow WMI lookup on
' repeat prints; self-heals if the cached printer can no longer be selected).
Private gCachedPrinter As String

' Medication-name wrapping - shared by the print label AND the gallery so the two stay
' in sync: a name longer than MED_WRAP_MAXLEN chars wraps to two lines at MED_WRAP_FONT pt
' instead of shrinking to a single small line.
Private Const MED_WRAP_MAXLEN As Long = 38
Private Const MED_WRAP_FONT As Single = 11

' ============================================================
'  BUSY / PROGRESS POPUP  (used during the print-prep delay)
' ============================================================
Public Sub BusyShow(ByVal pct As Long, ByVal msg As String)
    ' Show (or update) a small progress popup. Falls back to the status bar
    ' if the frmBusy UserForm isn't present in this workbook.
    On Error GoTo StatusFallback
    If busyFrm Is Nothing Then
        Set busyFrm = VBA.UserForms.Add("frmBusy")
    End If
    busyFrm.SetProgress pct, msg
    busyFrm.Show vbModeless
    DoEvents
    Exit Sub
StatusFallback:
    On Error Resume Next
    Set busyFrm = Nothing
    Application.StatusBar = "Please wait - " & msg & " (" & pct & "%)"
    DoEvents
    On Error GoTo 0
End Sub

Public Sub BusyHide()
    On Error Resume Next
    If Not busyFrm Is Nothing Then
        busyFrm.Hide
        Unload busyFrm
        Set busyFrm = Nothing
    End If
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ============================================================
'  WORKBOOK SETUP  (run once after importing this module)
' ============================================================
Private Sub MatchHeaderFormat(ByVal src As Range, ByVal dest As Range)
    ' Copy header styling without fragile multi-cell PasteSpecial calls.
    On Error Resume Next
    src.Copy
    dest.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    If Err.Number <> 0 Then
        Err.Clear
        With dest
            .Font.Name = src.Font.Name
            .Font.Size = src.Font.Size
            .Font.Bold = src.Font.Bold
            .Font.Color = src.Font.Color
            .Interior.Color = src.Interior.Color
            .HorizontalAlignment = src.HorizontalAlignment
            .VerticalAlignment = src.VerticalAlignment
            .WrapText = src.WrapText
        End With
    End If
    On Error GoTo 0
End Sub

' Debug trace to the VBE Immediate window (only when DEBUG_MODE = True). No-op otherwise.
Private Sub Dbg(ByVal msg As String)
    If DEBUG_MODE Then Debug.Print Format(Now(), "hh:nn:ss") & "  " & msg
End Sub

' Verify the required sheets + the Print? column marker exist; alert ONLY on a problem so
' a volunteer who deleted or renamed something gets a clear message instead of a cryptic
' runtime error later. Called on workbook open.
Public Sub CheckWorkbookStructure()
    On Error Resume Next
    Dim problems As String, i As Integer, found As Boolean
    Dim names As Variant, sh As Object
    problems = ""
    names = Array(SH_INPUT, SH_MEDS, SH_LABEL, SH_LOG, SH_ALL)
    For i = LBound(names) To UBound(names)
        found = False
        For Each sh In ThisWorkbook.Sheets
            If sh.Name = names(i) Then found = True: Exit For
        Next sh
        If Not found Then problems = problems & "   - Missing sheet: '" & names(i) & "'" & vbCrLf
    Next i

    Dim wsMed As Worksheet
    Set wsMed = Nothing
    Set wsMed = ThisWorkbook.Sheets(SH_MEDS)
    If Not wsMed Is Nothing Then
        If InStr(LCase(CStr(wsMed.Cells(2, C_SEL).Value)), "print") = 0 Then
            problems = problems & "   - The 'Print?' column header is not where expected " & _
                       "(Medications columns may have been changed)." & vbCrLf
        End If
    End If

    If problems <> "" Then
        MsgBox "The workbook structure looks off:" & vbCrLf & vbCrLf & problems & vbCrLf & _
               "If a sheet or column was changed, rebuild from the template (Build-Release) " & _
               "or restore a backup before dispensing.", _
               vbExclamation, "Workbook Check"
    End If
    On Error GoTo 0
    Dbg "CheckWorkbookStructure: " & IIf(problems = "", "OK", "PROBLEMS")
End Sub

' Jump from the guide to the working tab (wired to the guide's button).
Public Sub GoToInput()
    On Error Resume Next
    ThisWorkbook.Sheets(SH_INPUT).Activate
    ThisWorkbook.Sheets(SH_INPUT).Range("C5").Select
    On Error GoTo 0
End Sub

' One numbered step block on the Start Here sheet: colored badge + title + wrapped text.
Private Sub GuideStep(ws As Worksheet, ByVal topRow As Long, ByVal num As Integer, ByVal clr As Long, ByVal title As String, ByVal txt As String)
    On Error Resume Next
    ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow + 3, 2)).Merge
    With ws.Cells(topRow, 2)
        .Value = num
        .Interior.Color = clr
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(topRow, 3), ws.Cells(topRow, 8)).Merge
    With ws.Cells(topRow, 3)
        .Value = title
        .Font.Bold = True
        .Font.Size = 12.5
        .Font.Color = clr
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(topRow + 1, 3), ws.Cells(topRow + 3, 8)).Merge
    With ws.Cells(topRow + 1, 3)
        .Value = txt
        .Font.Bold = False
        .Font.Size = 10.5
        .Font.Color = RGB(38, 50, 56)
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
    On Error GoTo 0
End Sub

' Build (or rebuild) the "Start Here" quick-start sheet and make it the FIRST sheet, so
' the workbook opens to it. Best-effort/cosmetic - never fails the build.
Private Sub BuildQuickStartSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(SH_GUIDE)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SH_GUIDE
    End If
    ws.Move Before:=ThisWorkbook.Sheets(1)

    Application.ScreenUpdating = False
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B").ColumnWidth = 6
    ws.Columns("C:H").ColumnWidth = 15

    ' Header band
    ws.Range("A1:H3").Interior.Color = RGB(0, 121, 107)
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 8)).Merge
    With ws.Cells(2, 1)
        .Value = "   SCU Label Printing"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 22
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(3, 1), ws.Cells(3, 8)).Merge
    With ws.Cells(3, 1)
        .Value = "   Quick-Start Guide   -   print a medication label in 4 steps        v" & APP_VERSION
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With

    Call GuideStep(ws, 5, 1, RGB(21, 101, 192), "Enter & paste", _
        "On the 'Patient & Input' tab, type the patient Name and DOB, paste the medication list into the box, then click PARSE MEDICATIONS.")
    Call GuideStep(ws, 10, 2, RGB(0, 121, 107), "Review & add Exp / Lot", _
        "On the 'Medications' tab click Review & Validate and fix anything flagged (an amber Expiration cell means check the date). When prompted, enter Expiration (MM/YYYY) and Lot. For a med split across two bottles, separate values with commas.")
    Call GuideStep(ws, 15, 3, RGB(216, 67, 21), "Check what to print", _
        "Double-click the 'Print?' cell next to each medication you want (or use 'Check this label' in the gallery). A green row means it is selected.")
    Call GuideStep(ws, 20, 4, RGB(46, 125, 50), "Print", _
        "Click 'Print Checked Labels', confirm the list, and enter your initials. Any label missing Exp/Lot is skipped and named. When done you land on the Log.")

    ws.Range(ws.Cells(26, 2), ws.Cells(26, 8)).Merge
    With ws.Cells(26, 2)
        .Value = "Good to know"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(0, 121, 107)
    End With
    Dim gk As Variant
    gk = Array( _
        "-   Reprint Last Batch:  reprints the last set after a paper jam or misfeed.", _
        "-   Start NEW Patient:  clears the screen for the next patient (the Log is kept).", _
        "-   Closing the file clears patient info and the Log; a dated CSV backup is saved.", _
        "-   Labels print on the Brother QL-1100c with the DK-1202 (62 x 100 mm) roll.")
    Dim gi As Integer
    For gi = 0 To UBound(gk)
        ws.Range(ws.Cells(27 + gi, 2), ws.Cells(27 + gi, 8)).Merge
        With ws.Cells(27 + gi, 2)
            .Value = gk(gi)
            .Font.Size = 10
            .Font.Color = RGB(38, 50, 56)
            .WrapText = False
            .VerticalAlignment = xlCenter
        End With
    Next gi

    ' Row heights
    ws.Rows(1).RowHeight = 6
    ws.Rows(2).RowHeight = 32
    ws.Rows(3).RowHeight = 20
    ws.Rows(4).RowHeight = 10
    Dim tops As Variant, ti As Integer, s As Long
    tops = Array(5, 10, 15, 20)
    For ti = 0 To 3
        s = tops(ti)
        ws.Rows(s).RowHeight = 16
        ws.Rows(s + 1).RowHeight = 15
        ws.Rows(s + 2).RowHeight = 15
        ws.Rows(s + 3).RowHeight = 15
        ws.Rows(s + 4).RowHeight = 6
    Next ti
    ws.Rows(25).RowHeight = 8
    ws.Rows(26).RowHeight = 20
    For gi = 0 To 3
        ws.Rows(27 + gi).RowHeight = 16
    Next gi
    ws.Rows(31).RowHeight = 10

    Call AddButtonToSheet(ws, "btnGuideStart", "Go to Patient & Input  >", "GoToInput", 32, 2, 240, 30, RGB(21, 101, 192))

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells(1, 1).Select
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' Ensure the Tebra Template sheet exists and sits at the very end of the workbook.
Private Function EnsureTebraSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TEBRA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SH_TEBRA
    End If
    On Error Resume Next
    ws.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    On Error GoTo 0
    Set EnsureTebraSheet = ws
End Function

' One line of the pasteable Tebra note (col A; text overflows to the right visually).
Private Sub TebraLine(ws As Worksheet, ByRef r As Long, ByVal txt As String, ByVal bold As Boolean)
    With ws.Cells(r, 1)
        .Value = txt
        .Font.Bold = bold
        If bold Then
            .Font.Size = 12
            .Font.Color = RGB(0, 121, 107)
        Else
            .Font.Size = 11
            .Font.Color = RGB(38, 50, 56)
        End If
    End With
    ws.Rows(r).RowHeight = 15
    r = r + 1
End Sub

' Append a value to a comma-separated list (for the Tebra med-line detail tail).
Private Function TebraAppend(ByVal cur As String, ByVal addv As String) As String
    If cur <> "" Then cur = cur & ", "
    TebraAppend = cur & addv
End Function

' Format one medication as a Tebra note line from a LOG row (blank fields omitted):
'   Name Strength Form - Directions  (Qty X, Y refills, Exp MM/YYYY, Lot ####)
' Log columns: 4 name, 5 strength, 6 directions, 7 qty, 8 refills, 9 exp, 10 lot, 13 form.
Private Function TebraLogMedLine(wsLg As Worksheet, ByVal lg As Long) As String
    Dim s As String
    s = Trim(wsLg.Cells(lg, LG_NAME).Value & " " & wsLg.Cells(lg, LG_STR).Value)
    If Trim(wsLg.Cells(lg, LG_FORM).Value) <> "" Then s = Trim(s & " " & Trim(wsLg.Cells(lg, LG_FORM).Value))
    If Trim(wsLg.Cells(lg, LG_SIG).Value) <> "" Then s = s & " - " & Trim(wsLg.Cells(lg, LG_SIG).Value)
    Dim tail As String
    tail = ""
    If Trim(wsLg.Cells(lg, LG_QTY).Value) <> "" Then tail = TebraAppend(tail, "Qty " & Trim(wsLg.Cells(lg, LG_QTY).Value))
    If Trim(wsLg.Cells(lg, LG_REF).Value) <> "" Then tail = TebraAppend(tail, Trim(wsLg.Cells(lg, LG_REF).Value) & " refills")
    If Trim(wsLg.Cells(lg, LG_EXP).Value) <> "" Then tail = TebraAppend(tail, "Exp " & Trim(wsLg.Cells(lg, LG_EXP).Value))
    If Trim(wsLg.Cells(lg, LG_LOT).Value) <> "" Then tail = TebraAppend(tail, "Lot " & Trim(wsLg.Cells(lg, LG_LOT).Value))
    If tail <> "" Then s = s & "  (" & tail & ")"
    TebraLogMedLine = s
End Function

' Section header + one line per THIS patient's logged meds whose Source (Log col 11)
' matches srcKeys (pipe-delimited; blank counts as IN HOUSE). Duplicate lines removed.
Private Sub TebraLogSection(ws As Worksheet, ByRef r As Long, wsLg As Worksheet, ByVal pn As String, ByVal pdob As String, ByVal header As String, ByVal srcKeys As String)
    Call TebraLine(ws, r, header, True)
    Dim lastLog As Long, lg As Long, cnt As Integer, src As String, dupe As String, ln As String, dk As String
    lastLog = wsLg.Cells(wsLg.Rows.Count, 1).End(xlUp).Row
    cnt = 0
    dupe = "|"
    For lg = LOG_HDR_ROWS + 1 To lastLog
        If UCase(Trim(wsLg.Cells(lg, LG_PT).Value)) = UCase(pn) And UCase(Trim(wsLg.Cells(lg, LG_DOB).Value)) = UCase(pdob) Then
            src = UCase(Trim(wsLg.Cells(lg, LG_SRC).Value))
            If src = "" Then src = "IN HOUSE"
            If InStr("|" & srcKeys & "|", "|" & src & "|") > 0 Then
                ln = TebraLogMedLine(wsLg, lg)
                dk = UCase(ln)
                If InStr(dupe, "|" & dk & "|") = 0 Then
                    dupe = dupe & dk & "|"
                    Call TebraLine(ws, r, ln, False)
                    cnt = cnt + 1
                End If
            End If
        End If
    Next lg
    If cnt = 0 Then Call TebraLine(ws, r, "[copy & paste from weekly med list]", False)
    Call TebraLine(ws, r, "", False)
End Sub

' A patient sub-header band: name on the left, DOB on the right (light teal).
Private Sub TebraPatientHeader(ws As Worksheet, ByRef r As Long, ByVal pn As String, ByVal pdob As String)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(224, 242, 235)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Merge
    With ws.Cells(r, 1)
        .Value = "  " & IIf(pn <> "", pn, "[Patient]")
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(15, 110, 86)
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 8)).Merge
    With ws.Cells(r, 5)
        .Value = "DOB  " & IIf(pdob <> "", pdob, "--") & "  "
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(15, 110, 86)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 20
    r = r + 1
End Sub

' One patient's full pasteable Tebra note block (from the session Log).
Private Sub TebraPatientBlock(ws As Worksheet, ByRef r As Long, wsLg As Worksheet, ByVal pn As String, ByVal pdob As String)
    Call TebraPatientHeader(ws, r, pn, pdob)
    Call TebraLine(ws, r, "", False)
    Call TebraLine(ws, r, "Medication Reconciliation", True)
    Call TebraLine(ws, r, "Medication reconciliation completed by [Clinical Pharmacy Student name] on " & Format(Date, "MM/DD/YYYY") & ".", False)
    Call TebraLine(ws, r, "Medications added:", False)
    Call TebraLine(ws, r, "Medications modified:", False)
    Call TebraLine(ws, r, "Medications removed:", False)
    Call TebraLine(ws, r, "Refills needed:", False)
    Call TebraLine(ws, r, "", False)
    Call TebraLogSection(ws, r, wsLg, pn, pdob, "Medications Dispensed in Clinic:", "IN HOUSE")
    Call TebraLogSection(ws, r, wsLg, pn, pdob, "DOH & Outside Pharmacy Medications Prescribed:", "DOH|OTHER")
    Call TebraLogSection(ws, r, wsLg, pn, pdob, "RxAPs Medications Prescribed:", "RXAPS")
    Call TebraLine(ws, r, "Medication Counseling Note", True)
    Call TebraLine(ws, r, "", False)
    Call TebraLine(ws, r, "", False)
End Sub

' Build/refresh the Tebra Template sheet: patient name + DOB on the right, the pasteable
' reconciliation/counseling note, with "Medications Dispensed in Clinic" filled from the
' current patient's medications. Wired to a button and rebuilt each Setup.
Public Sub FillTebraTemplate()
    On Error Resume Next
    Dim ws As Worksheet, wsLg As Worksheet
    Set ws = EnsureTebraSheet()
    Set wsLg = ThisWorkbook.Sheets(SH_LOG)

    Application.ScreenUpdating = False
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    Dim ci As Integer
    For ci = 1 To 8
        ws.Columns(ci).ColumnWidth = 14
    Next ci
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' Session header band: title on the left, session date on the right.
    ws.Range("A1:H2").Interior.Color = RGB(0, 121, 107)
    ws.Range("A1:D1").Merge
    With ws.Cells(1, 1)
        .Value = "   Tebra Session Notes"
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 16
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A2:D2").Merge
    With ws.Cells(2, 1)
        .Value = "   One note per patient - copy each into that patient's chart"
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    ws.Range("E1:H2").Merge
    With ws.Cells(1, 5)
        .Value = "Session  " & Format(Date, "MM/DD/YYYY") & "   "
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 26
    ws.Rows(2).RowHeight = 18
    ws.Rows(3).RowHeight = 8

    Call AddButtonToSheet(ws, "btnTebraFill", "Refresh from session log", "FillTebraTemplate", 4, 1, 230, 24, RGB(21, 101, 192))
    ws.Rows(4).RowHeight = 30
    ws.Cells(5, 1).Value = "Each patient below is a full note - select its lines and copy them into that patient's Tebra chart."
    ws.Cells(5, 1).Font.Italic = True
    ws.Cells(5, 1).Font.Color = RGB(120, 120, 120)
    ws.Cells(5, 1).Font.Size = 9
    ws.Rows(5).RowHeight = 14

    Dim r As Long: r = 7
    Dim lastLog As Long
    lastLog = wsLg.Cells(wsLg.Rows.Count, 1).End(xlUp).Row
    If lastLog <= LOG_HDR_ROWS Then
        Call TebraLine(ws, r, "[no medications have been printed / dispensed yet this session]", False)
    Else
        Dim seen As String, lg As Long, pn As String, pdob As String, key As String
        seen = "|"
        For lg = LOG_HDR_ROWS + 1 To lastLog
            pn = Trim(wsLg.Cells(lg, LG_PT).Value)
            pdob = Trim(wsLg.Cells(lg, LG_DOB).Value)
            If pn <> "" Then
                key = UCase(pn & "~" & pdob)
                If InStr(seen, "|" & key & "|") = 0 Then
                    seen = seen & key & "|"
                    Call TebraPatientBlock(ws, r, wsLg, pn, pdob)
                End If
            End If
        Next lg
    End If

    ws.Cells(1, 1).Select
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

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
    Call AddButtonToSheet(ws1, "btnNewPt",  "Start NEW Patient",   "StartNewPatient",    53, 2, 230, 28, RGB(0, 121, 107))
    Call BuildLabelPreviewLayout(ws3)
    Call PreviewAllLabels
    Call AddButtonToSheet(ws3, "btnUpd",    "Update Label Preview", "UpdateLabelPreviewFromSelection", 20, 1, 220, 24, RGB(21, 101, 192))
    Call AddButtonToSheet(ws3, "btnPrint",  "Print This Label",    "PrintLabel",         23, 1, 220, 24, RGB(46, 125, 50))
    On Error Resume Next
    ws3.Shapes("btnUpd").DrawingObject.PrintObject = False
    ws3.Shapes("btnPrint").DrawingObject.PrintObject = False
    On Error GoTo 0
    Call AddButtonToSheet(ws2, "btnAddMed", "+ Add Medication",   "AddMedicationRow",         1, 19, 150, 22, RGB(46, 125, 50))
    Call AddButtonToSheet(ws2, "btnRemMed", "- Remove Selected",  "RemoveSelectedMedication", 3, 19, 150, 22, RGB(191, 54, 12))
    Call AddButtonToSheet(ws2, "btnRevMed", "Review & Validate",  "ReviewMedications",        5, 19, 150, 22, RGB(21, 101, 192))
    Call AddButtonToSheet(ws2, "btnPrvAll", "Preview ALL Labels",   "PreviewAllLabels",         7, 19, 150, 22, RGB(0, 121, 107))
    Call AddButtonToSheet(ws2, "btnPrnChk", "Print Checked Labels", "PrintCheckedLabels",       9, 19, 190, 30, RGB(216, 67, 21))
    Call AddButtonToSheet(ws2, "btnEditEnc", "Edit Past Encounter",  "EditEncounter",           12, 19, 190, 22, RGB(84, 110, 122))
    Call AddButtonToSheet(ws2, "btnSaveEnc", "Save Edited Encounter", "SaveEditedEncounter",    14, 19, 190, 22, RGB(0, 121, 107))
    Call EncStore   ' make sure the hidden encounter-snapshot sheet exists
    ' Removed the single "Print Selected Label" button - printing is now via Print Checked Labels
    On Error Resume Next
    ws2.Shapes("btnPrnMed").Delete
    On Error GoTo 0
    ' Medications header row (row 2), written by code so the column ORDER is authoritative
    ' regardless of the template: Source sits right of Lot #, Refills right of Date of Rx.
    ' Every column whose position the reorder changed is (re)labelled here so headers can
    ' never drift from the data.
    ws2.Cells(2, C_EXP).Value = "Expiration"
    ws2.Cells(2, C_LOT).Value = "Lot #"
    ws2.Cells(2, C_SRC).Value = "Source"
    ws2.Cells(2, C_DATE).Value = "Date of Rx"
    ws2.Cells(2, C_REF).Value = "Refills"
    ws2.Cells(2, C_CONF).Value = "Confidence"
    ws2.Cells(2, C_WARN).Value = "Warnings"
    ws2.Cells(2, C_RAW).Value = "Raw text"
    ws2.Cells(2, C_PRTD).Value = "Printed?"
    ws2.Cells(2, C_CNT).Value = "# of Prints"
    ws2.Cells(2, C_SEL).Value = "Print?"
    Dim hc As Variant
    For Each hc In Array(C_EXP, C_LOT, C_SRC, C_DATE, C_REF, C_CONF, C_WARN, C_RAW, C_PRTD, C_CNT, C_SEL)
        Call MatchHeaderFormat(ws2.Cells(2, C_QTY), ws2.Cells(2, CLng(hc)))
    Next hc
    ws2.Columns(C_EXP).NumberFormat = "@"    ' keep Expiration as text
    ws2.Columns(C_LOT).NumberFormat = "@"    ' keep Lot as text
    ws2.Columns(C_EXP).ColumnWidth = 10
    ws2.Columns(C_LOT).ColumnWidth = 11
    ws2.Columns(C_SRC).ColumnWidth = 12
    ws2.Columns(C_DATE).ColumnWidth = 12
    ws2.Columns(C_REF).ColumnWidth = 8
    ws2.Columns(C_CNT).ColumnWidth = 7
    ws2.Columns(C_SEL).ColumnWidth = 8
    ws2.Columns(C_CONF).ColumnWidth = 11
    ws2.Columns(C_WARN).ColumnWidth = 26

    ' --- Spruce up the Medications header area (code-controlled = resilient regardless of
    ' the template): a blue title banner and the header row both span the FULL table (A:Q),
    ' so nothing hangs off the right after the column reorder. The two internal columns
    ' (Raw text, Printed?) are hidden so the volunteer view stays clean.
    Dim medTitle As String
    On Error Resume Next
    ws2.Rows(1).UnMerge
    On Error GoTo 0
    medTitle = Trim(CStr(ws2.Cells(1, 1).Value))
    If medTitle = "" Then medTitle = "MEDICATIONS"
    ws2.Range(ws2.Cells(1, 1), ws2.Cells(1, C_SEL)).ClearContents
    With ws2.Range(ws2.Cells(1, 1), ws2.Cells(1, C_SEL))
        .Merge
        .Interior.Color = RGB(21, 101, 192)     ' app blue, matches the buttons
        .Font.Name = "Arial"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws2.Cells(1, 1).Value = medTitle
    ws2.Rows(1).RowHeight = 26
    ' Header row (row 2): uniform, wrapped, centered, with a divider under it
    With ws2.Range(ws2.Cells(2, 1), ws2.Cells(2, C_SEL))
        .Font.Name = "Arial"
        .Font.Bold = True
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws2.Rows(2).RowHeight = 30
    ' Light-gray grid across the data area so the whole table reads as a set of boxes,
    ' outlined all the way to the right edge (Print?), even before any meds are entered.
    With ws2.Range(ws2.Cells(MEDS_HDR_ROWS + 1, 1), ws2.Cells(MEDS_HDR_ROWS + 40, C_SEL)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(208, 208, 208)
    End With
    ws2.Columns(C_RAW).Hidden = True            ' internal: original parse text
    ws2.Columns(C_PRTD).Hidden = True           ' internal: superseded by "# of Prints"

    Call ApplySourceValidation(ws2)
    ' Re-point the double-click Print? handler to the current columns (it hardcodes the
    ' # of Prints / Print? column numbers, which the reorder shifted). Runs at build time
    ' only (Build-Release has VBA-project trust); volunteers get the baked-in handler.
    Call InstallMedSheetEvents(ws2)
    Call ApplyAllRowStates(ws2)

    ' Dispense Log header row written entirely by code (via LG_* constants) so the column
    ' order is authoritative: Encounter sits right after Timestamp, then patient + med fields.
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)
    wsLog.Cells(2, LG_TIME).Value = "Timestamp"
    wsLog.Cells(2, LG_ENC).Value = "Encounter"
    wsLog.Cells(2, LG_PT).Value = "Patient"
    wsLog.Cells(2, LG_DOB).Value = "DOB"
    wsLog.Cells(2, LG_NAME).Value = "Medication"
    wsLog.Cells(2, LG_STR).Value = "Strength"
    wsLog.Cells(2, LG_SIG).Value = "Directions"
    wsLog.Cells(2, LG_QTY).Value = "Qty"
    wsLog.Cells(2, LG_REF).Value = "Refills"
    wsLog.Cells(2, LG_EXP).Value = "Expiration"
    wsLog.Cells(2, LG_LOT).Value = "Lot #"
    wsLog.Cells(2, LG_SRC).Value = "Source"
    wsLog.Cells(2, LG_DATE).Value = "Rx Date"
    wsLog.Cells(2, LG_INIT).Value = "Initials"
    wsLog.Cells(2, LG_FORM).Value = "Dosage Form"
    wsLog.Cells(2, LG_CNT).Value = "Print #"
    Dim lh As Long
    For lh = 2 To LG_LAST
        Call MatchHeaderFormat(wsLog.Cells(2, LG_TIME), wsLog.Cells(2, lh))
    Next lh
    wsLog.Columns(LG_ENC).ColumnWidth = 10   ' Encounter
    wsLog.Columns(LG_SRC).ColumnWidth = 12   ' Source
    wsLog.Columns(LG_DATE).ColumnWidth = 12  ' Rx Date
    wsLog.Columns(LG_INIT).ColumnWidth = 8   ' Initials
    wsLog.Columns(LG_FORM).ColumnWidth = 13  ' Dosage Form
    wsLog.Columns(LG_CNT).ColumnWidth = 7    ' Print #
    wsLog.Cells(2, LG_ENC).HorizontalAlignment = xlCenter

    ' Default Date of Rx to today
    If Trim(ws1.Range("C7").Value) = "" Then
        ws1.Range("C7").Value = Format(Now(), "MM/DD/YYYY")
    End If

    ' Consolidate previews: migrate/rename the gallery and hide the internal print sheet
    Dim wsGallery As Worksheet
    Set wsGallery = EnsureAllLabelsSheet()

    ' Build the Start Here guide as the first sheet (the workbook opens to it).
    Call BuildQuickStartSheet

    ' Build the Tebra Template sheet at the end (auto-fills Dispensed from current meds).
    Call FillTebraTemplate

    ' Ship a clean sheet: wipe any leftover/test medication rows so a freshly built
    ' workbook never opens with stale data in the Medications tab.
    Call ClearMedArea(ws2)

    ' Version stamp so the loaded build is visible at a glance (support / troubleshooting).
    On Error Resume Next
    ws1.Cells(58, 2).Value = "SCU Label Tool  v" & APP_VERSION & "   -   built " & Format(Date, "YYYY-MM-DD")
    ws1.Cells(58, 2).Font.Size = 8
    ws1.Cells(58, 2).Font.Color = RGB(150, 150, 150)
    On Error GoTo 0

    On Error Resume Next
    ThisWorkbook.Sheets(SH_GUIDE).Activate    ' end on the Start Here guide (opens here)
    ws3.Visible = xlSheetHidden
    On Error GoTo 0

    MsgBox "Setup complete!  (v" & APP_VERSION & ")" & vbCrLf & _
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
Private Function LogoB64() As String
    Dim b As String
    b = "iVBORw0KGgoAAAANSUhEUgAAAaQAAALQCAYAAADfFwR/AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAA"
    b = b & "DsMAAA7DAcdvqGQAADCFSURBVHhe7d1pk1vHlaBh/P+YmOgOy5a7Z6x98XR7k2V321/my/yxideFdKEuUWQtwF2fJ+IEKYmk"
    b = b & "iiggz83MkydPJwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    b = b & "AAAAAAC44n+dTqffnE6nfz9H//y/z/+uH8e/G/9+/Hz89/Frr/2aot8/fu347+PP/Hr6xQCwD19cJIeRAC4TyjRh9Ou/OZ1O"
    b = b & "P55Op/88nU5/OJ1Ofz6dTj+f4y+n0+mv5/ivi/jvZ+Ly14zf158x/rz+7P4f/3E6nb4/nU5fXvmaxtff19a/70cAVmQklzEr"
    b = b & "uUwyI0ouJZafzolgJI9+HImhpPCnc2L4/el0+t359xQliuL/3CnGn9//q/9v//++jr6evq6+xsuvt//23UXSupxlSVQAd9SM"
    b = b & "YQy2lzOH8e8azBu4m3k0cPdjSaYE1OA9kss9k8qc0d+jv9MfJ3/v/s79t5Gket2aWRUAvEJ7JyWYBtLL5NPPG4AbcMdMoZ83"
    b = b & "g2g2sadk8964XGIcy4T980hSl68vAOen9jHzKfrnBskSzHjib3mqWc5IOtPBV7wsStaXSarXtf2x8Zr3fWhZE2D3mvlcVp8V"
    b = b & "7YGUbMZG/0g8ZjrzRAm+GeZ4/X87SVAq/YBdqMiggW1UtzW4lXDG/k4DYUtwks964jJBNYP69uL7Z3kP2IwxaI3ZT0/boyS6"
    b = b & "ZSIzn+1FDwx973qIqHhifI97uGgGBbAKY1Aas6AGsJGAesq237Ov6GGipDRmT2PmZGkPWMSo0BpPyVW5NUC1FCcBHSua8Y7D"
    b = b & "wRVDSE7A3VV2PWZCLeGMsy4NSNNBShwzRnLq4WQUr1jSA26iwWQkoUqFS0JmQeIlMUr2x7LeeC8BvNhluW+DykhCChHEW6OH"
    b = b & "mVHU0nvrKx0jgI9pkGiw6ExQSajBw0xI3DJ6qKnQpfdXPzdrAv7p8mm1/aCWWNofmg4kQtw6etgZxTCjQKYADmYsydUypqfV"
    b = b & "nlqnA4YQc8XYn+xhaOxZAjs3nkJHV2izIbGmGLOm3ptj9v5v0zcxsF2Va/fB7slzbCwrUBBrjw7fNmvqksIxowc2alTL1Yds"
    b = b & "tHyZfuiFWHs0ix9dPyznwcaMw4j1kCsRObgq9hDN6i+X88yYYMXG/lAf3J4o7Q+Jvca4qqSkVEeIWlkBKzBa+oxE5OyQOEqM"
    b = b & "faaxlGc5DxbSPUNjac6MSBw5Rtn4uLdJZR7MqEQ0ihUkIiEeYrS6qjKvz4jEBHc0NnMVKwjxfFwmpj4zv55+kIC3G/tEuioI"
    b = b & "8fIYiWnc02SPCd5hHGrtMGslr9MPnBDi0zESUw91ysXhDfrgjM7bOisI8f4YVXmjsz3wCZfLc1UPVUU3/WAJId4el+eYnGGC"
    b = b & "Z/TU1lOc5Tkh7h99zkpOzZh6EAQu9or+5mCrELPGOFD+g2U8eFg2qGihJzXLc0IsE6PwwfXqHFJr12ZFQqwrRkVrD4r/Ov3Q"
    b = b & "wh61NNB5ot78ZkVCrCtaxhtl4mZL7FatTHqTNyvS8keI9UYPii2j/0VSYo9aomvjtDf49M0vhFhntJzeA2QPkloQsQu9mbtY"
    b = b & "zLkiIbYZVeK15yspsVm/ulii021BiO3GWMJrqd1hWjans0XdZtnMaPrmFkJsL0pKlYc3W3Jmic1oE7SnqSrpLNEJsa+43FeC"
    b = b & "VSsZuThPiP3HOEj7+XQQgKV1iM5+kRDHiVY/qprt9ma307Iao3ihteXpm1YIsd8oKbU0X1NkDVpZXE9GFS/0pGS/SIhjhmIH"
    b = b & "Fvfv5+U5jVGFEI0FLdlLSsyu6fmYqktGQogRow/eZ9NBA+6hmVEdgZumS0ZCiMsYxQ46O3B3zYw67Foymr4RhRCiGJ0dGicU"
    b = b & "O3AXtQxxxkgI8ZIoKdW/sgdY7Ya4qZ5ySkYu0xPittEDXvuxzSgawPf2GdPZgZtqz6inHDMjIW4TzR76PFUq3bUsPfCNaADv"
    b = b & "8zY65O/loLlrLHi3zhlJRkLcLsb+Sj9vKWu6nNUDYIlp/LduVt7Lvu24xqK/I7xKb5o+DHv4IAixhhjJ6DUVaP3aZha16GnZ"
    b = b & "vKMW0z93K9HfvzGlHxU78GLNjEpELRlM31RCiLdFn6mW4Gq39Vo1Li4xlaCaaWz1wstR7FBikpT4pF+eTqcfXR8hxM2jRHKL"
    b = b & "RqQjMTVj2uoeU9sAff06O/CslhF6s49p9fRNJIR4W/SAd8sZwZgxdQh1qx1TRruh/h7dGABP9CZvA3WLb24h1hzNjm6ZkIYK"
    b = b & "H9pfqnvKVj+3vTYlpbcsZbJT46zRVt/UQqw1Wp6qvHtaUXcrLQM2oG/1YbKvuYT6/Y2WNNm43gRNnadvFCHE+2OODfxubh33"
    b = b & "km01KbX0WBXivRI3G1Ayah16byfEhVhL3Gu5bmokpS3fTzaKHdo+4GAqYmgpYctnG4RYe8yVkDIuzayIYvp1bCUuix1cY3Eg"
    b = b & "PYVs+WlKiLXHuMRyzmWocah966seo7PDSw8Rs2G9aXsKkYyEuF+UkDoQO/d5mz7fDejTr2dLMYodWsXRbmjHxpt1q4fqhNhK"
    b = b & "jIQ05wwpo93Q1lt/jWKHZnz2lXaojc/eqPaNhLh/jIS0xGC6h1nSiFHs4BqLnWlzdauloUJsMVp2mnuGlJJgA/nW95JGXBY7"
    b = b & "/GL6l2V7xr7R9BsthLhfzFllN9VnfssVd9fCNRY7UFuOZkVbX1MWYmtRJetSCan/bx0cpl/TlqNxrMrFZn9LLIVyA+PmV0t1"
    b = b & "QswbXbew1NP8nvaRLmM8XPd3k5Q2psNybQhOv6lCiPvHOIu0xMA59oynX9Ne4nJfqatzWLnekJ3cdtmeEMvFUvtIe50hTaPV"
    b = b & "n8a5JV5jXqE3pC7eQiwbHbNYYtmuAbozPNOvZ2/R+FbxxlIzUV6gb8weWogIsYfoKX7OJ/hKzb85b/5Pv5a9Rn/XsYTnfqWV"
    b = b & "6YmsMxDTb5oQYv7o4XDOWdJRluuuRatCLeG5X2kl+kYc9c0oxFqjz+Rcg2QzpKN2ZBml4f39LeGtQG/GvR2IE++PqpI+FtNf"
    b = b & "L24bY4C897UKR54dXYaWQyvQOrWODPuJEkX7gH24qpbsQaPln5ZjO3TZwNOHrh9HtF9R9N8vo9/X778W/bfx+y7/rP4//ffO"
    b = b & "fUha74ue3HtN79lKqOsa+p7ZO36I3rN9PuruMNfslAs9gR11qr7V6ENTwhmdjS8TwihM+e78sNFg1o8jxj/3VNzP+3HE+Pfj"
    b = b & "5x+Lfu/lnz3+vBF9oPv6RsLr5xLU2+JeSanvWw8Ojnk8jXGVxY/n14iZNIiYHa07GsQbNC4TT+vd358fJvrA9GPfy/HPI1Es"
    b = b & "qYTU0kd3+4yk1cDX02czKMnp5TGe2m95T9JYpu974pjHhzGusmil4B4PA0xUVdKA4eloXdHspg/CWAob+wgjyTQo9fMtLif0"
    b = b & "9xhJqgTbh91S0ctj7G+8Z4Ds9xf9WUcq8X5r9P4cpeHckdnReqIZUIPzGCTGLGcM4Hv8MIzk2gNRf/fpayKux9jX67V7zYxp"
    b = b & "/Hqv99ti7CtpOXQnDXj2jpaLBoZRBDBmPcVrBpk96O9bwu21sHz0smgJr6TSe6eHlrFUey3GfzMjfV/0vmzVQrHDHTTwaaA6"
    b = b & "fzQTGtVsl7Oflk+PrsRUIUaDrKT08iiJ95pNKybHkm/vtX6NPbv3R+/LEnvFDpLSDbVc18b49AUXt48Ggl7rHgBqzTI2+/nQ"
    b = b & "5+fXp4FUUhJrjNEHr4fLOTtp7FYvojXk+0dLI+MsT0//DbROgX9a52J6vSQlsdYYFXg9aEpK79R6soR0vygRjf2QZkIlIl6n"
    b = b & "jeNeO8t3Yq3R+3IcPJeU3qh1zwZMG5u3j5bmWrdvKj/2h3i73qs/KLwRK4+SkpnSGzU7Usxw+2hZbizNWZa7nZKSyjCx5hjL"
    b = b & "d71HW27mFUpInjhvE70Rmw2N8wkS0X30nnVeTqw5GguaJbmF9hV6oVpSmr6Y4m3Ra1n7Hktz99cxBUUOYs0xzim9p4vGobTG"
    b = b & "KSG9Ly5nRSUi68bzaPY52g1JSmLN8Xfjwqe1pNQZmAbT6QsoXhbjUFzFC5bn5td+Ut8HvRfFmqO9pKpDJaWPGEse0xdPvDzG"
    b = b & "XlGHN1nGKHLQFFSsOdpPsnT3Ec4evT2aEen2ux49efZwoBWOWGu0mtKYocDhil+dl+s8Vb4+es2aWR6t4ena6VQv1h4tLRs3"
    b = b & "rujDa7nu9TGuhDD1Xqe+L5KSWHM0kzdLmlDu/fooGdmYXL+KSxz0FmuNiqCMIRMlJJVJL4+W6SSj7WhZpAcu5eBibTE6jLiq"
    b = b & "4qxljSo+bAC/LEZjVNPs7WiPVFISaw3jyQW9614X1ny36RfnKkhJSawtJKQLvRAtP01fJPE0RpmmAobtkpTEGqP3o4R03vDt"
    b = b & "IKfuDB+PBq9xpTjbVrflkpKqUrGWKCEZW86b8qPVzfRFEo9RwcePWsfvxmfnpGSpWqwhzJDO7B99OkrWPU2rqNufkpJzSmLp"
    b = b & "sId0pl3Qx6OlOslo31oqKSm54E8sFc2QlH2fB1qX8T0fLdXVUkkhw76VlHrwsJcq5o7ONP4gIT0Msi1HeTJ8PlrONDs6hj4P"
    b = b & "PZx1Jk8FnpgrHLA/60WwXPd8VOxhXfdY+kx0xbQKPDFXONd41ovgg3c9mjn25OKNcjwtnYxiB6sH4p6h2/cFBQ3PR6+LafSx"
    b = b & "NVD0PviTJTxxp3Af0oVeiD5s0xfp6NHsSAde0r5SB8etJIhbR/uViqXOevprWUJV0Ydhk5FLvRf6vLTW3xKL2ZK4Rdg7utAH"
    b = b & "rGSkQ8OH4dQ011Qa/t15tuRzI94a42yjMeZCU8VmAj5YT6MlTG8UnjMKHtpbqjx8+v4R4lPRUt232pA9VULSMujDMI3mJZot"
    b = b & "jV54DpaLl0ZVmwqmrigh2ah9Gr1ZCgmJl2rpu6IHiUm8JP4uGV2n5PvDUFnHW4zrLMaMSfd8cS3cpfYR49qJ6Yt25LBcx3uN"
    b = b & "pbz2IluBaNYkOR07KmJobHHf0UeMvl3TF++o0aBR1eHhGxxyEy3jlZha0uvBr+RUEYTOD8eLklHvBT6ihOQM0mOUnC3XcQ/1"
    b = b & "xisx9ZnrabljBSWoElXnmiSp/cZIRr+cvil41LJUZYe1PZ++gEeNBgkJiXsrKbV002ewnxd9DjuCUZIq2tstWbXsV8LqwbFf"
    b = b & "M4pums2/JMavf0lMf+/08yFeF72G7RnpU/cCLUuNN+z0hTxqOKjGUsbsafxYjITVjz1hd1/O+MyWoEpUn4prieZalOzG72lZ"
    b = b & "sWR4mSCfixJnv65f3+9rlWGaQI+Y4Pq7V00nGb1Qb3KbrY/R69CHSEJijXqAbPbezGr8vPdqPz4X/fdP/ZqiQbNfN/7ckuD4"
    b = b & "/WMmN2KaLMfPp//cn/n9+XNVgiphjWXKEf1zyWwksssENv18bin6O/X3UcDwCr3heiNs/Zt/qygZ9WEC3q8y+JHsprO9afIa"
    b = b & "P+9W5pLSKABp9lUym35W1xqNIS3RVczS34tX6M2ibdBjuIgPltWYNGZklz82kyrWWoDVrK7ChZZULdG9Ud9sh2IfQzNVWJ8+"
    b = b & "k6N0vh/XlJx6mB93ZfW16Uv3Dn2j+8ZOX+SjRk84luxgvSqdvzzbNYou5t4Lb0Y0qiD7muwV3UAJqW/m9MU+avQGM0OCbWg2"
    b = b & "MmZNJYT2mnqoLEnd60qd/tyRAEdS5EaqqNHp+yF681bgISHBNpWUilY5ShbNXkaCGuXv08/9S+LyfNjl7IwbawNRQnqIcTDQ"
    b = b & "1Bu27/Is14ixzNaY11bFtUPH/djspwTUr+vX1zygP0/l3J1JSI/Rm7E3Xm86YH/a62l2M01UJZk+95WcjxnWqO4rEXlInYmE"
    b = b & "9Bg9GamSgWNSmLACPQVISA/RtN10HGAhihoeo3VjCQlgIa2VOof0EBISwIIcjH0MCQlgQXrZPYaEBLCgBuBq8SUkCQlgUe0h"
    b = b & "uX7iIVTZASyoAXjupoRrjXEOqQBgZqOdxlt7PO0pRqcGPaoAFtAM6bvzYDwdoI8WIylX6AHAzEpIzQjWcNHV0tGyZY0W7SMB"
    b = b & "LKQBeEt31t8zKoF3QR/AQkpIVdpNB+cjRvedmCEBLKQBuJnBdHA+YtTXzx4SwEL0s3uMLuOyZAewkGZILVVNB+cjhivMARbU"
    b = b & "AOwKioeouMMMCWAh3ZKowepDjEPCZkkACyghNTOQkB6ifSQJCWABvz63zNGt4SGUfgMsaDRZnQ7OR4yWLyUkgIU0AHcv0nRw"
    b = b & "PmK4FwlgQQ3AziI9REUN7ac5IAuwgAZfpd+P0T6S8m+ABWgf9DQUNgAspNLvBmIX9T1E11BYsgNYSAOwaygeYiRmy3YAC2jw"
    b = b & "VWn3GJbtABZSQtJk9TGqOrRsB7AA11A8jZYvzZAAFtBsoPZBeto9RK+DK80BFtKMQGHDY3Q2yywJYAHNBup2PR2Yjxq1EbKP"
    b = b & "BLAAl/U9jZbtqjw0SwKY2Wiyah/pMUrQ9pEAZlZC+vJ0Ov3+ysB81ND9G2Ah+tp9GL0e9pIAZmYf6cPQtQFgAc0EfrKP9CRa"
    b = b & "wvzaLAlgfhqtfhiKGwAWoI3Qh1H1oRkSwMwaeDVafRotYXZPUndHATCTNvB/dGHfB1GSNksCmFkDr/uRnkbNZ7+3lwQwL+Xf"
    b = b & "10PDVYCZNQuo3Nmy3dPQuQFgAeNM0nRQPnq0l2TZDmBGDbqW7T6MkrRZEsCMKnG2bHc9tBMCmNm4kmI6IB89HJQFmJlqu+ej"
    b = b & "bhZmSQAzaR+pQ7KdwZkOyEcPsySAmTUL+MuVAVk4lwQwqwZczVavRxV3SsABZtSg60qK69EsydIdwEwUNzwfJeovTqfTl9MX"
    b = b & "DYDbKyF1HsmZpA/jt2ZJAPNS3PB8VIVYKHAAmEEzAMUNz0fJWkICmImGqx+Plu7+bfqiAXB7ihs+HhU4fHU6nb6evnAA3F7V"
    b = b & "ZErAr0cFDhqvAszELOnj8R/npGTpDuDOupbiB/3tPhrNIOsBKCkB3FnFDWZJz0dLd1XdlbwBuKOW7Rp4HZT9eGi+CjCDZknt"
    b = b & "lUwHYfEYJew/6uIAcF8tR2kn9PFon+1bHcEB7s8s6ePxe+eSAOZRQmrgVXF3PUpIzSQt2QHMQMXd81H5t6IGgJn8+nQ6fXee"
    b = b & "DUwH5KOHhAQws2ZJf7syIB89/qSgAWB+7ZVU4jwdlI8cJSQzJICZ1SbHXtLT6KoOBQ0AC2g28OcrA/NRo4RkhgSwgPZLSkh1"
    b = b & "vJ4OzkeMn82QAJYzysBrMDodoI8WJSQzJICFVNxQGbhL/B46fktIAAuqwEEZ+On0VwkJYHntJ7VkNR2kjxSuMwdYgRJS53CO"
    b = b & "3A28hORgLMAKjA4ORy1wkJAAVqQih6OeTbJkB7Ai42zSEZfuSkjOIQGsyFGbr5ohAazQEZfuJCSAFepsUr3djrR0p6gBYKWO"
    b = b & "dmC2g7ESEsBKNUDXUmfvpeA1mNVcFWDFuvK8ZLT3XnclpJYozZAAVqxZQ/sre76mor9bN+gqagBYub13cah44/cSEsA2VAre"
    b = b & "Pssek1IJ6XeW7AC2odlDg/ce95NKSC3bKWoA2IhKwSuP3tv5pPH3sWQHsCEN2ns7n9Ry3bcSEsD2tNdS5d1e9pNKSF+fA4AN"
    b = b & "aa/li3O/uz0kpfbFSrIVbgCwMR2abTDv/M50gN9alJAs1wFsWEUOtRZqyWs6yG8pHIoF2IGW7/57450cJCSAndh65d2fJCSA"
    b = b & "/aggoJnSFoscaqwqIQHsSJV3HZzdWlKqWlDbIIAd+eXpdPpygz3v+nrNkAB25vPT6fTNeRlsK0lJQgLYqXGxX8UCW0hKLTNK"
    b = b & "SO/X7NjSJ7A6lYN3x1CHTteelCSk26iwpZZS7SXqnA6sSgdnW7pb+5UVDaIS0vuVkPp+dwSgvoBaMQGrUlKqiq3Z0jQRrCVK"
    b = b & "SJ7o36+k3jJtr2k/1sXjq9Pp9Nn0FwIspT2lBqe1JiUJ6TbaPyoRja4d/djZNHtLwKo04K81KVmyu41ew2bDl22kRnFL+3Tt"
    b = b & "LQGsQkmpEuu1JSUJ6TaaBU0T0ogxW2oJz2wJWIXLmdJaqu8kpNu4NkO6jL7f/fcKH8yWgFVo4BrVd2tIShLSbTTz6WFj+vpO"
    b = b & "4z/PlXjtLbW/CLCoqu9KSF39sHRSKiFZRnq/kvpLEtKI9pW+Vx4OrEFJqeWdpdsMSUi3UUIqyUxf3+ei73kPJP2eZksAi2rJ"
    b = b & "pqfk9haWSEolRN2+b6OEVOHC9DX+VPQ9aAnPmSVgcb86PyG33DN3UhoJyTmk9xu3B09f45dGM6Ua8/peAIvqybik1PLZdKC6"
    b = b & "Z5SQWjYyQ3q/t86QRowzSz0g2FcCFtXVFSWlOa9DLyFVXGEAfL+S+i0eKH53UYUHsKiRlJ47z3LLqAS5UH78frdKSCPGvlJL"
    b = b & "ugCLGVcZ3PsAbcmoxOcc0vv1Gt4yIRXtK5WU7CsBixrdo+950V8JqT9bQnq/eySkvjcdC+j75HsELKqzSpWF1wPvHkmp/Ypv"
    b = b & "7SHdRLOYWyekou97+3xVYUpKwKLaQ2jZ5j0VXM9FS4L92SU+3qc9pNccjH1t9PDQe8DDA7C4UezQ8s10sHpr9OTdU/cvp/8z"
    b = b & "Xq1Ecc+EVLTf93dJCViDBqKW727VA6+EZMP8Nu49Q7oMZeHAKjSj+e6VjTyfixKbfYnbeG0vu/dGy3ddY6HdELCoDtG29/Pe"
    b = b & "80q6NNzO3AmpqIjCNRbAKvSE3CD41ruVKik2Q7qNXsdbzFpfE33PS0q9DxygBRbXDOeHNw6GEtLt9H1of2/6Gt87Skp975sp"
    b = b & "SUrA4irbfksVnk7ft1NiXyIhFSMptYzbci7A4lq6Kcm89NK/fq0Z0m30OvZ6Tl/juaLv97jsz0MGsAoNjA1KVWF9quChJ3oJ"
    b = b & "6TaWTkjFSEqq74DVGPcrNTh9rBeedjS3s4aEVFwWOgCsRgdpR9uha7OlEpaEdBstk5X8p6/xElFS0mYIWJ3LK9LbW7ocuHqS"
    b = b & "lpBuo9dxLQlpREUukhKwOi3hjL2lGnU2YPVzCek21piQipKS7zGwSiWlElIzJjOk2+l1rPPFNCEsHS3VlpRU3gGr1DJOien/"
    b = b & "SUg3s9aEVHQ2rf1CSQlYrZbxJKTb6HWshdM0Gawl+tq6/8rBWYAd6z6pWgc14E8TwVpiXIfezFhSAtixtSekYhycVXkHsFOd"
    b = b & "9WrmsfaENEKRA8BOtRf3zUU5/dqjyrtK/t2jBLAzzY62lJCKZnMVOkhKADvSDKk7qV5z9cfS0X5SvffaT6ooA4AdaD+mQX5L"
    b = b & "CakYPe/sJwHsRBV2P24wIY0oKXXJIwAbN0q+r3VU30KMr91MCWDjRpeGrSakcQV6iRWADdt6Qhrxd7MkgG1rZtHVE1tPSJWt"
    b = b & "l1jtJwFs1F4SUvGzpTuA7WoAr3HpHhJSobUQwEaVkDpkupeE1NJdfxdNWAE2pqKGPSWk0RXcXVkAG9PAXdn0XhLSiK64t3QH"
    b = b & "sCEtbTWjmA7oW48KNRQ4AGxIM6Q9JqSiAgdl4AAb0SyifnDTwXwPUYFDe0qW7gA2oBnEXhNSoSM4wEa0ZLfnhFQX8/aTJCWA"
    b = b & "lWvJroq06UC+p2iPTEICWLlmSHtPSJW011ZIgQPAih0hIRWdtTJLAlixoySkolmSpASwUkdKSM2SLNsBrFQzhqMkpPaS6tun"
    b = b & "gwPACu25U8O1UHEHsFJHS0idS/q9pASwPkdLSIXuDQArNK6fmA7ae45mSN9ISgDr0gb/0RJSYZYEsDLNkDqfMx2w9x4/SUgA"
    b = b & "63LUhDTaCbnqHGAljpqQis5fSUgAK9GA3PLVdLA+QlTc8K3uDQDrcOSEVChuAFiJoyckDVcBVqKE1I2q04H6KFHnhn6UlAAW"
    b = b & "1kB85IRUtGynuAFgYQ3Ef7wySB8p6gAuIQEsrE4Nf7gySB8pLNsBrICE9BAt27knCWBBEtJD6NoAsKDPTqfTF+cDotMB+mjR"
    b = b & "a/D1+fUAYGa/lpCeRF3PzZIAFvCb871Av7syOB8x9LYDWEj7R/Vyk5Aeoo4VEhLAApohfSch/TMq//6tpAQwvwbe7gUa53DE"
    b = b & "w7Kd8m+AmXXtgoT0NOwjASyghFQykpAeo30kMySAmbWHVB+7ZknTgfmo0X5ahR7OIwHMaDRWlZCeRs1WzZIAZtQMqcFXQnoa"
    b = b & "rqMAmFkJqTJne0hP468SEsD8unKhljnTQfnI4YAswAIaeCt1ng7KR456+9VSyT4SwMxKSq6geIz21Cr2qCwegBmZJX0Y7SO5"
    b = b & "QRZgZiWkLqibDspHDpV2AAuxbPc0VNoBLKTlKct2j9H5LAkJYAENvs0KpgPzUaOiBgkJYAGfn06nr11n/s8YPe0kJYAFKG54"
    b = b & "jLpXlJQ0WQVYQAmp6rLp4HzUsI8EsJB62/1Js9V/hrNIAAsqKdlHegi3xwIsyD7SY7R8qZ8dwEIagO0jPYTDsQALas/EeaSH"
    b = b & "6FoOCQlgIXW4rtzZpX0PVXaKGgAW8uvz2RuFDQ8X9dlDAlhQy1SVf08H6KNFr4ElO4AFNStQaSchASyuQVjn74frOOwhASxI"
    b = b & "QnqIEpIZEsCCJKSHkJAAFuayvoeo0rBWSr+cvkAAzKOiBgnpISF9eTqdPpu+QADMwwzpIcyQABZmD+kh7CEBLKxBWD87CQlg"
    b = b & "cQ3Ctc2ZDtBHCwkJYGENwnrZPXRq0MsOYCENwCUj15hrHQSwKBV2j6HbN8CCFDQ8RvchmSEBLKQZ0h+vDM5HDDfGAiyk22Lt"
    b = b & "Hz1GM0VLdgALaPD97ysD81Gj10JCAlhAy1Ptm0wH5qNGxR2W7ABmVs+2BuH/vDIwHzVKSC7oA5hZMwHLdU+jogZLdgAzayag"
    b = b & "XdBjVNhRtaGEBDCjUeqtuu4xxtJl9yEBMJMSkuW6p+EuJIAFtCxVZ+vpoHzk0McOYGaKGa7HzxISwLxarnP26MNwBglgRuMg"
    b = b & "rGKGD0NCAphRvet09v4wStDtIVXUAMCdtVTXwKszw4fxu9Pp9LUzSADzUOr9fHRA2HIdwAy+OJ1OP5xnAtPBWNg/ApiN2dHH"
    b = b & "Q0ICmEEb9d+dOxFMB2LxsKfWa9MsEoA7Mjv6eNTTz+wI4M6aHTUDsHf0fFQGLyEB3FkDrXNHHw/7RwB31iFYS3Ufj2aPvz2/"
    b = b & "VgDcQftGPypk+GTURslhWIA7apBtqa6n/+kgLB6jGWTJG4A7aID925XBVzyNcWW5cm+AO2gvpCahquo+Hb1OZkcAd9CTflFf"
    b = b & "tungKz6MlutU1wHcQftGlTDbN/p0VF3XFe5mSAA31lKdfaOXR9V1khHAjZWM/uKeo1eF6jqAGysZVS3W8tN00BXXo4KPms06"
    b = b & "DAtwI5+fTqcvz8tP9o1eHu2zmR0B3FCNUxUxvC6aTdo/ArihKur0qXt9SEYAN1QyUlH3tnD2COBGxmV7LT1NB1vx8XARH8CN"
    b = b & "lIxqmKot0NuiWaWEBPBOlSi3/+E6ibdFZfFVJBYAvNGvz8tNzhq9PewdAbxTM6OW6EpIyrvfFs0qv1JdB/B2HXwtCXVNgmT0"
    b = b & "9rB3BPAOktFtYuwdSUgAbzCSUfcaSUbvi2ZHluoA3qDBs67dZkbvj17DDhED8EolowoYJKPbhCsmAN6ggbNEpJruNtH9UJIR"
    b = b & "wCu14e7Q6+1iLHkqZAB4hfY4ukJCO6DbRLPLlupcvgfwCl+cB09Xj98uWvKszLtKRQBeoGRUSbKu3bcNhQwAL/Sr8xO8+4xu"
    b = b & "Gy3V1QndvhHAC9QktWTUntF0QBXviwpCfjQ7Avi0nty/P51OPyvrvkvoyADwAiWjypAdeL19jKU6HRkAPuE35yU6Z4zuEzVP"
    b = b & "/UFCAnjeZxfFCyrp7hO9riV7Z44AntEZmC6EqwR5OoiK28Q4AKuqDuAZbax/e+6lZr/oPtHrWqulznIBcEX7GIoX7h+1WWrv"
    b = b & "SFUdwBWjDZDihftH+3L2jQAm/kUboNnCvhHAM2oDpHhhnigZdai4MnoALrR/8Y3ihdliXFxoqQ7gQsULDZCKF+aJruco8UtG"
    b = b & "ABcuOy9IRvOEIgaAiVG84EK9+aLXW1sggLNfuMNo9mj22UxUEQPAWUtFVdLVUdoS3TwxKup6COgOKYDD67xLl77VpkYymid6"
    b = b & "nX86nU7f6cQA8KB9iwbGUW48HTjFfaKWQLUGqkEtwOGVjNxhNH/0ejcbVVEHoJJusej1bp9OMgIO73+qpFssSka1X5KMgMP7"
    b = b & "pZ50i0XJyFkjgHNZcclIT7r5o+7oPQRIRsDhlYx+UNa9SJSM/u4qCYCHgbCqLg1S54+SkWU6gHMyqhNAZ14ko3lDMgI4Kxl1"
    b = b & "xqjDl9PBUtw37BkBnNWoswHRGaP5Y1TTaZYKHF4DYQNiT+nTwVLcN8Y5I8kIOLyRjKYDpbh/tDTaEqlqOuDwdF9YJioWqWik"
    b = b & "4hFdu4FD+5dzMtJ9Yf4oGdUlvdAOCDi0fz03SZWM5o9xn1E/ukICOLTPzjOj9i2cMZo3er1rwfS1m16Bo/uVZLRoNCOtL2Df"
    b = b & "B4DD+sU5GXWnjmQ0b4zuC73+LZcCHJZktFyMA6/t2QEcWncZlYxcHzFv9FrXnLaHAAdegcNr47wnc8lo3hiVdB161ZcOOLw2"
    b = b & "zttAt0w3f/Sa99rrvgAcntLuZWJ06+61b6kU4NB6Km+ZTjKaN1qeG8kIgIsODJLRPNHr3K267RkpXgA40w5o/qhg5BsNUgEe"
    b = b & "uUJi3rg87KoNEMBZpcWS0Twxzhe5UA9goiIGN73OEyWjP5+7LzhfBHChu3R6UpeM5onRHNV+EcCFklEb6pUbTwdOcbtoVjRK"
    b = b & "uisasV8EcKFBseuv28uYDqDidjFaAFXWbb8IYKJk9MP5+mtnje4bHS62RAdwxbhGoo11yeh+0RLduDLCEh3AFS0baQl0v7js"
    b = b & "uqCKDuAZzhrdP3TpBviEcdZoOoCK28Rll+7Ppy8+AA8q7+7JvcOY04FUvD8uS7oBeEbLdCWiSrynA6l4X7RfVKViBSL2iwA+"
    b = b & "olLjntpV1N0+RgugH5V0A3xaT+0q6u4T43yR/SKAT+ipXRHD7WNcGVEy+tfpiw7AUx3EVMRw+xjFC1oAAbzAZVug6YAq3hYt"
    b = b & "eVYU8rPiBYCX6+m9Dt72jW4TozlqS3UOuwK80Ng3koxuE72Oo/OCSjqAFxqHX120d7vQeQHglTpr9I19o5vFqKTTeQHgldrb"
    b = b & "cN7oNqGSDuCNWqpz3uj9cVlJJxkBvFIb7Z01cg35++Kykk7xAsAbdCamJ3pLdW+PXrvK5L+WjADexlLdbaL9osq6P5u+wAB8"
    b = b & "Wnscbb4X0wFWvDxGJZ1kBPBGVdXpxvD2UNYNcAOW6t4Xo6xbTzqAd+iJ/ju3v74pmk1WjdjMUjICeKdmRw7Avj56vepiUWiQ"
    b = b & "CvBOlSRX4q1X3evi8qrxruYA4J1aZmpgnQ644vkoGZXEK+uWjABuoNlRG/HTAVc8HyWjljdLRr+avqAAvF77RnURUMjwuhhX"
    b = b & "R0hGADdSQjI7el04YwRwYy3VVcSgI8PLYhx41a0b4MbGPUfTgVd8GHU9l4wA7qCqsM7NNNBOB1/xNHqNWtZ0xgjgDlqu++uV"
    b = b & "wVc8jZYzm0VKRgB3UDL6k0Own4zRCkgyAriTEpK9o+djXDfeQeGqEAG4gzoyNNjaO7oel33p3PAKcEdmR8/HSEYla8kI4I5a"
    b = b & "fqoJqHNHH0bJqH21kpFlOoA7a3NeV4YPo2T0k47dAPMZ+0fTAfnIIRkBzMzs6MMYyegHyQhgPm3Su+/oMS6TkQIGgJk04FY9"
    b = b & "5iDsQ1wmo8+nLxYA92O57jEuk5FqOoCZlZAUMzyWdpeMXKwHMLOW6+rHNh2cjxYjGfVzyQhgAe48euzA0B6aajqAhXTVdl2r"
    b = b & "p4P0kaLlysKeEcBCOgh79DuPSsaVu5sZASzo6Mt19exr/8zMCGBh4/zRdKA+QtQktWTs0CvAwpodHfVW2P7Of5OMANbhyIdh"
    b = b & "S0btnwGwAiWko/Wuq7y7JPyb6YsBwDIq9f76YOXeJaMqCiUjgBVpuernK4P2XqNk1GzwS10YANblaOXeHXqtkEHnboCVKSEd"
    b = b & "ZYbUWaP+rs4aAazQka4qV94NsFJHuYxvVNQp7wZYqQbove8flYxaplNRB7BiRyhoqJzdVRIAK3eEA7Et1SliAFi5PV9X3lJd"
    b = b & "s7/+jgCsWHsqDdx1up4O5nuIijUcfgXYgCrs9nr+qD0j10kAbMReCxos1QFszF4TUt0YSkpmRwAbUULqyu7pgL71UFUHsDEd"
    b = b & "iv3pyoC+5ejM0VdmRwDbsseSb+2BADamQfvb837LdFDfalS+3t6RYgaADWlJq2S0pzNI7YdZqgPYmGYRf9pZl++uJJeQADam"
    b = b & "gXtPFXYVM3wx/UsCsH4lpAoApgP7VqPZkb0jgA3a26FYbYIANmpPF/ONwgyHYQE2aE8zpIozLNcBbNSeZkj2jwA2rITUQD4d"
    b = b & "3LcYFWdISAAbtZeE1Dmq7nSSkAA2ai+X8412QQoaADaqGcUeElIHYruKHYCN2ktCUmEHsHF7aR1k/whg41rm2kNRQ0lVQgLY"
    b = b & "sL1U2fV3cCEfwIbt5WBsfwdFDQAbtpeE5MpygI3bSy+7EpIu3wAbtqeEpKgBYMP2ckGfhASwcXs5GGvJDmDjSkh/PDcnnQ7y"
    b = b & "WwozJICNqxlpjUnHbatbjfbBVNkBbFiziq/OzUmng/yWwsFYgB0Yy3bTQX5LoXUQwA40kP90ZZDfUlSYoagBYONKSFvvZ1dC"
    b = b & "NUMC2Lg9HI5tydEeEsDGVWm39YT0u9Pp9PW5QAOAjWpm0Y2rWz6LVNl6X79lO4AN69qGb8+zjOlAv5UoGf3ZFRQA2zdmSdOB"
    b = b & "fkvRsqMZEsDGNZB3lmc6yG8pdGsA2IES0ta7fncWyQwJYOOaWWx9hvQHCQlg+744nU4/brzJal/7bxU2AGxfs4utFzY0y9NC"
    b = b & "CGDj9tBCqMIGCQlg49pH2nrHhs4i2UcC2Lj2Xjocu+V9pO51+vJ0Ov16+pcDYFv2cDdSsyTnkQA2bi/7SBISwMbtYR/JdeYA"
    b = b & "O9B5pO833mhVQgLYiZbt2oeZDvRbic4iSUgAO7D1vnYlJKXfADvQwdLKp7da/i0hAexISemnK4P9FkJCAtiRLS/bSUgAO1LX"
    b = b & "hq4F3+KynaIGgJ1p2a5L76YD/trDDAlgZ0pIW1y2K4maIQHsSE1KvzpX3E0H/TWHhASwQy19ba2VkIQEsEMt23WLbAUO04F/"
    b = b & "rSEhAexUs6QtFTf0tVYlCMDONEva0rKd+5AAdqxZUkt308F/bdHSYhcMKvsG2KmtdG6QkAAOoH2ZP1xJAmsKCQngALZwUFZC"
    b = b & "AjiIDsuueZY0ElLJE4Ada6D/25VEsJYoIZUwVdkBHECDfbOQaTJYQ9SdvKRkhgRwAGveS5KQAA6mWdIab5SVkAAOpiq2v15J"
    b = b & "CEuHhARwQA36a0tKEhLAAXVQtgTwuyuJYakYCUmVHcDB/NvKysBLjt+fz0sBcDBrKnAYCcn1EwAH1CypKx9aLpsmiLlDQgI4"
    b = b & "uNHB4bdXksScISEB8I8k0ExpmiTmDAkJgH/MkrrEb8mqOwkJgH9Yuvnq70+n09fOIQGQqu7+a6H9pBLSV6fT6YvpFwXA8TQ7"
    b = b & "6RzQEqXgIyGZIQHwD5WCl5Dm3k+SkAD4QEmpaypq5TNNHPcKCQmAq+oKPmeRwyhq+Gz6hQBAJdhzFTl0fXlJEAA+0GylIoe/"
    b = b & "zJCUJCQAPqqk1N5OnRzumZQkJAA+6fNzMqqbw72SkoQEwItUeVcp+B/vlJQkJABe7J5JSUIC4FXulZT68yQkAF5lJKVb7ilJ"
    b = b & "SAC8SYUOP57bDN0iKUlIALzZr87dFX6+QVKSkAB4l2ZKnVOqo8M0ybwmJCQA3u0X544O9b57a0NWCQmAm+lyvbqEv+Xqigok"
    b = b & "jp6Q6nT+P6b/EoC3qSFrbYZeW4FXcUQ31h5ZCalCkV5DAG6gmc43p9Ppr1cSz3NRQjr6DKmE/H/Pr1tLoO3PAfBONWWtAq99"
    b = b & "pf+8koCmISE9/P0rDmkfrh9L6mZLADfSk35LeJ86ryQhPSakXo9eqwo9+udew0rsAXinlqIqDf/Ytej2kJ4mpBG9Xr1uvX5H"
    b = b & "f30AbqJ2Q1XhtT9yreDBDOl6QhqzpV6fwhIewI2UlMZs6XJvqWW9o88AnktIIyqnb0+u1/Bfpr8ZgNcbV6P3xF/boZGQzJA+"
    b = b & "npBGjCU8VXgAN9IAXGJqgO3JX0J6WUJqCe8vp9Pp+/NSKAA30r5Ig6yE9LKENJJSe3FdbNiBWgBupH54EtLLE9JISiWkZkuS"
    b = b & "EgA389qENKJih5Y9j14UAsCNvDUhFVUslpTMlAB4t/ckpMukdPSlTwDe6b0JaSSl/gxJCYA3u0VCKtpTqtBBSTgAb3KrhFRU"
    b = b & "fVdISgC82i0T0jin9N3pdPr19H8EAB9zy4RUjI4OGrIC8Cq3TkgjnFEC4FXulZC6U6legc4oAfAi90pIRZV3FTnYTwLgk+6Z"
    b = b & "kIqu+LCfBMAn3TshFZbuAPikORJSnRyaKTmfBMCz5khIheviAfiouRJSYekOgGfNmZB+fz44KykB8IE5E1LJqP+XhATAB+ZM"
    b = b & "SEUHZttPclUFAE/MnZCKet2ZJQHwxBIJqVmSpATAE0skpOJny3YAXFoqITVLKimZJQHwD0slpMKyHQD/tGRCqqVQZ5N0cABg"
    b = b & "0YRUdJGfvSQAFk9I3Zf05el0+mz6hQFwLEsnpMIsCYBVJCTVdgCsIiFV3FAZuKvOObz/D00oPap6/jeGAAAAAElFTkSuQmCC"
    LogoB64 = b
End Function

Private Function LogoFilePath() As String
    ' Use scu_emblem.png beside the workbook only (never the old embedded full-logo fallback).
    On Error GoTo Fail
    Dim folder As String, localPath As String
    folder = ThisWorkbook.Path
    If folder = "" Then Exit Function
    localPath = folder & Application.PathSeparator & LOGO_EMBLEM_FILE
    If Len(Dir(localPath)) > 0 Then
        LogoFilePath = localPath
        Exit Function
    End If
    Exit Function
Fail:
    LogoFilePath = ""
End Function

Private Sub InsertLabelLogo(ws As Worksheet, ByVal shapeName As String, ByVal rightPt As Single, ByVal heightPt As Single, Optional ByVal leftBoundPt As Single = 0, Optional ByVal bandFirstRow As Long = 0, Optional ByVal bandLastRow As Long = 0, Optional ByVal centerHoriz As Boolean = False)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Dim logoPath As String
    logoPath = LogoFilePath()
    If logoPath = "" Then Exit Sub

    Dim pic As Shape
    Dim maxW As Single, targetH As Single, initW As Single

    targetH = heightPt
    maxW = rightPt - LOGO_RIGHT_PAD_PT
    If leftBoundPt > 0 Then maxW = rightPt - leftBoundPt - LOGO_RIGHT_PAD_PT
    If maxW <= 0 Then Exit Sub
    initW = targetH * LOGO_ASPECT

    On Error Resume Next
    ' Insert at natural aspect (not maxW x targetH - that box is wider than the emblem and squashes it).
    Set pic = ws.Shapes.AddPicture2(logoPath, msoFalse, msoTrue, 0, 0, initW, targetH, 0)
    If pic Is Nothing Then
        Set pic = ws.Shapes.AddPicture(logoPath, msoFalse, msoTrue, 0, 0, initW, targetH)
    End If
    On Error GoTo 0
    If pic Is Nothing Then Exit Sub

    pic.name = shapeName
    pic.Placement = xlFreeFloating
    pic.LockAspectRatio = msoTrue
    pic.Height = targetH
    If pic.Width > maxW Then pic.Width = maxW
    If centerHoriz And leftBoundPt > 0 Then
        pic.Left = leftBoundPt + ((rightPt - leftBoundPt) - pic.Width) / 2
    Else
        pic.Left = rightPt - pic.Width - LOGO_RIGHT_PAD_PT
    End If
    If bandFirstRow > 0 And bandLastRow >= bandFirstRow Then
        Dim bandTop As Single, bandH As Single, r As Long
        bandTop = ws.Rows(bandFirstRow).Top
        bandH = 0
        For r = bandFirstRow To bandLastRow
            bandH = bandH + ws.Rows(r).RowHeight
        Next r
        pic.Top = bandTop + (bandH - pic.Height) / 2
        If pic.Top < bandTop Then pic.Top = bandTop
    End If
    pic.ZOrder msoBringToFront
End Sub

Private Function LogoTopInBand(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal logoH As Single) As Single
    ' Legacy helper - InsertLabelLogo now centers in-band after sizing.
    Dim bandTop As Single, bandH As Single, r As Long
    bandTop = ws.Rows(firstRow).Top
    bandH = 0
    For r = firstRow To lastRow
        bandH = bandH + ws.Rows(r).RowHeight
    Next r
    LogoTopInBand = bandTop + (bandH - logoH) / 2
    If LogoTopInBand < bandTop Then LogoTopInBand = bandTop
End Function

' Print surface (hidden Label Preview sheet): keep header merges and emblem in sync
' with the gallery. Workbooks built before the logo-slot fix may still have A2:H2 merged.
Private Sub EnsurePrintLabelHeaderLayout(ws As Worksheet)
    On Error Resume Next
    ws.Range("A2:H2").UnMerge
    ws.Range("A3:H3").UnMerge
    ws.Range("A2:F2").UnMerge
    ws.Range("A3:F3").UnMerge
    ws.Range("A2:C2").UnMerge
    ws.Range("A3:C3").UnMerge
    ws.Range("D2:E3").UnMerge
    ws.Range("F2:H2").UnMerge
    ws.Range("F3:H3").UnMerge
    ws.Range("G2:H3").UnMerge
    On Error GoTo 0
    ' Three zones: name A:C (2 lines), emblem D:E (centered), phone/address F:H (2 lines).
    ws.Range("A2:C2").Merge
    ws.Range("A3:C3").Merge
    ws.Range("D2:E3").Merge
    ws.Range("F2:H2").Merge
    ws.Range("F3:H3").Merge
    With ws.Range("D2:E3")
        .Interior.Pattern = xlNone
    End With
    ws.Rows(2).RowHeight = LOGO_HDR_ROW1_PT
    ws.Rows(3).RowHeight = LOGO_HDR_ROW2_PT
    If Trim(ws.Cells(2, 1).Value) = "" Then ws.Cells(2, 1).Value = "SATURDAY CLINIC"
    If Trim(ws.Cells(3, 1).Value) = "" Then ws.Cells(3, 1).Value = "FOR THE UNINSURED"
    If Trim(ws.Cells(2, 6).Value) = "" Then ws.Cells(2, 6).Value = "(414) 588-2865"
    If Trim(ws.Cells(3, 6).Value) = "" Then ws.Cells(3, 6).Value = "1121 E. North Ave, Milwaukee WI"
    Call FmtLblClinicName(ws.Cells(2, 1), CLINIC_NAME_FONT_PRINT)
    Call FmtLblNameSub(ws.Cells(3, 1), CLINIC_NAMESUB_FONT_PRINT)
    Call FmtLblContactRight(ws.Cells(2, 6), CLINIC_PHONE_FONT_PRINT, True, "B")
    Call FmtLblContactRight(ws.Cells(3, 6), CLINIC_ADDR_FONT_PRINT, False, "T")
End Sub

Private Sub RefreshPrintLabelLogo(ws As Worksheet)
    Call EnsurePrintLabelHeaderLayout(ws)
    Call InsertLabelLogo(ws, "scuLogo", ws.Range("F2").Left, LOGO_HEIGHT_PRINT, ws.Range("D2").Left, 2, 3, True)
End Sub

Private Sub BuildLabelPreviewLayout(ws As Worksheet)
    On Error Resume Next
    Application.ScreenUpdating = False

    ws.Range("A1:H22").UnMerge
    ws.Range("A1:H22").ClearContents
    ws.Range("A1:H22").Interior.Pattern = xlNone

    ' Width: scale A:H to LABEL_WIDTH_PT so content stays inside one DK-1202 label.
    ws.Columns("A:H").ColumnWidth = 5.8
    Call ApplyLabelContentWidth(ws)

    ws.Rows(1).RowHeight = 2
    ws.Rows(2).RowHeight = LOGO_HDR_ROW1_PT
    ws.Rows(3).RowHeight = LOGO_HDR_ROW2_PT
    ws.Rows(4).RowHeight = 5
    ws.Rows(5).RowHeight = 14
    ws.Rows(6).RowHeight = 12
    ws.Rows(7).RowHeight = 20
    ws.Rows(8).RowHeight = 12
    ws.Rows(9).RowHeight = 10
    ws.Rows(10).RowHeight = 15
    ws.Rows(11).RowHeight = 15
    ws.Rows(12).RowHeight = 15
    ws.Rows(13).RowHeight = 3
    ws.Rows(14).RowHeight = 2
    ws.Rows(15).RowHeight = 15
    ws.Rows(16).RowHeight = 10
    ws.Rows(17).RowHeight = 14
    ws.Rows(18).RowHeight = 14

    ws.Range("A2:C2").Merge
    ws.Range("A3:C3").Merge
    ws.Range("D2:E3").Merge
    ws.Range("F2:H2").Merge
    ws.Range("F3:H3").Merge
    ws.Range("A5:D6").Merge
    ws.Range("E5:H6").Merge
    ws.Range("A7:H7").Merge
    ws.Range("A8:E8").Merge
    ws.Range("F8:H8").Merge
    ws.Range("A9:H9").Merge
    ws.Range("A10:H12").Merge
    ws.Range("A15:D15").Merge
    ws.Range("E15:H15").Merge
    ws.Range("A17:H17").Merge
    ws.Range("A18:H18").Merge

    ws.Cells(2, 1).Value = "SATURDAY CLINIC"
    ws.Cells(3, 1).Value = "FOR THE UNINSURED"
    ws.Cells(2, 6).Value = "(414) 588-2865"
    ws.Cells(3, 6).Value = "1121 E. North Ave, Milwaukee WI"
    ws.Cells(9, 1).Value = "DIRECTIONS"
    With ws.Range("D2:E3")
        .Value = ""
        .Interior.Pattern = xlNone
    End With

    ws.Cells(5, 1).Formula = "=IF('Patient & Input'!C5<>"""",'Patient & Input'!C5,""[Patient Name]"")"
    ws.Cells(5, 5).Formula = "=IF('Patient & Input'!C6<>"""",""DOB  ""&'Patient & Input'!C6,""DOB  --"")"
    ws.Cells(8, 6).Formula = "=IF('Patient & Input'!C7<>"""",""Rx  ""&TEXT('Patient & Input'!C7,""m/d/yyyy""),""Rx  --"")"

    If Trim(ws.Cells(7, 1).Value) = "" Then _
        ws.Cells(7, 1).Value = "[Select a medication row, then Update]"
    Call SetMiniValue(ws.Cells(15, 1), "EXP", "--", 12, "L")
    Call SetMiniValue(ws.Cells(15, 5), "LOT", "--", 12, "R")

    ws.Cells(17, 1).Value = "Auto-fills from the Medications tab. Print via 'Print Checked Labels'."
    ws.Cells(18, 1).Value = "Brother QL-1100c  " & Chr(183) & "  DK-1202 62 x 100 mm  " & Chr(183) & "  Landscape"

    Call FmtLblClinicName(ws.Cells(2, 1), CLINIC_NAME_FONT_PRINT)
    Call FmtLblNameSub(ws.Cells(3, 1), CLINIC_NAMESUB_FONT_PRINT)
    Call FmtLblContactRight(ws.Cells(2, 6), CLINIC_PHONE_FONT_PRINT, True, "B")
    Call FmtLblContactRight(ws.Cells(3, 6), CLINIC_ADDR_FONT_PRINT, False, "T")
    Call FmtLbl(ws.Cells(5, 1), 12, True, "L", "C")
    Call FmtLbl(ws.Cells(5, 5), 12, True, "R", "C")
    Call FmtLbl(ws.Cells(7, 1), 14, True, "L", "C")
    Call FmtLbl(ws.Cells(8, 1), 8, False, "L", "C")
    Call FmtLbl(ws.Cells(8, 6), 8, False, "R", "C")
    Call FmtLbl(ws.Cells(9, 1), 6.5, True, "L", "C")
    Call FmtLbl(ws.Cells(10, 1), 11, True, "L", "T")
    Call FmtLbl(ws.Cells(17, 1), 8, False, "L", "C")
    Call FmtLbl(ws.Cells(18, 1), 8, False, "L", "C")
    ws.Cells(17, 1).Font.Color = RGB(150, 150, 150)
    ws.Cells(18, 1).Font.Color = RGB(150, 150, 150)

    With ws.Range("A10:H12")
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With

    ws.Range(ws.Cells(1, 1), ws.Cells(18, 8)).Borders.LineStyle = xlNone
    With ws.Range("A4:H4").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    With ws.Range("A15:H15").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    On Error Resume Next
    ws.Shapes("scuLogo").Delete
    On Error GoTo 0
    Call RefreshPrintLabelLogo(ws)

    Call ApplyLabelPageSetup(ws)

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub ApplyLabelContentWidth(ws As Worksheet)
    Dim currentW As Double, sf As Double, col As Long
    currentW = ws.Range("A1:H15").Width
    If currentW <= 0 Then Exit Sub
    sf = LABEL_WIDTH_PT / currentW
    For col = 1 To 8
        ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth * sf
    Next col
End Sub

Private Sub ApplyLabelPageSetup(ws As Worksheet)
    With ws.PageSetup
        .PrintArea = ws.Range("A1:H15").Address
        .LeftMargin = Application.InchesToPoints(0.04)
        .RightMargin = Application.InchesToPoints(0.04)
        .TopMargin = Application.InchesToPoints(0.04)
        .BottomMargin = Application.InchesToPoints(0.04)
        .HeaderMargin = 0
        .FooterMargin = 0
        ' Fit the whole A1:H15 label onto ONE page. Critical: the EXP/LOT row (row 15) sits
        ' at the very bottom of the print area, so on a printer whose printable height is a
        ' hair shorter than ours it would spill onto a (never-printed) page 2 and vanish -
        ' which is exactly the "Lot/Exp missing on print but fine in the gallery" symptom.
        ' Fit-to-1-page scales the label to the driver's real printable area on ANY PC.
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintGridlines = False
        .PrintHeadings = False
    End With
End Sub

' Only the emblem may print as a shape; buttons must not add a second (blank) page.
Private Sub PrepareLabelSheetForPrint(ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    ws.ResetAllPageBreaks
    For Each shp In ws.Shapes
        shp.PrintObject = (shp.Name = "scuLogo")
    Next shp
    On Error GoTo 0
End Sub

Private Function LabelHeaderFont() As String
    ' Helvetica on the clinic header; Arial if Excel cannot resolve Helvetica on this PC.
    On Error Resume Next
    ThisWorkbook.Styles("SCUFontProbeHdr").Delete
    Err.Clear
    Dim probe As Style
    Set probe = ThisWorkbook.Styles.Add("SCUFontProbeHdr")
    probe.Font.Name = FONT_LABEL_HDR
    If Err.Number = 0 And StrComp(probe.Font.Name, FONT_LABEL_HDR, vbTextCompare) = 0 Then
        LabelHeaderFont = FONT_LABEL_HDR
    Else
        LabelHeaderFont = FONT_LABEL_HDR_FB
    End If
    ThisWorkbook.Styles("SCUFontProbeHdr").Delete
    Err.Clear
    On Error GoTo 0
End Function

Private Sub FmtLbl(rng As Range, sz As Single, bld As Boolean, h As String, v As String, Optional fontName As String = "Arial")
    With rng.Font
        .Name = fontName
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

' Clinic header lines: one row each, shrink horizontally, never wrap (wrap clips on print).
Private Sub FmtLblHeader(rng As Range, sz As Single, bld As Boolean, Optional vAlign As String = "C")
    With rng
        .Font.Name = LabelHeaderFont()
        .Font.Size = sz
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .ShrinkToFit = True
        If bld Then
            .Font.Bold = True
        Else
            .Font.Bold = False
        End If
    End With
    Select Case vAlign
        Case "T": rng.VerticalAlignment = xlTop
        Case "B": rng.VerticalAlignment = xlBottom
        Case Else: rng.VerticalAlignment = xlCenter
    End Select
End Sub

Private Sub FmtLblClinicName(rng As Range, sz As Single)
    Call FmtLblHeader(rng, sz, True, "B")
End Sub

Private Sub FmtLblNameSub(rng As Range, sz As Single)
    With rng
        .Font.Name = LabelHeaderFont()
        .Font.Size = sz
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .ShrinkToFit = True
    End With
End Sub

Private Sub FmtLblContactRight(rng As Range, sz As Single, bld As Boolean, vAlign As String)
    With rng
        .Font.Name = "Arial"
        .Font.Size = sz
        .Font.Bold = bld
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlRight
        .WrapText = False
        .ShrinkToFit = True
    End With
    Select Case vAlign
        Case "T": rng.VerticalAlignment = xlTop
        Case "B": rng.VerticalAlignment = xlBottom
        Case Else: rng.VerticalAlignment = xlCenter
    End Select
End Sub

Private Sub SetMiniValue(c As Range, miniText As String, valueText As String, vSize As Single, hAlign As String)
    ' Footer field: small label + larger bold value in one cell (e.g. "EXP 05/2027")
    Dim s As String
    s = miniText & "  " & valueText
    c.Value = s
    c.WrapText = False
    c.Font.Name = "Arial"
    c.Font.Color = RGB(0, 0, 0)
    ' Shrink-to-fit for the value (merged cells ignore Excel's ShrinkToFit, so size by length)
    Dim vs As Single
    vs = vSize
    If Len(valueText) > 14 Then vs = 9
    If Len(valueText) > 20 Then vs = 7.5
    If Len(valueText) > 26 Then vs = 6.5
    On Error Resume Next
    c.Characters(1, Len(miniText)).Font.Size = 7
    c.Characters(1, Len(miniText)).Font.Bold = False
    c.Characters(Len(miniText) + 1, Len(s) - Len(miniText)).Font.Size = vs
    c.Characters(Len(miniText) + 1, Len(s) - Len(miniText)).Font.Bold = True
    On Error GoTo 0
    Select Case hAlign
        Case "R": c.HorizontalAlignment = xlRight
        Case "C": c.HorizontalAlignment = xlCenter
        Case Else: c.HorizontalAlignment = xlLeft
    End Select
    c.VerticalAlignment = xlCenter
End Sub

Private Function MedFontSize(ByVal s As String) As Single
    ' Keep the hero medication line on one line by stepping the size down
    Dim n As Long: n = Len(s)
    If n <= 24 Then
        MedFontSize = 14
    ElseIf n <= 30 Then
        MedFontSize = 12.5
    ElseIf n <= 38 Then
        MedFontSize = 11
    Else
        MedFontSize = 10
    End If
End Function

Private Function NameFontSize(ByVal s As String) As Single
    Dim n As Long: n = Len(s)
    If n <= 20 Then
        NameFontSize = 17
    ElseIf n <= 28 Then
        NameFontSize = 15
    ElseIf n <= 36 Then
        NameFontSize = 13
    ElseIf n <= 44 Then
        NameFontSize = 11.5
    Else
        NameFontSize = 10
    End If
End Function

Private Function PatientNameFontSize(ByVal patientName As String, ByVal medLine As String) As Single
    ' Keep the patient name larger than the medication line, stepping both down for long text.
    Dim sz As Single
    sz = NameFontSize(patientName)
    If sz <= MedFontSize(medLine) Then sz = MedFontSize(medLine) + 1
    PatientNameFontSize = sz
End Function

Private Function SigFontSize(ByVal s As String) As Single
    Dim n As Long: n = Len(s)
    If n <= 70 Then
        SigFontSize = 10.5
    ElseIf n <= 110 Then
        SigFontSize = 9.5
    Else
        SigFontSize = 8.5
    End If
End Function

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

    Dim shp As Shape
    Dim si As Long
    For si = ws.Shapes.Count To 1 Step -1
        ws.Shapes(si).Delete   ' clear ALL shapes (cards, buttons, stray/old images); rebuilt below
    Next si
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone
    ws.Rows.RowHeight = 14

    ws.Columns("A:F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 2
    ws.Columns("H:J").ColumnWidth = 16

    Dim patName As String, dob As String
    patName = Trim(wsI.Range("C5").Value)
    dob = Trim(wsI.Range("C6").Value)

    ws.Cells(1, 1).Value = "ALL MEDICATION LABELS" & IIf(patName <> "", "   -   " & patName, "")
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.name = "Arial"

    Dim logoPath As String
    logoPath = LogoFilePath()
    Dim logoOK As Boolean
    logoOK = (logoPath <> "")

    Dim lastRow As Long
    lastRow = wsM.Cells(wsM.Rows.Count, C_NAME).End(xlUp).Row

    ' Header band (rows 1-3) sized BEFORE the card loop so each card's logo is placed
    ' against the correct row tops (the logo is positioned relative to its base row).
    ' Two rows tall so Print Checked Labels + Reprint Last Batch stack one above the other.
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 8

    Dim n As Integer, base As Long, r As Long
    n = 0
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(wsM.Cells(r, C_NAME).Value) = "" Then GoTo NextR
        n = n + 1
        base = 4 + (n - 1) * 12    ' cards start at row 4 (rows 1-3 are the button header band)

        Dim medName As String, strength As String, qty As String, formTxt As String
        Dim sig As String, expv As String, lotv As String, dateRx As String, warn As String
        medName = Trim(wsM.Cells(r, C_NAME).Value)
        strength = Trim(wsM.Cells(r, C_STR).Value)
        formTxt = Trim(wsM.Cells(r, C_FORM).Value)
        qty = Trim(wsM.Cells(r, C_QTY).Value)
        sig = Trim(wsM.Cells(r, C_SIG).Value)
        expv = Trim(wsM.Cells(r, C_EXP).Value)
        lotv = Trim(wsM.Cells(r, C_LOT).Value)
        dateRx = Trim(wsM.Cells(r, C_DATE).Value)
        warn = Trim(wsM.Cells(r, C_WARN).Value)

        Dim medLine As String
        medLine = medName
        If strength <> "" Then medLine = medLine & " " & strength
        Dim fq As String
        fq = formTxt
        If qty <> "" Then
            If fq <> "" Then
                fq = fq & "   " & Chr(183) & "   Qty " & qty
            Else
                fq = "Qty " & qty
            End If
        End If
        Dim refillsG As String: refillsG = Trim(wsM.Cells(r, C_REF).Value)
        If refillsG <> "" Then
            If fq <> "" Then
                fq = fq & "   " & Chr(183) & "   Refills " & refillsG
            Else
                fq = "Refills " & refillsG
            End If
        End If

        ws.Range(ws.Cells(base, 1), ws.Cells(base, 2)).Merge
        ws.Cells(base, 1).Value = "SATURDAY CLINIC"
        Call FmtLblClinicName(ws.Cells(base, 1), CLINIC_NAME_FONT_GALLERY)
        ws.Range(ws.Cells(base + 1, 1), ws.Cells(base + 1, 2)).Merge
        ws.Cells(base + 1, 1).Value = "FOR THE UNINSURED"
        Call FmtLblNameSub(ws.Cells(base + 1, 1), CLINIC_NAMESUB_FONT_PRINT)
        ws.Range(ws.Cells(base, 5), ws.Cells(base, 6)).Merge
        ws.Cells(base, 5).Value = "(414) 588-2865"
        Call FmtLblContactRight(ws.Cells(base, 5), CLINIC_PHONE_FONT_PRINT, True, "B")
        ws.Range(ws.Cells(base + 1, 5), ws.Cells(base + 1, 6)).Merge
        ws.Cells(base + 1, 5).Value = "1121 E. North Ave, Milwaukee WI"
        Call FmtLblContactRight(ws.Cells(base + 1, 5), CLINIC_ADDR_FONT_PRINT, False, "T")
        With ws.Range(ws.Cells(base, 3), ws.Cells(base + 1, 4))
            .Merge
            .Value = ""
            .Interior.Pattern = xlNone
        End With

        ws.Range(ws.Cells(base + 2, 1), ws.Cells(base + 2, 3)).Merge
        ws.Cells(base + 2, 1).Value = IIf(patName <> "", patName, "[Patient Name]")
        Call FmtLbl(ws.Cells(base + 2, 1), 12, True, "L", "C")
        ws.Cells(base + 2, 1).Font.Size = 12       ' Name/DOB a bit smaller (matches print)
        ws.Range(ws.Cells(base + 2, 4), ws.Cells(base + 2, 6)).Merge
        ws.Cells(base + 2, 4).Value = "DOB " & IIf(dob <> "", dob, "--")
        Call FmtLbl(ws.Cells(base + 2, 4), 12, True, "R", "C")
        ws.Cells(base + 2, 4).Font.Size = 12

        ws.Range(ws.Cells(base + 3, 1), ws.Cells(base + 3, 6)).Merge
        ws.Cells(base + 3, 1).Value = medLine
        Call FmtLbl(ws.Cells(base + 3, 1), 13, True, "L", "C")
        ws.Cells(base + 3, 1).WrapText = True
        Dim medWrap As Boolean: medWrap = (Len(medLine) > MED_WRAP_MAXLEN)
        If medWrap Then
            ws.Cells(base + 3, 1).Font.Size = MED_WRAP_FONT
        Else
            ws.Cells(base + 3, 1).Font.Size = MedFontSize(medLine)
        End If

        ws.Range(ws.Cells(base + 4, 1), ws.Cells(base + 4, 4)).Merge
        ws.Cells(base + 4, 1).Value = fq
        Call FmtLbl(ws.Cells(base + 4, 1), 8, False, "L", "C")
        ws.Range(ws.Cells(base + 4, 5), ws.Cells(base + 4, 6)).Merge
        ws.Cells(base + 4, 5).Value = "Rx " & IIf(dateRx <> "", dateRx, "--")
        Call FmtLbl(ws.Cells(base + 4, 5), 8, False, "R", "C")

        ws.Range(ws.Cells(base + 5, 1), ws.Cells(base + 5, 6)).Merge
        ws.Cells(base + 5, 1).Value = "DIRECTIONS"
        Call FmtLbl(ws.Cells(base + 5, 1), 6.5, True, "L", "C")

        ws.Range(ws.Cells(base + 6, 1), ws.Cells(base + 7, 6)).Merge
        Dim sigText As String
        sigText = IIf(sig <> "", sig, "[Instructions not found - enter manually]")
        ws.Cells(base + 6, 1).Value = sigText
        Call FmtLbl(ws.Cells(base + 6, 1), 10, True, "L", "T")   ' larger directions (matches print)
        With ws.Range(ws.Cells(base + 6, 1), ws.Cells(base + 7, 6))
            .Interior.Color = RGB(0, 0, 0)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .WrapText = True
        End With

        ws.Range(ws.Cells(base + 8, 1), ws.Cells(base + 8, 3)).Merge
        Call SetMiniValue(ws.Cells(base + 8, 1), "EXP", IIf(expv <> "", expv, "--"), 11, "L")
        ws.Range(ws.Cells(base + 8, 4), ws.Cells(base + 8, 6)).Merge
        Call SetMiniValue(ws.Cells(base + 8, 4), "LOT", IIf(lotv <> "", lotv, "--"), 11, "R")

        ws.Rows(base).RowHeight = LOGO_HDR_ROW1_PT
        ws.Rows(base + 1).RowHeight = LOGO_HDR_ROW2_PT
        ws.Rows(base + 2).RowHeight = 22
        ws.Rows(base + 3).RowHeight = IIf(medWrap, 30, 20)
        ws.Rows(base + 4).RowHeight = 12
        ws.Rows(base + 5).RowHeight = 11
        ws.Rows(base + 6).RowHeight = 14
        ws.Rows(base + 7).RowHeight = 14
        ws.Rows(base + 8).RowHeight = 16

        With ws.Range(ws.Cells(base + 1, 1), ws.Cells(base + 1, 6)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        With ws.Range(ws.Cells(base + 8, 1), ws.Cells(base + 8, 6)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        With ws.Range(ws.Cells(base, 1), ws.Cells(base + 8, 6))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        If IsRowSelected(wsM, r) Then
            With ws.Range(ws.Cells(base, 1), ws.Cells(base + 5, 6))
                .Interior.Color = RGB(212, 237, 218)
            End With
        End If

        If logoOK Then
            Call InsertLabelLogo(ws, "al_logo_" & r, ws.Cells(base, 5).Left, LOGO_HEIGHT_GALLERY, ws.Cells(base, 3).Left, base, base + 1, True)
        End If

        If warn <> "" And UCase(warn) <> "OK" Then
            ws.Cells(base + 9, 1).Value = "! " & warn
            Call FmtLbl(ws.Cells(base + 9, 1), 8, False, "L", "C")
            ws.Cells(base + 9, 1).Font.Color = RGB(191, 54, 12)
        End If

        Dim chkCap As String, chkClr As Long
        If IsRowSelected(wsM, r) Then
            chkCap = "Uncheck this label": chkClr = RGB(46, 125, 50)
        Else
            chkCap = "Check this label": chkClr = RGB(84, 110, 122)
        End If
        Call AddRowButton(ws, "al_check_" & r, chkCap, "RowCheck", base, chkClr)
        Call AddRowButton(ws, "al_edit_" & r, "Edit this med", "RowEdit", base + 3, RGB(21, 101, 192))
        Call AddRowButton(ws, "al_remove_" & r, "Remove this med", "RowRemove", base + 6, RGB(191, 54, 12))
NextR:
    Next r

    If n = 0 Then
        ws.Cells(4, 1).Value = "No medications to preview yet. Parse or add medications first."
    End If

    Call AddButtonToSheet(ws, "galtop_prnchk", "Print Checked Labels", "PrintCheckedLabels", 1, 8, 176, 30, RGB(216, 67, 21))
    Call AddButtonToSheet(ws, "galtop_refresh", "Refresh Previews", "PreviewAllLabels", 1, 11, 150, 24, RGB(0, 121, 107))
    Call AddButtonToSheet(ws, "galtop_reprint", "Reprint Last Batch", "ReprintLastBatch", 2, 8, 176, 30, RGB(0, 131, 143))

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

Public Sub RowCheck()
    Dim r As Long
    r = CallerRow()
    If r <= MEDS_HDR_ROWS Then Exit Sub
    Call ToggleRowSelect(r)
    Call BuildAllLabelsPreview
End Sub

Public Sub RowRemove()
    Dim r As Long
    r = CallerRow()
    If r <= MEDS_HDR_ROWS Then Exit Sub
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)

    Dim nm As String
    nm = Trim(wsM.Cells(r, C_NAME).Value & " " & wsM.Cells(r, C_STR).Value)
    If nm = "" Then Exit Sub

    ' Confirm before removing THIS specific medication (named), so an accidental click
    ' on the gallery can't silently delete a row.
    If MsgBox("Remove this medication?" & vbCrLf & vbCrLf & "   " & nm & vbCrLf & vbCrLf & _
              "(This cannot be undone.)", vbYesNo + vbExclamation, "Remove Medication") = vbNo Then Exit Sub

    ' Delete the table columns (1..C_SEL, the full row incl. Source) and shift up; side buttons stay.
    Application.EnableEvents = False
    wsM.Range(wsM.Cells(r, 1), wsM.Cells(r, C_SEL)).Delete Shift:=xlUp
    Application.EnableEvents = True

    Call RenumberMeds
    Call ValidateMedications(False)
    Call ApplyAllRowStates(wsM)
    Call ApplySourceValidation(wsM)
    Call BuildAllLabelsPreview
End Sub

Public Sub RowEdit()
    Dim r As Long
    r = CallerRow()
    If r <= MEDS_HDR_ROWS Then Exit Sub
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)

    If Not EditMedWithForm(wsM, r) Then
        Dim t As String
        t = InputBox("Medication NAME:", "Edit medication", CStr(wsM.Cells(r, C_NAME).Value))
        If StrPtr(t) <> 0 Then wsM.Cells(r, C_NAME).Value = Trim(t)
        t = InputBox("STRENGTH (e.g. 10 mg):", "Edit medication", CStr(wsM.Cells(r, C_STR).Value))
        If StrPtr(t) <> 0 Then wsM.Cells(r, C_STR).Value = Trim(t)
        t = InputBox("DOSAGE FORM (tablet, capsule...):", "Edit medication", CStr(wsM.Cells(r, C_FORM).Value))
        If StrPtr(t) <> 0 Then wsM.Cells(r, C_FORM).Value = Trim(t)
        t = InputBox("QUANTITY:", "Edit medication", CStr(wsM.Cells(r, C_QTY).Value))
        If StrPtr(t) <> 0 Then wsM.Cells(r, C_QTY).Value = Trim(t)
        t = InputBox("DIRECTIONS (SIG):", "Edit medication", CStr(wsM.Cells(r, C_SIG).Value))
        If StrPtr(t) <> 0 Then wsM.Cells(r, C_SIG).Value = Trim(t)
        t = InputBox("EXPIRATION (MM/YYYY):", "Edit medication", CStr(wsM.Cells(r, C_EXP).Value))
        If StrPtr(t) <> 0 Then wsM.Cells(r, C_EXP).Value = Trim(t)
        t = InputBox("LOT number:", "Edit medication", CStr(wsM.Cells(r, C_LOT).Value))
        If StrPtr(t) <> 0 Then wsM.Cells(r, C_LOT).Value = Trim(t)
    End If

    Call ValidateMedications(False)
    Call BuildAllLabelsPreview
End Sub

Private Function EditMedWithForm(wsM As Worksheet, r As Long) As Boolean
    EditMedWithForm = False
    Dim f As Object
    On Error GoTo Done
    Set f = VBA.UserForms.Add("frmMedEdit")
    If f Is Nothing Then Exit Function
    f.txtName.Value = CStr(wsM.Cells(r, C_NAME).Value)
    f.txtStr.Value = CStr(wsM.Cells(r, C_STR).Value)
    f.txtForm.Value = CStr(wsM.Cells(r, C_FORM).Value)
    f.txtQty.Value = CStr(wsM.Cells(r, C_QTY).Value)
    f.txtSig.Value = CStr(wsM.Cells(r, C_SIG).Value)
    f.txtExp.Value = CStr(wsM.Cells(r, C_EXP).Value)
    f.txtLot.Value = CStr(wsM.Cells(r, C_LOT).Value)
    f.Result = ""
    f.Show
    If f.Result = "OK" Then
        wsM.Cells(r, C_NAME).Value = Trim(f.txtName.Value)
        wsM.Cells(r, C_STR).Value = Trim(f.txtStr.Value)
        wsM.Cells(r, C_FORM).Value = Trim(f.txtForm.Value)
        wsM.Cells(r, C_QTY).Value = Trim(f.txtQty.Value)
        wsM.Cells(r, C_SIG).Value = Trim(f.txtSig.Value)
        wsM.Cells(r, C_EXP).Value = Trim(f.txtExp.Value)
        wsM.Cells(r, C_LOT).Value = Trim(f.txtLot.Value)
    End If
    EditMedWithForm = True
    Unload f
    Exit Function
Done:
    On Error Resume Next
    Unload f
    EditMedWithForm = False
End Function

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
                    If Trim(wsMed.Cells(r, C_EXP).Value) = "" Or Trim(wsMed.Cells(r, C_LOT).Value) = "" Then
                        Dim eVal As String, lVal As String
                        eVal = Trim(wsMed.Cells(r, C_EXP).Value)
                        lVal = Trim(wsMed.Cells(r, C_LOT).Value)
                        Call PromptExpLotPair(medLabel, eVal, lVal)
                        wsMed.Cells(r, C_EXP).Value = eVal
                        wsMed.Cells(r, C_LOT).Value = lVal
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

    ' Clear medication rows (fixed-range wipe - never misses data in unnamed rows)
    Call ClearMedArea(wsMed)

    ' Clear the entire dispense Log (full reset only; Start NEW Patient keeps it)
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)
    Dim lastLog As Long
    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastLog > LOG_HDR_ROWS Then
        With wsLog.Range(wsLog.Cells(LOG_HDR_ROWS + 1, 1), wsLog.Cells(lastLog, LG_LAST))
            .ClearContents
            .Interior.ColorIndex = xlNone
        End With
    End If
    Call ClearEncounterStore   ' wipe saved encounter snapshots on a full reset

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

    Call ClearMedArea(wsMed)

    ' Forget the remembered batch (belongs to the previous patient).
    gLastBatchRows = ""
    gLastBatchVol = ""

    wsIn.Activate
    wsIn.Range("C5").Select
    MsgBox "Ready for the next patient." & vbCrLf & "The dispense Log was kept.", _
           vbInformation, "New Patient"
End Sub

Public Sub ClearSessionSilent()
    ' Silent Start-NEW-Patient: clears patient + meds + paste, KEEPS the Log.
    ' Used by the workbook Open / BeforeClose auto-reset (no prompts).
    On Error Resume Next
    Dim wsIn As Worksheet, wsMed As Worksheet
    Set wsIn = ThisWorkbook.Sheets(SH_INPUT)
    Set wsMed = ThisWorkbook.Sheets(SH_MEDS)
    If wsIn Is Nothing Or wsMed Is Nothing Then Exit Sub
    wsIn.Range("C5").Value = ""
    wsIn.Range("C6").Value = ""
    wsIn.Range("C7").Value = Format(Now(), "MM/DD/YYYY")
    wsIn.Range("B12").Value = ""
    wsIn.Range("C12").Value = ""
    Call ClearMedArea(wsMed)
    ' Forget any remembered batch so "Reprint Last Batch" can never reprint a previous
    ' patient's labels after the session resets.
    gLastBatchRows = ""
    gLastBatchVol = ""
    On Error GoTo 0
End Sub

' Wipe the whole working medication area by a FIXED range (rows 3..503, all table columns),
' NOT by hunting for the last row that has a Name. Detecting the last Name row means that if
' leftover data sits in other columns but the Name cell is blank (e.g. a partial row, or old
' data left after a column reorder), the clear finds "nothing" and skips it - which is exactly
' why Reset Session could leave rows behind. A fixed range always clears them.
Private Sub ClearMedArea(wsMed As Worksheet)
    On Error Resume Next
    Dim rng As Range
    Set rng = wsMed.Range(wsMed.Cells(MEDS_HDR_ROWS + 1, 1), wsMed.Cells(MEDS_HDR_ROWS + 500, C_SEL))
    rng.ClearContents
    rng.Interior.ColorIndex = xlNone
    wsMed.Columns(C_EXP).NumberFormat = "@"
    wsMed.Columns(C_LOT).NumberFormat = "@"
    Call ApplySourceValidation(wsMed)
    On Error GoTo 0
End Sub

Public Sub ClearLogSilent()
    ' Silently clear the dispense Log (used by the on-close full auto-reset).
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)
    If wsLog Is Nothing Then Exit Sub
    Dim lastLog As Long
    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastLog > LOG_HDR_ROWS Then wsLog.Range(wsLog.Cells(LOG_HDR_ROWS + 1, 1), wsLog.Cells(lastLog, LG_LAST)).ClearContents
    Call ClearEncounterStore   ' encounters share the Log's lifecycle (kept on New Patient, wiped on full reset/close)
    On Error GoTo 0
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
    If Trim(rec.Refills) = "" Then          ' Refills default to 0 when none parsed
        ws.Cells(r, C_REF).Value = "0"
    Else
        ws.Cells(r, C_REF).Value = rec.Refills
    End If
    ws.Cells(r, C_EXP).Value   = rec.Expiration
    ws.Cells(r, C_LOT).Value   = rec.LotNumber
    ws.Cells(r, C_DATE).Value  = dateRx
    ws.Cells(r, C_CONF).Value  = rec.Confidence
    ws.Cells(r, C_WARN).Value  = rec.Warnings
    ws.Cells(r, C_RAW).Value   = rec.RawText
    ws.Cells(r, C_PRTD).Value  = "No"
    ws.Cells(r, C_SRC).Value   = ""           ' Source left blank on purpose - volunteer MUST pick one

    ' Apply row formatting
    Call ApplyRowState(ws, r)
End Sub

' Apply the Source dropdown (DOH / IN HOUSE / RxAPS / Other) to the Medications Source
' column across a generous row range, so every med row can pick a source.
Private Sub ApplySourceValidation(ws As Worksheet)
    On Error Resume Next
    ' Strip any stray dropdowns from the whole table first so the Source list is the ONLY
    ' data-validation in the Medications grid (guards against a pre-reorder build that had
    ' the dropdown on the old far-right column).
    ws.Range(ws.Cells(MEDS_HDR_ROWS + 1, 1), ws.Cells(MEDS_HDR_ROWS + 500, C_SEL)).Validation.Delete
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(MEDS_HDR_ROWS + 1, C_SRC), ws.Cells(MEDS_HDR_ROWS + 500, C_SRC))
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="DOH,IN HOUSE,RxAPS,Other"
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
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
                ElseIf c = C_EXP And IsBadExpFormat(CStr(.Value)) Then
                    .Interior.Color = RGB(255, 224, 130)    ' filled but wrong format - amber
                Else
                    .Interior.Color = bg                     ' filled + valid - blends with row
                End If
            ElseIf c = C_QTY Then
                If Trim(.Value) = "" Then
                    .Interior.Color = RGB(255, 241, 118)    ' Quantity missing - yellow
                Else
                    .Interior.Color = bg
                End If
            ElseIf c = C_SRC Then
                .HorizontalAlignment = xlCenter
                If Trim(.Value) = "" Then
                    .Interior.Color = RGB(255, 241, 118)    ' Source required - yellow until picked
                Else
                    .Interior.Color = bg
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

    ' Tidy alignment: short data columns centered, text columns left, all vertically centered
    Dim ac As Variant
    For Each ac In Array(C_NUM, C_QTY, C_EXP, C_LOT, C_SRC, C_DATE, C_REF, C_CNT, C_SEL)
        ws.Cells(r, CLng(ac)).HorizontalAlignment = xlCenter
    Next ac
    For Each ac In Array(C_NAME, C_STR, C_FORM, C_SIG)
        ws.Cells(r, CLng(ac)).HorizontalAlignment = xlLeft
    Next ac
    ws.Range(ws.Cells(r, 1), ws.Cells(r, C_SEL)).VerticalAlignment = xlCenter

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
    Call ApplySourceValidation(ws)   ' keep the Source dropdown present after any refresh
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
        "    If Target.Column = " & C_CNT & " Then" & vbCrLf & _
        "        Cancel = True" & vbCrLf & _
        "        Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    ' Print? column - toggle the selection check" & vbCrLf & _
        "    If Target.Column = " & C_SEL & " And Target.Row > " & MEDS_HDR_ROWS & " Then" & vbCrLf & _
        "        If Trim(Me.Cells(Target.Row, " & C_NAME & ").Value) <> """" Then" & vbCrLf & _
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
    If MsgBox("About to print these " & cnt & " medication(s), " & LABEL_COPIES & " copies each, on the Brother QL-1100c:" & vbCrLf & vbCrLf & _
              listMsg & vbCrLf & _
              "Make sure the DK-1202 (62 x 100 mm) roll is loaded." & vbCrLf & vbCrLf & _
              "YES = print all       NO = cancel", _
              vbYesNo + vbQuestion, "Print Checked Labels") = vbNo Then Exit Sub

    Call BusyShow(15, "Locating the Brother QL-1100c...")
    Dim brother As String
    brother = SelectBrotherPrinter()
    If brother = "" Then
        Call BusyHide
        MsgBox "Brother QL-1100c not found. Make sure it is plugged in, powered on," & vbCrLf & _
               "and loaded with the DK-1202 roll, then try again.", _
               vbExclamation, "Printer Not Found"
        Exit Sub
    End If

    Call BusyShow(65, "Preparing the label page...")
    Call ApplyLabelContentWidth(wsL)
    Call ApplyLabelPageSetup(wsL)
    Call RefreshPrintLabelLogo(wsL)      ' insert the logo once; the loop reuses it
    Call BusyShow(100, "Ready.")
    Call BusyHide

    Dim batchVol As String
    batchVol = AskInitials()
    Dim encNum As Long
    encNum = NextEncounter()          ' this whole batch = one patient encounter
    Dim done As Integer, skipped As Integer
    Dim skippedNames As String, printedRows As String
    done = 0
    skipped = 0
    skippedNames = ""
    printedRows = ""
    Application.ScreenUpdating = False
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" And IsRowSelected(ws, r) Then
            If Trim(ws.Cells(r, C_EXP).Value) = "" Or Trim(ws.Cells(r, C_LOT).Value) = "" Then
                skipped = skipped + 1
                skippedNames = skippedNames & "   - " & _
                    Trim(Trim(ws.Cells(r, C_NAME).Value) & " " & Trim(ws.Cells(r, C_STR).Value)) & vbCrLf
            Else
                Call UpdateLabelPreviewForMedRow(r, False)
                If PrintLabelSurfaceSafe(LABEL_COPIES) Then
                    Call MarkPrinted(r)
                    Call LogPrint(r, batchVol, encNum)
                    Call ApplyRowState(ws, r)
                    done = done + 1
                    printedRows = printedRows & r & ","
                End If
            End If
        End If
    Next r
    Application.ScreenUpdating = True

    ' Snapshot this encounter's full med list so it can be reopened + edited later.
    If done > 0 Then Call SaveEncounterSnapshot(encNum)

    ' Remember this batch so "Reprint Last Batch" can re-run it after a jam/misfeed.
    gLastBatchRows = printedRows
    gLastBatchVol = batchVol
    Dbg "PrintCheckedLabels: done=" & done & " skipped=" & skipped

    Call ShowLogSheet      ' land on the dispense Log after printing

    Dim msg As String, icon As Integer
    msg = done & " medication(s) sent to the Brother, " & LABEL_COPIES & " copies each  (" & (done * LABEL_COPIES) & " labels)."
    icon = vbInformation
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & skipped & " SKIPPED (missing Expiration or Lot):" & vbCrLf & _
              skippedNames & vbCrLf & _
              "Add the missing Exp/Lot on the Medications tab, then reprint."
        icon = vbExclamation
    End If
    MsgBox msg, icon, "Print Complete"
End Sub

' Reprint the last successfully-printed batch (for paper jams / misfeeds). Reuses the
' same rows and volunteer initials; logs each as a reprint. Session-only.
Public Sub ReprintLastBatch()
    If Trim(gLastBatchRows) = "" Then
        MsgBox "There is no batch to reprint yet." & vbCrLf & _
               "Use 'Print Checked Labels' first, then this button can reprint it.", _
               vbInformation, "Reprint Last Batch"
        Exit Sub
    End If

    Dim ws As Worksheet, wsL As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)
    Set wsL = ThisWorkbook.Sheets(SH_LABEL)

    Dim rowArr() As String
    rowArr = Split(gLastBatchRows, ",")
    Dim i As Long, nRows As Integer
    nRows = 0
    For i = LBound(rowArr) To UBound(rowArr)
        If Trim(rowArr(i)) <> "" Then nRows = nRows + 1
    Next i

    If MsgBox("Reprint the last " & nRows & " medication(s), " & LABEL_COPIES & " copies each, on the Brother QL-1100c?" & vbCrLf & vbCrLf & _
              "Make sure the DK-1202 (62 x 100 mm) roll is loaded." & vbCrLf & vbCrLf & _
              "YES = reprint       NO = cancel", _
              vbYesNo + vbQuestion, "Reprint Last Batch") = vbNo Then Exit Sub

    Call BusyShow(15, "Locating the Brother QL-1100c...")
    Dim brother As String
    brother = SelectBrotherPrinter()
    If brother = "" Then
        Call BusyHide
        MsgBox "Brother QL-1100c not found. Make sure it is plugged in, powered on," & vbCrLf & _
               "and loaded with the DK-1202 roll, then try again.", _
               vbExclamation, "Printer Not Found"
        Exit Sub
    End If
    Call BusyShow(65, "Preparing the label page...")
    Call ApplyLabelContentWidth(wsL)
    Call ApplyLabelPageSetup(wsL)
    Call RefreshPrintLabelLogo(wsL)      ' insert the logo once; the loop reuses it
    Call BusyShow(100, "Ready.")
    Call BusyHide

    Dim r As Long, done As Integer
    done = 0
    Dim encNum As Long
    encNum = NextEncounter()          ' a reprint is logged as its own encounter
    Application.ScreenUpdating = False
    For i = LBound(rowArr) To UBound(rowArr)
        If Trim(rowArr(i)) <> "" Then
            r = CLng(rowArr(i))
            If Trim(ws.Cells(r, C_NAME).Value) <> "" Then
                Call UpdateLabelPreviewForMedRow(r, False)
                If PrintLabelSurfaceSafe(LABEL_COPIES) Then
                    Call MarkPrinted(r)
                    Call LogPrint(r, Trim(gLastBatchVol & " (reprint)"), encNum)
                    done = done + 1
                End If
            End If
        End If
    Next i
    Application.ScreenUpdating = True

    If done > 0 Then Call SaveEncounterSnapshot(encNum)   ' snapshot the reprint as its own encounter

    Call ShowLogSheet      ' land on the dispense Log after printing
    MsgBox done & " medication(s) reprinted, " & LABEL_COPIES & " copies each  (" & (done * LABEL_COPIES) & " labels).", vbInformation, "Reprint Complete"
End Sub

' Show the dispense Log sheet and scroll to the most recent entry (called after printing).
Private Sub ShowLogSheet()
    On Error Resume Next
    Dim wsLg As Worksheet
    Set wsLg = ThisWorkbook.Sheets(SH_LOG)
    wsLg.Activate
    Dim lastLg As Long
    lastLg = wsLg.Cells(wsLg.Rows.Count, 1).End(xlUp).Row
    If lastLg < 1 Then lastLg = 1
    wsLg.Cells(lastLg, 1).Select
    On Error GoTo 0
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
    ' True if any comma-separated expiration is not MM/YYYY or MM/DD/YYYY AFTER
    ' normalization (so ".", "-", spaces and 2-digit years are tolerated - they get
    ' standardized, not rejected). Multiple values (e.g. two bottles) are each checked.
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(0[1-9]|1[0-2])/(\d{4})$|^(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])/(\d{4})$"
    Dim parts() As String, i As Long, p As String
    parts = Split(NormalizeExp(s), ",")
    IsBadExpFormat = False
    For i = LBound(parts) To UBound(parts)
        p = Trim(parts(i))
        If p = "" Then IsBadExpFormat = True: Exit Function
        If Not re.Test(p) Then IsBadExpFormat = True: Exit Function
    Next i
End Function

' Standardize expiration input. Splits on commas (multiple bottles), and for each part
' converts ".", "-" and spaces to "/", zero-pads the month/day, and expands 2-digit
' years to 4 digits. Keeps MM/YYYY, or MM/DD/YYYY when a day was entered. Any part that
' is not a recognizable date is passed through unchanged (validation then flags it).
Public Function NormalizeExp(ByVal raw As String) As String
    Dim parts() As String, i As Long, outp As String
    parts = Split(raw, ",")
    outp = ""
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then outp = outp & ", "
        outp = outp & NormExpPart(Trim(parts(i)))
    Next i
    NormalizeExp = outp
End Function

Private Function NormExpPart(ByVal s As String) As String
    NormExpPart = s
    If s = "" Then Exit Function
    Dim t As String
    t = s
    t = Replace(t, ".", "/")
    t = Replace(t, "-", "/")
    t = Replace(t, " ", "/")
    Do While InStr(t, "//") > 0
        t = Replace(t, "//", "/")
    Loop
    Dim comp() As String, k As Long
    comp = Split(t, "/")
    For k = LBound(comp) To UBound(comp)
        If comp(k) = "" Or Not IsNumeric(comp(k)) Then Exit Function   ' not a date -> leave original
    Next k
    If UBound(comp) = 1 Then
        NormExpPart = ExpPad2(comp(0)) & "/" & ExpYear4(comp(1))
    ElseIf UBound(comp) = 2 Then
        NormExpPart = ExpPad2(comp(0)) & "/" & ExpPad2(comp(1)) & "/" & ExpYear4(comp(2))
    End If
End Function

Private Function ExpPad2(ByVal s As String) As String
    Dim n As Long
    n = CLng(Val(s))
    ExpPad2 = Right("0" & CStr(n), 2)
End Function

Private Function ExpYear4(ByVal s As String) As String
    Dim n As Long
    n = CLng(Val(s))
    If Len(CStr(n)) <= 2 Then
        ExpYear4 = CStr(2000 + (n Mod 100))
    Else
        ExpYear4 = CStr(n)
    End If
End Function

Private Function PromptExpiration(medLabel As String) As String
    ' Prompt for MM/YYYY expiration. If the format looks wrong, offer a correction
    ' but still ALLOW the entered value - never blocks the volunteer.
    Dim val As String
    val = ""
    Do
        val = Trim(InputBox( _
            "Enter EXPIRATION DATE for:" & vbCrLf & medLabel & vbCrLf & _
            "(format: MM/YYYY  -  check the bottle; use commas for multiple bottles)", _
            "Expiration Required", val))
        val = NormalizeExp(val)
        If val = "" Then Exit Do
        If Not IsBadExpFormat(val) Then Exit Do
        If MsgBox("'" & val & "' does not look like MM/YYYY (for example 05/2027)." & vbCrLf & vbCrLf & _
                  "Yes  =  re-enter it" & vbCrLf & _
                  "No  =  keep '" & val & "' as typed", _
                  vbYesNo + vbExclamation, "Check Expiration Format") = vbNo Then Exit Do
    Loop
    PromptExpiration = val
End Function

Private Sub PromptExpLotPair(medLabel As String, ByRef expVal As String, ByRef lotVal As String)
    ' Two-field popup (frmExpLot, built by Build-Release) asking a med's Expiration AND
    ' Lot together. If the form is not present, fall back to two plain input boxes.
    Dim f As Object, ok As Boolean
    ok = False
    On Error Resume Next
    Set f = VBA.UserForms.Add("frmExpLot")
    If Not (f Is Nothing) Then
        Do
            f.lblMed.Caption = medLabel
            f.txtExp.Value = expVal
            f.txtLot.Value = lotVal
            Err.Clear
            f.Show
            If Err.Number <> 0 Then Exit Do
            expVal = Trim(f.txtExp.Value)
            lotVal = Trim(f.txtLot.Value)
            If Err.Number <> 0 Then Exit Do
            ok = True
            expVal = NormalizeExp(expVal)      ' standardize separators / years, keep commas
            If expVal = "" Then Exit Do
            If Not IsBadExpFormat(expVal) Then Exit Do
            If MsgBox("'" & expVal & "' does not look like MM/YYYY (for example 05/2027)." & vbCrLf & vbCrLf & _
                      "Yes  =  re-enter it" & vbCrLf & "No  =  keep it as typed", _
                      vbYesNo + vbExclamation, "Check Expiration Format") = vbNo Then Exit Do
        Loop
        Unload f
    End If
    On Error GoTo 0
    If Not ok Then
        If Trim(expVal) = "" Then expVal = PromptExpiration(medLabel)
        If Trim(lotVal) = "" Then lotVal = Trim(InputBox( _
            "Enter LOT NUMBER for:" & vbCrLf & medLabel & vbCrLf & "(check the bottle or package)", _
            "Lot Number Required", lotVal))
    End If
End Sub

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
        If Trim(ws.Cells(r, C_SRC).Value) = "" Then w = w & "SOURCE missing (pick DOH / IN HOUSE / RxAPS / Other). "
        If Trim(ws.Cells(r, C_EXP).Value) <> "" Then
            ws.Cells(r, C_EXP).NumberFormat = "@"
            ws.Cells(r, C_EXP).Value = NormalizeExp(CStr(ws.Cells(r, C_EXP).Value))
            If IsBadExpFormat(CStr(ws.Cells(r, C_EXP).Value)) Then _
                w = w & "Check expiration format (use MM/YYYY, comma-separate multiples). "
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
                ws.Cells(r, C_SEL).Value = ChrW(10003)   ' auto-check passing meds for printing
                Call ApplyRowState(ws, r)
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
        foot = vbCrLf & "All rows look complete - checked and ready to print."
    Else
        foot = vbCrLf & flagged & " row(s) still need attention (red / orange cells)."
    End If
    foot = foot & vbCrLf & "Passing meds are auto-checked for printing (uncheck any you don't want)."
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

    ' Offer to collect Expiration + Lot now, exactly like the Parse flow does.
    ws.Cells(rr, C_EXP).NumberFormat = "@"
    ws.Cells(rr, C_LOT).NumberFormat = "@"
    If MsgBox("Enter EXPIRATION and LOT for " & nm & " now?" & vbCrLf & _
              "(No = fill the highlighted cells on the sheet later.)", _
              vbYesNo + vbQuestion, "Enter Exp / Lot?") = vbYes Then
        Dim eVal As String, lVal As String
        eVal = ""
        lVal = ""
        Call PromptExpLotPair(Trim(nm & " " & rec.Strength), eVal, lVal)
        ws.Cells(rr, C_EXP).Value = eVal
        ws.Cells(rr, C_LOT).Value = lVal
    End If

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

    ' Delete bottom-to-top so row indexes stay valid. Delete the table columns
    ' (1..C_SEL, the full row incl. Source) and shift up, so the side buttons (col 19+)
    ' do not move.
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
    Call ApplySourceValidation(ws)
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
Public Sub UpdateLabelPreviewForMedRow(ByVal medRow As Long, Optional ByVal refreshChrome As Boolean = True)
    Dim wsM As Worksheet, wsL As Worksheet, wsI As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsL = ThisWorkbook.Sheets(SH_LABEL)
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    If medRow <= MEDS_HDR_ROWS Then Exit Sub
    If Trim(wsM.Cells(medRow, C_NAME).Value) = "" Then Exit Sub

    Dim patName As String: patName = Trim(wsI.Range("C5").Value)
    Dim dob As String: dob = Trim(wsI.Range("C6").Value)
    Dim dateRx As String: dateRx = Trim(wsM.Cells(medRow, C_DATE).Value)
    Dim medName As String: medName = Trim(wsM.Cells(medRow, C_NAME).Value)
    Dim strength As String: strength = Trim(wsM.Cells(medRow, C_STR).Value)
    Dim formTxt As String: formTxt = Trim(wsM.Cells(medRow, C_FORM).Value)
    Dim qty As String: qty = Trim(wsM.Cells(medRow, C_QTY).Value)
    Dim sig As String: sig = Trim(wsM.Cells(medRow, C_SIG).Value)
    Dim expDate As String: expDate = Trim(wsM.Cells(medRow, C_EXP).Value)
    Dim lotNum As String: lotNum = Trim(wsM.Cells(medRow, C_LOT).Value)

    Dim medLine As String
    medLine = medName
    If strength <> "" Then medLine = medLine & " " & strength
    Dim refills As String: refills = Trim(wsM.Cells(medRow, C_REF).Value)
    Dim fq As String
    fq = formTxt
    If qty <> "" Then
        If fq <> "" Then
            fq = fq & "   " & Chr(183) & "   Qty " & qty
        Else
            fq = "Qty " & qty
        End If
    End If
    If refills <> "" Then
        If fq <> "" Then
            fq = fq & "   " & Chr(183) & "   Refills " & refills
        Else
            fq = "Refills " & refills
        End If
    End If

    Dim pn As String: pn = IIf(patName <> "", patName, "[Patient Name]")
    wsL.Cells(5, 1).Value = pn
    wsL.Cells(5, 1).Font.Size = 12       ' Name/DOB a bit smaller than the medication line
    wsL.Cells(5, 5).Value = "DOB  " & IIf(dob <> "", dob, "--")
    wsL.Cells(5, 5).Font.Size = 12

    wsL.Cells(7, 1).Value = medLine
    wsL.Cells(7, 1).Font.Bold = True
    wsL.Cells(7, 1).WrapText = True
    If Len(medLine) > MED_WRAP_MAXLEN Then
        ' Long name: wrap to two lines at a readable size (uses the white space in
        ' the med band) instead of shrinking to a tiny single line. Row 7 grows and
        ' the spacer rows 13/14 shrink so the A1:H15 print area height is preserved.
        wsL.Cells(7, 1).Font.Size = MED_WRAP_FONT
        wsL.Rows(7).RowHeight = 28
        wsL.Rows(13).RowHeight = 1
        wsL.Rows(14).RowHeight = 1
    Else
        wsL.Cells(7, 1).Font.Size = MedFontSize(medLine)
        wsL.Rows(7).RowHeight = 20
        wsL.Rows(13).RowHeight = 3
        wsL.Rows(14).RowHeight = 2
    End If

    wsL.Cells(8, 1).Value = fq
    wsL.Cells(8, 6).Value = "Rx  " & IIf(dateRx <> "", dateRx, "--")

    Dim sigText As String
    sigText = IIf(sig <> "", sig, "[Directions not found - enter manually]")
    wsL.Cells(10, 1).Value = sigText
    wsL.Cells(10, 1).Font.Color = RGB(255, 255, 255)
    wsL.Cells(10, 1).Font.Bold = True
    Dim sz As Single
    If Len(sigText) <= 80 Then
        sz = 11
    ElseIf Len(sigText) <= 130 Then
        sz = 10
    Else
        sz = 9
    End If
    wsL.Cells(10, 1).Font.Size = sz

    Call SetMiniValue(wsL.Cells(15, 1), "EXP", IIf(expDate <> "", expDate, "--"), 12, "L")
    Call SetMiniValue(wsL.Cells(15, 5), "LOT", IIf(lotNum <> "", lotNum, "--"), 12, "R")

    ' Chrome (column widths + logo + header merges) is identical on every label. Batch
    ' printing sets it once up-front and passes refreshChrome:=False here to skip the
    ' costly per-label logo re-insert from disk. Single-label preview keeps the default.
    If refreshChrome Then
        Call ApplyLabelContentWidth(wsL)
        Call RefreshPrintLabelLogo(wsL)
    End If
End Sub

Public Sub UpdateLabelPreviewFromSelection()
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Dim selRow As Long
    If ActiveSheet.name = SH_MEDS Then
        selRow = ActiveCell.Row
    Else
        selRow = MEDS_HDR_ROWS + 1
    End If
    If selRow <= MEDS_HDR_ROWS Then
        MsgBox "Please click on a medication row in the Medications tab first.", _
               vbExclamation, "No Row Selected"
        Exit Sub
    End If
    If Trim(wsM.Cells(selRow, C_NAME).Value) = "" Then
        MsgBox "Selected row appears to be empty. Click on a medication name cell first.", _
               vbExclamation, "Empty Row"
        Exit Sub
    End If
    Dim wsI As Worksheet, missing As String
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    missing = ""
    If Trim(wsI.Range("C5").Value) = "" Then missing = missing & "Patient Name  "
    If Trim(wsI.Range("C6").Value) = "" Then missing = missing & "DOB  "
    If Trim(wsM.Cells(selRow, C_EXP).Value) = "" Then missing = missing & "Expiration  "
    If Trim(wsM.Cells(selRow, C_LOT).Value) = "" Then missing = missing & "Lot Number  "
    If missing <> "" Then
        If MsgBox("The following fields are still empty:" & vbCrLf & missing & vbCrLf & vbCrLf & _
                  "Print anyway?", vbYesNo + vbExclamation, "Missing Label Fields") = vbNo Then Exit Sub
    End If
    Call UpdateLabelPreviewForMedRow(selRow)
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

Private Function PrintLabelSurfaceSafe(Optional ByVal copies As Long = 1) As Boolean
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets(SH_LABEL)
    Call PrepareLabelSheetForPrint(ws)
    ws.Visible = xlSheetVisible
    ws.Activate
    ' From/To 1: Brother QL treats each Excel page as one die-cut; avoid a blank 2nd page.
    ws.PrintOut From:=1, To:=1, Copies:=copies, Collate:=True
    PrintLabelSurfaceSafe = True
CleanExit:
    On Error Resume Next
    ThisWorkbook.Sheets(SH_MEDS).Activate
    ws.Visible = xlSheetHidden
    On Error GoTo 0
    Exit Function
Fail:
    PrintLabelSurfaceSafe = False
    Resume CleanExit
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

    Call ApplyLabelContentWidth(wsL)
    Call RefreshPrintLabelLogo(wsL)
    Call ApplyLabelPageSetup(wsL)

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
        Call LogPrint(medRowToMark, AskInitials(), NextEncounter())
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

    Call LogPrint(medRowToMark, AskInitials(), NextEncounter())

    If Not PrintLabelSurfaceSafe(LABEL_COPIES) Then
        MsgBox "Printing failed. Check the printer connection and that a label roll" & vbCrLf & _
               "is loaded, then try again.", vbExclamation, "Print Error"
        Exit Sub
    End If
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
    ' Session cache (B7): the WMI lookup + port probing is the slow part, so once we know
    ' the Brother's name we reuse it and only re-detect if selecting it fails.
    If gCachedPrinter <> "" Then
        If SetActivePrinterByName(gCachedPrinter) Then
            SelectBrotherPrinter = gCachedPrinter
            Dbg "SelectBrotherPrinter: cache hit '" & gCachedPrinter & "'"
            Exit Function
        End If
        gCachedPrinter = ""      ' cached printer no longer selectable -> full re-detect
    End If
    Dim nm As String
    nm = DetectBrotherPrinter()
    gCachedPrinter = nm          ' "" if not found (harmless; re-detected next time)
    Dbg "SelectBrotherPrinter: detected '" & nm & "'"
    SelectBrotherPrinter = nm
End Function

Private Function DetectBrotherPrinter() As String
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
            DetectBrotherPrinter = foundName
            Exit Function
        End If
        Err.Clear
        Application.ActivePrinter = foundName & " on " & foundPort
        If Err.Number = 0 Then
            On Error GoTo 0
            DetectBrotherPrinter = foundName
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
            DetectBrotherPrinter = candidates(i)
            Exit Function
        End If
    Next i

    DetectBrotherPrinter = ""
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

' The next encounter number = one more than the highest already in the Log (col 16).
' Derived from the Log itself (not a module counter) so it survives across prints and a
' macro reset, and every med in one print batch is stamped with the same value.
Private Function NextEncounter() As Long
    On Error Resume Next
    Dim wsLg As Worksheet, lastLog As Long, r As Long, mx As Long
    Set wsLg = ThisWorkbook.Sheets(SH_LOG)
    lastLog = wsLg.Cells(wsLg.Rows.Count, 1).End(xlUp).Row
    mx = 0
    For r = LOG_HDR_ROWS + 1 To lastLog
        If IsNumeric(wsLg.Cells(r, LG_ENC).Value) Then
            If CLng(wsLg.Cells(r, LG_ENC).Value) > mx Then mx = CLng(wsLg.Cells(r, LG_ENC).Value)
        End If
    Next r
    NextEncounter = mx + 1
End Function

Private Sub LogPrint(ByVal medRow As Long, ByVal vol As String, ByVal encounter As Long)
    On Error Resume Next
    Dim wsI As Worksheet, wsM As Worksheet, wsLg As Worksheet
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsLg = ThisWorkbook.Sheets(SH_LOG)

    Dim nextLog As Long
    nextLog = wsLg.Cells(wsLg.Rows.Count, 1).End(xlUp).Row + 1
    If nextLog <= LOG_HDR_ROWS Then nextLog = LOG_HDR_ROWS + 1

    wsLg.Cells(nextLog, LG_TIME).Value = Format(Now(), "MM/DD/YYYY HH:MM:SS")
    wsLg.Cells(nextLog, LG_ENC).Value = encounter                        ' Encounter #, right after Timestamp
    wsLg.Cells(nextLog, LG_PT).Value = Trim(wsI.Range("C5").Value)
    wsLg.Cells(nextLog, LG_DOB).Value = Trim(wsI.Range("C6").Value)
    If medRow > MEDS_HDR_ROWS Then
        wsLg.Cells(nextLog, LG_NAME).Value = wsM.Cells(medRow, C_NAME).Value
        wsLg.Cells(nextLog, LG_STR).Value = wsM.Cells(medRow, C_STR).Value
        wsLg.Cells(nextLog, LG_SIG).Value = wsM.Cells(medRow, C_SIG).Value
        wsLg.Cells(nextLog, LG_QTY).Value = wsM.Cells(medRow, C_QTY).Value
        wsLg.Cells(nextLog, LG_REF).Value = wsM.Cells(medRow, C_REF).Value
        wsLg.Cells(nextLog, LG_EXP).NumberFormat = "@"
        wsLg.Cells(nextLog, LG_EXP).Value = wsM.Cells(medRow, C_EXP).Value
        wsLg.Cells(nextLog, LG_LOT).NumberFormat = "@"
        wsLg.Cells(nextLog, LG_LOT).Value = wsM.Cells(medRow, C_LOT).Value
        wsLg.Cells(nextLog, LG_SRC).Value = wsM.Cells(medRow, C_SRC).Value
        wsLg.Cells(nextLog, LG_DATE).Value = wsM.Cells(medRow, C_DATE).Value
        wsLg.Cells(nextLog, LG_FORM).Value = wsM.Cells(medRow, C_FORM).Value
        wsLg.Cells(nextLog, LG_CNT).Value = wsM.Cells(medRow, C_CNT).Value
    End If
    wsLg.Cells(nextLog, LG_INIT).Value = vol                             ' Initials

    ' Shade the whole row by encounter, cycling 3 greens, so each patient's print reads
    ' as one clearly-bounded block in the Log.
    Dim encColor As Long
    Select Case (encounter - 1) Mod 3
        Case 0:    encColor = RGB(232, 245, 233)   ' lightest green
        Case 1:    encColor = RGB(200, 230, 201)   ' light green
        Case Else: encColor = RGB(165, 214, 167)   ' medium green
    End Select

    Dim c As Integer
    For c = 1 To LG_LAST
        With wsLg.Cells(nextLog, c)
            .Font.Name = "Arial"
            .Font.Size = 9
            .Interior.Color = encColor
        End With
    Next c
    wsLg.Cells(nextLog, LG_ENC).HorizontalAlignment = xlCenter
    wsLg.Cells(nextLog, LG_ENC).Font.Bold = True

    ' Mirror this row to the dated local CSV archive (best-effort; never blocks printing).
    Call ArchiveDispenseRow(wsLg, nextLog)
End Sub

' Append one dispense-Log row to a dated local CSV so the day's record survives the
' on-close Log wipe. The file lives in a "dispense-log" folder next to the workbook and
' is git-ignored (it contains PHI and must never leave this machine). Best-effort: any
' failure here is swallowed so it can never interrupt printing.
Private Sub ArchiveDispenseRow(wsLg As Worksheet, logRow As Long)
    On Error Resume Next
    If Not DISPENSE_CSV_ENABLED Then Exit Sub
    Dim basePath As String
    basePath = ThisWorkbook.Path
    If basePath = "" Then Exit Sub          ' workbook never saved -> no place to write

    Dim folder As String
    folder = basePath & "\dispense-log"
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    Dim fpath As String
    fpath = folder & "\" & Format(Date, "YYYY-MM-DD") & ".csv"
    Dim isNew As Boolean
    isNew = (Dir(fpath) = "")

    Dim ff As Integer
    ff = FreeFile
    Open fpath For Append As #ff
    If isNew Then
        Print #ff, "Timestamp,Encounter,Patient,DOB,Medication,Strength,Directions,Qty,Refills,Expiration,Lot,Source,RxDate,Initials,DosageForm,PrintCount"
    End If
    Dim ln As String, c As Integer
    ln = ""
    For c = 1 To LG_LAST
        If c > 1 Then ln = ln & ","
        ln = ln & CsvField(CStr(wsLg.Cells(logRow, c).Value))
    Next c
    Print #ff, ln
    Close #ff
    On Error GoTo 0
End Sub

' Quote one value as an RFC-4180 CSV field (double internal quotes; flatten newlines).
Private Function CsvField(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, """", """""")
    CsvField = """" & s & """"
End Function

' ============================================================
'  ENCOUNTER SNAPSHOTS + EDITING
'  Each print saves a FULL snapshot of the patient's med list under its Encounter # (hidden
'  SH_ENC sheet). A past encounter can be reopened, edited (add / remove / fix meds), then
'  re-saved - which REPLACES that encounter's Log + snapshot rows (no duplicates) and
'  refreshes the Tebra note. The dated CSV keeps history (append-only).
' ============================================================

' The hidden snapshot sheet (created on demand) with its header row.
Private Function EncStore() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_ENC)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SH_ENC
    End If
    ws.Cells(1, ES_ENC).Value = "Enc"
    ws.Cells(1, ES_PT).Value = "Patient"
    ws.Cells(1, ES_DOB).Value = "DOB"
    ws.Cells(1, ES_RXDATE).Value = "RxDate"
    ws.Cells(1, ES_NAME).Value = "Name"
    ws.Cells(1, ES_STR).Value = "Strength"
    ws.Cells(1, ES_FORM).Value = "Form"
    ws.Cells(1, ES_SIG).Value = "SIG"
    ws.Cells(1, ES_QTY).Value = "Qty"
    ws.Cells(1, ES_EXP).Value = "Exp"
    ws.Cells(1, ES_LOT).Value = "Lot"
    ws.Cells(1, ES_SRC).Value = "Source"
    ws.Cells(1, ES_DATE).Value = "Date"
    ws.Cells(1, ES_REF).Value = "Refills"
    ws.Visible = xlSheetHidden
    Set EncStore = ws
End Function

' Delete every data row on ws (below its header) whose column encCol = encNum. Bottom-up.
Private Sub DeleteRowsByEncounter(ws As Worksheet, ByVal encCol As Long, ByVal encNum As Long, ByVal firstDataRow As Long)
    Dim last As Long, r As Long
    last = ws.Cells(ws.Rows.Count, encCol).End(xlUp).Row
    For r = last To firstDataRow Step -1
        If IsNumeric(ws.Cells(r, encCol).Value) Then
            If CLng(ws.Cells(r, encCol).Value) = encNum Then ws.Rows(r).Delete
        End If
    Next r
End Sub

' Snapshot the CURRENT patient + every named medication row under encounter encNum.
Private Sub SaveEncounterSnapshot(ByVal encNum As Long)
    On Error Resume Next
    Dim wsE As Worksheet, wsM As Worksheet, wsI As Worksheet
    Set wsE = EncStore()
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    Call DeleteRowsByEncounter(wsE, ES_ENC, encNum, 2)   ' idempotent: clear any prior snapshot
    Dim pt As String, dob As String, rx As String
    pt = Trim(wsI.Range("C5").Value)
    dob = Trim(wsI.Range("C6").Value)
    rx = Trim(wsI.Range("C7").Value)
    Dim lastMed As Long, r As Long, nr As Long
    lastMed = wsM.Cells(wsM.Rows.Count, C_NAME).End(xlUp).Row
    For r = MEDS_HDR_ROWS + 1 To lastMed
        If Trim(wsM.Cells(r, C_NAME).Value) <> "" Then
            nr = wsE.Cells(wsE.Rows.Count, ES_ENC).End(xlUp).Row + 1
            If nr < 2 Then nr = 2
            wsE.Cells(nr, ES_ENC).Value = encNum
            wsE.Cells(nr, ES_PT).Value = pt
            wsE.Cells(nr, ES_DOB).Value = dob
            wsE.Cells(nr, ES_RXDATE).Value = rx
            wsE.Cells(nr, ES_NAME).Value = wsM.Cells(r, C_NAME).Value
            wsE.Cells(nr, ES_STR).Value = wsM.Cells(r, C_STR).Value
            wsE.Cells(nr, ES_FORM).Value = wsM.Cells(r, C_FORM).Value
            wsE.Cells(nr, ES_SIG).Value = wsM.Cells(r, C_SIG).Value
            wsE.Cells(nr, ES_QTY).Value = wsM.Cells(r, C_QTY).Value
            wsE.Cells(nr, ES_EXP).NumberFormat = "@"
            wsE.Cells(nr, ES_EXP).Value = wsM.Cells(r, C_EXP).Value
            wsE.Cells(nr, ES_LOT).NumberFormat = "@"
            wsE.Cells(nr, ES_LOT).Value = wsM.Cells(r, C_LOT).Value
            wsE.Cells(nr, ES_SRC).Value = wsM.Cells(r, C_SRC).Value
            wsE.Cells(nr, ES_DATE).Value = wsM.Cells(r, C_DATE).Value
            wsE.Cells(nr, ES_REF).Value = wsM.Cells(r, C_REF).Value
        End If
    Next r
    On Error GoTo 0
End Sub

' Wipe all snapshots (full reset / on-close only; kept across Start NEW Patient).
Private Sub ClearEncounterStore()
    On Error Resume Next
    Dim wsE As Worksheet, last As Long
    Set wsE = ThisWorkbook.Sheets(SH_ENC)
    If Not wsE Is Nothing Then
        last = wsE.Cells(wsE.Rows.Count, ES_ENC).End(xlUp).Row
        If last >= 2 Then wsE.Range(wsE.Rows(2), wsE.Rows(last)).ClearContents
    End If
    gEditingEncounter = 0
    On Error GoTo 0
End Sub

' Reopen a past encounter: show a list, pick a number, restore the patient + full med list.
Public Sub EditEncounter()
    Dim wsE As Worksheet
    Set wsE = EncStore()
    Dim last As Long
    last = wsE.Cells(wsE.Rows.Count, ES_ENC).End(xlUp).Row
    If last < 2 Then
        MsgBox "There are no saved encounters to edit yet." & vbCrLf & _
               "An encounter is saved each time you Print Checked Labels.", vbInformation, "Edit Encounter"
        Exit Sub
    End If
    ' Distinct encounter numbers, in order of first appearance.
    Dim seen As String, order As String, r As Long, e As Long
    seen = "|"
    order = ""
    For r = 2 To last
        If IsNumeric(wsE.Cells(r, ES_ENC).Value) Then
            e = CLng(wsE.Cells(r, ES_ENC).Value)
            If InStr(seen, "|" & e & "|") = 0 Then
                seen = seen & e & "|"
                order = order & e & ","
            End If
        End If
    Next r
    Dim listTxt As String, parts() As String, i As Long, mc As Long, pnm As String, pdb As String
    listTxt = ""
    parts = Split(order, ",")
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then
            e = CLng(parts(i))
            mc = 0
            pnm = ""
            pdb = ""
            For r = 2 To last
                If IsNumeric(wsE.Cells(r, ES_ENC).Value) Then
                    If CLng(wsE.Cells(r, ES_ENC).Value) = e Then
                        mc = mc + 1
                        pnm = Trim(wsE.Cells(r, ES_PT).Value)
                        pdb = Trim(wsE.Cells(r, ES_DOB).Value)
                    End If
                End If
            Next r
            listTxt = listTxt & "   " & e & ")   " & pnm
            If pdb <> "" Then listTxt = listTxt & "   (DOB " & pdb & ")"
            listTxt = listTxt & "    - " & mc & " med(s)" & vbCrLf
        End If
    Next i
    Dim ans As String
    ans = Trim(InputBox("Encounters you can edit:" & vbCrLf & vbCrLf & listTxt & vbCrLf & _
        "Type the ENCOUNTER NUMBER to edit:", "Edit Encounter"))
    If ans = "" Then Exit Sub
    If Not IsNumeric(ans) Then
        MsgBox "Please type one of the encounter numbers shown.", vbExclamation, "Edit Encounter"
        Exit Sub
    End If
    Dim encNum As Long
    encNum = CLng(ans)
    If InStr(seen, "|" & encNum & "|") = 0 Then
        MsgBox "Encounter " & encNum & " was not found in the list.", vbExclamation, "Edit Encounter"
        Exit Sub
    End If
    If MsgBox("Reopen encounter " & encNum & " for editing?" & vbCrLf & vbCrLf & _
        "This REPLACES the medication list currently on screen with that" & vbCrLf & _
        "encounter's saved meds. Finish/print the current patient first if needed.", _
        vbYesNo + vbQuestion, "Edit Encounter") = vbNo Then Exit Sub
    Call LoadEncounter(encNum)
End Sub

' Restore a snapshot into the Input + Medications tabs and enter editing mode.
Private Sub LoadEncounter(ByVal encNum As Long)
    Dim wsE As Worksheet, wsM As Worksheet, wsI As Worksheet
    Set wsE = EncStore()
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsI = ThisWorkbook.Sheets(SH_INPUT)
    Application.ScreenUpdating = False
    Call ClearMedArea(wsM)
    Dim last As Long, r As Long, rr As Long, gotHdr As Boolean
    last = wsE.Cells(wsE.Rows.Count, ES_ENC).End(xlUp).Row
    gotHdr = False
    For r = 2 To last
        If IsNumeric(wsE.Cells(r, ES_ENC).Value) Then
            If CLng(wsE.Cells(r, ES_ENC).Value) = encNum Then
                If Not gotHdr Then
                    wsI.Range("C5").Value = wsE.Cells(r, ES_PT).Value
                    wsI.Range("C6").Value = wsE.Cells(r, ES_DOB).Value
                    wsI.Range("C7").Value = wsE.Cells(r, ES_RXDATE).Value
                    gotHdr = True
                End If
                Dim rec As MedRecord
                rec.MedName = CStr(wsE.Cells(r, ES_NAME).Value)
                rec.Strength = CStr(wsE.Cells(r, ES_STR).Value)
                rec.DosageForm = CStr(wsE.Cells(r, ES_FORM).Value)
                rec.SIG = CStr(wsE.Cells(r, ES_SIG).Value)
                rec.Quantity = CStr(wsE.Cells(r, ES_QTY).Value)
                rec.Refills = CStr(wsE.Cells(r, ES_REF).Value)
                rec.Expiration = CStr(wsE.Cells(r, ES_EXP).Value)
                rec.LotNumber = CStr(wsE.Cells(r, ES_LOT).Value)
                rec.Confidence = "Restored"
                rec.Warnings = ""
                rec.RawText = "[restored from encounter " & encNum & "]"
                rr = FirstEmptyRow(wsM)
                Call WriteMedRow(wsM, rr, rec, CStr(wsE.Cells(r, ES_PT).Value), _
                                 CStr(wsE.Cells(r, ES_DOB).Value), CStr(wsE.Cells(r, ES_DATE).Value), 0)
                wsM.Cells(rr, C_SRC).Value = wsE.Cells(r, ES_SRC).Value   ' restore Source (WriteMedRow blanks it)
                Call ApplyRowState(wsM, rr)
            End If
        End If
    Next r
    Call RenumberMeds
    Call ValidateMedications(False)
    gEditingEncounter = encNum
    Application.ScreenUpdating = True
    wsM.Activate
    MsgBox "Editing encounter " & encNum & "." & vbCrLf & vbCrLf & _
        "Add, remove, or fix medications as needed, then click" & vbCrLf & _
        """Save Edited Encounter"" to update the Log and Tebra note.", _
        vbInformation, "Editing Encounter " & encNum
End Sub

' Save the edited encounter: replace its Log + snapshot rows, refresh Tebra, offer reprint.
Public Sub SaveEditedEncounter()
    If gEditingEncounter <= 0 Then
        MsgBox "You are not editing a saved encounter." & vbCrLf & _
               "Click ""Edit Encounter"" to reopen one first.", vbInformation, "Save Edited Encounter"
        Exit Sub
    End If
    Dim encNum As Long
    encNum = gEditingEncounter
    Dim wsM As Worksheet, wsLg As Worksheet
    Set wsM = ThisWorkbook.Sheets(SH_MEDS)
    Set wsLg = ThisWorkbook.Sheets(SH_LOG)
    Dim lastMed As Long, r As Long, medCount As Long
    lastMed = wsM.Cells(wsM.Rows.Count, C_NAME).End(xlUp).Row
    medCount = 0
    For r = MEDS_HDR_ROWS + 1 To lastMed
        If Trim(wsM.Cells(r, C_NAME).Value) <> "" Then medCount = medCount + 1
    Next r
    If medCount = 0 Then
        MsgBox "There are no medications to save for this encounter.", vbExclamation, "Save Edited Encounter"
        Exit Sub
    End If
    If MsgBox("Save changes to encounter " & encNum & "?" & vbCrLf & vbCrLf & _
        medCount & " medication(s) will REPLACE encounter " & encNum & " in the" & vbCrLf & _
        "Log and Tebra note (no duplicates).", vbYesNo + vbQuestion, "Save Edited Encounter") = vbNo Then Exit Sub

    Dim vol As String
    vol = AskInitials()

    Application.ScreenUpdating = False
    Call DeleteRowsByEncounter(wsLg, LG_ENC, encNum, LOG_HDR_ROWS + 1)   ' drop old Log rows
    Call SaveEncounterSnapshot(encNum)                                    ' refresh snapshot
    For r = MEDS_HDR_ROWS + 1 To lastMed                                  ' re-log under same #
        If Trim(wsM.Cells(r, C_NAME).Value) <> "" Then
            Call LogPrint(r, Trim(vol & " (edited)"), encNum)
        End If
    Next r
    Application.ScreenUpdating = True

    On Error Resume Next
    Call FillTebraTemplate                                                ' rebuild note from Log
    On Error GoTo 0

    gEditingEncounter = 0
    If MsgBox("Encounter " & encNum & " updated in the Log and Tebra note." & vbCrLf & vbCrLf & _
        "Reprint the corrected labels now (2 copies each)?", _
        vbYesNo + vbQuestion, "Reprint?") = vbYes Then
        Call PrintEncounterLabelsNoLog
    Else
        Call ShowLogSheet
    End If
End Sub

' Print every named med that has Exp + Lot (2 copies each) WITHOUT logging - used by the
' "reprint after edit" step, since the edited rows were already re-logged in SaveEditedEncounter.
Private Sub PrintEncounterLabelsNoLog()
    Dim ws As Worksheet, wsL As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_MEDS)
    Set wsL = ThisWorkbook.Sheets(SH_LABEL)
    Dim brotherName As String
    brotherName = SelectBrotherPrinter()
    If brotherName = "" Then
        If MsgBox("Brother QL-1100c not found. Choose a printer manually?", _
            vbOKCancel + vbExclamation, "Reprint") = vbCancel Then Exit Sub
        On Error Resume Next
        Application.Dialogs(xlDialogPrint).Show
        On Error GoTo 0
    End If
    Call ApplyLabelContentWidth(wsL)
    Call ApplyLabelPageSetup(wsL)
    Call RefreshPrintLabelLogo(wsL)
    Dim lastRow As Long, r As Long, done As Long
    lastRow = ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
    done = 0
    Application.ScreenUpdating = False
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" Then
            If Trim(ws.Cells(r, C_EXP).Value) <> "" And Trim(ws.Cells(r, C_LOT).Value) <> "" Then
                Call UpdateLabelPreviewForMedRow(r, False)
                If PrintLabelSurfaceSafe(LABEL_COPIES) Then
                    Call MarkPrinted(r)
                    done = done + 1
                End If
            End If
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox done & " medication(s) reprinted, " & LABEL_COPIES & " copies each  (" & _
        (done * LABEL_COPIES) & " labels).", vbInformation, "Reprint Complete"
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