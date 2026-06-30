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

' DK-1202 die-cut label: 62 mm x 100 mm. Landscape print uses the 100 mm edge
' as page width (~283 pt); keep content narrower so nothing bleeds to the next label.
Private Const LABEL_WIDTH_PT As Double = 228

Private Const FONT_LABEL_BODY As String = "Arial"
Private Const FONT_LABEL_HDR As String = "Helvetica"
Private Const LOGO_ASPECT As Double = 1.422   ' scu_emblem.png width / height (256 / 180, 1024 / 720)
Private Const LOGO_EMBLEM_FILE As String = "scu_emblem.png"

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
    Call AddButtonToSheet(ws3, "btnUpd",    "Update Label Preview", "UpdateLabelPreviewFromSelection", 20, 1, 220, 24, RGB(21, 101, 192))
    Call AddButtonToSheet(ws3, "btnPrint",  "Print This Label",    "PrintLabel",         23, 1, 220, 24, RGB(46, 125, 50))
    On Error Resume Next
    ws3.Shapes("btnUpd").DrawingObject.PrintObject = False
    ws3.Shapes("btnPrint").DrawingObject.PrintObject = False
    On Error GoTo 0
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
    Call MatchHeaderFormat(ws2.Cells(2, C_PRTD), ws2.Cells(2, C_CNT))
    Call MatchHeaderFormat(ws2.Cells(2, C_PRTD), ws2.Cells(2, C_SEL))
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
    Call MatchHeaderFormat(wsLog.Cells(2, 12), wsLog.Cells(2, 13))
    Call MatchHeaderFormat(wsLog.Cells(2, 12), wsLog.Cells(2, 14))

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
Private Function LogoB64() As String
    Dim b As String
    b = "iVBORw0KGgoAAAANSUhEUgAABAAAAALQCAYAAAAO8wKWAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAA"
    b = b & "DsMAAA7DAcdvqGQAADkUSURBVHhe7d1p121ldS7q+0dspbLAWlEURYkkIrpFZVsRTdSoAStUNBqUqLFCrGMVBKylrtmrrsk5"
    b = b & "384/O61njRkXnbHq+c45x5jX1drdsCWw1hzPM770Pp4iAWDK3pHkaJJdSf6v7FiOJfloH3wAAABYpa8PBWovWmV5qfH9SB94"
    b = b & "AAAAWLX7kxwYKVxlOXk6yf/pgw4AAACrdnWSw7YC7Fg0AAAAANgYtwyFai9e5fxT5yzUeQsAAACwEe5JcmikgJVzT62qqDH9"
    b = b & "2z7YAAAAsC6vSrJnSC9k5dyyaAC8pQ82AAAArNM/uBVgqdmdZH+S1/eBBgAAgHW7fWgCPDVS0MrZpRoAe5O8pg8yAAAAbIL7"
    b = b & "XA24lGgAAAAAsNHekOSg8wDOO9UAqH++qA8wAAAAbIp/cjXgeUcDAAAAgEm4w3kA55VqADyR5JI+sAAAALBp7h22A/TiVk6f"
    b = b & "2kLxaB9QAAAA2ESXD4VsHWbXC1w5dTQAAAAAmJQPDFsBdo0UuXLyVNOkVlAAAADAZHzFoYBnnX1J/tQHEgAAADbdL5MccSjg"
    b = b & "GacaAH/sgwgAAACb7qIkDzoU8IxzIMmv+yACAADAFFwx7G13KODpowEAAADApL0nyVGHAp42tVLiZ33wAAAAYEo+51DA0+ZQ"
    b = b & "kjv6wAEAAMDUfE8T4JSpBsAP+6ABAADAFN3lZoCTpsblW33AAAAAYIouOeFmAE2AZ6bOSfh6HzAAAACYqlcPBwLWvfe9CN7m"
    b = b & "VAPgtj5YAAAAMGVvHfa87xkphLc1tQXg832gAAAAYOo+MHz13j1SDG9jNAAAAACYrU8NNwPUloBeEG9bqgHwyT5AAAAAMBe3"
    b = b & "uh7wv1MNgA/3wQEAAIA5+Z4mwH83AD7SBwYAAADm5udJjm3x9YDVAHhvHxQAAACYo98MBwNuYxOgGgDv6gPC5L00yV+SXNP/"
    b = b & "HwAAANvsuUn+lOTwFjYBqgFwXR8QJu8Fw/tcja2vJ7mg/wsAAADb6nlJHkhyaMuaANUA+Js+GEzeK5PsTbJn2OLykJUeAAAA"
    b = b & "f/WyJA8nObhFTYD6QqwBMD+vS7L/hKsu63/XXH8nyfP7vwwAALCNqgnw6BY1AWrFwxv6IDB5r20NgEr971oNUO+3gx8BAACG"
    b = b & "4qmK/wMzbwJUQVjLxC/rA8DkXT7SAFhk37AaoK7BvLD/hwAAANvmii1oAiwaAK/uD8/kXTW8u2MNgMXcVxOgtrw4BBIAANh6"
    b = b & "r595E2B3kieTvLg/OJP31mF7x8kaAIsszgb4t/4HAAAAbJs5NwHqhPgnkrywPzSTd80ZNgAq9e88neQvSd7c/yAAAIBtUk2A"
    b = b & "KpTm1gSoBsDjToWfpbecRQNgkXq/6/DLm/sfBgAAsE0WKwHmdDtA7f9/sD8os1D7+qsB0Of8dKltIbUa4NfDjRgAAABbqQ7L"
    b = b & "e2RGTQANgPl6f5IjI3N+pjk8vOOuCwQAALbWy4cmwKJA6oXTlFINgPv6AzIL59sAWLwfdUDg1/sfDgAAsC1emuTeocCachOg"
    b = b & "7oP/U384ZqG+3J9vA6BSZwgcS/L7JJf1vwQAAGAb1MF5f1xSkbWuaADM14eGwr3P+bmmzhOo8wFu6H8RAADANnhukruGZdK9"
    b = b & "YJpC6tT3O/tDMQv/PBzm1+f8fFJbAqqpcFv/ywAAALbFz3ag2FpFqgHwq/4wzMInduidrC0B9efeneQl/S8FAADYBt8bCqOz"
    b = b & "uXd93akGwM/7gzALH9+hBsAidQjm40ne1v9iAACAbXDrUHTVXuleMG1ial/37f0hmIVP7XADoA6/rDMkqhFwU//LAQAAtkEV"
    b = b & "XnUmwJ6RomnTUg2A7/cHYBa+tKKzKarZVY2GWgEDAACwdT6Y5GCS/Rt+TWB9vf33/uOZhS+uqAGwSDUBfudcAAAAYBtdO5wH"
    b = b & "UI2ATW0C1BWGtW2B+bllxQ2AescX5wJc3X8MAADA3F2e5MGhMNrEJkA1AL7afzSz8JVhfvuc72TqHa+DJSu1CgYAAGCrPD/J"
    b = b & "PcP96b1gWnfqN32+/2Bm4VtD46nP+SpS51/Uu1WrEAAAALbOHRt4TWD9ns/2H8osfHONDYBKvecOBwQAALZWncxeX0Y35YaA"
    b = b & "KtA+3X8ks7DuBsAi9b7fmeSi/gMBAADmrvZG1x7pTbghoA6J+8f+A5mF7w7XPPY5X0fqLIJ7k7yi/0gAAIC5e3OSxzbgC201"
    b = b & "AD7cfxyz8NPhBoo+5+tINbqqGfFEkjf1HwoAADB3lya5e83nAtSX2ff1H8YsbFIDoLK4IWBfknf0HwsAALANvj00AXaPFE07"
    b = b & "nVqBcF3/QczCjzesAVCpJkA1AGo1wA39BwMAAGyDfx6KtVWfC1ANgHf2H8Ms/GL44t7nfBNSh2DW9pN67wEAALbOW5I8OizL"
    b = b & "7wXTTqS2HdSX2Kv7D2EWfjs0lPq8b0pqxUvdEPD5/sMBAAC2wQuT/HpF5wLUn1+rDq7qP4JZ2PQGQKWaAPWu/0v/8QAAANvi"
    b = b & "K8MS6dov3YumZWXRALiy/+XMwj0TaABU6j38ryRf7w8AAACwLa4/4fq0XjQtI1V41R5xd7PP0+8n0gCo1LtYKwE0AQAAgK31"
    b = b & "0h28KnCxAkADYJ7+vMMrSHYi9Z7XrRgAAABb6193YEtA7b+uJsCL+1/GLPxlye/LqlLbAX7QHwYAAGCbvCvJE0u8JaCuYns8"
    b = b & "yfP6X8QsTLUBUKmVAN/vDwQAALBN6paAnw/Xp9UX/F44nU0WDYDn97+EWbhvwg2AioMBAQAAknxqOOCt9vD3wulMszfJg0me"
    b = b & "0/9wJu+iJI8MTZ4+71NKrQS4rT8cAADAtqnr+/5wHgcEVgPg/v6HMgtzaQAsrgj8Wn9AAACAbVQHBB4ervTrBdSpUg2AWibO"
    b = b & "/Fw4kwZApZoAteXlC/0hAQAAttG1SR4YCqUzXQ1QWwjqrnjm55Ikj86kAVBZrAS4qT8oAADAtqr90me6GkADYL5eOhwSeb4H"
    b = b & "RW5S6lmqwfUP/WEBAAC21duHvf2nOxugmgR39f+YWZhjA6BSKxqqwfXO/sAAAADbqk72/2qSQ6e4KaD+77/q/yGz8JKZNgAq"
    b = b & "dbVhrV75m/7QAAAA2+zNw00BtXS67wfXAJivlw3zPccGQKUaAE8meWV/cAAAgG33mROWTy+2BdTqgB/2f5FZeN0wv6faAjLl"
    b = b & "PDU8X211ubg/PAAAwLarr6U/SXJ0+IJaBdTt/V9iFubeAKhUE+CIcywAAABO7t1J7kvy/2kAzNY2NAAWqcMuvccAAACn8C9D"
    b = b & "mJ8rtqgBUKkmwBf6IAAAAMDcvXU476EXynNNNTrqoMv394EAAACAOdu2BkClDrk8kOTKPhgAAAAwV9vYAKhUA+CRJM/rAwIA"
    b = b & "AABzdO2WNgAWNwP8ug8IAAAAzNH7hkK4F8jbkjoU8Kt9UAAAAGBuPrDlDYDFoYA1DgAAADBb294AqOwdcnkfHAAAAJgLDYDj"
    b = b & "OZjkL31wAAAAYC4+keToSEG8balDAWsrwLf7AAEAAMAc3KQB8IzUoYAf7IMEAAAAU3ejBsAzUmcB7ElyWR8oAAAAmDINgGem"
    b = b & "tgIcSvLbPlAAAAAwZV8c9r73QnjbU1sBvtwHCwAAAKbqKxoAo9k93I7wd33AAAAAYIrqK7cGwHj2J3mgDxgAAABMkQbAybO4"
    b = b & "GvBbfdAAAABgar6Z5PBI8SvHs2s4JPF/94EDAACAKfm2BsBpU1sBHu4DBwAAAFNSy9s1AE6f2gpQqyUAAABgkjQAziy1FaBu"
    b = b & "BXhrH0AAAACYgh8lOTRS8MqzcyDJfX0AAQAAYAp+muTgSLEr43k6yS19EAEAAGDTaQCcXXYPKwFe0wcSAAAANpkGwNmnzkz4"
    b = b & "VR9IAAAA2GR3Dl+0e5Erp05tBXhvH0wAAADYVPcM99z3AldOnX1JHu6DCQAAAJtKA+Dc8tSwCuCLfUABAABgE92lAXDOqQMB"
    b = b & "9yZ5eR9UAAAA2DR/GZaz9+JWzixHkvywDyoAAABsmgeHr9i9sJUzy66hCXB1H1gAAADYJA9oAJx3Dg23KQAAAMDG0gBYTo4l"
    b = b & "ub4PLgAAAGyKR5LsGSlo5exyIMmf++ACAADAJrg4yaMaAEtLXQv4oT7IAAAAsG4XaQAsNXWbwv19kAEAAGDdLtQAWHpqFcDH"
    b = b & "+0ADAADAOl2a5EkNgKWmVgHU1YoAAACwMV6RZPeQXsjKuccqAAAAADbKy5Ps0gBYemoVQF2vCAAAABvhZRoAO5ZjbgQAAABg"
    b = b & "U1yeZK8GwI5kf5I/9gEHAACAdXhjkoPDKoBewMr5p1YBvLsPOgAAAKzalUkOaADsWKq5cmcfdAAAAFi1N2gA7HgOJ3lzH3gA"
    b = b & "AABYpauHAlUDYOdS43tHH3gAAABYpbcN+9R70SrLSx2wWNcC1pWLAAAAsBbXagCsJDXGX+qDDwAAAKuiAbCa1FWLD/fBBwAA"
    b = b & "gFV5lwbAyvJ0khv6BAAAAMAq/H2SoyPFqiw/dduCKwEBAABYi/oirQGwmtRNCweTvK5PAgAAAOy0D2gArDS13eJf+yQAAADA"
    b = b & "TvuIBsBKU4cBPtQnAQAAAHbaJzUAVp5aBXB9nwgAAADYSTdpAKw8h5L8uE8EAAAA7KQbNQBWnt3DgYCX9skAAACAnXJzkiMj"
    b = b & "RarsbJ5O8tE+GQAAALBT6kR6DYDVp64D/HWfDAAAANgpX9EAWEtqC0DdCPCyPiEAAACwE76sAbC21G0AH+sTAgAAADvhNg2A"
    b = b & "tcU2AAAAAFbme0kOjxSnsvOp2wBqG8DL+6QAAADAsn1XA2CtcRsAAAAAK/EdDYC1prYB/KxPCgAAACzbHUkOjRSmsprsSfJk"
    b = b & "kkv6xAAAAMAy1dfn+grdC1NZXY4meXefGAAAAFgmDYD1p25h+GafGAAAAFgmDYD1Z1+Se/vEAAAAwDL9QgNg7dk1zMEVfXIA"
    b = b & "AABgWe5Jsn+kKJXV5liSG/vkAAAAwLJoAGxG6iaGn/bJAQAAgGW5WwNgI1LXAT6a5Dl9ggAAAGAZfqcBsBGpcwAOJ7m6TxAA"
    b = b & "AAAsw/1J9o4UpLL61DkAn+4TBAAAAMugAbA5cQ4AAAAAO+Y+DYCNSc3Dg32CAAAAYBmq4NQA2IzUOQAHk1zRJwkAAADOx0VD"
    b = b & "4bl7pBiV9aTOAfhwnygAAAA4HxoAm5cjSb7ZJwoAAADOx4VD0akBsDmpKxnv6RMFAAAA5+OSYd+5BsDmZE+SJ5I8r08WAAAA"
    b = b & "nKuXDQcAagBsVg4neXOfLAAAADgfdyc5MFKEyvpyNMnH+kQBAADA+fjH4eT5XoTK+lIHAX6jTxQAAACcj4uTPDnsPe+FqKwn"
    b = b & "tSLjzj5RAAAAcL6+O+w774WorCd1LsMjfZIAAADgfL1t2HfeC1FZT+pQxmoCvLpPFAAAAJyvB4aisxejsp5UQ+YdfZIAAADg"
    b = b & "fN1mFcBGpebixj5JAAAAcL7+bjh9vheisp7UmQz/3icJAAAAluH+JPtGilFZfQ4m+UWfIAAAAFiG2gZgFcBmpBox9/YJAgAA"
    b = b & "gGW41jkAG5O6CaD+eUmfJAAAAFiGun/ebQDrz64k+5Nc0ScIAAAAluGO4QC6XpDKalMNgNqOcV2fIAAAAFiGDyY5NlKQyurz"
    b = b & "dJKP9gkCAACAZXhZkj0n7EGX9aXOY7ilTxAAAAAsy2+H/ee9IJXVprZifKdPDgAAACzLrW4D2IgcTPKrPjkAAACwLHXwXB1A"
    b = b & "1wtSWW1qFcbv++QAAADAsrxoOIXeOQDrTV3HeF+fHAAAAFim3zkHYO2pwxgfT/K8PjkAAACwLN+wDWDtqVUY+4abGQAAAGBH"
    b = b & "fGi4h74XpbK6VAPgUJLX98kBAACAZbkyyYGhCO2FqawmiwbANX1yAAAAYFkuSPLYsA+9F6ayuhxOcn2fHAAAAFimu4ZVAL0o"
    b = b & "ldXl6LAdAwAAAHbMd4Yv0L0oldXlWJKP94kBAACAZfrU8AW6F6WyutRBjDf3iQEAAIBlevfwBboXpbK61Ph/qU8MAAAALFPd"
    b = b & "BHDQTQBrzZEkt/WJAQAAgGV6aZKnkuweKUxlNdEAAAAAYCUecRXgWnMoye19UgAAAGDZ/pxk30hhKqtJNQDu6JMCAAAAy3Zn"
    b = b & "kgMjhamsJhoAAAAArMRPh4MAe2Eqq0k1AL7fJwUAAACWTQNgvamxrzkAAACAHfWdJIdHClNZTWr7xa/7pAAAAMCyaQCsN9UA"
    b = b & "+FWfFAAAAFi2n9gCsNZYAQAAAMBK/FwDYK2psf+PPikAAACwbH9Ism+kMJXVxC0AAAAArMTDSfaMFKaymlQD4I4+KQAAALBM"
    b = b & "L03yVJLdI4WprCYaAAAAAOy4q4cbAHaNFKaymmgAAAAAsOM+nOTYSFEqq8uRJF/vEwMAAADL9G9DAdqLUlldavxv6xMDAAAA"
    b = b & "y3TncA99L0pldTma5Mt9YgAAAGBZLkryhBsA1p6nk9zcJwcAAACW5RoHAG5E6gyGm/rkAAAAwLJ8zgGAG5HaAlCHMQIAAMCO"
    b = b & "+LX9/xuRWoXxnj45AAAAsAzPT/KU/f8bkWoAvL1PEAAAACzD9cPS816MympT5y8cSnJVnyAAAABYhm8PX557QSqrTTUADia5"
    b = b & "rE8QAAAALMNDSfaOFKSy2uwemgAv6hMEAAAA5+tay/83JnUGw8NJLuiTBAAAAOfrm0mOjBSjsvrsS/LnPkEAAACwDPXF2fL/"
    b = b & "zUhdw3h3nyAAAAA4X9dZ/r9RqRsAftonCQAAAM7X7U7/36jUXNSWDAAAAFiai5M8ORw81wtRWU+OJflcnygAAAA4Hx8dCs5e"
    b = b & "hMr68nSSD/WJAgAAgPNx13DoXC9CZT3ZNZwB8NY+UQAAAHCurhyKzSo6eyEq68nu4TaGV/XJAgAAgHNVB80dGSlCZX2psxge"
    b = b & "S/LcPlkAAABwLi5K8oTD/zYu+5Pc0ycLAAAAztXHh8PmegEq601dAfiDPlkAAABwrv40fG3uBaisN0eT3NInCwAAAM7FdUOh"
    b = b & "2YtPWX/qSsb39AkDAACAc/HL4fT/XnzKelO3MdSqjCv6hAEAAMDZcvXf5qYOZHzUDQAAAAAsw+3DQXO9+JT150CSO/uEAQAA"
    b = b & "wNl65bDE3Nf/zUydy/C1PmkAAABwtv7d4X8bnZqbG/qkAQAAwNl48bDHfPdI4SnrT63KqC0Ar+0TBwAAAGfjq77+b3T2Jrm/"
    b = b & "TxoAAACcjUuHL/++/m9u6maGH/aJAwAAgLNxm6//G59jST7RJw4AAADO1CuGk/99/d/c1P7/WgFwZZ88AAAAOFO1rPzwSNEp"
    b = b & "m5Pa//9QnzgAAAA4U9dZ+j+JVIPm+33yAAAA4EzVV+V9IwWnbFZq//+H+uQBAADAmfj6UFj2YlM2K7X/v5o0dVYDAAAAnJVr"
    b = b & "kxwZistecMpm5UCS3/YJBAAAgDPx4HDyfy82ZfNSZzTc0icQAAAATuc7w9L/p0aKTdmsLK7/e1OfRAAAADiV9w/Fv6X/00jt"
    b = b & "/b+3TyIAAACcykuHorLulO+Fpmxmavn/bX0iAQAA4FTuHu6T70WmbG7qoMZr+kQCAADAyfyrK/8mF8v/AQAAOCvXD0vJ7fuf"
    b = b & "VmrObu2TCQAAAGNeMRT+9v1PKzVntV3D6f8AAACckT8M18i58m9a2Z/k930yAQAAYMz37fufbJ5O8sk+oQAAANB9eigie2Ep"
    b = b & "m5/dQ17SJxUAAABO9K7h+rgqIntxKZuf2rLx4z6pAAAAcKLXJtnj0L9Jp7ZtvLNPLAAAACxckuShJAdGikqZRvYlub9PLAAA"
    b = b & "AJzormHpvxP/p5v6+v/ZPrEAAACw8AMn/k8+izMbXtgnFwAAAMpXnPg/i9Tqje/0yQUAAIBy4/Dlf9dIQSnTSc1fnf5/ZZ9g"
    b = b & "AAAAeF+So677m0Wq+P9ln2AAAAC4diga68q/XkzK9FKrOK7rkwwAAMB2e9Pw1b+ujOuFpEwvdW3j7/skAwAAsN0uG675q6LR"
    b = b & "dX/zSB3geEOfaAAAALbXS5M8pvifVfYn+UufaAAAALZX3Q1//7DvX/E/n9TX/w/2yQYAAGA7vSDJfUkOK/5nFV//AQAA+B/P"
    b = b & "HQ6IO6L4n13q5H9f/wEAAPjv4v+3iv9Zxtd/AAAA/ttzhuL/qOJ/lqmv/x/okw4AAMB2WXz5V/zPM3WLw+/6pAMAALBdLlb8"
    b = b & "zzq7hq//7+gTDwAAwPZYnPZvz/98U9c4/qJPPAAAANvj0iT3uupv1qmv/9UAeEOffAAAALbDK5M8NBSHiv/5ppb+f6tPPgAA"
    b = b & "ANvhdUkeVfzPPnuGf76wvwAAAADM31VDUVinwiv+552nk3yyvwAAAADMX50CvzfJfsX/7FMNnj/3FwAAAID5+8Cw5H/fSLEo"
    b = b & "80od/FdXOl7bXwIAAADm7cahIFzsCZd5p+b6jv4SAAAAMG9fHvaC7x4pFGV+qS0e9U8H/wEAAGyR7w3Ffy0J74WizDM13//U"
    b = b & "XwQAAADm6cIkvx7ugO8Fosw3dcbDb/rLAAAAwDy9Yjj9/chIgSjzTW3xqNsdXtNfCAAAAObn6iRPDF+CXfO3Xaml/zf3FwIA"
    b = b & "AID5uSHJweErsOJ/u1INn9/1FwIAAID5uWXY7++av+1LLf0/kOSK/lIAAAAwL0763+7U3H+6vxQAAADMR93zfs9QAPaiULYj"
    b = b & "Tv0HAACYuTcmeSTJYfv9tzZ7h+X/L+8vBwAAAPPw3iT7hn3fiv/tTa38+FB/OQAAAJiHuubt6PD1txeEsh2ppk8d+HhHfzkA"
    b = b & "AACYh8Vhf7XsuxeFsj2plR8PJnluf0EAAACYthckudthfzI0f+rgvzf3lwQAAIBpuzLJw8Nhf70YlO1LNYE+2V8SAAAApu09"
    b = b & "w15/h/1JzX+d/fDT/pIAAAAwbZ9NcsRhfzKkmkC1EuTC/qIAAAAwXd912J+ckD1JDia5qr8oAAAATNPzkvxmuOKtF4Gyvalm"
    b = b & "0D/1lwUAAIBpet1wtVst++8FoGxvqhn0rf6yAAAAME3vPmGZt8P+pFLvQd388Nv+sgAAADBNNw6FXh32p/iXRfYneTzJi/oL"
    b = b & "AwAAwPT8m8P+ZCSL1SBX9xcGAACAaXlukp8Pxf+ukQJQtjf1PtS+/3/oLw0AAADT8ookf05ydKT4E6mm0Jf7SwMAAMC01JLu"
    b = b & "J5IcGin8ZLtT5z/Ul/8f9ZcGAACAaflAkgPD4W4O+5MTszjx/3f9pQEAAGBaPjss+a+T/nvxJ9udKv7rwL+Hk7ygvzgAAABM"
    b = b & "x7876V9OkX3DwX+X9xcHAACAaXhOkl8MxX8v+kQqdd1fnQdxbX95AAAAmIaXJPmTk/7lFKkVIfV+1NkQAAAATNCVSR4dvuw6"
    b = b & "7E/GUkv+68T/G/vLAwAAwDS8Y1jWXaf9K/5lLFX817aQL/WXBwAAgGn4yHCVWx3qpviXk6WK/2/2lwcAAIBp+MJQ2NXX/17w"
    b = b & "iSxS78gP+ssDAADANHxrKOxqaXcv+EQWqT3/P+kvDwAAANPwH675k9OktoPUaf939pcHAACAzXdJknuGr7q94BNZpIr/I0l+"
    b = b & "018gAAAANt8rktw/FHa94BNZZFH8/66/QAAAAGy+K5M8nuSQk/7lFFkU/39O8rz+EgEAALDZ3j6c8n9A8S+nSL0bdR2k4h8A"
    b = b & "AGCC3pfkYJJ9in85RRbF/5+GcyIAAACYkI8Py7n3jhR8IifGsn8AAICJ+txw0n8t/e/Fnsgiiz3/v1X8AwAATM9XkzydZNdI"
    b = b & "wSdyYo4muau/QAAAAGy+byv+5QxTK0R+018gAAAANt+PhuK/F3oiPfWe/Ky/QAAAAGy+Xw1fdHuhJ9JT70k1iwAAAJiQC4Y9"
    b = b & "3LWXuxd6IiemtoXUl//aJgIAAMCEXJTk98MX3TrNvRd8IovsHor/W/tLBAAAwGa7NMm9SQ6PFHsiJ6aK/2oS3dxfIgAAADbb"
    b = b & "i5PcPxT/vvzLqbI3yZEkH+0vEQAAAJvtNUkeSXJI8S+nSL0b+5McSPKe/hIBAACw2S5L8rjiX06TejcODv/7b/tLBAAAwGar"
    b = b & "L/9V/Fdhp/iXk6Xejdoa8tDwzgAAADAhr0vymOJfziB12N/vkrywv0QAAABstjcN97fXXm7Fv5ws9Y7UNX8/7S8QAAAAm++q"
    b = b & "obCrw9wU/3Ky1DV/Vfzf1l8gAAAANt/iy7/iX06Vuuav9vzf1F8gAAAANp/iX06Xei9qW8ieJNf3FwgAAIDNd7k9/3IGqa/+"
    b = b & "Dye5sr9AAAAAbL7LhtP+Ff9yqtRJ/79J8vz+AgEAALD5FsW/q/7kZFkc9vfd/vIAAAAwDa9Q/Mtpsi/JkSSf7i8PAAAA0/CC"
    b = b & "JPcnOaT4l5HUO1GNoToX4p395QEAAGAanpfkvuFAN8W/jOVoknuTvLq/PAAAAEzDc5P8YVjWrfiXnsV+/58kuaC/PAAAAEzH"
    b = b & "PYp/GUm9D3uHd+OW/tIAAAAwLT8frnLrxZ9InQVR+/2v7y8NAAAA0/KDYWl3L/xEqin0p+FKSAAAACbsG4p/Gcliv//t/YUB"
    b = b & "AABgev5F8S8ttd9//3DN3439hQEAAGB6bhqK/9rb3YtA2d7UQX+PJLmmvzAAAABMzw3DXe61zLsXgLKdqUZQNYR+meQF/YUB"
    b = b & "AABgeq4dTnXfM1IEyvallvzvG778f6m/LAAAAEzTFUPhXwVfLwRlO3M4yZNJ3tlfFgAAAKbpxUkeS3Jg+OrbC0HZriyW/N+d"
    b = b & "5KX9ZQEAAGCanjPc5V5fexX/snc4A+LW/qIAAAAwbb8YCr5eCMr2pc5/qK//7+svCQAAANP2zWGpdy8EZbuyWPL/2ySv6i8J"
    b = b & "AAAA0/YZxb8MS/6PJflGf0EAAACYvvcMy/53jxSEsh2p8x4ODg2AD/UXBAAAgOl7Y5L9Q+HXi0LZnlQD6L4kr+8vCAAAANN3"
    b = b & "6QnX/fWCULYji/3+P05yQX9BAAAAmId7XPe31alVH0eSfLG/GAAAAMzHd4fD3hT/25ea81r1sS/J+/uLAQAAwHx8yon/W51a"
    b = b & "9fFIkjf1FwMAAID5ePuw7NuJ/9uZWvVxV5IX9BcDAACA+Xj5sPy7ln73wlDmnWr41KqP2voBAADAzP0hySH7/rcqNdd12F9d"
    b = b & "8/fZ/kIAAAAwP4tD/3qBKPPN4rC/PUne218IAAAA5ufjDv3bytRhf48mubK/EAAAAMzPVUkODl+Be4Eo802t9vhtkkv7CwEA"
    b = b & "AMD8XJjk4WEZeC8QZZ7ZNaz2+EF/GQAAAJivHw+Hvzn0bztSqzzqy/+X+osAAADAfH3Svv+tSTV49g83PPxDfxEAAACYrzcP"
    b = b & "xWDd/d6LRZlXqvivMx7qn2/tLwIAAADz9dwkD9n3vzU5kuS+JK/qLwIAAADzdod9/1uT2u//n0ku7i8BAAAA8/Zh+/63IouT"
    b = b & "/r/TXwAAAADmr5aA7x1Ogu8Fo8wni5P+v9BfAAAAALbD3cPBf71glHmktnTsG+b4Q33yAQAA2A6ft/R/1lmc9F9L/530DwAA"
    b = b & "sKX+xpV/s0+d9P9gksv75AMAALA9/jJ8He5Fo8wjtd//niTP7xMPAADA9rh1KBB70SjTz+Kk/x/1SQcAAGC7vGVYGl6FYi8e"
    b = b & "Zdqp7RxV/N/WJx0AAIDtc1+SAyPFo0w7dZVjNXY+3SccAACA7fNlp/7PLotr/qqp84E+4QAAAGyfK536P7uceM3fNX3CAQAA"
    b = b & "2E6/d+r/rFLF/+EkDyd5bZ9sAAAAttMnLf2fXY4m+WOSS/tkAwAAsJ1eOiz73zNSRMo0U82cXyZ5Tp9sAAAAttdPh9PhexEp"
    b = b & "00vt9a/i/3t9kgEAANhu70pybKSQlOmlVnFU8f+vfZIBAADg/iT7R4pJmVZq+0Y1cm7qEwwAAABfcPDfLLJvuL7x7/sEAwAA"
    b = b & "QB38V1+NHfw33dQ1f7V6oxoA7+gTDAAAAOUHDv6bdKr4P5jkySRv6pMLAAAA5W+H4r9OjO+FpWx+qvivJf8PJbmsTy4AAAAs"
    b = b & "3D18Pe6FpUwj1bz5c5JL+8QCAADAwvsd/Dfp1En/v0lyYZ9YAAAAOJFr/6abatz8rE8oAAAAdDf6+j/Z1Lx9r08oAAAAdBck"
    b = b & "eTTJ3pHiUjY3dVDjfyW5rU8oAAAAjLnZ1//JZfcwZ7f0yQQAAIAxlyR5wtf/SWVPkqNJPtUnEwAAAE7mX4bT43uRKZuZKv7r"
    b = b & "qr+P9okEAACAk3lBkqeGorIXmrJ5qVUah5O8r08kAAAAnMoX7P2fTPYNVzS+q08iAAAAnEqd/P+Yr/8bn1qhUYV/5a19EgEA"
    b = b & "AOB0PuPr/8aniv8Dw/9+c59AAAAAOBMPOfl/o7Mo/uufV/bJAwAAgDPxYSf/b3Sq6D+Y5JEkr+mTBwAAAGfqj8Oe8l54yvpT"
    b = b & "xf+h4XyGV/WJAwAAgDN1na//G5tF8V/bM17RJw4AAADOxs+HIrMXn7LeVPF/OMm9SS7tkwYAAABn43XD3vJdIwWorC+L4v++"
    b = b & "JC/skwYAAABn69+SHB0pQGV9ObH4f0GfMAAAADhbFyR5PMmekSJU1pNF8X+/L/8AAAAsyz8meXqkCJX15MQv//b8AwAAsDS/"
    b = b & "SXJgpBCV1ceXfwAAAHbEFcPJ/w7/W38c+AcAAMCO+YrD/zYiVfxXI6a+/Fv2DwAAwNI9kGTfSEEqq8uJy/4V/wAAACzd//b1"
    b = b & "f+058cu/Zf8AAADsiO8kOTJSlMpqUsX/weEKxpf1yQEAAIBluGAoPPeMFKay86niv25eeCzJq/vkAAAAwLJcb/n/2rIo/ut/"
    b = b & "v6ZPDAAAACzT7cPBc704lZ3P/uHaxav6pAAAAMCyPWr5/1qyd/j6/5Y+IQAAALBsb3f431pSxX+tunhHnxAAAADYCbfZ/7/y"
    b = b & "7B6aLjf0yQAAAICdcm+SfSNFquxMar//sSQf6xMBAAAAO+WK4e75Kkp7oSrLT43z00k+3ycCAAAAdtJNw9foXqjKzqSK/6/3"
    b = b & "SQAAAICd9h/DCoBeqMpy89TQaPlRnwAAAADYaRcmedz1fzueKv7rwL+7+wQAAADAKlzj+r8dTxX/h5Lcn+SiPgEAAACwCl90"
    b = b & "/d+OZ3+SJ5O8og8+AAAArMqdSQ6MFK2ynNTWijpf4e/6wAMAAMCqXJLkCfv/dyx13V8d+vf3feABAABglez/39nUdX+1xQIA"
    b = b & "AADW6jP2/+9IXPcHAADARvnxcDp9L2Dl/FJj+qc+2AAAALAuDyTZO1LAyrln3/DPl/fBBgAAgHV49VCs1kF1vYiVc8vu4UyF"
    b = b & "d/TBBgAAgHV5z7BPvRexcu6pQ/8+1wcaAAAA1qlOp3cA4HJSh/7Vl/+f9UEGAACAdfuPJAdHilk5++xP8kiSi/sgAwAAwLrd"
    b = b & "6wDApaT2/R9O8pY+wAAAALBuLxuKfwcAnn9q339tpwAAAICN87bhq7UGwPnlUJLf9sEFAACATfEJNwCcd/YMqyjqOkUAAADY"
    b = b & "SF8fTq3vRa2ceWrp/019YAEAAGCT/CrJgZGiVs4stfT/nj6oAAAAsGnuS7JvpLCV06dO/a/rE1/fBxUAAAA2yYuSPDUUsr24"
    b = b & "ldOnzk74ah9UAAAA2DRvGJaw98JWTp869O+xJBf0QQUAAIBN824HAJ5z6uC/ukEBAAAANl4VsEdHils5derMhHv7YAIAAMCm"
    b = b & "qv3rGgBnn/r6/+E+mAAAALCp7nAGwFmn9v4/1AcSAAAANtmvkxwYKXLl5KmT/7/QBxIAAAA22V+G/ey9yJXx1HWJNV4v7wMJ"
    b = b & "AAAAm+o5SR5Nsmek0JXx1HaJn/WBBAAAgE32kiRPaQCcVWr5//v7QAIAAMAme+2wnL2WtfdCV56dapQ8nuSCPpAAAACwyf42"
    b = b & "yeEku0aKXXl2aqzq1gQAAACYlOuGJe290JXx1Fj9fR9EAAAA2HQ3aACccWqbROXFfRABAABg030iydGRYleenQNJ7uwDCAAA"
    b = b & "AFPwaQ2AM06N05f6AAIAAMAUfCXJkZFiV56dGqd39AEEAACAKbhVA+CMstj//6I+gAAAADAF3x6utusFrzwz+5L8oQ8eAAAA"
    b = b & "TEXdaa8BcPrUGN3eBw8AAACm4vsaAGeUOgDwU33wAAAAYCqsADizHEtyXR88AAAAmIqfJzk4UvDKX7NrOAPg8j54AAAAMBUa"
    b = b & "AKfPniSPJbmgDx4AAABMhQbA6VNf///cBw4AAACmRAPg9Knx+Y8+cAAAADAlv9AAOG3qkMRv9YEDAACAKflhkkMjRa/8NXUD"
    b = b & "wJf7wAEAAMCU3O4awNPmaJKb+sABAADAlNyhAXDaHEnywT5wAAAAMCXf0wA4bWp83tUHDgAAAKbk1uELdy965a+pBsA1feAA"
    b = b & "AABgSjQATp1dwyGJb+oDBwAAAFPy2eGQu174yvEsGgCX94EDAACAKflnDYBTZtEAeG0fOAAAAJiSOt2+7rnvha8cjwYAAAAA"
    b = b & "s/BuDYBTRgMAAACAWXiLQwBPGQ0AAAAAZuH1SfYPhW4vfkUDAAAAgJl42VDk7h4pfsUtAAAAAMzERUkeT7JnpPiV4zmc5Ko+"
    b = b & "cAAAADA19yfZO1L4yvFUA+CaPmgAAAAwNXcN5wD0wleOpxoA1/VBAwAAgKn54bDPvRe+cjx1S8INfdAAAABgam5LcnSk8JXj"
    b = b & "qbH5WB80AAAAmJpPaQCcMseSfL4PGgAAAEzNe4Zl7r3wleOp5kitkgAAAIBJe9Nw0F0vfOV46nyE7/dBAwAAgKl5yVDo7h4p"
    b = b & "fiU5kOQ/+6ABAADAFD2QZO9I8SvJviT39QEDAACAKbpz+NLdi19J9iR5LMlFfdAAAABgar7lHICTprZGVBPgsj5oAAAAMDV1"
    b = b & "FeDTI8WvHE/dBPCWPmgAAAAwNdcNKwB2jRS/khxL8qE+aAAAADA1rxwOu9MAGM+RJF/qgwYAAABT9KCbAE6aQ0l+2AcMAAAA"
    b = b & "puiXSQ6OFL+S7E/y+z5gAAAAMEW3Dkvde/Erx28CeDLJxX3QAAAAYGpuGA6768WvHD8boVZHvKkPGgAAAEzNa4el7g4CHE9d"
    b = b & "k/jBPmgAAAAwRQ85CPCkqe0RtU0CAAAAJu9nw4n3vfiV5ECSO/uAAQAAwBR9LsnRkeJXkj1JHk1yYR80AAAAmJq3ugngpKmz"
    b = b & "EWp1hIMAAQAAmLznJXliuPauF8By/JaEf+6DBgAAAFNU+9xrv3svfiU5nOSOPmAAAAAwRbc4B+CkqRsSHugDBgAAAFP0NucA"
    b = b & "nDR1DsDBJG/sgwYAAABTc/FwDkCdet8LYDm+OuIf+6ABAADAFP18+NLdi185vjriI33AAAAAYIpuGk6878WvHG8AfLgPGAAA"
    b = b & "AEzRFcMKgNrz3gvgbY8GAAAAALPy5yT7RgrgbY8tAAAAAMzK19wGMJoak/f1wQIAAICpunY48d42gGemGgDv7IMFAAAAU/Zw"
    b = b & "kr0jRfA2pxoA7+oDBQAAAFP23SSHR4rgbY4GAAAAALPz7mEbQC+CtzkaAAAAAMzOc5I8lmTPSCG8rdEAAAAAYJa+5TaAZ0QD"
    b = b & "AAAAgFl6u20Az4gGAAAAALP1gNsA/icaAAAAAMzWV6wC+J9UA+Bv+wABAADAHFyR5FCSXSMF8balGgBv6gMEAAAAc3FXkgMj"
    b = b & "BfG2pRoAV/XBAQAAgLn4cJKnRwribYsGAAAAALN2UZInkuwZKYq3KRoAAAAAzN43HQaoAQAAAMD8XekwQA0AAAAAtsN/Jjk4"
    b = b & "UhhvSzQAAAAA2ArvSXJspDDelmgAAAAAsDXuS7JvpDiee2rrw/4kr+0DAgAAAHN005ZeCVgNgL1JLusDAgAAAHN04XAlYBXD"
    b = b & "vUiecxYNgFf3AQEAAIC5+uIWngWgAQAAAMDWeeFQEO8eKZTnGg0AAAAAttJXt2wVgAYAAAAAW+lFwwqAPSPF8hyjAQAAAMDW"
    b = b & "2qZVABoAAAAAbK3FKoBtOAvANYAAAABstX/dklUAtdWhrj+8tA8AAAAAbINLkjy1BWcBLBoAdQMCAAAAbKWbkzw9UjTPKRoA"
    b = b & "AAAAkOShJPtGCue5RAMAAAAAktww81UAGgAAAAAwuDvJwZHieQ7RAAAAAIDB1UmODFfm9QJ66tEAAAAAgBN8J8nR4WaAXkRP"
    b = b & "ORoAAAAAcILnD8X/3pEiesqpBsDjSV7QHxgAAAC21cdmeCBgNTQe7A8KAAAA264OBDw0UkhPNRoAAAAAMOL1QwNg90gxPcVo"
    b = b & "AAAAAMBJ3DKjrQAaAAAAAHAKf0lycKSgnlo0AAAAAOAU3jiTrQAaAAAAAHAaX5jBVgANAAAAADgDv534rQAaAAAAAHAGXj0U"
    b = b & "0XtGiusppH77A/2hAAAAgGf78IS3AuxL8qf+QAAAAMC47yc5luSpkSJ7k6MBAAAAAGfhOUnuS3JgpMje5GgAAAAAwFl6/dAA"
    b = b & "mNJ5ABoAAAAAcA4W5wHsGim2NzEaAAAAAHCOvjGhQwE1AAAAAOA83JXk8EjBvWnRAAAAAIDz8Pwkj0zgUMBqAPy5/3gAAADg"
    b = b & "zL0hyf4ke0cK701JNQB+3384AAAAcHbek+Rokt0jxfcmpFYo/LL/aAAAAODs3bjBNwNUA+BX/QcDAAAA5+arG3ozgAYAAAAA"
    b = b & "LNkdSY6NFOHrjAYAAAAA7IBfDE2Ap0aK8XVEAwAAAAB2yN3DwYCb0ATQAAAAAIAdcmGSPyY5sgFNAA0AAAAA2EEXbUgTQAMA"
    b = b & "AAAAdtjFG9AEOJjkZ/2HAQAAAMu17ibAoSS39x8FAAAALN9iO8A6DgbUAAAAAIAVqpUAvxmuCOxF+k5GAwAAAADW4JcrbgJo"
    b = b & "AAAAAMCa/CDJ00l2jRTsy44GAAAAAKzR14YmwO6Ron2Z0QAAAACANfvUcDDg3pHCfVnRAAAAAIAN8J4k+5Mc2KEbAqoB8P3+"
    b = b & "lwIAAACrd1WSR5McHingzzf1Z36r/4XsqLr28ZX9/wgAAADlRUnu3oHDAY8kubX/Zeyo1w+rOj7T/x8AAACwUF/rqwmwZ6SY"
    b = b & "P5dUA+Cr/S9hR706ycEk/0+SPyW5pv8LAAAAUD42FJDLOBdAA2D1qgFQBzvWSo6aw9qG8Y0kF/R/EQAAAP4myYPDLQG9qD+b"
    b = b & "aACs3okNgJqD+uexJA8leXf/lwEAAODCJD86zy0BGgCr1xsAi9S5ANXQ+Wb/DwAAAKB8Yigea1tAL/BPFw2A1TtZA6CyWA1w"
    b = b & "X5Kr+38IAAAAdbL8H4bVALtHCsuTRQNg9U7VAKjUuQ51NsChJDf3/xgAAABKXelXh8pVAdkLy7FUA+Br/Q9hR52uAbBINXKq"
    b = b & "ofOLJC/ofwgAAABcm+T+YSn56YrM+ne+2P8AdtSZNgAWqSbNo0ne0v8gAAAAeM7wZb9WA5zqbID6wvyZ/h+zo862AVBbAuqM"
    b = b & "h9oScGP/wwAAAKD8XZI/Dl/6x24KqAbAZ/t/xI462wbAIostAd/pfyAAAAAs1GFyVXTWigANgPU61wbAItXMuSfJC/sfDAAA"
    b = b & "AOWyJD8dCshaUq4BsB7n2wCoVCPn4SRX9D8cAAAAFq4f7pmv4v//1QBYuWU0AOpcgDrbobYFvK3/BQAAAHCiTw9NgC/3/wc7"
    b = b & "ahkNgEUTYN9wOOAN/S8BAACAE70qydX9/8iOWlYDYJH6s+qqwI/3vwgAAABYn2U3ACp1w0Ot5qhVHQAAAMAG2IkGQKXOAzia"
    b = b & "5Jb+FwIAAACrt1MNgEr9mf/lXAcAAABYv51sAFTqz63tAJ/pfzEAAACwOjvdAKjUn30sySf7Xw4AAACsxioaAJU6E6BWAny4"
    b = b & "/wAAAABg562qAVCp2wHqisB39R8BAAAA7KxVNgAq9XftT3JV/yEAAADAzll1A6BSDYAnkrys/xgAAABgZ6yjAfBUkkNJ/tR/"
    b = b & "DAAAALAz1tEAqFQT4GiSn/QfBAAAACzfuhoAi9TNAF/oPwoAAABYrnU3AOp6wFoJcH3/YQAAAMDyrLsBUNk3/NOhgAAAALBD"
    b = b & "NqEBsDgU8J7+4wAAAIDl2IQGwCJ1HsCX+w8EAAAAzt8mNQDqPIAjSd7afyQAAABwfjapAVDZn+ThJBf0HwoAAACcu01rANR5"
    b = b & "AMeSfLv/UAAAAODcbVoDoFK/pa4GfEf/sQAAAMC52cQGQKW2AjzUfywAAABwbja1AVCpWwG+2n8wAAAAcPY2uQFQv+lQktf3"
    b = b & "Hw0AAACcnU1uAFSqAXBn/9EAAADA2dn0BkCltgLc0H84AAAAcOam0ADY50BAAAAAOD9TaABUahXAzf3HAwAAAGdmKg2APUme"
    b = b & "SvK8/gAAAADA6U2lAVA55lpAAAAAODdTagDsHn7rS/pDAAAAAKc2pQZApVYB3NYfAgAAADi1qTUAahVAxSoAAAAAOAtTawBU"
    b = b & "ahXAl/uDAAAAACc3xQZA3QjwZJKL+8MAAAAA46bYAKg8neTT/WEAAACAcVNtAOxLcn9/GAAAAGDcVBsAlVoF8N7+QAAAAMCz"
    b = b & "TbkBcCjJr/sDAQAAAM825QZA/eaDSV7XHwoAAAB4pik3ACpHk9zaHwoAAAB4pqk3AOq3P9wfCgAAAHimqTcAKseSvLs/GAAA"
    b = b & "APBXc2gAHE7yw/5gAAAAwF/NoQGwO8lTSZ7XHw4AAAA4bg4NgEptA7ihPxwAAABw3FwaAIeS/Lg/HAAAAHDcXBoAtQ3gSdsA"
    b = b & "AAAAYNxcGgCV2gbw/v6AAAAAwLwaAHUbwO39AQEAAIB5NQDqOR7uDwgAAADMqwFQqVUAV/eHBAAAgG03twbA0SSf7w8JAAAA"
    b = b & "2+61SQ7OqAFwIMmd/SEBAABg212c5NEk+0eK6SmmrgN8Iskl/UEBAABg27112DtfxXMvqKeWWslwJMnb+kMCAAAAyZeSHBsp"
    b = b & "qKeYOgfgi/0BAQAAgON+P5wH0AvqqaWe4Vf94QAAAIDjrpzJgYB7kjyW5IL+gAAAAMBxt81gK0A1MKqR8cb+cAAAAMBxFw5f"
    b = b & "z/eOFNZTSjUxPtofDgAAAPirf0ry9EhRPaXUTQDf6A8GAAAAPNNfkuwfKaynkvrtd/WHAgAAAJ7pQxM/C6C2MDyS5Ln9wQAA"
    b = b & "AIBnui/JvpHiegqpgwAPJLm8PxQAAADwTJ+Y+CqAo0ne3R8KAAAAeKYLkjyeZM9IcT2FVAPg0/2hAAAAgGf72lBI9+J6Cqmb"
    b = b & "AL7VHwgAAAB4tiuSHBz21PcCe9NTZwD8qj8QAAAAMK6u06tiuhfYm566CeCB/jAAAADAuDoM8OmRAnvTU2cXPJnkhf2BAAAA"
    b = b & "gGd7+fA1ffdIkb3JqW0L+4dtDAAAAMAZuHOi2wDqGsO394cBAAAAxt080dsAqgHwD/1hAAAAgHFXJTk0UmBveqpp8Zn+MAAA"
    b = b & "AMDJPTycBdCL7E3OkSRf7w8CAAAAnNxPkxwcKbI3OYeT/LA/CAAAAHByUzwHoBoWv+gPAgAAAJzcO4Yl9b3I3uTUNYD39AcB"
    b = b & "AAAATu4VwxkAu0cK7U3NviT39gcBAAAATu3BiR0EWL+1Di/8X/1BAAAAgJO7K8mBkUJ7U1OrFZ5KcnF/EAAAAODk7khyaKTQ"
    b = b & "3tRUA6Dywv4gAAAAwMl9ZWIHAe4atgFc1h8EAAAAOLlPJnl6pNDe1CwaAK/qDwIAAACc3AeTHBsptDc11QCoqwBf3x8EAAAA"
    b = b & "OLn3TnALwMEkb+4PAgAAAJzc24ZDAKuw7sX2JmbRAPib/iAAAADAyb1FAwAAAADm7xoNAAAAAJg/KwAAAABgC7wzyeGRQntT"
    b = b & "owEAAAAA5+D9E70FQAMAAAAAzsJHkjw9UmhvalwDCAAAAOfgsxNsAOxL8tr+IAAAAMDJ3TbBLQB7k7yqPwgAAABwcj8ZltT3"
    b = b & "QntTowEAwHn7/wGefi6xPnp3wAAAAABJRU5ErkJggg=="
    LogoB64 = b
End Function

Private Function LogoFilePath() As String
    ' Prefer the high-quality emblem PNG beside the workbook; fall back to embedded copy.
    On Error GoTo Fail
    Dim folder As String, localPath As String
    folder = ThisWorkbook.Path
    If folder <> "" Then
        localPath = folder & Application.PathSeparator & LOGO_EMBLEM_FILE
        If Len(Dir(localPath)) > 0 Then
            LogoFilePath = localPath
            Exit Function
        End If
    End If

    Dim tmp As String
    tmp = Environ$("TEMP")
    If tmp = "" Then tmp = Environ$("TMP")
    If tmp = "" Then Exit Function
    tmp = tmp & "\scu_emblem_embedded.png"
    Dim dom As Object, node As Object, stream As Object
    On Error Resume Next
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    If dom Is Nothing Then Set dom = CreateObject("MSXML2.DOMDocument")
    On Error GoTo Fail
    If dom Is Nothing Then Exit Function
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = LogoB64()
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write node.nodeTypedValue
    stream.SaveToFile tmp, 2
    stream.Close
    LogoFilePath = tmp
    Exit Function
Fail:
    LogoFilePath = ""
End Function

Private Sub InsertLabelLogo(ws As Worksheet, ByVal shapeName As String, ByVal topPt As Single, ByVal rightPt As Single, ByVal heightPt As Single)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Dim logoPath As String
    logoPath = LogoFilePath()
    If logoPath = "" Then Exit Sub

    Dim lh As Single, lw As Single, leftPt As Single, pic As Shape
    lh = heightPt
    lw = lh * LOGO_ASPECT
    leftPt = rightPt - lw - 2

    On Error Resume Next
    Set pic = ws.Shapes.AddPicture2(logoPath, msoFalse, msoTrue, leftPt, topPt, lw, lh, msoPictureCompressNone)
    If pic Is Nothing Then
        Set pic = ws.Shapes.AddPicture(logoPath, msoFalse, msoTrue, leftPt, topPt, lw, lh)
    End If
    If Not pic Is Nothing Then
        pic.name = shapeName
        pic.Placement = xlMove
        pic.LockAspectRatio = msoTrue
    End If
    On Error GoTo 0
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
    ws.Rows(2).RowHeight = 15
    ws.Rows(3).RowHeight = 10
    ws.Rows(4).RowHeight = 5
    ws.Rows(5).RowHeight = 17
    ws.Rows(6).RowHeight = 15
    ws.Rows(7).RowHeight = 20
    ws.Rows(8).RowHeight = 12
    ws.Rows(9).RowHeight = 11
    ws.Rows(10).RowHeight = 13
    ws.Rows(11).RowHeight = 12
    ws.Rows(12).RowHeight = 12
    ws.Rows(13).RowHeight = 14
    ws.Rows(14).RowHeight = 3
    ws.Rows(15).RowHeight = 2
    ws.Rows(16).RowHeight = 10
    ws.Rows(17).RowHeight = 14
    ws.Rows(18).RowHeight = 14

    ws.Range("A2:H2").Merge
    ws.Range("A3:H3").Merge
    ws.Range("A5:E6").Merge
    ws.Range("F5:H5").Merge
    ws.Range("F6:H6").Merge
    ws.Range("A7:H7").Merge
    ws.Range("A8:H8").Merge
    ws.Range("A9:H9").Merge
    ws.Range("A10:H12").Merge
    ws.Range("A13:D13").Merge
    ws.Range("E13:H13").Merge
    ws.Range("A17:H17").Merge
    ws.Range("A18:H18").Merge

    ws.Cells(2, 1).Value = "SATURDAY CLINIC FOR THE UNINSURED"
    ws.Cells(3, 1).Value = "1121 E. North Ave, Milwaukee WI    " & Chr(183) & "    (414) 588-2865"
    ws.Cells(9, 1).Value = "DIRECTIONS"

    ws.Cells(5, 1).Formula = "=IF('Patient & Input'!C5<>"""",'Patient & Input'!C5,""[Patient Name]"")"
    ws.Cells(5, 6).Formula = "=IF('Patient & Input'!C7<>"""",""Rx  ""&TEXT('Patient & Input'!C7,""m/d/yyyy""),""Rx  --"")"
    ws.Cells(6, 6).Formula = "=IF('Patient & Input'!C6<>"""",""DOB  ""&'Patient & Input'!C6,""DOB  --"")"

    If Trim(ws.Cells(7, 1).Value) = "" Then _
        ws.Cells(7, 1).Value = "[Select a medication row, then Update]"
    Call SetMiniValue(ws.Cells(13, 1), "EXP", "--", 12, "L")
    Call SetMiniValue(ws.Cells(13, 5), "LOT", "--", 12, "R")

    ws.Cells(17, 1).Value = "Auto-fills from the Medications tab. Print via 'Print Checked Labels'."
    ws.Cells(18, 1).Value = "Brother QL-1100c  " & Chr(183) & "  DK-1202 62 x 100 mm  " & Chr(183) & "  Landscape"

    Call FmtLbl(ws.Cells(2, 1), 9, True, "L", "C", LabelHeaderFont())
    Call FmtLbl(ws.Cells(3, 1), 6.5, False, "L", "C", LabelHeaderFont())
    Call FmtLbl(ws.Cells(5, 1), 17, True, "L", "C")
    Call FmtLbl(ws.Cells(5, 6), 7, False, "R", "B")
    Call FmtLbl(ws.Cells(6, 6), 11, True, "R", "T")
    Call FmtLbl(ws.Cells(7, 1), 14, True, "L", "C")
    Call FmtLbl(ws.Cells(8, 1), 8, False, "L", "C")
    Call FmtLbl(ws.Cells(9, 1), 6.5, True, "L", "C")
    Call FmtLbl(ws.Cells(10, 1), 8, True, "L", "T")
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
    With ws.Range("A13:H13").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    On Error Resume Next
    ws.Shapes("scuLogo").Delete
    On Error GoTo 0
    Call InsertLabelLogo(ws, "scuLogo", ws.Rows(2).Top + 1, ws.Range("A1:H15").Left + ws.Range("A1:H15").Width, 28)

    With ws.PageSetup
        .PrintArea = ws.Range("A1:H15").Address
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

Private Sub ApplyLabelContentWidth(ws As Worksheet)
    Dim currentW As Double, sf As Double, col As Long
    currentW = ws.Range("A1:H15").Width
    If currentW <= 0 Then Exit Sub
    sf = LABEL_WIDTH_PT / currentW
    For col = 1 To 8
        ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth * sf
    Next col
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

Private Sub FmtLbl(rng As Range, sz As Single, bld As Boolean, h As String, v As String, Optional fontName As String = FONT_LABEL_BODY)
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

Private Sub SetMiniValue(c As Range, miniText As String, valueText As String, vSize As Single, hAlign As String)
    ' Footer field: small label + larger bold value in one cell (e.g. "EXP 05/2027")
    Dim s As String
    s = miniText & "  " & valueText
    c.Value = s
    c.WrapText = False
    c.Font.Name = "Arial"
    c.Font.Color = RGB(0, 0, 0)
    On Error Resume Next
    c.Characters(1, Len(miniText)).Font.Size = 7
    c.Characters(1, Len(miniText)).Font.Bold = False
    c.Characters(Len(miniText) + 1, Len(s) - Len(miniText)).Font.Size = vSize
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
    For Each shp In ws.Shapes
        If Left(shp.name, 3) = "al_" Then shp.Delete
    Next shp
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone
    ws.Rows.RowHeight = 14

    ws.Columns("A:F").ColumnWidth = 9
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

    Dim n As Integer, base As Long, r As Long
    n = 0
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(wsM.Cells(r, C_NAME).Value) = "" Then GoTo NextR
        n = n + 1
        base = 3 + (n - 1) * 12

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

        ws.Range(ws.Cells(base, 1), ws.Cells(base, 6)).Merge
        ws.Cells(base, 1).Value = "SATURDAY CLINIC FOR THE UNINSURED"
        Call FmtLbl(ws.Cells(base, 1), 8.5, True, "L", "C", LabelHeaderFont())
        ws.Range(ws.Cells(base + 1, 1), ws.Cells(base + 1, 6)).Merge
        ws.Cells(base + 1, 1).Value = "1121 E. North Ave, Milwaukee WI   " & Chr(183) & "   (414) 588-2865"
        Call FmtLbl(ws.Cells(base + 1, 1), 6, False, "L", "C", LabelHeaderFont())

        ws.Range(ws.Cells(base + 2, 1), ws.Cells(base + 2, 4)).Merge
        ws.Cells(base + 2, 1).Value = IIf(patName <> "", patName, "[Patient Name]")
        Call FmtLbl(ws.Cells(base + 2, 1), 11, True, "L", "C")
        ws.Cells(base + 2, 1).Font.Size = PatientNameFontSize(IIf(patName <> "", patName, "[Patient Name]"), medLine)
        ws.Range(ws.Cells(base + 2, 5), ws.Cells(base + 2, 6)).Merge
        ws.Cells(base + 2, 5).Value = "Rx " & IIf(dateRx <> "", dateRx, "--") & "   DOB " & IIf(dob <> "", dob, "--")
        Call FmtLbl(ws.Cells(base + 2, 5), 7, True, "R", "C")

        ws.Range(ws.Cells(base + 3, 1), ws.Cells(base + 3, 6)).Merge
        ws.Cells(base + 3, 1).Value = medLine
        Call FmtLbl(ws.Cells(base + 3, 1), 13, True, "L", "C")
        ws.Cells(base + 3, 1).Font.Size = MedFontSize(medLine)

        ws.Range(ws.Cells(base + 4, 1), ws.Cells(base + 4, 6)).Merge
        ws.Cells(base + 4, 1).Value = fq
        Call FmtLbl(ws.Cells(base + 4, 1), 8, False, "L", "C")

        ws.Range(ws.Cells(base + 5, 1), ws.Cells(base + 5, 6)).Merge
        ws.Cells(base + 5, 1).Value = "DIRECTIONS"
        Call FmtLbl(ws.Cells(base + 5, 1), 6.5, True, "L", "C")

        ws.Range(ws.Cells(base + 6, 1), ws.Cells(base + 7, 6)).Merge
        Dim sigText As String
        sigText = IIf(sig <> "", sig, "[Instructions not found - enter manually]")
        ws.Cells(base + 6, 1).Value = sigText
        Call FmtLbl(ws.Cells(base + 6, 1), 8, True, "L", "T")
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

        ws.Rows(base).RowHeight = 14
        ws.Rows(base + 1).RowHeight = 11
        ws.Rows(base + 2).RowHeight = 18
        ws.Rows(base + 3).RowHeight = 20
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
            Call InsertLabelLogo(ws, "al_logo_" & r, ws.Rows(base).Top + 1, ws.Cells(base, 7).Left, 24)
        End If

        If warn <> "" And UCase(warn) <> "OK" Then
            ws.Cells(base + 9, 1).Value = "! " & warn
            Call FmtLbl(ws.Cells(base + 9, 1), 8, False, "L", "C")
            ws.Cells(base + 9, 1).Font.Color = RGB(191, 54, 12)
        End If

        Call AddRowButton(ws, "al_print_" & r, "Print this label", "RowPrint", base, RGB(123, 31, 162))
        Call AddRowButton(ws, "al_edit_" & r, "Edit this med", "RowEdit", base + 3, RGB(21, 101, 192))
        Call AddRowButton(ws, "al_remove_" & r, "Remove this med", "RowRemove", base + 6, RGB(191, 54, 12))
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

    Call ApplyLabelContentWidth(wsL)

    With wsL.PageSetup
        .PrintArea = wsL.Range("A1:H15").Address
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
    Application.ScreenUpdating = False
    For r = MEDS_HDR_ROWS + 1 To lastRow
        If Trim(ws.Cells(r, C_NAME).Value) <> "" And IsRowSelected(ws, r) Then
            If Trim(ws.Cells(r, C_EXP).Value) = "" Or Trim(ws.Cells(r, C_LOT).Value) = "" Then
                skipped = skipped + 1
            Else
                Call UpdateLabelPreviewForMedRow(r)
                If PrintLabelSurfaceSafe(1) Then
                    Call MarkPrinted(r)
                    Call LogPrint(r, batchVol)
                    Call ApplyRowState(ws, r)
                    done = done + 1
                End If
            End If
        End If
    Next r
    Application.ScreenUpdating = True

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
Public Sub UpdateLabelPreviewForMedRow(ByVal medRow As Long)
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
    Dim fq As String
    fq = formTxt
    If qty <> "" Then
        If fq <> "" Then
            fq = fq & "   " & Chr(183) & "   Qty " & qty
        Else
            fq = "Qty " & qty
        End If
    End If

    Dim pn As String: pn = IIf(patName <> "", patName, "[Patient Name]")
    wsL.Cells(5, 1).Value = pn
    wsL.Cells(5, 1).Font.Size = PatientNameFontSize(pn, medLine)
    wsL.Cells(5, 6).Value = "Rx  " & IIf(dateRx <> "", dateRx, "--")
    wsL.Cells(6, 6).Value = "DOB  " & IIf(dob <> "", dob, "--")

    wsL.Cells(7, 1).Value = medLine
    wsL.Cells(7, 1).Font.Bold = True
    wsL.Cells(7, 1).Font.Size = MedFontSize(medLine)

    wsL.Cells(8, 1).Value = fq

    Dim sigText As String
    sigText = IIf(sig <> "", sig, "[Directions not found - enter manually]")
    wsL.Cells(10, 1).Value = sigText
    wsL.Cells(10, 1).Font.Color = RGB(255, 255, 255)
    wsL.Cells(10, 1).Font.Bold = True
    Dim sz As Single
    If Len(sigText) <= 80 Then
        sz = 9
    ElseIf Len(sigText) <= 130 Then
        sz = 8
    Else
        sz = 7
    End If
    wsL.Cells(10, 1).Font.Size = sz

    Call SetMiniValue(wsL.Cells(13, 1), "EXP", IIf(expDate <> "", expDate, "--"), 12, "L")
    Call SetMiniValue(wsL.Cells(13, 5), "LOT", IIf(lotNum <> "", lotNum, "--"), 12, "R")
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
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.PrintOut Copies:=copies, Collate:=True
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

    ' Set up page for Brother QL-1100c  DK-1202  62mm x 100mm die-cut label
    ' LANDSCAPE: 100mm wide x 62mm tall  =  approx. 3.9" x 2.4"
    ' The label content block (A1:H15) is proportioned to print at ~100% scale.
    With wsL.PageSetup
        .PrintArea       = wsL.Range("A1:H15").Address
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

    If Not PrintLabelSurfaceSafe(1) Then
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