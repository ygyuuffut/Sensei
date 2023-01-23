Attribute VB_Name = "coreFX_NukeData"
Option Explicit
Public config As Worksheet, ecsp As Worksheet, acsp As Worksheet, f110a As Worksheet, f110b As Worksheet, _
    depIO As Worksheet, f2424 As Worksheet
Public cspRng As Range, f110aRng As Range, f110aRng0 As Range, f110bRng As Range
Public f110bRng0 As Range, f2424Rng As Range, f2424expl As Object
Public rejTemp As Worksheet, rejCard As Worksheet
Const cEnd = 302 ' last row
Const cCap = 300 ' capacity

Sub Initialize()

Set config = ThisWorkbook.Worksheets("SENSEI.CONFIG") ' Unified Config Sheet
Set ecsp = ThisWorkbook.Worksheets("CSP.TR") ' Main Table
Set acsp = ThisWorkbook.Worksheets("CSP.ACH")
Set f110a = ThisWorkbook.Worksheets("DEBT.A")
Set f110b = ThisWorkbook.Worksheets("DEBT.B")
Set depIO = ThisWorkbook.Worksheets("DEP.IO") ' nuke dep io as well
Set f2424 = ThisWorkbook.Worksheets("ADV.PAY")
Set rejTemp = ThisWorkbook.Worksheets("DATA.TMP")
Set rejCard = ThisWorkbook.Worksheets("REJECT.RPT")


Set cspRng = Union(ecsp.Range("C3:D" & cEnd), ecsp.Range("F3:H" & cEnd), ecsp.Range("J3:M" & cEnd))
Set f2424Rng = Union(f2424.Range("B9"), f2424.Range("F9"), f2424.Range("C10"), f2424.Range("C11"), _
                    f2424.Range("C12"), f2424.Range("G10"), f2424.Range("G12"), f2424.Range("I10"), _
                    f2424.Range("B14"), f2424.Range("G14"), f2424.Range("B16"), f2424.Range("J16"))
Set f2424expl = f2424.Shapes("f2424_expl").TextFrame.Characters

Set f110aRng = Union(f110a.Range("A5:A17"), f110a.Range("C5:E17"), _
    f110a.Range("J5:J17"), f110a.Range("H5:H17"), f110a.Range("E23"), _
    f110a.Range("H2"), f110a.Range("M2"), f110a.Range("N5:N25"))
Set f110aRng0 = Union(f110a.Range("H5:H17"), f110a.Range("K5:K17"), f110a.Range("J20:J23"), _
    f110a.Range("L20"))

Set f110bRng = Union(f110b.Range("A5:A26"), f110b.Range("C5:E26"), _
    f110b.Range("J5:J26"), f110b.Range("N5:N26"))
Set f110bRng0 = Union(f110b.Range("H5:H26"), f110b.Range("K5:K26"))


End Sub
Public Sub nukeData()
Application.ScreenUpdating = False
Dim aRow As Long

Initialize
aRow = acsp.Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
acsp.Range("C3:N" & aRow).ClearContents
cspRng.ClearContents
f110aRng.ClearContents
f110aRng0.Value = 0
f110bRng.ClearContents
f110bRng0.Value = 0
depIO.Range("A2:L9999").ClearContents
f2424Rng.Value = vbNullString
f2424expl.Text = ""
rejTemp.Range("A1:Z9999").ClearContents
rejCard.Range("B2:S2").ClearContents
rejCard.Range("B3:T9999").Delete xlShiftUp ' NUKE REJ

With config
    ' The Document Link Data Wipe
    .Range("B2:B3").Value = "" ' R3R Wipe
    .Range("B6:B9").Value = "" ' 114 Wipe and unified name wipe
    ' Form Agreement Reset
    .Range("D2").Value = 0 ' User agreement
    .Range("D3").Value = 0 ' Debug Risk Agreement
    ' Form Common Setting
    .Range("D9:D11").Value = 2 ' D9 (1-ZH, 2-EN), all others (2 - off)
    .Range("D13").Value = False ' Pull 2 Excel Cards for Dual Update of entires
    .Range("D14").Value = 0 ' Reset Dual Update warning to enabled
    .Range("D23").Value = False ' TURN OFF AUTOSAVE
    .Range("D24").Value = 0 ' RESET ACTION COUNTER
    .Range("D25").Value = 25 ' RESET ACTION CAP
    .Range("D26").Value = "D" ' SET TYPE TO CSP
    .Range("D27").Value = Format(Now(), "YYYY") ' SET TO THIS YEAR
    .Range("D28").Value = 1 ' RESET MISC COUNTING
    .Range("D29").Value = "A" ' RESET SEARCH TO ALL
    .Range("D30").Value = "" ' RESET RECORD START
    .Range("D31").Value = "" ' RESET RECORD END
    .Range("D32").Value = True ' DEFALUT ENABLE FINAL LOG
    ' Form Distiller General Setting
    .Range("F5").Value = False ' Reset Distiller to variable locations (appoint per time)
    .Range("F6").Value = "" ' Reset Distiller Fixed Export Path
    .Range("F7").Value = True ' Reset Distiller Deletion warning to enable
    .Range("F33:F39").Value = False ' 110 - config group
    .Range("F64:F67").Value = False ' 2424 - CONFIG GROUP
    .Range("F68").Value = "" ' 2424
    .Range("F69").Value = False ' 2424 SSN LINK
    ' Deployment Scantron Settings
    .Range("J4").Value = False ' Resume to unified standard
    .Range("J5:J9").Value = 180 ' SET ALL DATES BACK TO 180
    .Range("J10").Value = False
    .Range("J11").Value = ""
End With

Application.ScreenUpdating = True
End Sub
