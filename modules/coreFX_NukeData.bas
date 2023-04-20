Attribute VB_Name = "coreFX_NukeData"
Option Explicit
Public config As Worksheet, ecsp As Worksheet, acsp As Worksheet, f110a As Worksheet, f110b As Worksheet, _
    depIO As Worksheet, f2424 As Worksheet, f117 As Worksheet
Public cspRng As Range, f110aRng As Range, f110aRng0 As Range, f110bRng As Range
Public f110bRng0 As Range, f2424Rng As Range, f2424expl As Object, f117Rng As Range
Public rejTemp As Worksheet, dataCard As Worksheet
Const cEnd = 502 ' last row
Const cCap = 500 ' capacity

Sub Initialize()

Set config = ThisWorkbook.Worksheets("SENSEI.CONFIG") ' Unified Config Sheet
Set ecsp = ThisWorkbook.Worksheets("CSP.TR") ' Main Table
Set acsp = ThisWorkbook.Worksheets("CSP.ACH")
Set f110a = ThisWorkbook.Worksheets("DEBT.A")
Set f110b = ThisWorkbook.Worksheets("DEBT.B")
Set depIO = ThisWorkbook.Worksheets("DEP.IO") ' nuke dep io as well
Set f2424 = ThisWorkbook.Worksheets("ADV.PAY")
Set f117 = ThisWorkbook.Worksheets("CAGE.PAY")
Set rejTemp = ThisWorkbook.Worksheets("DATA.TMP")
Set dataCard = ThisWorkbook.Worksheets("SENSEI.DATA")



Set cspRng = Union(ecsp.Range("C3:D" & cEnd), ecsp.Range("F3:H" & cEnd), ecsp.Range("J3:N" & cEnd))
Set f117Rng = Union(f117.Range("U2"), f117.Range("C40"), f117.Range("J10"), f117.Range("G10"), _
                    f117.Range("G11"), f117.Range("G12"), f117.Range("G13"), f117.Range("V13"), _
                    f117.Range("V27"), f117.Range("U9"), f117.Range("B19"), f117.Range("K19"), _
                    f117.Range("I5"), f117.Range("V23"), f117.Range("B56"), f117.Range("J56"), _
                    f117.Range("G14"))
Set f2424Rng = Union(f2424.Range("B9"), f2424.Range("F9"), f2424.Range("C10"), f2424.Range("C11"), _
                    f2424.Range("C12"), f2424.Range("G10"), f2424.Range("G12"), f2424.Range("I10"), _
                    f2424.Range("B14"), f2424.Range("G14"), f2424.Range("B16"), f2424.Range("J16"), _
                    f2424.Range("J28"), f2424.Range("J30"), f2424.Range("J32"))
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
acsp.Range("C3:O" & aRow).ClearContents
cspRng.ClearContents
f110aRng.ClearContents
f110aRng0.Value = 0
f110bRng.ClearContents
f110bRng0.Value = 0
depIO.Range("A2:L9999").ClearContents
f2424Rng.Value = vbNullString
f117Rng.Value = vbNullString
f2424expl.Text = ""
rejTemp.Range("B3:ZZ9999").Delete xlShiftUp ' NUKE REJ
f117.Range("V27") = 0 ' 117 TOTAL AMOUNT

With config
    ' The Document Link Data Wipe
    .Range("B2:B3").Value = "" ' R3R Wipe
    .Range("B6:B9").Value = "" ' 114 Wipe and unified name wipe
    ' The Rejection Report
    .Range("B35").Value = 0 ' UPDATE NUMBER TO DEFAULT
    .Range("B36").Value = False ' DISABLE AUTO STORAGE USING FOR HTML AP
    .Range("B37").Value = "" ' STORAGE PATH TO NAN
    .Range("B38").Value = False ' DISABLE ASSOCIATED PRINTING OF HTML
    .Range("B39").Value = "" ' WIPE FAKE PEOPLE
    .Range("B40").Value = False ' DISABLE SENT ON BEHALF OF
    .Range("B41").Value = False ' DISABLE AUTO SOURCE
    .Range("B42").Value = "" ' WIPE AUTO SOURCE TARGET
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
    .Range("D33").Value = False ' DISABLE AUTO SCROLL
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
    ' Form 117 Setting
    .Range("F95:F98").Value = False
    .Range("F99:F100").Value = True
    .Range("F101").Value = False
End With


nukeDataCard
Application.ScreenUpdating = True
End Sub
Public Sub nukeDataCard() ' cleans the data sheet
Initialize
Application.ScreenUpdating = False

With dataCard
    ' NUKE EMAIL LIST FOR REJC
    .Range("N2:N21").Value = "##"
    .Range("O2:Q21").Value = False
    .Range("R2:R21").Value = ""
    ' RESERVED POSITION FOR NUKE SSID LIST
    ' put here
End With

Application.ScreenUpdating = True
End Sub

Public Sub nukeDataBuffer() ' nuking temporary data card

Initialize
rejTemp.Range("A1:ZZ9999").Delete xlShiftUp ' NUKE REJ

End Sub
