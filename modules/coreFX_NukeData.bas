Attribute VB_Name = "coreFX_NukeData"
Option Explicit
Public config As Worksheet, ecsp As Worksheet, acsp As Worksheet, f110a As Worksheet, f110b As Worksheet, _
    depIO As Worksheet
Public cspRng As Range, f110aRng As Range, f110aRng0 As Range, f110bRng As Range, f110bRng0 As Range
Sub Initialize()

Set config = ThisWorkbook.Worksheets("SENSEI.CONFIG") ' Unified Config Sheet
Set ecsp = ThisWorkbook.Worksheets("CSP.TR") ' Main Table
Set acsp = ThisWorkbook.Worksheets("CSP.ACH")
Set f110a = ThisWorkbook.Worksheets("DEBT.A")
Set f110b = ThisWorkbook.Worksheets("DEBT.B")
Set depIO = ThisWorkbook.Worksheets("DEP.IO") ' nuke dep io as well

Set cspRng = Union(ecsp.Range("C3:D102"), ecsp.Range("F3:H102"), ecsp.Range("J3:K102"))

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
aRow = acsp.Cells.Find("*", SearchOrder:=xlByRows, searchDirection:=xlPrevious).Row
acsp.Range("C3:L" & aRow).ClearContents
cspRng.ClearContents
f110aRng.ClearContents
f110aRng0.Value = 0
f110bRng.ClearContents
f110bRng0.Value = 0
depIO.Range("A2:L9999").ClearContents

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
    ' Form 110 Setting
    .Range("F5").Value = False ' Reset Distiller to variable locations (appoint per time)
    .Range("F6").Value = "" ' Reset Distiller Fixed Export Path
    .Range("F7").Value = True ' Reset Distiller Deletion warning to enable
End With

Application.ScreenUpdating = True
End Sub
