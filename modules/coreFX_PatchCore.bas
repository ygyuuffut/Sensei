Attribute VB_Name = "coreFX_PatchCore"
Option Explicit
' #################################
' Worksheet and Range Declarations
' #################################
' ==============Workbook=================
Public Sensei As Workbook

' =============Worksheets================
Public tCSP As Worksheet, aCSP As Worksheet
Public f110a As Worksheet, f110b As Worksheet
Public sConfig As Worksheet, sData As Worksheet

' ===============Ranges==================
Public cspIQ As Range
' Common Patches Placed Here

Sub RepairRef() ' Repairing Reference Library
' Workbook
    Set Sensei = ThisWorkbook
'Worksheets
    Set tCSP = Sensei.Worksheets("CSP.TR")
    Set aCSP = Sensei.Worksheets("CSP.ACH")
    Set f110a = Sensei.Worksheets("DEBT.A")
    Set f110b = Sensei.Worksheets("DEBT.B")
    Set sConfig = Sensei.Worksheets("SENSEI.CONFIG")
    Set sData = Sensei.Worksheets("SENSEI.DATA")
'Range
    Set cspIQ = Range("entryTable[ID]")
End Sub
Sub RepairAPI() ' Force VB reloading
    Unload trackerAPI
    trackerAPI.Show
End Sub

Sub RepairFreeFloaters() ' complimentary for deleteEntry
' Remove Freefloating Informations not associated with IQID
RepairRef
Application.ScreenUpdating = False
Dim pCell As Range, pRow As Long, pEnd As Long
With tCSP
    For Each pCell In cspIQ
        If pCell.Value = "" And pRow = 0 Then ' Mark starting row
            pRow = pCell.Row
        End If
        If pCell.Value <> "" And pRow > 2 Then ' Mark end row upon defined content
            pEnd = pCell.Row - 1
            .Range("D" & pRow & ":D" & pEnd).ClearContents
            .Range("F" & pRow & ":H" & pEnd).ClearContents
            .Range("J" & pRow & ":K" & pEnd).ClearContents
            pRow = 0
        ElseIf pCell.Row = 102 And pRow > 2 Then ' or till the end
            pEnd = pCell.Row
            .Range("D" & pRow & ":D" & pEnd).ClearContents
            .Range("F" & pRow & ":H" & pEnd).ClearContents
            .Range("J" & pRow & ":K" & pEnd).ClearContents
            pRow = 0
        End If
    Next pCell
End With
Application.ScreenUpdating = True
End Sub
