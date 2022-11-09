VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} utilityDepScantron 
   Caption         =   "Sensei Deployment Scantron"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "utilityDepScantron.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "utilityDepScantron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public config As Worksheet, pCard As Worksheet
Public lastR As Long, currR As Long
Public m114 As Workbook, sensei As Workbook
Sub updateDisp() ' read only information
    If pCard.Range("A" & currR).Value <> "" Then ' SSAN
        dsCssan.Caption = Format(pCard.Range("A" & currR).Value, "000000000")
        SetClipboard (Format(pCard.Range("A" & currR).Value, "000000000"))
    Else
        dsCssan.Caption = ""
    End If
    If pCard.Range("B" & currR).Value <> "" Then ' NAME
        dsCname.Caption = " " & pCard.Range("B" & currR).Value
    Else
        dsCname.Caption = ""
    End If
    If pCard.Range("C" & currR).Value = "X" Then ' FL
        dsCeFL.BackColor = &H8000000F
    Else
        dsCeFL.BackColor = &H885985
    End If
    If pCard.Range("D" & currR).Value = "X" Then ' 14
        dsCe14.BackColor = &H8000000F
    Else
        dsCe14.BackColor = &H885985
    End If
    If pCard.Range("E" & currR).Value = "X" Then ' 23
        dsCe23.BackColor = &H8000000F
    Else
        dsCe23.BackColor = &H885985
    End If
    If pCard.Range("F" & currR).Value = "X" Then ' 65
        dsCe65.BackColor = &H8000000F
    Else
        dsCe65.BackColor = &H885985
    End If
    dsClvD.Value = pCard.Range("G" & currR).Value
    dsCarrD.Value = pCard.Range("H" & currR).Value
    dsC65D.Caption = Format(pCard.Range("I" & currR).Value, "YYYY-MM-DD")
    dsCflD.Caption = Format(pCard.Range("J" & currR).Value, "YYYY-MM-DD")
    If pCard.Range("K" & currR).Value = "O" Then
        dsCdReady.Caption = "ALL SET"
    Else
        dsCdReady.Caption = "NOT YET"
    End If
    If pCard.Range("L" & currR).Value Like "*OMIT*" Then
        dsCdOmit.Caption = ChrW(&HD8)
    Else
        dsCdOmit.Caption = "O"
    End If
End Sub
Private Sub dsCaOmit_Click() ' TOGGLE OMIT
    If pCard.Range("L" & currR).Value Like "*OMIT*" Then
        pCard.Range("L" & currR).Value = ""
    Else
        pCard.Range("L" & currR).Value = "OMIT. on " & Format(Now(), "YYMMDD-HH:MM:SS")
    End If
    updateDisp
End Sub

Private Sub dsCrow_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If pCard.Visible = xlSheetHidden Then ' DOUBLE CLICK TO DISPLAY FORM
        pCard.Visible = xlSheetVisible
        pCard.Activate
    Else
        pCard.Visible = xlSheetHidden
        sensei.Sheets("CSP.TR").Activate
    End If
End Sub

Private Sub dsRun_Click()
    MsgBox "Due to policy config, this is not quite plausible at the moment..." & vbNewLine & vbNewLine & "Please goto 114 and use DEP.IO there instead"
    Exit Sub
    readScantron ' UNABLE DUE TO POLICY
End Sub

Private Sub dsCaNG_Click()
    pCard.Range("K" & currR).Value = ""
    updateDisp
End Sub
Private Sub dsCaOK_Click() ' CANNOT MARK EMPTY ENTRY
    pCard.Range("K" & currR).Value = "O"
    If Not dsCforce Then
        If pCard.Range("G" & currR).Value = "" Or pCard.Range("H" & currR).Value = "" Then pCard.Range("K" & currR).Value = ""
    End If
    updateDisp
End Sub

Private Sub dsCarrD_Change()
    pCard.Range("H" & currR).Value = dsCarrD.Value
End Sub
Private Sub dsClvD_Change()
    pCard.Range("G" & currR).Value = dsClvD.Value
End Sub
Private Sub dsCrowAdjust_SpinDown()
    lastR = pCard.Cells.Find("*", SearchOrder:=xlByRows, searchDirection:=xlPrevious).Row
    currR = currR - 1
    If currR < 2 Then currR = lastR
    If lastR < 2 Then currR = 2
    dsCrow.Caption = currR
    updateDisp
End Sub
Private Sub dsCrowAdjust_SpinUp() ' remember to trigger the update
    lastR = pCard.Cells.Find("*", SearchOrder:=xlByRows, searchDirection:=xlPrevious).Row
    currR = currR + 1
    If currR > lastR Then currR = 2
    dsCrow.Caption = currR
    updateDisp
End Sub

Private Sub hidePanel_Click()
sensei.Sheets("CSP.TR").Activate
Me.Hide
trackerAPI.Show
End Sub

Private Sub loadScantron_Click() ' generate the scantron and write last row
    generateScantron
    lastR = pCard.Cells.Find("*", SearchOrder:=xlByRows, searchDirection:=xlPrevious).Row
    updateDisp
    MsgBox "Data Loaded", vbOKOnly, "Sensei Scantron"
End Sub

Private Sub Reset_Click()
Dim qa As String: qa = MsgBox("Wipe Scantron?", vbYesNo, "Wipe Scantron")
    If qa = vbYes Then pCard.Range("A2:l9999").ClearContents
    updateDisp
End Sub

Private Sub UserForm_Initialize()
Set config = ThisWorkbook.Worksheets("SENSEI.CONFIG")
Set pCard = ThisWorkbook.Worksheets("DEP.IO")
Set sensei = ThisWorkbook

' VERSION IFORMATION
dsVer.Caption = config.Range("J2").Value
dsVerType.Caption = config.Range("J3").Value

' INITIALIZE ROW NUMBER and panels
currR = 2
dsCrow.Caption = currR
dsCssan.Caption = ""

updateDisp
End Sub

Sub readScantron() ' direct 114 to read scantron
Dim rq As String: rq = MsgBox("Ensure your Master 114 is operating!", vbOKCancel, "Sensei Scantron Operation Warning")
    If rq = vbCancel Then Exit Sub
Dim aTrans As String, aSSN As String, aNmn As String, aDate As String
Dim c As Long, lc As Long: lc = 0 ' for iteration purposes
Set m114 = Workbooks(config.Range("B8").Value) ' use link data
Dim n114 As String: n114 = config.Range("B8").Value
Dim wSSN As Range, wNmn As Range, wTrans As Range, wDate As Range
    Set wSSN = m114.Sheets("Master DD114").Range("C3")
    Set wNmn = m114.Sheets("Master DD114").Range("C2")
    Set wTrans = m114.Sheets("Master DD114").Range("C6")
    Set wDate = m114.Sheets("Master DD114").Range("C9")
Application.ScreenUpdating = False
    lastR = pCard.Cells.Find("*", SearchOrder:=xlByRows, searchDirection:=xlPrevious).Row
For c = 2 To lastR
    If pCard.Range("K" & c).Value = "O" And pCard.Range("L" & c).Value = "" Then
        If lc = 99 Then
            Application.Run n114 & "!Print114"
            Application.Run n114 & "!ExecuteExport"
            MsgBox "Cycle capacity reached, exiting..."
            Exit Sub
        End If
        With pCard
            aSSN = Format(.Range("A" & c).Value, "000000000")
            aNmn = .Range("B" & c).Value
        End With
        ' DO IN ORDER FL 14 23 65
        If pCard.Range("C" & c).Value = "X" Then
            aTrans = "FL02"
            aDate = pCard.Range("J" & c).Value
            wTrans = aTrans
            wDate = aDate
            wSSN = aSSN
            wNmn = aNmn
            Application.Run n114 & "!Addto114"   ' boot macro
            lc = lc + 1 ' count trans
        End If
        If pCard.Range("D" & c).Value = "X" Then
            aTrans = "1402"
            aDate = pCard.Range("G" & c).Value
            wTrans = aTrans
            wDate = aDate
            wSSN = aSSN
            wNmn = aNmn
            Application.Run n114 & "!Addto114" ' boot macro
            lc = lc + 1 ' count trans
        End If
        If pCard.Range("E" & c).Value = "X" Then
            aTrans = "2302"
            aDate = pCard.Range("G" & c).Value
            wTrans = aTrans
            wDate = aDate
            wSSN = aSSN
            wNmn = aNmn
            Application.Run n114 & "!Addto114" ' boot macro
            lc = lc + 1 ' count trans
        End If
        If pCard.Range("F" & c).Value = "X" Then
            aTrans = "6502"
            aDate = pCard.Range("I" & c).Value
            wTrans = aTrans
            wDate = aDate
            wSSN = aSSN
            wNmn = aNmn
            Application.Run n114 & "!Addto114" ' boot macro
            lc = lc + 1 ' count trans
        End If
        pCard.Range("L" & c).Value = "READ"
    End If
Next c
Application.ScreenUpdating = True
End Sub
