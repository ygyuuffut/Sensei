VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} utilityDataScantron 
   Caption         =   "Sensei Deployment Scantron"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8685.001
   OleObjectBlob   =   "utilityDataScantron.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "utilityDataScantron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public config As Worksheet, pCard As Worksheet
Public lastR As Long, currR As Long
Public m114 As Workbook, sensei As Workbook
Public prntIt As String
Public na As String
Public Tholder, Uholder, Vholder, Wholder, Xholder ' to temporary held value for setting reversion

Sub updateDsDisp() ' read only information' Dep Scantron
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
    ' UPDATE DISP
    lastR = pCard.Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
    DsCountLbl.Caption = Format(lastR - 1, "000") & " ENTRIES TOTAL"
End Sub
Private Sub dsCaOmit_Click() ' TOGGLE OMIT' Dep Scantron
    If pCard.Range("L" & currR).Value Like "*OMIT*" Then
        pCard.Range("L" & currR).Value = ""
    Else
        pCard.Range("L" & currR).Value = "OMIT. on " & Format(Now(), "YYMMDD-HH:MM:SS")
    End If
    updateDsDisp
    globalSave
End Sub

Private Sub dsCrow_DblClick(ByVal Cancel As MSForms.ReturnBoolean) ' Dep Scantron
    If pCard.Visible = xlSheetHidden Then ' DOUBLE CLICK TO DISPLAY FORM
        pCard.Visible = xlSheetVisible
        pCard.Activate
    Else
        pCard.Visible = xlSheetHidden
        sensei.Sheets("CSP.TR").Activate
    End If
End Sub

Private Sub dsDate14_Change() ' 14 REDLINE' Dep Scantron

On Error GoTo handleThis
config.Range("J7") = CInt(dsDate14.Value)
Exit Sub

handleThis:
dsDateUniCt.Value = Uholder

End Sub

Private Sub dsDate14_Enter() ' Dep Scantron
    Uholder = CInt(dsDate14.Value)
End Sub
Private Sub dsDate14adj_SpinDown() ' Dep Scantron
If dsDate14.Value <> "" Then
    dsDate14.Value = dsDate14.Value - 1
Else
    dsDate14.Value = 999
End If
End Sub

Private Sub dsDate14adj_SpinUp() ' Dep Scantron
If dsDate14.Value <> "" Then
    dsDate14.Value = dsDate14.Value + 1
Else
    dsDate14.Value = 0
End If
End Sub

Private Sub dsDate23_Change() ' 23 REDLINE' Dep Scantron

On Error GoTo handleThis
config.Range("J8") = CInt(dsDate23.Value)
Exit Sub

handleThis:
dsDateUniCt.Value = Vholder

End Sub

Private Sub dsDate23_Enter() ' Dep Scantron
    Vholder = CInt(dsDate23.Value)
End Sub

Private Sub dsDate23adj_SpinDown() ' Dep Scantron
If dsDate23.Value <> "" Then
    dsDate23.Value = dsDate23.Value - 1
Else
    dsDate23.Value = 999
End If
End Sub

Private Sub dsDate23adj_SpinUp() ' Dep Scantron
If dsDate23.Value <> "" Then
    dsDate23.Value = dsDate23.Value + 1
Else
    dsDate23.Value = 0
End If
End Sub

Private Sub dsDate65_Change() ' 65 REDLINE' Dep Scantron

On Error GoTo handleThis
config.Range("J9") = CInt(dsDate65.Value)
Exit Sub

handleThis:
dsDateUniCt.Value = Wholder

End Sub

Private Sub dsDate65_Enter() ' Dep Scantron
    Wholder = CInt(dsDate65.Value)
End Sub


Private Sub dsDate65adj_SpinDown() ' Dep Scantron
If dsDate65.Value <> "" Then
    dsDate65.Value = dsDate65.Value - 1
Else
    dsDate65.Value = 999
End If
End Sub

Private Sub dsDate65adj_SpinUp() ' Dep Scantron
If dsDate65.Value <> "" Then
    dsDate65.Value = dsDate65.Value + 1
Else
    dsDate65.Value = 0
End If
End Sub

Private Sub dsDateFL_Change() ' FL REDLINE' Dep Scantron

On Error GoTo handleThis
config.Range("J6") = CInt(dsDateFL.Value)
Exit Sub

handleThis:
dsDateUniCt.Value = Xholder

End Sub

Private Sub dsDateFL_Enter() ' Dep Scantron
    Xholder = CInt(dsDateFL.Value)
End Sub
Private Sub dsDateFLadj_SpinDown() ' Dep Scantron
If dsDateFL.Value <> "" Then
    dsDateFL.Value = dsDateFL.Value - 1
Else
    dsDateFL.Value = 999
End If
End Sub

Private Sub dsDateFLadj_SpinUp() ' Dep Scantron
If dsDateFL.Value <> "" Then
    dsDateFL.Value = dsDateFL.Value + 1
Else
    dsDateFL.Value = 0
End If
End Sub

Private Sub dsDateUni_Click() ' DO WE WISH TO UNIFY THE CONTROL OR SPLIT CONTROL EACH FILTER' Dep Scantron
If dsDateUni Then ' if split enable else disable
    dsDateUniCt.Enabled = False
    dsDateUniAdj.Enabled = False
    
    dsDateUni.Caption = "ISOLATE"
    config.Range("J4") = True
    dsDateFL.Enabled = True
    dsDateFLadj.Enabled = True
    dsDate14.Enabled = True
    dsDate14adj.Enabled = True
    dsDate23.Enabled = True
    dsDate23adj.Enabled = True
    dsDate65.Enabled = True
    dsDate65adj.Enabled = True
    
    dsDatelblFL.Enabled = True
    dsDatelbl14.Enabled = True
    dsDatelbl23.Enabled = True
    dsDatelbl65.Enabled = True
    Tholder = vbNullString ' delete backup
    Uholder = CInt(dsDateFL.Value)
    Vholder = CInt(dsDate14.Value)
    Wholder = CInt(dsDate23.Value)
    Xholder = CInt(dsDate65.Value)
Else
    dsDateUniCt.Enabled = True
    dsDateUniAdj.Enabled = True
    
    dsDateUni.Caption = "UNIFIED"
    config.Range("J4") = False
    dsDateFL.Enabled = False
    dsDateFLadj.Enabled = False
    dsDate14.Enabled = False
    dsDate14adj.Enabled = False
    dsDate23.Enabled = False
    dsDate23adj.Enabled = False
    dsDate65.Enabled = False
    dsDate65adj.Enabled = False
    
    dsDatelblFL.Enabled = False
    dsDatelbl14.Enabled = False
    dsDatelbl23.Enabled = False
    dsDatelbl65.Enabled = False
    Tholder = CInt(dsDateUniCt.Value) ' back up original
    Uholder = vbNullString
    Vholder = vbNullString
    Wholder = vbNullString
    Xholder = vbNullString
End If
globalSave
End Sub

Private Sub dsDateUniAdj_SpinDown() ' Dep Scantron
If dsDateUniCt.Value <> "" Then
    dsDateUniCt.Value = dsDateUniCt.Value - 1
Else
    dsDateUniCt.Value = 999
End If
End Sub

Private Sub dsDateUniAdj_SpinUp() ' Dep Scantron
If dsDateUniCt.Value <> "" Then
    dsDateUniCt.Value = dsDateUniCt.Value + 1
Else
    dsDateUniCt.Value = 0
End If
End Sub

Private Sub dsDateUniCt_Change() ' something went bad there, it need to block alpha' Dep Scantron
' see if TypeName() works ^

On Error GoTo handleThis
    ' MsgBox TypeName(CInt(dsDateUniCt.Value)) ' Debug to see if the thing worked
config.Range("J5") = CInt(dsDateUniCt.Value)
Exit Sub

handleThis:
dsDateUniCt.Value = Tholder

End Sub

Private Sub dsDateUniCt_Enter() ' BACK VALUE UP WHEN ENTER
    Tholder = CInt(dsDateUniCt.Value) ' back up original
End Sub

Private Sub dsPrintOptn_Click() ' print path config
If dsPrintOptn Then ' ALLOW FIXED PATH
    config.Range("J10").Value = True
    dsPrintOptn.Caption = "ASSIGN"
    dsPrintPath.Enabled = True
    dsPrintPathAssign.Enabled = True
    dsPrintPathRemove.Enabled = True
Else ' OR NOT
    config.Range("J10").Value = False
    dsPrintOptn.Caption = "MANUAL"
    dsPrintPath.Enabled = False
    dsPrintPathAssign.Enabled = False
    dsPrintPathRemove.Enabled = False
End If
globalSave
End Sub

Private Sub dsPrintPathAssign_Click() ' assign a path
Dim tempPath As String, tempFinder As FileDialog

Set tempFinder = Application.FileDialog(msoFileDialogFolderPicker)
    With tempFinder
        .Title = "Future Scantrons will be exported to here"
        .ButtonName = "Assign"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo datawrite
        tempPath = .SelectedItems(1)
    End With
datawrite:
Set tempFinder = Nothing
config.Range("J11").Value = tempPath

initializeConfig
globalSave
End Sub

Private Sub dsPrintPathRemove_Click() ' remove a path
Dim resB As String
    resB = MsgBox("Remove Current Path?", vbYesNo, "Scantron Path Deletion")
If resB = vbNo Then Exit Sub
config.Range("J11").Value = ""
initializeConfig
globalSave
End Sub

Private Sub dsRun_Click()
    MsgBox "Due to policy config, this is not quite plausible at the moment..." & vbNewLine & vbNewLine & "Please goto 114 and use DEP.IO there instead"
    Exit Sub
    readScantron ' UNABLE DUE TO POLICY
    globalSave
End Sub

Private Sub dsCaNG_Click()
    pCard.Range("K" & currR).Value = ""
    updateDsDisp
    globalSave
End Sub
Private Sub dsCaOK_Click() ' CANNOT MARK EMPTY ENTRY
    pCard.Range("K" & currR).Value = "O"
    If Not dsCforce Then
        If pCard.Range("G" & currR).Value = "" Or pCard.Range("H" & currR).Value = "" Then pCard.Range("K" & currR).Value = ""
    End If
    updateDsDisp
    globalSave
End Sub

Private Sub dsCarrD_Change() ' Dep Scantron
    pCard.Range("H" & currR).Value = dsCarrD.Value
End Sub
Private Sub dsClvD_Change() ' Dep Scantron
    pCard.Range("G" & currR).Value = dsClvD.Value
End Sub
Private Sub dsCrowAdjust_SpinDown() ' Dep Scantron
    lastR = pCard.Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
    currR = currR - 1
    If currR < 2 Then currR = lastR
    If lastR < 2 Then currR = 2
    dsCrow.Caption = currR
    updateDsDisp
End Sub
Private Sub dsCrowAdjust_SpinUp() ' remember to trigger the update' Dep Scantron
    lastR = pCard.Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
    currR = currR + 1
    If currR > lastR Then currR = 2
    dsCrow.Caption = currR
    updateDsDisp
End Sub


Private Sub hidePanel_Click()
sensei.Sheets("CSP.TR").Activate
pCard.Visible = xlSheetHidden
Me.Hide
    'saveThis
trackerAPI.Show
globalSave
End Sub

Private Sub loadScantron_Click() ' generate the scantron and write last row' Dep Scantron
    generateScantron
    lastR = pCard.Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
    updateDsDisp
    MsgBox "Data Loaded", vbOKOnly, "Sensei Scantron"
    globalSave
End Sub

Private Sub printDsScantron_Click() ' Print for Deployment Scantron
    prntIt = "DEPLOY"
    directPrint
    globalSave
End Sub

Sub directPrint() ' General Printing Prompt

Dim saveToPrompt
Dim Cpath As String: Cpath = na ' current path contains no name
Dim Cexist As String, saveTo As String

    saveTo = config.Range("J11").Value
On Error GoTo handleIt
If Not dsPrintOptn Or saveTo = "" Then ' ALWAYS PROMPT IF DISABLED PATHWAY OR SAVETO IS BLANK
    Set saveToPrompt = Application.FileDialog(msoFileDialogFolderPicker)
    With saveToPrompt
        .Title = "Sending Scantron Export to here..."
        .ButtonName = "Save"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = 0 Then Exit Sub ' if cancelled exit immediately
        saveTo = .SelectedItems(1)
    End With
Else
    If prntIt = "DEPLOY" Then ' DS SCANTRON
        Cexist = Dir(saveTo & "\Deployment Scantron Exports\", vbDirectory) ' check parent
        If Cexist = "" Then
            Cexist = saveTo & "\Deployment Scantron Exports\"
            MkDir Cexist ' A FIXED DIRECTORY
        End If
        Cexist = Dir(saveTo & "\Deployment Scantron Exports\" & Format(Now(), "YYYY-MM"), vbDirectory) ' check child
        If Cexist = "" Then
            Cexist = saveTo & "\Deployment Scantron Exports\" & Format(Now(), "YYYY-MM")
            MkDir Cexist ' A FIXED DIRECTORY
        End If
        Cpath = saveTo & "\Deployment Scantron Exports\" & Format(Now(), "YYYY-MM") ' assign to Cpath
    Else 'if
    End If
End If

exportForm:
Set saveToPrompt = Nothing ' UNLOAD OBJECT >>Does not save any way ?
If Not dsPrintOptn Or saveTo = "" Then Cpath = saveTo  ' only wrote when fixed path is not activated

' add a Mkdir, or make directory for when under constant method
If prntIt = "DEPLOY" Then
    pCard.ExportAsFixedFormat xlTypePDF, _
        Filename:=Cpath & "\deployScantron." & Format(Now(), "YYMMDD-HHMMSS")
    Application.StatusBar = "The Scantron has been exported to " & Cpath
    Exit Sub
Else 'if
End If

handleIt:


End Sub

Private Sub ResetDs_Click() ' wipe Dep scantron
Dim qa As String: qa = MsgBox("Wipe Scantron?", vbYesNo, "Wipe Scantron")
    If qa = vbYes Then pCard.Range("A2:l9999").ClearContents
    updateDsDisp
    globalSave
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
dsVersion.Caption = "SENSEI DATA SCANTRON ver. " & config.Range("J2").Value & "-" & config.Range("J3").Value

initializeConfig ' load config
updateDsDisp
End Sub
Sub initializeConfig() ' Initialize simple configs

dsDateUni = config.Range("J4") ' unified date
dsDateUniCt.Value = config.Range("J5").Value ' unified parameter
dsDate14.Value = config.Range("J7").Value ' 14 23 65 FL PARAM
dsDate23.Value = config.Range("J8").Value
dsDate65.Value = config.Range("J9").Value
dsDateFL.Value = config.Range("J6").Value
dsPrintOptn = config.Range("J10").Value
dsPrintPath.Value = config.Range("J11").Value


'Update GUI
If dsDateUni Then ' if split enable else disable
    dsDateUniCt.Enabled = False
    dsDateUniAdj.Enabled = False

    dsDateUni.Caption = "ISOLATE"
    config.Range("J4") = True
    dsDateFL.Enabled = True
    dsDateFLadj.Enabled = True
    dsDate14.Enabled = True
    dsDate14adj.Enabled = True
    dsDate23.Enabled = True
    dsDate23adj.Enabled = True
    dsDate65.Enabled = True
    dsDate65adj.Enabled = True
    
    dsDatelblFL.Enabled = True
    dsDatelbl14.Enabled = True
    dsDatelbl23.Enabled = True
    dsDatelbl65.Enabled = True
Else
    dsDateUniCt.Enabled = True
    dsDateUniAdj.Enabled = True

    dsDateUni.Caption = "UNIFIED"
    config.Range("J4") = False
    dsDateFL.Enabled = False
    dsDateFLadj.Enabled = False
    dsDate14.Enabled = False
    dsDate14adj.Enabled = False
    dsDate23.Enabled = False
    dsDate23adj.Enabled = False
    dsDate65.Enabled = False
    dsDate65adj.Enabled = False
    
    dsDatelblFL.Enabled = False
    dsDatelbl14.Enabled = False
    dsDatelbl23.Enabled = False
    dsDatelbl65.Enabled = False
End If

If dsPrintOptn Then ' ALLOW FIXED PATH UPDATE THE PRINT CONTROL (THIS THING GETTING LONG)
    config.Range("J10").Value = True
    dsPrintOptn.Caption = "ASSIGN"
    dsPrintPath.Enabled = True
    dsPrintPathAssign.Enabled = True
    dsPrintPathRemove.Enabled = True
Else ' OR NOT
    config.Range("J10").Value = False
    dsPrintOptn.Caption = "MANUAL"
    dsPrintPath.Enabled = False
    dsPrintPathAssign.Enabled = False
    dsPrintPathRemove.Enabled = False
End If

na = vbNullString

End Sub


Sub readScantron() ' direct 114 to read scantron' Dep Scantron
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
    lastR = pCard.Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
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
