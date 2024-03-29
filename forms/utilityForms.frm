VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} utilityForms 
   Caption         =   "Sensei Form Distiller"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12555
   OleObjectBlob   =   "utilityForms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "utilityForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ########## INFO LIB ###########
' formlib is the multipage that contains diff forms
'
'
' ###############################
'
' ########## TODO LIST ##########
' [x] 2424 Form
' [ ] 117 Initialization, Operation and Exportation
'
'
'
'
'
' ###############################
' =========== GLOBAL VARIABLE ============
' gconfig_delwarn as global management
Public config As Worksheet ' config
Public formVer As Range, Sver As Range ' Form version and Sensei Version
Public saveTo As String, saveOptn As Boolean ' the location and is prompt allowed?
Const na As String = vbNullString ' the nothing string
' ========================================
' =========== FORM 110 VARIABLE ====================================================
' Data for Row and Range selection
Public f110Row As Long
' Data for 110 pages
Public F110p1 As Worksheet
Public F110p2 As Worksheet
' Data Erase Operational Range
Public taxTotal As Range
Public paidRateP1 As Range, dueRateP1 As Range, dueUSP1 As Range ' Portion that is $ - 110
Public paidRateP2 As Range, dueRateP2 As Range, dueUSP2 As Range ' p2
Public periodStartP1 As Range, periodEndP1 As Range ' Portion that is date - 110
Public periodStartP2 As Range, periodEndP2 As Range ' p2
Public itemP1 As Range, typeP1 As Range, gradeP1 As Range ' According 3 coloumns - 110
Public itemP2 As Range, typeP2 As Range, gradeP2 As Range ' p2
Public taxIso As Range ' isolated tax on p1
Public name110 As Range, ssn110 As Range ' person info
Public f110delW As Range ' delete indicator config
' Fixed Display Data
Public stPer As Range ' State %
Public debtTotal As Range
' State ID
Public stID As Range ' State ID
' Floating Variable for Edit Purpose
Public entryCount As Long, entryRow As Long ' number for entry editing
Public period1 As Range, period2 As Range ' Start and End
Public item1 As Range, type1 As Range ' Item and type
Public grade1 As Range, paid1 As Range, due1 As Range ' Grade, paid and due
Public fica1 As Range, med1 As Range, sitw1 As Range ' FICA$, MEDIC$, and SITW$
' Data Range for State Tax
Public Data As Worksheet, Tab1 As Range ' Data sheet and table range for SITW %
' Data for Display Page flip
Public DispPage As Long, Disp_all As Boolean, Disp_Link As Boolean
Public f110cSSNLink As Range
' ==================================================================================

' =========== FORM 2424 VARIABLE ===================================================
' Data for 2424 Pages
Public f2424 As Worksheet
' Data for Specific 2424 Fill Range
Public f2424amount As Range, f2424amountF As Range ' amount
Public f2424PA As Range, f2424PC As Range, f2424PK As Range, f2424PQ As Range ' options
Public f2424O As Range, f2424Oex As Range, f2424Cat As Range ' options others and addition
Public f2424SSN As Range, f2424Name As Range, f2424Rank As Range ' Member info
Public f2424expl As Object ' the text box for computation
' Data for 2424 config switches
Public f2424cDelw As Range, f2424cStartNew As Range ' delete warning, new page
Public f2424cMNA As Range, f2424cSSNlink As Range ' not available, ssn linkage
Public f2424cPrev As Range, f2424cPrevType As Range ' prior selected pay type
Public f2424Ptype As String, f2424Cancel As Boolean ' type flag and remove flag
' Data for 2424 input data for JPBB
Public f2424JPBBname As Range, f2424JPBBssn As Range, f2424JPBBamount As Range
' ==================================================================================

' =========== Form 117 VARIABLE ====================================================
' Data for 117 Pages
Public f117 As Worksheet
' Data for 117 applicable fields
Public f117vchNo As Range, f117ADSN As Range, f117payPurpose As Range ' Generic top form
Public f117Name As Range, f117Rank As Range, f117SSN As Range, f117Routing As Range, _
       f117Account As Range, f117Ptype As Range, f117Pdate As Range, _
       f117Amu As Range, f117EO As Range, f117accType As Range, _
       f117Service As Range ' Member Info, Pay type, paid date, paid amount, 530/510-EO
Public f117Lctn As Range, f117DO As Range, f117DSSN As Range, f117FSO As Range, _
       f117LOA As Range ' where is paying, authority, dssn, local office, accounting Info
Public f117PdDate As Range, f117prepBy As Range, f117receiveBy As Range ' MAKE AND TAKE
' Misc data for LOA Automation
Public thisFY As String, thisFYid As String
' Configs
Public f117cDelWarn As Range, f117cStartNew As Range, f117cDMO As Range, _
       f117cSSNLink As Range, f117cAutoPrep As Range, f117cAutoLOA As Range, _
       f117cDelTotal As Range
' Flags
Public f117PKPQ As Boolean, f117Nuking As Boolean
' ==================================================================================

Private Sub amountList_Change() ' 110 CHANGE LOGIC REQUIRED, STR IS PREFERRED
On Error GoTo handler
    If CDbl(amountList.Value) - CDbl(amountSum.Value) <> 0 Or amountList.Value = "" Or amountSum.Value = "" Then
        amountMatch.Value = "NG"
        amountMatch.BackColor = &H84496B
        amountMatch.BorderColor = &H84496B
    Else
        amountMatch.Value = "OK"
        amountMatch.BackColor = &HC000&
        amountMatch.BorderColor = &HC000&
    End If
    updateCompDebt ' UPDATE MATCH
    Exit Sub
handler:
    amountMatch.Value = "--"
    amountMatch.BackColor = &H84496B
    amountMatch.BorderColor = &H84496B
    updateCompDebt ' UPDATE MATCH
End Sub

Private Sub amountSum_Change() ' 110 UPDATE DISPLAY WHEN CHANGED
On Error GoTo handler
    If CDbl(amountList.Value) - CDbl(amountSum.Value) <> 0 Or amountList.Value = "" Or amountSum.Value = "" Then
        amountMatch.Value = "NG"
        amountMatch.BackColor = &H84496B
        amountMatch.BorderColor = &H84496B
    Else
        amountMatch.Value = "OK"
        amountMatch.BackColor = &HC000&
        amountMatch.BorderColor = &HC000&
    End If
    updateCompDebt ' UPDATE MATCH
    Exit Sub
handler:
    amountMatch.Value = "--"
    amountMatch.BackColor = &H84496B
    amountMatch.BorderColor = &H84496B
    updateCompDebt ' UPDATE MATCH
End Sub

Private Sub f110_delall_Click() ' 110 delete all items
f110nukeAll
End Sub
Sub f110nukeAll() ' 110 nuke function
Dim ResQ As String, npy As String
    npy = 0
    
If Gconfig_delWarn Or Not f110c_delWarn Then
        ResQ = MsgBox("Reset form?", vbYesNo, "Form Distiller")
    If ResQ = vbNo Then Exit Sub
End If

    paidRateP1 = npy
    paidRateP2 = npy
    dueUSP1 = na
    dueUSP2 = na
    dueRateP1 = npy
    dueRateP2 = npy
    periodStartP1 = na
    periodStartP2 = na
    periodEndP1 = na
    periodEndP2 = na
    itemP1 = na
    itemP2 = na
    typeP1 = na
    typeP2 = na
    gradeP1 = na
    gradeP2 = na
    name110 = na
    ssn110 = na
    If Not f110c_keepMbr Then ' do when unprotected
        f110_name.Value = na
        f110_ssn.Value = na
    End If
    f110rowDispUpdate
    
update110Display
End Sub

Private Sub f110_delone_Click() ' remove one entry
If Gconfig_delWarn Or Not f110c_delWarn Then
    Dim ResQ As String
        ResQ = MsgBox("Delete this Entry?", vbYesNo, "Form Distiller")
    If ResQ = vbNo Then Exit Sub
End If
If f110_PageCt.Value = "P.1" Then ' PAGE ONE
    With F110p1
        f110_strDate.Value = na
        f110_endDate.Value = na
        f110_itemName.Value = na
        f110_itemType.Value = na
        f110_itemGrade.Value = na
        f110_paidRate.Value = 0
        f110_dueRate.Value = 0
        f110_dueUS.Value = na
        f110_dueClaimant.Caption = Format(.Range("M" & f110Row).Value, "$0.00")
    End With
ElseIf f110_PageCt.Value = "P.2" Then ' PAGE TWO
    With F110p2
        f110_strDate.Value = na
        f110_endDate.Value = na
        f110_itemName.Value = na
        f110_itemType.Value = na
        f110_itemGrade.Value = na
        f110_paidRate.Value = 0
        f110_dueRate.Value = 0
        f110_dueUS.Value = na
        f110_dueClaimant.Caption = Format(.Range("M" & f110Row).Value, "$0.00")
    End With
End If

f110rowDispUpdate
End Sub

Private Sub f110_dispM1_Click() ' 110 goto previous page
If Disp_all Or Disp_Link Then Exit Sub
If DispPage > 1 Then ' a simple minus one page
    DispPage = DispPage - 1
Else
    DispPage = 4
End If
f110_dispP.Caption = DispPage
update110Display ' DO SOME AUTO UPDATE WHEN CHANGE PAGE
End Sub

Private Sub f110_dispM2_Click() ' 110 goto next page
If Disp_all Or Disp_Link Then Exit Sub
If DispPage < 4 Then ' a simple plus one page
    DispPage = DispPage + 1
Else
    DispPage = 1
End If
f110_dispP.Caption = DispPage
update110Display ' SAME HERE, AUTO PAGE INFO UPDATE
End Sub

Private Sub f110_dueRate_Change() ' 110 update due amount
    On Error GoTo WND
    If CDbl(f110_dueRate.Value) > 99999.99 Then f110_dueRate.Value = 99999.99
    If CDbl(f110_dueRate.Value) < -9999.99 Then f110_dueRate.Value = -9999.99
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("K" & f110Row).Value = f110_dueRate.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("K" & f110Row).Value = f110_dueRate.Value
End If

WND:
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_dueUS_Change() ' 110 update due us amount
    On Error GoTo WND
    If CDbl(f110_dueUS.Value) > 99999999999.99 Then f110_dueUS.Value = 99999999999.99
    If CDbl(f110_dueUS.Value) < -9999999999.99 Then f110_dueUS.Value = -9999999999.99
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("N" & f110Row).Value = f110_dueUS.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("N" & f110Row).Value = f110_dueUS.Value
End If

WND:
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_endDate_Change() ' 110 edit end date
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("C" & f110Row).Value = f110_endDate.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("C" & f110Row).Value = f110_endDate.Value
End If
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_export_Click() ' 110 export function, add support for fixed directory
globalSave
Dim saveToPrompt
Dim Cpath As String: Cpath = na ' current path contains no name
Dim Cexist As String

    saveTo = config.Range("F6").Value
On Error GoTo handleIt
If Not saveOptn Or saveTo = "" Then ' ALWAYS PROMPT IF DISABLED PATHWAY OR SAVETO IS BLANK
    Set saveToPrompt = Application.FileDialog(msoFileDialogFolderPicker)
    With saveToPrompt
        .Title = "Sending Form 110 to here..."
        .ButtonName = "Save"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = 0 Then Exit Sub ' if cancelled exit immediately
        saveTo = .SelectedItems(1)
    End With
Else
    Cexist = Dir(saveTo & "\Form 110 Exports\", vbDirectory) ' check parent
    If Cexist = "" Then
        Cexist = saveTo & "\Form 110 Exports\"
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    Cexist = Dir(saveTo & "\Form 110 Exports\" & Format(Now(), "YYYY-MM"), vbDirectory) ' check child
    If Cexist = "" Then
        Cexist = saveTo & "\Form 110 Exports\" & Format(Now(), "YYYY-MM")
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    Cpath = saveTo & "\Form 110 Exports\" & Format(Now(), "YYYY-MM") ' assign to Cpath
End If


exportForm:
Set saveToPrompt = Nothing ' UNLOAD OBJECT >>Does not save any way ?
If Not saveOptn Or saveTo = "" Then Cpath = saveTo  ' only wrote when fixed path is not activated

' add a Mkdir, or make directory for when under constant method
    F110p1.ExportAsFixedFormat xlTypePDF, _
        Filename:=Cpath & "\110.COMP." & Left(f110_name.Value, 5) & "." & Format(Now(), "YYMMDD-HHMMSS") & ".01"
    If F110p1.Range("M18").Value <> 0 Then F110p2.ExportAsFixedFormat xlTypePDF, _
        Filename:=Cpath & "\110.COMP." & Left(f110_name.Value, 5) & "." & Format(Now(), "YYMMDD-HHMMSS") & ".02"
    Application.StatusBar = "Form 110 has been exported to " & Cpath
    Exit Sub
    If f110c_startNew Then f110nukeAll
handleIt:

End Sub

Private Sub f110_fica_Change() ' 110 fica
    On Error GoTo WND
    If CDbl(f110_fica.Value) > 99999.99 Then f110_fica.Value = 99999.99
    If CDbl(f110_fica.Value) < -9999.99 Then f110_fica.Value = -9999.99
    F110p1.Range("J21").Value = f110_fica.Value
    f110_ficaD.Value = Round(F110p1.Range("L21").Value, 2)
    
WND:
f110rowDueAmuUpdate
updateCompDebt
End Sub

Private Sub f110_inherit_Click() ' 110 Inherit from previous Entry
    f110rowDispInherit ' inherit function
    globalSave
End Sub

Private Sub f110_itemGrade_Change() ' 110 GRADE/YEAR
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("J" & f110Row).Value = f110_itemGrade.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("J" & f110Row).Value = f110_itemGrade.Value
End If
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_itemName_Change() ' 110 ITEM
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("D" & f110Row).Value = f110_itemName.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("D" & f110Row).Value = f110_itemName.Value
End If
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_itemType_Change() ' 110 UPDATE ITEM TYPE
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("E" & f110Row).Value = f110_itemType.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("E" & f110Row).Value = f110_itemType.Value
End If
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_med_Change() ' 110 medicare
    On Error GoTo WND
    If CDbl(f110_med.Value) > 99999.99 Then f110_med.Value = 99999.99
    If CDbl(f110_med.Value) < -9999.99 Then f110_med.Value = -9999.99
    F110p1.Range("J22").Value = f110_med.Value
    f110_medD.Value = Round(F110p1.Range("L22").Value, 2)
    
WND:
f110rowDueAmuUpdate
updateCompDebt
End Sub

Private Sub f110_name_Change() ' 110 name
    name110.Value = f110_name.Value
End Sub

Private Sub f110_paidRate_Change() ' 110 update paid amount
    On Error GoTo WND
    If CDbl(f110_paidRate.Value) > 99999.99 Then f110_paidRate.Value = 99999.99
    If CDbl(f110_paidRate.Value) < -9999.99 Then f110_paidRate.Value = -9999.99
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("H" & f110Row).Value = f110_paidRate.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("H" & f110Row).Value = f110_paidRate.Value
End If
    
WND:
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_Rank_Change()

If Left(f110_Rank.Value, 1) = "E" Then f117EO.Value = 530
If Left(f110_Rank.Value, 1) = "O" Then f117EO.Value = 510
f117Rank.Value = "  " & f110_Rank.Text
update117display

End Sub

Private Sub f110_sitw_Change() ' 110 fitw
    On Error GoTo WND
    If CDbl(f110_sitw.Value) > 99999.99 Then f110_sitw.Value = 99999.99
    If CDbl(f110_sitw.Value) < -9999.99 Then f110_sitw.Value = -9999.99
    F110p1.Range("J23").Value = f110_sitw.Value
    f110_sitwD.Value = Round(F110p1.Range("L23").Value, 2)
    
WND:
f110rowDueAmuUpdate
updateCompDebt
End Sub

Private Sub f110_ssn_Change() ' 110 SSN
    ssn110.Value = f110_ssn.Value
End Sub

Private Sub f110_strDate_Change() ' 110 Edit Start date
If f110_PageCt.Value = "P.1" Then ' page 1
    F110p1.Range("A" & f110Row).Value = f110_strDate.Value
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    F110p2.Range("A" & f110Row).Value = f110_strDate.Value
End If
f110rowDueAmuUpdate ' update due amount
update110Display
End Sub

Private Sub f110_SwitchRow_SpinUp() ' 110 Row selector minus
    f110_EntryCt.Value = f110_EntryCt.Value - 1
    If f110_EntryCt.Value < 1 Then f110_EntryCt.Value = 35
      
    f110rowChange
    f110rowSwitchLink
    f110Row = f110_RowCt.Value ' Make row Number valid
    f110rowDispUpdate ' UPDATE FIELD
End Sub
Private Sub f110_SwitchRow_SpinDown() ' 110 row selector plus
    f110_EntryCt.Value = f110_EntryCt.Value + 1
    If f110_EntryCt.Value > 35 Then f110_EntryCt.Value = 1
    
    f110rowChange
    f110rowSwitchLink
    f110Row = f110_RowCt.Value ' Make row Number valid
    f110rowDispUpdate ' UPDATE FIELD
End Sub
Sub f110rowSwitchLink() ' if dislay is linked, do this
If Disp_Link Then ' if display is linked to spin
    If f110_EntryCt.Value < 11 Then
        DispPage = 1
    ElseIf f110_EntryCt.Value < 21 Then
        DispPage = 2
    ElseIf f110_EntryCt.Value < 31 Then
        DispPage = 3
    Else
        DispPage = 4
    End If
    f110_dispP.Caption = DispPage
    update110Display
End If
End Sub
Sub f110rowChange() ' 110 row selector adjusted
    If f110_EntryCt.Value < 14 Then
        f110_PageCt.Value = "P.1"
        f110_RowCt.Value = f110_EntryCt.Value + 4
    Else
        f110_PageCt.Value = "P.2"
        f110_RowCt.Value = f110_EntryCt.Value - 9
    End If
End Sub
Sub f110rowDispUpdate() ' 110 entry initial loading
If f110_PageCt.Value = "P.1" Then ' PAGE ONE
    With F110p1
        f110_strDate.Value = Format(.Range("A" & f110Row).Value, "YYYY-MM-DD")
        f110_endDate.Value = Format(.Range("C" & f110Row).Value, "YYYY-MM-DD")
        f110_itemName.Value = .Range("D" & f110Row).Value
        f110_itemType.Value = .Range("E" & f110Row).Value
        f110_itemGrade.Value = .Range("J" & f110Row).Value
        f110_paidRate.Value = .Range("H" & f110Row).Value
        f110_dueRate.Value = .Range("K" & f110Row).Value
        f110_dueUS.Value = .Range("N" & f110Row).Value
        f110_dueClaimant.Caption = Format(.Range("M" & f110Row).Value, "$0.00")
        SITWdrop.Value = .Range("E23").Value
        f110_sitw.Value = .Range("J23").Value
        f110_med.Value = .Range("J22").Value
        f110_fica.Value = .Range("J21").Value
    End With
ElseIf f110_PageCt.Value = "P.2" Then ' PAGE TWO
    With F110p2
        f110_strDate.Value = Format(.Range("A" & f110Row).Value, "YYYY-MM-DD")
        f110_endDate.Value = Format(.Range("C" & f110Row).Value, "YYYY-MM-DD")
        f110_itemName.Value = .Range("D" & f110Row).Value
        f110_itemType.Value = .Range("E" & f110Row).Value
        f110_itemGrade.Value = .Range("J" & f110Row).Value
        f110_paidRate.Value = .Range("H" & f110Row).Value
        f110_dueRate.Value = .Range("K" & f110Row).Value
        f110_dueUS.Value = .Range("N" & f110Row).Value
        f110_dueClaimant.Caption = Format(.Range("M" & f110Row).Value, "$0.00")
    End With
End If
End Sub
Sub f110rowDispInherit() ' 110 entry inherition from previous
' 5-17 page one; 5-26 page two
Dim isPrev As Long, isFine As String ' see if current line exist
If f110Row = 5 And f110_PageCt.Value = "P.1" Then ' page one first entry nothing happen
    MsgBox "This is the first entry...", vbOKOnly, "Distiller Form 110"
ElseIf f110Row > 5 And f110Row < 18 And f110_PageCt.Value = "P.1" Then ' PAGE ONE others
    With F110p1
        isPrev = Application.WorksheetFunction.CountA(.Range("A" & f110Row & ":" & "N" & f110Row))
            If isPrev <> 8 And Not f110c_inherit Then ' if current not blank, huston we have a problem
                isFine = MsgBox("You are about to overwrite this line, Continue?", vbYesNo, "Distiller Form 110")
                If isFine = vbNo Then Exit Sub
            End If
        f110_strDate.Value = .Range("A" & f110Row - 1).Value
        f110_endDate.Value = .Range("C" & f110Row - 1).Value
        f110_itemName.Value = .Range("D" & f110Row - 1).Value
        f110_itemType.Value = .Range("E" & f110Row - 1).Value
        f110_itemGrade.Value = .Range("J" & f110Row - 1).Value
        f110_paidRate.Value = .Range("H" & f110Row - 1).Value
        f110_dueRate.Value = .Range("K" & f110Row - 1).Value
        f110_dueUS.Value = .Range("N" & f110Row - 1).Value
        f110_dueClaimant.Caption = Format(.Range("M" & f110Row - 1).Value, "$0.00")
        SITWdrop.Value = .Range("E23").Value
        f110_sitw.Value = .Range("J23").Value
        f110_med.Value = .Range("J22").Value
        f110_fica.Value = .Range("J21").Value
    End With
ElseIf f110Row = 5 And f110_PageCt.Value = "P.2" Then ' first in page two inherit p1 end
    With F110p2
        isPrev = Application.WorksheetFunction.CountA(.Range("A" & f110Row & ":" & "N" & f110Row))
            If isPrev <> 8 And Not f110c_inherit Then ' if current not blank, huston we have a problem
                isFine = MsgBox("You are about to overwrite this line, Continue?", vbYesNo, "Distiller Form 110")
                If isFine = vbNo Then Exit Sub
            End If
    End With
    With F110p1
        f110_strDate.Value = .Range("A17").Value
        f110_endDate.Value = .Range("C17").Value
        f110_itemName.Value = .Range("D17").Value
        f110_itemType.Value = .Range("E17").Value
        f110_itemGrade.Value = .Range("J17").Value
        f110_paidRate.Value = .Range("H17").Value
        f110_dueRate.Value = .Range("K17").Value
        f110_dueUS.Value = .Range("N17").Value
        f110_dueClaimant.Caption = Format(.Range("M17").Value, "$0.00")
    End With
ElseIf f110Row > 5 And f110Row < 27 And f110_PageCt.Value = "P.2" Then ' page two others
    With F110p2
        isPrev = Application.WorksheetFunction.CountA(.Range("A" & f110Row & ":" & "N" & f110Row))
            If isPrev <> 8 And Not f110c_inherit Then ' if current not blank, huston we have a problem
                isFine = MsgBox("You are about to overwrite this line, Continue?", vbYesNo, "Distiller Form 110")
                If isFine = vbNo Then Exit Sub
            End If
        f110_strDate.Value = .Range("A" & f110Row - 1).Value
        f110_endDate.Value = .Range("C" & f110Row - 1).Value
        f110_itemName.Value = .Range("D" & f110Row - 1).Value
        f110_itemType.Value = .Range("E" & f110Row - 1).Value
        f110_itemGrade.Value = .Range("J" & f110Row - 1).Value
        f110_paidRate.Value = .Range("H" & f110Row - 1).Value
        f110_dueRate.Value = .Range("K" & f110Row - 1).Value
        f110_dueUS.Value = .Range("N" & f110Row - 1).Value
        f110_dueClaimant.Caption = Format(.Range("M" & f110Row - 1).Value, "$0.00")
    End With
End If
End Sub
Sub f110rowDueAmuUpdate() ' 110 due claimant update
If f110_PageCt.Value = "P.1" Then ' page 1
    f110_dueClaimant.Caption = Format(F110p1.Range("M" & f110Row).Value, "$0.00")
ElseIf f110_PageCt.Value = "P.2" Then ' page 2
    f110_dueClaimant.Caption = Format(F110p2.Range("M" & f110Row).Value, "$0.00")
End If
End Sub

Private Sub f110c_delWarn_Click() ' 110 DELETE WARNING
If f110c_delWarn Then
    f110delW.Value = True
    f110c_delWarn.Caption = "Mute"
Else
    f110delW.Value = False
    f110c_delWarn.Caption = "Warn"
End If

End Sub

Private Sub f110c_dispAll_Click() ' 110 When this is enabled, display all entries regardless
If f110c_dispAll Then
    f110c_dispAll.Caption = "Disp. All"
    f110c_dispFollow = False
    Disp_all = True
Else
    f110c_dispAll.Caption = "This Page"
    Disp_all = False
End If
config.Range("F33").Value = Disp_all
update110Display ' immediately flip the display
End Sub

Private Sub f110c_dispFollow_Click() ' 110 Make display follows user's editor number
If f110c_dispFollow Then
    f110c_dispFollow.Caption = "Linked"
    f110c_dispAll = False
    Disp_Link = True
Else
    f110c_dispFollow.Caption = "Split"
    Disp_Link = False
End If
config.Range("F34").Value = Disp_Link
f110rowSwitchLink ' immediately inflict linkage update
End Sub

Private Sub f110c_inherit_Click() ' 110 inherit warning control
If f110c_inherit Then
    f110c_inherit.Caption = "Mute"
Else
    f110c_inherit.Caption = "Warn"
End If
config.Range("F36").Value = f110c_inherit
End Sub

Private Sub f110c_keepMbr_Click() ' 110 do we want to keep personal info?
If f110c_keepMbr Then
    f110c_keepMbr.Caption = "Retain"
Else
    f110c_keepMbr.Caption = "Discard"
End If
config.Range("F35").Value = f110c_keepMbr
End Sub

Private Sub f110c_SSNlookup_Click() ' 110 SSN look up

' Require function of dictionary write, dictionary append and dictionary look
If f110c_SSNlookup.Value Then f110c_SSNlookup.Caption = "Enabled"
If Not f110c_SSNlookup.Value Then f110c_SSNlookup.Caption = "Disabled"
f110cSSNLink.Value = f110c_SSNlookup.Value

End Sub

Private Sub f110c_startNew_Click() ' 110 Export erase option
If f110c_startNew Then
    config.Range("F38").Value = True
    f110c_startNew.Caption = "Enabled"
Else
    config.Range("F38").Value = False
    f110c_startNew.Caption = "Disabled"
End If
End Sub

Private Sub f117_Account_Change()

f117Account.Value = f117_Account.Value

End Sub

Private Sub f117_accType_Click()

If f117_accType.Value Then
    f117accType.Value = "Saving"
Else
    f117accType.Value = "Checking"
End If
update117display

End Sub

Private Sub f117_ADSN_Change() ' 117 ADSN

f117ADSN.Value = f117_ADSN.Text

End Sub

Private Sub f117_Amu_Change() ' 117 $ AMOUNT

f117Amu.Value = "$" & Format(f117_Amu.Value, "0.00")

End Sub

Private Sub f117_DO_Change()

f117DO.Value = f117_DO.Value

End Sub

Private Sub f117_DSSN_Change()

f117DSSN.Value = f117_DSSN.Value

End Sub

Private Sub f117_editPaidBy_Click() ' 117 EDIT LOCK

If f117_editPaidBy.Value Then
    f117_editPaidBy.Caption = "Editing Mode"
    f117_Erase.Enabled = False
    f117_Export.Enabled = False
    ' SERIES OF DISABLING AND ENABLING <DISABLE EXECUTIVE OPTION>
Else
    f117_editPaidBy.Caption = "Operating Mode"
    f117_Erase.Enabled = True
    f117_Export.Enabled = True
    ' SERIES OF DISABLING AND ENABLING <REENGAGE THEM>
End If
update117display

End Sub

Private Sub f117_Erase_Click()

globalSave
nukef117
initialize117a
f117Nuking = False

End Sub

Sub nukef117()

f117Nuking = True
Dim Qres As String
If f117cDelWarn.Value = False Then  ' prompt only if warn is on
    Qres = MsgBox("Reset form?", vbYesNo, "Form Distiller")
    If Qres = vbNo Then Exit Sub
End If

' the basic information
Dim unionBase117 As Range: Set unionBase117 = Union(f117vchNo, f117Name, _
                                            f117Rank, f117SSN, f117Routing, _
                                            f117Account, f117Pdate, f117Amu, _
                                            f117Service, f117PdDate, f117prepBy, _
                                            f117receiveBy, f117accType)
' the detailed payment information that are usually constant
Dim unionINFO117 As Range: Set unionINFO117 = Union(f117ADSN, f117Lctn, f117DO, _
                                                    f117DSSN, f117FSO)
'total delete if enabled
If f117c_DelTotal.Value Then unionINFO117.Value = na
unionBase117.Value = na
f117_Amu.Value = 0

End Sub

Private Sub f117_Export_Click()
If InStr(f117_Name.Text, " ") = 0 Or f117_Name.Value = "" Then Exit Sub ' kick if empty

globalSave
Dim saveToPrompt
Dim Cpath As String: Cpath = na ' current path contains no name
Dim Cexist As String

    saveTo = config.Range("F6").Value
On Error GoTo handleIt
If Not saveOptn Or saveTo = "" Then ' ALWAYS PROMPT IF DISABLED PATHWAY OR SAVETO IS BLANK
    Set saveToPrompt = Application.FileDialog(msoFileDialogFolderPicker)
    With saveToPrompt
        .Title = "Sending Form 117 to here..."
        .ButtonName = "Save"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = 0 Then Exit Sub ' if cancelled exit immediately
        saveTo = .SelectedItems(1)
    End With
Else
    Cexist = Dir(saveTo & "\Form 117 Exports\", vbDirectory) ' check parent
    If Cexist = "" Then
        Cexist = saveTo & "\Form 117 Exports\"
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    Cexist = Dir(saveTo & "\Form 117 Exports\" & Format(Now(), "YYYY-MM"), vbDirectory) ' check child
    If Cexist = "" Then
        Cexist = saveTo & "\Form 117 Exports\" & Format(Now(), "YYYY-MM")
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    Cpath = saveTo & "\Form 117 Exports\" & Format(Now(), "YYYY-MM") ' assign to Cpath
End If


exportForm:
Set saveToPrompt = Nothing ' UNLOAD OBJECT >>Does not save any way ?
If Not saveOptn Or saveTo = "" Then Cpath = saveTo ' only wrote when fixed path is not activated

' add a Mkdir, or make directory for when under constant method
    f117.ExportAsFixedFormat xlTypePDF, _
        Filename:=Cpath & "\117." & f117_vchNo & "-" & nameInverter(f117_Name.Text) & Format(Now(), "YYMMDD-HHMMSS")
    Application.StatusBar = "Form 117 has been exported to " & Cpath
    If f117c_StartNew Then nukef117
    Exit Sub
    
handleIt:


End Sub

Private Sub f117_FSO_Change()

f117FSO.Value = f117_FSO.Value

End Sub

Private Sub f117_Lctn_Change()

f117Lctn.Value = f117_Lctn.Value

End Sub

Private Sub f117_LOA_Change() ' 117 LOA

f117LOA.Value = f117_LOA.Text

End Sub

Private Sub f117_Name_Change() ' 117 Name invert

f117Name = nameInverter(f117_Name.Text)

End Sub

Private Sub f117_payPurpose_Change() ' 117 CHANGE REASON

f117payPurpose.Value = f117_payPurpose.Text

End Sub

Private Sub F117_Pdate_Change() ' 117 PREPARE DATE

f117Pdate.Value = f117_Pdate.Value

End Sub

Private Sub f117_PdDate_Change()

f117PdDate.Value = f117_PdDate.Value

End Sub

Private Sub f117_PK_Click() ' 117 PK PQ MUTUALLY EXCLUSIVE

If f117PKPQ Then Exit Sub ' do not execute if found it sequential

f117PKPQ = True ' PROTECT
If f117_PK.Value Then
    F117_PQ.Value = False
    f117Ptype.Value = "PK03"
    f117_payPurpose.Text = "PARTIAL PAYMENT THRU CAGE UNDER SPECIAL CIRCUMSTANCES"
Else
    f117Ptype.Value = ""
    f117_payPurpose.Text = ""
End If
f117PKPQ = False

End Sub

Private Sub F117_PQ_Click() ' 117 PK PQ MUTUALLY EXCLUSIVE

If f117PKPQ Then Exit Sub ' do not execute if found it sequential

f117PKPQ = True ' PROTECT
If F117_PQ.Value Then
    f117_PK.Value = False
    f117Ptype.Value = "PQ03"
    f117_payPurpose.Text = "SEP/RET PAYMENT THRU CAGE FOR " & """V - SEPARATED""" & " ACCOUNT"
Else
    f117Ptype.Value = ""
    f117_payPurpose.Text = ""
End If
f117PKPQ = False

End Sub

Private Sub f117_prepBy_Change()

f117prepBy.Value = f117_prepBy.Value

End Sub

Private Sub f117_Rank_Change()

f117Rank.Value = f117_Rank.Value
update117display

End Sub

Private Sub f117_receiveBy_Change()

f117receiveBy.Value = f117_receiveBy.Value

End Sub

Private Sub f117_Routing_Change()

f117Routing.Value = f117_Routing.Value

End Sub

Private Sub f117_Service_Change() ' 117 AGENCY PAID CODE

f117Service.Value = Left(f117_Service.Text, 1)

End Sub

Private Sub f117_SSN_Change()

f117SSN.Value = f117_SSN.Value

End Sub

Private Sub f117_vchNo_Change()

f117vchNo.Value = f117_vchNo.Value

End Sub

Private Sub f117c_AutoLOA_Click() ' 117c Adaptive LOA

If f117c_AutoLOA.Value Then f117c_AutoLOA.Caption = "Enabled"
If Not f117c_AutoLOA.Value Then f117c_AutoLOA.Caption = "Disabled"
f117cAutoLOA.Value = f117c_AutoLOA.Value
update117display

End Sub

Private Sub f117c_AutoPrep_Click() ' 117c prep date assume

If f117c_AutoPrep.Value Then f117c_AutoPrep.Caption = "Enabled"
If Not f117c_AutoPrep.Value Then f117c_AutoPrep.Caption = "Disabled"
f117cAutoPrep.Value = f117c_AutoPrep.Value
update117display

End Sub

Private Sub f117c_DelTotal_Click()

If f117c_DelTotal.Value Then f117c_DelTotal.Caption = "Enabled"
If Not f117c_DelTotal.Value Then f117c_DelTotal.Caption = "Disabled"
f117cDelTotal.Value = f117c_DelTotal.Value
update117display

End Sub

Private Sub f117c_DelWarn_Click() ' 117c mute warning

If f117c_DelWarn.Value Then f117c_DelWarn.Caption = "Mute"
If Not f117c_DelWarn.Value Then f117c_DelWarn.Caption = "Warn"
f117cDelWarn.Value = f117c_DelWarn.Value
update117display

End Sub

Private Sub f117c_DMO_Click() ' 117c DMO Forwarding - develop after 114 Core completed

If f117c_DMO.Value Then f117c_DMO.Caption = "Enabled"
If Not f117c_DMO.Value Then f117c_DMO.Caption = "Disabled"
f117cDMO.Value = f117c_DMO.Value
update117display

End Sub

Private Sub f117c_SSNLink_Click() ' 117c LINK DATA - develop after Alpha Core complete

If f117c_SSNLink.Value Then f117c_SSNLink.Caption = "Enabled"
If Not f117c_SSNLink.Value Then f117c_SSNLink.Caption = "Disabled"
f117cSSNLink.Value = f117c_SSNLink.Value
update117display

End Sub

Private Sub f117c_StartNew_Click() ' 117c clear after export

If f117c_StartNew.Value Then f117c_StartNew.Caption = "Enabled"
If Not f117c_StartNew.Value Then f117c_StartNew.Caption = "Disabled"
f117cStartNew.Value = f117c_StartNew.Value
update117display

End Sub

Private Sub f2424_amount_Change() ' 2424 place an amount

f2424amount.Value = f2424_amount.Value


On Error GoTo reducedExit ' block error
If f2424_typePQ.Value Then '  PQ - 9 DIGIT
    f2424JPBBamount.Value = "'" & Format(f2424_amount.Value * 100, "000000000")
ElseIf Not f2424_typeO.Value Then ' EVERYTHING ELSE - 7 DIGIT
    f2424JPBBamount.Value = "'" & Format(f2424_amount.Value * 100, "0000000")
Else
reducedExit:
    f2424JPBBamount.Value = "INPUT WITH CAUTION"
End If

On Error GoTo problems ' do not need error to break the thing, wrote the words

If f2424_amount.Value <> na Then
    f2424_amountFig.Caption = SpellDollar(Format(f2424_amount.Value, "0.00"))
Else
    f2424_amountFig.Caption = ""
End If
f2424amountF.Value = f2424_amountFig.Caption
Exit Sub

problems:
End Sub

Private Sub f2424_delall_Click() ' 2424 nuke the form
f2424nuke
globalSave
End Sub

Sub f2424nuke() ' 2424 Delete file

Dim Qres As String
If Not f2424c_DelWarn Or Gconfig_delWarn Then ' prompt only if warn is on
    Qres = MsgBox("Reset form?", vbYesNo, "Form Distiller")
    If Qres = vbNo Then Exit Sub
End If

Dim unionPay As Range: Set unionPay = Union(f2424O, f2424PA, _
                                      f2424PC, f2424PK, f2424PQ, f2424Oex, f2424amountF)
    unionPay.Value = na
f2424_amount.Value = na
f2424_amountFig.Caption = na

f2424Cancel = True ' do not trigger loop
    f2424_typePA.Value = False
    f2424_typePC.Value = False
    f2424_typePK.Value = False
    f2424_typePQ.Value = False
    f2424_typeO.Value = False
f2424Cancel = False ' reset flag

f2424_typeOex.Value = na
f2424_typeCat.Value = na
f2424_mbrSSN.Value = na
f2424_mbrRank.Value = na
f2424_mbrName.Value = na
f2424_explain.Value = na
f2424.Range("B23").Value = na

If f2424c_Prev Then ' adapt the prior transaction type
    f2424Ptype = f2424cPrevType.Value
    f2424payType
    If f2424PA.Value = "X" Then f2424_typePA.Value = True
    If f2424PC.Value = "X" Then f2424_typePC.Value = True
    If f2424PK.Value = "X" Then f2424_typePK.Value = True
    If f2424PQ.Value = "X" Then f2424_typePQ.Value = True
    If f2424O.Value = "X" Then f2424_typeO.Value = True
End If

End Sub

Private Sub f2424_explain_Change() ' 2424 additional comments
f2424expl.Text = f2424_explain.Value
f2424_explainCount.Caption = 1000 - Len(f2424_explain.Text)
End Sub


Private Sub f2424_export_Click() ' 2424 Export, now support fixed directory
globalSave
Dim saveToPrompt
Dim Cpath As String: Cpath = na ' current path contains no name
Dim Cexist As String

    saveTo = config.Range("F6").Value
On Error GoTo handleIt
If Not saveOptn Or saveTo = "" Then ' ALWAYS PROMPT IF DISABLED PATHWAY OR SAVETO IS BLANK
    Set saveToPrompt = Application.FileDialog(msoFileDialogFolderPicker)
    With saveToPrompt
        .Title = "Sending Form 2424 to here..."
        .ButtonName = "Save"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = 0 Then Exit Sub ' if cancelled exit immediately
        saveTo = .SelectedItems(1)
    End With
Else
    Cexist = Dir(saveTo & "\Form 2424 Exports\", vbDirectory) ' check parent
    If Cexist = "" Then
        Cexist = saveTo & "\Form 2424 Exports\"
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    Cexist = Dir(saveTo & "\Form 2424 Exports\" & Format(Now(), "YYYY-MM"), vbDirectory) ' check child
    If Cexist = "" Then
        Cexist = saveTo & "\Form 2424 Exports\" & Format(Now(), "YYYY-MM")
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    Cpath = saveTo & "\Form 2424 Exports\" & Format(Now(), "YYYY-MM") ' assign to Cpath
End If


exportForm:
Set saveToPrompt = Nothing ' UNLOAD OBJECT >>Does not save any way ?
If Not saveOptn Or saveTo = "" Then Cpath = saveTo ' only wrote when fixed path is not activated

' add a Mkdir, or make directory for when under constant method
    f2424.ExportAsFixedFormat xlTypePDF, _
        Filename:=Cpath & "\2424." & Left(f2424_mbrName.Value, 5) & "." & Format(Now(), "YYMMDD-HHMMSS")
    Application.StatusBar = "Form 2424 has been exported to " & Cpath
    If f2424c_StartNew Then f2424nuke
    Exit Sub
    
handleIt:

End Sub

Private Sub f2424_mbrName_Change() ' 2424 Member Name
f2424Name = f2424_mbrName.Value
f2424JPBBname.Value = Left(Replace(f2424_mbrName.Text, ",", ""), 5) ' replace the comma regardless
End Sub

Private Sub f2424_mbrRank_Change() ' 2424 Member rank
f2424Rank.Value = f2424_mbrRank.Value
End Sub

Private Sub f2424_mbrSSN_Change() ' 2424 Member SSN
If f2424cSSNlink Then ' if found Link Setting is on then
    f2424_mbrName.Enabled = False
    f2424_mbrRank.Enabled = False
    ' do lookup
Else
    f2424_mbrName.Enabled = True
    f2424_mbrRank.Enabled = True
    ' lock name and Rank then find them
End If

f2424SSN.Value = f2424_mbrSSN.Value
f2424JPBBssn.Value = f2424_mbrSSN.Value
End Sub

Private Sub f2424_typeCat_Change() ' 2424 Category of Adv Pay
f2424Cat.Value = f2424_typeCat.Value
End Sub

Private Sub f2424_typeO_Click() ' 2424 Other Pay
If f2424Cancel Then Exit Sub
f2424Ptype = "PO"
f2424cPrevType.Value = f2424Ptype
f2424payType
f2424Cancel = False
End Sub

Private Sub f2424_typeOex_Change() ' 2424 Specification of Others
f2424Oex.Value = f2424_typeOex
End Sub

Private Sub f2424_typePA_Click() ' 2424 PA pay
If f2424Cancel Then Exit Sub
f2424Ptype = "PA"
f2424cPrevType.Value = f2424Ptype
f2424payType
f2424Cancel = False
End Sub

Private Sub f2424_typePC_Click() ' 2424 PC pay
If f2424Cancel Then Exit Sub
f2424Ptype = "PC"
f2424cPrevType.Value = f2424Ptype
f2424payType
f2424Cancel = False
End Sub

Private Sub f2424_typePK_Click() ' 2424 PK pay
If f2424Cancel Then Exit Sub
f2424Ptype = "PK"
f2424cPrevType.Value = f2424Ptype
f2424payType
f2424Cancel = False
End Sub

Private Sub f2424_typePQ_Click() ' 2424 PQ pay
If f2424Cancel Then Exit Sub
f2424Ptype = "PQ"
f2424cPrevType.Value = f2424Ptype
f2424payType
f2424Cancel = False
End Sub
Sub f2424payType() ' 2424 Adjustment for all options, not the best, but it works 221121
Dim unionPay As Range: Set unionPay = Union(f2424O, f2424PA, f2424PC, f2424PK, f2424PQ)
    unionPay.Value = na
    
If f2424Ptype = "PO" Then
    f2424Cancel = True
    f2424_typePA.Value = False
    f2424_typePC.Value = False
    f2424_typePK.Value = False
    f2424_typePQ.Value = False
    f2424_typeOex.Enabled = True
        f2424Oex.Value = f2424_typeOex
    f2424O.Value = "X"
    GoTo JPBBupdater
End If
If f2424Ptype = "PA" Then
    f2424Cancel = True
    f2424_typeO.Value = False
    f2424_typePC.Value = False
    f2424_typePK.Value = False
    f2424_typePQ.Value = False
    f2424_typeOex.Enabled = False
        f2424Oex.Value = na
    f2424PA.Value = "X"
    GoTo JPBBupdater
End If
If f2424Ptype = "PC" Then
    f2424Cancel = True
    f2424_typeO.Value = False
    f2424_typePA.Value = False
    f2424_typePK.Value = False
    f2424_typePQ.Value = False
    f2424_typeOex.Enabled = False
        f2424Oex.Value = na
    f2424PC.Value = "X"
    GoTo JPBBupdater
End If
If f2424Ptype = "PK" Then
    f2424Cancel = True
    f2424_typeO.Value = False
    f2424_typePA.Value = False
    f2424_typePC.Value = False
    f2424_typePQ.Value = False
    f2424_typeOex.Enabled = False
        f2424Oex.Value = na
    f2424PK.Value = "X"
    GoTo JPBBupdater
End If
If f2424Ptype = "PQ" Then
    f2424Cancel = True
    f2424_typeO.Value = False
    f2424_typePA.Value = False
    f2424_typePC.Value = False
    f2424_typePK.Value = False
    f2424_typeOex.Enabled = False
        f2424Oex.Value = na
    f2424PQ.Value = "X"
    GoTo JPBBupdater
End If

JPBBupdater:
If f2424_typePQ.Value Then '  PQ - 9 DIGIT
    f2424JPBBamount.Value = "'" & Format(f2424_amount.Value * 100, "000000000")
ElseIf Not f2424_typeO.Value Then ' EVERYTHING ELSE - 7 DIGIT
    f2424JPBBamount.Value = "'" & Format(f2424_amount.Value * 100, "0000000")
Else
    f2424JPBBamount.Value = "INPUT WITH CAUTION"
End If

End Sub

Private Sub f2424c_DelWarn_Click() ' 2424 config for warning upon removing Content on 2424

If f2424c_DelWarn Then ' baseline config update trigger
    f2424cDelw.Value = True
    f2424c_DelWarn.Caption = "Mute"
Else
    f2424cDelw.Value = False
    f2424c_DelWarn.Caption = "Warn"
End If

End Sub

Private Sub f2424c_MNA_Click() ' 2424 config should we put ADMIN ACTION

If f2424c_MNA Then
    f2424cMNA.Value = True
    f2424c_MNA.Caption = "N/A"
    f2424.Range("B23").Value = "M.N.A."
    f2424_admin.BackColor = &H855988
Else
    f2424cMNA.Value = False
    f2424c_MNA.Caption = "Available"
    f2424.Range("B23").Value = ""
    f2424_admin.BackColor = &H8000000F
End If

End Sub

Private Sub f2424c_Prev_Click() ' 2424 config Inherit previous

If f2424c_Prev Then
    f2424cPrev.Value = True
    f2424c_Prev.Caption = "Inherit"
Else
    f2424cPrev.Value = False
    f2424c_Prev.Caption = "Discard"
End If

End Sub

Private Sub f2424c_SSNlink_Click() ' 2424 alievate thru ALPHA

' Require function of dictionary write, dictionary append and dictionary look
If f2424c_SSNlink.Value Then f2424c_SSNlink.Caption = "Enabled"
If Not f2424c_SSNlink.Value Then f2424c_SSNlink.Caption = "Disabled"
f2424cSSNlink.Value = f2424c_SSNlink.Value

End Sub

Private Sub f2424c_StartNew_Click() ' 2424 config start a new page after print

If f2424c_StartNew Then
    f2424cStartNew.Value = True
    f2424c_StartNew.Caption = "Enabled"
Else
    f2424cStartNew.Value = False
    f2424c_StartNew.Caption = "Disabled"
End If

End Sub

Private Sub formLib_Change() ' GLOBAL PAGE SWITCHER #PENDING

sheetDisplayInitialHide

Select Case formLib.Value
Case 2
    F110p1.Visible = xlSheetVisible
    F110p2.Visible = xlSheetVisible
Case 3
    f2424.Visible = xlSheetVisible
Case 4
    f117.Visible = xlSheetVisible
Case Else
    ' do nothing
End Select

End Sub

Private Sub Gconfig_delWarn_Click() ' GLOBAL CONFIG WARN BEFORE DELETE
If Gconfig_delWarn Then
    config.Range("F7").Value = True
    Gconfig_delWarn.Caption = "ENABLED"
Else
    config.Range("F7").Value = False
    Gconfig_delWarn.Caption = "DISABLED"
End If

Gconfig_DelOverride ' adjust override
globalSave
End Sub
Sub Gconfig_DelOverride() ' global config to override localized warning
If Gconfig_delWarn Then
    f2424c_DelWarn.Enabled = False
    f110c_delWarn.Enabled = False
Else
    f2424c_DelWarn.Enabled = True
    f110c_delWarn.Enabled = True
End If
End Sub

Private Sub Gconfig_saveOptn_Click() ' global config save way in some way
If Gconfig_saveOptn Then ' ALLOW FIXED PATH
    config.Range("F5").Value = True
    saveOptn = True
    Gconfig_saveOptn.Caption = "CONSTANT"
    Gconfig_saveTo.Enabled = True
    Gconfig_saveToAssign.Enabled = True
    Gconfig_saveToRemove.Enabled = True
Else ' OR NOT
    config.Range("F5").Value = False
    saveOptn = False
    Gconfig_saveOptn.Caption = "VARIED"
    Gconfig_saveTo.Enabled = False
    Gconfig_saveToAssign.Enabled = False
    Gconfig_saveToRemove.Enabled = False
End If
globalSave
End Sub

Private Sub Gconfig_saveToAssign_Click() ' global config assign path
Dim tempPath As String, tempFinder As FileDialog

Set tempFinder = Application.FileDialog(msoFileDialogFolderPicker)
    With tempFinder
        .Title = "Global forms will be exported to here"
        .ButtonName = "Assign"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo datawrite
        tempPath = .SelectedItems(1)
    End With
datawrite:
Set tempFinder = Nothing
config.Range("F6").Value = tempPath
loadGconfig
globalSave

End Sub

Private Sub Gconfig_saveToRemove_Click() ' Global config remove path
Dim resB As String
    resB = MsgBox("Remove Current Path?", vbYesNo, "Distiller Path Deletion")
If resB = vbNo Then Exit Sub
config.Range("F6").Value = ""
loadGconfig
globalSave

End Sub

Private Sub Gconfig_SSNlookup_Click() ' GCONFIG for GLOBAL FORCED SSN LOOK UP

' Require function of dictionary write, dictionary append and dictionary look

End Sub

Private Sub hidePanel_Click()
    utilityForms.Hide
    Application.StatusBar = False
    trackerAPI.Show
    globalSave
    sheetDisplayInitialHide ' hide all
End Sub


Private Sub SITWdrop_Change() ' 110 State list
    stID.Value = SITWdrop.Value
    SITWamu.Caption = "Tax at " & Format(stPer.Value * 100, "0.00") & "%"
    If SITWdrop <> "" Then
        f110_sitw.Locked = False
    Else
        f110_sitw.Locked = True
    End If
End Sub


Private Sub UserForm_Initialize()
' FIGURE WORKSHEETS
    Set F110p1 = Worksheets("DEBT.A") ' 110
    Set F110p2 = Worksheets("DEBT.B") ' 110
    Set f2424 = Worksheets("ADV.PAY") ' 2424
    Set f117 = Worksheets("CAGE.PAY") ' 117
    Set config = Worksheets("SENSEI.CONFIG")
' FIGURE DATA CONTAINER
    Set Data = Worksheets("SENSEI.DATA")
    Set Tab1 = Data.Range("D19:D70")
    Set formVer = config.Range("F4") ' version Form
    Set Sver = config.Range("D4") ' sensei Verison
    distVer = config.Range("F2").Value ' flip version
    distVerType = config.Range("F3").Value ' flip version type
    FormVersion.Caption = formVer.Value & " on " & Sver.Value
    saveTo = config.Range("F6").Value
    saveOptn = config.Range("F5").Value

    initialize110 ' initialize 110
    initialize2424 ' initialize 2424
    initialize117 ' initialize 117
    
    loadGconfig ' GLOBAL CONFIG LOADER
    sheetDisplayInitialHide ' hide all forms
End Sub

Sub sheetDisplayInitialHide() ' Hide display

F110p1.Visible = xlSheetHidden
F110p2.Visible = xlSheetHidden
f2424.Visible = xlSheetHidden
f117.Visible = xlSheetHidden

End Sub

Sub initialize110() ' initialize 110 content
' FIGURE ERASE RANGE
    With F110p1 ' 110 setup pAGE 1
        Set paidRateP1 = .Range("H5:H17") '$ PAID
        Set dueRateP1 = .Range("K5:K17") '$ DUE
        Set periodStartP1 = .Range("A5:A17")
        Set periodEndP1 = .Range("C5:C17")
        Set itemP1 = .Range("D5:D17")
        Set typeP1 = .Range("E5:E17")
        Set gradeP1 = .Range("J5:J17")
        Set dueUSP1 = .Range("N5:N25") '$ TO US
        Set taxTotal = .Range("J20:J23") ' TAXABLE $ LOCATION
        Set taxIso = .Range("L20") ' isolated tax $
        Set stID = .Range("E23") ' STATE ID FOR SITW
        Set stPer = .Range("O23") ' STATE TAX % calc point
        Set debtTotal = .Range("M25") ' TOTAL DEBT
        Set name110 = .Range("H2") ' NAME
        Set ssn110 = .Range("M2") ' SSN
    End With
    With F110p2
        Set paidRateP2 = .Range("H5:H26") ' $ PAID
        Set dueRateP2 = .Range("K5:K26") ' $ DUE
        Set periodStartP2 = .Range("A5:A26")
        Set periodEndP2 = .Range("C5:C26")
        Set itemP2 = .Range("D5:D26")
        Set typeP2 = .Range("E5:E26")
        Set gradeP2 = .Range("J5:J26")
        Set dueUSP2 = .Range("N5:N26") '$ TO US
    End With
' FIGURE SITW CHART
    writeSITWlist ' 110
' Write update 0 for Display
    DispPage = 1 ' 110 display page #
    
    Disp_all = config.Range("F33").Value ' temporary 110 to display all
        f110c_dispAll = Disp_all ' update in settings
    Disp_Link = config.Range("F34").Value ' temporary 110 to follow the display
        f110c_dispFollow = Disp_Link ' update from settings
        f110c_keepMbr = config.Range("F35").Value
        f110c_inherit = config.Range("F36").Value
    Set f110delW = config.Range("F37") ' INDIVIDUAL WARNING
        f110c_delWarn = f110delW.Value
    
    With config
        Set f110cSSNLink = .Range("F39")
    End With
    
    f110Row = f110_RowCt.Value ' 110 - Make row Number valid
    f110rowDispUpdate ' 110 - UPDATE FIELD
    f110_name.Value = name110.Value ' 110 LOAD NAME
    f110_ssn = Format(ssn110.Value, "000000000") ' 110 LOAD SSN
    update110Display ' 110

End Sub

Sub initialize2424() ' 2424 intitlize all config
f2424Cancel = True ' Disable Updates while initializing

With f2424 ' ASSIGN TO RANGE
    Set f2424amount = .Range("B9") ' AMOUNT
    Set f2424amountF = .Range("F9") ' amount in words
    Set f2424PA = .Range("C10")
    Set f2424PC = .Range("G10")
    Set f2424PK = .Range("I10")
    Set f2424PQ = .Range("C11")
    Set f2424O = .Range("C12")
    Set f2424Oex = .Range("G12")
    Set f2424Cat = .Range("B14") ' additional pay info
    Set f2424SSN = .Range("G14")
    Set f2424Name = .Range("B16")
    Set f2424Rank = .Range("J16")
    Set f2424expl = .Shapes("f2424_expl").TextFrame.Characters ' computations
    Set f2424JPBBssn = .Range("J28") ' JPBB DATA, TO 3 BELOW
    Set f2424JPBBname = .Range("J30")
    Set f2424JPBBamount = .Range("J32")
End With

With config ' USED FOR CONFIG IF
    Set f2424cDelw = .Range("F64")
    Set f2424cStartNew = .Range("F65")
    Set f2424cMNA = .Range("F66")
    Set f2424cPrev = .Range("F67") ' is it enabled? the prior inherit
    Set f2424cPrevType = .Range("F68") ' prior inherit specific type
    Set f2424cSSNlink = .Range("F69") ' Auto SSN, suspended due to lack of The list
End With

' Load Existing Value
' f2424_admin.BackColor = &H8000000F ' hold on to this till it was worked
f2424_amount.Value = f2424amount.Value
f2424_amountFig.Caption = f2424amountF.Value
If f2424PA.Value = "X" Or (f2424cPrevType.Value = "PA" And f2424cPrev.Value = True) Then f2424_typePA.Value = True
If f2424PC.Value = "X" Or (f2424cPrevType.Value = "PC" And f2424cPrev.Value = True) Then f2424_typePC.Value = True
If f2424PK.Value = "X" Or (f2424cPrevType.Value = "PK" And f2424cPrev.Value = True) Then f2424_typePK.Value = True
If f2424PQ.Value = "X" Or (f2424cPrevType.Value = "PQ" And f2424cPrev.Value = True) Then f2424_typePQ.Value = True
If f2424O.Value = "X" Or (f2424cPrevType.Value = "PO" And f2424cPrev.Value = True) Then f2424_typeO.Value = True
f2424_typeOex.Value = f2424Oex.Value
f2424_mbrSSN.Value = f2424SSN.Value
f2424_mbrName.Value = f2424Name.Value
f2424_mbrRank.Value = f2424Rank.Value
f2424_explain.Value = f2424expl.Text
f2424_typeCat.Value = f2424Cat.Value

' Adjust local setting toggles
f2424c_DelWarn.Value = f2424cDelw.Value
f2424c_StartNew.Value = f2424cStartNew.Value
f2424c_MNA.Value = f2424cMNA.Value
f2424c_Prev.Value = f2424cPrev.Value
f2424c_SSNlink.Value = f2424cSSNlink

' Input field update
On Error GoTo reducedExit ' to cut it off
If f2424_typePQ.Value Then '  PQ - 9 DIGIT
    f2424JPBBamount.Value = "'" & Format(f2424_amount.Value * 100, "000000000")
ElseIf Not f2424_typeO.Value Then ' EVERYTHING ELSE - 7 DIGIT
    If f2424_amount.Value <> "" Then
        f2424JPBBamount.Value = "'" & Format(f2424_amount.Value * 100, "0000000")
    Else
        f2424JPBBamount.Value = "'" & Format(0 * 100, "0000000")
    End If
Else
reducedExit:
    f2424JPBBamount.Value = "INPUT WITH CAUTION"
End If
f2424JPBBssn.Value = f2424_mbrSSN.Value
f2424JPBBname.Value = Left(Replace(f2424_mbrName.Text, ",", ""), 5) ' replace the comma regardless


f2424Cancel = False ' Re-engage Updater
End Sub
Sub initialize117() ' initalize 117 and its associated config
' Config loading range F95,96,97,98
' Add toggle ability to Partial/SepRet
' Add toggle lock to PAID BY section
' Add Toggle for Auto/Manual Payment input

' Addigning FY for LOA
    thisFYid = Format(Now(), "YYYY") & "-10-01" ' separator of FY
    If DateValue(Format(Now(), "YYYY-MM-DD")) < DateValue(thisFYid) Then
        thisFY = Right(Format(Now(), "YYYY"), 1)
    Else
        thisFY = Str(Int(Right(Format(Now(), "YYYY"), 1)) + 1)
    End If
    If thisFY = "10" Then thisFY = "0" ' End of assigning FY id
    f117Nuking = False
    
' Assign ranges [Named]
With f117
    Set f117vchNo = .Range("U2")
    Set f117ADSN = .Range("C40")
    Set f117payPurpose = .Range("B7")
    Set f117Name = .Range("J10")
    Set f117Rank = .Range("G10")
    Set f117SSN = .Range("G11")
    Set f117Routing = .Range("G12")
    Set f117Account = .Range("G13")
    Set f117Ptype = .Range("C34")
    Set f117Pdate = .Range("V13") ' PREPARE DATE
    Set f117Amu = .Range("V27")
    Set f117EO = .Range("C46") ' Sheet exclusive, depends on Rank
    Set f117Service = .Range("C38") ' A-AF B-ARMY C-USN D-MARINE G-COAST GUARD M-US STATE DEPT
    Set f117Lctn = .Range("U9") ' Institution address
    Set f117DO = .Range("B19")
    Set f117DSSN = .Range("K19")
    Set f117FSO = .Range("I5")
    Set f117LOA = .Range("C27") ' AUTO POPULATE DISABLE EDIT BY DEFAULT
    Set f117PdDate = .Range("V23")
    Set f117prepBy = .Range("B56")
    Set f117receiveBy = .Range("J56")
    Set f117accType = .Range("G14")
End With

' initialize Rank List
With f117_Rank
    .Clear
    .AddItem "E-1/AB"
    .AddItem "E-2/AMN"
    .AddItem "E-3/A1C"
    .AddItem "E-4/SRA"
    .AddItem "E-5/SSGT"
    .AddItem "E-6/TSGT"
    .AddItem "E-7/MSGT"
    .AddItem "E-8/SMSGT"
    .AddItem "E-9/CMSGT"
    .AddItem "O-1/2 LT"
    .AddItem "O-2/1 LT"
    .AddItem "O-3/CAPT"
    .AddItem "O-4/MAJ"
    .AddItem "O-5/LTC"
    .AddItem "O-6/COL"
    .AddItem "O-7/B.GEN"
    .AddItem "O-8/M.GEN"
    .AddItem "O-9/L.GEN"
    .AddItem "O-10/GEN"
End With

' initialize Service List
With f117_Service
    .Clear
    .AddItem "A - USAF"
    .AddItem "B - USA"
    .AddItem "C - USN"
    .AddItem "D - USMC"
    .AddItem "G - USCG"
    .AddItem "M - STATE"
End With

' populate config and update Config
With config
    Set f117cDelWarn = .Range("F95")
    Set f117cStartNew = .Range("F96")
    Set f117cDMO = .Range("F97")
    Set f117cSSNLink = .Range("F98")
    Set f117cAutoPrep = .Range("F99")
    Set f117cAutoLOA = .Range("F100")
    Set f117cDelTotal = .Range("F101")
End With
    f117c_DelWarn.Value = f117cDelWarn.Value
    f117c_StartNew.Value = f117cStartNew.Value
    f117c_DMO.Value = f117cDMO.Value
    f117c_SSNLink.Value = f117cSSNLink.Value
    f117c_AutoPrep.Value = f117cAutoPrep.Value
    f117c_AutoLOA.Value = f117cAutoLOA.Value
    f117c_DelTotal.Value = f117cDelTotal.Value
    f117PKPQ = False ' no flag initially

initialize117a ' to populate
update117display ' to change display

End Sub
Sub initialize117a() ' Second part, take care of factoring the value amd populate them
    
    f117_vchNo.Value = f117vchNo.Value
    f117_payPurpose.Value = f117payPurpose.Value
    f117_ADSN.Value = f117ADSN.Value
    f117_Name.Value = f117Name.Value
    f117_Rank.Value = f117Rank.Value
    Select Case f117Service.Value ' LOAD SERVICE LIST DISPLAY(A,B,C,D,G,M)
    Case "A"
        f117_Service.Value = "A - USAF"
    Case "B"
        f117_Service.Value = "B - USA"
    Case "C"
        f117_Service.Value = "C - USN"
    Case "D"
        f117_Service.Value = "D - USMC"
    Case "G"
        f117_Service.Value = "G - USCG"
    Case "M"
        f117_Service.Value = "M - STATE"
    Case Else
        f117_Service.Value = ""
    End Select
    f117_SSN.Value = f117SSN.Value
    f117_Routing.Value = f117Routing.Value
    f117_Account.Value = f117Account.Value
    If f117Ptype.Value = "PK03" Then f117_PK.Value = True
    If f117Ptype.Value = "PQ03" Then F117_PQ.Value = True
    f117_Pdate.Value = f117Pdate.Value
    f117_Amu.Value = f117Amu.Value
    If f117Amu.Value = "$" Then f117_Amu.Value = 0 ' invert this
    f117_Lctn.Value = f117Lctn.Value
    f117_DO.Value = f117DO.Value
    f117_DSSN.Value = f117DSSN.Value
    f117_FSO.Value = f117FSO.Value
    f117_LOA.Value = f117LOA.Value
    f117_PdDate.Value = f117PdDate.Value
    f117_prepBy.Value = f117prepBy.Value
    f117_receiveBy.Value = f117receiveBy.Value
    If f117accType.Value = "" And Not f117Nuking Then f117_accType.Value = False
    
End Sub
Sub update117display() ' the common interrelated update when flipping toggles in config

If f117c_AutoLOA.Value Then ' if adapative, prep for auto calculate
    f117_LOA.Enabled = False
    f117pb5.BackColor = &H938A94
    ' default
    f117_LOA.Value = "57" & thisFY & " 32" & thisFY & " P###" & " 525725"
    ' to write with rank
    If Left(f117_Rank.Value, 1) = "E" Then f117_LOA.Value = Replace(f117_LOA.Value, "###", "530")
    If Left(f117_Rank.Value, 1) = "O" Then f117_LOA.Value = Replace(f117_LOA.Value, "###", "510")
Else ' if not adapative, unlock only
    f117_LOA.Enabled = True
    f117pb5.BackColor = &H855988
End If

If f117c_AutoPrep.Value Then ' if adapative, automatically place today's date in it
    f117_Pdate.Enabled = False
    f117pt1.BackColor = &H938A94
    f117_Pdate.Value = Format(Now(), "YYYY-MM-DD")
Else ' if not enable it for yyyy-mm-dd input
    f117_Pdate.Enabled = True
    f117pt1.BackColor = &H855988
End If

' PaidBy Enabling sequence, appearance deal
If Not f117_editPaidBy.Value Then ' to disable edit
    fm02.Enabled = False
    f117pb1.BackColor = &H938A94
    f117pb2.BackColor = &H938A94
    f117pb3.BackColor = &H938A94
    f117pb4.BackColor = &H938A94
    If Not f117c_AutoLOA.Value Then f117pb5.BackColor = &H938A94
Else ' to enable edit
    fm02.Enabled = True
    f117pb1.BackColor = &H855988
    f117pb2.BackColor = &H855988
    f117pb3.BackColor = &H855988
    f117pb4.BackColor = &H855988
    If Not f117c_AutoLOA.Value Then f117pb5.BackColor = &H855988
End If

If f117accType.Value = "Checking" Then
    f117_accType.Value = False
    f117_accType.Caption = "C"
ElseIf f117accType.Value = "Saving" Then
    f117_accType.Value = True
    f117_accType.Caption = "S"
Else ' DO NOTHING IF ACCOUNT IS UNDEFINED
End If

End Sub

Sub loadGconfig() ' GLOBAL Config loader

Gconfig_saveOptn.Value = config.Range("F5").Value
Gconfig_saveTo.Value = config.Range("F6").Value
Gconfig_delWarn.Value = config.Range("F7").Value ' WARN DELETE
Gconfig_SSNlookup.Value = config.Range("F8").Value ' GLOBAL FORCE SSN LOOKUP
Gconfig_DelOverride ' Must handle enabling

End Sub
Sub writeSITWlist() ' 110 FX_pre, loads SITW list FOR 110
Dim tCell As Range ' for cell counting within this

For Each tCell In Tab1 ' Loop through the table. append
    If tCell.Value <> "" Then
        With SITWdrop
            .AddItem tCell.Value
        End With
    End If
Next tCell

End Sub
Sub updateCompDebt() ' 110 FX for Summing Tax
    amountSum.Value = Format(-Round(debtTotal.Value, 2), "0.00")
End Sub
Sub update110Display() ' 110 update display
' 13 Lines on Page 1, 22 Lines on Page 2
Dim defValue As String  ' Default loader
    defValue = "Entry  Period         Item     Type   Mth-Dys  P.rate    P.amount  Grade   D.rate    D.amount  Diff.       Diff.US "
               'E-00:  220000-229999  1234567  12345  000-000  99999.00  3333.30_  E9-99Y  99999.00  3333.30_  +99999.00    99999.00
Dim apArray As String, apLoopRw As Long, apActual As Long, apLoopRwP As Long  ' Appointed array list that amends disp110 as it goes
 '  Row data           Loop Number       Matching Row      page limiter
Dim sP As String  ' Space bar
    sP = "  "
Dim cA As String, cC As String, cD As String, cE As String, cF As String, cG As String, cH As String, cI As String, cJ As String, cK As String, cL As String, cM As String, cN As String
    If Not Disp_all Then ' per page display
        Select Case DispPage
        Case 1
            apLoopRwP = 1
        Case 2
            apLoopRwP = 11
        Case 3
            apLoopRwP = 21
        Case 4
            apLoopRwP = 31
        End Select
    Else
        apLoopRwP = 1
    End If
    disp110.Value = "" ' CURRENTLY JUST WIPE IT, TILL WE ADDED DISPLAY ALL THEN WE STOP.
For apLoopRw = apLoopRwP To 35
    ' page fx inserted here \/
    If Not Disp_all Then ' display by page
        Select Case DispPage
        Case 1
            If apLoopRw = 11 Then Exit For
        Case 2
            If apLoopRw = 21 Then Exit For
        Case 3
            If apLoopRw = 31 Then Exit For
        Case 4
            If apLoopRw = 36 Then Exit For
        End Select
    Else
    End If
    If (apLoopRw - 1) Mod 10 = 0 Then  ' test write reminder every 10 rows
        If apArray = "" And Not Disp_all Then
            apArray = defValue
        Else
            apArray = apArray & vbNewLine & defValue
        End If
    End If
    If apLoopRw < 14 Then ' Page writing Logics
        With F110p1  ' sample experiment block for Page 1 writing
            apActual = apLoopRw + 4
            cA = Format(.Range("A" & apActual), "YYMMDD") & "-"
                If .Range("A" & apActual) = "" Or .Range("A" & apActual) = 0 Then cA = "       "  ' PERIOD START
            cC = Format(.Range("C" & apActual), "YYMMDD")
                If .Range("C" & apActual) = "" Or .Range("C" & apActual) = 0 Then cC = "      "  ' PERIOD END
            cD = Format(.Range("D" & apActual), "@@@@@@@")
                If .Range("D" & apActual) = "" Or .Range("D" & apActual) = 0 Then cD = "       "  ' ITEM
            cE = Format(.Range("E" & apActual), "@@@@@")
                If .Range("E" & apActual) = "" Or .Range("E" & apActual) = 0 Then cE = "     "  ' TYPE
            cF = Format(.Range("F" & apActual), "000")
                If .Range("F" & apActual) = "" Or .Range("F" & apActual) = 0 Then cF = "   "  ' MONTH
            cG = Format(.Range("G" & apActual), "000")
                If .Range("G" & apActual) = "" Or .Range("G" & apActual) = 0 Then cG = "   "  ' DAY
            cH = Format(.Range("H" & apActual), "00000.00")
                If .Range("H" & apActual) = "" Or .Range("H" & apActual) = 0 Then cH = "        "  ' PAID RATE
            cI = Format(.Range("I" & apActual), "00000.00")
                If .Range("I" & apActual) = "" Or .Range("I" & apActual) = 0 Then cI = "        "  ' PAID TOT.
            cJ = Format(.Range("J" & apActual), "@@@@@@")
                If .Range("J" & apActual) = "" Or .Range("J" & apActual) = 0 Then cJ = "      "  ' GRADE
            cK = Format(.Range("K" & apActual), "00000.00")
                If .Range("K" & apActual) = "" Or .Range("K" & apActual) = 0 Then cK = "        "  ' DUE RATE
            cL = Format(.Range("L" & apActual), "00000.00")
                If .Range("L" & apActual) = "" Or .Range("L" & apActual) = 0 Then cL = "        "  ' DUE TOT.
            cM = Format(.Range("M" & apActual), "+00000.00")
                If .Range("I" & apActual) = 0 And .Range("M" & apActual) = 0 Then cM = "         "  ' DUE MBR
                If .Range("M" & apActual) < 0 Then cM = Format(.Range("M" & apActual), "00000.00") ' lesser than 0
            cN = Format(.Range("N" & apActual), "+00000.00")
                If .Range("N" & apActual) = "" And .Range("N" & apActual) = 0 Then cN = "        "  ' DUE US
                If .Range("N" & apActual) < 0 Then cN = Format(.Range("N" & apActual), "00000.00") ' lesser than 0
                
            apArray = apArray & vbNewLine & "E-" & Format(apLoopRw, "00") & ":" & sP & cA & cC & sP & cD & sP & cE & sP & cF & " " & cG & sP & cH & sP & cI & sP & cJ & sP & cK & sP & cL & sP & cM & sP & cN
        End With
    Else
        With F110p2  ' sample experiment block for Page 2
            apActual = apLoopRw - 9
            cA = Format(.Range("A" & apActual), "YYMMDD") & "-"
                If .Range("A" & apActual) = "" Or .Range("A" & apActual) = 0 Then cA = "       "  ' PERIOD START
            cC = Format(.Range("C" & apActual), "YYMMDD")
                If .Range("C" & apActual) = "" Or .Range("C" & apActual) = 0 Then cC = "      "  ' PERIOD END
            cD = Format(.Range("D" & apActual), "@@@@@@@")
                If .Range("D" & apActual) = "" Or .Range("D" & apActual) = 0 Then cD = "       "  ' ITEM
            cE = Format(.Range("E" & apActual), "@@@@@")
                If .Range("E" & apActual) = "" Or .Range("E" & apActual) = 0 Then cE = "     "  ' TYPE
            cF = Format(.Range("F" & apActual), "000")
                If .Range("F" & apActual) = "" Or .Range("F" & apActual) = 0 Then cF = "   "  ' MONTH
            cG = Format(.Range("G" & apActual), "000")
                If .Range("G" & apActual) = "" Or .Range("G" & apActual) = 0 Then cG = "   "  ' DAY
            cH = Format(.Range("H" & apActual), "00000.00")
                If .Range("H" & apActual) = "" Or .Range("H" & apActual) = 0 Then cH = "        "  ' PAID RATE
            cI = Format(.Range("I" & apActual), "00000.00")
                If .Range("I" & apActual) = "" Or .Range("I" & apActual) = 0 Then cI = "        "  ' PAID TOT.
            cJ = Format(.Range("J" & apActual), "@@@@@@")
                If .Range("J" & apActual) = "" Or .Range("J" & apActual) = 0 Then cJ = "      "  ' GRADE
            cK = Format(.Range("K" & apActual), "00000.00")
                If .Range("K" & apActual) = "" Or .Range("K" & apActual) = 0 Then cK = "        "  ' DUE RATE
            cL = Format(.Range("L" & apActual), "00000.00")
                If .Range("L" & apActual) = "" Or .Range("L" & apActual) = 0 Then cL = "        "  ' DUE TOT.
            cM = Format(.Range("M" & apActual), "600000.00")
                If .Range("I" & apActual) = 0 And .Range("M" & apActual) = 0 Then cM = "         "  ' DUE MBR
                If .Range("M" & apActual) < 0 Then cM = Format(.Range("M" & apActual), "0000.00") ' lesser than 0
            cN = Format(.Range("N" & apActual), "+00000.00")
                If .Range("N" & apActual) = "" And .Range("N" & apActual) = 0 Then cN = "        "  ' DUE US
                If .Range("N" & apActual) < 0 Then cN = Format(.Range("N" & apActual), "00000.00") ' lesser than 0

                
            apArray = apArray & vbNewLine & "E-" & Format(apLoopRw, "00") & ":" & sP & cA & cC & sP & cD & sP & cE & sP & cF & " " & cG & sP & cH & sP & cI & sP & cJ & sP & cK & sP & cL & sP & cM & sP & cN
        End With
    End If
Next apLoopRw

' INITIALIZE
    disp110.Value = apArray ' was defValue & apArray
    updateCompDebt ' MIGHT AS WELL UPDATE THIS
End Sub

