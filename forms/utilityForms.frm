VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} utilityForms 
   Caption         =   "Sensei Debt Computation"
   ClientHeight    =   7695
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
' =========== GLOBAL VARIABLE ============
Public config As Worksheet ' config
Public formVer As Range, Sver As Range ' Form version and Sensei Version
Public saveTo As String, saveOptn As Boolean ' the location and is prompt allowed?
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
' ==================================================================================

Private Sub amountList_Change() ' 110 CHANGE LOGIC REQUIRED, STR IS PREFERRED
On Error GoTo handler
    If CDbl(amountList.Value) - CDbl(amountSum.Value) <> 0 Or amountList.Value = "" Or amountSum.Value = "" Then
        amountMatch.Value = "NG"
    Else
        amountMatch.Value = "OK"
    End If
    updateCompDebt ' UPDATE MATCH
    Exit Sub
handler:
    amountMatch.Value = "--"
    updateCompDebt ' UPDATE MATCH
End Sub


Private Sub f110_delall_Click() ' 110 delete all items
Dim ResQ As String, na As String, npy As String
    na = vbNullString
    npy = 0
    
If Gconfig_delWarn Then
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
If Gconfig_delWarn Then
    Dim ResQ As String
        ResQ = MsgBox("Delete this Entry?", vbYesNo, "Form Distiller")
    If ResQ = vbNo Then Exit Sub
End If
If f110_PageCt.Value = "P.1" Then ' PAGE ONE
    With F110p1
        f110_strDate.Value = ""
        f110_endDate.Value = ""
        f110_itemName.Value = ""
        f110_itemType.Value = ""
        f110_itemGrade.Value = ""
        f110_paidRate.Value = 0
        f110_dueRate.Value = 0
        f110_dueUS.Value = ""
        f110_dueClaimant.Caption = Format(.Range("M" & f110Row).Value, "$0.00")
    End With
ElseIf f110_PageCt.Value = "P.2" Then ' PAGE TWO
    With F110p2
        f110_strDate.Value = ""
        f110_endDate.Value = ""
        f110_itemName.Value = ""
        f110_itemType.Value = ""
        f110_itemGrade.Value = ""
        f110_paidRate.Value = 0
        f110_dueRate.Value = 0
        f110_dueUS.Value = ""
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

Private Sub f110_export_Click() ' 110 export function
Dim saveToPrompt

On Error GoTo handleIt
If Not saveOptn Or saveTo = "" Then ' ALWAYS PROMPT IF DISABLED PATHWAY OR SAVETO IS BLANK
    Set saveToPrompt = Application.FileDialog(msoFileDialogFolderPicker)
    With saveToPrompt
        .Title = "Sending Form 110 to here..."
        .ButtonName = "Save"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo exportForm
        saveTo = .SelectedItems(1)
    End With
End If


exportForm:
Set saveToPrompt = Nothing ' UNLOAD OBJECT >>Does not save any way ?
    F110p1.ExportAsFixedFormat xlTypePDF, _
        Filename:=saveTo & "\COMP." & Left(f110_name.Value, 5) & ".01"
    If F110p1.Range("M18").Value <> 0 Then F110p2.ExportAsFixedFormat xlTypePDF, _
        Filename:=saveTo & "\COMP." & Left(f110_name.Value, 5) & ".02"
    Exit Sub

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
        f110_strDate.Value = .Range("A" & f110Row).Value
        f110_endDate.Value = .Range("C" & f110Row).Value
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
        f110_strDate.Value = .Range("A" & f110Row).Value
        f110_endDate.Value = .Range("C" & f110Row).Value
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

Private Sub Gconfig_delWarn_Click() ' GLOBAL CONFIG WARN BEFORE DELETE
If Gconfig_delWarn Then
    config.Range("F7").Value = True
    Gconfig_delWarn.Caption = "ENABLED"
Else
    config.Range("F7").Value = False
    Gconfig_delWarn.Caption = "DISABLED"
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

End Sub

Private Sub Gconfig_saveToRemove_Click() ' Global config remove path
Dim resB As String
    resB = MsgBox("Remove Current Path?", vbYesNo, "Distiller Path Deletion")
If resB = vbNo Then Exit Sub
config.Range("F6").calue = ""
loadGconfig

End Sub

Private Sub hidePanel_Click()
    utilityForms.Hide
    trackerAPI.Show
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
    
    loadGconfig ' GLOBAL CONFIG LOADER
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
        
    update110Display ' 110
    f110Row = f110_RowCt.Value ' 110 - Make row Number valid
    f110rowDispUpdate ' 110 - UPDATE FIELD
    f110_name.Value = name110.Value ' 110 LOAD NAME
    f110_ssn = ssn110.Value ' 110 LOAD SSN

End Sub

Sub loadGconfig() ' GLOBAL Config loader

Gconfig_saveOptn.Value = config.Range("F5").Value
Gconfig_saveTo.Value = config.Range("F6").Value
Gconfig_delWarn.Value = config.Range("F7").Value ' WARN DELETE

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
Sub update110Display()
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
