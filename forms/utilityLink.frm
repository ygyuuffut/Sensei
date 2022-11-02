VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} utilityLink 
   Caption         =   "SENSEI LINK - A Node within Network"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "utilityLink.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "utilityLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 20221102
Public config As Worksheet ' Core Config data
Public LcCell As Range, LcTrim As Range, Vrrr As Range, VrHere As Range, Vrelease As Range
Public Cell114 As Range, Trim114 As Range ' 114 Data Cell reference
Public T114 As Range, TR2R As Range ' 114 TRIMMED NAME AND RRR TRIMMED NAME
Public SenseiVer As Range ' Sensei's version
Public FileLc As String ' Default Location carrier
Public R2Rx As String ' R2R Label Version

'todo :
' APPOINT RRR
' RUN IT
'
'
'
' taskings:
' VALIDATE IF PEOPLE ARE WORKING REJECT
' REMOVE PROCESS WITH QA(BONUS)
' POWER BI UNDER (DASHBOARD AND EDITING):
' DRAG REJECT INFORMATION AND VALIDATION PROCESS WITHIN EXCEL
'
'

Private Sub AddressFinder_Click()
'Dim F_Locater As FileDialog
'Set F_Locater = Application.FileDialog(msoFileDialogFilePicker)
'
'With F_Locater ' sample select
'    .AllowMultiSelect = False
'    .Filters.Add "3R Engine", "*.xlsm", 1
'    .Title = "Linking to RRR Engine"
'    .ButtonName = "Link"
'    .Show
'
'    FileLc = .SelectedItems.Item(1) ' Wrote address to Variable
'End With
'
'LcCell.Value = FileLc
'uRRR_Address.Value = LcCell.Value ' inflict updates
CommonLocationFinder
End Sub

Private Sub hidePanel_Click()
    Unload utilityLink
    trackerAPI.Show
End Sub

Private Sub linkGuide_Change()
Select Case linkGuide.Value
Case 0 ' cover
    AddressFinder.Visible = False
    Trimmer.Visible = False
Case 1 ' the RRR page
    AddressFinder.Visible = True
    Trimmer.Visible = True
    With OperateFiles
        .Visible = True
        .Caption = "Operate"
        .ControlTipText = "Execute Reject Report"
    End With
Case 2 ' the 114 page
    AddressFinder.Visible = True
    Trimmer.Visible = True
    With OperateFiles
        .Visible = True
        .Caption = "Launch"
        .ControlTipText = "Launch 114 by RTC team"
    End With
End Select
If linkGuide.Value = 0 Then OperateFiles.Visible = False
End Sub



Private Sub OperateFiles_Click() ' test carry
On Error Resume Next ' ingnore if workbook is cloing an closed instance, that is not a bug
    If uRRR_Address <> vbNullString Then
        Select Case linkGuide.Value
        Case 1
            Workbooks.Open (uRRR_Address)
            Application.Run Trimmed & "!Dupe_Main" ' adjust this if the destination module has been changed
            Workbooks("SENSEI - dev.xlsm").Sheets("CSP.TR").Range("S1").Value = Workbooks(Trimmed).Sheets("DRIVE").Range("A18").Value ' version Update
            Workbooks(Trimmed).Close True
            Unload utilityLink ' de-load this
        Case 2
            Workbooks.Open (e114address)
            Unload utilityLink ' de-load this
        Case Else
            ' do nothing for this
        End Select
    End If
End Sub

Private Sub Trimmed_Change()
    TR2R.Value = Trimmed.Value
End Sub

Private Sub Trimmed114_Change()
    T114.Value = Trimmed114.Value
End Sub

Private Sub Trimmer_SpinDown()
Select Case linkGuide.Value
Case 1 'RRR
    If LcTrim.Value - 1 > 0 Then
        LcTrim.Value = LcTrim.Value - 1
    End If
    trimRedux
Case 2 '114
    If Trim114.Value - 1 > 0 Then
        Trim114.Value = Trim114.Value - 1
    End If
    trimRedux
Case Else ' DO NOTHIN
End Select
End Sub

Private Sub Trimmer_SpinUp()
Select Case linkGuide.Value
Case 1 ' RRR
    If LcTrim.Value + 1 < 100 Then
        LcTrim.Value = LcTrim.Value + 1
    End If
    trimRedux
Case 2 ' 114
    If Trim114.Value + 1 < 100 Then
        Trim114.Value = Trim114.Value + 1
    End If
    trimRedux
Case Else ' DO NOTHING
End Select
End Sub

Private Sub uRRR_webDMO_Click()
    CreateObject("Shell.Application").ShellExecute _
        "microsoft-edge:https//https://dmoapps.csd.disa.mil/WebDMO/Login.aspx"
End Sub

Private Sub UserForm_Initialize()
    Set config = Worksheets("SENSEI.CONFIG")
    Set LcCell = config.Range("B2") 'address
    Set LcTrim = config.Range("B3") 'trim length for RRR
    Set VrHere = config.Range("B4") ' current ver
    Set Vrelease = config.Range("B5") ' release type
    Set Cell114 = config.Range("B6") ' 114 PATH
    Set Trim114 = config.Range("B7") ' 114 TRIMMING
    Set SenseiVer = config.Range("D4") ' sensei version
    Set T114 = config.Range("B8") ' 114 TRIMMED ADDRESS
    Set TR2R = config.Range("B9") ' RRR TRIMMED ADDRESS
    uRRR_Address.Value = LcCell.Value ' auto assume stored address
    e114address.Value = Cell114.Value ' ASSUME 114 ADDRESS
    ' TRIM THE RRR
    If Len(uRRR_Address.Value) > LcTrim.Value Then
        Trimmed.Value = Right(uRRR_Address.Value, LcTrim.Value)
    Else
        Trimmed.Value = vbNullString
    End If
    Length.Caption = LcTrim.Value
    ' TRIM THE 114
    If Len(e114address.Value) > Trim114.Value Then
        Trimmed114.Value = Right(e114address.Value, Trim114.Value)
    Else
        Trimmed114.Value = vbNullString
    End If
    Length114.Caption = Trim114.Value
    
    R2Rx = VrHere.Value
    R2Rver.Caption = R2Rx
    Lrelease.Caption = Vrelease.Value
    LinkVersion.Caption = "SENSEI LINK ver." & VrHere & "-" & Vrelease & " on " & SenseiVer
    'Set Vrrr = Workbooks(Trimmed).Sheets(1).Range("A18")
End Sub
Sub trimRedux() ' screen updating when adjusting trimming features
Dim trimming As Long
Select Case linkGuide.Value
Case 1 ' for RRR
    trimming = LcTrim.Value
    If Len(uRRR_Address.Value) > trimming Then
        Trimmed.Value = Right(uRRR_Address.Value, trimming)
    Else
        Trimmed.Value = vbNullString
    End If
    Length.Caption = trimming
Case 2 ' FOR 114
    trimming = Trim114.Value
    If Len(e114address.Value) > trimming Then
        Trimmed114.Value = Right(e114address.Value, trimming)
    Else
        Trimmed114.Value = vbNullString
    End If
    Length114.Caption = trimming
Case Else ' DO NOTHING
End Select
End Sub

Sub CommonLocationFinder()
Dim F_Locater As FileDialog
Set F_Locater = Application.FileDialog(msoFileDialogFilePicker)

' Branched Picker
With F_Locater
    Select Case linkGuide.Value
    Case 1 ' for RRR linking
        .AllowMultiSelect = False
        .Filters.Add "3R Engine", "*.xlsm", 1
        .Title = "Linking to RRR Engine"
        .ButtonName = "Link"
        .Show
    Case 2 ' for 114 Linking
        .AllowMultiSelect = False
        .Filters.Add "114 Infinite", "*.xlsm", 1
        .Title = "Linking to 114 - Infinite"
        .ButtonName = "Link"
        .Show
    Case Else ' do nothing
    End Select
    FileLc = .SelectedItems.Item(1) ' Write file Path to Location
End With

' Branched Assigner for update purpose
Select Case linkGuide.Value
Case 1 ' RRR
    LcCell.Value = FileLc
    uRRR_Address.Value = LcCell.Value
Case 2 ' 114
    Cell114.Value = FileLc
    e114address.Value = Cell114.Value
Case Else ' do nothing
End Select

End Sub
