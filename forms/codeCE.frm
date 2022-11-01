VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} codeCE 
   Caption         =   "Sensei CE - Transcation Coding Engine"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12555
   OleObjectBlob   =   "codeCE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "codeCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public null0 As String
Public gotoTrans As String, Fid As String, Actn As String
' TN is the Primary Tab, T<FID> is specific tab
' AC_<FID> is the secondary, A<ACTN> is specific actn
' NOTE - Three part for Director:
' 1.Select case to pull Trans##
' 2.select case again to pull AC_<action code>
' 3.Build the actial page and sync the name

Private Sub CEmanualSW_Click()

If CEmanualSW.Value = False Then
    CEmanualSW.Caption = "Auto"
    With mbrName
        .Locked = True
        .Value = null0
    End With
Else
    CEmanualSW.Caption = "Manual"
    mbrName.Locked = False
End If
End Sub

Private Sub CEtransGo_Click() ' search fx that navigates to transaction directly
If CEtransStr.Value <> null0 And Len(CEtransStr.Value) = 4 Then
    Fid = Left(CEtransStr.Value, 2)
    Actn = Right(CEtransStr.Value, 2)
Else
    Exit Sub
End If
Call selectFid
End Sub

Private Sub CEtransStr_Change()
If CEtransStr.Value <> null0 And Len(CEtransStr.Value) = 4 Then
    Fid = Left(CEtransStr.Value, 2)
    Actn = Right(CEtransStr.Value, 2)
Else
    Exit Sub
End If
selectFid
End Sub


Private Sub CEuserInfo_Click()
codeInfo.Show
End Sub

Private Sub hidePanel_Click()
Sheets("CSP.TR").Activate
codeCE.Hide
trackerAPI.Show
End Sub

Private Sub mbrDoDID_Change()
    Call sdSwap
End Sub

Private Sub mbrName_Change()
    If Len(mbrName.Value) > 4 Then
        mbrMMPA.Value = Left(mbrName.Value, 5)
    Else
        mbrMMPA.Value = "TOO SHORT"
    End If
End Sub

Private Sub mbrSSN_Change()
    Call sdSwap
End Sub



Private Sub UserForm_Initialize()
    null0 = vbNullString
End Sub
Sub sdSwap() ' SSN or EDIPI, not both

If mbrSSN.Value <> null0 Then
mbrDoDID.Enabled = False
End If
If mbrDoDID.Value <> null0 Then
mbrSSN.Value = False
End If
If mbrDoDID.Value = null0 And mbrSSN.Value = null0 Then
mbrSSN.Enabled = True
mbrDoDID.Enabled = True
End If

End Sub
Sub FidHandler() ' when Fid is missing
    MsgBox "Transaction " & Fid & " does not seems to exist..."
End Sub
Sub ActnHandler() ' when action ismissing
    MsgBox "Action " & Actn & " for Transaction " & Fid & " does not seems to exist..."
End Sub

Sub selectFid()

Select Case Fid
    Case "10"
        Trans10
    Case "12"
        Trans12
    Case "14"
        Trans14
    Case "15"
        Trans15
    Case "21"
        Trans21
    Case "23"
        Trans23
    
    
    
    
    
    
    
    
    
    Case Else
        Call FidHandler
End Select

End Sub

Sub Trans10() ' Demolition Pay
Select Case Actn
    Case "01"
        AC_10.Value = 0
    Case "02"
        AC_10.Value = 1
    Case "03"
        AC_10.Value = 2
    Case "05"
        AC_10.Value = 3
    Case "06"
        AC_10.Value = 4
    Case Else
        Call ActnHandler
        Exit Sub
End Select
TN.Value = 1
End Sub
Sub Trans12() ' (ACIP) Fly Pay
Select Case Actn
    Case "01"
        AC_12.Value = 0
    Case "02"
        AC_12.Value = 1
    Case "03"
        AC_12.Value = 2
    Case "04"
        AC_12.Value = 3
    Case "05"
        AC_12.Value = 4
    Case "06"
        AC_12.Value = 5
    Case Else
        Call ActnHandler
        Exit Sub
End Select
TN.Value = 2
End Sub
Sub Trans14() ' (IDP) Imminent Danger Pay
Select Case Actn
    Case "01"
        AC_14.Value = 0
    Case "02"
        AC_14.Value = 1
    Case "03"
        AC_14.Value = 2
    Case "04"
        AC_14.Value = 3
    Case "05"
        AC_14.Value = 4
    Case "06"
        AC_14.Value = 5
    Case Else
        Call ActnHandler
        Exit Sub
End Select
TN.Value = 3
End Sub
Sub Trans15() ' JUMPS Pay
Select Case Actn
    Case "01"
        AC_15.Value = 0
    Case "02"
        AC_15.Value = 1
    Case "03"
        AC_15.Value = 2
    Case "04"
        AC_15.Value = 3
    Case "05"
        AC_15.Value = 4
    Case "06"
        AC_15.Value = 5
    Case Else
        Call ActnHandler
        Exit Sub
End Select
TN.Value = 4
End Sub
Sub Trans21() ' Dive Pay
Select Case Actn
    Case "01"
        AC_21.Value = 0
    Case "02"
        AC_21.Value = 1
    Case "03"
        AC_21.Value = 2
    Case "04"
        AC_21.Value = 3
    Case "05"
        AC_21.Value = 4
    Case "06"
        AC_21.Value = 5
    Case Else
        Call ActnHandler
        Exit Sub
End Select
TN.Value = 5
End Sub
Sub Trans23() ' Hostile Fire / Imminent Danger Pay
Select Case Actn
    Case "01"
        AC_23.Value = 0
    Case "02"
        AC_23.Value = 1
    Case "03"
        AC_23.Value = 2
    Case "04"
        AC_23.Value = 3
    Case "05"
        AC_23.Value = 4
    Case "06"
        AC_23.Value = 5
    Case Else
        Call ActnHandler
        Exit Sub
End Select
TN.Value = 6
End Sub
