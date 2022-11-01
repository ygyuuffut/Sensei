VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} codeInfo 
   Caption         =   "Sensei CE Settings"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "codeInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "codeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ADSNstr As String
Public ADSNList As Range ' List of FMF
Public ADSNFlt As Range ' Looping Pointer
Public Dat_Val As Worksheet ' Validation Page
Public Dat_114 As Worksheet ' 114 itself
' todo: PASSED 220607

Private Sub hidePanel_Click()
codeInfo.Hide
End Sub
Private Sub infoADSN_Change()
Call FindADSN
End Sub
Private Sub infoCycle_Change()

On Error GoTo errH
    Range("G3").Value = infoUser.Value & " | " & infoCycle.Value
    Exit Sub
errH:
    MsgBox "Information inputted does not seems Qualified...", vbOKOnly, "User Info Issues"
    
End Sub
Private Sub infoUser_Change()

If infoCycle.Value = "" Then
    Exit Sub
End If
On Error GoTo errH
    Range("G3").Value = infoUser.Value & " | " & infoCycle.Value
    Exit Sub
errH:
    MsgBox "Information inputted does not seems Qualified...", vbOKOnly, "User Info Issues"

End Sub
Private Sub UserForm_Initialize()
Application.Calculation = xlCalculationManual
Set Dat_Val = Sheets("VALIDATION")
Set Dat_114 = Sheets("RUN.114")
Set ADSNList = Dat_Val.Range("E20:E112")
infoUser.Value = Left(Dat_114.Range("G3").Value, 2)
infoCycle.Value = Right(Dat_114.Range("G3").Value, 2)
Call FindADSN
Application.Calculation = xlCalculationAutomatic
End Sub
Sub FindADSN()
Call S_Opt
    With ADSNList
        Set ADSNFlt = .Find(infoADSN.Value, LookIn:=xlValues)
        If Not ADSNFlt Is Nothing Then ' when found it in list, do it
            infoFSO.Value = Dat_Val.Range("F" & ADSNFlt.Row).Value
        End If
    End With
    Set ADSNFlt = Nothing
Call S_Xit
End Sub


