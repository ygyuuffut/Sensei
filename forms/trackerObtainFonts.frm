VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} trackerObtainFonts 
   Caption         =   "Sensei Font Repair"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "trackerObtainFonts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "trackerObtainFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hidePanel_Click()
    trackerObtainFonts.Hide
    trackerAPI.Show
End Sub

Private Sub obtainFonts_Click()
Dim resX As String: resX = MsgBox("By Downloading:" & vbCrLf & vbCrLf & "- Microsoft Edge will begin to download Fonts." & vbCrLf & "- Sensei will Save your work and Quit" & vbCrLf & vbCrLf & "Please restart the excel completely to complete activation once installed fonts.", vbOKCancel, "Sensei Dependency Installtion")
If resX = vbCancel Then Exit Sub

CreateObject("Shell.Application").ShellExecute _
        "microsoft-edge:https://fonts.google.com/download?family=JetBrains%20Mono"
CreateObject("Shell.Application").ShellExecute _
        "microsoft-edge:https://fonts.google.com/download?family=Noto%20Sans%20Symbols%202"

ThisWorkbook.Activate
saveThis
ThisWorkbook.Close
End Sub

