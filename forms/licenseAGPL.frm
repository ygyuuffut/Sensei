VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} licenseAGPL 
   Caption         =   "GNU AGPL v3 License"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   OleObjectBlob   =   "licenseAGPL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "licenseAGPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub apCOPY_Click()
SetClipboard (apLicense.Text)
End Sub

Private Sub apOK_Click()
licenseAGPL.Hide
End Sub
