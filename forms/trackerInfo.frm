VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} trackerInfo 
   Caption         =   "Sensei EUA (End-User Acknowledgement)"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6945
   OleObjectBlob   =   "trackerInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "trackerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub euaConsent_Click()
Worksheets("SENSEI.CONFIG").Range("D2").Value = 1
Unload trackerInfo
trackerAPI.Show
End Sub
