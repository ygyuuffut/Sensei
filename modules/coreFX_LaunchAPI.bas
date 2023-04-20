Attribute VB_Name = "coreFX_LaunchAPI"
Sub runAPI()
If Worksheets("SENSEI.CONFIG").Range("D2").Value = 0 Then
    trackerInfo.Show
Else
    Application.Caption = "Sensei - " & Worksheets("SENSEI.CONFIG").Range("D6").Value & "." & Worksheets("SENSEI.CONFIG").Range("D7").Value ' applet name
    trackerAPI.Show
End If
End Sub
