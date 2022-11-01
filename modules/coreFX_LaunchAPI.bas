Attribute VB_Name = "coreFX_LaunchAPI"
Sub runAPI()
If Worksheets("SENSEI.CONFIG").Range("D2").Value = 0 Then
    trackerInfo.Show
Else
    trackerAPI.Show
End If
End Sub
