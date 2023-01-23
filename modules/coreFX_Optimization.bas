Attribute VB_Name = "coreFX_Optimization"
Sub S_Opt()
Application.ScreenUpdating = False
End Sub
Sub S_Xit()
Application.ScreenUpdating = True
End Sub

Sub testInstr()

Dim strA As String, i As Long
    strA = "ALPHA BETA OMEGA POS-X"
Dim strArray() As String
strArray = Split(strA, " ")
For i = 0 To UBound(strArray)
    MsgBox strArray(i)
Next i

End Sub

