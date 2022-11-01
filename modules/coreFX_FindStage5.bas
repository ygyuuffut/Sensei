Attribute VB_Name = "coreFX_FindStage5"
Option Explicit
Sub findAllStage5()
Dim floater As Range ' floating check in loop
Dim acFlter As Range ' floating check in archive
Dim typeResponse As String
    typeResponse = MsgBox("Yes - Remove" & vbNewLine & "No - Archive" & vbNewLine & "Cancel - Loop through with no action", vbYesNoCancel, "Action?")
Dim void As String
    void = ""
Dim currentLimit As Range
    Set currentLimit = Range("D3:D52")

For Each floater In currentLimit
    If floater.Value = 5 And typeResponse = vbCancel Then ' ok
        MsgBox "found Stage 5 at row " & floater.Row
    End If
    If floater.Value = 5 And typeResponse = vbYes Then ' ok
        MsgBox "Remove entry on row " & floater.Row
        Range("C" & floater.Row).Value = void
        Range("D" & floater.Row).Value = void
        Range("F" & floater.Row).Value = void
        Range("G" & floater.Row).Value = void
        Range("H" & floater.Row).Value = void
        Range("J" & floater.Row).Value = void
        Range("K" & floater.Row).Value = void
    End If
    If floater.Value = 5 And typeResponse = vbNo Then ' OK
        MsgBox "found Stage 5 at row " & floater.Row
        With Range("B" & floater.Row & ":L" & floater.Row)
            .Select
            .Copy
        End With
        With Sheets("CSP.ACH")
            .Activate
            For Each acFlter In Range("B3:B1000")
                If acFlter.Value = "" Then
                    With Range("B" & acFlter.Row)
                        .Select
                        .PasteSpecial xlPasteAll
                    End With
                    Exit For
                End If
            Next acFlter
        End With
        Sheets("CSP.TR").Select
        Range("C" & floater.Row).Value = void ' OPTIMIZATION 1
        Range("D" & floater.Row).Value = void
        Range("F" & floater.Row).Value = void
        Range("G" & floater.Row).Value = void
        Range("H" & floater.Row).Value = void
        Range("J" & floater.Row).Value = void
        Range("K" & floater.Row).Value = void
    End If
Next floater
End Sub
