Attribute VB_Name = "expdFX_CommonProcedures"
Option Explicit
Dim config As Worksheet, autoSaveCounter As Range, autoSaveTrigger As Range, _
    isAutoSave As Range

''''' Commonly used procedures '''''

Sub initializeCommon()  ' intitalize common variable

Set config = ThisWorkbook.Worksheets("SENSEI.CONFIG")
Set autoSaveCounter = config.Range("D24")
Set autoSaveTrigger = config.Range("D25")
Set isAutoSave = config.Range("D23")

End Sub


Sub saveThis() ' save the file

    Application.SendKeys ("{ENTER}")
    ActiveWorkbook.Save

End Sub

Sub acLastRow()

    MsgBox Cells.Find("*", Range("C1"), LookIn:=xlValues, SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row

End Sub

Sub globalSave() ' add one to count, found equal to some number wipe and save
initializeCommon
If isAutoSave.Value = False Then Exit Sub

    autoSaveCounter.Value = autoSaveCounter.Value + 1
    If autoSaveCounter.Value = autoSaveTrigger.Value Then
        Application.EnableEvents = False
        autoSaveCounter.Value = 0
        saveThis
        Application.EnableEvents = True
    End If

End Sub
