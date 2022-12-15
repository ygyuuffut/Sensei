Attribute VB_Name = "expdFX_CommonProcedures"
Option Explicit
''''' Commonly used procedures '''''

Sub saveThis() ' save the file

    Application.SendKeys ("{ENTER}")
    ActiveWorkbook.Save

End Sub
