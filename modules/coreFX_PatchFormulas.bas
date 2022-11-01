Attribute VB_Name = "coreFX_PatchFormulas"
Public Sub localeRepair()
Dim rangeSID As Range, rangeRID As Range, rangeSEL As Range
    Set rangeSID = Range("E3:E102")
    Set rangeRID = Range("I3:I102")
Dim patchSID As String
    patchSID = "=IFERROR(VLOOKUP([@SID],IF(SENSEI.CONFIG!$D$9=1,tableStage,tableStageEN),2,TRUE),"""")"
Dim patchRID As String
    patchRID = "=IFERROR(VLOOKUP([@RID],IF(SENSEI.CONFIG!$D$9=1,tableRequest,tableRequestEN),2,TRUE),"""")"

Range("E3").Value = patchSID
Range("I3").Value = patchRID
rangeSID.FillDown
rangeRID.FillDown
End Sub
