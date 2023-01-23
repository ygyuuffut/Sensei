Attribute VB_Name = "coreFX_PatchFormulas"
Public Sub localeRepair() ' updated for new method and inclusion for archive on 221206
Dim rangeSID As Range, rangeRID As Range, rangeDisp As Range, rangeCount As Range, rangeSEL As Range
    Set rangeSID = Worksheets("CSP.TR").Range("entryTable[STAGE]")
    Set rangeRID = Worksheets("CSP.TR").Range("entryTable[REQUEST]")
    Set rangeDisp = Worksheets("CSP.TR").Range("entryTable[DISP]")
    Set rangeCount = Worksheets("CSP.TR").Range("entryTable[COUNT]")
    
Dim patchSID As String
    patchSID = "=IFERROR(VLOOKUP([@SID],IF(SENSEI.CONFIG!$D$9=1,tableStage,tableStageEN),2,TRUE),"""")"
Dim patchRID As String
    patchRID = "=IFERROR(VLOOKUP([@RID],IF(SENSEI.CONFIG!$D$9=1,tableRequest,tableRequestEN),2,TRUE),"""")"
Dim patchDisp As String
    patchDisp = "=[@SID]"
Dim patchCount As String
    patchCount = "=IF([@ID]<>"""",1,0)"
    
    
rangeSID.Value = patchSID
rangeRID.Value = patchRID
rangeDisp.Value = patchDisp
rangeCount.Value = patchCount

' fix archive
    Set rangeSID = Worksheets("CSP.ACH").Range("entryArchive[STAGE]")
    Set rangeRID = Worksheets("CSP.ACH").Range("entryArchive[REQUEST]")
    Set rangeDisp = Worksheets("CSP.ACH").Range("entryArchive[DISP]")
    Set rangeCount = Worksheets("CSP.ACH").Range("entryArchive[COUNT]")
rangeSID.Value = patchSID
rangeRID.Value = patchRID
rangeDisp.Value = patchDisp
rangeCount.Value = patchCount
    
End Sub
