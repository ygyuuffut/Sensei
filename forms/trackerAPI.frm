VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} trackerAPI 
   Caption         =   "Client Record Management Sensei"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "trackerAPI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "trackerAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ws As Worksheet, tbl As ListObject, SenXcel As Workbook, Ecsp As Worksheet
Public sid As Range, rid As Range, doDate As Range, clDate As Range, IQid As Range
Public sortFlag As Integer, filterFlag As Integer, sortRng As Range, filtRID As String
Public sortOrder As Boolean
Public stackFilterFlag As Boolean, stackCompoundFlag As Boolean
Public autoSave As Boolean
Public debugNotice As String, debugHH As String, doDebug As Boolean  ' Debug's Append
Public mainTrackerNextEmptRow As Range, cyclingCell As Range ' Addtion's Append
Public descRIDBox As String, descSIDBox As String, descRID As Integer, descSID As Integer ' Indexing RID - Passed 220511
Public ctrlSrc As String, editIQID As String, editSID As String, editSIDex As String, editCYC As String, editDDate As String, editRID As String, editRIDex As String, editComm As String, editReceive As String ' EDIT Variable Lib
' initialize common variable
Public Rcount As Long, editSorted As Integer, uCancel As Integer, Nvoid As String
Public config As Worksheet ' SENSEI CONFIG
Public Data As Worksheet ' SENSEI DATA (TABLES)
Public SconsoleVer As String, SlogVer As String, StdVer As String ' Sensei Version migrate
Public TypeVer As String ' WHATKIND OF RELEASE IS IT?
Public PtchVer As String ' PATCH VERSION
' append sort indicator
Public isOnAppendSort As Boolean
' Search Lib
Public searchDirection As Boolean, searchEdit As Boolean, searchShield As Boolean
Public srchObj As String, srchResult As String
' Archive Lib
Public Mfloater As Range, Afloater As Range, AchLimit As Range, ACsht As Worksheet
' Migration Lib, c FOR csp; s FOR sensei
Public migrateTY As String, formDataThisSensei As String
' SENSEI VERSION LIB
Public senseiVersion As String, senseiLogVer As String, senseiCoverLog As String
Const currentLong = 102 ' Table Last Row number
Const currentCapacity = 100 ' Max capacity
'TODO LIST on Tracker:
' > Appending need to prevent duplication
' > Search FX in Edit Tab should be able to perform drop-down
' > Dictionary search fro ADSN in ADSN Utility
'
'


Private Sub debugEmptyLocater_Click()
Call locateNextEmptySpot
End Sub
Private Sub debugLookForEntryDirectional_SpinDown()
searchEdit = False
    searchDirection = True
    Call debugLookForEntry
End Sub

Private Sub debugLookForEntryDirectional_SpinUp()
searchEdit = False
    searchDirection = False
    Call debugLookForEntry
End Sub

Private Sub debugReloadInfo_Click()
Call loadMatchestoDebug
End Sub


Private Sub formAppendExecute_Click() ' add a record to table - Operational 220510
Call SDT
Dim tempRow As Integer
If Range("M1").Value < 100 Then
    For Each cyclingCell In IQid
        If cyclingCell.Value = "" Then
            tempRow = cyclingCell.Row
            Range("C" & tempRow).Value = formAppendID.Value
            Range("D" & tempRow).Value = formAppendSID.Value
            Range("G" & tempRow).Value = formAppendDoDate.Value
            Range("H" & tempRow).Value = formAppendRID.Value
            Range("J" & tempRow).Value = formAppendActnComment.Value
            Range("K" & tempRow).Value = formAppendDate.Value
            Exit For
        End If
    Next cyclingCell
    MsgBox "Appended Inquiry: " & formAppendID.Value & " to Record!", vbOKOnly, "Client Record Management Sensei"
Else
    MsgBox "Record Table is full, please consider remove some Entries", vbOKOnly, "Client Record Management Sensei"
End If
Call RDT
debugNotice = debugHH & "[User]: Appended entry with ID " & formAppendID.Value & " to table"

If formAppendClean.Value = False Then
    formAppendID.Value = ""
    formAppendSID.Value = 1
    formAppendDoDate.Value = ""
    formAppendRID.Value = 1
    formAppendActnComment.Value = ""
End If

If formAppendAutoSort.Value = True Then
    Call SDT
    sortFlag = 1
    isOnAppendSort = True
    Call postActionSeries
    Call sortCaseMaster
    Call RDT
Else
    isOnAppendSort = False
    Call postActionSeries
End If
    Range("N1").Value = "=TODAY()"
update_Occupacy
Call RDT
End Sub

Private Sub formAppendHelpRID_Click()
trackerRIDHelp.Show
End Sub

Private Sub formAppendRID_Change() ' Append Page FX
If formAppendRID.Value < 1 Then
    formAppendRID.Value = 16
End If
If formAppendRID.Value > 16 Then
    formAppendRID.Value = 1
End If
    descRID = formAppendRID.Value
Call RIDexExplain
    formAppendRIDex.Value = descRIDBox
End Sub

Private Sub formCoverDelLog_Click()
Call resetLog
formCoverLog.Text = senseiCoverLog
End Sub

Private Sub formCoverDisplayDebug_Click() ' TODO: FIX THE DISPLAYING ISSUE
If formCoverDisplayDebug.Value = True Then
    fxSwitcher.Pages("MainDebug").Visible = True
    formCoverDebugTitle.Visible = True
    formCoverLog.Visible = True
    formCoverExportLog.Visible = True
    formCoverDelLog.Visible = True
End If
If formCoverDisplayDebug.Value = False Then
    fxSwitcher.Pages("MainDebug").Visible = False
    formCoverDebugTitle.Visible = False
    formCoverLog.Visible = False
    formCoverExportLog.Visible = False
    formCoverDelLog.Visible = False
End If
End Sub

Private Sub formCoverExportLog_Click() ' export log fx - Operational 220513
Dim debugLogCreate, debugLogItem, debugLog, debugLogLocation
Dim logLocator As FileDialog, logLocatorSel As String, logtime As String
logtime = (Format(Now(), "yymmddhhnnss"))
GoTo initialProcess

errorHandle:
    MsgBox "Exported nothing."
Exit Sub

initialProcess:
On Error GoTo errorHandle ' locate a path
Set logLocator = Application.FileDialog(msoFileDialogFolderPicker)
With logLocator
    .Title = "Storing Log to this location"
    .AllowMultiSelect = False
    .InitialFileName = Application.DefaultFilePath
    If .Show <> -1 Then GoTo secondProcess
    logLocatorSel = .SelectedItems(1)
    debugLogLocation = logLocatorSel
End With

secondProcess:
Set logLocator = Nothing
Set debugLogCreate = CreateObject("Scripting.FileSystemObject")
Set debugLogItem = debugLogCreate.CreateTextFile(debugLogLocation & "\SenseiLog - " & logtime & ".txt", True)
    debugLogItem.Close
    debugLog = debugLogLocation & "\SenseiLog - " & logtime & ".txt"
Open debugLog For Output As #1
    Print #1, formCoverLog.Text
Close #1
    Call resetLog
    formCoverLog.Text = senseiCoverLog
    MsgBox "Log exported to " & debugLogLocation
End Sub

Private Sub formCoverGuide_Click()
trackerAssist.Show
End Sub

Private Sub formCoverLang_Click()
SDT
If formCoverLang Then
    formCoverLang.Caption = "Lang: [EN-US]"
    config.Range("D9").Value = 2
Else
    formCoverLang.Caption = "Lang: [ZH-TW]"
    config.Range("D9").Value = 1
End If
labelLocaleAdj
    Ecsp.Columns("E").ColumnWidth = 12
    Ecsp.Columns("I").ColumnWidth = 20
debugNotice = debugHH & "[User]: Amended Display Language to " & Right(formCoverLang.Caption, 7)
RDT
postActionSeries
End Sub
Private Sub formCoverLang_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim resP As String, originalLC As Boolean
originalLC = Not formCoverLang
resP = MsgBox("Fix Language Display?", vbYesNo, "Locale Repair")
If resP = vbYes Then
    localeRepair ' fix formula
End If
formCoverLang = originalLC
End Sub

Private Sub formCoverLaunchCE_Click()
Sheets("RUN.114").Activate
trackerAPI.Hide
codeCE.Show
End Sub

Private Sub formCoverLaunchCL_Click()
    trackerAPI.Hide
    utilityDictionary.Show
End Sub

Private Sub formCoverLaunchDistill_Click()
trackerAPI.Hide
utilityForms.Show
End Sub

Private Sub formCoverLaunchLINK_Click()
On Error GoTo removal
    trackerAPI.Hide
    utilityLink.Show
removal:
    Set ws = Workbooks("SENSEI - dev.xlsm").Sheets("CSP.TR") ' redefine the target
End Sub

Private Sub formCoverTitle_Click()
licenseAGPL.Show
End Sub

Private Sub formCoverUpdateEntry_Click()
SDT
If formCoverUpdateEntry Then
    formCoverUpdateEntry.Caption = "Boot Update [I]"
    config.Range("D12").Value = True
Else
    formCoverUpdateEntry.Caption = "Boot Update [O]"
    config.Range("D12").Value = False
End If
debugNotice = debugHH & "[User]: Changed Auto Update by Reminder to " & formCoverUpdateEntry.Value
RDT
postActionSeries
End Sub

Private Sub FormDataAppendAmend_Click()
If FormDataAppendAmend Then
    config.Range("D10").Value = 1
Else
    config.Range("D10").Value = 2
End If
End Sub
Private Sub formDataConExp_Click()
On Error Resume Next

config.Range("D13").Value = formDataConExp.Value
End Sub


Private Sub formDataInflictUpdate_Click()
If FormDataAppendAmend Then
    config.Range("D11").Value = 1
Else
    config.Range("D11").Value = 2
End If
End Sub

Private Sub formDataMigrate_Click()
' File Locator Local Lib
Dim thisSensei As Workbook
    Set thisSensei = Workbooks(formDataThisSensei)
Dim senseiLocation, cspLocation, sourceXcel As Workbook
Dim xcelLocator As FileDialog, xcelLocation As String
' pointer Lib
Dim cellPointer As Range, countPointer As Long
    countPointer = 0
Call SDT
GoTo initiateImport

errorImport:
    MsgBox "Migration terminated by user actions", vbOKOnly, "Migration Info"
    Call RDT
Exit Sub

initiateImport:
On Error GoTo errorImport
If WorksheetFunction.Sum(Range("C3:K102").Value) <> 0 And formDataMigrateAll.Value = False Then
    MsgBox "Sensei Prevented you from losing potentially useful entries." & vbNewLine & "You may disable this in the config on the side", vbOKOnly, "Migration Protection"
    Exit Sub
End If

Set xcelLocator = Application.FileDialog(msoFileDialogFilePicker)
With xcelLocator
    .AllowMultiSelect = False
    .InitialFileName = Application.DefaultFilePath
    If migrateTY = "S" Then
        .Title = "Migrating from previous Sensei..."
        .Filters.Clear
        .Filters.Add "SENSEI Data", "*.xlsm", 1
    End If
    If migrateTY = "C" Then
        .Title = "Migrating from CSP exported xlsx..."
        .Filters.Clear
        .Filters.Add "CSP Records", "*.xlsx; *.xls", 1
    End If
    If .Show <> -1 Then GoTo secondaryImport
    If migrateTY = "S" Then
        senseiLocation = .SelectedItems(1)
    ElseIf migrateTY = "C" Then
        cspLocation = .SelectedItems(1)
    End If
End With

secondaryImport:
Set xcelLocator = Nothing
If migrateTY = "S" Then ' PASSING AS OF 220517
    Set sourceXcel = Workbooks.Open(senseiLocation)
    With sourceXcel.Sheets(1)
        For Each cellPointer In Range("C3:C102")
            If cellPointer.Value = "" Then
                Exit For
            End If
            countPointer = countPointer + 1
        Next cellPointer
        .Range("C3:K" & cellPointer.Row).Select
        .Range("C3:K" & cellPointer.Row).Copy
    End With
    thisSensei.Activate
    If formDataImportOptn = True Then
        thisSensei.Sheets(1).Range("C3:K102").Value = ""
    End If
    With thisSensei.Sheets(1)
        .Range("C3").Select
        .Range("C3").PasteSpecial Paste:=xlPasteValues
    End With
    Application.CutCopyMode = False
    sourceXcel.Close False
    debugNotice = debugHH & "[User]: Conducted Migration from Previous Version of Sensei (" & countPointer & " total) to Table"

ElseIf migrateTY = "C" Then ' passing as of 220517
    Set sourceXcel = Workbooks.Open(cspLocation)
    With sourceXcel.Sheets(1)
        .Range("A2:A101").Select
        .Range("A2:A101").Copy
    End With
    thisSensei.Activate
    If formDataImportOptn = True Then
        thisSensei.Sheets(1).Range("C3:K102").Value = ""
    End If
    With thisSensei.Sheets(1)
        .Range("C3").Select
        .Range("C3").PasteSpecial Paste:=xlPasteValues
        For Each cellPointer In IQid
            If cellPointer.Value <> "" Then
                If formDataInflictUpdate Then ' ADDED LOGIC WHEN UPDATE IS ENABLED
                    .Range("D" & cellPointer.Row).Value = 2
                Else
                    .Range("D" & cellPointer.Row).Value = 1
                End If
                .Range("J" & cellPointer.Row).Value = "New Migrated Entry"
                .Range("K" & cellPointer.Row).Value = formAppendDate.Value
                countPointer = countPointer + 1
            End If
        Next cellPointer
    End With
    Application.CutCopyMode = False
    sourceXcel.Close False
    debugNotice = debugHH & "[User]: Conducted Migration from CSP Export (" & countPointer & " total) to Table"
End If

'todo: all completed as of 220517
Call postActionSeries

If formDataInflictUpdate Then updateExistingEntry ' Trigger this update block

update_Occupacy
Call RDT
End Sub
Private Sub formDataFromCSP_Click()
formDataFromCSP.Value = True
formDataFromSensei.Value = False
migrateTY = "C"
End Sub

Private Sub formDataFromSensei_Click()
formDataFromCSP.Value = False
formDataFromSensei.Value = True
migrateTY = "S"
End Sub



Private Sub formDataNuke_Click()
Call SDT
Dim agrClear As String
    agrClear = MsgBox("Erase Sensei?" & vbNewLine & "This action is not reversible", vbOKCancel, "Erase Sensei")

If agrClear = vbOK And formDataNuke.Value = False Then
    ThisWorkbook.Sheets(1).Range("C3:K102").Value = ""
    With ACsht
        .Activate
        .Range("B3:L3002").Value = ""
    End With
    ThisWorkbook.Sheets(1).Activate
    debugNotice = debugHH & "[User] Removed All Entries in Sensei"
End If

If agrClear = vbOK And formDataNuke.Value = True Then
    With ACsht
        .Visible = xlSheetVisible
        .Activate
        .Range("B3:L3002").Value = ""
    End With
    ThisWorkbook.Sheets(1).Activate
    debugNotice = debugHH & "[User] Removed All Entries in Archive"
End If

update_Occupacy
Call RDT
Call postActionSeries
End Sub

Private Sub formDataNukeAll_Click()

Dim nukeClear As String
    nukeClear = MsgBox("Resetting to Default" & vbNewLine & "This action is not reversible!", vbOKCancel, "Sensei Factory Initialization")

If nukeClear = vbOK Then
    Ecsp.Range("C3:K102").Value = ""
    With ACsht
        .Activate
        .Range("B3:L3002").Value = ""
    End With
    With config
        .Range("B2:B3").Value = ""
        .Range("B6:B9").Value = ""
        .Range("D2").Value = 0
        .Range("D3").Value = 0
        .Range("D9:D11").Value = 2 ' for D9, it sets locale to EN, for switches, 2 is off
        .Range("D13").Value = False ' DISABLE DUAL UPDATE IN DATA
        .Range("D14").Value = 0 ' RE-ENGAGE WARNING PROMPT IN DATA.NUKE IT ALL
        .Range("F5").Value = False ' set distiller to per time
        .Range("F6").Value = "" ' set distiller path to null
        .Range("F7").Value = True ' ENABLE DISTILLER DELETE WARNING
    End With
    Ecsp.Activate
    debugNotice = debugHH & "[Master] Sensei Default Initialized"
End If

update_Occupacy
End Sub

Private Sub formDataUpdate_Click()
SDT
If formDataConExp.Value = True Then
    Dim resX As String
    If config.Range("D14").Value <> 2 And formDataConExp.Value = True Then resX = MsgBox("Dual Update is enabled, Now when updating to Stage 1, please read the file selection prompt carefully to avoid form breakage" & vbNewLine & vbNewLine & "By selecting Yes, this warning will nolonger display. ", vbYesNoCancel, "Dual Update Info Card")
    If resX = vbYes Then
        config.Range("D14").Value = 2
    End If
    updateExpiredEntry
    updateArchiveThem
End If

updateExistingEntry
update_Occupacy
RDT
End Sub

Sub updateExistingEntry() ' Updater for entries
Dim updaterV1 As FileDialog, EWPxcell As Workbook, KC As Long
Dim updateFile As String, updateFile2 As String, updaterLine As Long, updaterTar As Range
Dim trackerEnd As Range ' current end line of tracker
Dim searchRt As Range ' return result
Set updaterV1 = Application.FileDialog(msoFileDialogFilePicker)
KC = 0


On Error GoTo handler
With updaterV1
    .AllowMultiSelect = False
    .Filters.Add "Need to Work", "*.xlsx", 1
    .Title = "Looking for self-assigned RegAF-Total Inquiries.xlsx..."
    .ButtonName = "Update"
    .Show
    updateFile = .SelectedItems.Item(1)
End With

If updateFile = "" Then GoTo handler

' Run Workbooks if found
Set EWPxcell = Workbooks.Open(updateFile)
' validate file, loop through documents
If EWPxcell.Worksheets(1).Range("A1") <> "Inquiry ID" Then GoTo handler
updaterLine = EWPxcell.Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row

For Each updaterTar In EWPxcell.Worksheets(1).Range("A2:A" & updaterLine + 1)
    If updaterTar = vbNullString Then Exit For
    With Ecsp.Range("C1" & ":C" & currentLong) ' match and update
        Set searchRt = .Find(Left(updaterTar.Value, 18), after:=Range("C1"), LookIn:=xlValues)
    End With
    If Not searchRt Is Nothing Then ' if found it passed
        If Ecsp.Range("D" & searchRt.Row).Value = 2 Then Ecsp.Range("D" & searchRt.Row).Value = 1 ' only alter stage 2
    Else
        If FormDataAppendAmend Then ' go ahead and put the new entry on the form 221003
            With Ecsp
                If Not .Range("M1").Value < 100 Then ' if full, append nothing
                    MsgBox "Reached Max Capacity! Cannot append more entries, exiting..."
                    GoTo handler
                End If
                For Each trackerEnd In IQid ' Find next cell in available space
                    If trackerEnd.Value = "" Then
                        .Range("C" & trackerEnd.Row).Value = Left(updaterTar.Value, 18)
                        .Range("D" & trackerEnd.Row).Value = 1
                        .Range("J" & trackerEnd.Row).Value = "New Entry"
                        .Range("k" & trackerEnd.Row).Value = Format(Now(), "YYYY-MM-DD")
                        Exit For
                    End If
                Next trackerEnd
            End With
        End If
    End If
    KC = KC + 1
Next updaterTar
EWPxcell.Close False

' Adjustment to sorting
    sortFlag = 1
    isOnAppendSort = True
    Call postActionSeries
    Call sortCaseMaster
    MsgBox "Entries successfully updated/appended against Exports", vbOKOnly, "Entry Update Completion"
    Exit Sub
handler:
    MsgBox "Actions is either cancelled or incomplete due to exceptions.", vbOKOnly, "Entry Update encountered Exception(s)"
End Sub

Sub updateExpiredEntry() ' check all entires that doesnt exist to stage 5 then archive
Dim updaterV1 As FileDialog, EWPxcell As Workbook, KC As Long
Dim updateFile As String, updaterLine As Long, updaterTar As Range
Dim trackerEnd As Range ' current end line of tracker
Dim searchRt As Range ' return result
Set updaterV1 = Application.FileDialog(msoFileDialogFilePicker)
KC = 0

On Error GoTo handler
With updaterV1
    .AllowMultiSelect = False
    .Filters.Add "Complete Export", "*.xlsx", 1
    .Title = "Looking for Total RegAF-Total Inquiries.xlsx..."
    .ButtonName = "Apply Expirations"
    .Show
    updateFile = .SelectedItems.Item(1)
End With
If updateFile = "" Then GoTo handler ' no file exit

Set EWPxcell = Workbooks.Open(updateFile) ' open destination
    If EWPxcell.Worksheets(1).Range("A1") <> "Inquiry ID" Then GoTo handler ' csp export format wrong exit
updaterLine = EWPxcell.Worksheets(1).Range("A" & Rows.Count).End(xlUp).Row ' last line in source

For Each updaterTar In Ecsp.Range("C1" & ":C" & currentLong)
    If updaterTar.Value <> "" And Ecsp.Range("G" & updaterTar.Row).Value = "" Then ' blank reminder entry will be checked
        With EWPxcell.Worksheets(1).Range("A1:A" & updaterLine)
            Set searchRt = .Find(Left(updaterTar.Value, 18), after:=Range("A1"), LookIn:=xlValues)
        End With
        If searchRt Is Nothing Then ' DID NOT FIND IT
            With Ecsp
                .Range("D" & updaterTar.Row).Value = 5
            End With
            KC = KC + 1
        End If
    End If
Next updaterTar
EWPxcell.Close False ' close it when done

    sortFlag = 1
    isOnAppendSort = True
    Call postActionSeries
    Call sortCaseMaster
    Exit Sub
handler:
    MsgBox "Actions is either cancelled or incomplete due to exceptions.", vbOKOnly, "Entry Update encountered Exception(s)"
End Sub

Sub updateArchiveThem() ' Separated Archive function to here
Call SDT
ACsht.Visible = xlSheetVisible
Dim edCount As Long, acCount As Long, ACresponse As String
    edCount = 0
    acCount = 0

If formEditArchive1.Value = True Then ' DELETE operational 220516
    For Each Mfloater In sid
        If Mfloater.Value = 5 Then
            deleteEntry
            edCount = edCount + 1
        End If
        If Mfloater.Value = 3 And formEditConv3 Then ' if allow Archive 3 then ok
            Mfloater.Value = 5
            deleteEntry
            acCount = acCount + 1
        End If
        ' add expired function (maybe, also thinking of a separate export)
    Next Mfloater
    If Not formEditConv3 Then
        debugNotice = debugHH & "[Info]: Removed " & edCount & " Entries with Stage 5 from Record"
    Else
        debugNotice = debugHH & "[Info]: Removed " & edCount & " Entries with Stage 5; Converted " & acCount & " Entries and Removed them"
    End If
    Call postActionSeries
    GoTo endPoint
End If

If formEditArchive2.Value = True Then ' MOVE TO ARCHIVE
    For Each Mfloater In sid
        If Mfloater.Value = 5 Then
            With Range("B" & Mfloater.Row & ":L" & Mfloater.Row)
                .Select
                .Copy
            End With
            With ACsht
                .Activate
                For Each Afloater In AchLimit
                    If Afloater.Value = "" Then
                        With Range("B" & Afloater.Row)
                            .Select
                            .PasteSpecial xlPasteAll
                        End With
                        Exit For
                    End If
                Next Afloater
            End With
            Sheets("CSP.TR").Select
            deleteEntry
            edCount = edCount + 1
        End If
        If Mfloater.Value = 3 And formEditConv3 Then ' when stage 3 is authorized
            Mfloater.Value = 5
            With Range("B" & Mfloater.Row & ":L" & Mfloater.Row)
                .Select
                .Copy
            End With
            With ACsht
                .Activate
                For Each Afloater In AchLimit
                    If Afloater.Value = "" Then
                        With Range("B" & Afloater.Row)
                            .Select
                            .PasteSpecial xlPasteAll
                        End With
                        Exit For
                    End If
                Next Afloater
            End With
            Sheets("CSP.TR").Select
            deleteEntry
            edCount = edCount + 1
        End If
    Next Mfloater
    If Not formEditConv3 Then
        debugNotice = debugHH & "[Info]: Archived " & edCount & " Entries with Stage 5 to Archive"
    Else
        debugNotice = debugHH & "[Info]: Archived " & edCount & " Entries with Stage 5 to Archive; Converted " & acCount & " Entries and Archived them."
    End If
    Call postActionSeries
    GoTo endPoint
End If

endPoint:
RepairFreeFloaters
update_Occupacy
Call RDT
ACsht.Visible = xlSheetHidden
End Sub

Private Sub formDataUpdateRemind_Click()
    updateByRmindButton
    update_Occupacy
End Sub

Private Sub formDebugReset_Click()
    With ThisWorkbook.Sheets(1)
        .Range("C3:K102").Value = ""
        .Range("O1").Value = 0
    End With
    With ThisWorkbook.Sheets("CSP.ACH")
        .Activate
        .Range("B3:L3002").Value = ""
    End With
    ThisWorkbook.Sheets(1).Activate
End Sub

Private Sub formDebugShowConfig_Click()
If formDebugShowConfig Then
    config.Visible = xlSheetVisible
    Data.Visible = xlSheetVisible
Else
    config.Visible = xlSheetVeryHidden
    Data.Visible = xlSheetVeryHidden
End If
End Sub

Private Sub formDebugShowDialogs_Click()
doDebug = formDebugShowDialogs.Value
End Sub

Private Sub formDebugUnlock_Click()
If formDebugUnlock.Value = False Then
    Call hideDebug
    Exit Sub
End If
Dim uStr As String
If config.Range("D3").Value <> 1 Then
    uStr = MsgBox("You will be able to access Developer's scaffolding work" & vbNewLine & "Using them may cause stability issues (This is an One-Time Notice)", vbOKCancel, "Unlocking Debug - Consent")
    If uStr = vbCancel Then
        uCancel = 1
        formDebugUnlock.Value = False
    End If
End If
If formDebugUnlock.Value = True Then
    config.Range("D3").Value = 1
    Call displayDebug
End If
End Sub

Private Sub formEditArchive1_Click()
formEditArchive1.Value = True
formEditArchive2.Value = False
End Sub
Private Sub formEditArchive2_Click()
formEditArchive1.Value = False
formEditArchive2.Value = True
End Sub


Private Sub formEditComment_Change()
If formEditLoader And Not searchShield And formEditRowDisp.Value <> vbNullString Then ws.Range("J" & formEditRowDisp.Value).Value = formEditComment.Value
End Sub


Private Sub formEditConv3_Click()
If formEditConv3 Then
    formEditConv3.Caption = "Convert AC [I]"
Else
    formEditConv3.Caption = "Convert AC [O]"
End If
End Sub

Private Sub formEditCycle_Change()
If formEditLoader And formEditRowDisp.Value <> vbNullString And Not searchShield Then ws.Range("F" & formEditRowDisp.Value).Value = formEditCycle.Value
End Sub

Private Sub formEditDelEntry_Click() ' test row delete
Call SDT
Dim eDelRes As String
Dim sDelLoc As String
    sDelLoc = formEditID.Value
If formEditIQID.Value = "" Or formEditRowDisp.Value = "" Then  ' exit on blank
    Exit Sub
End If

If formEditPromptDel.Value = True Then
    eDelRes = MsgBox("Remove this entry?" & vbNewLine & vbNewLine & """" & formEditID.Value & """" & vbNewLine & vbNewLine & "From Record?", vbYesNo, "Delete Record")
Else
    Call formRemoveSingleEntry
    GoTo rollOver
End If

If formEditPromptDel.Value = True And eDelRes = vbYes Then
    Call formRemoveSingleEntry
End If

rollOver:
If formEditRollOver.Value = True Then
    Call formEditNavi_SpinDown
End If

If formEditAutoSort.Value = True Then
    sortFlag = 1
    editSorted = 1
    Call sortCaseMaster
End If
debugNotice = debugHH & "[User] Removed Entry " & """" & sDelLoc & """" & " from Tracker"
Call postActionSeries
Call RDT
End Sub

Private Sub formEditDoDate_Change()
If formEditLoader And formEditRowDisp.Value <> vbNullString And Not searchShield Then ws.Range("G" & formEditRowDisp.Value).Value = formEditDoDate.Value
End Sub

Private Sub formEditDODID_Click()
If formEditDODID Then
    formEditDODID.Caption = "Copy DODID [I]"
Else
    formEditDODID.Caption = "Copy DODID [O]"
End If
End Sub

Private Sub formEditID_Change()
If formEditLoader And formEditRowDisp.Value <> vbNullString And Not searchShield Then ws.Range("C" & formEditRowDisp.Value).Value = formEditID.Value

On Error GoTo halt
If formEditDODID Then ' Old MS fORM cOPY mETHOD DEPRECIATED, NEED REWORK
    SetClipboard (Left(formEditID.Value, 10))
End If
halt:
End Sub


Private Sub formEditLoader_Click()

If formEditLoader Then
    formEditLoader.Caption = "Save by Instant"
    formEditManSave.Enabled = False
ElseIf Not formEditLoader Then
    formEditLoader.Caption = "Save by Manual"
    formEditManSave.Enabled = True
End If

End Sub

Private Sub formEditManSave_Click()
SDT
amendEntry
debugNotice = debugHH & "[User]: Manually Amended Entry " & formEditID.Value & " on Row " & formEditRowDisp.Value
postActionSeries
RDT
End Sub

Private Sub formEditNavi_SpinDown()
editSorted = 0
searchEdit = False
searchShield = True
Call SDT
'If formEditLoader Then amendEntry (Disabled on 220824)
    searchDirection = True
    Call EditLookForEntry
Call RDT
If formEditAutoSort.Value = True Then
    sortFlag = 1
    editSorted = 1
    Call sortCaseMaster
End If
If editSorted > 0 Then
    searchEdit = False
    Call SDT
        searchDirection = True
        Call EditLookForEntry
    Call RDT
End If
Call editBoxValidate
debugNotice = debugHH & "[User]: Amended Entry " & formEditID.Value & " on Row " & formEditRowDisp.Value
If formEditLoader Then postActionSeries
searchShield = False ' release shield
End Sub

Private Sub formEditNavi_SpinUp()
editSorted = 0
searchEdit = False
searchShield = True
Call SDT
'If formEditLoader Then amendEntry (disabled on 220824)
    searchDirection = False
    Call EditLookForEntry
Call RDT
If formEditAutoSort.Value = True Then
    sortFlag = 1
    editSorted = 1
    Call sortCaseMaster
End If
If editSorted > 0 Then
    searchEdit = False
    Call SDT
        searchDirection = True
        Call EditLookForEntry
    Call RDT
End If
Call editBoxValidate
debugNotice = debugHH & "[User]: Amended Entry " & formEditID.Value & " on Row " & formEditRowDisp.Value
If formEditLoader Then postActionSeries
searchShield = False ' Release Shield
End Sub

Private Sub formEditRID_Change()
If formEditRID.Value = "" Then
    descRID = 0
    GoTo blankPt
End If
If formEditRID.Value > 16 Then
    formEditRID.Value = 1
End If
If formEditRID.Value < 1 Then
    formEditRID.Value = 16
End If
If formEditLoader And Not searchShield And formEditRowDisp.Value <> vbNullString Then ws.Range("H" & formEditRowDisp.Value).Value = formEditRID.Value

descRID = formEditRID.Value
blankPt:
Call RIDexExplain
formEditRIDex.Value = descRIDBox
End Sub

Private Sub formEditRIDadjust_SpinUp()
If formEditRID.Value = "" Then
    Exit Sub
Else
    formEditRID.Value = formEditRID.Value + 1
End If
End Sub
Private Sub formEditRIDadjust_SpinDown()
If formEditRID.Value = "" Then
    Exit Sub
Else
    formEditRID.Value = formEditRID.Value - 1
End If
End Sub

Private Sub formEditRunArchive_Click()
Dim ACresponse As String
    If formEditPromptDel.Value = True Then
        ACresponse = MsgBox("Want Sensei to handle Stage 5 Entries now?", vbYesNo, "Sensei - A Client Record Management Tool")
        If ACresponse = vbNo Then
            Exit Sub
        End If
    End If
updateArchiveThem
End Sub

Private Sub formEditSID_Change()
If formEditSID.Value = "" Then
    descSID = 0
    GoTo blankPt
End If
If formEditSID.Value > 5 Then
    formEditSID.Value = 5
End If
If formEditSID.Value < 1 Then
    formEditSID.Value = 1
End If
If formEditLoader And Not searchShield And formEditRowDisp.Value <> vbNullString Then ws.Range("D" & formEditRowDisp.Value).Value = formEditSID.Value

descSID = formEditSID.Value
blankPt:
Call SIDexExplain
formEditSIDex.Value = descSIDBox
End Sub

Private Sub formEditSIDadjust_SpinUp()
If formEditSID.Value = "" Then
    Exit Sub
Else
    formEditSID.Value = formEditSID.Value + 1
End If
End Sub

Private Sub formEditSIDadjust_SpinDown()
If formEditSID.Value = "" Then
    Exit Sub
Else
    formEditSID.Value = formEditSID.Value - 1
End If
End Sub

Private Sub formFilterAdjust_SpinDown() ' minus 1 to RID search box
formFilterRID.Value = formFilterRID.Value - 1

debugNotice = debugHH & "[Info]: RID Filter Criteria reduced to" & formFilterRID.Value
Call postActionSeries
Call formFilterAdjust_AutoPilot
End Sub

Private Sub formFilterAdjust_SpinUp() ' add 1 to the RID search box
formFilterRID.Value = formFilterRID.Value + 1

debugNotice = debugHH & "[Info]: RID Filter Criteria increased to" & formFilterRID.Value
Call postActionSeries
Call formFilterAdjust_AutoPilot
End Sub

Sub formFilterAdjust_AutoPilot() ' Auto FX on RID search
If formFormAutoRID.Value = True Then
    Call SDT
    filterFlag = 6
    filtRID = formFilterRID.Value
    Call filterCaseMaster
    debugNotice = debugHH & "[Info]: Triggered Filtering by RID"
    Call postActionSeries
    Call RDT
End If
End Sub

Private Sub formFilterHelpRID_Click()
trackerRIDHelp.Show
End Sub

Private Sub formFilterRID_Change() ' Operational 220420

If formFilterRID.Value < 1 Or formFilterRID.Value = "" Then
    formFilterRID.Value = 16
End If
If formFilterRID.Value > 16 Then
    formFilterRID.Value = 1
End If
formFilterRID.Value = Format(formFilterRID.Value, "00")

    descRID = Int(formFilterRID.Value)
Call RIDexExplain
Call formFilterAdjust_AutoPilot
    formFormRIDex.Value = descRIDBox
End Sub

Private Sub formFilterRunRID_Click() ' RID Filtering re-director
' Operational 220420
On Error GoTo errNotInt
Call SDT
filterFlag = 6
filtRID = formFilterRID.Value
Call filterCaseMaster
Call RDT
Exit Sub

errNotInt: ' portion operational 220419
debugNotice = debugHH & "[ERROR]: Exited - RID not Integer"
Call postActionSeries
Call RDT

End Sub

Private Sub formFilterS1_Click() ' RID Sub-procedure for Filtering
Call SDT
filterFlag = 1
Call filterCaseMaster
Call RDT
End Sub

Private Sub formFilterS2_Click() ' RID Sub-procedure for Filtering
Call SDT
filterFlag = 2
Call filterCaseMaster
Call RDT
End Sub

Private Sub formFilterS3_Click() ' RID Sub-procedure for Filtering
Call SDT
filterFlag = 3
Call filterCaseMaster
Call RDT
End Sub

Private Sub formFilterS4_Click() ' RID Sub-procedure for Filtering
Call SDT
filterFlag = 4
Call filterCaseMaster
Call RDT
End Sub

Private Sub formFilterS5_Click() ' RID Sub-procedure for Filtering
Call SDT
filterFlag = 5
Call filterCaseMaster
Call RDT
End Sub


Private Sub formFormStackCompound_Click()
    stackCompoundFlag = formFormStackCompound.Value
End Sub

Private Sub formFormStackFilter_Click()
    stackFilterFlag = formFormStackFilter.Value
End Sub

Private Sub formResetForm_Click() ' Reset Main Table
Call SDT
Call restoreForm
debugNotice = debugHH & "[User]: Restored Table"
Call postActionSeries
Call RDT
End Sub

Private Sub formSortDate_Click() ' RID Sub-procedure for Sorting
Call SDT
sortFlag = 4
Call sortCaseMaster
Call RDT
End Sub

Private Sub formSortDoDate_Click() ' RID Sub-procedure for Sorting
Call SDT
sortFlag = 3
Call sortCaseMaster
Call RDT
End Sub

Private Sub formSortOrder_Click()
' Mono Execution Switch Passed 220415

If sortOrder = False Then
    sortOrder = True
    With formSortOrder
        .Caption = "Descending"
        .ForeColor = &H40C0&
        .BorderColor = &H40C0&
    End With
ElseIf sortOrder = True Then
    sortOrder = False
    With formSortOrder
        .Caption = "Ascending"
        .ForeColor = &HC0C000
        .BorderColor = &HC0C000
    End With
End If

End Sub

Private Sub formSortRID_Click() ' RID Sub-procedure for Sorting
Call SDT
sortFlag = 2
Call sortCaseMaster
Call RDT
End Sub

Private Sub hidePanel_Click()
trackerAPI.Hide
End Sub

Private Sub formSortSID_Click() ' RID Sub-procedure for Sorting
Call SDT
sortFlag = 1
Call sortCaseMaster
Call RDT
End Sub


Private Sub Reset_Click()
RepairAPI
End Sub

Private Sub RIDAddMinusOne_spinup()
    formAppendRID.Value = formAppendRID.Value + 1
End Sub

Private Sub RIDAddMinusOne_spindown()
    formAppendRID.Value = formAppendRID.Value - 1
End Sub

Private Sub UserForm_Initialize()
' RANGE LOCK
Workbooks("SENSEI - dev.xlsm").Worksheets("CSP.TR").Activate
Set SenXcel = Workbooks("SENSEI - dev.xlsm")
Set Ecsp = SenXcel.Worksheets("CSP.TR")
Set config = Worksheets("SENSEI.CONFIG")
Set Data = Worksheets("SENSEI.DATA")
Set ACsht = Worksheets("CSP.ACH")
Set ws = Workbooks("SENSEI - dev.xlsm").Sheets("CSP.TR")
Set tbl = ws.ListObjects("entryTable")
Set sid = Range("entryTable[SID]")
Set rid = Range("entryTable[RID]")
Set doDate = Range("entryTable[DO.DATE]")
Set clDate = Range("entryTable[DATE]")
Set IQid = Range("entryTable[ID]")
Set AchLimit = Sheets("CSP.ACH").Range("C3:C3002") ' archive limit
' VARIABLE FIX
SconsoleVer = config.Range("D4").Value
SlogVer = config.Range("D5").Value ' log version
StdVer = config.Range("D6").Value
TypeVer = config.Range("D7").Value
PtchVer = Format(config.Range("D8").Value, "000")
ctrlSrc = "CSP.TR!"
stackFilterFlag = formFormStackFilter.Value
sortOrder = False
debugHH = SlogVer & "[" & Format(Now(), "hh:nn:ss") & "]"
formCoverLog.Text = SconsoleVer & " " & Format(Now(), "hh:nn:ss")
senseiCoverLog = SconsoleVer & " " & Format(Now(), "hh:nn:ss")
formAppendRIDex.Value = trackerRIDHelp.RID01.Value
formFormRIDex.Value = trackerRIDHelp.RID01.Value
searchDirection = False
isOnAppendSort = False
srchResult = 3
Nvoid = ""
migrateTY = "S"
formDataThisSensei = "SENSEI - dev.xlsm"
senseiVersion = "Sensei 1.4.0R"
senseiLogVer = SlogVer
' FIX SPREADSHEET
    initialize_Sheet
' Fix Locale and others
    initialize_Locale
    update_Occupacy
    updateConfig
' UPDATE COVER
CoverVersion.Caption = StdVer ' UPDATE MAIN VALUE
LogVersion.Caption = StdVer & " " & TypeVer & " " & PtchVer ' UPDATE UPDATE LOG VALUE
If TypeVer = "RELEASE" Then
    CoverVerType.Caption = TypeVer 'Left(TypeVer, 4) & "." ' Type of version (4+1 bytes)
Else
    CoverVerType.Caption = Left(TypeVer, 5)
End If
Call editBoxValidate
LinkVersion.Caption = SlogVer

End Sub
Sub updateConfig() ' update global config
    formDataConExp.Value = config.Range("D13").Value
End Sub
Sub update_Occupacy() ' update % occupied
    formDataOccupacy = Format(Ecsp.Range("M1").Value, "000") & " In Use / 100 Available"
    formDataOccupacyP = Format(Ecsp.Range("M1").Value, "000") & " %"
    formDataOccuP.Value = Ecsp.Range("M1").Value
End Sub
Sub initialize_Locale() ' not just locale
If config.Range("D9").Value = 1 Then ' locale
    formCoverLang.Value = False
Else
    formCoverLang.Value = True
End If
If config.Range("D10").Value = 1 Then ' Updater Auto amend function
    FormDataAppendAmend.Value = True
Else
    FormDataAppendAmend.Value = False
End If
If config.Range("D11").Value = 1 Then ' The Sorting upon Migration
    formDataInflictUpdate.Value = True
Else
    formDataInflictUpdate.Value = False
End If
If config.Range("D12") = True Then ' Auto Update sequence
    formCoverUpdateEntry.Value = True
Else
    formCoverUpdateEntry.Value = False
End If
    updateByRmind
    labelLocaleAdj
End Sub
Sub updateByRmind()
SDT
Dim aCell As Range, aRow As Long
If formCoverUpdateEntry Then ' IF MARK IS NOT GREATER THAN TODAY, CHANGE DATE
    For Each aCell In doDate
        aRow = aCell.Row
        If aCell <> "" And Ecsp.Range("C" & aRow).Value <> "" Then ' OMIT NA
            If Not DateValue(Format(aCell.Value, "YYYY-MM-DD")) > DateValue(Format(Now(), "YYYY-MM-DD")) And Ecsp.Range("D" & aRow).Value < 3 Then ' TODAY OR OLDER
                Ecsp.Range("D" & aRow).Value = 1
            ElseIf Ecsp.Range("D" & aRow).Value < 3 Then
                Ecsp.Range("D" & aRow).Value = 2
            End If
        End If
    Next aCell
    sortFlag = 1 ' Sort it by stage
    sortCaseMaster
Else
End If
RDT
End Sub
Sub updateByRmindButton()
SDT
Dim aCell As Range, aRow As Long

    For Each aCell In doDate
        aRow = aCell.Row
        If aCell <> "" And Ecsp.Range("C" & aRow).Value <> "" Then ' OMIT NA
            If Not DateValue(Format(aCell.Value, "YYYY-MM-DD")) > DateValue(Format(Now(), "YYYY-MM-DD")) And Ecsp.Range("D" & aRow).Value < 3 Then ' TODAY OR OLDER
                Ecsp.Range("D" & aRow).Value = 1
            ElseIf Ecsp.Range("D" & aRow).Value < 3 Then
                Ecsp.Range("D" & aRow).Value = 2
            End If
        End If
    Next aCell
    sortFlag = 1 ' Sort it by stage
    sortCaseMaster

RDT
End Sub
Sub initialize_Sheet()
    config.Visible = xlSheetVeryHidden
    Data.Visible = xlSheetVeryHidden
    ACsht.Visible = xlSheetHidden
End Sub
Sub labelLocaleAdj()
On Error GoTo handler
If formCoverLang Then
    With trackerRIDHelp
        .RID01.ControlSource = "=SENSEI.DATA!E79"
        .RID02.ControlSource = "=SENSEI.DATA!E80"
        .RID03.ControlSource = "=SENSEI.DATA!E81"
        .RID04.ControlSource = "=SENSEI.DATA!E82"
        .RID05.ControlSource = "=SENSEI.DATA!E83"
        .RID06.ControlSource = "=SENSEI.DATA!E84"
        .RID07.ControlSource = "=SENSEI.DATA!E85"
        .RID08.ControlSource = "=SENSEI.DATA!E86"
        .RID09.ControlSource = "=SENSEI.DATA!E87"
        .RID10.ControlSource = "=SENSEI.DATA!E88"
        .RID11.ControlSource = "=SENSEI.DATA!E89"
        .RID12.ControlSource = "=SENSEI.DATA!E90"
        .RID13.ControlSource = "=SENSEI.DATA!E91"
        .RID14.ControlSource = "=SENSEI.DATA!E92"
        .RID15.ControlSource = "=SENSEI.DATA!E93"
        .RID16.ControlSource = "=SENSEI.DATA!E94"
    End With
    With trackerSIDHelp
        .SID01.ControlSource = "=SENSEI.DATA!C73"
        .SID02.ControlSource = "=SENSEI.DATA!C74"
        .SID03.ControlSource = "=SENSEI.DATA!C75"
        .SID04.ControlSource = "=SENSEI.DATA!C76"
        .SID05.ControlSource = "=SENSEI.DATA!C77"
    End With
Else
    With trackerRIDHelp
        .RID01.ControlSource = "=SENSEI.DATA!E2"
        .RID02.ControlSource = "=SENSEI.DATA!E3"
        .RID03.ControlSource = "=SENSEI.DATA!E4"
        .RID04.ControlSource = "=SENSEI.DATA!E5"
        .RID05.ControlSource = "=SENSEI.DATA!E6"
        .RID06.ControlSource = "=SENSEI.DATA!E7"
        .RID07.ControlSource = "=SENSEI.DATA!E8"
        .RID08.ControlSource = "=SENSEI.DATA!E9"
        .RID09.ControlSource = "=SENSEI.DATA!E10"
        .RID10.ControlSource = "=SENSEI.DATA!E11"
        .RID11.ControlSource = "=SENSEI.DATA!E12"
        .RID12.ControlSource = "=SENSEI.DATA!E13"
        .RID13.ControlSource = "=SENSEI.DATA!E14"
        .RID14.ControlSource = "=SENSEI.DATA!E15"
        .RID15.ControlSource = "=SENSEI.DATA!E16"
        .RID16.ControlSource = "=SENSEI.DATA!E17"
    End With
    With trackerSIDHelp
        .SID01.ControlSource = "=SENSEI.DATA!C2"
        .SID02.ControlSource = "=SENSEI.DATA!C3"
        .SID03.ControlSource = "=SENSEI.DATA!C4"
        .SID04.ControlSource = "=SENSEI.DATA!C5"
        .SID05.ControlSource = "=SENSEI.DATA!C6"
    End With
End If
Exit Sub
handler:
    MsgBox "Excel's compile service did not start fully:" & vbNewLine & "It is recommended to restart Excel, or continue with degraded functions."
End Sub

Sub SDT() ' Performance Boost Part A
debugHH = senseiLogVer & "[" & Format(Now(), "hh:nn:ss") & "]"
' Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False
End Sub

Sub RDT() ' Performance Boost Part B
Application.ScreenUpdating = True
' Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
End Sub

Sub sortCaseMaster()
' Mono Execution Order|Stability Passed 220415

Select Case sortFlag
Case 1
    Set sortRng = sid
Case 2
    Set sortRng = rid
Case 3
    Set sortRng = doDate
Case 4
    Set sortRng = clDate
End Select


If sortOrder = False Then
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add key:=sortRng, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
ElseIf sortOrder = True Then
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add key:=sortRng, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With
End If
debugNotice = debugHH & "[User]: Sorting by Category done successfully"

Call postActionSeries

endPT:
End Sub

Sub filterCaseMaster()
' Mono Exexution with False Passed 220415

If stackFilterFlag = False Then
    If filterFlag < 6 Then
        With tbl.Range
            .AutoFilter
            .AutoFilter field:=3, Criteria1:=filterFlag
        End With
    Else
        With tbl.Range
            .AutoFilter
            .AutoFilter field:=7, Criteria1:=filtRID
        End With
    End If
ElseIf stackFilterFlag = True Then
    If filterFlag < 6 Then
        With tbl.Range
            .AutoFilter field:=3, Criteria1:=filterFlag
        End With
    Else
        With tbl.Range
            .AutoFilter field:=7, Criteria1:=filtRID
        End With
    End If
End If

debugNotice = debugHH & "[User]: Applied user appointed Filter successfully"
Call postActionSeries


End Sub


Sub postActionSeries()
    If formDebugShowDialogs.Value = True Then
        formCoverLog.Text = formCoverLog.Text & Chr(10) & debugNotice
    End If
    If formGlobalAutoSave = True Then
        Application.EnableEvents = False
        ActiveWorkbook.Save
    End If
    Range("N1").Formula = "=TODAY()"
End Sub

Sub restoreForm()
    With tbl.Range
        .AutoFilter
    End With
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add key:=sid, SortOn:=xlSortOnValues, Order:=xlAscending
    End With
End Sub

Sub locateNextEmptySpot() ' todo: passed 220419

For Each cyclingCell In IQid
    If cyclingCell.Value = "" Then
        MsgBox "Next Available Spot in Table is on Row: " & cyclingCell.Row
        Exit For
    End If
Next cyclingCell

End Sub
Sub debugLookForEntry() ' debug loading procedure
srchObj = debugFindID.Value
If debugLocatedRow <> "" Then
    srchResult = debugLocatedRow.Value
End If
    If debugFindID.Value = "" Then
        MsgBox "Blank entry, exiting"
        Exit Sub
    End If
Call findNextMatchingValue
Call loadMatchestoDebug
debugLocatedRow.Value = srchResult
End Sub
Sub EditLookForEntry() ' debug loading procedure
srchObj = formEditIQID.Value

If debugLocatedRow <> "" Then
    srchResult = formEditRowDisp.Value
End If
If formEditIQID.Value = "" Then
    Call editBoxValidate
    Call editCtrlSrcRemoval
    Exit Sub
End If

Call findNextMatchingValue
Call LoadResultEdit ' Temp disabled for this issue
formEditRowDisp.Value = srchResult

If searchEdit = True Then ' remove all data when found empty
    Call editCtrlSrcRemoval
End If
End Sub
Sub findNextMatchingValue() ' todo: Dual Directional Passed 220512
Dim Srch As Range

SearchCore: ' dual direction loop
With Range("C1" & ":C" & currentLong)
    If searchDirection = False Then ' FORWARD SEARCH LOOP
        If srchResult <> "" Then
            Set Srch = .Find(srchObj, after:=Range("C" & srchResult), LookIn:=xlValues)
        Else
            Set Srch = .Find(srchObj, after:=Range("C1"), LookIn:=xlValues)
        End If
    ElseIf searchDirection = True Then ' BACKWARD SEARCH LOOP
        If srchResult <> "" Then
            Set Srch = .Find(srchObj, after:=Range("C" & srchResult), LookIn:=xlValues, searchDirection:=xlPrevious)
        Else
            Set Srch = .Find(srchObj, after:=Range("C1"), LookIn:=xlValues, searchDirection:=xlPrevious)
        End If
    End If
    If Not Srch Is Nothing Then
        srchResult = Srch.Row ' Write Row Number for Reference
        'MsgBox "Found " & debugFindID.Value & " located at row: " & Srch.Row
    Else
        searchEdit = True
        Exit Sub ' WAS GOTO CUTOFF
    End If
    GoTo cutOff
End With

cutOff:
End Sub

Sub loadMatchestoDebug() ' todo: Operational Update Mpdule 220510
Dim SrchFB As Integer
    SrchFB = srchResult ' assign found row number to elements
Dim debugStrIQID As String
    debugStrIQID = ctrlSrc & "C" & SrchFB ' IQID
debugIQID.ControlSource = debugStrIQID
    debugStrIQID = ctrlSrc & "D" & SrchFB ' SID
debugSID.ControlSource = debugStrIQID

debugSIDex.Value = Range("E" & SrchFB).Value ' SID EXPLN

    debugStrIQID = ctrlSrc & "F" & SrchFB ' CYC
debugCYC.ControlSource = debugStrIQID
    debugStrIQID = ctrlSrc & "G" & SrchFB ' DO DATE
debugDoDate.ControlSource = debugStrIQID
    debugStrIQID = ctrlSrc & "H" & SrchFB ' TYPE RID
debugRID.ControlSource = debugStrIQID

debugRIDex.Value = Range("I" & SrchFB).Value ' RID EXPLN

    debugStrIQID = ctrlSrc & "J" & SrchFB ' ACTN COMMENT
debugActnComm.ControlSource = debugStrIQID
    debugStrIQID = ctrlSrc & "K" & SrchFB ' RECEIVING DATE
debugDate.ControlSource = debugStrIQID

End Sub

Sub LoadResultEdit() ' On: Debugging

Dim SrchFB As Integer
    SrchFB = srchResult ' assign found row number to elements
    Rcount = 0
Dim editStrIQID As String
     ' IQID
        formEditID.Value = ws.Range("C" & SrchFB).Value
     ' SID
        formEditSID.Value = ws.Range("D" & SrchFB).Value
     ' SID EXPLN
        formEditSIDex.Value = ws.Range("E" & SrchFB).Value
     ' CYC
        formEditCycle.Value = ws.Range("F" & SrchFB).Value
     ' DO DATE
        formEditDoDate.Value = ws.Range("G" & SrchFB).Value
     ' TYPE RID
        formEditRID.Value = ws.Range("H" & SrchFB).Value
     ' RID EXPLN
        formEditRIDex.Value = ws.Range("I" & SrchFB).Value
     ' ACTN COMMENT
        formEditComment.Value = ws.Range("J" & SrchFB).Value
     ' RECEIVING DATE
        formEditDate.Value = ws.Range("K" & SrchFB).Value
End Sub

Sub editCtrlSrcRemoval() ' data protection
    formEditID.ControlSource = ""
    formEditSID.ControlSource = ""
    formEditCycle.ControlSource = ""
    formEditComment.ControlSource = ""
    formEditDate.ControlSource = ""
    formEditDoDate.ControlSource = ""
    formEditRID.ControlSource = ""
    formEditID.Value = ""
    formEditSID.Value = ""
    formEditCycle.Value = ""
    formEditComment.Value = ""
    formEditDate.Value = ""
    formEditDoDate.Value = ""
    formEditRID.Value = ""
    formEditRowDisp.Value = ""
    Call editBoxDisable
End Sub

Sub formRemoveSingleEntry()
Dim Crow As Long
    Crow = formEditRowDisp.Value
    Range("C" & Crow).Value = ""
    Range("D" & Crow).Value = ""
    Range("F" & Crow).Value = ""
    Range("G" & Crow).Value = ""
    Range("H" & Crow).Value = ""
    Range("J" & Crow).Value = ""
    Range("K" & Crow).Value = ""
    formEditID.ControlSource = ""
    formEditSID.ControlSource = ""
    formEditCycle.ControlSource = ""
    formEditComment.ControlSource = ""
    formEditDate.ControlSource = ""
    formEditDoDate.ControlSource = ""
    formEditRID.ControlSource = ""
End Sub

Sub RIDexExplain() ' Append RID explain box display update - OP 220511

Select Case descRID
Case 1
    descRIDBox = trackerRIDHelp.RID01.Value
Case 2
    descRIDBox = trackerRIDHelp.RID02.Value
Case 3
    descRIDBox = trackerRIDHelp.RID03.Value
Case 4
    descRIDBox = trackerRIDHelp.RID04.Value
Case 5
    descRIDBox = trackerRIDHelp.RID05.Value
Case 6
    descRIDBox = trackerRIDHelp.RID06.Value
Case 7
    descRIDBox = trackerRIDHelp.RID07.Value
Case 8
    descRIDBox = trackerRIDHelp.RID08.Value
Case 9
    descRIDBox = trackerRIDHelp.RID09.Value
Case 10
    descRIDBox = trackerRIDHelp.RID10.Value
Case 11
    descRIDBox = trackerRIDHelp.RID11.Value
Case 12
    descRIDBox = trackerRIDHelp.RID12.Value
Case 13
    descRIDBox = trackerRIDHelp.RID13.Value
Case 14
    descRIDBox = trackerRIDHelp.RID14.Value
Case 15
    descRIDBox = trackerRIDHelp.RID15.Value
Case 16
    descRIDBox = trackerRIDHelp.RID16.Value
Case Else
    descRIDBox = ""
End Select

End Sub

Sub SIDexExplain()
Select Case descSID
Case 1
    descSIDBox = trackerSIDHelp.SID01.Value
Case 2
    descSIDBox = trackerSIDHelp.SID02.Value
Case 3
    descSIDBox = trackerSIDHelp.SID03.Value
Case 4
    descSIDBox = trackerSIDHelp.SID04.Value
Case 5
    descSIDBox = trackerSIDHelp.SID05.Value
Case Else
    descSIDBox = ""
End Select
End Sub
Sub displayDebug()

Label6.Visible = True
debugEmptyLocater.Visible = True
debugLookForEntryDirectional.Visible = True
Label22.Visible = True
debugFindID.Visible = True
Label19.Visible = True
debugLocatedRow.Visible = True
Label13.Visible = True
Label14.Visible = True
debugIQID.Visible = True
debugSID.Visible = True
debugSIDex.Visible = True
Label15.Visible = True
Label16.Visible = True
debugCYC.Visible = True
debugDoDate.Visible = True
debugRID.Visible = True
debugRIDex.Visible = True
Label17.Visible = True
Label18.Visible = True
debugActnComm.Visible = True
debugDate.Visible = True
debugReloadInfo.Visible = True
formDebugShowDialogs.Visible = True
formDebugReset.Visible = True
formDebugShowConfig.Visible = True

End Sub

Sub hideDebug()

Label6.Visible = False
debugEmptyLocater.Visible = False
debugLookForEntryDirectional.Visible = False
Label22.Visible = False
debugFindID.Visible = False
Label19.Visible = False
debugLocatedRow.Visible = False
Label13.Visible = False
Label14.Visible = False
debugIQID.Visible = False
debugSID.Visible = False
debugSIDex.Visible = False
Label15.Visible = False
Label16.Visible = False
debugCYC.Visible = False
debugDoDate.Visible = False
debugRID.Visible = False
debugRIDex.Visible = False
Label17.Visible = False
Label18.Visible = False
debugActnComm.Visible = False
debugDate.Visible = False
debugReloadInfo.Visible = False
formDebugShowDialogs.Visible = False
formDebugReset.Visible = False
formDebugShowConfig.Visible = False

End Sub

Sub deleteEntry()
        Range("C" & Mfloater.Row).Value = Nvoid
        Range("D" & Mfloater.Row).Value = Nvoid
        Range("F" & Mfloater.Row).Value = Nvoid
        Range("G" & Mfloater.Row).Value = Nvoid
        Range("H" & Mfloater.Row).Value = Nvoid
        Range("J" & Mfloater.Row).Value = Nvoid
        Range("K" & Mfloater.Row).Value = Nvoid
End Sub

Sub editBoxValidate()
If formEditIQID.Value = "" Or searchEdit = True Then
    formEditID.Enabled = False
    formEditSIDadjust.Enabled = False
    formEditSID.Enabled = False
    formEditCycle.Enabled = False
    formEditDoDate.Enabled = False
    formEditRIDadjust.Enabled = False
    formEditRID.Enabled = False
    formEditComment.Enabled = False
ElseIf formEditIQID.Value <> "" And searchEdit = False Then
    formEditID.Enabled = True
    formEditSIDadjust.Enabled = True
    formEditSID.Enabled = True
    formEditCycle.Enabled = True
    formEditDoDate.Enabled = True
    formEditRIDadjust.Enabled = True
    formEditRID.Enabled = True
    formEditComment.Enabled = True
End If
End Sub

Sub editBoxDisable()
    formEditID.Enabled = False
    formEditSIDadjust.Enabled = False
    formEditSID.Enabled = False
    formEditCycle.Enabled = False
    formEditDoDate.Enabled = False
    formEditRIDadjust.Enabled = False
    formEditRID.Enabled = False
    formEditComment.Enabled = False
End Sub

Sub resetLog()
senseiCoverLog = SconsoleVer & " " & Format(Now(), "hh:nn:ss")
End Sub
Sub amendEntry() 'entry amendment procedure
Dim RLC As Long

If formEditRowDisp.Value <> vbNullString Then
    RLC = formEditRowDisp.Value
Else
    Exit Sub
End If

With ws
    .Range("C" & RLC).Value = formEditID.Value
    .Range("D" & RLC).Value = formEditSID.Value
    .Range("F" & RLC).Value = formEditCycle.Value
    .Range("G" & RLC).Value = formEditDoDate.Value
    .Range("H" & RLC).Value = formEditRID.Value
    .Range("J" & RLC).Value = formEditComment.Value
    .Range("K" & RLC).Value = formEditDate.Value
End With

End Sub
Sub codingConfigInitialize()

End Sub
