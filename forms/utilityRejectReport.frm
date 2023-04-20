VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} utilityRejectReport 
   Caption         =   "Sensei Rejection Report"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "utilityRejectReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "utilityRejectReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Publicly declare all the worksheets necessary to include: _
    SENSEI.DATA, REJECT.RPT, DATA.TEMP
Public sData As Worksheet, rRpt As Worksheet, rTemp As Worksheet, _
rejcMember As Range, ecsp As Worksheet
Public config As Worksheet, rData As Worksheet ' CONFIG AND SENSEI DATA
Public protectEmail As Boolean, rejcDeleteIO As Boolean, rejcAddIO As Boolean
Public exportOnly As Boolean
' aRow as the field keeping row number for modification _
  rejcMember as ws.listobjects("rejcMemberTable") _
  ## COL > N: NAME, O: ADMIN/MEMBER, P: ACT/DEACTIVE EMAIL, Q: ADMIN COPY?, R: EMAIL _
          rejcColN, rejcColO       , rejcColP             , rejcColQ,     , rejcColR _
          String    Boolean          Boolean                Boolean         String
' ## ROWID > rejcColRow as long
' ## AMEND > rejcAmend :: DELETE > rejcDelete, rejcDeleteIO, rejcAddIO

' ==================== HTML ======================
Public rejHtml As String, urejHtml As String, recHtml As String, memoHtml As String, _
       processedHtml As String
Public rejHtmlAdmin As String, urejHtmlAdmin As String, _
       recHtmlAdmin As String, memoHtmlAdmin As String, processedHtmlAdmin As String
Const tableHead = "<tr>", tableEnd = "</tr>"
Const rejTdColor = "<td style=""text-align: center; color: #912929"">", _
      UrejTdColor = "<td style=""text-align: center; color: #a44211"">", _
      recTdColor = "<td style=""text-align: center; color: #a46c11"">", _
      memoTdColor = "<td style=""text-align: center; color: #656565"">", _
      normTd = "<td style=""text-align: center"">", _
      flatTd = "<td>", _
      normTdEnd = "</td>"
Public techList As Dictionary, rejcTech As Dictionary ' use tech list for email ref
Public techMemo As Dictionary ' the memorandum dictionary for person
Public fsoRejectCU As Long, fsoRecycleCU As Long, fsoMemoCU As Long, _
       fsoMiscCU As Long, fsoProcessedCU As Long ' count All
Public tradBuffLast As Long ' buffed total raw entry count
Public fsoRejectCUA As Long, fsoRecycleCUA As Long, fsoMemoCUA As Long, _
       fsoMiscCUA As Long, fsoProcessedCUA As Long ' admin COUNT ALL
' ================================================
' ==================== EMAIL =====================
Public exportHTML As String, bCClist As String
' ================================================



Private Sub mainExecute_Click() ' boot reject report locally
    rejcReportCore
End Sub

Sub rejcReportCore() ' The actual thing
Dim txtRpt As String, txtCov ' file name for import and variant for reduced friction.
Dim filePk As FileDialog: Set filePk = Application.FileDialog _
                                           (msoFileDialogFilePicker) ' Bind it
Dim thisTech As String, thisEmail As String
Dim refactoredEmail As String
Dim thisUpdate As String, thisUpdateLn As Long ' the update number

' move email argument here
Dim msOutLookAP ' the ol app
Dim msOutlookMA ' the ol mail item

rTemp.Visible = xlSheetVisible
exportHTML = dataRejectNotice.Mail.Text ' html data
'memo: final;ize base html, test the block and complete email structure

' >>> Load the file
' if <realized we have a given location> and <known file name.txt> then
'   Pull it automatically based on path and name
' else

'========== debug skip to main
'GoTo debugSkipLoadPreReq
'=============================
' -----> Pull file from Presets when enabled + has path
If rejcSourceOptn.Value And rejcSourcePath.Value <> "" Then
    txtCov = rejcSourcePath.Text
    txtRpt = txtCov ' convert
    txtCov = vbNullString
    GoTo preLoaded
End If

' -----> Pull file from Explorer
Do While txtRpt = vbNullString
    With filePk
        .Filters.Clear
        .Filters.Add "DMO Report", "*.txt"
        .Title = "Sensei - Pulling Report from this Source."
        .ButtonName = "Make"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = -1 Then ' open window
            txtCov = .SelectedItems.Item(1)
            txtRpt = txtCov ' convert
            txtCov = vbNullString
        Else
            Exit Sub
        End If
    End With
Loop
' end if

preLoaded:
' >>> Write the Report rTemp as buffer, rRpt as actual
With rTemp
    .Activate
    With rTemp.QueryTables.Add(Connection:="text;" & txtRpt, Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 11
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    tradBuffLast = rTemp.Cells.Find("*", searchDirection:=xlPrevious, SEARCHORDER:=xlByRows).Row
        .Range("B1:B" & tradBuffLast).NumberFormat = "000000000"
    thisUpdate = Replace(Left(.Range("O" & 1).Value, 5), "-", "") ' THIS UPDATE NUMBER STR
End With

' >>> Skip Pre-Req on export only
If exportOnly Then GoTo bootAdmin

' >>> Check Pre-requisites
thisUpdateLn = config.Range("B35").Value ' LOAD LTS TO CACHE
config.Range("B35").Value = thisUpdate ' LOAD NEW TO LTS

If thisUpdateLn < config.Range("B35").Value Then '(CONFIG B35) COMPARE NEW TO CACHE
    ' PROCEED, WE ALREADY WROTE THE NEW DATE IN
    thisUpdateLn = config.Range("B35").Value
Else
    MsgBox "This update might've been done already", vbOKOnly, "Duplicated Work"
    rTemp.Range("B3:ZZ9999").Delete xlShiftUp ' NUKE
    rTemp.Visible = xlSheetVeryHidden
    ecsp.Activate
    Exit Sub
End If

' >>> Write Traditional Full Sheet
' (i) ALL COMPONENTS: Status  Transaction  mbrSSAN  mbrName  CycleID  Techn  errorCD  expln
'                     COL.A   COL.D        COL.B    COL.C    COL.F    COL.H  COL.Q    COL.R
'
' (i) Need memo specific: expln
'                         COL.U
'

debugSkipLoadPreReq:
' >> ESTABLISH DICTIONARY ITEM FOR ITERATION; ITERATION ITEM
Dim i As Long, tKey
tradBuffLast = rTemp.Cells.Find("*", searchDirection:=xlPrevious, SEARCHORDER:=xlByRows).Row


' >> INITIAL LOAD OF ADMIN DATA
bootAdmin:
loadAdminTable

' >> Skip Main on export only
If exportOnly Then GoTo terminalExitNoError

' >> Individual Block of Execution
For Each tKey In techList.Keys
    
    ' >> Yoink Person, reset counter, select email
    thisTech = tKey
    thisEmail = rejcTech.Item(tKey)
    fsoRejectCU = 0
    fsoRecycleCU = 0
    fsoMemoCU = 0
    fsoMiscCU = 0
    fsoProcessedCU = 0
    rejHtml = ""
    recHtml = ""
    memoHtml = ""
    processedHtml = ""
    urejHtml = ""
    techMemo.RemoveAll
    
    ' >> Use do while to prevent 9999+ email item crash PC
emailInspect:
    ' >> INIT THE EMAIL AND OUTLOOK
    Set msOutLookAP = CreateObject("Outlook.Application")
    Set msOutlookMA = msOutLookAP.CreateItem(0)
    
    Do While msOutLookAP.Application.Inspectors.Count > 9
        MsgBox "There are 10 Email instances running already, please send or dispose them first before press ok to continue"
        On Error GoTo -1
        On Error GoTo issueInspect
    Loop


' >>> NEED TO GO AHEAD DETERMINE WHEN TO SEND THE ADMIN COPY

    ' ----->[MEMBER COPY]<-----
    If techList.Item(tKey) <> 1 Then
    
    ' >> Pass to load memorandum dictionary
    For i = 1 To tradBuffLast
        If rTemp.Range("H" & i).Value = tKey And _
            Not techMemo.Exists(rTemp.Range("B" & i).Value) Then ' if MATCH AND NEW
            techMemo.Add rTemp.Range("B" & i).Value, "CACHE"
        End If
    Next i
    
    ' >> Pass 1 to write Table data
    For i = 1 To tradBuffLast ' construct sample reject table content for OL
    ' HIERARCHY: REJECT -> RECYCLE -> MANAGEMENT NOTICE -> PROCESSED -> ANYTHING ELSE
        If rTemp.Range("H" & i).Value = tKey And (UCase(rTemp.Range("A" & i).Value) = "REJECT" Or UCase(rTemp.Range("A" & i).Value) = "REJ PART") Then
            rejHtml = rejHtml & tableHead & vbNewLine & "  " & _
                      rejTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & Format(rTemp.Range("B" & i).Value, "000000000") & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                      rejTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                      flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                      tableEnd & vbNewLine
            fsoRejectCU = fsoRejectCU + 1
        ElseIf rTemp.Range("H" & i).Value = tKey And UCase(rTemp.Range("A" & i).Value) = "RECYCLE" Then
            recHtml = recHtml & tableHead & vbNewLine & "  " & _
                      recTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & Format(rTemp.Range("B" & i).Value, "000000000") & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                      recTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                      flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                      tableEnd & vbNewLine
            fsoRecycleCU = fsoRecycleCU + 1
        ElseIf UCase(rTemp.Range("A" & i).Value) = "MGMT NTC" And _
            techMemo.Exists(rTemp.Range("B" & i).Value) Then ' MANAGEMENT NTC TO MAPPED (U)
            memoHtml = memoHtml & tableHead & vbNewLine & "  " & _
                      memoTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & Format(rTemp.Range("B" & i).Value, "000000000") & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                      memoTdColor & tKey & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                      flatTd & rTemp.Range("U" & i).Value & normTdEnd & vbNewLine & _
                      tableEnd & vbNewLine
            fsoMemoCU = fsoMemoCU + 1
        ElseIf rTemp.Range("H" & i).Value = tKey _
          And UCase(rTemp.Range("A" & i).Value) = "PROCESSED" Then ' PROCESSED
            processedHtml = processedHtml & tableHead & vbNewLine & "  " & _
                      memoTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & Format(rTemp.Range("B" & i).Value, "000000000") & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                      memoTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                      flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                      tableEnd & vbNewLine
            fsoProcessedCU = fsoProcessedCU + 1
        ElseIf techMemo.Exists(rTemp.Range("B" & i).Value) And _
        rTemp.Range("H" & i).Value = thisTech Then  ' anything else goes here
            urejHtml = urejHtml & tableHead & vbNewLine & "  " & _
                      UrejTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & Format(rTemp.Range("B" & i).Value, "000000000") & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                      UrejTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                      normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                      flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                      tableEnd & vbNewLine
            fsoMiscCU = fsoMiscCU + 1
        End If
    Next i
    ' >> Re-format if some table data is blank
    reformEmptyTable
    End If
    ' ----->[END MEMBER COPY]<-----
    
    ' >> Draft Individual Email Copy
    With msOutlookMA
        .SentOnBehalfOfName = rejcEmailOnBehalf.Value
        If rejcCCthis.Value = True Then
            .cC = rejcEmailOnBehalf.Value
        End If
        .Importance = 2
        .To = thisEmail
        .Subject = "(CUI)(USE HTML VIEW) Daily Rejection Report - " & Format(Now(), "YYYYMMDD")
        ' CHANGE BODY
        refactoredEmail = dataRejectNotice.Mail.Text ' default body
        If techList.Item(tKey) = 0 Then ' the basic
            ' the memo
            refactoredEmail = Replace(refactoredEmail, "[memoAssignedTo]", "Assigned to")
            ' the table
            refactoredEmail = Replace(refactoredEmail, "[managementNoticeData]", memoHtml)
            refactoredEmail = Replace(refactoredEmail, "[recycleData]", recHtml)
            refactoredEmail = Replace(refactoredEmail, "[unassignedData]", urejHtml)
            refactoredEmail = Replace(refactoredEmail, "[rejectData]", rejHtml)
            refactoredEmail = Replace(refactoredEmail, "[processedData]", processedHtml)
            ' the counts
            refactoredEmail = Replace(refactoredEmail, "[fsoMemoCount]", fsoMemoCU)
            refactoredEmail = Replace(refactoredEmail, "[fsoRecycleCount]", fsoRecycleCU)
            refactoredEmail = Replace(refactoredEmail, "[fsoRejectCount]", fsoRejectCU)
            refactoredEmail = Replace(refactoredEmail, "[fsoMiscCount]", fsoMiscCU)
            refactoredEmail = Replace(refactoredEmail, "[fsoProcessedCount]", fsoProcessedCU)
            ' update#, date and copy type
            refactoredEmail = Replace(refactoredEmail, "[YYMMDD-HH:MM:SS]", Format(Now(), "YYMMDD-HH:MM:SS"))
            refactoredEmail = Replace(refactoredEmail, "[acUpdate]", Format(thisUpdateLn, "0000"))
            refactoredEmail = Replace(refactoredEmail, "[personalCopy]", "EXECUTION DISTRIBUTION")
            ' >> POST IT
            .HTMLBody = refactoredEmail
            .BCC = bCClist
        ElseIf techList.Item(tKey) = 1 Then ' the admin
            ' the memo
            refactoredEmail = Replace(refactoredEmail, "[memoAssignedTo]", "Auto ID")
            ' the table
            refactoredEmail = Replace(refactoredEmail, "[managementNoticeData]", memoHtmlAdmin)
            refactoredEmail = Replace(refactoredEmail, "[recycleData]", recHtmlAdmin)
            refactoredEmail = Replace(refactoredEmail, "[unassignedData]", urejHtmlAdmin)
            refactoredEmail = Replace(refactoredEmail, "[rejectData]", rejHtmlAdmin)
            refactoredEmail = Replace(refactoredEmail, "[processedData]", processedHtmlAdmin)
            ' the counts
            refactoredEmail = Replace(refactoredEmail, "[fsoMemoCount]", fsoMemoCUA)
            refactoredEmail = Replace(refactoredEmail, "[fsoRecycleCount]", fsoRecycleCUA)
            refactoredEmail = Replace(refactoredEmail, "[fsoRejectCount]", fsoRejectCUA)
            refactoredEmail = Replace(refactoredEmail, "[fsoMiscCount]", fsoMiscCUA)
            refactoredEmail = Replace(refactoredEmail, "[fsoProcessedCount]", fsoProcessedCUA)
            ' update#, date and copy type
            refactoredEmail = Replace(refactoredEmail, "[YYMMDD-HH:MM:SS]", Format(Now(), "YYMMDD-HH:MM:SS"))
            refactoredEmail = Replace(refactoredEmail, "[acUpdate]", Format(thisUpdateLn, "0000"))
            refactoredEmail = Replace(refactoredEmail, "[personalCopy]", "ADMINISTRATIVE ARCHIVE")
            ' >> POST IT
            .HTMLBody = refactoredEmail
        Else
        End If
        '.send
        .Display ' OR .SEND, BUT IT MAY BE BLOCKED
    End With
    
    ' >> Unload email for next Person
    Set msOutLookAP = Nothing
    Set msOutlookMA = Nothing

    
Next tKey
GoTo terminalExitNoError


' >>> SIEVE THE INFORMATION
' Go through the iteration loop and Mark all the transactions with appropriate actions
' Append those require actions to arrays

' >>> Finalize information
' Export the traditional Report with AC date and Make date
' Wipe Traditional Report and BUFFER

' >>> Save and Send the Email
' >>> Save the html if option is available

issueInspect:
' STALL APPLICATION INCASE OF ERROR OCCURING DUE TO DELAYED PROCESS
If msOutLookAP = Empty Then ' >>>> error 93 ??
    Application.Wait (Now + TimeValue("0:00:01")) 'wait a second
    Err = 0
    GoTo emailInspect
End If
If Err = 462 Then
    Set msOutLookAP = CreateObject("Outlook.Application")
    Err = 0
    GoTo emailInspect
End If
If Err = 424 Then
    Set msOutLookAP = CreateObject("Outlook.Application")
    Err = 0
    GoTo emailInspect
End If

terminalExitNoError:
'SetClipboard (rejHtml)
Application.StatusBar = "Pulled Update from DJMS " & _
    Format(config.Range("B35").Value, "0000") & "; Verified " & _
    tradBuffLast & " DMO Entries."
If rejcPrintEnable.Value Then
    refactoredEmail = dataRejectNotice.Mail.Text
    ' the memo
    refactoredEmail = Replace(refactoredEmail, "[memoAssignedTo]", "Auto ID")
    ' the table
    refactoredEmail = Replace(refactoredEmail, "[managementNoticeData]", memoHtmlAdmin)
    refactoredEmail = Replace(refactoredEmail, "[recycleData]", recHtmlAdmin)
    refactoredEmail = Replace(refactoredEmail, "[unassignedData]", urejHtmlAdmin)
    refactoredEmail = Replace(refactoredEmail, "[rejectData]", rejHtmlAdmin)
    refactoredEmail = Replace(refactoredEmail, "[processedData]", processedHtmlAdmin)
    ' the counts
    refactoredEmail = Replace(refactoredEmail, "[fsoMemoCount]", fsoMemoCUA)
    refactoredEmail = Replace(refactoredEmail, "[fsoRecycleCount]", fsoRecycleCUA)
    refactoredEmail = Replace(refactoredEmail, "[fsoRejectCount]", fsoRejectCUA)
    refactoredEmail = Replace(refactoredEmail, "[fsoMiscCount]", fsoMiscCUA)
    refactoredEmail = Replace(refactoredEmail, "[fsoProcessedCount]", fsoProcessedCUA)
    ' update#, date and copy type
    refactoredEmail = Replace(refactoredEmail, "[YYMMDD-HH:MM:SS]", Format(Now(), "YYMMDD-HH:MM:SS"))
    refactoredEmail = Replace(refactoredEmail, "[acUpdate]", Format(thisUpdateLn, "0000"))
    refactoredEmail = Replace(refactoredEmail, "[personalCopy]", "ADMINISTRATIVE ARCHIVE")
    exportHTML = refactoredEmail
    ' Weite to file
    expHtmlToFile
End If

rTemp.Visible = xlSheetVeryHidden ' HIDE HIS ASS
ecsp.Activate ' return
globalSave ' trigger save
nukeDataBuffer
End Sub

Sub expHtmlToFile() ' spit ADMINISTRATIVE ARCHIVE to folder
Dim makeObjStem, makeObjHtml, objHtml
Dim expFLDR, selFLDR As String, htmlFile As String, Cexist As String

On Error GoTo ExceptionPathway
If Not rejcPrintOptn.Value Or rejcPrintPath.Value = "" Then ' when nothing is destination
    Set expFLDR = Application.FileDialog(msoFileDialogFolderPicker)
    With expFLDR
        .Title = "Exporting Rejection Web Archive to here..."
        .ButtonName = "Save"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = 0 Then Exit Sub ' if cancelled exit immediately
        selFLDR = .SelectedItems(1)
    End With
Else ' we have something in destination and we are sending document to the
    Cexist = Dir(rejcPrintPath.Value & "\Reject Archive Exports\", vbDirectory) ' check parent
    If Cexist = "" Then
        Cexist = rejcPrintPath.Value & "\Reject Archive Exports\"
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    Cexist = Dir(rejcPrintPath.Value & "\Reject Archive Exports\" & Format(Now(), "YYYY-MM"), vbDirectory) ' check child
    If Cexist = "" Then
        Cexist = rejcPrintPath.Value & "\Reject Archive Exports\" & Format(Now(), "YYYY-MM")
        MkDir Cexist ' A FIXED DIRECTORY
    End If
    ' confirm selFLDR path
    selFLDR = rejcPrintPath.Value & "\Reject Archive Exports\" & Format(Now(), "YYYY-MM")
End If

' >> Make Object
'Set makeObjStem = CreateObject("Scripting.FileSystemObject")
'Set makeObjHtml = makeObjStem.CreateTextFile(selFLDR & "\Archive Record - DJMS " & _
        Format(config.Range("B35").Value, "0000") & _
        "." & Format(Now(), "YYMMDD") & ".HTML", True)
'    makeObjHtml.Close

' >> Appoint Object, Write data
objHtml = selFLDR & "\Archive Record - DJMS " & _
        Format(config.Range("B35").Value, "0000") & _
        "." & Format(Now(), "YYMMDD") & ".HTML"
Open objHtml For Output As #1
    Print #1, exportHTML;
Close #1
Application.StatusBar = "Today's Archive (DJMS " & _
    Format(config.Range("B35").Value, "0000") & _
    "-" & tradBuffLast & ") has been exported to " & selFLDR
GoTo terminalNoError

ExceptionPathway:
MsgBox "Run into some issue during saving, you may save the file later on a separately", vbOKOnly, "File Saving Interrupted"
Exit Sub

terminalNoError:
End Sub


Sub loadAdminTable() ' Save the Raw Total Table for admin
' nuke all long to 0
Dim i As Long
fsoRejectCUA = 0
fsoRecycleCUA = 0
fsoProcessedCUA = 0
fsoMemoCUA = 0
fsoMiscCUA = 0


For i = 1 To tradBuffLast ' CONSTRUCT TOTAL ADMIN COPY
    If (UCase(rTemp.Range("A" & i).Value) = "REJECT" Or UCase(rTemp.Range("A" & i).Value) = "REJ PART") Then
        rejHtmlAdmin = rejHtmlAdmin & tableHead & vbNewLine & "  " & _
                  rejTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & Format(rTemp.Range("B" & i).Value, "000000000") & vbNewLine & "  " & _
                  normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                  rejTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                  flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                  tableEnd & vbNewLine
        fsoRejectCUA = fsoRejectCUA + 1
    ElseIf UCase(rTemp.Range("A" & i).Value) = "RECYCLE" Then
        recHtmlAdmin = recHtmlAdmin & tableHead & vbNewLine & "  " & _
                  recTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & Format(rTemp.Range("B" & i).Value, "000000000") & vbNewLine & "  " & _
                  normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                  recTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                  flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                  tableEnd & vbNewLine
        fsoRecycleCUA = fsoRecycleCUA + 1
    ElseIf UCase(rTemp.Range("A" & i).Value) = "PROCESSED" Then ' PROCESSED
        processedHtmlAdmin = processedHtmlAdmin & tableHead & vbNewLine & "  " & _
                  memoTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & Format(rTemp.Range("B" & i).Value, "000000000") & vbNewLine & "  " & _
                  normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                  memoTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                  flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                  tableEnd & vbNewLine
        fsoProcessedCUA = fsoProcessedCUA + 1
    ElseIf UCase(rTemp.Range("A" & i).Value) = "MGMT NTC" Then ' MANAGEMENT NTC USE (U)
        memoHtmlAdmin = memoHtmlAdmin & tableHead & vbNewLine & "  " & _
                  memoTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & Format(rTemp.Range("B" & i).Value, "000000000") & vbNewLine & "  " & _
                  normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                  memoTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                  flatTd & rTemp.Range("U" & i).Value & normTdEnd & vbNewLine & _
                  tableEnd & vbNewLine
        fsoMemoCUA = fsoMemoCUA + 1
    Else ' anything else goes here
        urejHtmlAdmin = urejHtmlAdmin & tableHead & vbNewLine & "  " & _
                  UrejTdColor & rTemp.Range("A" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("D" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & Format(rTemp.Range("B" & i).Value, "000000000") & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("C" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("F" & i).Value & normTdEnd & vbNewLine & "  " & _
                  UrejTdColor & rTemp.Range("H" & i).Value & normTdEnd & vbNewLine & "  " & _
                  normTd & rTemp.Range("Q" & i).Value & normTdEnd & vbNewLine & "  " & _
                  flatTd & rTemp.Range("R" & i).Value & normTdEnd & vbNewLine & _
                  tableEnd & vbNewLine
        fsoMiscCUA = fsoMiscCUA + 1
    End If
Next i

End Sub

Sub reformEmptyTable() ' format the empty table imtp some kind of celebratory note
Dim i As Long
' (i) ALL COMPONENTS: Status  Transaction  mbrSSAN  mbrName  CycleID  Techn  errorCD  expln
'                     COL.A   COL.D        COL.B    COL.C    COL.F    COL.H  COL.Q    COL.R
'
' (i) Need memo specific: expln
'                         COL.U
'

If rejHtml = "" Then
    rejHtml = rejHtml & tableHead & vbNewLine & "  " & _
              rejTdColor & "-----" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "---------" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "--" & normTdEnd & vbNewLine & "  " & _
              rejTdColor & "No Rejects Today :)" & normTdEnd & vbNewLine & "  " & _
              normTd & "---" & normTdEnd & vbNewLine & "  " & _
              flatTd & "-------------------------------------------------" & normTdEnd & vbNewLine & _
              tableEnd & vbNewLine
    fsoRejectCU = 0
End If
If recHtml = "" Then
    recHtml = recHtml & tableHead & vbNewLine & "  " & _
              recTdColor & "-----" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "---------" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "--" & normTdEnd & vbNewLine & "  " & _
              recTdColor & "No Recycles Today :)" & normTdEnd & vbNewLine & "  " & _
              normTd & "---" & normTdEnd & vbNewLine & "  " & _
              flatTd & "-------------------------------------------------" & normTdEnd & vbNewLine & _
              tableEnd & vbNewLine
    fsoRecycleCU = 0
End If
If memoHtml = "" Then
    memoHtml = memoHtml & tableHead & vbNewLine & "  " & _
              memoTdColor & "-----" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "---------" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "--" & normTdEnd & vbNewLine & "  " & _
              memoTdColor & "No Management Notices Today :)" & normTdEnd & vbNewLine & "  " & _
              normTd & "---" & normTdEnd & vbNewLine & "  " & _
              flatTd & "-------------------------------------------------" & normTdEnd & vbNewLine & _
              tableEnd & vbNewLine
    fsoMemoCU = 0
End If
If urejHtml = "" Then
    urejHtml = urejHtml & tableHead & vbNewLine & "  " & _
              UrejTdColor & "-----" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "---------" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "--" & normTdEnd & vbNewLine & "  " & _
              UrejTdColor & "No Unassigned/Misc Item today :)" & normTdEnd & vbNewLine & "  " & _
              normTd & "---" & normTdEnd & vbNewLine & "  " & _
              flatTd & "-------------------------------------------------" & normTdEnd & vbNewLine & _
              tableEnd & vbNewLine
    fsoMiscCU = 0
End If
If processedHtml = "" Then
    processedHtml = processedHtml & tableHead & vbNewLine & "  " & _
              memoTdColor & "-----" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "---------" & normTdEnd & vbNewLine & "  " & _
              normTd & "----" & normTdEnd & vbNewLine & "  " & _
              normTd & "--" & normTdEnd & vbNewLine & "  " & _
              memoTdColor & "No Processed Item today :)" & normTdEnd & vbNewLine & "  " & _
              normTd & "---" & normTdEnd & vbNewLine & "  " & _
              flatTd & "-------------------------------------------------" & normTdEnd & vbNewLine & _
              tableEnd & vbNewLine
    fsoProcessedCU = 0
End If

End Sub

Private Sub mainExit_Click() ' LINK COMPLETE
utilityRejectReport.Hide
trackerAPI.Show
Application.StatusBar = False
End Sub

Private Sub mainExportOnly_Click()
exportOnly = True
    rejcReportCore
exportOnly = False
End Sub

Private Sub mainLaunchDMO_Click()
    CreateObject("Shell.Application").ShellExecute _
        "https://dmoapps.csd.disa.mil/WebDMO/Feedback.aspx"
End Sub


Private Sub rejcAmend_Click()
If rejcDeleteIO Then Exit Sub
protectEmail = False
addToGroup
iniConfig
End Sub

Private Sub rejcCCthis_Click() ' TOGGLE SELF CC
config.Range("B40").Value = rejcCCthis.Value
End Sub

Private Sub rejcColN_Change()
Dim Temp As String, icell As Range, deathCount As Integer
    deathCount = 0

If rejcDeleteIO Then Exit Sub
protectEmail = True
On Error GoTo skipErr
With rData
    For Each icell In rejcMember
        If icell.Value = rejcColN.Value Then
            rejcColRow.Value = Format(icell.Row, "00")
            GoTo resumeNext
        End If
        deathCount = deathCount + 1
        If deathCount > 20 Then GoTo skipErr
    Next icell
End With

skipErr:
rejcColRow.Value = ""
rejcAddIO = True

resumeNext:
rejcColR.Value = ""
updateGroup
protectEmail = False
rejcAddIO = False

If rejcColN.Value = "" Then ' kick if found the source block blank
    rejcAmend.Enabled = False
    rejcDelete.Enabled = False
    rejcColO.Value = False
    rejcColO.Caption = "MEMBER"
    rejcColP.Value = False
    rejcColP.Caption = "DISABLED"
    rejcColQ.Value = False
    rejcColQ.Caption = "DISABLED"
    rejcColR.Value = ""
    rejcColO.Enabled = False
    rejcColP.Enabled = False
    rejcColQ.Enabled = False
    rejcColR.Enabled = False
    Exit Sub
End If

End Sub

Private Sub rejcColO_Click()
If rejcDeleteIO Or rejcAddIO Then Exit Sub
protectEmail = True
editGroup
protectEmail = False
End Sub

Private Sub rejcColP_Click()
If rejcDeleteIO Or rejcAddIO Then Exit Sub
protectEmail = True
editGroup
protectEmail = False
End Sub

Private Sub rejcColQ_Click()
If rejcDeleteIO Or rejcAddIO Then Exit Sub
protectEmail = True
editGroup
protectEmail = False
End Sub

Private Sub rejcColR_Change()
If rejcDeleteIO Or rejcAddIO Then Exit Sub
If protectEmail Then Exit Sub
editGroup
End Sub

Private Sub rejcDelete_Click()
rejcDeleteIO = True
editGroup
rejcDeleteIO = False
End Sub

Private Sub rejcEmailOnBehalf_Change() ' change email sent on behalf of
config.Range("B39").Value = rejcEmailOnBehalf.Value
End Sub

Private Sub rejcPrintEnable_Click() ' print exclusive html copy?
If rejcPrintEnable Then ' YES, ENABLE ASSOCIATE PRINT FOR HTML
    config.Range("B38").Value = True
    rejcPrintEnable.Caption = "HTML MAKE ENABLED"
Else ' OR NOT
    config.Range("B38").Value = False
    rejcPrintEnable.Caption = "HTML MAKE DISABLED"
End If
globalSave
End Sub

Private Sub rejcPrintOptn_Click() ' APPOINT OR ASSIGN
If rejcPrintOptn Then ' ALLOW FIXED PATH
    config.Range("B36").Value = True
    rejcPrintOptn.Caption = "ASSIGN"
    rejcPrintPath.Enabled = True
    rejcPrintPathAssign.Enabled = True
    rejcPrintPathRemove.Enabled = True
Else ' OR NOT
    config.Range("B36").Value = False
    rejcPrintOptn.Caption = "MANUAL"
    rejcPrintPath.Enabled = False
    rejcPrintPathAssign.Enabled = False
    rejcPrintPathRemove.Enabled = False
End If
globalSave
End Sub

Private Sub rejcPrintPathAssign_Click() ' Assign the folder
Dim tempPath As String, tempFinder As FileDialog

Set tempFinder = Application.FileDialog(msoFileDialogFolderPicker)
    With tempFinder
        .Title = "Future Printouts will be exported to here"
        .ButtonName = "Assign"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo datawrite
        tempPath = .SelectedItems(1)
    End With
datawrite:
Set tempFinder = Nothing
config.Range("B37").Value = tempPath

iniConfig
globalSave
End Sub

Private Sub rejcPrintPathRemove_Click() ' NUKE ADDRESS
Dim resB As String
    resB = MsgBox("Remove Current Path?", vbYesNo, "Rejection Printouts Path Deletion")
If resB = vbNo Then Exit Sub
config.Range("B37").Value = ""
iniConfig
globalSave
End Sub

Private Sub rejcSourceOptn_Click() ' AUTO PULL FROM THIS
If rejcSourceOptn Then ' ALLOW FIXED PATH FOR AUTO PULL
    config.Range("B41").Value = True
    rejcSourceOptn.Caption = "ENABLE"
    rejcSourcePath.Enabled = True
    rejcSourcePathAssign.Enabled = True
    rejcSourcePathRemove.Enabled = True
Else ' OR NOT
    config.Range("B41").Value = False
    rejcSourceOptn.Caption = "DISABLE"
    rejcSourcePath.Enabled = False
    rejcSourcePathAssign.Enabled = False
    rejcSourcePathRemove.Enabled = False
End If
globalSave
End Sub


Private Sub rejcSourcePathAssign_Click() ' FOLDER ASSIGNMENT
Dim tempPath As String, tempFinder As FileDialog

Set tempFinder = Application.FileDialog(msoFileDialogFilePicker)
    With tempFinder
        .Title = "Assigning Auto Source Path..."
        .ButtonName = "Assign"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo datawrite
        tempPath = .SelectedItems(1)
    End With
datawrite:
Set tempFinder = Nothing
config.Range("B42").Value = tempPath

iniConfig
globalSave
End Sub

Private Sub rejcSourcePathRemove_Click() ' NUKE AUTO SOURCE ADDRESS

Dim resB As String
    resB = MsgBox("Remove Auto Source Target?", vbYesNo, "Rejection Auto Source Target Deletion")
If resB = vbNo Then Exit Sub
config.Range("B42").Value = ""
iniConfig
globalSave

End Sub

Private Sub rejcUnlockModification_Click()
If rejcUnlockModification.Caption = ChrW(&H26BF) Then
    rejcUnlockModification.Caption = ChrW(&H270D)
    rejcUser.BackColor = &HA0&
    rejcUser.Caption = "Receiving Opions - EDITING"
    mainExecute.Enabled = False
Else
    rejcUnlockModification.Caption = ChrW(&H26BF)
    rejcUser.BackColor = &H84496B
    rejcUser.Caption = "Receiving Opions"
    mainExecute.Enabled = True
    iniConfig ' RELOAD CONFIG
End If
updateGroup
End Sub

Private Sub UserForm_Initialize()
    
Set rTemp = ThisWorkbook.Sheets("DATA.TMP")
Set rData = ThisWorkbook.Sheets("SENSEI.DATA")
Set config = ThisWorkbook.Sheets("SENSEI.CONFIG")
Set ecsp = ThisWorkbook.Sheets("CSP.TR")
rejcUnlockModification.Caption = ChrW(&H26BF) ' UNLOCK BUTTON
Set rejcMember = rData.Range("rejcMemberTable[REJC MEMBER]") ' rejc table
Set rejcTech = New Scripting.Dictionary

protectEmail = True
rejcDeleteIO = False
rejcAddIO = False
REJrelease.Caption = config.Range("B34").Value
REJver.Caption = config.Range("B33").Value
mainVersion.Caption = "Sensei REJC " & config.Range("B33").Value & " on " & config.Range("D4").Value


iniConfig

End Sub
Sub iniConfig()
Dim icell As Range
Set techList = New Scripting.Dictionary ' re define the dictionary each time
Set rejcTech = New Scripting.Dictionary ' CONTACT LIST
Set techMemo = New Scripting.Dictionary ' memorandum list matching to person

bCClist = "" ' NO BCC INITIALLY

' ### insert disable sequence here

With config
    ' #### CONFIG - PRINT ####
    rejcPrintOptn.Value = .Range("B36").Value ' FALSE FOR MANUAL TRUE FOR FIXED POSITION
    rejcPrintPath.Value = .Range("B37").Value ' the storage path
    rejcPrintEnable.Value = .Range("B38").Value ' IS ASSOCIATED PRINT ENABLED (print to same place as pdf, if pdf existed)
    rejcSourceOptn.Value = .Range("B41").Value ' AUTO SOURCE MASTER IO
    rejcSourcePath.Value = .Range("B42").Value ' AUTO SOURCE TARGER PATH
    rejcEmailOnBehalf.Value = .Range("B39").Value ' Sending on Behalf of
    rejcCCthis.Value = .Range("B40").Value ' CC SELF?
End With
rejcColN.Clear
With rData ' load first field
    For Each icell In rejcMember
        If icell.Value <> "" And icell.Value <> "##" Then ' prevent ## from entering
            rejcColN.AddItem icell.Value
            
            ' LOAD THE TABLE TEST FOR ADMIN AS WELL AS EMAIL CATALOG
            ' Block ##, and those who do not need an email
            If (Not techList.Exists(icell.Value)) And rData.Range("P" & icell.Row).Value Then
                If rData.Range("O" & icell.Row).Value = False Then
                    techList.Add UCase(icell.Value), 0 ' 0 as member
                ElseIf rData.Range("O" & icell.Row).Value = True Then
                    techList.Add UCase(icell.Value), 1 ' 1 AS ADMIN
                End If
                If rData.Range("R" & icell.Row).Value <> "" Then _
                    rejcTech.Add UCase(icell.Value), rData.Range("R" & icell.Row).Value
            End If
            
            ' LOAD BCC LIST
            If rData.Range("Q" & icell.Row).Value = True Then
                bCClist = bCClist & rData.Range("R" & icell.Row).Value
            End If
            
        End If
    Next icell
End With
    rejcAmend.Enabled = False
    rejcDelete.Enabled = False

End Sub

Sub updateGroup() ' ENABLE OR DISABLE THE GROUP + UPDATE ROW INFO
Dim tRow As Long
' ADD. DELETE FUNCTION JACKED UP

If rejcAddIO Then ' KICK OUT IF FOUND OUT WE ARE AMENDING >> CONFIG ERROR EXIST, TOGGLE DID NOT UPDATE
    rejcAmend.Enabled = True
    rejcDelete.Enabled = False
    rejcColO.Value = False
    rejcColO.Caption = "MEMBER"
    rejcColP.Value = False
    rejcColP.Caption = "DISABLED"
    rejcColQ.Value = False
    rejcColQ.Caption = "DISABLED"
    rejcColR.Value = ""
    Exit Sub
End If

If rejcUnlockModification.Caption = ChrW(&H270D) Then ' MASTER UNLOCK
    rejcColN.Enabled = True
    rejcColO.Enabled = True
    rejcColP.Enabled = True
    rejcColQ.Enabled = True
    rejcColR.Enabled = True
    rejcColRow.Enabled = True
    rejcAmend.Enabled = True
    rejcDelete.Enabled = True
Else
    rejcColN.Enabled = False
    rejcColO.Enabled = False
    rejcColP.Enabled = False
    rejcColQ.Enabled = False
    rejcColR.Enabled = False
    rejcColRow.Enabled = False
    rejcAmend.Enabled = False
    rejcDelete.Enabled = False
    Exit Sub
End If

With rData ' ENTRY LOADING FX AND PRE DETERMINING FX
    
    If rejcColO.Value Then
        rejcColO.Caption = "ADMIN"
    Else
        rejcColO.Caption = "MEMBER"
    End If
    If rejcColP.Value Then
        rejcColP.Caption = "ENABLED"
    Else
        rejcColP.Caption = "DISABLED"
    End If
    If rejcColQ.Value Then
        rejcColQ.Caption = "ENABLED"
    Else
        rejcColQ.Caption = "DISABLED"
    End If
    
    If rejcColN.Value = "##" Then
        rejcAmend.Enabled = False
        rejcDelete.Enabled = True
    End If
    
    If rejcColRow.Value <> "" Then
        If rejcColRow.Value < 2 Or rejcColRow.Value > 21 Then
            rejcColRow.Value = ""
            rejcAmend.Enabled = True
            rejcDelete.Enabled = False
            Exit Sub
        End If
        rejcAmend.Enabled = False
        rejcDelete.Enabled = True
        tRow = rejcColRow.Value
        If .Range("O" & tRow).Value = True Then
            rejcColO.Value = True
            rejcColO.Caption = "ADMIN"
        Else
            rejcColO.Value = False
            rejcColO.Caption = "MEMBER"
        End If
        If .Range("P" & tRow).Value = True Then
            rejcColP.Value = True
            rejcColP.Caption = "ENABLED"
        Else
            rejcColP.Value = False
            rejcColP.Caption = "DISABLED"
        End If
        If .Range("Q" & tRow).Value = True Then
            rejcColQ.Value = True
            rejcColQ.Caption = "ENABLED"
        Else
            rejcColQ.Value = False
            rejcColQ.Caption = "DISABLED"
        End If
        rejcColR.Value = Replace(.Range("R" & tRow).Value, ";", "")
    Else
    End If
End With


End Sub
Sub editGroup() ' editing the group
Dim tRow As Long
If rejcColRow.Value = "" Then Exit Sub

tRow = rejcColRow.Value

With rData
    .Range("O" & tRow).Value = rejcColO.Value
    .Range("P" & tRow).Value = rejcColP.Value
    .Range("Q" & tRow).Value = rejcColQ.Value
    If Not protectEmail Then .Range("R" & tRow).Value = rejcColR.Value & ";"
    
    If rejcDeleteIO Then '(NUKE)
        .Range("N" & tRow).Value = "##"
        .Range("O" & tRow).Value = False
        .Range("P" & tRow).Value = False
        .Range("Q" & tRow).Value = False
        .Range("R" & tRow).Value = ""
        rejcColN.Value = "##"
        rejcColO.Value = False
        rejcColP.Value = False
        rejcColQ.Value = False
        rejcColR.Value = "##"
        rejcColRow.Value = ""
    End If
    
End With

' update id
    If rejcColO.Value Then
        rejcColO.Caption = "ADMIN"
    Else
        rejcColO.Caption = "MEMBER"
    End If
    If rejcColP.Value Then
        rejcColP.Caption = "ENABLED"
    Else
        rejcColP.Caption = "DISABLED"
    End If
    If rejcColQ.Value Then
        rejcColQ.Caption = "ENABLED"
    Else
        rejcColQ.Caption = "DISABLED"
    End If

End Sub

Sub addToGroup() ' add, if spaceous, or warn if no space
Dim icell As Range, thisR As Long
Dim ct As Long, ks As Long
ct = 0
ks = 20
If Len(rejcColN.Value) <> 2 Then
    MsgBox "Initial must be 2 character in Length!", vbOKOnly, "Sensei REJC Notice"
    Exit Sub
End If

With rData ' load first empty field
    For Each icell In rejcMember
        If icell.Value = "##" Then
            thisR = icell.Row
            Exit For
        End If
        ct = ct + 1
        If ct > ks Then ' reached max exit total
            MsgBox "It appears there are no more empty space on the List" & vbNewLine & _
                   vbNewLine & _
                   "Please consider remove or edit the existing list", _
                   vbOKOnly, "Sensei REJC Email List Capacity Reached!"
            Exit Sub
        End If
    Next icell
End With

With rData
    .Range("N" & thisR).Value = rejcColN.Value
    .Range("O" & thisR).Value = rejcColO.Value
    .Range("P" & thisR).Value = rejcColP.Value
    .Range("Q" & thisR).Value = rejcColQ.Value
    If Not protectEmail Then .Range("R" & thisR).Value = rejcColR.Value & ";"
End With

End Sub

