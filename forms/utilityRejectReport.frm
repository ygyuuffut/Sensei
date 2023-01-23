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
Public sData As Worksheet, rRpt As Worksheet, rTemp As Worksheet
Private Sub mainExecute_Click() ' boot reject report locally
Dim txtRpt As String, txtCov ' file name for import and variant for reduced friction.
Dim filePk As FileDialog: Set filePk = Application.FileDialog _
                                           (msoFileDialogFilePicker) ' Bind it
Dim tradBuffLast As Long ' for standard report making

' >>> Load the file
' if <realized we have a given location> and <known file name.txt> then
'   Pull it automatically based on path and name
' else
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
End With

' >>> Check Pre-requisites
' If <Prior update number is not same as today> then
'   Update the <Made> to today, <Update #> to today's note
' Else
'   Notify this update number is done already
'   Wipe rTemp and Exit Sub
' end if

' >>> Write Traditional Full Sheet
rTemp.Range("A1:F" & tradBuffLast).Copy ' Current Status, SSN NAME, TRANS, DESC, CYCLE
    rRpt.Range("C2").PasteSpecial xlPasteValues
rTemp.Range("H1:H" & tradBuffLast).Copy ' THE RESPONSEE
    rRpt.Range("I2").PasteSpecial xlPasteValues
rTemp.Range("J1:J" & tradBuffLast).Copy ' ADSN
    rRpt.Range("J2").PasteSpecial xlPasteValues
rTemp.Range("M1:U" & tradBuffLast).Copy ' POST DATE, RECEIPT, UPDATE NUMBER, DAY CT, ERR, ERR REASON, ERR ADDITIONAL, CARD DATA
    rRpt.Range("K2").PasteSpecial xlPasteValues

' >>> SIEVE THE INFORMATION
' Go through the iteration loop and Mark all the transactions with appropriate actions
' Append those require actions to arrays

' >>> Finalize information
' Export the traditional Report with AC date and Make date
' Wipe Traditional Report and BUFFER

' >>> Save and Send the Email

MsgBox tradBuffLast
ecsp.Activate ' return
End Sub

Private Sub mainLaunchDMO_Click()
    CreateObject("Shell.Application").ShellExecute _
        "microsoft-edge:https://dmoapps.csd.disa.mil/WebDMO/Login.aspx"
End Sub

Private Sub UserForm_Initialize()
    
Set rTemp = ThisWorkbook.Sheets("DATA.TMP")
    
End Sub


