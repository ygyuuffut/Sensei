Attribute VB_Name = "expdFX_Deployment"
Option Explicit
Const scantronStart = 2

Public Sub generateScantron() ' Generate Scan tron for Master 114
' iterate document first to elimate duplicate and wrong format
Dim depIO As Worksheet: Set depIO = ThisWorkbook.Sheets("DEP.IO")
Dim rData As Workbook, sh1 As Worksheet, sh2 As Worksheet, _
    sh3 As Worksheet, sh4 As Worksheet, shx As Worksheet ' raw data and worksheets
Dim config As Worksheet: Set config = ThisWorkbook.Sheets("SENSEI.CONFIG")
Dim dOption As Boolean: dOption = config.Range("J4").Value ' not use omega date?

Dim baseLine As Long, ssnBegin As Long ' last row for raw data, SSN START LINE
Dim iv As Long, ix As Long ' used to loop thru worksheets
Dim flag14 As Boolean, flagFL As Boolean, flag23 As Boolean, flag65 As Boolean
    flag14 = False: flagFL = False: flag23 = False: flag65 = False ' reset flags
Dim is14 As Boolean, isFL As Boolean, is23 As Boolean, is65 As Boolean, isDO As Boolean
    is14 = False: isFL = False: is23 = False: is65 = False: isDO = False ' reset flags for confirmation
'Dim aSSN() As Variant, aNmn() As Variant, _
'    aeFL() As Variant, ae14() As Variant, ae23() As Variant, ae65() As Variant ' person, ssn, entitlements *4, then Dates

Dim ssnDict As Dictionary: Set ssnDict = New Scripting.Dictionary ' SSN to name
Dim ssnDictFL As Dictionary: Set ssnDictFL = New Scripting.Dictionary ' SSN to FL
Dim ssnDict14 As Dictionary: Set ssnDict14 = New Scripting.Dictionary ' SSN to 14
Dim ssnDict23 As Dictionary: Set ssnDict23 = New Scripting.Dictionary ' SSN to 23
Dim ssnDict65 As Dictionary: Set ssnDict65 = New Scripting.Dictionary ' SSN to 65
Dim tempA() As Variant, ib As Long: ib = 0 ' dictionary Index Transpose

' Day Count Lib, w, x, y, z according to FL, 14, 23, 65 Micro adjusting date
Dim wDate: wDate = Format(Now() - config.Range("J6").Value, "YYYY-MM-DD")
Dim xDate: xDate = Format(Now() - config.Range("J7").Value, "YYYY-MM-DD")
Dim yDate: yDate = Format(Now() - config.Range("J8").Value, "YYYY-MM-DD")
Dim zDate: zDate = Format(Now() - config.Range("J9").Value, "YYYY-MM-DD")
Dim omegaDate: omegaDate = Format(Now() - config.Range("J5").Value, "YYYY-MM-DD") ' TODAY IN yyyy-mm-dd < Using

' date separator for 65
Dim zx, zy

Dim ssnL ' temporary SSN
Dim rawdata: rawdata = Application.GetOpenFilename("Pay Entitlements~ .xls (*.xls; *.xlsx),*.xls;*.xlsx", , "Select Raw Deployment Data", "Generate", False)
Application.ScreenUpdating = False
    If rawdata = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If


Set rData = Workbooks.Open(rawdata) ' appoint data card and find first sheet's last row
    Do While Not rData.Sheets(1).Range("C3").Value Like "Pay Entitlements" ' not the right one? try again!
        rData.Close False
        rawdata = Application.GetOpenFilename("Pay Entitlements~ .xls (*.xls; *.xlsx),*.xls;*.xlsx", , "re-Select Raw Deployment Data", "Generate", False)
        If rawdata = False Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        Set rData = Workbooks.Open(rawdata)
    Loop

baseLine = depIO.Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
depIO.Range("A2:L" & baseLine).ClearContents ' get rid of old information

For iv = 1 To 4 ' go thru all 4 sheets and build the reference table
    ' KNOW WHERE TO START AND STOP
    baseLine = rData.Sheets(iv).Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
    ssnBegin = rData.Sheets(iv).Cells.Find("SSN", SEARCHORDER:=xlByRows, searchDirection:=xlAfter).Row + 1
    Set shx = rData.Sheets(iv) ' lock into variable
    
    With shx.Columns("B:B") ' Deformat SSN back to number
        .Value = .Value
        .NumberFormat = "000000000"
    End With
    

    For ix = ssnBegin To baseLine ' RECORD ALL PEOPLE'S NAME JUST INCASE
        ssnL = shx.Range("B" & ix).Value
        If ssnL <> "" Then
            If ssnDict.Exists(ssnL) Then
            Else
                ssnDict.Add ssnL, shx.Range("G" & ix).Value
            End If
        End If
    Next ix

Next iv

' Second iteration to physically record what is missing
For iv = 1 To 4
    ' KNOW WHERE TO START AND STOP
    baseLine = rData.Sheets(iv).Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
    ssnBegin = rData.Sheets(iv).Cells.Find("SSN", SEARCHORDER:=xlByRows, searchDirection:=xlAfter).Row + 1
    Set shx = rData.Sheets(iv) ' lock into variable
    
    ' Ensure the formatting based on entitlements (What are we looking at?)
    If shx.Range("B5").Value <> "" And shx.Range("B4").Value = "" Then ' if title is on B5, check B5
        If shx.Range("B5").Value Like "*Combat Zone Tax*" Then
            flagFL = True
        ElseIf shx.Range("B5").Value Like "*Family Separation*" Then
            flag65 = True
        ElseIf shx.Range("B5").Value Like "*Hardship Duty*" Then
            flag14 = True
        ElseIf shx.Range("B5").Value Like "*Hostile Fire*" Then
            flag23 = True
        End If
    Else ' else check B4
        If shx.Range("B4").Value Like "*Combat Zone Tax*" Then
            flagFL = True
        ElseIf shx.Range("B4").Value Like "*Family Separation*" Then
            flag65 = True
        ElseIf shx.Range("B4").Value Like "*Hardship Duty*" Then
            flag14 = True
        ElseIf shx.Range("B4").Value Like "*Hostile Fire*" Then
            flag23 = True
        End If
    End If
    
    For ix = ssnBegin To baseLine ' Now match record, if found then wrote entry at current row
        ssnL = shx.Range("B" & ix).Value
        If ssnL <> "" Then
            If ssnDict.Exists(ssnL) Then ' detect and write
            
                If flagFL And Not dOption Then ' do fl IF OLDER THAN or euqal omegaDate
                    If Not Format(shx.Range("L" & ix).Value, "YYYY-MM-DD") > omegaDate Then
                        If ssnDictFL.Exists(ssnL) Then ' see if FL dictionary has entry
                        Else ' LOAD IT TO FL DICTIONARY
                            ssnDictFL.Add ssnL, "X"
                        End If
                    End If
                End If
                If flagFL And dOption Then ' do fl IF OLDER THAN or euqal wDate
                    If Not Format(shx.Range("L" & ix).Value, "YYYY-MM-DD") > wDate Then
                        If ssnDictFL.Exists(ssnL) Then ' see if FL dictionary has entry
                        Else ' LOAD IT TO FL DICTIONARY
                            ssnDictFL.Add ssnL, "X"
                        End If
                    End If
                End If
                
                If flag14 And Not dOption Then ' do 14 IF OLDER THAN or euqal omegaDate
                    If Not Format(shx.Range("L" & ix).Value, "YYYY-MM-DD") > omegaDate Then
                        If ssnDict14.Exists(ssnL) Then ' see if 14 dictionary has entry
                        Else ' LOAD IT TO 14 DICTIONARY
                            ssnDict14.Add ssnL, "X"
                        End If
                    End If
                End If
                If flag14 And dOption Then ' do 14 IF OLDER THAN or euqal xDate
                    If Not Format(shx.Range("L" & ix).Value, "YYYY-MM-DD") > xDate Then
                        If ssnDict14.Exists(ssnL) Then ' see if 14 dictionary has entry
                        Else ' LOAD IT TO 14 DICTIONARY
                            ssnDict14.Add ssnL, "X"
                        End If
                    End If
                End If
                
                
                If flag23 And Not dOption Then ' do 23 IF OLDER THAN or euqal omegaDate
                    If Not Format(shx.Range("R" & ix).Value, "YYYY-MM-DD") > omegaDate Then
                        If ssnDict23.Exists(ssnL) Then ' see if 23 dictionary has entry
                        Else ' LOAD IT TO 23 DICTIONARY
                            ssnDict23.Add ssnL, "X"
                        End If
                    End If
                End If
                If flag23 And dOption Then ' do 23 IF OLDER THAN or euqal yDate
                    If Not Format(shx.Range("R" & ix).Value, "YYYY-MM-DD") > yDate Then
                        If ssnDict23.Exists(ssnL) Then ' see if 23 dictionary has entry
                        Else ' LOAD IT TO 23 DICTIONARY
                            ssnDict23.Add ssnL, "X"
                        End If
                    End If
                End If
                
                If flag65 And Not dOption Then ' do 65 IF OLDER THAN or euqal omegaDate
                    zx = shx.Range("O" & ix).Value ' RECOMPOSE DATE
                    zy = Format(DateSerial(CInt(Left(zx, 2)), CInt(Mid(zx, 3, 2)), CInt(Right(zx, 2))), "YYYY-MM-DD")
                    If Not zy > omegaDate Then
                        If ssnDict65.Exists(ssnL) Then ' see if 65 dictionary has entry
                        Else ' LOAD IT TO 65 DICTIONARY
                            ssnDict65.Add ssnL, "X"
                        End If
                    End If
                End If
                If flag65 And dOption Then ' do 65 IF OLDER THAN or euqal zDate
                    zx = shx.Range("O" & ix).Value ' RECOMPOSE DATE
                    zy = Format(DateSerial(CInt(Left(zx, 2)), CInt(Mid(zx, 3, 2)), CInt(Right(zx, 2))), "YYYY-MM-DD")
                    If Not zy > zDate Then
                        If ssnDict65.Exists(ssnL) Then ' see if 65 dictionary has entry
                        Else ' LOAD IT TO 65 DICTIONARY
                            ssnDict65.Add ssnL, "X"
                        End If
                    End If
                End If
                
                
            Else
            End If
        End If
    Next ix ' END OF INSHEET LOOP
    
    'reset AFTER SHEET COMPLETE
    flag14 = False: flagFL = False: flag23 = False: flag65 = False

Next iv



' after wrote all dictionaries, go through all entries in main dictionary
' 3rd iteration to sieve all that are in creteria, make counts
Dim ba, bb, ecc As Long, cSSN, cxx As Long: cxx = 0
ecc = ssnDict.Count - 1
cSSN = ssnDict.Keys
For ix = 0 To ecc ' Now match record, if found then wrote entry at current row
    ' elementary finding of what is on
    If ssnDictFL.Exists(cSSN(ix)) Then ' found fl
        isFL = True
    End If
    If ssnDict14.Exists(cSSN(ix)) Then ' found 14
        is14 = True
    End If
    If ssnDict23.Exists(cSSN(ix)) Then ' found 23
        is23 = True
    End If
    If ssnDict65.Exists(cSSN(ix)) Then ' found 65
        is65 = True
    End If
    
    ' primary sieve before writing into tempA
    If is23 Then isFL = False
    If isFL Or is14 Or is23 Or is65 Then isDO = True ' if has any open, mark action flag
    cxx = cxx + 1 ' confirm numbers

    is14 = False: isFL = False: is23 = False: is65 = False: isDO = False ' reset all action flags
Next ix



' 4 th iteration, remove blank spaces, use count to construct solid blocks
Dim acx As Long: acx = 0
Dim strN: strN = ssnDict.Items
ReDim Preserve tempA(cxx, 5) ' finalize dictionary matrix size (6 column from 0-5)
For ix = 0 To ecc ' Now match record, if found then wrote entry at current row
    ' elementary finding of what is on
    If ssnDictFL.Exists(cSSN(ix)) Then ' found fl
        isFL = True
    End If
    If ssnDict14.Exists(cSSN(ix)) Then ' found 14
        is14 = True
    End If
    If ssnDict23.Exists(cSSN(ix)) Then ' found 23
        is23 = True
    End If
    If ssnDict65.Exists(cSSN(ix)) Then ' found 65
        is65 = True
    End If
    
    ' primary sieve before writing into tempA
    If is23 Then isFL = False
    If isFL Or is14 Or is23 Or is65 Then isDO = True ' if has any open, mark action flag
    
    ' write element ix to tempA
    If isDO Then ' contain action flag
        tempA(acx, 0) = cSSN(ix) ' assign ssn to column 0
        tempA(acx, 1) = strN(ix) ' assign name to column 1
        If isFL Then
            tempA(acx, 2) = "X"
        Else
            tempA(acx, 2) = ""
        End If
        If is14 Then
            tempA(acx, 3) = "X"
        Else
            tempA(acx, 3) = ""
        End If
        If is23 Then
            tempA(acx, 4) = "X"
        Else
            tempA(acx, 4) = ""
        End If
        If is65 Then
            tempA(acx, 5) = "X"
        Else
            tempA(acx, 5) = ""
        End If
        acx = acx + 1
    End If
    is14 = False: isFL = False: is23 = False: is65 = False: isDO = False ' reset all action flags
Next ix

rData.Close False

' wrote to actual scantron for future purposes
For ba = 0 To ecc

    For bb = 0 To 5
        depIO.Cells(ba + 2, bb + 1).Value = tempA(ba, bb)
    Next bb

Next ba

' CLEAN THE REPORT UP and formulate it
depIO.Columns("A:A").NumberFormat = "000000000"
With depIO
    baseLine = .Cells.Find("*", SEARCHORDER:=xlByRows, searchDirection:=xlPrevious).Row
    .Range("I2").Formula = "=IF(AND($K2<>"""",$F2<>""""),$H2-1,"""")"
    .Range("J2").Formula = "=IF(AND($K2<>"""",$C2<>""""),EOMONTH($G2,0),"""")"
    .Range("I2:I" & baseLine).FillDown
    .Range("J2:J" & baseLine).FillDown
    .Range("A1").Value = "SSN"
    .Range("B1").Value = "NAME"
    .Range("C1").Value = "FL"
    .Range("D1").Value = "14"
    .Range("E1").Value = "23"
    .Range("F1").Value = "65"
    .Range("G1").Value = "DEPART AREA"
    .Range("H1").Value = "ARRIVE HERE"
    .Range("I1").Value = "FSA STOP"
    .Range("J1").Value = "TAX STOP"
    .Range("K1").Value = "ACTION"
    .Range("L1").Value = "STATUS"
End With


Application.ScreenUpdating = True

End Sub
Public Function matchKey(sourceDictionary As Dictionary, matchString As String, Optional Include As Boolean = True, Optional Compare As VbCompareMethod = vbTextCompare) As String() ' A lookup Function for Dictionary, soucrced from Stack Overflow
    matchKey = Filter(sourceDictionary.Keys, matchString, Include, Compare)
End Function

Sub enumLatest_Table118() ' Enumerate and draw Latest item from TB 118
' Find the last row
' get the source range
' from top to bottom, receive a value
' compare ID with destination (set to col "g" @ g2)
' if unique ID, log all information, else compare the date and get the newer one
' next value from source

'###############################################
' Make necessary Parameter adjustment as Needed
'###############################################

Application.ScreenUpdating = False
Dim Lrow As Long, LdRow As Long, Crow As Long, destOffset As Long, Drow As Long
'   source last, destinat. last, current srce, destination offset, current dest
Dim HDPsht As Worksheet, srcCell As Range, dstCell As Range
'   worksheet,           source loop,      destination
    Set HDPsht = Worksheets("118-HDP")
    Lrow = HDPsht.Range("D" & Rows.Count).End(xlUp).Row
    destOffset = 2
With HDPsht
    .Range("G1").Value = "Country ID"
    .Range("H1").Value = "Location ID"
    .Range("I1").Value = "Detail"
    .Range("J1").Value = "Amount"
    .Range("K1").Value = "Effective Since"
End With

HDPsht.Range("D:D").NumberFormat = "YYYY-MM-DD"
For Each srcCell In Range("D2:D" & Lrow)
With HDPsht
    Crow = srcCell.Row
    .Range("D" & Crow).Value = .Range("D" & Crow).Value
    For Each dstCell In .Range("G2:G" & destOffset)
        Drow = dstCell.Row
        If .Range("C" & Crow).Value = 0 Then Exit For ' $0 ENTITLMENT EXIT DIRECTLY
        If .Range("A" & Crow).Value <> .Range("H" & Drow).Value And .Range("G" & Drow).Value = "" Then
            .Range("G" & Drow).Value = Left(.Range("A" & Crow).Value, 2) ' ID
            .Range("H" & Drow).Value = .Range("A" & Crow).Value ' CODE
            .Range("I" & Drow).Value = .Range("B" & Crow).Value ' Location
            .Range("J" & Drow).Value = .Range("C" & Crow).Value ' Amount
            .Range("K" & Drow).Value = .Range("D" & Crow).Value ' DATE
            destOffset = destOffset + 1
            Exit For
        End If
        If .Range("A" & Crow).Value = .Range("H" & Drow).Value And .Range("D" & Crow).Value > .Range("J" & Drow).Value Then
            .Range("G" & Drow).Value = Left(.Range("A" & Crow).Value, 2) ' ID
            .Range("H" & Drow).Value = .Range("A" & Crow).Value ' CODE
            .Range("I" & Drow).Value = .Range("B" & Crow).Value ' Location
            .Range("J" & Drow).Value = .Range("C" & Crow).Value ' amount
            .Range("K" & Drow).Value = .Range("D" & Crow).Value ' date
            destOffset = destOffset + 1
            Exit For
        End If
    Next dstCell
'        .Range("E" & Crow) = .Range("D" & Crow).Value
End With
Next srcCell

Application.ScreenUpdating = True
End Sub

Sub enumLatest_Table054()
' Find the Last row
' get source range
' iterate through source range
' compare date against today (Country ID will be placed on G2 and onward)
' if greater than today, write this


'###############################################
' Make necessary Parameter adjustment as Needed
'###############################################

Application.ScreenUpdating = False
Dim Crow As Long, Drow As Long ' Used for simplifying Row
Dim Lrow As Long, LdRow As Long, Crdate As String, Ddate As Date, Doffset As Long
' Last row source.Last row dest. Date on Source,   Today's date, Offest on destina.
Dim HFPsht As Worksheet, srcCell As Range, dstCell As Range
' Worksheet              source sel Cell,  destin sel Cell
    Set HFPsht = Worksheets("054-HFP") ' Alt this as needed
    Lrow = HFPsht.Range("A" & Rows.Count).End(xlUp).Row
    Doffset = 2
With HFPsht
    .Range("G1").Value = "Country ID"
    .Range("H1").Value = "Country"
    .Range("I1").Value = "Effective Till"
End With
HFPsht.Range("D:D").NumberFormat = "YYYY-MM-DD"
Ddate = Format(Now(), "YYYY-MM-DD")

With HFPsht
For Each srcCell In .Range("A2:A" & Lrow)
    Crow = srcCell.Row
    .Range("D" & Crow).Value = .Range("D" & Crow).Value
    If Ddate < Format(.Range("D" & Crow).Value, "YYYY-MM-DD") Then
        .Range("G" & Doffset).Value = Left(.Range("B" & Crow).Value, 2)
        If Mid(.Range("B" & Crow).Value, 6, 80) <> "" Then
            .Range("H" & Doffset).Value = Mid(.Range("B" & Crow).Value, 6, 80)
        Else
            .Range("H" & Doffset).Value = "Location Code - " & Left(.Range("B" & Crow).Value, 2)
        End If
        .Range("I" & Doffset).Value = .Range("D" & Crow).Value
        Doffset = Doffset + 1
    End If
Next srcCell

Application.ScreenUpdating = True
End With

End Sub

Sub enumLatest_Table154()
' Find the Last row
' get source range
' iterate through source range
' compare date against today if is 1/1/0001 then record it
' if greater than today, write this

'###############################################
' Make necessary Parameter adjustment as Needed
'###############################################

Application.ScreenUpdating = False
Dim Crow As Long, Drow As Long ' Minimal Row
Dim Lrow As Long, LdRow As Long, Doffset As Long
Dim CZTEsht As Worksheet, srcCell As Range, dstCell As Range
    Set CZTEsht = Worksheets("154-CZTE")
    Lrow = CZTEsht.Range("A" & Rows.Count).End(xlUp).Row
    Doffset = 2
With CZTEsht
    .Range("G1").Value = "Country ID"
    .Range("H1").Value = "Is Effective?"
    .Range("D:D").NumberFormat = "YYYY-MM-DD"
    For Each srcCell In .Range("A2:A" & Lrow)
        Crow = srcCell.Row
        .Range("D" & Crow).Value = .Range("D" & Crow).Value
        If Range("D" & Crow).Value = "1/1/0001" Then
            .Range("G" & Doffset).Value = .Range("B" & Crow).Value
            .Range("H" & Doffset).Value = "Yes"
            Doffset = Doffset + 1
        End If
    Next srcCell
End With
Application.ScreenUpdating = True
End Sub

