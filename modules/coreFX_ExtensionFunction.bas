Attribute VB_Name = "coreFX_ExtensionFunction"
Option Explicit
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As LongPtr) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr

Public Sub SetClipboard(sUniText As String) ' Migrated From Microsoft Support
    Dim iStrPtr As LongPtr
    Dim iLen As LongPtr
    Dim iLock As LongPtr
    Const GMEM_MOVEABLE = &H2
    Const GMEM_ZEROINIT = &H40
    Const CF_UNICODETEXT = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub


Public Function GetClipboard() As String ' Migrated From Microsoft Support
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function
Public Function indexDict(sourceDictionary As Dictionary, targetStr As String, Optional Include As Boolean = True, Optional Compare As VbCompareMethod = vbTextCompare) As String() ' compare in the dictionary

    indexDict = Filter(sourceDictionary.Keys, targetStr, Include, Compare)
    ' supposively finds by partial match
    
End Function

Function SpellDollar(ByVal numIn) ' 2424 Dollar speller
On Error GoTo reducedExit
    
    Dim LSide, RSide, Temp, DecPlace, Count, oNum
    oNum = numIn
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    numIn = Trim(Str(numIn)) 'String representation of amount
    ' Edit 2.(0)/Internationalisation
    ' Don't change point sign here as the above assignment preserves the point!
    DecPlace = InStr(numIn, ".") 'Pos of dec place 0 if none
    If DecPlace > 0 Then 'Convert Right & set numIn
        RSide = GetTens(Left(Mid(numIn, DecPlace + 1) & "00", 2))
        numIn = Trim(Left(numIn, DecPlace - 1))
    End If
    RSide = numIn
    Count = 1
    Do While numIn <> ""
        Temp = GetHundreds(Right(numIn, 3))
        If Temp <> "" Then LSide = Temp & Place(Count) & LSide
        If Len(numIn) > 3 Then
            numIn = Left(numIn, Len(numIn) - 3)
        Else
            numIn = ""
        End If
        Count = Count + 1
    Loop

    SpellDollar = LSide
    If InStr(oNum, Application.DecimalSeparator) > 0 Then    ' << Edit 2.(1)
        SpellDollar = SpellDollar & " Dollars and " & fractionWords(oNum) & " Cents"
    End If
    SpellDollar = Replace(SpellDollar, "  ", " ") ' nuke extra spaces
    If Right(SpellDollar, 9) = "and Cents" Then SpellDollar = Replace(SpellDollar, "and Cents", "and Zero Cents")
    If Left(SpellDollar, 8) = " Dollars" Then SpellDollar = Replace(SpellDollar, " Dollars", "Zero Dollars")

reducedExit:
End Function

Function GetHundreds(ByVal numIn) 'Converts a number from 100-999 into text
    Dim w As String
    If Val(numIn) = 0 Then Exit Function
    numIn = Right("000" & numIn, 3)
    If Mid(numIn, 1, 1) <> "0" Then 'Convert hundreds place
        w = GetDigit(Mid(numIn, 1, 1)) & " Hundred "
    End If
    If Mid(numIn, 2, 1) <> "0" Then 'Convert tens and ones place
        w = w & GetTens(Mid(numIn, 2))
    Else
        w = w & GetDigit(Mid(numIn, 3))
    End If
    GetHundreds = w
End Function

Function GetTens(TensText)  'Converts a number from 10 to 99 into text
    Dim w As String
    w = ""           'Null out the temporary function value
    If Val(Left(TensText, 1)) = 1 Then   'If value between 10-19
        Select Case Val(TensText)
            Case 10: w = "Ten"
            Case 11: w = "Eleven"
            Case 12: w = "Twelve"
            Case 13: w = "Thirteen"
            Case 14: w = "Fourteen"
            Case 15: w = "Fifteen"
            Case 16: w = "Sixteen"
            Case 17: w = "Seventeen"
            Case 18: w = "Eighteen"
            Case 19: w = "Nineteen"
            Case Else
        End Select
    Else      'If value between 20-99..
        Select Case Val(Left(TensText, 1))
            Case 2: w = "Twenty "
            Case 3: w = "Thirty "
            Case 4: w = "Forty "
            Case 5: w = "Fifty "
            Case 6: w = "Sixty "
            Case 7: w = "Seventy "
            Case 8: w = "Eighty "
            Case 9: w = "Ninety "
            Case Else
        End Select
        w = w & GetDigit _
            (Right(TensText, 1))  'Retrieve ones place
    End If
    GetTens = w
End Function

Function GetDigit(Digit) 'Converts a number from 1 to 9 into text
    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

Function fractionWords(n) As String ' counts decimals
    Dim fraction As String, x As Long
    fraction = Split(n, Application.DecimalSeparator)(1)   ' << Edit 2.(2)
    For x = 1 To Len(fraction)
        If fractionWords <> "" Then fractionWords = fractionWords & " "
        If Len(fraction) > 2 Then fractionWords = fractionWords & GetDigit(Mid(fraction, x, 1))
    Next x
    If Len(fraction) = 2 Then
        fractionWords = GetTens(fraction)
    End If
End Function

Public Function tableReject(nStatus As String, nRequest As String, nSSN As String, nid As String, nCycle As String, nTech As String, nRej As String, nRejReason As String)
    Dim htmlRejTable As String: htmlRejTable = ""
    
' >>> Go through the pre-fab Arrays and wrote the whole table _
      then out put the string formatted Html raw code with some kind of style

End Function

Public Function isInMonoArray(searchString As String, array1 As Variant) As Long
Dim i As Long
isInMonoArray = -1 ' error message

    For i = LBound(array1) To UBound(array1)
        If StrComp(searchString, array1(i), vbTextCompare) = 0 Then
            isInDualArray = i ' find the index (start at 0) of where this is sitting at
            Exit For
        End If
    Next i

End Function

Public Function nameInverter(documentName) As String
Dim nameArray() As String

If InStr(documentName, " ") = 0 Then Exit Function
nameArray = Split(documentName, " ") 'Split the name
nameInverter = nameArray(1) & " " & nameArray(0) ' Invert it

End Function
