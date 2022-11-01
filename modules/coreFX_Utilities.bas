Attribute VB_Name = "coreFX_Utilities"
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
Public Function indexDict(sourceDictionary As Dictionary, targetStr As String, Optional Include As Boolean = True, Optional Compare As VbCompareMethod = vbTextCompare) As String()

    indexDict = Filter(sourceDictionary.Keys, targetStr, Include, Compare)
    ' supposively finds by partial match
    
End Function
