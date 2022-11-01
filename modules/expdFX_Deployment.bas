Attribute VB_Name = "expdFX_Deployment"
Option Explicit

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
