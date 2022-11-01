VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} utilityDictionary 
   Caption         =   "Sensei Dictionary"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "utilityDictionary.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "utilityDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Default Form Dimension: 408.75 * 262.5
' SENSEI GLOBAL VAR
Public config As Worksheet
' SENSEI ADSN DATABASE and its reverse variant
Public adsnLib As Scripting.Dictionary, biLnsda As Scripting.Dictionary
' SENSEI Country code Dictionary
Public countryID As Scripting.Dictionary, DIyrtnuoc As Scripting.Dictionary ' Country
Public countryHDP As Scripting.Dictionary, PHDyrtnuoc As Scripting.Dictionary ' HDP not yet used
Public countryQual As Scripting.Dictionary ' country validation qualification
' SENSEI BAQ DATABASE
Public BAQ1st As Scripting.Dictionary, BAQ2nd As Scripting.Dictionary, BAQ3rd As Scripting.Dictionary, BAQ4th As Scripting.Dictionary
Public BAQ1 As String, BAQ2 As String, BAQ3 As String, BAQ4 As String ' For BAQ Builder
Public key ' key is a counter for each array entry
Public sData As Worksheet ' data
Public CIDtbl As Range, HDPtbl As Range ' The Country and HDP table

Private Sub BAQ1_0_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ1_1_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ1_2_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ2_0_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ2_2_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ2_3_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ2_4_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ3_0_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ3_1_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4A_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4C_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4D_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4G_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4I_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4K_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4L_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4N_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4R_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4S_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4T_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4W_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQ4X_Click() ' BAQ.MISC.DIGITS
BAQbuilderAdjust
End Sub

Private Sub BAQinput_Change() ' BAQ INPUTTER
Dim BAQarray() As String ' BAQ array
If BAQinput.Value = vbNullString Then ' destroy blanks
    BAQw1.Caption = "..."
    BAQe1.Caption = "..."
    BAQe2.Caption = "..."
    BAQe3.Caption = "..."
    BAQe4.Caption = "..."
    Exit Sub
End If
BAQw1.Caption = "..."
'If Not BAQops Then
    ' Generic Forward Writing
    'assignment
    BAQarray = matchKey(sourceDictionary:=BAQ1st, matchString:=Format(Left(BAQinput.Value, 1), "0"), Include:=True, Compare:=vbTextCompare)
    If Len(BAQinput.Value) = 1 Then
        For Each key In BAQarray
            With BAQe1
                .Caption = BAQ1st(key)
            End With
            BAQe2.Caption = "..."
            BAQe3.Caption = "..."
            BAQe4.Caption = "..."
        Next
    End If
    ' adequacy
    If Len(BAQinput.Value) = 2 Then
        BAQarray = matchKey(sourceDictionary:=BAQ2nd, matchString:=Format(Mid(BAQinput.Value, 2), "0"), Include:=True, Compare:=vbTextCompare)
        For Each key In BAQarray
            With BAQe2
                .Caption = BAQ2nd(key)
            End With
        Next
        BAQe3.Caption = "..."
        BAQe4.Caption = "..."
    End If
    ' dependent count
    If Len(BAQinput.Value) = 3 Then
        BAQarray = matchKey(sourceDictionary:=BAQ3rd, matchString:=Format(Mid(BAQinput.Value, 3), "0"), Include:=True, Compare:=vbTextCompare)
        For Each key In BAQarray
            With BAQe3
                .Caption = BAQ3rd(key)
            End With
        Next
        BAQe4.Caption = "..."
    End If
    'dependent code
    If Len(BAQinput.Value) = 4 Then
        BAQarray = matchKey(sourceDictionary:=BAQ4th, matchString:=Format(Mid(BAQinput.Value, 4), "0"), Include:=True, Compare:=vbTextCompare)
        For Each key In BAQarray
            With BAQe4
                .Caption = BAQ4th(key)
            End With
        Next
    End If
'ELSE

'END IF
BAQwarning
End Sub

Private Sub BAQops_Click() ' BAQ.REVERSE.GUI
If BAQops Then
    BAQops.Caption = "B"
    'BAQinput.Value = ""
    BAQinput.Locked = True
    V_01.Visible = True
    V_02.Visible = True
    V_03.Visible = True
    V_04.Visible = True
    BAQnotifyD.Visible = False
    BAQexplain.Visible = False
Else
    BAQops.Caption = "V"
    BAQinput.Locked = False
    V_01.Visible = False
    V_02.Visible = False
    V_03.Visible = False
    V_04.Visible = False
    BAQnotifyD.Visible = True
    BAQexplain.Visible = True
End If
End Sub

Private Sub hidePanel_Click() ' FORM.HIDE
    utilityDictionary.Hide
    trackerAPI.Show
End Sub

Private Sub Lib_Change()

End Sub

Private Sub matchInputADSN_Change() ' ADSN.ACTION.LOAD_ENTRIES
Dim responseArray() As String
If matchInputADSN.Value = vbNullString Then
    With matchOutput
        .Clear
    End With
    Exit Sub
End If
    If Not ReverseADSN Then
        responseArray = matchKey(sourceDictionary:=adsnLib, matchString:=matchInputADSN.Value, Include:=True, Compare:=vbTextCompare)
        For Each key In responseArray
            With matchOutput
                .Clear
                .AddItem adsnLib(key)
                .Value = adsnLib(key)
            End With
        Next
    Else
        responseArray = matchKey(sourceDictionary:=biLnsda, matchString:=Format(matchInputADSN.Value, "0000"), Include:=True, Compare:=vbTextCompare)
        For Each key In responseArray
            With matchOutput
                .Clear
                .AddItem biLnsda(key)
                .Value = biLnsda(key)
            End With
        Next
    End If
End Sub


Private Sub matchInputCountry_Change() ' depC.ACTION.LOAD_ENTRIES
Dim responseArray() As String, responseQual() As String
If matchInputCountry.Value = vbNullString Then
    With matchOutput
        .Clear
    End With
    Exit Sub
End If

On Error Resume Next ' hold on
countryQualReset ' Wipe display
DEPinfoHDP.Value = "" ' WIPE HDP DISP

If Not ReverseCountry Then ' Name Look up
    matchInputCountry.MaxLength = 0 ' infinite
    responseArray = matchKey(sourceDictionary:=DIyrtnuoc, matchString:=matchInputCountry.Value, Include:=True, Compare:=vbTextCompare)
    For Each key In responseArray
        With matchOutputCountry
            .Clear
            .AddItem DIyrtnuoc(key)
            .Value = DIyrtnuoc(key) ' yep
        End With
    Next ' below is for country finding
    If Mid(countryQual(matchOutputCountry.Value), 1, 1) = "T" Then
        DEPQ1.BackColor = &H885985
    End If
    If Mid(countryQual(matchOutputCountry.Value), 2, 1) = "T" Then
        DEPQ2.BackColor = &H885985
    End If
    If Mid(countryQual(matchOutputCountry.Value), 3, 1) = "T" Then
        DEPQ3.BackColor = &H885985
    End If
Else ' 2 Digit ID Look up
    matchInputCountry.MaxLength = 2 ' 2 character max
    responseArray = matchKey(sourceDictionary:=countryID, matchString:=Format(matchInputCountry.Value, "00"), Include:=True, Compare:=vbTextCompare)
    For Each key In responseArray
        With matchOutputCountry
            .Clear
            .AddItem countryID(key)
            .Value = countryID(key)
        End With
    Next ' below is for country finding
    If Mid(countryQual(UCase(matchInputCountry.Value)), 1, 1) = "T" Then
        DEPQ1.BackColor = &H885985
    End If
    If Mid(countryQual(UCase(matchInputCountry.Value)), 2, 1) = "T" Then
        DEPQ2.BackColor = &H885985
    End If
    If Mid(countryQual(UCase(matchInputCountry.Value)), 3, 1) = "T" Then
        DEPQ3.BackColor = &H885985
    End If
End If

countryWriteHDP

If Len(matchInputCountry) < 2 Then ' Reset if less than 2 bytes
    countryQualReset
    DEPinfoHDP.Value = ""
End If

End Sub
Sub countryQualReset() ' depC.ACTION.WIPE_DISPLAY

DEPQ1.BackColor = &H8000000F
DEPQ2.BackColor = &H8000000F
DEPQ3.BackColor = &H8000000F

End Sub

Sub countryWriteHDP() '  depC.ACTION.Load_HDP
Dim dCell As Range, Drow As Long, dString As String

If Not ReverseCountry Then ' use output
    For Each dCell In HDPtbl ' loop
        If dCell.Value = matchOutputCountry.Value Then ' found a match hdp
            Drow = dCell.Row
            DEPinfoHDP.Value = DEPinfoHDP.Value & Application.WorksheetFunction.Text(dCell.Value, "00") & "   " & sData.Range("C" & Drow).Value & "   " & Format(Left(sData.Range("D" & Drow).Value, 25), "!@@@@@@@@@@@@@@@@@@@@@@@@@") & "   " & Format(sData.Range("E" & Drow).Value, "$ 000.00") & vbNewLine
        End If
    Next dCell
Else ' use input
    For Each dCell In HDPtbl ' loop IN REVERSE MODE
        If dCell.Value = matchInputCountry.Value Then ' found a match hdp
            Drow = dCell.Row
            DEPinfoHDP.Value = DEPinfoHDP.Value & Application.WorksheetFunction.Text(dCell.Value, "00") & "   " & sData.Range("C" & Drow).Value & "   " & Format(Left(sData.Range("D" & Drow).Value, 25), "!@@@@@@@@@@@@@@@@@@@@@@@@@") & "   " & Format(sData.Range("E" & Drow).Value, "$ 000.00") & vbNewLine
        End If
    Next dCell
End If

End Sub

Private Sub matchOutput_Enter() ' ADSN.ACTION.ENGAGE_COPY
If AutoCopy Then writeClip
End Sub


Private Sub ReverseADSN_Click() ' ADSN.REVERSE.DISPLAY
    With ReverseADSN
        If ReverseADSN Then
            .Caption = "R"
            .ControlTipText = "Looking for Base"
        Else
            .Caption = "F"
            .ControlTipText = "Looking for ADSN"
        End If
    End With
    If ReverseADSN Then
        resultIn.Caption = "Results for Location Lookup"
        srchIn.Caption = "Searching with ADSN"
    Else
        resultIn.Caption = "Results for ADSN Lookup"
        srchIn.Caption = "Searching with Location"
    End If
End Sub


Private Sub ReverseCountry_Click() ' depC.REVERSE.DISPLAY
If ReverseCountry Then
    ReverseCountry.Caption = "ID"
    DEPlabel.Caption = "Input ID Here:"
    matchInputCountry.Value = Left(matchInputCountry.Value, 2)
Else
    ReverseCountry.Caption = "Country"
    DEPlabel.Caption = "Input Country Here:"
End If
End Sub

Private Sub UserForm_Initialize() ' FORM.INITIALIZE
    Set config = Worksheets("SENSEI.CONFIG")
    Set adsnLib = New Scripting.Dictionary ' adsn
    Set biLnsda = New Scripting.Dictionary ' reverse adsn
    Set BAQ1st = New Scripting.Dictionary ' BAQ 1 ST POSITION
    Set BAQ2nd = New Scripting.Dictionary ' BAQ 2 ND POSITION
    Set BAQ3rd = New Scripting.Dictionary ' BAQ 3 RD POSITION
    Set BAQ4th = New Scripting.Dictionary ' BAQ 4 TH Position
    Set countryID = New Scripting.Dictionary ' Forward Country ID
    Set DIyrtnuoc = New Scripting.Dictionary ' Backward Country ID
    Set countryQual = New Scripting.Dictionary ' The CZ, HFP, IDP display dictionary
    Set CIDtbl = Range("tableCountries[COUNTRIES]")
    Set HDPtbl = Range("tableCountriesHDP[COUNTRY]")
    Set sData = Worksheets("SENSEI.DATA")
    initializeUI ' UPDATE DISPLAY
    initializeADSN ' Populate ADSN list
    initializeBAQ ' Populate BAQ List
    initializeCountry ' Pupulate Country List
End Sub
Function matchKey(sourceDictionary As Dictionary, matchString As String, Optional Include As Boolean = True, Optional Compare As VbCompareMethod = vbTextCompare) As String() ' A lookup Function for Dictionary, soucrced from Stack Overflow
    matchKey = Filter(sourceDictionary.Keys, matchString, Include, Compare)
End Function

Sub writeClip() 'ADSN.WRITE.CLIPBOARD
On Error GoTo halt
SetClipboard matchOutput.Value
halt:
End Sub
Sub initializeUI() ' UI.update
    clVer.Caption = config.Range("H2").Value
    clVerType.Caption = config.Range("H3").Value
End Sub
Sub initializeADSN() ' ADSN.DICTIONARY.LOAD
With adsnLib ' write forward lib
    .Add "JOINT PERSONAL PROPERTY SHIPPING OFFICE - JPPSO, SAN ANTONIO TX", "3898"
    .Add "AFFSC, Ellsworth AFB, SD", "8630"
    .Add "343 WG, EIELSON AFB, AK 99702", "4001"
    .Add "51 WG, KUNSAN AB, APO AP 96278", "4002"
    .Add "EUR AFELM NATO AGS SQ SIGONELLA ITALY 096270000", "4017"
    .Add "432 FW, MISAWA AB, APO AP 96319", "4003"
    .Add "7241 ABG, IZMIR AB, APO AE 09821", "4004"
    .Add "18 WG, KADENA AB, APO AP 96368", "4005"
    .Add "52 FW/FMFP, SPANGDAHLEM, APO AE 09126", "4008"
    .Add "501 CSW/FM, UNIT 5555 BOX 9, APO AE 09470 - ALCONBURY", "4009"
    .Add "48 CPTS/FMFPM, LAKENHEATH AB, APO AE 09464", "4011"
    .Add "89 AW/FMFP, ANDREWS AFB, MD 20331", "4012"
    .Add "470 ABF/FMFP, GEILENKIRCHEN AB, APO AE 09104", "4013"
    .Add "DET 1, 786 FSS,UNIT 30400, ATTN:  FSO, APO AE 09131 - STUTTGART", "4014"
    .Add "86 CPTS/FMFP, RAMSTEIN, APO AE 09094-0006", "4015"
    .Add "100 CPTS/FMFP, MILDENHALL, APO AE 09459-0006", "4016"
    .Add "31 CPTS/FMFPM, AVIANO, APO AE 09601-0006", "4017"
    .Add "39 CPTF/FMFP, INCIRLIK AFB, APO AE 09824", "4018"
    .Add "51 WG, OSAN AFB, APO 92678", "4019"
    .Add "55 WG/FMFP, OFFUTT AFB, NE 681113-5000", "4020"
    .Add "97 CPTS/FMFP, ALTUS AFB, OK 73523-5000", "4021"
    .Add "28 BW/FMFP, ELLSWORTH AFB, SD 57706-5000", "4022"
    .Add "19 CPTS/FMF, LITTLE ROCK AFB, AR 72099", "4023"
    .Add "314 CPTS/FMFFS, LITTLE ROCK AFB, AR 72099", "4023"
    .Add "319 CPTS/FMFP, GRAND FORKS AFB, NE 58205", "4024"
    .Add "509 CPTS, WHITEMAN AFB, MO 65305", "4025"
    .Add "22 CPTS/FMFP, MCCONNELL AFB, KS 67221", "4026"
    .Add "375 AW, SCOTT AFB, IL 62225", "4027"
    .Add "71 CPTS/FMFPM, VANCE AFB, OK 73705", "4028"
    .Add "21 SPW/FMFP, PETERSON AFB, CO 80914", "4029"
    .Add "USAFA/FMFPM, USAF ACADEMY, CO 80840", "4030"
    .Add "5 BW/FMFP, MINOT AFB, ND 58705", "4031"
    .Add "OC ALC, TINKER AFB, OK 73145", "4032"
    .Add "460 CPTF/FMFD, BUCKLEYAFB, CO 80011", "4033"
    .Add "50 CPTF, SCHRIEVER AFB, CO 80912", "4034"
    .Add "62 CPTS/FMFPM, MCCHORD AFB, WA 98438", "4036"
    .Add "341 CPTS/FMFP, MALSTROM AFB, MT 59402", "4037"
    .Add "366 WG/FMFP, MT, HOME AFB, ID 83648", "4038"
    .Add "92 CPTS/FMFP, FAIRCHILD AFB, WA 99011", "4039"
    .Add "9 WG/FMFP, BEALE AFB, CA 95903", "4040"
    .Add "90 MW/FMFP, F.E. WARREN AFB, WY 82005", "4041"
    .Add "60 AW, TRAVIS AFB, CA 94535", "4042"
    .Add "99 CPTS/FMFP, NELLIS AFB, NV 89191", "4043"
    .Add "CREECH AFB NV 89018", "4043"
    .Add "AFFTC/FMFPM, EDWARDS AFB, CA 93524", "4045"
    .Add "30 SPW/FMFP, VANDENBERG AFB, CA 93437", "4046"
    .Add "OO-ALC, HILL AFB, UT 84056", "4047"
    .Add "SMC, LOS ANGELES AFBS, CA 90009", "4048"
    .Add "2 CPTS/FMFP, BARKSDALE AFB, LA 71110-5000", "4050"
    .Add "14 CPTS/FMFP, COLUMBUS AFB, MS 39701-1101", "4051"
    .Add "347 CPTS/FMFP, MOODY AFB, GA 31699", "4052"
    .Add "42ND CPTS/FMFP, MAXWELL AFB, AL 36112-6335", "4053"
    .Add "AEDC/FMFPM, ARNOLD AS, TN 37389", "4054"
    .Add "20 FW/FMFPM, SHAW AFB, SC 29152", "4055"
    .Add "81 CPTS/FMFPS, KEESLER AFB, MS 39534-2555", "4056"
    .Add "1 CPTS/FMFP, LANGLEY AFB, VA 23665", "4057"
    .Add "AFDTC/FMFPM, EGLIN AFB, FL 32403", "4058"
    .Add "AFI, KKEFLAVIK, APO AE 09725", "4059"
    .Add "65 CPTF, LAJES FIELD, AE 09720", "4060"
    .Add "USAFCENT/FM Shaw AFB 29152", "4061"
    .Add "HQ AFOATS, MAXWELL AFB, AL", "4064"
    .Add "437 AW/FMFP, CHARLESTON AFB, SC 29404-5000", "4065"
    .Add "ASC/FMFPM, WRIGHT-PATTERSON AFB, OH 45433", "4066"
    .Add "23 CPTS/FMFPM, POPE AFB, NC 28308", "4067"
    .Add "436 CPTS, DOVER AFB, DE 19902", "4068"
    .Add "305 AMW/FMFS, MCGUIRE AFB, NJ 08641", "4069"
    .Add "ESC/FMFPM, HANSCOM AFB, MA 01731", "4070"
    .Add "4 WG/FMFP, SEYMOUR-JOHNSON AFB, NC 27531", "4072"
    .Add "WR-ALC, ROBINS AFB, GA 31098", "4073"
    .Add "ROME LABORATORY, GRIFFISS AFB, NY 13441", "4074"
    .Add "45 SW/FMFS, PATRICK AFB, FL 32925", "4080"
    .Add "1 SOCPTS, HURLBURT FLD, FL 32544", "4081"
    .Add "325 CPTS/FMFP, TYNDALL AFB, FL 32403-5535", "4082"
    .Add "6 CPTS/FMF, MACDILL AFB FL 33621", "4083"
    .Add "FT BRAGG", "4090"
    .Add "AIR FORCE SECURITY FORCES CENTER", "4094"
    .Add "LACKLAND  TECH SCHOOL, LACKLAND AFB TX", "4095"
    .Add "JOINT BASE SAN ANTONIO (includes Lackland, Randolph, Ft Sam Houston)", "4096"
    .Add "82 CPTS//FMFP, SHEPPARD AFB, TX 76311", "4100"
    .Add "47 FTW/FMFP, LAUGHLIN AFB, TX 78843-5241", "4101"
    .Add "17 CPTS/FMFP, GOODFELLOW AFB, TX76908-4418", "4102"
    .Add "7 CPTS, DYESS AFB, TX 79607", "4103"
    .Add "355 CPTS/FMFP, DAVIS-MONTHAN AFB, AZ 85707", "4126"
    .Add "FT MEADE", "4127"
    .Add "HQ 11WG/FMFP-B, BOLLING AFB,", "4128"
    .Add "27 FW/FMFP, CANNON AFB, NM 88103", "4129"
    .Add "49 FW, HOLLOMAN AFB, NM 88330", "4130"
    .Add "58 CPTS, LUKE AFB, AZ 85309", "4131"
    .Add "377 ABW/FMFPM, KIRTLAND AFB,NM 87117", "4132"
    .Add "15 ABW/FMF, HICKAM AFB, HI 96853", "4150"
    .Add "36 ABW/FMF, ANDERSON AFB, APO AP 96543", "4152"
    .Add "3 WG/FM, ELMENDORF AFB, AK 99506", "4153"
    .Add "YOKOTA AB JAPAN", "6688"
    .Add "HQ AIR FORCEPERSONNEL CENTER, RANDOLPH AFB, TX  78150", "8888"
    .Add "CADETS ¨C UNITED STATES AIR FORCE ACADEMY, CO", "8890"
    .Add "BASIC MILITARY TRAINING", "9998"
End With
With biLnsda ' write reverse Lib
    .Add "3898", "JOINT PERSONAL PROPERTY SHIPPING OFFICE - JPPSO SAN ANTONIO TX"
    .Add "8630", "AFFSC Ellsworth AFB SD"
    .Add "4001", "343 WG EIELSON AFB AK 99702"
    .Add "4002", "51 WG KUNSAN AB APO AP 96278"
    .Add "4017", "EUR AFELM NATO AGS SQ SIGONELLA ITALY 096270000; 31 CPTS/FMFPM AVIANO APO AE 09601-0006"
    .Add "4003", "432 FW MISAWA AB APO AP 96319"
    .Add "4004", "7241 ABG IZMIR AB APO AE 09821"
    .Add "4005", "18 WG KADENA AB APO AP 96368"
    .Add "4008", "52 FW/FMFP SPANGDAHLEM APO AE 09126"
    .Add "4009", "501 CSW/FM UNIT 5555 BOX 9 APO AE 09470 - ALCONBURY"
    .Add "4011", "48 CPTS/FMFPM LAKENHEATH AB APO AE 09464"
    .Add "4012", "89 AW/FMFP ANDREWS AFB MD 20331"
    .Add "4013", "470 ABF/FMFP GEILENKIRCHEN AB APO AE 09104"
    .Add "4014", "DET 1 786 FSSUNIT 30400 ATTN:  FSO APO AE 09131 - STUTTGART"
    .Add "4015", "86 CPTS/FMFP RAMSTEIN APO AE 09094-0006"
    .Add "4016", "100 CPTS/FMFP MILDENHALL APO AE 09459-0006"
    .Add "4018", "39 CPTF/FMFP INCIRLIK AFB APO AE 09824"
    .Add "4019", "51 WG OSAN AFB APO 92678"
    .Add "4020", "55 WG/FMFP OFFUTT AFB NE 681113-5000"
    .Add "4021", "97 CPTS/FMFP ALTUS AFB OK 73523-5000"
    .Add "4022", "28 BW/FMFP ELLSWORTH AFB SD 57706-5000"
    .Add "4023", "19 CPTS/FMF LITTLE ROCK AFB AR 72099; 314 CPTS/FMFFS LITTLE ROCK AFB AR 72099"
    .Add "4024", "319 CPTS/FMFP GRAND FORKS AFB NE 58205"
    .Add "4025", "509 CPTS WHITEMAN AFB MO 65305"
    .Add "4026", "22 CPTS/FMFP MCCONNELL AFB KS 67221"
    .Add "4027", "375 AW SCOTT AFB IL 62225"
    .Add "4028", "71 CPTS/FMFPM VANCE AFB OK 73705"
    .Add "4029", "21 SPW/FMFP PETERSON AFB CO 80914"
    .Add "4030", "USAFA/FMFPM USAF ACADEMY CO 80840"
    .Add "4031", "5 BW/FMFP MINOT AFB ND 58705"
    .Add "4032", "OC ALC TINKER AFB OK 73145"
    .Add "4033", "460 CPTF/FMFD BUCKLEYAFB CO 80011"
    .Add "4034", "50 CPTF SCHRIEVER AFB CO 80912"
    .Add "4036", "62 CPTS/FMFPM MCCHORD AFB WA 98438"
    .Add "4037", "341 CPTS/FMFP MALSTROM AFB MT 59402"
    .Add "4038", "366 WG/FMFP MT HOME AFB ID 83648"
    .Add "4039", "92 CPTS/FMFP FAIRCHILD AFB WA 99011"
    .Add "4040", "9 WG/FMFP BEALE AFB CA 95903"
    .Add "4041", "90 MW/FMFP F.E. WARREN AFB WY 82005"
    .Add "4042", "60 AW TRAVIS AFB CA 94535"
    .Add "4043", "99 CPTS/FMFP NELLIS AFB NV 89191; CREECH AFB NV 89018"
    .Add "4045", "AFFTC/FMFPM EDWARDS AFB CA 93524"
    .Add "4046", "30 SPW/FMFP VANDENBERG AFB CA 93437"
    .Add "4047", "OO-ALC HILL AFB UT 84056"
    .Add "4048", "SMC LOS ANGELES AFBS CA 90009"
    .Add "4050", "2 CPTS/FMFP BARKSDALE AFB LA 71110-5000"
    .Add "4051", "14 CPTS/FMFP COLUMBUS AFB MS 39701-1101"
    .Add "4052", "347 CPTS/FMFP MOODY AFB GA 31699"
    .Add "4053", "42ND CPTS/FMFP MAXWELL AFB AL 36112-6335"
    .Add "4054", "AEDC/FMFPM ARNOLD AS TN 37389"
    .Add "4055", "20 FW/FMFPM SHAW AFB SC 29152"
    .Add "4056", "81 CPTS/FMFPS KEESLER AFB MS 39534-2555"
    .Add "4057", "1 CPTS/FMFP LANGLEY AFB VA 23665"
    .Add "4058", "AFDTC/FMFPM EGLIN AFB FL 32403"
    .Add "4059", "AFI KKEFLAVIK APO AE 09725"
    .Add "4060", "65 CPTF LAJES FIELD AE 09720"
    .Add "4061", "USAFCENT/FM Shaw AFB 29152"
    .Add "4064", "HQ AFOATS MAXWELL AFB AL"
    .Add "4065", "437 AW/FMFP CHARLESTON AFB SC 29404-5000"
    .Add "4066", "ASC/FMFPM WRIGHT-PATTERSON AFB OH 45433"
    .Add "4067", "23 CPTS/FMFPM POPE AFB NC 28308"
    .Add "4068", "436 CPTS DOVER AFB DE 19902"
    .Add "4069", "305 AMW/FMFS MCGUIRE AFB NJ 08641"
    .Add "4070", "ESC/FMFPM HANSCOM AFB MA 01731"
    .Add "4072", "4 WG/FMFP SEYMOUR-JOHNSON AFB NC 27531"
    .Add "4073", "WR-ALC ROBINS AFB GA 31098"
    .Add "4074", "ROME LABORATORY GRIFFISS AFB NY 13441"
    .Add "4080", "45 SW/FMFS PATRICK AFB FL 32925"
    .Add "4081", "1 SOCPTS HURLBURT FLD FL 32544"
    .Add "4082", "325 CPTS/FMFP TYNDALL AFB FL 32403-5535"
    .Add "4083", "6 CPTS/FMF MACDILL AFB FL 33621"
    .Add "4090", "FT BRAGG"
    .Add "4094", "AIR FORCE SECURITY FORCES CENTER"
    .Add "4095", "LACKLAND  TECH SCHOOL LACKLAND AFB TX"
    .Add "4096", "JOINT BASE SAN ANTONIO (includes Lackland Randolph Ft Sam Houston)"
    .Add "4100", "82 CPTS//FMFP SHEPPARD AFB TX 76311"
    .Add "4101", "47 FTW/FMFP LAUGHLIN AFB TX 78843-5241"
    .Add "4102", "17 CPTS/FMFP GOODFELLOW AFB TX76908-4418"
    .Add "4103", "7 CPTS DYESS AFB TX 79607"
    .Add "4126", "355 CPTS/FMFP DAVIS-MONTHAN AFB AZ 85707"
    .Add "4127", "FT MEADE"
    .Add "4128", "HQ 11WG/FMFP-B BOLLING AFB"
    .Add "4129", "27 FW/FMFP CANNON AFB NM 88103"
    .Add "4130", "49 FW HOLLOMAN AFB NM 88330"
    .Add "4131", "58 CPTS LUKE AFB AZ 85309"
    .Add "4132", "377 ABW/FMFPM KIRTLAND AFBNM 87117"
    .Add "4150", "15 ABW/FMF HICKAM AFB HI 96853"
    .Add "4152", "36 ABW/FMF ANDERSON AFB APO AP 96543"
    .Add "4153", "3 WG/FM ELMENDORF AFB AK 99506"
    .Add "6688", "YOKOTA AB JAPAN"
    .Add "8888", "HQ AIR FORCEPERSONNEL CENTER RANDOLPH AFB TX  78150"
    .Add "8890", "CADETS ¨C UNITED STATES AIR FORCE ACADEMY CO"
    .Add "9998", "BASIC MILITARY TRAINING"
End With
End Sub

Sub initializeBAQ() ' BAQ.DICITIONARY.LOAD
'1 st digit assignment
With BAQ1st
    .Add "0", "Undefined Quarter Assignment"
    .Add "1", "Assigned to Government Quarter"
    .Add "2", "Not Assigned to Government Quarter"
End With
'2 nd digit adequacy
With BAQ2nd
    .Add "0", "Adequcy unavailable due to Not Assigned to Gov't Quarter"
    .Add "1", "Adequate Quarter"
    .Add "2", "Inadequate Quarter"
    .Add "3", "Partial BAH or In Dorm Airmen"
    .Add "4", "Assigned but paying Dependent(s) Support"
End With
'3 rd digit dependency status
With BAQ3rd
    .Add "0", "No Valid Dependent"
    .Add "1", "Has Valid Dependent(s)"
End With
'4 th digit dependent type
With BAQ4th
    .Add "A", "Has Civilian Spouse"
    .Add "B", "Paying Child Support (Expired 1991-12-04)"
    .Add "C", "Has the Custody of Child(ren)"
    .Add "D", "Has Parent(s) as Secondary Dependent(s)"
    .Add "F", "Has Stepchild(ren) (Expired 1991-12-04)"
    .Add "G", "The Grandfathered Clause (DA 26.16)"
    .Add "I", "Has Military Spouse without Child(ren)"
    .Add "K", "Has Ward of the Court as Secondary Dependent"
    .Add "L", "Has Parent-in-Law as Secondary Dependent"
    .Add "N", "Does not have Custody of Child(ren)"
    .Add "Q", "Has Adopted Child (Expired 1991-12-04)"
    .Add "R", "Has no Dependents or Single"
    .Add "S", "Has Student(s) (21 or 22) as Secondary Dependent"
    .Add "T", "Has Incapitated Child(ren) as Secondary Dependent"
    .Add "V", "Has illegitimate Child(ren) (Expired 1991-12-04)"
    .Add "W", "Has Military Spouse and Child(ren)"
    .Add "X", "Has the Custody of Child(ren) (Legacy Indicator)"
End With
End Sub

Sub initializeCountry() ' depC.DICTIONARY.LOAD
Dim fCell As Range, fRow As Long
' Forward
With CIDtbl
    For Each fCell In CIDtbl
        fRow = fCell.Row
        countryID.Add Application.WorksheetFunction.Text(fCell.Value, "00"), sData.Range("G" & fRow).Value ' use ID
        DIyrtnuoc.Add sData.Range("G" & fRow).Value, Application.WorksheetFunction.Text(fCell.Value, "00") ' use Name
    Next fCell
End With

' Generic Qualification
With CIDtbl
    For Each fCell In CIDtbl
        fRow = fCell.Row
        countryQual.Add Application.WorksheetFunction.Text(fCell.Value, "00"), sData.Range("H" & fRow).Value & sData.Range("I" & fRow).Value & sData.Range("J" & fRow).Value
    Next fCell
End With


End Sub

Sub BAQwarning() ' BAQ.REVERSE.GUI.BUILDER
If Len(BAQinput.Value) > 4 Then
    BAQw1.Caption = "BAQ Code is longer than 4 Byte"
    Exit Sub
End If
On Error GoTo BAQERROR
Select Case Mid(BAQinput.Value, 1, 1) ' Check assignment
Case Is < 1
    BAQw1.Caption = "Undefined Assignment"
Case Is > 2
    BAQw1.Caption = "Assignment Invalid"
Case Is = 1 ' second tier
If Len(BAQinput.Value) > 1 Then
    Select Case Mid(BAQinput.Value, 2, 1)
    Case 2 ' 3rd on depn
    If Len(BAQinput.Value) > 2 Then
        Select Case Mid(BAQinput.Value, 3, 1)
        Case 0 ' 4th on depn type
            If Len(BAQinput.Value) > 3 Then
                Select Case UCase(Mid(BAQinput.Value, 4))
                Case "I"
                Case "R"
                Case "N"
                Case "X"
                    BAQw1.Caption = "With type X, Start Date should be > 1991-08-01"
                Case Else
                    BAQw1.Caption = "Type Incompatibility: must be I, N, R, X"
                End Select
            End If
        Case 1 ' 4th on depn code for valid count
            If Len(BAQinput.Value) > 3 Then
                Select Case UCase(Mid(BAQinput.Value, 4))
                Case "A"
                Case "C"
                Case "D"
                Case "G"
                Case "K"
                    BAQw1.Caption = "With type K, Start Date should be > 1994-07-01"
                Case "L"
                Case "S"
                Case "T"
                Case "W"
                Case Else
                    BAQw1.Caption = "Type Incompatibility: must be A, C, D, G, K, L, S, T, W"
                End Select
            End If
        Case Else
        End Select
    End If
    Case 3 ' LEVEL 3
    If Len(BAQinput.Value) > 2 Then
        If Mid(BAQinput.Value, 3, 1) <> 0 Then
            BAQw1.Caption = "Adequacy-Dependent Count Incompatibility: Must be 0"
        Else
            If Len(BAQinput.Value) > 3 Then
                Select Case UCase(Mid(BAQinput.Value, 4))
                Case "I"
                Case "R"
                Case "N"
                Case "X"
                    BAQw1.Caption = "With type X, Start Date should be > 1991-08-01"
                Case Else
                    BAQw1.Caption = "Type Incompatibility: must be I, N, R, X"
                End Select
            End If
        End If
    End If
    Case 4 ' LEVEL
    If Len(BAQinput.Value) > 2 Then
        If Mid(BAQinput.Value, 3, 1) <> 1 Then
            BAQw1.Caption = "Adequacy-Dependent Count Incompatibility: Must be 1"
        Else
            If Len(BAQinput.Value) > 3 Then
                Select Case UCase(Mid(BAQinput.Value, 4))
                Case "A"
                Case "C"
                Case "D"
                Case "G"
                Case "K"
                    BAQw1.Caption = "With type K, Start Date should be > 1994-07-01"
                Case "L"
                Case "S"
                Case "T"
                Case "W"
                Case Else
                    BAQw1.Caption = "Type Incompatibility: must be A, C, D, G, K, L, S, T, W"
                End Select
            End If
        End If
    End If
    Case Else
        BAQw1.Caption = "Assignment-Adequacy Incompatibility: Must be 2, 3, 4"
    End Select
End If
Case Is = 2 ' second tier
If Len(BAQinput.Value) > 1 Then
    Select Case Mid(BAQinput.Value, 2, 1)
    Case 0 ' 3rd on depn
    If Len(BAQinput.Value) > 2 Then
        Select Case Mid(BAQinput.Value, 3, 1)
        Case 0 ' 4th on depn type
            If Len(BAQinput.Value) > 3 Then
                Select Case UCase(Mid(BAQinput.Value, 4))
                Case "I"
                Case "N"
                Case "R"
                Case "X"
                    BAQw1.Caption = "With type X, Start Date should be > 1991-08-01"
                Case Else
                    BAQw1.Caption = "Type Incompatibility: must be I, N, R, X"
                End Select
            End If
        Case 1 ' 4th on depn code for valid count
            If Len(BAQinput.Value) > 3 Then
                Select Case UCase(Mid(BAQinput.Value, 4))
                Case "A"
                Case "C"
                Case "D"
                Case "G"
                Case "K"
                    BAQw1.Caption = "With type K, Start Date should be > 1994-07-01"
                Case "L"
                Case "S"
                Case "T"
                Case "W"
                Case Else
                    BAQw1.Caption = "Type Incompatibility: must be A, C, D, G, K, L, S, T, W"
                End Select
            End If
        Case Else
            BAQw1.Caption = "Dependent Present Status invalid, must be 1 or 0"
        End Select
    End If
    Case Else
        BAQw1.Caption = "Assignment-Adequacy Incompatibility: Must be 0"
    End Select
End If
End Select
Exit Sub

BAQERROR:
    BAQw1.Caption = "Format of the 4th Digit is not an Alphabet"
End Sub

Sub BAQbuilderAdjust() ' BAQ.REVERSE.GUI.BUILDER.QUALIFY
    BAQbuilderRecover
'Character space 1
If BAQ1_0 Then
    BAQbuilderInvalid
End If
If BAQ1_1 Then
    BAQ100
    BAQ1_2.Enabled = False
    BAQ2_0.Enabled = False
    BAQ1 = "1"
End If
If BAQ1_2 Then
    BAQ100
    BAQ1_1.Enabled = False
    BAQ2_2.Enabled = False
    BAQ2_3.Enabled = False
    BAQ2_4.Enabled = False
    BAQ1 = "2"
End If

'character space 2
If BAQ2_0 Then
    BAQ100
    BAQ1_1.Enabled = False
    BAQ2_2.Enabled = False
    BAQ2_3.Enabled = False
    BAQ2_4.Enabled = False
    BAQ2 = "0"
End If
If BAQ2_2 Then
    BAQ100
    BAQ1_2.Enabled = False
    BAQ2_0.Enabled = False
    BAQ2_3.Enabled = False
    BAQ2_4.Enabled = False
    BAQ2 = "2"
End If
If BAQ2_3 Then
    BAQ100
    BAQ1_2.Enabled = False
    BAQ2_0.Enabled = False
    BAQ2_2.Enabled = False
    BAQ2_4.Enabled = False
    BAQ3_1.Enabled = False
    BAQ2 = "3"
End If
If BAQ2_4 Then
    BAQ100
    BAQ1_2.Enabled = False
    BAQ2_0.Enabled = False
    BAQ2_2.Enabled = False
    BAQ2_3.Enabled = False
    BAQ2 = "4"
End If

'character space 3
If BAQ3_0 Then
    BAQ100
    BAQ2_4.Enabled = False
    BAQ3_1.Enabled = False
    BAQ3 = "0"
End If
If BAQ3_1 Then
    BAQ100
    BAQ2_3.Enabled = False
    BAQ3_0.Enabled = False
    BAQ3 = "1"
End If

'CHARACTER SPACE 4

BAQbuilder4th
BAQinput.Value = BAQ1 & BAQ2 & BAQ3 & BAQ4 ' WRITE BAQ TO INPUT
End Sub
Sub BAQ100() ' BAQ.INVALID_BUILDER
    BAQ1_0.Enabled = False
End Sub

Sub BAQbuilderRecover() ' BAQ.RESET_QUALIFICATION
BAQ1_0.Enabled = True
BAQ1_1.Enabled = True
BAQ1_2.Enabled = True
BAQ2_0.Enabled = True
BAQ2_2.Enabled = True
BAQ2_3.Enabled = True
BAQ2_4.Enabled = True
BAQ3_0.Enabled = True
BAQ3_1.Enabled = True
BAQ4A.Enabled = True
BAQ4C.Enabled = True
BAQ4D.Enabled = True
BAQ4G.Enabled = True
BAQ4I.Enabled = True
BAQ4K.Enabled = True
BAQ4L.Enabled = True
BAQ4N.Enabled = True
BAQ4R.Enabled = True
BAQ4S.Enabled = True
BAQ4T.Enabled = True
BAQ4W.Enabled = True
BAQ4X.Enabled = True
BAQ1 = ""
BAQ2 = ""
BAQ3 = ""
BAQ4 = ""
End Sub
Sub BAQbuilderInvalid() ' BAQ.NULLIFY_ALL
BAQ1_1.Enabled = False
BAQ1_2.Enabled = False
BAQ2_0.Enabled = False
BAQ2_2.Enabled = False
BAQ2_3.Enabled = False
BAQ2_4.Enabled = False
BAQ3_0.Enabled = False
BAQ3_1.Enabled = False
BAQ4A.Enabled = False
BAQ4C.Enabled = False
BAQ4D.Enabled = False
BAQ4G.Enabled = False
BAQ4I.Enabled = False
BAQ4K.Enabled = False
BAQ4L.Enabled = False
BAQ4N.Enabled = False
BAQ4R.Enabled = False
BAQ4S.Enabled = False
BAQ4T.Enabled = False
BAQ4W.Enabled = False
BAQ4X.Enabled = False
End Sub
Sub BAQbuilder4th() ' BAQ.VALIDATE_4TH.PREV
If BAQ3_0 Then
    BAQ4A.Enabled = False
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4I.Enabled = False
End If
If BAQ3_1 Then
    BAQ4I.Enabled = False
    BAQ4R.Enabled = False
    BAQ4N.Enabled = False
    BAQ4X.Enabled = False
End If
If BAQ4A Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "A"
End If
If BAQ4C Then
    BAQ4A.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "C"
End If
If BAQ4D Then
    BAQ4C.Enabled = False
    BAQ4A.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "D"
End If
If BAQ4G Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4A.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "G"
End If
If BAQ4I Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4A.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "I"
End If
If BAQ4K Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4A.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "K"
End If
BAQbuilder4th2
End Sub

Sub BAQbuilder4th2() ' BAQ.VALUDATE_4TH.NEXT
If BAQ4L Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4A.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "L"
End If
If BAQ4N Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4A.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "N"
End If
If BAQ4R Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4A.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "R"
End If
If BAQ4S Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4A.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "S"
End If
If BAQ4T Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4A.Enabled = False
    BAQ4W.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "T"
End If
If BAQ4W Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4A.Enabled = False
    BAQ4X.Enabled = False
    BAQ4 = "W"
End If
If BAQ4X Then
    BAQ4C.Enabled = False
    BAQ4D.Enabled = False
    BAQ4G.Enabled = False
    BAQ4I.Enabled = False
    BAQ4K.Enabled = False
    BAQ4L.Enabled = False
    BAQ4N.Enabled = False
    BAQ4R.Enabled = False
    BAQ4S.Enabled = False
    BAQ4T.Enabled = False
    BAQ4W.Enabled = False
    BAQ4A.Enabled = False
    BAQ4 = "X"
End If
End Sub
