Attribute VB_Name = "Lib_CE"
' TN TABLE TRANSLATOR - BY INDEX (AC_[fid])
' UI - Have Tab System Linked
' XE - Have Input Interface Setup
' CE - Have Coding Engine Linked up
'01 - FID 10 [UI]
'   0 - start
'   1 - stop
'   2 - report
'   3 - correct
'   4 - cancel
'02 - FID 12 [UI]
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'03 - FID 14 [UI]
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'04 - FID 15 [UI]
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'05 - FID 21 [UI]
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'06 - FID 23 [UI]
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'07 - FID 35
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'08 - FID 36
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'09 - FID 40
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'10 - FID 43
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'11 - FID 46
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'12 - FID 59
'   0 - report
'   1 - correct
'   2 - cancel
'13 - FID 61
'   0 - report
'   1 - correct
'   2 - cancel
'14 - FID 65
'   0 - start
'   1 - stop
'   2 - report
'   3 - correct
'   4 - cancel
'15 - FID 67
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'16 - FID 68
'   0 - start
'   1 - stop
'   2 - report
'   3 - change
'   4 - correct
'   5 - cancel
'17 - FID 81
'   0 - start
'   1 - report
'   2 - correct
'   3 - cancel
'18 - FID AD
'   0 - start
'   1 - stop
'   2 - change
'   3 - increase (07)
'   4 - decrease (08)
'19 - FID C2
'(C203/C903)
'   0 - report
'20 - FID C9
'(C203/C903)
'   0 - report
'21 - FID DE
'   0 - start
'   1 - change
'22 - FID DN
'   0 - start
'   1 - stop
'   2 - report
'   3 - correct
'   4 - cancel
'23 - FID DQ
'   0 - start
'   1 - change
'   2 - suspend (18)
'   3 - resume (20)
'24 - FID DR
'   0 - change
'   1 - suspend (18)
'   2 - resume (20)
'25 - FID DS
'   0 - start
'   1 - cancel
'   2 - suspend (18)
'   3 - resume (20)
'26 - FID DT
'   0 - change
'27 - FID E7
'   0 - start
'   1 - stop
'   2 - cancel
'   3 - release
'28 - FID E8
'
'
'
'29 - FID FJ
'30 - FID FK
'31 - FID FL
'32 - FID LC
'33 - FID LH
'34 - FID MG
'35 - FID MH
'36 - FID PA
'37 - FID PQ
'38 - FID PZ
'39 - FID PT
'40 - FID SB
'41 - FID SC
'42 - FID SH
'43 - FID ST
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
