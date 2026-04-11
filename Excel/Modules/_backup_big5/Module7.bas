Attribute VB_Name = "Module7"
Sub Macro15()
'
' mikeon _ 2015/10/18 _____ªº¥¨¶°
' ____¦¬ÂÃ

'
Application.ScreenUpdating = False
Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²

Application.Calculation = xlManual

Sheets("__").Selectct
    Range("A1:R29").Select
    Selection.Copy
    
Sheets("__").Selectct
[f1] = "_"""
[f2] = Sheets("__").[a2]2]

Call lst
  
Sheets("__").Selectct
    [i16] = "___"ÂÃ"
    Application.CutCopyMode = False
    [y16].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"

End Sub

Sub Macro16()
'
' mikeon _ 2015/10/19 _____ªº¥¨¶°
' ____¦¬ÂÃ

'
Application.ScreenUpdating = False
Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Application.Calculation = xlManual

Sheets("__").Selectct
    Range("A1:R24").Select
    Selection.Copy
    
Sheets("__").Selectct
[f1] = "_"""
[f2] = mkw(Sheets("__").[a2])])

Call lst
    
Sheets("__").Selectct
    [i14] = "___"ÂÃ"
    Application.CutCopyMode = False
    [y16].Select
    
Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
    
End Sub


Sub Macro17()
'
' mikeon _ 2015/10/19 _____ªº¥¨¶°
' ____¦¬ÂÃ

'
Application.ScreenUpdating = False
Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Application.Calculation = xlManual

Sheets("__").Selectct
    [c1].NumberFormatLocal = "@"
    Range("A1:R28").Select
    Selection.Copy
    
Sheets("__").Selectct
[f1] = "_"""
[f2] = Format(Sheets("__").[a2], "0000")")

Call lst
    
Sheets("__").Selectct
    [i14] = "___"ÂÃ"
    Application.CutCopyMode = False
    [y16].Select
    
Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
    
End Sub


Sub Macro18()
'
' mikeon _ 2015/10/19 _____ªº¥¨¶°
' ____¦¬ÂÃ

'
Application.ScreenUpdating = False
Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Application.Calculation = xlManual

Sheets("__").Selectct
    Range("A1:R23").Select
    Selection.Copy
    
Sheets("__").Selectct
[f1] = "_"""
[f2] = Format(Sheets("__").[a2], "000000")")

Call lst
    
Sheets("__").Selectct
    [i14] = "___"ÂÃ"
    Application.CutCopyMode = False
    [x16].Select
    
Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
    
End Sub


Sub Macro19()
'
' mikeon _ 2015/10/19 _____ªº¥¨¶°
' ____¦¬ÂÃ

'
Application.ScreenUpdating = False
Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Application.Calculation = xlManual

Sheets("__").Selectct
    Range("A1:R24").Select
    Selection.Copy
    
Sheets("__").Selectct
[f1] = Sheets("__").[a1]1]
[f2] = mkw(Sheets("__").[a2])])
If UCase(Trim(Sheets("__").[a1])) = "CN" Then [f2] = Format(Sheets("__").[a2], "000000")00")
If UCase(Trim(Sheets("__").[a1])) = "HK" Then [f2] = Format(Sheets("__").[a2], "0000")00")
       
Call lst

Sheets("__").Selectct
    [i14] = "___"ÂÃ"
    Application.CutCopyMode = False
    [ae16].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
    
End Sub



Public Sub lst()

k = 1 '_______h»¡©ú¦C
Do Until Cells(k, 1) = ""
k = k + 1
Loop
k = k - 1

i = k: j = k
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
   j = i
   f2a$ = mkw(UCase(Trim(Cells(i, 1))))
   If f2a$ = UCase(Trim([f2])) And Trim(Cells(i + 1, 4)) = [f1] Then
   i = i - 1
   GoTo ps
   End If
End If
i = i + 1
Loop

i = j
Do Until i - j > 100
If IsError(Cells(i, 9)) Then
j = i
GoTo i100
End If
If Trim(Cells(i, 9)) <> "" Then j = i
i100: i = i + 1
Loop
i = j + 4

ps: Cells(i, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
      
Cells(i, 1) = UCase(Cells(i, 1))
Cells(i + 1, 1).NumberFormatLocal = "@"
If [f1] = "_" Or [f1] = "CN" Then Cells(i + 1, 1) = Format(Cells(i + 1, 1), "000000"))
If [f1] = "_" Or [f1] = "HK" Then Cells(i + 1, 1) = Format(Cells(i + 1, 1), "0000"))

Range("v:v").NumberFormatLocal = "@"
    
k = 2
Do Until Cells(k, 22) = ""
If [f1] = "_" Or [f1] = "CN" Then Cells(k, 22) = Format(Cells(k, 22), "000000"))
If [f1] = "_" Or [f1] = "HK" Then Cells(k, 22) = Format(Cells(k, 22), "0000"))
f22a$ = mkw(UCase(Trim(Cells(k, 22))))
If f22a = UCase(Trim([f2])) And UCase(Trim(Cells(k, 28))) = UCase(Trim([f1])) Then Exit Do '__´ä
If f22a = UCase(Trim([f2])) And UCase(Trim(Cells(k, 28))) = UCase(Trim(Cells(i, 1))) Then Exit Do '__²y
k = k + 1
Loop
Cells(k, 22) = Cells(i + 1, 1) '__¦W
Cells(k, 23) = Cells(i + 3, 11) '_____ø¹S²v
Cells(k, 24) = Cells(i + 2, 11) '__»ù
Cells(k, 25) = Cells(i + 4, 11) '_Q
Cells(k, 26) = Cells(i + 6, 11) '_Q
Cells(k, 28) = [f1]
If Cells(k, 30) = "" Then Cells(k, 30) = 1 '__­È
Cells(k, 31) = Cells(i, 1) '__¤¶
If Cells(i + 2, 4) = "" Then
Cells(k, 28) = Cells(i, 1)
Cells(k, 31) = Cells(i, 2)
End If
Cells(k, 36) = Cells(i, 9)  ' __´Á
Cells(k, 37) = Cells(i + 3, 11) '_____ø¹S²v
Cells(k, 38) = Cells(i + 2, 11) '__»ù

i = 1: j = 1: b = 0
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i
b = b + 1
End If
i = i + 1
Loop

i = 1: j = 1: n = 0
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i: n = n + 1
Cells(i - 2, 1) = "  " & n & " / " & b & " _"""
For c = 2 To k
If Trim(Cells(c, 22)) = Trim(Cells(i, 1)) And Cells(c, 28) = [f1] Then Cells(c, 27) = n
Next c
End If
i = i + 1
Loop

Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop

Call fm7
[f1] = "": [f2] = ""
Cells(j + 32, 26).Select

End Sub

Public Sub fm7()


Sheets("__").Selectct

Cells.Select
Selection.RowHeight = 16
Selection.ColumnWidth = 8
With Selection.Font
    .Name = "____"úÅé"
    .Name = "Arial"
    .FontStyle = "__"Ç"
    .Size = 10
End With

[a1] = "     ."
[a2] = "     ."
[a3] = "     ."
[a4] = "     ."
[a5] = "     ."
[m3] = "____+_______________15%"~³ø¹S²vÃ­©wÁÍ¦V15%"
[m4] = "__Alex_________"¶WÁpµ²µ{¦¡"
[y1] = "_$"""
[Z1] = "_$"""
[k3] = "___________ [__]"ö¤@¦¸ [§ó·s]"
[k4] = "_____3___5___8___11__"¡B8¤ë¤¤¡B11¤ë¤¤"
[n5] = "__________(___AN_)"«(¤¶²Ð­¶ANÄæ)"

[L1] = " ___ [__]_U _______ [__]"é´Á¦A«ö [§ó·s]"
[L1].Select: With Selection.Font
        .ColorIndex = xlAutomatic
End With
    
i = 2: Do Until Cells(i, 22) = ""
If IsDate(Cells(i, 21)) Then
[L1] = " U ______ [__] __"ó·s] °]³ø"
Range("L1").Select
    With Selection.Font
        .Color = -16776961
    End With
Exit Do
End If
i = i + 1
Loop

[t1] = "____U _ __"æ ªÅ¥Õ"
[t2] = "___X"GX"
[ae1] = "_________"Áä¥i±Æ¦æ"
[t6] = "___ [__] ________"ºâ¹w´Á³ø¹S²v"
[t7] = "[__] ___"§ó§ï"
Range("T5").Select: ActiveCell.FormulaR1C1 = "=IF(R[-1]C=TODAY(),""[__] ___"",""[__] ___"")"¥¼§ó·s"")"

Range("t4:t7").Select
With Selection
.HorizontalAlignment = xlRight
.ShrinkToFit = False
End With
Range("t4").Select
    With Selection
    .NumberFormatLocal = "yyyy/m/d;@"
    .ShrinkToFit = True
End With

Range("M5").Select
    With Selection.Font
        .Color = -65536
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Cells(1, 40) = "____ " & "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date=&type=ALLBUT0999"999"
Cells(2, 40) = "____ " & "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=htm" & " ________________"£¨Ñ¥xªÑ¤WÂdªÑ»ùºô§}"
Cells(3, 40) = "__ " & [m5] & " ______________________Alex__"å©M®á¡B§d«a½å®á¡BAlex§d®á"
Cells(4, 40) = "________ = ______ - ((___/___)^(1/8)-1)"»ù/ÂÂªÑ»ù)^(1/8)-1)"
Cells(5, 40) = "_____________________"»ù¡BÀËÅç·s°]³ø¶i¨Ó§_"
Cells(6, 40) = "___________M5"Ä³¼Ð¥Ü¯ÊM5"

[j1].ColumnWidth = 8
Range("w:w").ColumnWidth = 9
Range("aa1:aa1, ac1:ac1").ColumnWidth = 8
Range("ab1").ColumnWidth = 6
Range("ac:ac").ColumnWidth = 12
Range("ae:ae").ColumnWidth = 70
Range("af:am").ColumnWidth = 3
Range("an:an").ColumnWidth = 25
Range("w:w").NumberFormatLocal = "0%"
Range("t1:t2,v:v, ad:ad").HorizontalAlignment = xlLeft
Range("k3:k4").HorizontalAlignment = xlRight
Range("u:u,w:ad, ao:ap").HorizontalAlignment = xlCenter
Range("x:z, ad:ad").NumberFormatLocal = "#,##0.0_);(#,##0.0)"
Columns("u:u").NumberFormatLocal = "G/____"æ¦¡"
Range("aj:am").Font.ThemeColor = xlThemeColorDark1
[g3].Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font
        .Color = -16776961
    End With
    
[t9].Select
With Selection.Font
        .Color = -16776961
End With

End Sub


Sub Macro20()
'
' mikeon _ 2016/1/12 _____ªº¥¨¶°
' __·s

'
Dim NumberOfIE As Integer
NumberOfIE = 1 ' for control of creating IE
Sheets("__").Range("zz1") = NumberOfIE ' for set new IEIE
Sheets("__").Range("zz1") = NumberOfIE  'for set new IEIE
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)

Sheets("__").Selectct
Application.ScreenUpdating = False
Application.StatusBar = ""
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Application.Calculation = xlManual

Range("j6:j7").Select
With Selection.Font
    .Color = -16776961
End With
Range("j6:j7").HorizontalAlignment = xlLeft

Range("a:a, v:v").NumberFormatLocal = "@"
Application.ScreenUpdating = True
[j6] = "___......".."
[b2].Select
Application.ScreenUpdating = False

i = 2 '__§R
Do Until Cells(i, 22) = ""

    If UCase(Cells(i, 21)) = "X" Then
        Application.ScreenUpdating = True
        [j6] = "______......"....."
        Application.Wait Now() + TimeValue("00:00:01")
        Application.ScreenUpdating = False
        
        Application.Calculation = xlManual
        
        k = 2: j = 2: p = 1
        Do Until Cells(k, 22) = "" And p > 20 '________20_£¶W¹L20¦C
           If Cells(k, 22) = "" Then
              p = p + 1: GoTo kk
           End If
           Sheets("__").Cells(j, 81) = Cells(k, 21)1)
           If Trim(Cells(k, 28)) = "_" Or UCase(Trim(Cells(k, 28))) = "CN" Then Cells(k, 22) = Format(Cells(k, 22), "000000"))
           If Trim(Cells(k, 28)) = "_" Or UCase(Trim(Cells(k, 28))) = "HK" Then Cells(k, 22) = Format(Cells(k, 22), "0000"))
           Sheets("__").Cells(j, 82) = Cells(k, 22) '__ªÑ¦W
           Sheets("__").Cells(j, 83) = Cells(k, 23) '____´Á³ø¹S
           Sheets("__").Cells(j, 84) = Cells(k, 24) '__ªÑ»ù
           Sheets("__").Cells(j, 85) = Cells(k, 25) ' _ ²Q
           Sheets("__").Cells(j, 86) = Cells(k, 26) '_'¶Q
           'Sheets("__").Cells(j, 87) = Cells(k, 27) '__§Ç¸¹
           Sheets("__").Cells(j, 88) = Cells(k, 28) '__ªÑ¥«
           Sheets("__").Cells(j, 89) = Cells(k, 29) '__¤ÀÃþ
           Sheets("__").Cells(j, 90) = Cells(k, 30) '__¥«­È
           Sheets("__").Cells(j, 91) = Cells(k, 31) '__¤¶²Ð
           Sheets("__").Cells(j, 96) = Cells(k, 36) '__¤é´Á
           Sheets("__").Cells(j, 97) = Cells(k, 37) '____´Á³ø¹S
           Sheets("__").Cells(j, 98) = Cells(k, 38) '__ªÑ»ù
           Sheets("__").Cells(j, 99) = Cells(k, 39) '____­ÈªÑ»ù
           p = 1: j = j + 1
kk:     k = k + 1
        Loop
        GoTo __§R
    End If

    i = i + 1
Loop

GoTo __§R

__:R:
Application.Calculation = xlManual

k = 2
Do Until Sheets("__").Cells(k, 82) = """"

If UCase(Sheets("__").Cells(k, 81)) = "X" Thenen
Application.StatusBar = "___ " & k - 1 & " _" ¤ä"
i = 1: j = 1

Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i
If Trim(Cells(i + 1, 4)) = "_" Then Cells(i, 1) = Format(Cells(i, 1), "000000"))
If Trim(Cells(i + 1, 4)) = "_" Then Cells(i, 1) = Format(Cells(i, 1), "0000"))

If Trim(Sheets("__").Cells(k, 82)) = Trim(Cells(i, 1)) Thenen

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 29
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 24
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 28
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 23
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If UCase(Trim(Sheets("__").Cells(k, 88))) = UCase(Trim(Cells(i - 1, 1))) Thenen
g = i - 2
h = i + 24
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

End If
End If

i = i + 1
Loop

End If

k = k + 1
Loop

k = 2: j = 2: p = 1
Do Until Sheets("__").Cells(k, 82) = """"
If UCase(Sheets("__").Cells(k, 81)) = "X" Then GoTo dddd
Cells(j, 21) = Sheets("__").Cells(k, 81)1)
Cells(j, 22) = Sheets("__").Cells(k, 82) '__ªÑ¦W
Cells(j, 23) = Sheets("__").Cells(k, 83) '____´Á³ø¹S
Cells(j, 24) = Sheets("__").Cells(k, 84) '__ªÑ»ù
Cells(j, 25) = Sheets("__").Cells(k, 85) '_'²Q
Cells(j, 26) = Sheets("__").Cells(k, 86) '_'¶Q
'Cells(j, 27) = Sheets("__").Cells(k, 87) '__§Ç¸¹
Cells(j, 28) = Sheets("__").Cells(k, 88) '__ªÑ¥«
Cells(j, 29) = Sheets("__").Cells(k, 89) '__¤ÀÃþ
Cells(j, 30) = Sheets("__").Cells(k, 90) '__¥«­È
Cells(j, 31) = Sheets("__").Cells(k, 91) '__¤¶²Ð
Cells(j, 36) = Sheets("__").Cells(k, 96) '__¤é´Á
Cells(j, 37) = Sheets("__").Cells(k, 97) '____´Á³ø¹S
Cells(j, 38) = Sheets("__").Cells(k, 98) '__ªÑ»ù
Cells(j, 39) = Sheets("__").Cells(k, 99) '____­ÈªÑ»ù
j = j + 1
dd: k = k + 1
Loop

If k > j Then
Set a = Range(Cells(j, 21), Cells(k, 38))
a.Clear
End If


Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Range("cc:cz").Delete

[a118].Select
Sheets("__").Selectct
[b2] = ""

__:R:
Application.Calculation = xlManual

i = 2 'V__¼W
Do Until Cells(i, 22) = "" '(_______W¸ò­«½Æ
If Trim(Cells(i, 28)) = "_" Or UCase(Trim(Cells(i, 28))) = "CN" Then Cells(i, 22) = Format(Cells(i, 22), "000000"))
If Trim(Cells(i, 28)) = "_" Or UCase(Trim(Cells(i, 28))) = "HK" Then Cells(i, 22) = Format(Cells(i, 22), "0000"))

If IsError(Cells(i, 23)) Then Cells(i, 23) = "_" '_____³ø¹S²v
If Cells(i, 23) <> "" Then GoTo ii '___s¼W

k = 2
Do Until Cells(k, 22) = "" '(___«½Æ

If i <> k And UCase(Trim(Cells(i, 22))) = UCase(Trim(Cells(k, 22))) And Cells(i, 28) = "" Then '(________ªÑ¥«ªÅ¥Õ
If Cells(k, 28) = "_" Then  '__«½Æ
Cells(i, 22) = ""
GoTo ii
End If
If Cells(k, 28) = "_" Then '__«½Æ
Cells(i, 22) = ""
GoTo ii
End If
If Cells(k, 28) = "_" Then  '__«½Æ
Cells(i, 22) = ""
GoTo ii
End If
End If '________)Ñ¥«ªÅ¥Õ)

If i <> k And UCase(Trim(Cells(i, 22))) = UCase(Trim(Cells(k, 22))) And UCase(Trim(Cells(i, 28))) = UCase(Trim(Cells(k, 28))) Then '(__________¥«´ä©M¥þ²y
Cells(i, 22) = "" '__½Æ
GoTo ii
End If '__________)«´ä©M¥þ²y)

k = k + 1
Loop '___)½Æ)

ii: i = i + 1
Loop '_______)¸ò­«½Æ)


i = 2: p = 2: k = 1 '__VzV
Do Until i - p > 200
If Cells(i, 22) <> "" Then
p = i: k = k + 1
Cells(k, 21) = Cells(i, 21)
Cells(k, 22) = Cells(i, 22) '__¦W
Cells(k, 23) = Cells(i, 23) '____³ø¹S
Cells(k, 24) = Cells(i, 24) '__»ù
Cells(k, 25) = Cells(i, 25) '_Q
Cells(k, 26) = Cells(i, 26) '_Q
'Cells(k, 27) = Cells(i, 27) '__¸¹
Cells(k, 28) = Cells(i, 28) '__¥«
Cells(k, 29) = Cells(i, 29) '__Ãþ
Cells(k, 30) = Cells(i, 30) '__­È
Cells(k, 31) = Cells(i, 31) '__²Ð
Cells(k, 36) = Cells(i, 36) '__´Á
Cells(k, 37) = Cells(i, 37) '____³ø¹S
Cells(k, 38) = Cells(i, 38) '__»ù
Cells(k, 39) = Cells(i, 39) '____ªÑ»ù

End If
i = i + 1
Loop

i = k + 1: p = 1 '____Vh¾lV
Do Until Cells(i, 22) = "" And p > 20
If Cells(i, 22) = "" Then p = p + 1
Cells(i, 21) = ""
Cells(i, 22) = ""
Cells(i, 23) = ""
Cells(i, 24) = ""
Cells(i, 25) = ""
Cells(i, 26) = ""
Cells(i, 27) = ""
Cells(i, 28) = ""
Cells(i, 29) = ""
Cells(i, 30) = ""
Cells(i, 31) = ""
Cells(i, 36) = ""
Cells(i, 37) = ""
Cells(i, 38) = ""
Cells(i, 39) = ""
i = i + 1
Loop


Application.Calculation = xlManual

i = 1: j = 1 '________ j±`§Q³B j
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i
End If
i = i + 1
Loop

c = 1 '_______h»¡©ú¦C
Do Until Cells(c, 1) = ""
c = c + 1
Loop
c = c - 1

i = 2: x = 0 '_____s´X¤ä
Do Until Cells(i, 22) = ""
If IsError(Cells(i, 23)) Then Cells(i, 23) = "_"""
If Cells(i, 21) <> "" And Cells(i, 22) <> "" And Cells(i, 28) <> "" Then x = x + 1 '__¦³
If Cells(i, 22) <> "" And Cells(i, 23) = "" Then x = x + 1 '__¼W
i = i + 1
Loop

Application.ScreenUpdating = True
[j6] = "___......_ " & x & " _"" ¤ä"
[b2].Select
Application.Wait Now() + TimeValue("00:00:01")

With Application '_____}¼sºÖ
        .CutCopyMode = False
        .Calculation = xlManual
End With

k = 2: Y = 0: Z = c - 2
Do Until Cells(k, 22) = ""
If Cells(k, 21) = "" And Trim(Cells(k, 25)) <> "" Then GoTo ss
[j7] = Cells(k, 22)
Application.ScreenUpdating = False
[t3] = Cells(k, 21): Cells(k, 21) = ""

For i = c To j
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
If Trim(Cells(k, 28)) = "_" And Trim(Cells(k, 22)) = Trim(Cells(i, 1)) Thenn
TWW: With Application '_____}¼sºÖ
        .CutCopyMode = False
        .Calculation = xlAutomatic
End With
Sheets("__").Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = Z - 2
Cells(Z, 1).Select '_______@¤ä¦ì¸m
Application.ScreenUpdating = True
Y = Y + 1
Cells(i - 2, 1) = "__ " & y & " / " & x x

Application.ScreenUpdating = False
Sheets("__").[a2] = Cells(k, 22)2)

If Sheets("__").[m1] = "N" Then '_______ø©|¥¼¶i¨Ó
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _______"³ø©|¥¼¶i¨Ó"
Sheets("__").[m1] = """"
GoTo xx
End If
Sheets("__").[m1] = """"

If Sheets("__").[af2] = "" Thenen
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _____"¤£¨ì°]³ø"
GoTo stopp
End If

Application.Calculation = xlManual

Sheets("__").Range("A1:R29").Copypy
Sheets("__").[y16].Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    
Sheets("__").Selectct
Cells(i - 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(k, 23) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 24) = Cells(i + 1, 11) '__»ù
Cells(k, 25) = Cells(i + 3, 11) '_Q
Cells(k, 26) = Cells(i + 5, 11) '_Q
Cells(k, 30) = Cells(i - 2, 2) '__­È
If Cells(k, 30) = "" Then Cells(k, 30) = 1
Cells(k, 31) = Sheets("__").Cells(1, 1)1)
Cells(k, 36) = Cells(i - 1, 9) '__´Á
Cells(k, 37) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 38) = Cells(i + 1, 11) '__»ù
Cells(k, 39) = Cells(i - 2, 9) '____ªÑ»ù
'If y Mod [n5] = 0 Then
'Application.StatusBar = "_____"ÀÉ¤¤"
'ActiveWorkbook.Save
'End If
Z = i '_______@¤ä¦ì¸m
[j6] = ""
'Exit for
GoTo stopp
End If

If UCase(Trim(Cells(k, 28))) = "US" Then Cells(k, 28) = "_"""
If Trim(Cells(k, 28)) = "_" And Trim(Cells(k, 22)) = Trim(Cells(i, 1)) Thenn
USS: With Application '_____}¼sºÖ
        .CutCopyMode = False
        .Calculation = xlAutomatic
End With
Sheets("__").Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = Z - 2
Cells(Z, 1).Select '_______@¤ä¦ì¸m
Application.ScreenUpdating = True
Y = Y + 1
Cells(i - 2, 1) = "__ " & y & " / " & x x

Application.ScreenUpdating = False
Sheets("__").[a2] = Cells(k, 22)2)

If Sheets("__").[m1] = "N" Then '_______ø©|¥¼¶i¨Ó
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _______"³ø©|¥¼¶i¨Ó"
Sheets("__").[m1] = """"
GoTo xx
End If
Sheets("__").[m1] = """"

NumberOfIE = NumberOfIE + 1
Sheets("__").Range("zz1") = NumberOfIE: Sheets("__").Range("zz1") = NumberOfIEOfIE

If Sheets("__").[ae2] = "" Thenen
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _____"¤£¨ì°]³ø"
GoTo stopp
End If

Application.Calculation = xlManual

Sheets("__").Range("A1:R24").Copypy
Sheets("__").[y16].Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    
Sheets("__").Selectct
Cells(i - 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(k, 23) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 24) = Cells(i + 1, 11) '__»ù
Cells(k, 25) = Cells(i + 3, 11) '_Q
Cells(k, 26) = Cells(i + 5, 11) '_Q
Cells(k, 30) = Cells(i - 2, 2) '__­È
If Cells(k, 30) = "" Then Cells(k, 30) = 1
Cells(k, 31) = Sheets("__").Cells(1, 1)1)
Cells(k, 36) = Cells(i - 1, 9) '__´Á
Cells(k, 37) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 38) = Cells(i + 1, 11) '__»ù
Cells(k, 39) = Cells(i - 2, 9) '____ªÑ»ù
'If y Mod [n5] = 0 Then
'Application.StatusBar = "_____"ÀÉ¤¤"
'ActiveWorkbook.Save
'End If
Z = i
[j6] = ""
'Exit for
GoTo stopp
End If

If UCase(Trim(Cells(k, 28))) = "HK" Then Cells(k, 28) = "_"""
If Trim(Cells(k, 28)) = "_" And Trim(Cells(k, 22)) = Trim(Cells(i, 1)) Thenn
HKK: With Application '_____}¼sºÖ
        .CutCopyMode = False
        .Calculation = xlAutomatic
End With
Sheets("__").Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = Z - 2
Cells(Z, 1).Select '_______@¤ä¦ì¸m
Application.ScreenUpdating = True
Y = Y + 1
Cells(i - 2, 1) = "__ " & y & " / " & x x

Application.ScreenUpdating = False
Sheets("__").[a2] = Cells(k, 22)2)

If Sheets("__").[m1] = "N" Then '_______ø©|¥¼¶i¨Ó
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _______"³ø©|¥¼¶i¨Ó"
Sheets("__").[m1] = """"
GoTo xx
End If
Sheets("__").[m1] = """"

If Sheets("__").[ae2] = "" Thenen
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _____"¤£¨ì°]³ø"
GoTo stopp
End If

Application.Calculation = xlManual

Sheets("__").Range("A1:R28").Copypy
Sheets("__").[y16].Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    
Sheets("__").Selectct
Cells(i - 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(k, 23) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 24) = Cells(i + 1, 11) '__»ù
Cells(k, 25) = Cells(i + 3, 11) '_Q
Cells(k, 26) = Cells(i + 5, 11) '_Q
Cells(k, 30) = Cells(i - 2, 2) '__­È
If Cells(k, 30) = "" Then Cells(k, 30) = 1
Cells(k, 31) = Sheets("__").Cells(1, 1)1)
Cells(k, 36) = Cells(i - 1, 9) '__´Á
Cells(k, 37) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 38) = Cells(i + 1, 11) '__»ù
Cells(k, 39) = Cells(i - 2, 9) '____ªÑ»ù
'If y Mod [n5] = 0 Then
'Application.StatusBar = "_____"ÀÉ¤¤"
'ActiveWorkbook.Save
'End If
Z = i
[j6] = ""
'Exit for
GoTo stopp
End If

If UCase(Trim(Cells(k, 28))) = "CN" Then Cells(k, 28) = "_"""
If Trim(Cells(k, 28)) = "_" And Trim(Cells(k, 22)) = Trim(Cells(i, 1)) Thenn
CNN: With Application '_____}¼sºÖ
        .CutCopyMode = False
        .Calculation = xlAutomatic
End With
Sheets("__").Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = Z - 2
Cells(Z, 1).Select '_______@¤ä¦ì¸m
Application.ScreenUpdating = True
Y = Y + 1
Cells(i - 2, 1) = "__ " & y & " / " & x x

Application.ScreenUpdating = False
Sheets("__").[a2] = Cells(k, 22)2)

If Sheets("__").[m1] = "N" Then '_______ø©|¥¼¶i¨Ó
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _______"³ø©|¥¼¶i¨Ó"
Sheets("__").[m1] = """"
GoTo xx
End If
Sheets("__").[m1] = """"

If Sheets("__").[ae2] = "" Thenen
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _____"¤£¨ì°]³ø"
GoTo stopp
End If

Application.Calculation = xlManual

Sheets("__").Range("A1:R24").Copypy
Sheets("__").[y16].Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1

Sheets("__").Selectct
Cells(i - 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(k, 23) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 24) = Cells(i + 1, 11) '__»ù
Cells(k, 25) = Cells(i + 3, 11) '_Q
Cells(k, 26) = Cells(i + 5, 11) '_Q
Cells(k, 30) = Cells(i - 2, 2) '__­È
If Cells(k, 30) = "" Then Cells(k, 30) = 1
Cells(k, 31) = Sheets("__").Cells(1, 1)1)
Cells(k, 36) = Cells(i - 1, 9) '__´Á
Cells(k, 37) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 38) = Cells(i + 1, 11) '__»ù
Cells(k, 39) = Cells(i - 2, 9) '____ªÑ»ù
'If y Mod [n5] = 0 Then
'Application.StatusBar = "_____"ÀÉ¤¤"
'ActiveWorkbook.Save
'End If
Z = i
[j6] = ""
'Exit for
GoTo stopp
End If

If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) Then
GBB: With Application '_____}¼sºÖ
        .CutCopyMode = False
        .Calculation = xlAutomatic
End With
Sheets("__").Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = Z - 2
Cells(Z, 1).Select '_______@¤ä¦ì¸m
Application.ScreenUpdating = True
Y = Y + 1
Cells(i - 2, 1) = "__ " & y & " / " & x x

Application.ScreenUpdating = False
Sheets("__").[a1] = Trim(Cells(k, 28))))
Sheets("__").[a2] = Cells(k, 22)2)

If Sheets("__").[m1] = "N" Then '_______ø©|¥¼¶i¨Ó
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _______"³ø©|¥¼¶i¨Ó"
Sheets("__").[m1] = """"
GoTo xx
End If
Sheets("__").[m1] = """"

NumberOfIE = NumberOfIE + 1
Sheets("__").Range("zz1") = NumberOfIE: Sheets("__").Range("zz1") = NumberOfIEOfIE

If Sheets("__").[ak2] = "" Thenen
Sheets("__").[j6] = Sheets("__").Cells(k, 22) & " _____"¤£¨ì°]³ø"
GoTo stopp
End If

Application.Calculation = xlManual

Sheets("__").Range("A1:R24").Copypy
Sheets("__").[ae16].Selectct
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
  
Sheets("__").Selectct
Cells(i - 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(k, 23) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 24) = Cells(i + 1, 11) '__»ù
Cells(k, 25) = Cells(i + 3, 11) '_Q
Cells(k, 26) = Cells(i + 5, 11) '_Q
Cells(k, 30) = Cells(i - 2, 2) '__­È
If Cells(k, 30) = "" Then Cells(k, 30) = 1
Cells(k, 31) = Sheets("__").Cells(1, 1)1)
Cells(k, 36) = Cells(i - 1, 9) '__´Á
Cells(k, 37) = Cells(i + 2, 11) '____³ø¹S
Cells(k, 38) = Cells(i + 1, 11) '__»ù
Cells(k, 39) = Cells(i - 2, 9) '____ªÑ»ù
'If y Mod [n5] = 0 Then
'Application.StatusBar = "_____"ÀÉ¤¤"
'ActiveWorkbook.Save
'End If
Z = i
[j6] = ""
'Exit for
GoTo stopp
End If

End If

Next i

If Cells(k, 28) = "" And IsNumeric(Left(Cells(k, 22), 1)) And Len(Cells(k, 22)) = 4 Then Cells(k, 28) = "_"""
If Cells(k, 28) = "" And IsNumeric(Left(Cells(k, 22), 1)) And Len(Cells(k, 22)) = 6 Then Cells(k, 28) = "_"""
If Cells(k, 28) = "" Then Cells(k, 28) = "_"""
i = c: j = c
Do Until i - j > 50 '_V___A£¦bA
If IsError(Cells(i, 9)) Then
j = i
GoTo v50
End If
If Trim(Cells(i, 9)) <> "" Then j = i
v50: i = i + 1
Loop
i = j + 5
If Trim(Cells(k, 28)) = "_" Then GoTo TWWW
If Trim(Cells(k, 28)) = "_" Then GoTo USSS
If Trim(Cells(k, 28)) = "_" Then GoTo HKKK
If Trim(Cells(k, 28)) = "_" Then GoTo CNNN
If Trim(Cells(k, 28)) <> "_" And Trim(Cells(k, 28)) = "_" And Trim(Cells(k, 28)) = "_" And Trim(Cells(k, 28)) = "_" Then GoTo GBB GBB

ss: k = k + 1
Loop


stopp: '__¸¹
Sheets("__").Selectct
i = 1: j = 1: b = 0 '_____Ã´X¤ä
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i: b = b + 1
End If
i = i + 1
Loop

Application.Calculation = xlManual

i = 1: j = 1: n = 0
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
n = n + 1: j = i
Cells(i - 2, 1) = "  " & n & " / " & b & " _"""

k = 2
Do Until Cells(k, 22) = ""
If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And Trim(Cells(k, 28)) = Trim(Cells(i + 1, 4)) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
Cells(i - 2, 2) = Cells(k, 30) '__­È
Cells(k, 27) = n
Exit Do
End If

If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And UCase(Trim(Cells(k, 28))) = UCase(Trim(Cells(i - 1, 1))) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
Cells(i - 2, 2) = Cells(k, 30) '__­È
Cells(k, 27) = n
Exit Do
End If
k = k + 1
Loop

End If
i = i + 1
Loop

xx: Sheets("__").[t3] = "": Sheets("__").[j7] = "" = ""
If Right(Sheets("__").[j6], 2) = "__" Then GoTo xxa xxa
If Right(Sheets("__").[j6], 5) = "_____" Then GoTo xxaoTo xxa
Sheets("__").[j6] = """"
xxa: Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop


Call fm7
Sheets("__").Selectct
Call ProtectSheet(ActiveSheet)
Sheets("__").Selectct
[dz100].Select
ActiveWindow.ScrollColumn = 5: ActiveWindow.ScrollRow = 1

   With Application '_____}¼sºÖ
        .CutCopyMode = False
        .Calculation = xlAutomatic
    End With
    
 '----delete IE---------------

       ' Sheets("__").Range("zz1") = """"
       ' Sheets("__").Range("zz1") = """"
        
        'Call DelIE
 '-----------------------------
Application.StatusBar = "__"¨"



End Sub


Sub Macro21()
'
' mikeon _ 20168/10/29 _____ªº¥¨¶°
' ____ªÑ»ù

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)

Sheets("__").Selectct


Application.ScreenUpdating = True
[t9] = "___......".."
Application.Wait Now() + TimeValue("00:00:01")
Application.ScreenUpdating = False


'___ÀÃþ
b = 2 '_____Ã´X¤ä
Do Until Cells(b, 22) = ""
Cells(b, 39) = Cells(b, 24) '____ªÑ»ù
b = b + 1
Loop
b = b - 2

i = 1: j = 1: n = 0
Do Until i - j > 100 Or n > b
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
n = n + 1: j = i

k = 2
Do Until Cells(k, 22) = ""
If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And Trim(Cells(k, 28)) = Trim(Cells(i + 1, 4)) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
If Cells(k, 30) = "" Then Cells(k, 30) = 1
Cells(i - 2, 2) = Cells(k, 30) '__­È
Exit Do
End If

If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And UCase(Trim(Cells(k, 28))) = UCase(Trim(Cells(i - 1, 1))) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
If Cells(k, 30) = "" Then Cells(k, 30) = 1
Cells(i - 2, 2) = Cells(k, 30) '__­È
Exit Do
End If
k = k + 1
Loop

End If
i = i + 1
Loop

Range("ai:aj").Select
    With Selection
    .NumberFormatLocal = "yyyy/m/d"
End With

Range("T5").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=TODAY(),""[__] ___"",""[__] ___"")"¥¼§ó·s"")"

If [t4] = Date Then '______©wªÑ»ù

i = 2
Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""
If Not (IsNumeric(Cells(i, 24))) Then GoTo ip
If Not (IsNumeric(Cells(i, 38))) Then GoTo ip
If Cells(i, 38) <= 0 Then GoTo ip
If Not (IsNumeric(Cells(i, 37))) Then GoTo ip
If Cells(i, 24) = Cells(i, 38) Then GoTo ip
Cells(i, 23) = Cells(i, 37) - ((Cells(i, 24) / Cells(i, 38)) ^ (1 / 8) - 1)
Call w2k(i)
ip: i = i + 1
Loop

GoTo aa
End If


Macro22


'__ªÑ
Application.Calculation = xlManual

i = 2
Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""
If Cells(i, 28) = "_" Or Cells(i, 28) = "TW" Then GoTo TWWW
i = i + 1
Loop
GoTo _~

TWW: '__¥«
Application.StatusBar = "________"«ªÑ»ù¤¤"
On Error GoTo err1
Dim url As String
       url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date=&type=ALLBUT0999"
       Call ConnectMarketWatch(url, Cells(1, 40), 2)
       
Application.Calculation = xlManual

Range("an:an").NumberFormatLocal = "@"
i = 2
Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""

If Not (IsNumeric(Cells(i, 30))) Then Cells(i, 30) = 1
If Not (IsNumeric(Cells(i, 24))) Then Cells(i, 24) = 0
If Not (IsNumeric(Cells(i, 39))) Then Cells(i, 39) = Cells(i, 24)

If Cells(i, 28) = "_" Or Cells(i, 28) = "TW" Thenn
Application.StatusBar = i - 1
j = 2
Do Until Cells(j, 40) = "" And Cells(j + 1, 40) = "" And Cells(j + 2, 40) = "" And Cells(j + 3, 40) = ""
If Trim(Cells(i, 22)) = Trim(Cells(j, 40)) And Left(Cells(i, 31), 1) = Left(Cells(j, 41), 1) Then
If Not (IsNumeric(Cells(j, 40 + 8))) Then GoTo iL
Cells(i, 24) = Cells(j, 40 + 8)
If Not (IsNumeric(Cells(i, 38))) Then GoTo iL
If Cells(i, 38) <= 0 Then GoTo iL
If Not (IsNumeric(Cells(i, 37))) Then GoTo iL
Cells(i, 23) = Cells(i, 37) - ((Cells(j, 40 + 8) / Cells(i, 38)) ^ (1 / 8) - 1)

If Cells(i, 39) <> 0 Then Cells(i, 30) = Cells(i, 30) * Cells(i, 24) / Cells(i, 39) '__­È
Call w2k(i)
GoTo iL
End If
j = j + 1
Loop
End If
iL: i = i + 1
Loop

err1:
Range("an:bk").ClearContents


'__Âd
Application.Calculation = xlManual

Application.StatusBar = "________"dªÑ»ù¤¤"
On Error GoTo err2
       url = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=htm"
       Call ConnectMarketWatch(url, Cells(1, 40), 2)
       
Application.Calculation = xlManual

Range("an:an").NumberFormatLocal = "@"
i = 2
Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""

If Not (IsNumeric(Cells(i, 30))) Then Cells(i, 30) = 1
If Not (IsNumeric(Cells(i, 24))) Then Cells(i, 24) = 0
If Not (IsNumeric(Cells(i, 39))) Then Cells(i, 39) = Cells(i, 24)

If Cells(i, 28) = "_" Or Cells(i, 28) = "TW" Thenn
Application.StatusBar = i - 1
j = 2
Do Until Cells(j, 40) = "" And Cells(j + 1, 40) = ""
If Trim(Cells(i, 22)) = "6201" Then GoTo io '____¥X¿ù
If Trim(Cells(i, 22)) = Trim(Cells(j, 40)) And Left(Cells(i, 31), 1) = Left(Cells(j, 41), 1) Then
If Not (IsNumeric(Cells(j, 42))) Then GoTo io
Cells(i, 24) = Cells(j, 42)
If Not (IsNumeric(Cells(i, 38))) Then GoTo io
If Cells(i, 38) <= 0 Then GoTo io
If Not (IsNumeric(Cells(i, 37))) Then GoTo io
Cells(i, 23) = Cells(i, 37) - ((Cells(j, 42) / Cells(i, 38)) ^ (1 / 8) - 1)
If Cells(i, 39) <> 0 Then Cells(i, 30) = Cells(i, 30) * Cells(i, 24) / Cells(i, 39) '__­È
Call w2k(i)
GoTo io
End If
j = j + 1
Loop
End If
io: i = i + 1
Loop

err2:
Range("an:bk").ClearContents


_::
Application.Calculation = xlManual

Application.StatusBar = "_______ 8 _"®É 8 ¬í"
On Error GoTo err3
    url = [m5]
    Call ConnectMarketWatch(url, Cells(1, 40), 1)
    
i = 2: Dim i22, ii22 As String
Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""

If Not (IsNumeric(Cells(i, 30))) Then Cells(i, 30) = 1
If Not (IsNumeric(Cells(i, 24))) Then Cells(i, 24) = 0
If Not (IsNumeric(Cells(i, 39))) Then Cells(i, 39) = Cells(i, 24)

If Trim(Cells(i, 28)) <> "_" And Trim(Cells(i, 28)) <> "TW" Thenn
Application.StatusBar = i - 1

j = 2
Do Until Cells(j, 41) = "" And Cells(j + 1, 41) = ""
If [m5] = "" Then Exit Do
i22 = Cells(j, 41)
If Trim(Cells(i, 28)) = "_" Or UCase(Trim(Cells(i, 28))) = "HK" Thenn
i22 = Format(Right(i22, 4), "0000")
Cells(i, 22) = Format(Right(Cells(i, 22), 4), "0000")
End If

If Trim(Cells(i, 28)) = "_" Or UCase(Trim(Cells(i, 28))) = "CN" Thenn
i22 = Format(Right(i22, 6), "000000")
Cells(i, 22) = Format(Right(Cells(i, 22), 6), "000000")
End If

If Trim(Cells(i, 28)) = "_" Or UCase(Trim(Cells(i, 28))) = "JP" Thenn
i22 = Right(i22, 4)
End If

ii22 = mkw(Cells(i, 22))
If Trim(UCase(ii22)) = Trim(UCase(i22)) Then
If Not (IsNumeric(Cells(j, 42))) Then GoTo iw
Cells(i, 24) = Cells(j, 42)
If Not (IsNumeric(Cells(i, 38))) Then GoTo iw
If Cells(i, 38) <= 0 Then GoTo iw
If Not (IsNumeric(Cells(i, 37))) Then GoTo iw
Cells(i, 23) = Cells(i, 37) - ((Cells(j, 42) / Cells(i, 38)) ^ (1 / 8) - 1)

If Cells(i, 39) <> 0 Then Cells(i, 30) = Cells(i, 30) * Cells(i, 24) / Cells(i, 39) '__­È

Call w2k(i)
GoTo iw
End If
j = j + 1
Loop
Cells(i, 21) = "_M5"""
End If
iw: i = i + 1
Loop
err3:
Range("an:bk").ClearContents


i = 2 '_A__«½Æ
Do Until Cells(i, 22) = "" '(_ii
If Trim(Cells(i, 28)) = "_" Or UCase(Trim(Cells(i, 28))) = "CN" Then Cells(i, 22) = Format(Cells(i, 22), "000000"))
If Trim(Cells(i, 28)) = "_" Or UCase(Trim(Cells(i, 28))) = "HK" Then Cells(i, 22) = Format(Cells(i, 22), "0000"))
If IsError(Cells(i, 23)) Then Cells(i, 23) = "_" '_____³ø¹S²v
If UCase(Trim(Cells(i, 21))) = "X" Then GoTo ii

k = 2
Do Until Cells(k, 22) = "" '(_kk

If i <> k And UCase(Trim(Cells(i, 22))) = UCase(Trim(Cells(k, 22))) And UCase(Trim(Cells(i, 28))) = UCase(Trim(Cells(k, 28))) Then   '(_______AªÑ¥«¦P
Cells(i, 29) = Cells(k, 29) '__Ãþ
If Cells(i, 35) >= Cells(k, 35) Then
Cells(k, 21) = "X"
Else
Cells(i, 21) = "X"
End If
GoTo ii
End If
k = k + 1
Loop '_k))

ii: i = i + 1
Loop '_i))

Range("ai:ai").ClearContents


'_____«½ÆªÑ
Range("j6:j7").Select
With Selection.Font
    .Color = -16776961
End With
Range("j6:j7").HorizontalAlignment = xlLeft

Range("a:a, v:v").NumberFormatLocal = "@"

[dz100].Select
i = 2
Do Until Cells(i, 22) = "" '(_VV
    If UCase(Trim(Cells(i, 21))) = "X" Then '(__§R
        Application.ScreenUpdating = True
        [j6] = "______......"....."
        Application.ScreenUpdating = False
        
        Application.Calculation = xlManual
        
        k = 2: j = 2: p = 1
        Do Until Cells(k, 22) = "" And p > 20 '________20_£¶W¹L20¦C
           If Cells(k, 22) = "" Then
              p = p + 1: GoTo kk
           End If
           Sheets("__").Cells(j, 81) = Cells(k, 21)1)
           Sheets("__").Cells(j, 82) = Cells(k, 22) '__ªÑ¦W
           Sheets("__").Cells(j, 83) = Cells(k, 23) '____´Á³ø¹S
           Sheets("__").Cells(j, 84) = Cells(k, 24) '__ªÑ»ù
           Sheets("__").Cells(j, 85) = Cells(k, 25) ' _ ²Q
           Sheets("__").Cells(j, 86) = Cells(k, 26) '_'¶Q
           Sheets("__").Cells(j, 87) = Cells(k, 27) '__§Ç¸¹
           Sheets("__").Cells(j, 88) = Cells(k, 28) '__ªÑ¥«
           Sheets("__").Cells(j, 89) = Cells(k, 29) '__¤ÀÃþ
           Sheets("__").Cells(j, 90) = Cells(k, 30) '__¥«­È
           Sheets("__").Cells(j, 91) = Cells(k, 31) '__¤¶²Ð
           Sheets("__").Cells(j, 96) = Cells(k, 36) '__¤é´Á
           Sheets("__").Cells(j, 97) = Cells(k, 37) '____´Á³ø¹S
           Sheets("__").Cells(j, 98) = Cells(k, 38) '__ªÑ»ù
           Sheets("__").Cells(j, 99) = Cells(k, 39) '____­ÈªÑ»ù
           p = 1: j = j + 1
kk:     k = k + 1
        Loop
        GoTo __§R
    End If '__)R)

    i = i + 1
Loop '_V))

GoTo sr

__:R:
Application.Calculation = xlManual

k = 2
Do Until Sheets("__").Cells(k, 82) = "" '(_81ä81

If UCase(Sheets("__").Cells(k, 81)) = "X" Thenen
Application.StatusBar = "___ " & k - 1 & " _" ¤ä"
i = 1: j = 1

Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i

If Trim(Sheets("__").Cells(k, 82)) = Trim(Cells(i, 1)) And Trim(Sheets("__").Cells(k, 87)) = Trim(Cells(i - 2, 1)) Then '_______W¦P¡B§Ç¸¹¦P

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 29
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 24
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 28
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If Trim(Sheets("__").Cells(k, 88)) = "_" Thenhen
g = i - 2
h = i + 23
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

If UCase(Trim(Sheets("__").Cells(k, 88))) = UCase(Trim(Cells(i - 1, 1))) Thenen
g = i - 2
h = i + 24
Rows(g & ":" & h).Delete Shift:=xlUp
Exit Do
End If

End If
End If

i = i + 1
Loop

End If

k = k + 1
Loop '_81))

k = 2: j = 2: p = 1
Do Until Sheets("__").Cells(k, 82) = """"
If UCase(Sheets("__").Cells(k, 81)) = "X" Then GoTo dddd
Cells(j, 21) = Sheets("__").Cells(k, 81)1)
Cells(j, 22) = Sheets("__").Cells(k, 82) '__ªÑ¦W
Cells(j, 23) = Sheets("__").Cells(k, 83) '____´Á³ø¹S
Cells(j, 24) = Sheets("__").Cells(k, 84) '__ªÑ»ù
Cells(j, 25) = Sheets("__").Cells(k, 85) '_'²Q
Cells(j, 26) = Sheets("__").Cells(k, 86) '_'¶Q
'Cells(j, 27) = Sheets("__").Cells(k, 87) '__§Ç¸¹
Cells(j, 28) = Sheets("__").Cells(k, 88) '__ªÑ¥«
Cells(j, 29) = Sheets("__").Cells(k, 89) '__¤ÀÃþ
Cells(j, 30) = Sheets("__").Cells(k, 90) '__¥«­È
Cells(j, 31) = Sheets("__").Cells(k, 91) '__¤¶²Ð
Cells(j, 36) = Sheets("__").Cells(k, 96) '__¤é´Á
Cells(j, 37) = Sheets("__").Cells(k, 97) '____´Á³ø¹S
Cells(j, 38) = Sheets("__").Cells(k, 98) '__ªÑ»ù
Cells(j, 39) = Sheets("__").Cells(k, 99) '____­ÈªÑ»ù
j = j + 1
dd: k = k + 1
Loop

If k > j Then
Set a = Range(Cells(j, 21), Cells(k, 38))
a.Clear
End If

Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Range("cc:cz").Delete

sr: '__¸¹
Sheets("__").Selectct

i = 1: j = 1: b = 0 '_____Ã´X¤ä
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i: b = b + 1
End If
i = i + 1
Loop

Application.Calculation = xlManual

i = 1: j = 1: n = 0
Do Until i - j > 100

If Left(Trim(Cells(i, 3)), 2) = "RO" Then

n = n + 1: j = i
Cells(i - 2, 1) = "  " & n & " / " & b & " _"""

k = 2
Do Until Cells(k, 22) = ""
If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And Trim(Cells(k, 28)) = Trim(Cells(i + 1, 4)) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
Cells(i - 2, 2) = Cells(k, 30) '__­È
Cells(k, 27) = n
Exit Do
End If

If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And UCase(Trim(Cells(k, 28))) = UCase(Trim(Cells(i - 1, 1))) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
Cells(i - 2, 2) = Cells(k, 30) '__­È
Cells(k, 27) = n
Exit Do
End If
k = k + 1
Loop

End If
i = i + 1
Loop

[j6] = "": [j7] = ""

Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop



'____----------------------------------------------------------


Application.StatusBar = "___......AC _" Äæ"

i = 2: Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""
Cells(i, 50) = Cells(i, 29)
i = i + 1
Loop

axx: [ax1] = 0 'ax = 50, an = 40
i = 2: Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""
If Cells(i, 50) <> "" Then
[ax1] = 1
Exit Do
End If
i = i + 1
Loop

If [ax1] = 1 Then
i = 2: Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = ""
If Cells(i, 50) = "" Then GoTo ii7
a = InStr(1, Cells(i, 50), "/")
If a > 0 Then 'PPP/
ax = Left(Cells(i, 50), a - 1)
Call stat(i, ax)
Cells(i, 50) = Mid(Cells(i, 50), a + 1, Len(Cells(i, 50)))
End If

If a = 0 Then 'PPP
ax = Cells(i, 50)
Call stat(i, ax)
Cells(i, 50) = ""
End If

ii7: i = i + 1
Loop

End If

If [ax1] = 1 Then GoTo axx


For k = 1 To 4
ax = ""
Application.StatusBar = "___......AE _" Äæ"
i = 2: Do Until Cells(i, 22) = "" And Cells(i + 1, 22) = "" 'AE _æ
If k = 1 And InStr(1, Cells(i, 31), "*") > 0 Then ax = Right(Cells(i, 31), Len(Cells(i, 31)) - InStr(1, Cells(i, 31), "*"))
If k = 2 And Cells(i, 28) = "_" And InStr(1, Cells(i, 31), "_") > 0 Then ax = Mid(Cells(i, 31), InStr(1, Cells(i, 31), "_") + 1, Len(Cells(i, 31)) - InStr(1, Cells(i, 31), "_"))G"))
If k = 3 And Cells(i, 28) = "_" And InStr(1, Cells(i, 31), "_") > 0 Then ax = Mid(Cells(i, 31), InStr(1, Cells(i, 31), "_") + 1, Len(Cells(i, 31)) - InStr(1, Cells(i, 31), "_"))G"))
If k = 4 And Cells(i, 28) = "_" And InStr(1, Cells(i, 31), "_") > 0 Then ax = Right(Cells(i, 31), Len(Cells(i, 31)) - InStr(1, Cells(i, 31), "*"))))
Call stat(i, ax)
i = i + 1
Loop
Next k


Cells(10, 40) = 0
i = 2: Do Until Cells(i, 22) = ""
Cells(10, 40) = Cells(10, 40) + Cells(i, 30)
i = i + 1
Loop


i = 11: Do Until Cells(i, 40) = ""
Cells(i, 42) = 100 * Cells(i, 41) / Cells(10, 40)
i = i + 1
Loop
If Cells(i, 40) = "" Then Cells(i, 41) = ""


Cells(9, 40) = "____ " & Format(b, "##,##0") & "_" "¤ä"
Cells(10, 40) = "______ " & Format(Cells(10, 40), "#,##0.0")#0.0")
Cells(10, 41) = "__": Cells(10, 42) = "%"%"
Range(Columns(48), Columns(60)).Clear

Columns(41).NumberFormatLocal = "#,##0.0_);(#,##0.0)"
Columns(42).NumberFormatLocal = "#,##0_);(#,##0)"


'----------------------------------------------------------------------

aa:
[t4] = Date
Call fm7
Sheets("__").Selectct
Call ProtectSheet(ActiveSheet)
Sheets("__").Selectct
[t9] = ""
[dz100].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
Beep
   
End Sub

Private Sub stat(i, ax)

j = 11: Do Until Cells(j, 40) = ""
If Trim(LCase(ax)) = Trim(LCase(Cells(j, 40))) Then
Cells(j, 41) = Cells(j, 41) + Cells(i, 30)
GoTo ifc
End If
j = j + 1
Loop

Cells(j, 40) = LCase(ax)
Cells(j, 41) = Cells(i, 30)


ifc:
ax = ""


End Sub


Private Sub w2k(i)
    
k = 1: j = 1
Do Until k - j > 100
If Left(Trim(Cells(k, 3)), 2) = "RO" Then
j = k
If Trim(Cells(k + 1, 4)) = "_" Then Cells(k, 1) = Format(Cells(k, 1), "000000"))
If Trim(Cells(k + 1, 4)) = "_" Then Cells(k, 1) = Format(Cells(k, 1), "0000"))

If Trim(Cells(i, 22)) = Trim(Cells(k, 1)) And Cells(i, 28) = Trim(Cells(k + 1, 4)) Then
Cells(k - 1, 9) = Date
Cells(k - 1, 9).Select
    With Selection
    .NumberFormatLocal = "yyyy/m/d;@"
    .ShrinkToFit = True
End With
Cells(k + 2, 11) = Cells(i, 23) '____³ø¹S
Cells(k + 1, 11) = Cells(i, 24) '__»ù
Cells(k - 2, 9) = Cells(i, 39) '____ªÑ»ù
Exit Do
End If
End If

k = k + 1
Loop

End Sub



Sub Macro22()
'
' mikeon _ 2016/3/26 _____ªº¥¨¶°
' __³ò

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct
Application.Calculation = xlManual


Call ak2k
Range("u:am").ClearContents


[y1] = "_$"""
[Z1] = "_$"""
Range("a:a, v:v").Select
    Selection.NumberFormatLocal = "@"

i = 1: j = 1: b = 0 '_____Ã´X¤ä
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i: b = b + 1
End If
i = i + 1
Loop

i = 1: j = 1: k = 1: n = 0
Do Until i - j > 100 '(_AA
If Left(Trim(Cells(i, 3)), 2) = "RO" Then '(A_b
j = i: n = n + 1: k = k + 1
Cells(i - 2, 1) = n '_ä

Cells(k, 22) = Cells(i, 1) '__¦W
Cells(k, 23) = Cells(i + 2, 11) '_____ø¹S²v
Cells(k, 24) = Cells(i + 1, 11) '__»ù
Cells(k, 25) = Cells(i + 3, 11) '_Q
Cells(k, 26) = Cells(i + 5, 11) '_Q
Cells(k, 27) = n '__¸¹
Cells(k, 28) = Cells(i + 1, 4) '__¥«
Cells(k, 29) = Cells(i - 2, 4) '__Ãþ
Cells(k, 30) = Cells(i - 2, 2) '__­È
Cells(k, 31) = Cells(i - 1, 1) '__¤¶
Cells(k, 36) = Cells(i - 1, 9) '__´Á
Cells(k, 37) = Cells(i + 2, 11) '_____ø¹S²v
Cells(k, 38) = Cells(i + 1, 11) '__»ù
Cells(k, 39) = Cells(i - 2, 9) '____ªÑ»ù
If Trim(Cells(i + 1, 4)) <> "_" And Trim(Cells(i + 1, 4)) <> "_" And Trim(Cells(i + 1, 4)) <> "_" And Trim(Cells(i + 1, 4)) <> "_" ThenThen
Cells(k, 28) = Cells(i - 1, 1) '__¥«
Cells(k, 31) = Cells(i - 1, 2) '__¤¶
End If

Application.StatusBar = k - 1
For j = i To i + 100 '___]³ø
If Left(Cells(j, 5), 2) = "__" Or Cells(j, 5) = "Net" Then Exit For '____________2020____´«¡Aµ¥2020§R°£¥»¦C
If Left(Cells(j, 6), 2) = "__" Or Cells(j, 6) = "Net" Then Exit For '____________2020____´«¡Aµ¥2020§R°£¥»¦C
If Cells(j, 6) = Sheets("__").[f20] Or Cells(j, 6) = Sheets("__").[f18] Then Exit For For
Next j
Do Until Cells(j, 5) = "" And Cells(j + 1, 5) = ""
j = j + 1
Loop
j = j - 1
Cells(k, 21) = Cells(j, 1) '_____]³ø¤é
If Mid(Cells(j, 1), 2, 1) = "Q" Then '_x
If Left(Cells(j, 1), 1) = 1 Then Cells(k, 21) = "20" + Right(Cells(j, 1), 2) + "/3"
If Left(Cells(j, 1), 1) = 2 Then Cells(k, 21) = "20" + Right(Cells(j, 1), 2) + "/6"
If Left(Cells(j, 1), 1) = 3 Then Cells(k, 21) = "20" + Right(Cells(j, 1), 2) + "/9"
If Left(Cells(j, 1), 1) = 4 Then Cells(k, 21) = "20" + Right(Cells(j, 1), 2) + "/12"
End If
Cells(k, 35) = Cells(j - 1, 1) '______°]³ø¤é
If Mid(Cells(j - 1, 1), 2, 1) = "Q" Then '_x
If Left(Cells(j - 1, 1), 1) = 1 Then Cells(k, 35) = "20" + Right(Cells(j - 1, 1), 2) + "/3"
If Left(Cells(j - 1, 1), 1) = 2 Then Cells(k, 35) = "20" + Right(Cells(j - 1, 1), 2) + "/6"
If Left(Cells(j - 1, 1), 1) = 3 Then Cells(k, 35) = "20" + Right(Cells(j - 1, 1), 2) + "/9"
If Left(Cells(j - 1, 1), 1) = 4 Then Cells(k, 35) = "20" + Right(Cells(j - 1, 1), 2) + "/12"
End If

End If 'A_))

iii: i = i + 1
Loop '_A))

k = 2 '_____é¥¼¨ì
Do Until Cells(k, 22) = ""
If Not (IsNumeric(DateDiff("d", Cells(k, 35), Cells(k, 21)))) Then GoTo kk

If DateDiff("d", Cells(k, 21), Now) < 30 * 4 + 30 * 1.5 - 5 And DateDiff("d", Cells(k, 35), Cells(k, 21)) < 120 Then Cells(k, 21) = ""
If DateDiff("d", Cells(k, 21), Now) < 30 * 7 + 30 * 1.5 + 5 And DateDiff("d", Cells(k, 35), Cells(k, 21)) > 120 Then Cells(k, 21) = "" '_____b¦~³ø
If Trim(Cells(k, 28)) = "_" Or UCase(Trim(Cells(k, 28))) = "TW" Or UCase(Trim(Cells(k, 28))) = "DE" Or Trim(Cells(k, 28)) = "_" Or UCase(Trim(Cells(k, 28))) = "CN" Thenen
If DateDiff("d", Cells(k, 21), Now) < 30 * 4 + 30 * 1.5 + 2 And DateDiff("d", Cells(k, 35), Cells(k, 21)) < 100 Then Cells(k, 21) = ""
If DateDiff("d", Cells(k, 21), Now) < 30 * 4 + 30 * 3 + 2 And Right(Cells(k, 21), 1) = "9" Then Cells(k, 21) = ""
End If
kk: k = k + 1
Loop

Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop


[t4] = ""
Call fm7
Application.Calculation = xlAutomatic           '_____}¼sºÖ

   
End Sub


Sub Macro23()
'
' mikeon _ 2015/10/19 _____ªº¥¨¶°
' ____²MªÅ

'
Application.ScreenUpdating = False
Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²

Application.Calculation = xlManual

Sheets("__").Selectct
Sheets("__").[m5] = [m5]5]
Cells.Clear


[m5] = Sheets("__").[m5]5]
Sheets("__").[m5] = """"

Call fm7
[dz100].Select
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"

End Sub

Public Sub addlink()

'____Alex_ __  2016/5/1716/5/17
    
  
    
    Application.ScreenUpdating = False
    Dim companyid, companyname As String
    Dim companyid_row, i, m As Double
    Dim rng1, rng2 As Range
    
    Sheets("__").Selectct
    Set rng1 = Range("a1:a60000")
    Set rng2 = Range("b1:b60000")
    i = 2
    Do
        If Range("v" & i) = "" Then Exit Do
        
        companyid = Range("v" & i)
        If IsNumeric(companyid) = True Then companyid = Str(companyid)
        
        companyname = Range("ae" & i)
        
        If IsError(Application.Match(companyname, rng1, 0)) = False Then
          
            companyid_row = Application.Match(companyname, rng1, 0)
            
        ElseIf IsError(Application.Match(companyname, rng2, 0)) = False Then
            companyid_row = Application.Match(companyname, rng2, 0)
        
        Else
        
            GoTo nexti
            
        End If
        
        '  Debug.Print companyid, companyname, companyid_row
        
           
        m = 25 '__ªÑ
        'If Range("ab" & i) = "_" Then m = 25 '__xªÑ
        ' add hyperlink for specific company
        Range("V" & i).Select
        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="__!A" & companyid_row + m, TextToDisplay:=companyidid
             
        ' return to stock list
        Range("C" & companyid_row - 1).Select
             
            ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="__!X" & i, TextToDisplay:="Back"k"
             
nexti:

        i = i + 1
    Loop
    
End Sub

Public Sub Delete_Pictures()
      
       Select Case ActiveSheet.Name
             Case Is = "__", "__", "__", "__", "__", "__"¤ªÑ", "¤j½L"
                    
                   If ActiveSheet.Pictures.Count > 0 Then ActiveSheet.Pictures.Delete
                   
                            
             Case Else
               
       End Select
       
 
End Sub

Sub Macro24()
'
' mikeon _ 2016/5/5 _____ªº¥¨¶°
' __:____´Á³ø¹S

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct

Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("w2"), SortOn _ _
        :=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("u2:am10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop

Call fm7
[dz100].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"

End Sub


Sub Macro25()
'
' mikeon _ 2016/5/5 _____ªº¥¨¶°
' __:__ªÑ¥«

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct

Application.Calculation = xlManual

    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("ab2"), _ _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("w2"), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("u2:am10000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop

Call fm7
[dz100].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
   
End Sub

Sub Macro41()
'
' mikeon _ 2016/5/5 _____ªº¥¨¶°
' __:__¥«­È

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct

Application.Calculation = xlManual

    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("ad2"), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("w2"), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("u2:am10000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop

Call fm7
[dz100].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
   
End Sub




Sub Macro26()
'
' mikeon _ 2016/5/16 _____ªº¥¨¶°
' __:__ªÑ¦W

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct

Application.Calculation = xlManual


[s1] = [t4]
Columns("t:t").ClearContents
a = 2
Do Until Cells(a, 22) = ""
Cells(a, 20) = Cells(a, 22)
a = a + 1
Loop

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("t2"), _ _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("t2:am10000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Columns("t:t").ClearContents

Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop

[t4] = [s1]: [s1] = ""
Call fm7
[dz100].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
   
End Sub


Sub Macro27()
'
' mikeon _ 2016/5/17 _____ªº¥¨¶°
' __:__§Ç¸¹

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)

Sheets("__").Selectct
Application.Calculation = xlManual
    
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("aa2"), _ _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("w2"), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("u2:am10000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
 Sheets("__").Selectct
 [t3] = [m5]

 c = 1 '_______h»¡©ú¦C
 Do Until Cells(c, 1) = ""
 c = c + 1
 Loop
 c = c - 1
 
 i = 1: j = 1 '_____á¤@¦C
Do Until i - j > 100
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
j = i
End If
i = i + 1
Loop
 
g = c + 3
h = i + 50
Range("A" & g & ":" & "R" & h).Copy
Sheets("__").Selectct
Range("CC1").Select
    ActiveSheet.Paste
    
Sheets("__").Selectct

u = 2
Do Until Cells(u, 27) = ""
u = u + 1
Loop
u = u - 2

Columns("A:R").Select
Selection.Clear
    
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop


k = 2
Do Until Cells(k, 27) = ""
Application.StatusBar = k - 1
i = 1: j = 1


Do Until i - j > 100
If Left(Trim(Sheets("__").Cells(i, 83)), 2) = "RO" Then '(__±`§Q
j = i

If Trim(Cells(k, 22)) = Trim(Sheets("__").Cells(i, 81)) Then '(____¦W¬Û¦P
p = c: q = c
Do Until p - q > 50
If IsError(Cells(p, 9)) Then
q = p
GoTo p50
End If
If Trim(Cells(p, 9)) <> "" Then q = p
p50: p = p + 1
Loop
m = q + 3

If Trim(Sheets("__").Cells(i + 1, 84)) = "_" Then '(_'(¥x
g = i - 2
h = i + 29
Sheets("__").Range("CC" & g & ":" & "CT" & h).Copypy
Cells(m, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(m, 1) = Cells(k, 27) & " / " & u & " _"""
Exit Do
End If '_))

If Trim(Sheets("__").Cells(i + 1, 84)) = "_" Then '(_'(¬ü
g = i - 2
h = i + 24
Sheets("__").Range("CC" & g & ":" & "CT" & h).Copypy
Cells(m, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(m, 1) = Cells(k, 27) & " / " & u & " _"""
Exit Do
End If '_))

If Trim(Sheets("__").Cells(i + 1, 84)) = "_" Then '(_'(´ä
g = i - 2
h = i + 28
Sheets("__").Range("CC" & g & ":" & "CT" & h).Copypy
Cells(m, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(m, 1) = Cells(k, 27) & " / " & u & " _"""
Exit Do
End If '_))

If Trim(Sheets("__").Cells(i + 1, 84)) = "_" Then '(_'(¤¤
g = i - 2
h = i + 23
Sheets("__").Range("CC" & g & ":" & "CT" & h).Copypy
Cells(m, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(m, 1) = Cells(k, 27) & " / " & u & " _"""
Exit Do
End If '_))

If Trim(Sheets("__").Cells(i + 1, 84)) <> "_" And Trim(Sheets("__").Cells(i + 1, 84)) <> "_" And Trim(Sheets("__").Cells(i + 1, 84)) <> "_" And Trim(Sheets("__").Cells(i + 1, 84)) <> "_" Then '(_¤¤" Then '(¥þ
g = i - 2
h = i + 24
Sheets("__").Range("CC" & g & ":" & "CT" & h).Copypy
Cells(m, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Cells(m, 1) = Cells(k, 27) & " / " & u & " _"""
Exit Do
End If '_))

End If '____)Û¦P)

End If '__)Q)

i = i + 1
Loop

k = k + 1
Loop
    
Sheets("__").Selectct
Columns("CB:CV").Select
  Selection.Delete Shift:=xlToLeft
[be1].Select
Sheets("__").Selectct
[m5] = [t3]
[t3] = ""
   
Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop

Call fm7
Sheets("__").Selectct
Call ProtectSheet(ActiveSheet)
Sheets("__").Selectct
[dz100].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
Beep

End Sub

Sub Macro28()
'
' mikeon _ 2016/5/29 _____ªº¥¨¶°
' __:__¤ÀÃþ

'
Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Sheets("__").Selectct

Application.Calculation = xlManual
    
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("ac2"), _ _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("w2"), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("u2:am10000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call addlink
a = 2
Do Until Cells(a, 22) = ""
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "CN" Then Cells(a, 22) = Format(Cells(a, 22), "000000"))
If Trim(Cells(a, 28)) = "_" Or UCase(Trim(Cells(a, 28))) = "HK" Then Cells(a, 22) = Format(Cells(a, 22), "0000"))
a = a + 1
Loop

Call fm7
[dz100].Select

Application.Calculation = xlAutomatic           '_____}¼sºÖ
Application.StatusBar = "__"¨"
   
End Sub


Sub Macro29()
'
' mikeon _ 2019/10/4 _____ªº¥¨¶°
'______¶iÂÂÀÉ
'

Sheets("__").Selectct
Application.ScreenUpdating = True
[g3] = "___...".."
Application.Wait Now() + TimeValue("00:00:01")

Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\¨q¶²

Application.Calculation = xlManual
Application.ScreenUpdating = False
Call ak2k

Dim FileToOpen As Variant
Dim OpenBook As Workbook
FileToOpen = Application.GetOpenFilename(Title:="______")¦Aªí")
If FileToOpen <> False Then
Set OpenBook = Application.Workbooks.Open(FileToOpen)
OpenBook.Sheets("__").Activatete
Call ak2k

'___ÀÃþ
b = 2 '_____Ã´X¤ä
Do Until Cells(b, 22) = ""
b = b + 1
Loop
b = b - 2

i = 1: j = 1: n = 0
Do Until i - j > 100 Or n > b
If Left(Trim(Cells(i, 3)), 2) = "RO" Then
n = n + 1: j = i

k = 2
Do Until Cells(k, 22) = ""
If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And Trim(Cells(k, 28)) = Trim(Cells(i + 1, 4)) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
Cells(i - 2, 2) = Cells(k, 30) '__­È
Exit Do
End If

If Trim(Cells(k, 22)) = Trim(Cells(i, 1)) And UCase(Trim(Cells(k, 28))) = UCase(Trim(Cells(i - 1, 1))) Then
Cells(i - 2, 4) = Cells(k, 29) '__Ãþ
Cells(i - 2, 2) = Cells(k, 30) '__­È
Exit Do
End If
k = k + 1
Loop

End If
i = i + 1
Loop



k = 1 '_______h»¡©ú¦C
Do Until Cells(k, 1) = ""
k = k + 1
Loop
k = k - 1

i = k: j = k
Do Until i - j > 50
If IsError(Cells(i, 9)) Then
j = i
GoTo i50
End If
If Trim(Cells(i, 9)) <> "" Then j = i
i50: i = i + 1
Loop
i = j + 3

RR = i
Set a = Range(Cells(k + 3, 1), Cells(i, 18))
a.Copy

ThisWorkbook.Worksheets("__").Activatete

k = 1 '_______h»¡©ú¦C
Do Until Cells(k, 1) = ""
k = k + 1
Loop
k = k - 1

i = k: j = k
Do Until i - j > 50
If IsError(Cells(i, 9)) Then
j = i
GoTo ii50
End If
If Trim(Cells(i, 9)) <> "" Then j = i
ii50: i = i + 1
Loop
i = j + 3

Cells(i, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False



OpenBook.Sheets("__").[m5].Copypy
ThisWorkbook.Worksheets("__").[m5].Select: ActiveSheet.Pastete

OpenBook.Close False
End If

[t4] = ""
Macro21

i = 1
Do Until i > RR
If Cells(i, 4) = "_" Thenn
For i = i + 17 To i + 17 + 9
Cells(i, 1).NumberFormatLocal = "yyyy/m"
Next i
GoTo i17
End If

If Cells(i, 4) = "_" Thenn
For i = i + 16 To i + 16 + 4
Cells(i, 1).NumberFormatLocal = "yyyy/m"
Next i
GoTo i17
End If

i17: i = i + 1
Loop

[g3] = ""

Application.ScreenUpdating = True

End Sub


Public Sub ak2k()

Application.Calculation = xlManual

i = 2: Do Until Cells(i, 22) = ""
k = 1: j = 1
Do Until k - j > 100 '_____________________ù¡B¥«­ÈªÑ»ù¶ñ¦^¦¬ÂÃªÑ
If Left(Trim(Cells(k, 3)), 2) = "RO" Then
j = k
If Trim(Cells(k + 1, 4)) = "_" Then Cells(k, 1) = Format(Cells(k, 1), "000000"))
If Trim(Cells(k + 1, 4)) = "_" Then Cells(k, 1) = Format(Cells(k, 1), "0000"))

If Trim(Cells(i, 22)) = Trim(Cells(k, 1)) And Cells(i, 28) = Trim(Cells(k + 1, 4)) Then
Cells(k - 1, 9) = Cells(i, 36)  '__´Á
If Len(Cells(k - 1, 9)) < 11 Then Cells(k - 1, 9).ShrinkToFit = True
Cells(k + 2, 11) = Cells(i, 37) '____³ø¹S
Cells(k + 1, 11) = Cells(i, 38) '__»ù
Cells(k - 2, 9) = Cells(i, 39) '____ªÑ»ù
Exit Do
End If
End If

k = k + 1
Loop
i = i + 1
Loop

End Sub

Sub Macro40()
'
' mikeon _ 2015/10/28 _____ªº¥¨¶°
'

'
    Sheets("__").Selectct
   ActiveWindow.ScrollColumn = 13: ActiveWindow.ScrollRow = 1
End Sub




