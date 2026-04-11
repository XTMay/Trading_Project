Attribute VB_Name = "Module4"
Sub Macro4()

' mikeon_ 2007/5/7 _____的巨集



Sheets("__").Selectct
Application.Calculation = xlAutomatic '_____}廣福
[b11] = "__2_______..."h擔待..."
Application.Calculation = xlManual '____徐廣
Call UnprotectSheet(ActiveSheet)
With Application '_____}廣福
         .MaxChange = 0.001    '____速度
         .EnableCancelKey = xlInterrupt '_____\秀雯
         .ScreenUpdating = False
 End With
 
 On Error GoTo err
 With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
 End With

   [f15] = 1
   [a54] = "______________(Michael On)__"Michael On)所有"
   [i12] = Left(ActiveWorkbook.Name, InStr(1, ActiveWorkbook.Name, ".") - 1)
 
    Range("o7:p37").ClearContents
    [i14] = "": [q6] = "": [a1] = "": pnd = 1: [i1] = "": [y23] = 0: [y24] = 0
    dc1 = 31 '__ 5 ____(______)眴p，非單季)
    dc11 = 60 '__要
    dc12 = 65 '____股價
    dc13 = 75 '__+__股子
    dc14 = 86 '____簡介
    [a12] = "Yahoo___________________ Y"在中國請於右格打 Y"
    
    t = 7: b = 37
        
    Range("E9").FormulaR1C1 = "=IFERROR(R[39]C[14],""_"")"""
    Range("k3").FormulaR1C1 = "=R[3]C[6]"
    [k10] = 12
    Range("k11").FormulaR1C1 = "=R[2]C[13]"
    Range("O2").FormulaR1C1 = "=VLOOKUP(R[5]C[-1]-ROUNDUP(RC[2],0)+1,R[5]C[-1]:R[35]C[2],2,FALSE)"
    Range("P2").FormulaR1C1 = "=VLOOKUP(R[5]C[-2]-ROUNDUP(RC[1],0)+1,R[5]C[-2]:R[35]C[1],4,FALSE)"
    Range("q2") = "=YEAR(TODAY())-IF(R[13]C[-2]<>"""",R[13]C[-3],IF(R[12]C[-2]<>"""",R[12]C[-3],IF(R[11]C[-2]<>"""",R[11]C[-3],IF(R[10]C[-2]<>"""",R[10]C[-3],IF(R[9]C[-2]<>"""",R[9]C[-3],IF(R[8]C[-2]<>"""",R[8]C[-3],IF(R[7]C[-2]<>"""",R[7]C[-3],IF(R[6]C[-2]<>"""",R[6]C[-3],R[5]C[-3]-MONTH(TODAY())/12))))))))"

Range("A2").NumberFormatLocal = "@"
cd$ = [a2]
dc = dc1
Range(Columns(dc), Columns(dc + 100)).Clear
Application.StatusBar = "___ 5 ____(______)  1 / 14"璈u)  1 / 14"
    Cells(1, dc) = "1 / 14 __ 5 ____(______) https://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/" & cd$ & "/ctrl/part/displaytype/4.phtml"ype/4.phtml"
        
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/" & cd$ & "/ctrl/part/displaytype/4.phtml" _
        , Destination:=Cells(2, dc))
        .Name = "4.phtml"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingAll
        .WebTables = """ProfitStatementNewTable0"""
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
    Delete_Pictures '__Alex_x吳
    
         [m1] = ""
         If Len(Sheets("__").[t3]) > 5 Thenen
         [m1] = Sheets("__").[t3]3]
         If DateDiff("d", [m1], [af4]) < 40 Then
         [m1] = "N"
         [b11] = "_______"未進來"
         GoTo __式
         End If
         End If
         
         For m = 2 To 20
         If Cells(m, dc1) <> "" Then Exit For
         Next m
         If m > 19 Then GoTo __式
    
'_~
'For i = 1 To 4
'Cells(24 - i, 1) = Cells(4, dc + i)
'Next i

Application.StatusBar = "___ 5 ____(______)  2 / 14"璈u)  2 / 14"
Cells(57, dc) = "2 / 14 __ 5 ____(______) https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/" & cd$ & "/ctrl/part/displaytype/4.phtml"ype/4.phtml"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/" & cd$ & "/ctrl/part/displaytype/4.phtml" _
        , Destination:=Cells(58, dc))
        .Name = "4.phtml"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingAll
        .WebTables = """BalanceSheetNewTable0"""
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
    Delete_Pictures '__Alex_x吳
    
    
For j = 2 To 50
If Cells(j, dc1 + 1) <> "" And IsNumeric(Year(Cells(j, dc1 + 1))) Then Exit For
Next j
[a9] = Year(Cells(j, dc1 + 1))
If Month(Cells(j, dc1 + 1)) = 12 Then [a9] = Year(Cells(j, dc1 + 1)) + 1

ii = 2
For i = 1 To 4
    If Not IsDate(Cells(32 - i, 1)) Then Exit For
    yr$ = Year(Cells(j, dc1 + 1)) - i - 1
    If Month(Cells(j, dc + 1)) = 12 Then yr$ = Year(Cells(j, dc1 + 1)) - i
    ii = ii + 1
     Application.StatusBar = "_" & yr$ & "____(______)  " & ii & " / 14"i & " / 14"
     Cells(1, dc + 1 + 5 * i) = ii & " / 14_" & yr$ & " ____(______) https://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/" & cd$ & "/ctrl/" & yr$ & "/displaytype/4.phtml"pe/4.phtml"
                        
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/" & cd$ & "/ctrl/" & yr$ & "/displaytype/4.phtml" _
        , Destination:=Cells(2, dc + 1 + 5 * i))
        .Name = "4.phtml"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingAll
        .WebTables = """ProfitStatementNewTable0"""
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
    Delete_Pictures '__Alex_x吳
    ii = ii + 1
    Application.StatusBar = "_" & yr$ & "____(______)  " & ii & " / 14"i & " / 14"
    Cells(57, dc + 1 + 5 * i) = ii & " / 14_" & yr$ & " ____(______) https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/" & cd$ & "/ctrl/" & yr$ & "/displaytype/4.phtml"pe/4.phtml"
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/" & cd$ & "/ctrl/" & yr$ & "/displaytype/4.phtml" _
         , Destination:=Cells(58, dc + 1 + 5 * i))
        .Name = "4.phtml"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingAll
        .WebTables = """BalanceSheetNewTable0"""
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
    Delete_Pictures '__Alex_x吳
    
Next i

'_____磞洶J
k = 1
Do Until Mid(Trim(Cells(k, dc)), 1, 1) = "_" Or k > 588
k = k + 1
Loop
If k < 58 Then
Cells(k, dc) = [aa14]
End If

'_____磞^轉
k = 1
Do Until Mid(Trim(Cells(k, dc)), 2, 1) = "_" Or k > 588
k = k + 1
Loop
If k < 58 Then
Cells(k, dc) = [aa15]
End If

'___b利
k = 1
Do Until Mid(Trim(Cells(k, dc)), 3, 2) = "__" Or k > 5858
k = k + 1
Loop
If k < 58 Then
Cells(k, dc) = [aa16]
End If

'___B分
k = 1
Do Until Mid(Trim(Cells(k, dc)), 1, 2) = "__" Or k > 5858
k = k + 1
Loop
If k < 58 Then
Cells(k, dc) = [aa17]
End If

'_____ぎb利
k = 1
Do Until Mid(Trim(Cells(k, dc)), 1, 1) = "_" Or k > 588
k = k + 1
Loop
If k < 58 Then
Cells(k, dc) = [aa21]
End If

 '______、淨值
k = 66
Do Until Trim(Cells(k, dc)) = "" Or k > 200
k = k + 1
Loop
If k < 200 Then
k = k - 1
Cells(k - 1, dc) = [aa30] '__值
Cells(k, dc) = [aa31] '__產
End If

'___悒
k = 62
Do Until Mid(Trim(Cells(k, dc)), 1, 1) = "_" Or k > 2000
k = k + 1
Loop
If k < 200 Then
k = k + 1
Cells(k, dc) = [aa32]
End If

For i = 1 To 4

'_____磞洶J
k = 1
Do Until Mid(Trim(Cells(k, dc + 1 + 5 * i)), 1, 1) = "_" Or k > 588
k = k + 1
Loop
If k < 58 Then Cells(k, dc + 1 + 5 * i) = [aa14]

'_____磞^轉
k = 1
Do Until Mid(Trim(Cells(k, dc + 1 + 5 * i)), 2, 1) = "_" Or k > 588
k = k + 1
Loop
If k < 58 Then Cells(k, dc + 1 + 5 * i) = [aa15]

'___b利
k = 1
Do Until Mid(Trim(Cells(k, dc + 1 + 5 * i)), 3, 2) = "__" Or k > 5858
k = k + 1
Loop
If k < 58 Then Cells(k, dc + 1 + 5 * i) = [aa16]

'___B分
k = 1
Do Until Mid(Trim(Cells(k, dc + 1 + 5 * i)), 1, 2) = "__" Or k > 5858
k = k + 1
Loop
If k < 58 Then Cells(k, dc + 1 + 5 * i) = [aa17]

'_____ぎb利
k = 1
Do Until Mid(Trim(Cells(k, dc + 1 + 5 * i)), 1, 1) = "_" Or k > 588
k = k + 1
Loop
If k < 58 Then Cells(k, dc + 1 + 5 * i) = [aa21]

'______、淨值
k = 66
Do Until Trim(Cells(k, dc + 1 + 5 * i)) = "" Or k > 200
k = k + 1
Loop
If k < 200 Then
k = k - 1
Cells(k - 1, dc + 1 + 5 * i) = [aa30] '__值
Cells(k, dc + 1 + 5 * i) = [aa31] '__產
End If

'___悒
k = 62
Do Until Mid(Trim(Cells(k, dc + 1 + 5 * i)), 1, 1) = "_" Or k > 2000
k = k + 1
Loop
If k < 200 Then
k = k + 1
Cells(k, dc + 1 + 5 * i) = [aa32]
End If

Next i


'__________Alex______-------------------------------------------------------------

ex$ = "SHA"
If Left(cd$, 1) = "0" Then ex$ = "SHE"
If Left(cd$, 1) = "2" Then ex$ = "SHE"
If Left(cd$, 1) = "3" Then ex$ = "SHE"
Application.StatusBar = "___  11 / 14"14"

dc = dc11
Cells(1, dc) = "11 / 14 __ https://www.aastocks.com/tc/cnhk/analysis/company-fundamental/basic-information?shsymbol=" & cd$d$
Call ConnectMarketWatch("https://www.aastocks.com/tc/cnhk/analysis/company-fundamental/basic-information?shsymbol=" & cd$, Cells(2, dc), 5)
 
[a1] = Trim([ae2]) '__+__市場
[a1] = Left([a1], Len([a1]) - 4)

For k = k To 500
   If Left(Cells(k, dc), 4) = "____" Then Exit For For
Next k ' Date
[i1] = Left(Right(Cells(k, dc), Len(Cells(k, dc)) - 4), Len(Right(Cells(k, dc), Len(Cells(k, dc)) - 4)) - 6)
Cells(k, dc - 1) = 1
If Not (IsDate([i1])) Then [i1] = Date

For k = k To 500
 If Cells(k, dc) = "____" Then Exit For For
Next k
[a1] = [a1] + "_" + Cells(k - 1, dc))
Cells(k - 1, dc - 1) = 1

For k = k To 500
   If Cells(k, dc) = "___" Then Exit ForFor
Next k
[q6] = Cells(k, dc + 1) '__價
Cells(k, dc - 1) = 1

For k = k To 500
 If Cells(k, dc) = "___" Then Exit For '___總市值
Next k
Cells(k, dc - 1) = 1
If IsNumeric(Left(Cells(k, dc + 1), Len(Cells(k, dc + 1)) - 1)) Then [y23] = 100 * Left(Cells(k, dc + 1), Len(Cells(k, dc + 1)) - 1)
If Right(Cells(k, dc + 1), 1) = "_" Or UCase(Right(Cells(k, dc + 1), 1)) = "T" Then [y23] = 1000000 * Left(Cells(k, dc + 1), (Len(Cells(k, dc + 1)) - 2)))
If Right(Cells(k, dc + 1), 1) = "b" Then [y23] = 1000 * Left(Cells(k, dc + 1), Len(Cells(k, dc + 1)) - 1)


ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
Delete_Pictures '__Alex_x吳

Range(Columns(dc + 2), Columns(dc + 4)).Clear
ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
Delete_Pictures '__Alex_x吳

[g13].Select
Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.google.cn/finance/company_news?q=" & ex$ & ":" & cd$, TextToDisplay:="__g"g"


'__+__+____ __Alex_______lex吳桑和方家晟桑
Application.StatusBar = "_____  12 / 14"/ 14"

If UCase([f12]) = "Y" Then GoTo CN1

ex$ = "SS"
    If Left(cd$, 1) = "0" Then ex$ = "SZ"
    If Left(cd$, 1) = "2" Then ex$ = "SZ"
    If Left(cd$, 1) = "3" Then ex$ = "SZ"
    
yrs$ = [n7] - 30
yre$ = [n7]

dc = dc12 'BM _______ ' for pricer price
Range(Columns(dc), Columns(dc + 6)).Clear

'Call cnhis '__20_______yahoo____yahoo已不給抓

'___8_____歷史股價
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp

url = "https://finance.yahoo.com/quote/" & cd$ & "." & ex$ & "/history?period1=573436800&period2=" & myday & "&interval=1mo&filter=history&frequency=1mo"
Cells(1, dc) = "12 / 14 ____ " + url url
'Call ConnectMarketWatch(url, Cells(3, dc), 2)

Call CN_Yahoo_Price_Dividend_Split(cd$, ex$, dc)

CN1:


 '__+_____1___10__ x _每10股配 x 元
Application.StatusBar = "___+__  13 / 14"/ 14"
dc = dc13
   Cells(1, dc) = "13 / 14 __+__ https://money.finance.sina.com.cn/corp/go.php/vISSUE_ShareBonus/stockid/" & cd$ & ".phtml"tml"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://money.finance.sina.com.cn/corp/go.php/vISSUE_ShareBonus/stockid/" & cd$ & ".phtml" _
        , Destination:=Cells(2, dc))
        .Name = cd$ & ".phtml"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingAll
        .WebTables = """sharebonus_1"""
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
    Delete_Pictures '__Alex_x吳
    
i = 5
Do Until Cells(i, dc + 5) = ""
If Cells(i, dc + 5) <> "" And IsNumeric(Left(Cells(i, dc + 5), 1)) Then Cells(i, dc - 1) = Year(Cells(i, dc + 5))
If Cells(i, dc - 1) <> "" And Not IsNumeric(Cells(i, dc - 1)) Then Cells(i, dc - 1) = 0
i = i + 1
Loop

dc = dc14
Application.StatusBar = "_____  14 / 14"/ 14"
Cells(1, dc) = "14 / 14 ____ http://www.aastocks.com/tc/cnhk/analysis/company-fundamental/company-profile?shsymbol=" & cd$ cd$
Call ConnectMarketWatch("http://www.aastocks.com/tc/cnhk/analysis/company-fundamental/company-profile?shsymbol=" & cd$, Cells(2, dc), 5)

For k = 2 To 500
If Cells(k, dc) = "____" Then Exit For For
Next k
[a1] = [a1] + "*" + Cells(k, dc + 1)


err:
[b11] = ""
    
 '----------------------------------------------------------------------------------------------------

If UCase([f12]) = "Y" Then GoTo CN2

dc = dc12
For c = 2 To 500
If Cells(c, dc) = "Date" Or Cells(c, dc) = "__" Then Exit Foror
Next c
If c > 500 Then GoTo __式
c = c + 1

Call highlow(dc, c, t, b, pnd)

CN2:

If [q6] = "" Or Not (IsNumeric([q6])) Then
For i = 1 To 100
If Cells(i, dc + 4) <> "" And IsNumeric(Cells(i, dc + 4)) Then Exit For
Next i
[q6] = Cells(i, dc + 4) * pnd
End If

If [q6] < [o7] Then [o7] = [q6]
If [q6] > [p7] Then [p7] = [q6]
If [o7] = "" Then [o7] = [q6]
If [p7] = "" Then [p7] = [q6]


'--------------------------------------------------------------------------------------------------

__: [y24] = [q6]6]
    Range("o2:q2").Select
    Selection.NumberFormatLocal = "#,##0.0_);(#,##0.0)"
    
    [h13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/echarts?s=" & cd$ & "." & ex$ & "#symbol=" & cd$ & "." & ex$ & ";range=5y", TextToDisplay:="___"圖"
    
    ex$ = "sh"
      If Left(cd$, 1) = "0" Then ex$ = "sz"
      If Left(cd$, 1) = "2" Then ex$ = "sz"
    [g13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://gu.qq.com/" + ex$ + cd$ + "/gp/news", TextToDisplay:="__"D"
    
    [f13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="http://www.aastocks.com/tc/cnhk/analysis/company-fundamental/company-profile?shsymbol=" & cd$, TextToDisplay:="__""
    
    [e13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.google.com/finance/stockscreener", _
    TextToDisplay:="___"器"
    
   [d13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.tradingeconomics.com/china/gdp-growth-annual", _
    TextToDisplay:="GDP"

If [b11] = "_______" Then GoTo y16oTo y16

For m = 2 To 20
If Cells(m, dc13) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "13 / 14 _____+__"+股子"

If [q6] = "" Then [b11] = "_____"股價"

For m = 2 To 20
If Cells(m, dc1) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "1 / 14 _____________"報，別再問了"

For i = t To b
Cells(i, 21) = ""
Next i

y16:
With ActiveSheet.Cells
        .Font.Name = "____"體"
        .Font.Name = "Arial"
        .Font.FontStyle = "__""
        .Font.Size = 10
        .RowHeight = 16
        .ColumnWidth = 7.5
    End With
    Range("a16:a17").Font.Size = 9
    
    Range(Cells(2, 2), Cells(52, 10)).ShrinkToFit = True '___Y小
    [f9].ShrinkToFit = False '____縮小
    [i12].ShrinkToFit = False '____縮小
    [e15].ShrinkToFit = False '____縮小
    [g42].ShrinkToFit = False '____縮小
    [g43].ShrinkToFit = False '____縮小
    Range(Cells(18, 11), Cells(23, 11)).ShrinkToFit = False '____縮小
    Range(Cells(36, 10), Cells(43, 10)).ShrinkToFit = False '____縮小
    
    Range(Columns(dc1), Columns(dc1 + 100)).Select '___Y小
        With Selection
        .WrapText = False
        .ShrinkToFit = True
    End With
    
    Range(Cells(1, dc1), Cells(1, dc1 + 100)).ShrinkToFit = False '____縮小
    Range(Cells(57, dc1), Cells(57, dc1 + 100)).ShrinkToFit = False '____縮小
    
    [j1].ColumnWidth = 8
    Cells(1, dc1).ColumnWidth = 20
    Cells(1, dc1 + 6).ColumnWidth = 20
    Cells(1, dc1 + 5 * 2 + 1).ColumnWidth = 20
    Cells(1, dc1 + 5 * 3 + 1).ColumnWidth = 20
    Cells(1, dc1 + 5 * 4 + 1).ColumnWidth = 20

    Columns(dc12).NumberFormatLocal = "yyyy/m/d;@"
    Columns(dc13).NumberFormatLocal = "yyyy/m/d;@"
    
Range("c3:e9, b20:b23, e20:j23, b28:b33, e28:j33,j34:j34, b38:f41, h38:i41, b46:f52, h46:i52, r6:r11").Select
    Selection.NumberFormatLocal = "#,##0_);(#,##0)"
    
Range("k1,d13:i13").Select
    With Selection.Font
        .Size = 12
        .Name = "Arial"
    End With
    
Range("e15").HorizontalAlignment = xlRight
Range("f15").HorizontalAlignment = xlLeft

Range("k1,c13:i13").Select
    With Selection.Font
        .Size = 12
        .Name = "Arial"
    End With

Range("B11").ShrinkToFit = False
Application.StatusBar = "__""
ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
Delete_Pictures '__Alex_x吳

dc = dc14
For k = 2 To 500
        If Cells(k, dc) = "____" Then Exit For For
Next k
     Cells(k, dc + 1).Copy:
     [y16].Select: ActiveSheet.Paste
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
Application.Calculation = xlAutomatic           '_____}廣福

End Sub


Private Sub cnhis() '__20_______yahoo____yahoo已不給抓

Dim myurl(2)
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & ex$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=history&crumb="
myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & ex$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=div&crumb="
myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & ex$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=split&crumb="

Dim url As String, crumbrng As Range, crumbrng1 As Range
Dim serr As String, csvt As Variant, csv As Variant
Dim cii As Long, cjj As Long, ci As Long, cj As Long

On Error Resume Next
For iurl = 0 To 2
    url = myurl(iurl)
    If iurl = 0 Then 'price
        Cells(1, dc) = "____ " & url url
        endrow = 3
    Else
        endrow = Range("bm100000").End(xlUp).Row
    End If
    Set crumbrng = Cells(endrow, dc) ' for price + dividend+split
   
retryno = 0
rest:

    csvt = DCSV(url)
     If InStr(1, csvt(0), "<!doctype html public") >= 1 Or InStr(1, csvt(0), "Method Not Allowed") >= 1 Then
  
          If retryno > 3 Then
              GoTo err_yh
          
          Else
              Debug.Print "DCSV didnot get data, retryno: "; retryno, url
         
              Application.Wait Now() + TimeValue("00:00:01") * 8
              retryno = retryno + 1
              GoTo rest:
         End If
    End If
    
    cii = UBound(csvt)
    cjj = UBound(split(csvt(0), ",")) + 1
    If Not err.Number = 0 Then
        If serr = "err" Then GoTo err_yh
        serr = "err"
        If Not CCRUMBT(crumbrng) = 0 Then GoTo err_yh
        GoTo rest
    End If
   
    
    ReDim csvarray(1 To cii, 1 To cjj)
    For ci = 1 To cii
        csv = split(csvt(ci - 1), ",")
        For cj = 1 To cjj
            If iurl = 2 And cj = cjj Then
                csvarray(ci, cj) = "'" + csv(cj - 1)
            Else
               csvarray(ci, cj) = csv(cj - 1)
            End If
        Next cj
    Next ci
    crumbrng.Offset(1, 0).Resize(cii, cjj) = csvarray
    
err_yh:
Next iurl


'----clear and sort data-----------------------------------------------------------------------
    
    endrow = Range("BL10000").End(xlUp).Row
    For i = endrow To 7 Step -1
        If Range("BL" & i) = "Date" Or Len(Range("BL" & i)) = 0 Then
             Range("BL" & i & ":BS" & i).Delete Shift:=xlUp
      
        End If
    Next i
    
    endrow = Range("bm10000").End(xlUp).Row
    
    Range("bm4:BS" & endrow).Select
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("bm5:bm" & endrow), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("bm4:BS" & endrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
   
'----end clear and sort data--------------------------------------------------------------------------


End Sub



Sub CN_Yahoo_Price_Dividend_Split(cd$, ex$, dc)
           'Call UnprotectSheet(ActiveSheet)
  
  
            'cd$ = "600519"
            'ex$ = "SS"
            'dc = 65
            
            ActiveSheet.Range(Cells(3, dc - 1), Cells(10000, dc + 6)).ClearContents
            Dim myurl(2)
            myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
            
            myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & ex$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=history&includeAdjustedClose=true"
            myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & ex$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=div&includeAdjustedClose=true"
            myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & ex$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=split&includeAdjustedClose=true"
            
            
            
            Dim url As String, crumbrng As Range, crumbrng1 As Range
            Dim serr As String, csvt As Variant, csv As Variant
            Dim cii As Long, cjj As Long, ci As Long, cj As Long
        
            
            For iurl = 0 To 2
                url = myurl(iurl)
                If iurl = 0 Then 'price
                    
                    endrow = 3
                Else
                    endrow = Cells(100000, dc).End(xlUp).Row
                End If
                Set crumbrng = Cells(endrow, dc)  ' for price + dividend+split
                retryno = 0
rest:
                  
               
                 Set httpreq = CreateObject("MSXML2.XMLHTTP.3.0")
              
                 httpreq.Open "GET", myurl(iurl), False
             
                 httpreq.send
                  
                 csvt = split(httpreq.responseText, Chr(10))
                 httpreq.abort
                 Set httpreq = Nothing
                 
                 
                 If UBound(csvt) = 0 Then
                     GoTo err_yh
                 
                 ElseIf InStr(1, csvt(0), "<!doctype html public") >= 1 Then
              
                      If retryno > 3 Then
                          GoTo err_yh
                      
                      Else
                          Debug.Print "DCSV didnot get data, retryno ", retryno
                     
                          Application.Wait Now() + TimeValue("00:00:01") * 5
                          retryno = retryno + 1
                          GoTo rest:
                     End If
                End If
                
                cii = UBound(csvt)
                cjj = UBound(split(csvt(0), ",")) + 1
               
                
               If iurl = 1 Or iurl = 2 Then
                   ReDim csvarray(1 To cii, 1 To cjj + 1)
               Else
                  ReDim csvarray(1 To cii, 1 To cjj)
               End If
               For ci = 1 To cii
                    csv = split(csvt(ci - 1), ",")
                    For cj = 1 To cjj
                        If iurl = 2 And cj = cjj Then
                            csvarray(ci, cj) = "'" + csv(cj - 1)
                        Else
                           csvarray(ci, cj) = csv(cj - 1)
                        End If
                    Next cj
                    
                    If iurl = 1 Then csvarray(ci, cj) = "Dividend"
                    If iurl = 2 Then csvarray(ci, cj) = "Split"
                    
                Next ci
                If iurl = 0 Then
                   crumbrng.Offset(1, 0).Resize(cii, cjj) = csvarray
                Else
                   crumbrng.Offset(1, 0).Resize(cii, cjj + 1) = csvarray
                End If
                
err_yh:
            Next iurl
            
          '----clear and sort data-----------------------------------------------------------------------
            
                endrow = Cells(10000, dc).End(xlUp).Row
                For i = endrow To 7 Step -1
                    If Cells(i, dc) = "Date" Or Len(Cells(i, dc)) = 0 Then
                         'Range("CE" & i & ":CK" & i).Delete Shift:=xlUp
                         Range(Cells(i, dc), Cells(i, dc + 6)).Delete Shift:=xlUp
                  
                    End If
                Next i
                
            
            
                endrow = Cells(10000, dc).End(xlUp).Row
                
                'Range("CE4:CK" & endrow).Select
                Range(Cells(4, dc), Cells(endrow, dc + 6)).Select
                
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range(Cells(5, dc), Cells(endrow, dc)), _
                    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                With ActiveWorkbook.ActiveSheet.Sort
                    .SetRange Range(Cells(4, dc), Cells(endrow, dc + 6))
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
               
            '----end clear and sort data--------------------------------------------------------------------------
            
            
End Sub


