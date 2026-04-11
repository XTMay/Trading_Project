Attribute VB_Name = "Module3"
Sub Macro3()
'
' mikeon_ 2007/10/8 _____ªº¥¨¶°
'

    Sheets("__").Selectct
    Application.Calculation = xlAutomatic '_____}¼sºÖ
    [b11] = "__55_____..."y«Ý..."
    Application.Calculation = xlManual '_____}¼sºÖ
    Call UnprotectSheet(ActiveSheet)
     
    With ActiveWorkbook '____³t«×
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With

With Application '_____}¼sºÖ
     .MaxChange = 0.001
     .EnableCancelKey = xlInterrupt '_____\¨q¶²
End With
    
    [d11] = ""
    On Error GoTo err
    Application.ScreenUpdating = False
    
    [a68] = "______________(Michael On)__"Michael On)©Ò¦³"
    [i12] = Left(ActiveWorkbook.Name, InStr(1, ActiveWorkbook.Name, ".") - 1)
    
[i14] = "": [q6] = "": [f15] = 1: pnd = 1
Range("o7:p37").ClearContents
dc1 = 31 '_____l¯qªí
dc2 = 31 + 8 '_____ê²£ªí
dc3 = 31 + 16 '____¯qªí
dc4 = 31 + 24 '____²£ªí
dc5 = 64 '__®§
dc6 = 78 '__+____½æ³æ¦ì
dc7 = 83 '____ªÑ»ù
dc8 = 92 '__ªp
dc9 = 96 '__²v
t = 7: b = 37

 [a12] = "Yahoo___________________ Y"¦b¤¤°ê½Ð©ó¥k®æ¥´ Y"


    Range("E9").FormulaR1C1 = "=IFERROR(R[38]C[14],""_"")"""
    Range("k3").FormulaR1C1 = "=R[3]C[6]"
    [k10] = 12
    Range("K11").FormulaR1C1 = "=R[2]C[13]"
    Range("O2").FormulaR1C1 = "=VLOOKUP(R[5]C[-1]-ROUNDUP(RC[2],0)+1,R[5]C[-1]:R[35]C[2],2,FALSE)"
    Range("P2").FormulaR1C1 = "=VLOOKUP(R[5]C[-2]-ROUNDUP(RC[1],0)+1,R[5]C[-2]:R[35]C[1],4,FALSE)"
    Range("q2") = "=YEAR(TODAY())-IF(R[13]C[-2]<>"""",R[13]C[-3],IF(R[12]C[-2]<>"""",R[12]C[-3],IF(R[11]C[-2]<>"""",R[11]C[-3],IF(R[10]C[-2]<>"""",R[10]C[-3],IF(R[9]C[-2]<>"""",R[9]C[-3],IF(R[8]C[-2]<>"""",R[8]C[-3],IF(R[7]C[-2]<>"""",R[7]C[-3],IF(R[6]C[-2]<>"""",R[6]C[-3],R[5]C[-3]-MONTH(TODAY())/12))))))))"
    
Range("A2").NumberFormatLocal = "@"
cd$ = Format([a2], "0000")
'_____l¯qªí
Application.StatusBar = "______  1 / 9"1 / 9"
dc = dc1
Range(Columns(dc), Columns(dc + 100)).Clear

Cells(1, dc) = "1 / 9 __/______ https://www.aastocks.com/tc/stocks/analysis/company-fundamental/profit-loss?symbol=" & [a2] & "&period=2"eriod=2"
Dim url As String
  url = "https://www.aastocks.com/tc/stocks/analysis/company-fundamental/profit-loss?symbol=" & [a2] & "&period=2"
  Call ConnectMarketWatch(url, Cells(2, dc), 5)
  
         [m1] = ""
         Dim aj4 As Date
         If Len(Sheets("__").[t3]) > 5 Thenen
         [m1] = Sheets("__").[t3]3]
         aj4 = [aj4]
         If aj4 = "" Then aj4 = [ai4]
         If aj4 = "" Then aj4 = [ah4]
         If aj4 = "" Then aj4 = [ag4]
         If aj4 = "" Then aj4 = [af4]
         If DateDiff("d", [m1], aj4) < 40 Then
         [m1] = "N"
         [b11] = "_______"¥¼¶i¨Ó"
         GoTo __¦¡
         End If
         End If
         
         For m = 2 To 20
         If Cells(m, dc1) <> "" Then Exit For
         Next m
         If m > 19 Then GoTo __¦¡

k = 200
Do Until Cells(k, dc) = "____" Or k > 500 500
k = k + 1
Loop
i = k

Do Until Left(Cells(k, dc), 5) = "_____" Or k > 500 > 500
k = k + 1
Loop
j = k

Set a = Range(Cells(i, dc), Cells(j, dc + 5))
a.Copy: Cells(4, dc).Select: ActiveSheet.Paste

Set a = Range(Cells(j - i + 5, dc), Cells(j - i + 500, dc + 5))
a.ClearContents

For i = 5 To 50
If Left(Cells(i, dc), 2) = "__" Then Exit Foror
For j = dc + 1 To dc + 5
If Trim(Cells(i, j)) = "-" Then Cells(i, j) = 0
Next j
Next i

'_____ê²£ªí
dc = dc2
Application.StatusBar = "______  2 / 9"2 / 9"
Cells(1, dc) = "2 / 9 _____ https://www.aastocks.com/tc/stocks/analysis/company-fundamental/balance-sheet?symbol=" & [a2] & "&period=2"od=2"
       url = "https://www.aastocks.com/tc/stocks/analysis/company-fundamental/balance-sheet?symbol=" & [a2] & "&period=2"
       Call ConnectMarketWatch(url, Cells(2, dc), 5)

k = 200
Do Until Cells(k, dc) = "____" Or k > 500 500
k = k + 1
Loop
i = k

Do Until Left(Cells(k, dc), 5) = "_____" Or k > 500 > 500
k = k + 1
Loop
j = k

Set a = Range(Cells(i, dc), Cells(j, dc + 5))
a.Copy: Cells(4, dc).Select: ActiveSheet.Paste


Set a = Range(Cells(j - i + 5, dc), Cells(j - i + 500, dc + 5))
a.ClearContents
 
For i = 3 To 90 '_________vlook__§Kvlook§ä¿ù
If Cells(i, dc) = "_______" Then Cells(i, dc) = "_______" '1¤Îµu´Á¸êª÷" '1
Next i

For i = 5 To 50
If Left(Cells(i, dc), 2) = "__" Then Exit Foror
For j = dc + 1 To dc + 5
If Trim(Cells(i, j)) = "-" Then Cells(i, j) = 0
Next j
Next i
    
  '____¯qªí
Application.StatusBar = "_____  3 / 9" / 9"
dc = dc3
       Cells(1, dc) = "3 / 9 ____ https://www.aastocks.com/tc/stocks/analysis/company-fundamental/profit-loss/?symbol=" & [a2][a2]
       url = "https://www.aastocks.com/tc/stocks/analysis/company-fundamental/profit-loss/?symbol=" & [a2]
       Call ConnectMarketWatch(url, Cells(2, dc), 5)
       
k = 200
Do Until Cells(k, dc) = "____" Or k > 500 500
k = k + 1
Loop
i = k

Do Until Left(Cells(k, dc), 5) = "_____" Or k > 500 > 500
k = k + 1
Loop
j = k

Set a = Range(Cells(i, dc), Cells(j, dc + 5))
a.Copy: Cells(4, dc).Select: ActiveSheet.Paste

Set a = Range(Cells(j - i + 5, dc), Cells(j - i + 500, dc + 5))
a.ClearContents

For i = 5 To 50
If Left(Cells(i, dc), 2) = "__" Then Exit Foror
For j = dc + 1 To dc + 5
If Trim(Cells(i, j)) = "-" Then Cells(i, j) = 0
Next j
Next i

If Cells(5, dc + 1) <> "" And Not IsNumeric(Cells(5, dc + 1)) Then GoTo __¦¡
If Cells(5, dc + 1) = 0 Then GoTo __¦¡

'____²£ªí
dc = dc4
Application.StatusBar = "_____  4 / 9" / 9"
Cells(1, dc) = "4 / 9 ____ https://www.aastocks.com/tc/stocks/analysis/company-fundamental/balance-sheet?symbol=" & [a2][a2]
       url = "https://www.aastocks.com/tc/stocks/analysis/company-fundamental/balance-sheet?symbol=" & [a2]
       Call ConnectMarketWatch(url, Cells(2, dc), 5)
    
k = 200
Do Until Cells(k, dc) = "____" Or k > 500 500
k = k + 1
Loop
i = k

Do Until Left(Cells(k, dc), 5) = "_____" Or k > 500 > 500
k = k + 1
Loop
j = k

Set a = Range(Cells(i, dc), Cells(j, dc + 5))
a.Copy: Cells(4, dc).Select: ActiveSheet.Paste


Set a = Range(Cells(j - i + 5, dc), Cells(j - i + 500, dc + 5))
a.ClearContents
 
For i = 3 To 90 '_________vlook__§Kvlook§ä¿ù
If Cells(i, dc) = "_______" Then Cells(i, dc) = "_______" '1¤Îµu´Á¸êª÷" '1
Next i

For i = 5 To 50
If Left(Cells(i, dc), 2) = "__" Then Exit Foror
For j = dc + 1 To dc + 5
If Trim(Cells(i, j)) = "-" Then Cells(i, j) = 0
Next j
Next i

Application.Calculation = xlAutomatic '_____}¼sºÖ
For i = 71 To 76
If Cells(i, 22) = "_" Then Exit Forr
Next i
If i > 75 Then i = 70

For j = 1 To 5
Cells(8 - j, 1) = Cells(i + j, 22)
Next j
Application.Calculation = xlManual '_____}¼sºÖ


'__®§
Application.StatusBar = "___  5 / 9" 9"
dc = dc5

Cells(1, dc) = "5 / 9 __ https://www.aastocks.com/tc/stocks/analysis/company-fundamental/dividend-history?symbol=" & [a2]2]
       url = "https://www.aastocks.com/tc/stocks/analysis/company-fundamental/dividend-history?symbol=" & [a2]
       Call ConnectMarketWatch(url, Cells(2, dc), 5)
    
k = 200
Do Until Cells(k, dc) = "____" Or k > 500 500
k = k + 1
Loop
i = k

Do Until Left(Cells(k, dc), 5) = "_____" Or k > 500 > 500
k = k + 1
Loop
j = k

Set a = Range(Cells(i, dc), Cells(j, dc + 8))
a.Copy: Cells(4, dc).Select: ActiveSheet.Paste

Set a = Range(Cells(j - i + 5, dc), Cells(j - i + 500, dc + 8))
a.ClearContents

     
'__+____½æ³æ¦ì
Application.StatusBar = "___+____  6 / 9" 6 / 9"
dc = dc6
url = "https://www.aastocks.com/tc/stocks/analysis/company-fundamental/basic-information?symbol=" & [a2]
Cells(1, dc) = "6 / 9 __+____ " + url + url
Call ConnectMarketWatch(url, Cells(2, dc), 2)

k = 1
Do Until Cells(k, dc) = "____" Or k > 500 500
k = k + 1
Loop
[a1] = Cells(k, dc + 1)
i = k

Do Until Cells(k, dc) = "___" Or k > 500500
k = k + 1
Loop
j = k

[q6] = Cells(k, dc + 1) '__»ù

Do Until Cells(k, dc) = "___" Or k > 500500
k = k + 1
Loop

Set a = Range(Cells(i, dc), Cells(j, dc + 2))
a.Copy: Cells(2, dc).Select: ActiveSheet.Paste

Set a = Range(Cells(j - i + 3, dc), Cells(j - i + 500, dc + 8))
a.ClearContents

k = 1
Do Until Cells(k, dc) = "____" Or k > 20 > 20
k = k + 1
Loop
[y24] = Cells(k, dc + 1)

Range(Columns(dc + 4), Columns(dc + 7)).Clear
ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
Delete_Pictures '__Alex_x§d



'____ __Alex_______d®á©M¤è®aÑÔ®á

Application.StatusBar = "_____  7 / 9" / 9"

If UCase([f12]) = "Y" Then GoTo CN1

dc = dc7
yrs$ = [n7] - 30
yre$ = [n7]

'Call hkhis '__20_______yahoo____yahoo¤w¤£µ¹§ì


'___8_____¾ú¥vªÑ»ù
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
url = "https://finance.yahoo.com/quote/" & cd$ & ".HK/history?period1=573436800&period2=" & myday & "&interval=1mo&filter=history&frequency=1mo"
Cells(1, dc) = "7 / 9 ____ " + url url
'Call ConnectMarketWatch(url, Cells(3, dc), 2)

Call HK_Yahoo_Price_Dividend_Split(cd$, dc)

CN1:


dc = dc8
Application.StatusBar = "___  8 / 9" 9"
Cells(1, dc) = "8 / 9 __ https://www.aastocks.com/tc/Stock/CompanyFundamental.aspx?CFType=1&symbol=" & [a2]2]
With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.aastocks.com/tc/Stock/CompanyFundamental.aspx?CFType=1&symbol=" & [a2] _
        , Destination:=Cells(2, dc + 18))
        .Name = "CompanyFundamental.aspx?CFType=1&symbol=" & [a2]
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
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingAll
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
Delete_Pictures '__Alex_x§d

k = 1
Do Until Cells(k, dc + 18) = "____" Or k > 500 500
k = k + 1
Loop
[a1] = [a1] + "_" + Cells(k, dc + 18 + 1))


k = 1: Do Until Left(Cells(k, dc + 18), 2) = "__" Or k > 70000
k = k + 1
Loop

For i = 1 To 15
Cells(1 + i, dc) = Cells(k - 3 + i, dc + 18)
Cells(1 + i, dc + 1) = Cells(k - 3 + i, dc + 19)
Next i
Cells(1 + i - 2, dc + 1) = Cells(k - 3 + 14, dc + 20)

Range(Columns(dc + 18), Columns(dc + 100)).Clear
ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
Delete_Pictures '__Alex_x§d

'___ÑªF
  Cells(18, dc) = "___ https://www.aastocks.com/tc/Stock/CompanyFundamental.aspx?CFType=2&symbol=" & [a2]a2]
  With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.aastocks.com/tc/Stock/CompanyFundamental.aspx?CFType=2&symbol=" & [a2] _
        , Destination:=Cells(2, dc + 18))
        .Name = "CompanyFundamental.aspx?CFType=2&symbol=" & [a2]
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
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingAll
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d
    
k = 1: Do Until Left(Cells(k, dc + 18), 2) = "__" Or k > 40000
k = k + 1
Loop
j = k
Do Until Left(Cells(k, dc + 18), 2) = "__" Or k > 40000
k = k + 1
Loop

For i = 1 To 4 + k - j

Cells(18 + i, dc) = Cells(j + i - 3, dc + 18)
Cells(18 + i, dc + 1) = Cells(j + i - 3, dc + 19)
Cells(18 + i, dc + 2) = Cells(j + i - 3, dc + 20)
Next i

Range(Columns(dc + 18), Columns(dc + 100)).Clear
ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
Delete_Pictures '__Alex_x§d

If Left([y24], 1) = "_" Then [y25] = "HKD"""
If Left([y24], 1) = "_" Then [y25] = "USD"""
If Left([y24], 1) = "_" Then [y25] = "JPY"""
If Left([y24], 1) = "_" Then [y25] = "CNY"""
If Left([y24], 1) = "_" Then [y25] = "EUR"""
If Left([y24], 1) = "_" Then [y25] = "GBP"""

dc = dc9
   If [y25] = "HKD" Then GoTo cf
   Application.StatusBar = "___  9 / 9" 9"
   
   GoTo HL '_____yahooyahoo
   url = "https://finance.yahoo.com/quote/HKD" + [y25] + "=x?ltr=1" '__Alex_x§d
   Cells(1, dc) = "9 / 9 __ " & urlrl
   Call ConnectMarketWatch(url, Cells(2, dc), 2)
   
   For i = 2 To 20
    If Cells(i, dc) = "Previous Close" And Cells(i, dc + 1) <> "" And IsNumeric(Cells(i, dc + 1)) Then
    [f15] = Cells(i, dc + 1)
    Exit For
    End If
    If Cells(i, dc) = "Open" And Cells(i, dc + 1) <> "" And IsNumeric(Cells(i, dc + 1)) Then [f15] = Cells(i, dc + 1)
  Next i
  

HL:
    'url = "https://transferwise.com/zh-hk/currency-converter/HKD-to-" & [y25] & "-rate" '_____yahooyahoo
    url = "https://wise.com/zh-hk/currency-converter/HKD-to-" & [y25] & "-rate" '_____yahooyahoo
    Cells(1, dc) = "9 / 9 __ " & urlrl
    
         Call ConnectWinHttp(url, 1)
  
'        Call ConnectXMLHTTP(url)
                            
        If InStr(1, doc.body.innerHTML, "HTTP ERROR") >= 1 Then
        
           [f15] = "FX not found"
        
         Else
              
                For Each kk In doc.getElementsByTagName("h3")
                    If kk.className = "cc__source-to-target" Then
                           
                        For Each ff In kk.getElementsByTagName("span")
                            
                            If InStr(1, ff.className, "text-success") >= 1 Then
                             
                                [f15] = ff.innerText
                                 Exit For
                             End If
                        Next ff
                        
                    End If
                Next kk
                

        End If


  
  
cf:
  
err:
[b11] = ""

'------------------------------------------------------------------------------------

dc = dc5
c = 1
For c = 1 To 20
If Cells(c, dc + 1) <> "" And IsNumeric(Left(Cells(c, dc + 1), 1)) Then Exit For
Next c

If c = 20 Then GoTo cx5

Cells(c - 2, dc + 9) = "__"§"
Cells(c - 2, dc + 10) = "__": Cells(c - 1, dc + 10) = [y24]4]
Cells(c - 2, dc + 11) = "__": Cells(c - 1, dc + 11) = "__"ä¤¸"
i = c
Do Until Cells(i, dc) = "" Or i > 1000
If Cells(i, dc) <> "" And IsNumeric(Left(Cells(i, dc), 1)) Then

Cells(i, dc - 1) = Year(Cells(i, dc)) 'year

If InStr(1, Cells(i, dc + 3), "__") = 0 Thenen
Cells(i, dc + 8) = 0
Cells(i, dc + 9) = "_"""
Cells(i, dc + 10) = 0
Cells(i, dc + 11) = "na"
GoTo idiv
End If

If InStr(1, Cells(i, dc + 3), "_") = 0 And InStr(1, Cells(i, dc + 3), "_") = 0 And InStr(1, Cells(i, dc + 3), "_") = 0 Thenhen
Cells(i, dc + 8) = 0
Else:
'a = Application.WorksheetFunction.Max(InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"))"))
g = 0
If InStr(1, Cells(i, dc + 3), "_") > 0 Then g = g + 11
If InStr(1, Cells(i, dc + 3), "_") > 0 Then g = g + 11
If InStr(1, Cells(i, dc + 3), "_") > 0 Then g = g + 11
If g = 1 Then a = Application.WorksheetFunction.Max(InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"))"))
If g = 2 Then a = Application.WorksheetFunction.Median(InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"))"))
If g = 3 Then a = Application.WorksheetFunction.Min(InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"), InStr(1, Cells(i, dc + 3), "_"))"))

Cells(i, dc + 8) = Mid(Cells(i, dc + 3), a + 1, 6)
If Not (IsNumeric(Cells(i, dc + 8))) Then Cells(i, dc + 8) = Left(Cells(i, dc + 8), Len(Cells(i, dc + 8)) - 1)
End If

If InStr(1, Cells(i, dc + 3), "_") = 0 And InStr(1, Cells(i, dc + 3), ":") = 0 Thenn
Cells(i, dc + 9) = "_"""
Else:
If InStr(1, Cells(i, dc + 3), "_") = 0 Thenn
Cells(i, dc + 9) = Mid(Cells(i, dc + 3), InStr(1, Cells(i, dc + 3), ":") + 2, 1)
Else:
Cells(i, dc + 9) = Mid(Cells(i, dc + 3), InStr(1, Cells(i, dc + 3), "_") + 1, 1))
End If
End If


Cells(i, dc + 10) = Mid(Cells(i, dc + 3), InStr(1, Cells(i, dc + 3), Left([y24], 1)) + Len([y24]), 6)


If InStr(1, Cells(i, dc + 3), "_") = 0 Thenn
Cells(i, dc + 11) = "na"
Else:
Cells(i, dc + 11) = Mid(Cells(i, dc + 3), InStr(1, Cells(i, dc + 3), "_") + 2, 6))
End If

End If
idiv:
i = i + 1
Loop


cx5: If Cells(4, dc7) = "" Then GoTo __¦¡

If UCase([f12]) = "Y" Then GoTo CN2

dc = dc7
For c = 2 To 500
If Cells(c, dc) = "Date" Or Cells(c, dc) = "__" Then Exit Foror
Next c
If c > 500 Then GoTo __¦¡
c = c + 1
'[i1] = Cells(c, dc)
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


__:¡:

  [h13].Select
Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/echarts?s=" & cd$ & ".HK#symbol=" & cd$ & ".HK;range=5y", TextToDisplay:="___"¹Ï"
    
    [g13].Select
Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://hk.finance.yahoo.com/q/h?s=" & cd$ & ".HK", TextToDisplay:="__"D"
    
    [f13].Select
Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.google.com.hk/finance/company_news?q=HKG:" & cd$, TextToDisplay:="__g"g"
    
    [e13].Select
Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.aastocks.com/tc/stocks/analysis/company-fundamental/company-profile?symbol=" & [a2], TextToDisplay:="__"Ð"
    
    [d13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://mikeon88.blogspot.tw/2014/02/blog-post_2047.html", _
    TextToDisplay:="___"¾¹"
    
    [c13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.tradingeconomics.com/hong-kong/gdp-growth-annual", _
    TextToDisplay:="GDP"
    
If [b11] = "_______" Then GoTo y16oTo y16

If UCase([f12]) <> "Y" Then

For m = 2 To 20
If Cells(m, dc7) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "7 / 9 _______"¥vªÑ»ù"

End If

If [q6] = "" Then [b11] = "_____"ªÑ»ù"

For m = 2 To 20
If Cells(m, dc4) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "4 / 9 _______"¸ê²£ªí"

For m = 2 To 20
If Cells(m, dc3) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "3 / 9 _______"·l¯qªí"

For m = 2 To 20
If Cells(m, dc2) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "2 / 9 ________"~¸ê²£ªí"

For m = 2 To 20
If Cells(m, dc1) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "1 / 9 ________"~·l¯qªí"
    
For i = t To b
Cells(i, 21) = ""
Next i

With ActiveSheet.Cells
        .Font.Name = "____"úÅé"
        .Font.Name = "Arial"
        .Font.FontStyle = "__"Ç"
        .Font.Size = 10
        .RowHeight = 16
        .ColumnWidth = 7.5
    End With
    Range("a16:a17").Font.Size = 9
    
    Range(Cells(2, 2), Cells(66, 10)).ShrinkToFit = True '___Y¤p
    [f9].ShrinkToFit = False '____ÁY¤p
    [i12].ShrinkToFit = False '____ÁY¤p
    [e15].ShrinkToFit = False '____ÁY¤p
    [g55].ShrinkToFit = False '____ÁY¤p
    [g56].ShrinkToFit = False '____ÁY¤p
    Range(Cells(18, 11), Cells(23, 11)).ShrinkToFit = False '____ÁY¤p
    Range(Cells(43, 10), Cells(56, 10)).ShrinkToFit = False '____ÁY¤p
    
  Range(Columns(dc1), Columns(dc1 + 100)).Select '___Y¤p
        With Selection
        .WrapText = False
        .ShrinkToFit = True
    End With
    
  Range(Cells(1, dc1), Cells(1, dc1 + 100)).ShrinkToFit = False '____ÁY¤p

    [j1].ColumnWidth = 8 '____Äæ¼e
    Cells(1, dc1).ColumnWidth = 20
    Cells(1, dc1 + 8).ColumnWidth = 20
    Cells(1, dc1 + 16).ColumnWidth = 20
    Cells(1, dc1 + 24).ColumnWidth = 20
    
Range("c3:e9, b20:b29, e20:j29, b34:b39, e34:j39, b45:f54, h45:i54, b59:f64, h59:i64, j41:j41, ae:bv").Select
    Selection.NumberFormatLocal = "#,##0_);(#,##0)"

Range(Cells(4, dc1 + 1), Cells(4, dc1 + 5)).NumberFormatLocal = "yyyy/m/d;@"
Range(Cells(4, dc2 + 1), Cells(4, dc2 + 5)).NumberFormatLocal = "yyyy/m/d;@"
Range(Cells(4, dc3 + 1), Cells(4, dc3 + 5)).NumberFormatLocal = "yyyy/m/d;@"
Range(Cells(4, dc4 + 1), Cells(4, dc4 + 5)).NumberFormatLocal = "yyyy/m/d;@"
Range(Columns(dc5), Columns(dc5 + 1)).NumberFormatLocal = "yyyy/m/d;@"
Range(Columns(dc5 + 5), Columns(dc5 + 7)).NumberFormatLocal = "yyyy/m/d;@"
Columns(dc7).NumberFormatLocal = "yyyy/m/d;@"
Range("a20:a28").NumberFormatLocal = "yyyy/m;@"
    
Range("e15").HorizontalAlignment = xlRight
Range("f15").HorizontalAlignment = xlLeft

Range("k1,c13:i13").Select
    With Selection.Font
        .Size = 12
        .Name = "Arial"
    End With
    
y16: Range("B11").ShrinkToFit = False
Application.StatusBar = "__"¨"
ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
Delete_Pictures '__Alex_x§d

For k = 2 To 400
         If Left(Trim(Cells(k, dc8)), 11) = "____" Then Exit For For
Next k
Cells(k, dc8 + 1).Copy: [y16].Select: ActiveSheet.Paste
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
Application.Calculation = xlAutomatic           '_____}¼sºÖ

End Sub


Private Sub hkhis() '__20_______yahoo____yahoo¤w¤£µ¹§ì

Dim myurl(2)
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & ".HK?period1=57600&period2=" & myday & "&interval=1mo&events=history&crumb="
myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & ".HK?period1=57600&period2=" & myday & "&interval=1mo&events=div&crumb="
myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & ".HK?period1=57600&period2=" & myday & "&interval=1mo&events=split&crumb="

Dim crumbrng As Range, crumbrng1 As Range
Dim serr As String, csvt As Variant, csv As Variant
Dim cii As Long, cjj As Long, ci As Long, cj As Long

On Error Resume Next
For iurl = 0 To 2
    url = myurl(iurl)
    If iurl = 0 Then 'price
        Cells(1, dc) = "____ " & url url
        endrow = 3
    Else
        endrow = Range("cw100000").End(xlUp).Row
    End If
    Set crumbrng = Cells(endrow, dc) ' for price + dividend+split
   
retryno = 0
rest:

    csvt = DCSV(url)
     If InStr(1, csvt(0), "<!doctype html public") >= 1 Or InStr(1, csvt(0), "Method Not Allowed") >= 1 Then
  
          If retryno > 3 Then
              GoTo err_yh
          
          Else
              Debug.Print "DCSV didnot get data, retryno:"; retryno, url
         
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
    
    endrow = Range("cw10000").End(xlUp).Row
    For i = endrow To 7 Step -1
        If Range("cw" & i) = "Date" Or Len(Range("cw" & i)) = 0 Then
             Range("CW" & i & ":DC" & i).Delete Shift:=xlUp
      
        End If
    Next i
    
    endrow = Range("cw10000").End(xlUp).Row
    
    Range("CW4:DC" & endrow).Select
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("cw5:cw" & endrow), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("CW4:DC" & endrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
   
'----end clear and sort data--------------------------------------------------------------------------

ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
Delete_Pictures '__Alex_x§d



End Sub






Sub HK_Yahoo_Price_Dividend_Split(cd$, dc)
            
            ActiveSheet.Range("CD3:CK100000").ClearContents
            Dim myurl(2)
            myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
            
            myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & ".HK?period1=57600&period2=" & myday & "&interval=1mo&events=history&includeAdjustedClose=true"
            myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & ".HK?period1=57600&period2=" & myday & "&interval=1mo&events=div&includeAdjustedClose=true"
            myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & ".HK?period1=57600&period2=" & myday & "&interval=1mo&events=split&includeAdjustedClose=true"
            
            
            
            Dim url As String, crumbrng As Range, crumbrng1 As Range
            Dim serr As String, csvt As Variant, csv As Variant
            Dim cii As Long, cjj As Long, ci As Long, cj As Long
        
            
            For iurl = 0 To 2
                url = myurl(iurl)
                If iurl = 0 Then 'price
                    
                    endrow = 3
                Else
                    endrow = Range("CE100000").End(xlUp).Row
                End If
                Set crumbrng = Cells(endrow, dc) ' for price + dividend+split
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
            
                endrow = Range("CE10000").End(xlUp).Row
                For i = endrow To 7 Step -1
                    If Range("CE" & i) = "Date" Or Len(Range("CE" & i)) = 0 Then
                         Range("CE" & i & ":CK" & i).Delete Shift:=xlUp
                    
                  
                    End If
                Next i
                
            
            
                endrow = Range("CE10000").End(xlUp).Row
                
                Range("CE4:CK" & endrow).Select
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("CE5:CE" & endrow), _
                    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                With ActiveWorkbook.ActiveSheet.Sort
                    .SetRange Range("CE4:CK" & endrow)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
               
            '----end clear and sort data--------------------------------------------------------------------------
            
            
End Sub


