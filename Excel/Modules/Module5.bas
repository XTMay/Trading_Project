Attribute VB_Name = "Module5"
Sub Macro5()
'
' mikeon _ 2014/5/1 _____的巨集
'

Sheets("__").Selectct

Dim waitT As Integer: waitT = 5

If UCase([a1]) = "" Or UCase([a1]) = "US" Then
[b11] = "It takes 1 min. Please wait..."
Else
[b11] = "It takes 1 min 57 sec. Be patient..."
End If



With Application
    '.Calculation = xlManual '____徐廣
    .Calculation = xlAutomatic '_____}廣福
    .EnableCancelKey = xlInterrupt '_____\秀雯
    .MaxChange = 0.001 '____速度
    .ScreenUpdating = False
End With




Call UnprotectSheet(ActiveSheet)

     With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
    

On Error GoTo err


[a56] = "______________(Michael On)__"Michael On)所有"
[i12] = Left(ActiveWorkbook.Name, InStr(1, ActiveWorkbook.Name, ".") - 1)

    Range("E9") = "=IFERROR(R[38]C[14],""na"")"
    Range("k3") = "=R[3]C[6]"
    [k10] = 12
    Range("k11") = "=R[2]C[19]"
    Range("K12") = "=IF(UPPER(R[-11]C[-10])=""HK"",R[27]C[18],R[18]C[18])"
    Range("O2") = "=VLOOKUP(R[5]C[-1]-ROUNDUP(RC[2],0)+1,R[5]C[-1]:R[35]C[2],2,FALSE)"
    Range("P2") = "=VLOOKUP(R[5]C[-2]-ROUNDUP(RC[1],0)+1,R[5]C[-2]:R[35]C[1],4,FALSE)"
    Range("q2") = "=YEAR(TODAY())-IF(R[13]C[-2]<>"""",R[13]C[-3],IF(R[12]C[-2]<>"""",R[12]C[-3],IF(R[11]C[-2]<>"""",R[11]C[-3],IF(R[10]C[-2]<>"""",R[10]C[-3],IF(R[9]C[-2]<>"""",R[9]C[-3],IF(R[8]C[-2]<>"""",R[8]C[-3],IF(R[7]C[-2]<>"""",R[7]C[-3],IF(R[6]C[-2]<>"""",R[6]C[-3],R[5]C[-3]-MONTH(TODAY())/12))))))))"
    Range("A2").NumberFormatLocal = "@"
    
    [q6] = "": [f15] = 1: [ac25] = "USD": [ae23] = 0: [ae24] = 0: [ae25] = "": [ac39] = "": [i14] = "": pnd = 1
    
Range("o7:p37").ClearContents
Range("ag78:ag88").ClearContents
dc1 = 37 '____益表
dc5 = 73 '____ bu bu
dc6 = 82 '___+__+__+__股價+日期
dc7 = 85 '__率
dccf = 89 '____流表
dcmc = 98 '__值
dc10 = 102 '____單位
gr = "Y"
[a12] = "If you are in China, please enter Y in the right cell."


If [a1] = "" Then [a1] = "US"
nd$ = UCase([a1])

dc = dc1
Range(Columns(dc), Columns(dc + 100)).Clear

cd$ = wsj([a2])
If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
Application.StatusBar = "Quarterly income statement  1 / 9"

Dim url As String
    url = "https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/financials/quarter/income-statement"
        Cells(1, dc) = "1 / 9 Quarterly income statement " & url
        
        
        no_retry_report = 1
        
PL_Q:
        
        Call ConnectMarketWatch(url, Cells(2, dc), 2)
           
        ' Debug.Print Cells(3, dc)
         
        If Len(Cells(3, dc)) = 0 Then
        
            If no_retry_report < 4 Then
              
               Application.Wait Now() + TimeValue("00:00:01") * waitT
               no_retry_report = no_retry_report + 1
                Debug.Print url, "Retry:" & no_retry_report
               GoTo PL_Q
               
            ElseIf no_retry_report >= 4 Then
        
                 
                MsgBox "_____________(" & url & "), _____", vbOKOnlyy後再試", vbOKOnly
                
                Call ProtectSheet(ActiveSheet)
                
                Exit Sub
            End If
        End If
           
        
        'Call ConnectXMLHTTP(url)
        
        'Call ListData(Cells(2, dc))
        
        
For i = 1 To 5
yw = Cells(2, dc1 + i)
Cells(2, dc1 + i) = yymmx(yw)
Next i

If Right(Cells(2, dc1 + 1), 4) = "0001" Then   '--------0001---------------------------
If Abs(Month(Cells(2, dc1 + 3)) - Month(Cells(2, dc1 + 4))) < 5 Then
Cells(2, dc1 + 1) = Cells(2, dc1 + 5) + 365
Else
Cells(2, dc1 + 1) = Cells(2, dc1 + 3) + 365
End If
End If

If Right(Cells(2, dc1 + 2), 4) = "0001" Then '--------0001-------1218.TW--------------------
If Abs(Month(Cells(2, dc1 + 3)) - Month(Cells(2, dc1 + 4))) < 5 Then
Cells(2, dc1 + 2) = Cells(2, dc1 + 3) + 91
Else
Cells(2, dc1 + 2) = Cells(2, dc1 + 3) + 183
End If
End If

   For i = 1 To 10
   If (Left(Cells(i, dc), 1) = "F") Then Exit For
   Next i
   
   [ag13] = Cells(i + 1, dc)
   [ab33] = Cells(i, dc)
   
         [m1] = ""
         If Len(Sheets("__").[t3]) > 5 Thenen
         [m1] = Sheets("__").[t3]3]
         
         If DateDiff("d", [m1], [af78]) < 40 Then
         [m1] = "N"
         [b11] = "The new financial report is not yet ready."
         GoTo __式
         End If
         End If
         
         For m = 2 To 20
         If Cells(m, dc1) <> "" Then Exit For
         Next m
         If m > 19 Then GoTo __式
         
   
    cd$ = wsj([a2])
    If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
    
    Application.StatusBar = "Quarterly balance sheet  2 / 9"
    url = "https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/financials/quarter/balance-sheet"
        Cells(1, dc + 9) = "2 / 9 Quarterly balance sheet " & url
        
        no_retry_report = 1
BS_Q:

        Call ConnectMarketWatch(url, Cells(2, dc + 9), 2)

        
        If Len(Cells(3, dc + 9)) = 0 Then
           If no_retry_report < 4 Then
        
              Application.Wait Now() + TimeValue("00:00:01") * waitT
              no_retry_report = no_retry_report + 1
               Debug.Print url, "Retry:" & no_retry_report
              GoTo BS_Q
            ElseIf no_retry_report >= 4 Then
                MsgBox "_____________(" & url & "), _____", vbOKOnlyy後再試", vbOKOnly
                 Call ProtectSheet(ActiveSheet)
                Exit Sub
            End If
        End If
         
         
[ab34] = Cells(i, dc + 9)

For i = 1 To 5
yw = Cells(2, dc1 + 9 + i)
Cells(2, dc1 + 9 + i) = yymmx(yw)
Next i

If Right(Cells(2, dc1 + 10), 4) = "0001" Then '--------0001---------------------------
If Abs(Month(Cells(2, dc1 + 13)) - Month(Cells(2, dc1 + 14))) < 5 Then
Cells(2, dc1 + 10) = Cells(2, dc1 + 14) + 365
Else
Cells(2, dc1 + 10) = Cells(2, dc1 + 12) + 365
End If
End If


If Right(Cells(2, dc1 + 11), 4) = "0001" Then '--------0001-------1218.TW--------------------
If Abs(Month(Cells(2, dc1 + 12)) - Month(Cells(2, dc1 + 13))) < 5 Then
Cells(2, dc1 + 11) = Cells(2, dc1 + 12) + 91
Else
Cells(2, dc1 + 11) = Cells(2, dc1 + 12) + 183
End If
End If

  
cd$ = wsj([a2])
If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
Application.StatusBar = "Annual income statement  3 / 9"
url = "https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/financials/annual/income-statement"
        Cells(1, dc + 18) = "3 / 9 Annual income statement " & url
        
         no_retry_report = 1
PL_Y:

        Call ConnectMarketWatch(url, Cells(2, dc + 18), 2)
     
        If Len(Cells(3, dc + 18)) = 0 Then
            If no_retry_report < 4 Then
        
               Application.Wait Now() + TimeValue("00:00:01") * waitT
               no_retry_report = no_retry_report + 1
               Debug.Print url, "Retry:" & no_retry_report
               GoTo PL_Y
            ElseIf no_retry_report >= 4 Then
                 MsgBox "_____________(" & url & "), _____", vbOKOnlyy後再試", vbOKOnly
                 Call ProtectSheet(ActiveSheet)
                Exit Sub
            End If
        End If
    
    
    For i = 1 To 10
   If (Left(Cells(i, dc + 18), 1) = "F") Then Exit For
   Next i
   
   For j = 1 To 5
   Cells(77 + j + 6, 32) = Cells(i, dc + 18 + j)
   Next j
   
   [af64] = Cells(i, dc + 18)
   [ab35] = Cells(i, dc + 18)
   [ae25] = Mid([af64], InStr(1, [af64], "values") + 7, 3) '___v用
    
 cd$ = wsj([a2])
 If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
Application.StatusBar = "Annual balance sheet  4 / 9"
url = "https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/financials/annual/balance-sheet"
        Cells(1, dc + 27) = "4 / 9 Annual balance sheet " & url
        
        no_retry_report = 1
BS_Y:
        
        Call ConnectMarketWatch(url, Cells(2, dc + 27), 2)
   
        If Len(Cells(3, dc + 27)) = 0 Then
           If no_retry_report < 4 Then
        
               Application.Wait Now() + TimeValue("00:00:01") * waitT
               no_retry_report = no_retry_report + 1
               Debug.Print url, "Retry:" & no_retry_report
               GoTo BS_Y
            ElseIf no_retry_report >= 4 Then
                MsgBox "_____________(" & url & "), _____", vbOKOnlyy後再試", vbOKOnly
                Call ProtectSheet(ActiveSheet)
                Exit Sub
            End If
        End If
   
   
   [ab36] = Cells(i, dc + 27)
   
   
Application.StatusBar = "Historical stock price  5 / 9" '__Alex_______M方家晟桑



If UCase([f12]) = "Y" Then GoTo CN1


yrs$ = [n7] - 30
yre$ = [n7]

'Call glhis '__20_______yahoo____yahoo已不給抓

'___8_____歷史股價
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp

dc = dc5
cd$ = yho([a2])
If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
Call yhondd(nd$, cd$)
url = "https://finance.yahoo.com/quote/" & cd$ & "." & nd$ & "/history?period1=573436800&period2=" & myday & "&interval=1mo&filter=history&frequency=1mo"
Cells(1, dc) = "5 / 9 Historical stock price " + url

Call WW_Yahoo_Price_Dividend_Split(cd$, nd$, dc)


CN1:

   
cd$ = wsj([a2])
If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
Application.StatusBar = "Comapny name + Profile + Stock price + Date  6 / 9"
dc = dc6
    Cells(1, dc) = "6 / 9 Company name + Profile + Stock price + Date https://www.wsj.com/market-data/" & [a1] & "/" & cd$ & "/company-people"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/company-people", Destination:=Cells(2, dc))
        .Name = "company-people"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With 'by ___s福
 '   ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
 '   Delete_Pictures '__Alex_x吳


   i = 2: Do Until Cells(i, dc) = "View All companies" Or i > 1000
    i = i + 1
   Loop
   If i < 999 Then
   [b1] = Cells(i - 2, dc) '___q名
   Cells(i - 2, dc - 1) = 1
   Else: [b1] = Cells(i - 3, dc)
   Cells(i - 3, dc - 1) = 1
   End If
   
   For i = i To i + 10
   If InStr(1, Cells(i, dc), "/") > 0 Then Exit For
   Next i
   [i1] = Right(Cells(i, dc), 8) '__期
   Cells(i, dc - 1) = 1
   If Cells(i + 1, dc) <> "" Then
   [q6] = Cells(i + 1, dc) '__價
   Cells(i + 1, dc - 1) = 1
   Else: [q6] = Cells(i + 2, dc)
   Cells(i + 2, dc - 1) = 1
   End If
   [ac25] = UCase(Right(Trim([q6]), 3)) '____幣別
   
   a = Left(Trim([q6]), Len(Trim([q6])) - 3)
   b = Right(Trim(a), Len(Trim(a)) - 1)
   c = Right(Trim(b), Len(Trim(b)) - 1)
   D = Right(Trim(c), Len(Trim(c)) - 1)
   If D <> "" And IsNumeric(D) Then [q6] = D
   If c <> "" And IsNumeric(c) Then [q6] = c
   If b <> "" And IsNumeric(b) Then [q6] = b
   If a <> "" And IsNumeric(a) Then [q6] = a
   
   If UCase([a1]) = "UK" Or UCase([a1]) = "ZA" Then pnd = 0.01
   If UCase([a1]) = "UK" And Left(cd$, 1) = "0" Then pnd = 1
   [q6] = [q6] * pnd
   
   Do Until Left(Cells(i, dc), 6) = "Sector" Or i > 1000
   i = i + 1
   Loop
   If i < 999 Then
   [b1] = [b1] & "_" & Trim(Cells(i, dc)))
   Cells(i, dc - 1) = 1
   End If
   
   Do Until Left(Cells(i, dc), 8) = "Industry" Or i > 1000
   i = i + 1
   Loop
   If i < 999 Then
   [b1] = [b1] & "*" & Trim(Cells(i, dc))
   Cells(i, dc - 1) = 1
   End If
   
   aa = InStr(1, [b1], "Sector")
   If aa > 0 Then [b1] = Left([b1], aa - 1) + Right([b1], Len([b1]) - aa - Len("Sector ") + 1)
   aa = InStr(1, [b1], "Industry")
   If aa > 0 Then [b1] = Left([b1], aa - 1) + Right([b1], Len([b1]) - aa - Len("Industry ") + 1)
   aa = InStr(1, [b1], "Companies on the")
   If aa > 0 Then [b1] = Left([b1], aa - 1) + Right([b1], Len([b1]) - aa - Len("Companies on the ") + 1)


'-----------------------------------------------------------------



Application.StatusBar = "Foreign Exchange  7 / 9"

   If [ae25] = "" Then
   For i = 1 To 10
   If (Left(Cells(i, 37), 1) = "F") Then Exit For
   Next i
   [af64] = Cells(i, 37)
   [ae25] = Mid([af64], InStr(1, [af64], "values") + 7, 3) '___v用
   End If
   
   If [ae25] = "" Then
   For i = 1 To 10
   If (Left(Cells(i, 37), 1) = "F") Then Exit For
   Next i
   [af64] = Cells(i, 37)
   Range("AE25").Select
    ActiveCell.FormulaR1C1 = "= MID(R[39]C[1], FIND(""values"",R[39]C[1])+ 7, 3)"
   End If

   If Left([ac25], 2) = Left([ae25], 2) Then GoTo cf
   
   dc = dc7
     'url = "https://transferwise.com/zh-hk/currency-converter/" & [ac25] & "-to-" & [ae25] & "-rate" '_____yahooyahoo
     url = "https://wise.com/zh-hk/currency-converter/" & [ac25] & "-to-" & [ae25] & "-rate" '_____yahooyahoo
     
     
     Cells(1, dc) = "7 / 9 Foreign exchange " & url
                                
            Call ConnectWinHttp(url, 1)
            Debug.Print url
            
                                
                              
                                
            If InStr(1, doc.body.innerHTML, "HTTP ERROR") >= 1 Then
                
               Debug.Print "FX not found"
               
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
   
Application.StatusBar = "Annual cashflow statement  8 / 9"
cd$ = wsj([a2])
If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
dc = dccf
   Cells(1, dc) = "8 / 9 Annual cashflow statement https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/financials/annual/cash-flow"
    
     
     url = "https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/financials/annual/cash-flow"
           
     no_retry_report = 1
CF_Y:
        
        Call ConnectMarketWatch(url, Cells(2, dc), 2)
   
        If Len(Cells(3, dc)) = 0 Then
           If no_retry_report < 4 Then
        
               Application.Wait Now() + TimeValue("00:00:01") * waitT
               no_retry_report = no_retry_report + 1
               Debug.Print url, "Retry:" & no_retry_report
               GoTo CF_Y
            ElseIf no_retry_report >= 4 Then
                MsgBox "_____________(" & url & "), _____", vbOKOnlyy後再試", vbOKOnly
                Call ProtectSheet(ActiveSheet)
                Exit Sub
            End If
        End If
    

    
    
'    With ActiveSheet.QueryTables.Add(Connection:= _
'        "URL;https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/financials/annual/cash-flow", _
'        Destination:=Cells(2, dc))
'        .Name = "annual cash-flow"
'        .FieldNames = True
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlOverwriteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = False
'        .RefreshPeriod = 0
'        .WebSelectionType = xlSpecifiedTables
'        .WebFormatting = xlWebFormattingNone
'        .WebTables = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99"
'        .WebPreFormattedTextToColumns = True
'        .WebConsecutiveDelimitersAsOne = True
'        .WebSingleBlockTextImport = False
'        .WebDisableDateRecognition = False
'        .WebDisableRedirections = False
'        .Refresh BackgroundQuery:=False
'   End With 'by ___s福
'   ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
'   Delete_Pictures '__Alex_x吳
'
   For i = 1 To 10
   If (Left(Cells(i, dc), 1) = "F") Then Exit For
   Next i
   [ab37] = Cells(i, dc)
   


For i = 2 To 5
For j = dc + 8 To dc Step -1
If Right(Trim(Cells(i, j)), 5) = "trend" Then GoTo cyyrr
Next j
Next i

cyyrr:
For k = 1 To 6
Cells(90 + k, 32) = ""
Next k


For k = 1 To 5
Cells(90 + k, 32) = Cells(i, j - k)
If Not IsNumeric(Cells(90 + k, 32)) Then Cells(90 + k, 32) = ""
Next k


If UCase([a1]) = "US" Then
Application.StatusBar = "US Market capitalization  9 / 9"
dc = dcmc

If gr = "Y" Then
cd$ = mkw([a2])
        url = "https://www.gurufocus.com/stock/" & cd$ & "/guru-trades" ' RDS.B ____錯誤
        Cells(1, dc) = "9 / 9 US Market capitalization " & url
        Call ConnectMarketWatch(url, Cells(2, dc), 2)
        
        i = 2
        Do Until Trim(Cells(i, dc)) = "Market Cap $ M" Or i > 100
        i = i + 1
        Loop
        [ae23] = Cells(i, dc + 1)
End If
        
If gr = "N" Then
cd$ = yho([a2])
url = "https://finance.yahoo.com/quote/" & cd$
        Cells(1, dc) = "9 / 9 Market capitalization " & url
        With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://finance.yahoo.com/quote/" & cd$, Destination:=Cells(2, dc))
        .Name = "company-people"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With 'by ___s福
        
    i = 2
    Do Until Trim(Cells(i, dc)) = "Market Cap" Or i > 100
    i = i + 1
    Loop
    unitt = Cells(i, dc + 1)
    [ae23] = unitconversion(unitt)
    Cells(i, dc - 1) = 1
    
End If
End If


If UCase([a1]) = "HK" Then
Application.StatusBar = "Trading unit  10 / 10"
dc = dc10
url = "https://www.aastocks.com/tc/stocks/analysis/company-fundamental/basic-information?symbol=" & cd$
Cells(1, dc) = "10 / 10 Trading unit " + url
Call ConnectMarketWatch(url, Cells(2, dc), 2)
    i = 2
    Do Until Cells(i, dc) = "____" Or i > 10001000
    i = i + 1
   Loop
   Cells(2, dc) = "____"璁"
   Cells(3, dc) = Cells(i, dc + 1): [ac39] = Cells(i, dc + 1)
End If


err:
[b11] = ""

'---------------------------------------------------------------------------------------------------

If UCase([f12]) = "Y" Then GoTo CN2

dc = dc5
For c = 2 To 500
If Cells(c, dc) = "Date" Or Cells(c, dc) = "__" Then Exit Foror
Next c
If c = 500 Then GoTo __式
c = c + 1

t = 7: b = 37
Call highlow(dc, c, t, b, pnd)


CN2:

If [q6] = "" Or [q6] = 0 Then
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

__:   [ae24] = [q6]6]
        
If [b11] = "The new financial report is not yet ready." Then GoTo y16
     
If [q6] = 0 Or [q6] = "" Then [b11] = "___Stock prices are missing."g."

If UCase([a1]) = "US" Then
For m = 2 To 20
If Cells(m, dcmc) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___9 / 9 Latest market capitalization is missing."g."
End If

For m = 2 To 20
If Cells(m, dccf) <> "" Then Exit For
Next m
If m > 19 Then
[aa46] = "___"流"
If [ac25] <> [ae25] Then [b11] = "___7 / 9 Cashflow statements are missing."g."
End If

If UCase([f12]) <> "Y" Then
For m = 5 To 20
If Cells(m, dc5) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___5 / 9 Div, split & historic prices are missing."g."
End If

For m = 2 To 20
If Cells(m, dc1 + 27) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___4 / 9 Annual balance sheets are missing."g."

For m = 2 To 20
If Cells(m, dc1 + 18) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___3 / 9 Annual income statements are missing."g."

For m = 2 To 20
If Cells(m, dc1 + 9) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___2 / 9 Quarterly balance sheets are missing."g."

For m = 2 To 20
If Cells(m, dc1) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___1 / 9 Quarterly income statements is missing."g."
     
For m = 2 To 20
If Cells(m, dc7) <> "" Then Exit For
Next m

For i = 4 To 40
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
    
    Range(Cells(2, 2), Cells(54, 10)).ShrinkToFit = True '___Y小
    [f9].ShrinkToFit = False '____縮小
    [i12].ShrinkToFit = False '____縮小
    [e15].ShrinkToFit = False '____縮小
    [h14].ShrinkToFit = False '____縮小
    [g44].ShrinkToFit = False '____縮小
    [g45].ShrinkToFit = False '____縮小
    Range(Cells(18, 11), Cells(25, 11)).ShrinkToFit = False '____縮小
    Range(Cells(37, 11), Cells(47, 11)).ShrinkToFit = False '____縮小
    
    Range(Columns(dc1), Columns(dc1 + 100)).Select '___Y小
        With Selection
        .WrapText = False
        .ShrinkToFit = True
    End With
    
    Range(Cells(1, dc1), Cells(1, dc1 + 100)).ShrinkToFit = False '____縮小
  
    [j1].ColumnWidth = 8
    Cells(1, dc1).ColumnWidth = 20
    Cells(1, dc1 + 9).ColumnWidth = 20
    Cells(1, dc1 + 9 * 2).ColumnWidth = 20
    Cells(1, dc1 + 9 * 3).ColumnWidth = 20
    Cells(1, dccf).ColumnWidth = 30
    
    Call GDPSCR(nd$, cd$)
    
    If UCase([a1]) = "HK" Then
    Range("K12").NumberFormatLocal = "#,##0_);(#,##0)"
    Else
    Range("K12").NumberFormatLocal = "0%"
    End If
    
    Range("c3:d8,f3:i8, k3:k7, r1:r11, g38:g42, g47:g52, a13:i13").Select
    With Selection
        .HorizontalAlignment = xlCenter '___m中
        .VerticalAlignment = xlCenter
    End With
    
    Columns(dc5).NumberFormatLocal = "yyyy/m/d;@"

    Range("c3:e9, b20:b24, e20:j24, b29:b34, e29:j34, j35:j35,b39:f43, h39:i43, b48:f54, h48:i54").Select
    Selection.NumberFormatLocal = "#,##0_);(#,##0)"
        
    Range("e15").HorizontalAlignment = xlRight
    Range("f15").HorizontalAlignment = xlLeft
    
    Range("k1,a13:i13").Select
    With Selection.Font
        .Size = 12
        .Name = "Arial"
    End With

Range("B11").ShrinkToFit = False
Application.StatusBar = "Done !"
    '__紹
    i = 2: Do Until Left(Trim(Cells(i, dc6)), 11) = "Description" Or i > 1000
    i = i + 1
   Loop
   ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
  ' Delete_Pictures '__Alex_x吳
   
   Cells(i, dc6 - 1) = 1
   If Cells(i + 1, dc6) <> "" Then
   Cells(i + 1, dc6 - 1) = 1
     Cells(i + 1, dc6).Copy
   Else
   Cells(i + 2, dc6 - 1) = 1
   Cells(i + 2, dc6).Copy
   End If
   
     [ae16].Select: ActiveSheet.Paste
     ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
Application.Calculation = xlAutomatic           '_____}廣福

Exit Sub

'--------------------------
'Errorhandler_IE:

'    If InStr(1, err.Description, "Automation") >= 1 Then
'        Debug.Print err.Number, err.Description
'        Call DelIE
       
'        GoTo ResetIE
        
'    Else
'
'       Debug.Print "not IE automation issue", err.Number, err.Description
'
'       Resume Next
'    End If
    

End Sub


Public Function CCRUMBT_BAK(rng As Range) '______家晟桑
    On Error Resume Next
    Dim crumbt As String, crumb As String
    err.Clear
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", "https://finance.yahoo.com/recent-quotes", False
        .Option(4) = 13056
        .Option(6) = False
        .SetTimeouts 500, 5000, 5000, 5000
        .send
        .waitForResponse
        crumbt = .responseText
        crumb = InStrRev(crumbt, """crumb""" & ":" & """")
        crumb = Mid(crumbt, Val(crumb) + 9, 11)
        rng.Offset(-1, 0).value = crumb
    End With
    CCRUMBT = err.Number
    err.Clear
    
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
   ' Delete_Pictures '__Alex_x吳
End Function

Public Function DCSV_BAK(url As String) '______家晟桑
    On Error Resume Next
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "POST", url, False
        .Option(4) = 13056
        .Option(6) = False
        .SetTimeouts 500, 5000, 5000, 5000
        .send
        .waitForResponse
        DCSV = split(.responseText, Chr(10))
    End With
    
 '   ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
 '   Delete_Pictures '__Alex_x吳
End Function

Public Sub yhondd(nd$, cd$)

If UCase([a1]) = "JP" Then nd$ = "T"
If UCase([a1]) = "UK" Then nd$ = "L"
If UCase([a1]) = "UK" And Left(cd$, 1) = "0" Then nd$ = "IL"
If UCase([a1]) = "FR" Then nd$ = "PA"
If UCase([a1]) = "CA" Then nd$ = "TO"
If UCase([a1]) = "AU" Then nd$ = "AX"
If UCase([a1]) = "KR" Then nd$ = "KS"
If UCase([a1]) = "DK" Then nd$ = "CO"
If UCase([a1]) = "SE" Then nd$ = "ST"
If UCase([a1]) = "FI" Then nd$ = "HE"
If UCase([a1]) = "NL" Then nd$ = "AS"
If UCase([a1]) = "CH" Then nd$ = "SW"
If UCase([a1]) = "MY" Then nd$ = "KL"
If UCase([a1]) = "SG" Then nd$ = "SI"
If UCase([a1]) = "IT" Then nd$ = "MI"
If UCase([a1]) = "ES" Then nd$ = "MC"
'If ucase([a1]) = "CH" And [ac46] Then nd$ = "VX"
'[ac46]==ISNUMBER(FIND("Europe",B1)/1)

If UCase([a1]) = "CN" Then
    nd$ = "SS"
    If Left(cd$, 1) = "0" Or Left(cd$, 1) = "2" Or Left(cd$, 1) = "3" Then
    nd$ = "SZ"
    End If
End If



End Sub

Public Sub GDPSCR(nd$, cd$)

    [a13] = ""
    [h13].Select: cd$ = yho([a2])
    If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
    Call yhondd(nd$, cd$)
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/quote/" & cd$ & "." & nd$ & "?p=" & cd$ & "." & nd$, _
    TextToDisplay:="Chart"
    
    [g13].Select: cd$ = yho([a2])
    If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/quote/" & cd$ & "." & nd$ & "?p=" & cd$ & "." & nd$, _
    TextToDisplay:="News"
    
    [f13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://mikeon88.blogspot.tw/2014/02/blog-post_2047.html", _
    TextToDisplay:="Screen"
    
    [d13].Select: cd$ = mkw([a2])
    If UCase([a1]) = "HK" Then cd$ = Format([a2], "0000")
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.wsj.com/market-data/quotes/" & [a1] & "/" & cd$ & "/company-people", _
    TextToDisplay:="Profile"
    
If UCase([a1]) = "HK" Then
    [g13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://hk.finance.yahoo.com/q/h?s=" & cd$ & ".HK", TextToDisplay:="News"
End If

If UCase([a1]) = "CN" Then
    [h13].Select: cd$ = yho([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/quote/" & cd$ & ".SS" & "?p=" & cd$ & ".SS", _
    TextToDisplay:="Chart"
    
    [g13].Select: cd$ = yho([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/quote/" & cd$ & ".SS" & "?p=" & cd$ & ".SS", _
    TextToDisplay:="News"
    
    If Left(cd$, 1) = "0" Or Left(cd$, 1) = "2" Or Left(cd$, 1) = "3" Then
      [h13].Select: cd$ = yho([a2])
      Selection.Hyperlinks.Add Anchor:=Selection, _
      Address:="https://finance.yahoo.com/quote/" & cd$ & ".SZ" & "?p=" & cd$ & ".SZ", _
      TextToDisplay:="Chart"
      
       [g13].Select: cd$ = yho([a2])
      Selection.Hyperlinks.Add Anchor:=Selection, _
      Address:="https://finance.yahoo.com/quote/" & cd$ & ".SZ" & "?p=" & cd$ & ".SZ", _
      TextToDisplay:="News"
    End If
    
End If
    If UCase([a1]) = "TW" Then
    [f13].Select
     Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://mikeon88.blogspot.tw/2014/02/blog-post_16.html", TextToDisplay:="Screen"
    
    [g13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://tw.stock.yahoo.com/q/q?s=" & cd$, TextToDisplay:="News"
End If

If UCase([a1]) = "" Or UCase([a1]) = "US" Then 'US
    [h13].Select: cd$ = yho([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/quote/" & cd$ & "/?p=" & cd$, _
    TextToDisplay:="Chart"
                   
    [g13].Select: cd$ = yho([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/quote/" & cd$ & "/?p=" & cd$, _
    TextToDisplay:="News"
    
    [f13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finviz.com/screener.ashx?v=111&f=fa_pe_u15,fa_roe_o10,idx_sp500&ft=4", _
    TextToDisplay:="Screen"
    
    [d13].Select: cd$ = mkw([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://abxusa.com/" & cd$, _
    TextToDisplay:="Profile"
    'Address:="https://www.mg21.com/" & cd$ & ".html", _

    [a13].Select: cd$ = mkw([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://wealth.esunbank.com.tw/usstock/esun/basic0001.xdjhtm?a=" & cd, _
    TextToDisplay:="Profl-e"
    'Address:="https://www.moneydj.com/us/basic/basic0001/" & cd, _

End If

sk:
End Sub

Sub Macro6()
'
'
' ASUS _ 2014/9/8 _____的巨集
'

'
Sheets("__").Selectct
Application.Calculation = xlAutomatic '_____}廣福
Application.Calculation = xlManual '____徐廣
Call UnprotectSheet(ActiveSheet)

    Application.MaxChange = 0.001 '____速度
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
    
With Application '_____}廣福
     .EnableCancelKey = xlInterrupt '_____\秀雯
End With

On Error GoTo err
Application.ScreenUpdating = False

If [a1] = "" Then [a1] = "US"
If UCase([a1]) = "US" Then nd$ = "united-states"
If UCase([a1]) = "HK" Then nd$ = "hong-kong"
If UCase([a1]) = "CN" Then nd$ = "china"
If UCase([a1]) = "JP" Then nd$ = "japan"
If UCase([a1]) = "DE" Then nd$ = "germany"
If UCase([a1]) = "UK" Then nd$ = "united-kingdom"
If UCase([a1]) = "FR" Then nd$ = "france"
If UCase([a1]) = "CA" Then nd$ = "canada"
If UCase([a1]) = "AU" Then nd$ = "australia"
If UCase([a1]) = "ES" Then nd$ = "spain"
If UCase([a1]) = "BR" Then nd$ = "brazil"
If UCase([a1]) = "RU" Then nd$ = "russia"
If UCase([a1]) = "TH" Then nd$ = "thailand"
If UCase([a1]) = "MY" Then nd$ = "malaysia"
If UCase([a1]) = "ID" Then nd$ = "indonesia"
If UCase([a1]) = "KR" Then nd$ = "south-korea"
If UCase([a1]) = "ZA" Then nd$ = "south-africa"
If UCase([a1]) = "SG" Then nd$ = "singapore"
If UCase([a1]) = "TW" Then nd$ = "taiwan"
If UCase([a1]) = "FI" Then nd$ = "finland"
If UCase([a1]) = "NZ" Then nd$ = "new-zealand"

[e13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://tradingeconomics.com/" & nd$ & "/gdp-growth-annual", _
    TextToDisplay:="GDP"


If UCase([a1]) = "TW" Then
    [e13].Select
     Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://mikeon88.blogspot.tw/2012/02/blog-post_12.html", TextToDisplay:="GDP"
End If


err:

[e13].Select
    Selection.RowHeight = 16
    Selection.ColumnWidth = 6.5
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "__""
        .Size = 12
    End With

    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic           '_____}廣福
Application.StatusBar = "__""
Call ProtectSheet(ActiveSheet)

End Sub






Private Sub glhis() '__20_______yahoo____yahoo已不給抓
Dim myurl(2)
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & nd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=history&crumb="
myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & nd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=div&crumb="
myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "." & nd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=split&crumb="

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
        endrow = Range("bu100000").End(xlUp).Row
    End If
    Set crumbrng = Cells(endrow, dc) ' for price + dividend+split
   
retryno = 0
rest:

    csvt = DCSV(url)
     If InStr(1, csvt(0), "<!doctype html public") >= 1 Then
  
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
    
    endrow = Range("bu10000").End(xlUp).Row
    For i = endrow To 7 Step -1
        If Range("bu" & i) = "Date" Or Len(Range("bu" & i)) = 0 Then
             Range("BU" & i & ":CA" & i).Delete Shift:=xlUp
      
        End If
    Next i
    
    endrow = Range("bu10000").End(xlUp).Row
    
    Range("BU4:CA" & endrow).Select
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("bu5:bu" & endrow), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("BU4:CA" & endrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'----end clear and sort data--------------------------------------------------------------------------


End Sub





Sub WW_Yahoo_Price_Dividend_Split(cd$, nd$, dc)


            
            
            nd$ = "." & nd$
            If InStr(1, UCase(nd$), "US") >= 1 Then nd$ = ""
            
            ActiveSheet.Range(Cells(2, dc - 1), Cells(10000, dc + 6)).ClearContents
            Dim myurl(2)
            myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
            
            myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & nd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=history&includeAdjustedClose=true"
            myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & nd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=div&includeAdjustedClose=true"
            myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & nd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=split&includeAdjustedClose=true"
            
            
            
            
            Dim url As String, crumbrng As Range, crumbrng1 As Range
            Dim serr As String, csvt As Variant, csv As Variant
            Dim cii As Long, cjj As Long, ci As Long, cj As Long
        
            
            For iurl = 0 To 2
                url = myurl(iurl)
                ' Debug.Print url
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
                
                cii = UBound(csvt) + 1
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
            
nd$ = UCase([a1])

End Sub


Public Function yymmx(yw)

If yw = "" Then Exit Function
[ar2].Select
Selection.NumberFormat = "yyyy/m/d"
[ar2] = yw
yymmx = [ar2]


[ar2] = ""
End Function

