Attribute VB_Name = "Module2"
Public doc As Object, ConnectErr As Boolean, IE As Object


Sub Macro2()
'
' mikeon _ 2005/2/13 _____ªº¥¨¶°
'

Sheets("__").Selectct
Application.Calculation = xlAutomatic '_____}¼sºÖ
[b11] = "It takes 1 min 41 sec. Be patient..."
Application.Calculation = xlManual '_____}¼sºÖ
Call UnprotectSheet(ActiveSheet)

    With ActiveWorkbook '____³t«×
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
    With Application
            .EnableCancelKey = xlInterrupt '_____\¨q¶²
            .MaxChange = 0.001
            .ScreenUpdating = False
            .DisplayAlerts = False
    End With

On Error GoTo err
Dim x As Integer
Dim sUrl(5, 3) ' Define parameters for reports
Dim IEVisible As Boolean: IEVisible = False ' visible or not when navigate IE
Dim WT As Integer: WT = 3  ' IE wait time
 
 [a16] = "Please indicate the source when citing this table. Those who quote or rewrite other people's ideas, formulas,"
 [a17] = "and stock tips without citing the source are thieves."
 [a56] = "______________(Michael On)__"Michael On)©Ò¦³"
 [i12] = Left(ActiveWorkbook.Name, InStr(1, ActiveWorkbook.Name, ".") - 1)

    Range("E9").FormulaR1C1 = "=IFERROR(R[41]C[14],""na"")"
    Range("k3").FormulaR1C1 = "=R[6]C[6]"
    [k10] = 12
    Range("k11").FormulaR1C1 = "=R[2]C[13]"
    Range("K12").FormulaR1C1 = "=R[18]C[12]" '_____©µ|²v
 
    Range("O5").FormulaR1C1 = "=VLOOKUP(R[5]C[-1]-ROUNDUP(RC[2],0)+1,R[5]C[-1]:R[35]C[2],2,FALSE)"
    Range("P5").FormulaR1C1 = "=VLOOKUP(R[5]C[-2]-ROUNDUP(RC[1],0)+1,R[5]C[-2]:R[35]C[1],4,FALSE)"
    Range("q5").FormulaR1C1 = "=YEAR(TODAY())-IF(R[13]C[-2]<>"""",R[13]C[-3],IF(R[12]C[-2]<>"""",R[12]C[-3],IF(R[11]C[-2]<>"""",R[11]C[-3],IF(R[10]C[-2]<>"""",R[10]C[-3],IF(R[9]C[-2]<>"""",R[9]C[-3],IF(R[8]C[-2]<>"""",R[8]C[-3],IF(R[7]C[-2]<>"""",R[7]C[-3],IF(R[6]C[-2]<>"""",R[6]C[-3],R[5]C[-3]-MONTH(TODAY())/12))))))))"
   
    Range("A2").NumberFormatLocal = "@"
    [q9] = "": [f15] = 1: [y23] = 0: [y24] = 0: [i14] = "": [y16] = "": [a1] = [a2]: pnd = 1: [w25] = "USD"
 Range("o10:p40").ClearContents
 dc1 = 31 '__³ø
 dc5 = 67 '____ªÑ»ù
 dc6 = 76 '___+__+__+__ªÑ»ù+¤é´Á
 dc7 = 79 '__²v
 dccf = 83 '____¬yªí
 dcmc = 92 '__­È
 t = 10: b = 40
 wj = "Y": gr = "Y"
 [a12] = "Yahoo___________________ Y"¦b¤¤°ê½Ð©ó¥k®æ¥´ Y"

cd$ = mkw([a2])
dc = dc1
Range(Columns(dc), Columns(dc + 100)).Clear






  sUrl(1, 0) = "1 / 9 Quarterly income statement https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials/income/quarter"
  sUrl(1, 1) = "https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials/income/quarter"
  sUrl(1, 2) = dc '_______Alex__ÂAlex§d®á
  sUrl(1, 3) = "Quarterly income statement  1 / 9"
  
  sUrl(2, 0) = "2 / 9 Quarterly balance sheet https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials/balance-sheet/quarter"
  sUrl(2, 1) = "https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials/balance-sheet/quarter"
  sUrl(2, 2) = dc + 9 '____²£ªí
  sUrl(2, 3) = "Quarterly balance sheet  2 / 9"
  
  sUrl(3, 0) = "3 / 9 Annual income statement https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials"
  sUrl(3, 1) = "https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials"
  sUrl(3, 2) = dc + 18 ' ____¯qªí
  sUrl(3, 3) = "Annual income statement  3 / 9"
  
  sUrl(4, 0) = "4 / 9 Annual balance sheet https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials/balance-sheet"
  sUrl(4, 1) = "https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials/balance-sheet"
  sUrl(4, 2) = dc + 27 '____²£ªí
  sUrl(4, 3) = "Annual balance sheet  4 / 9"
  
  'sUrl(5, 0) = "5 / 9 ___+__+__+__ https://www.marketwatch.com/investing/Stock/" & cd$k/" & cd$
  'sUrl(5, 1) = "https://www.marketwatch.com/investing/Stock/" & cd$
  'sUrl(5, 2) = 76 'CC ___+__+__»ù+¤é´Á
  'sUrl(5, 3) = "____+__+__+__  5 / 9"²Ð  5 / 9"
  


'On Error GoTo Errorhandler_IE:
  
'ResetIE:
   
'    If [zz1] = "" Or [zz1] = 1 Or [zz1] Mod 3 = 1 Then
'       Set IE = New InternetExplorer
'    End If





   For x = 1 To 4
        Cells(1, sUrl(x, 2)) = sUrl(x, 0)
        Application.StatusBar = sUrl(x, 3)
        'Call ConnectMarketWatch(sUrl(x, 1), Cells(2, sUrl(x, 2)), x)
        ' Call ConnectIE(sUrl(x, 1), (WT + x / 2), IEVisible)
          Call ConnectXMLHTTP(sUrl(x, 1))
          
            
          '-------------------------------------------------------
          'extract data
          
             ii = 0
             
         
             
             For Each tbl In doc.getElementsByTagName("table")
                      
                      If tbl.className = "table table--overflow align--right" Then
                      
                             endrow = Cells(10000, sUrl(x, 2)).End(xlUp).Row + 1
                             
                             Set rng = Cells(endrow, sUrl(x, 2))
                              
                                For Each rw In tbl.Rows
                                    For Each cl In rw.Cells
                                           
                                            
                                           If InStr(1, cl.className, "overflow__heading") >= 1 Then
                                               rng.value = cl.getElementsByTagName("div")(0).innerText
                                           
                                           ElseIf cl.className = "overflow__cell fixed--column" Then
                                               
                                              rng.value = cl.getElementsByTagName("div")(0).innerText
                                              
                                           Else
                                           
                                              rng.value = Trim(cl.innerText)

                                           End If
                                    
                                    
                                           Set rng = rng.Offset(, 1)
                                           ii = ii + 1 'column
                                    Next cl
                                    Set rng = rng.Offset(1, -ii)
                                    ii = 0
                                Next rw
                                
                          End If
               Next tbl
               
               'get fiscal year
               
               For Each L1 In doc.getElementsByTagName("small")
               
                   If L1.className = "small" And InStr(1, L1.innerText, "All value") >= 1 Then
                          
                         Cells(2, sUrl(x, 2)) = L1.innerText
                       
                         Exit For
                         
                   End If
                   
               Next L1
               
               
             
         '---------------------------------------------------------------
         Set doc = Nothing
        
If Right(Cells(2, dc1 + 1), 4) = "0001" And Right(Cells(2, dc1 + 2), 4) = "0001" Then '--------0001---------------------------
Cells(2, dc1 + 9) = Cells(2, dc1 + 6)
Range("AG2:AG100").Select
Selection.Copy
Range("AK2").Select
ActiveSheet.Paste

Range("AF2:AF100").Select
Selection.Copy
Range("AL2").Select
ActiveSheet.Paste

Range("AH2:AM100").Select
Selection.Copy
Range("AF2").Select
ActiveSheet.Paste
Range("AK2:AL100").Clear
Cells(2, dc1 + 6) = Cells(2, dc1 + 9)
Cells(2, dc1 + 9) = ""

If Abs(Month(Cells(2, dc1 + 2)) - Month(Cells(2, dc1 + 3))) < 5 Then
Cells(2, dc1 + 5) = Cells(2, dc1 + 1) + 365
Cells(2, dc1 + 4) = Cells(2, dc1 + 5) - 92
Else
Cells(2, dc1 + 5) = Cells(2, dc1 + 3) + 365
Cells(2, dc1 + 4) = Cells(2, dc1 + 5) - 182
End If
End If

If Right(Cells(2, dc1 + 1), 4) = "0001" Then '--------0001---------------------------
Cells(2, dc1 + 7) = Cells(2, dc1 + 6)
Range("AF2:AF100").Select
Selection.Copy
Range("AK2").Select
ActiveSheet.Paste
Range("AG2:AL100").Select
Selection.Copy
Range("AF2").Select
ActiveSheet.Paste
Cells(2, dc1 + 7) = ""

If Abs(Month(Cells(2, dc1 + 2)) - Month(Cells(2, dc1 + 3))) < 5 Then
Cells(2, dc1 + 5) = Cells(2, dc1 + 1) + 365
Else
Cells(2, dc1 + 5) = Cells(2, dc1 + 3) + 365
End If
End If
        
         
         [m1] = ""
         If Len(Sheets("__").[t3]) > 5 Thenen
         [m1] = Sheets("__").[t3]3]
         If DateDiff("d", [m1], [aj2]) < 40 Then
         [m1] = "N"
         [b11] = "The new financial report is not yet ready."
         GoTo __¦¡
         End If
         End If
         
         For m = 2 To 20
         If Cells(m, dc1) <> "" Then Exit For
         Next m
         If m > 19 Then GoTo __¦¡
         
   Next x '__Alex__§d®á


For i = 2 To 5
For j = dc + 18 + 6 To dc + 18 Step -1
If Right(Trim(Cells(i, j)), 5) = "trend" Then GoTo yyrr
Next j
Next i

yyrr:
For k = 1 To 6
Cells(93 + k, 22) = ""
Next k

Cells(94, 22) = j - dc - 18
For k = 1 To 5
Cells(94 + k, 22) = Cells(i, dc + 18 + 6 - k)
Next k

For k = 1 To 5
If Right(Trim(Cells(94 + k, 22)), 5) = "trend" Then
Cells(94 + k, 22) = Cells(99, 22) - 1
Exit For
End If
Next k

dc = dc1 + 27
For i = 1 To 10
   If (Left(Cells(i, dc), 1) = "F") Then Exit For
Next i
[x25] = Cells(i, dc)


If Right(Cells(2, dc1 + 10), 4) = "0001" And Right(Cells(2, dc1 + 11), 4) = "0001" Then '-------0001---------------
Cells(2, dc1 + 17) = Cells(2, dc1 + 15)
Range("AP2:AP100").Select
Selection.Copy
Range("AT2").Select
ActiveSheet.Paste

Range("AO2:AO100").Select
Selection.Copy
Range("AU2").Select
ActiveSheet.Paste

Range("AQ2:AV100").Select
Selection.Copy
Range("AO2").Select
ActiveSheet.Paste
Range("AT2:AU100").Clear
Cells(2, dc1 + 15) = Cells(2, dc1 + 17)
Cells(2, dc1 + 17) = ""

If Abs(Month(Cells(2, dc1 + 11)) - Month(Cells(2, dc1 + 12))) < 5 Then
Cells(2, dc1 + 14) = Cells(2, dc1 + 10) + 365
Cells(2, dc1 + 13) = Cells(2, dc1 + 14) - 92
Else
Cells(2, dc1 + 14) = Cells(2, dc1 + 12) + 365
Cells(2, dc1 + 13) = Cells(2, dc1 + 14) - 182
End If
End If

If Right(Cells(2, dc1 + 10), 4) = "0001" Then '-------0001---------------
Cells(2, dc1 + 16) = Cells(2, dc1 + 15)
Range("AO2:AO100").Select
Selection.Copy
Range("AT2").Select
ActiveSheet.Paste
Range("AP2:AV100").Select
Selection.Copy
Range("AO2").Select
ActiveSheet.Paste
Cells(2, dc1 + 16) = ""

If Abs(Month(Cells(2, dc1 + 2)) - Month(Cells(2, dc1 + 3))) < 5 Then
Cells(2, dc1 + 14) = Cells(2, dc1 + 10) + 365
Else
Cells(2, dc1 + 14) = Cells(2, dc1 + 12) + 365
End If
End If


'--------------------------------------------------------------------



If wj = "Y" Then

If InStr(1, [an2], "values") > 0 Then
[w25] = Mid(Cells(i, dc), InStr(1, Cells(i, dc), "values") + 7, 3)
GoTo xwj
End If


dc = 104 'ADR_______WSJ__§ìWSJ¼È¥N
nd$ = "US"
cd$ = wsj([a2])

  Dim url As String
    url = "https://quotes.wsj.com/" & nd$ & "/" & cd$ & "/financials/quarter/income-statement"
    Cells(1, dc) = "6-2 / 9 Company name + Profile + Stock price + Date " & url
    Call ConnectMarketWatch(url, Cells(2, dc), 2)
        
For i = 1 To 10
   If InStr(1, Cells(i, dc), "values") > 0 Then Exit For
Next i
[w25] = Mid(Cells(i, dc), InStr(1, Cells(i, dc), "values") + 7, 3) '___v¥Î

xwj:

End If


'--------------------------------------------------------------------------------------
Application.StatusBar = "Historical stock price  5 / 9"


If UCase([f12]) = "Y" Then GoTo CN1

If UCase([f12]) = "Y" Then

dc = dc5
cd$ = mkw([a2])
    url = "https://seekingalpha.com/symbol/" & cd$ & "/dividends/history"
    Cells(1, dc) = "5 / 9 Historical stock price " + url
      
        Cells(4, dc) = "Date"
        Cells(4, dc + 1) = "Open"
        Cells(4, dc + 2) = "High"
        Cells(4, dc + 3) = "Low"
        Cells(4, dc + 4) = "Close"
        Cells(4, dc + 5) = "Volume"
        Cells(4, dc + 6) = "PercentChange"
    
      
     Dim json As Object
    '---get token---
          url = "https://seekingalpha.com/market_data/xignite_token"
           Call ConnectXMLHTTP(url)
           Set json = ParseJson(doc.body.innerHTML)
           mytoken = json("_token")
           myid = json("_token_userid")
           mydate = Application.Rept("0", 2 - Len(Month(Now()))) & Month(Now()) & "/" & Application.Rept("0", 2 - Len(Day(Now()))) & Day(Now()) & "/" & Year(Now())
     
        
      '---get price---
        ' https://globalhistorical.xignite.com/v3/xGlobalHistorical.json/GetGlobalHistoricalMonthlyQuotesRange?IdentifierType=Symbol&Identifier=aapl&AdjustmentMethod=All&StartDate=01/01/1991&EndDate=09/19/2021&IdentifierAsOfDate=&_callback=SA.Utils.SymbolData.clb16320592966848&_token=fd1f4a4e836e3c67486c178d82228ff7d616c6af732d8d8de92c602095f9e20e6986177b9c96cc13bbac6c5a6ea77d1d48e022f1&_token_userid=122&_=1632059230580
          url = "https://globalhistorical.xignite.com/v3/xGlobalHistorical.json/GetGlobalHistoricalMonthlyQuotesRange?IdentifierType=Symbol&Identifier=" & cd$ & "&AdjustmentMethod=All&StartDate=01/01/1991&EndDate=" & mydate & "&IdentifierAsOfDate=&_callback=SA.Utils.SymbolData.clb16320592966848&_token=" & mytoken & "&_token_userid=" & myid
          
          Call ConnectXMLHTTP(url)
          
          If InStr(1, doc.body.innerText, "HistoricalQuotes") = 0 Or InStr(1, doc.body.innerText, "No data available") >= 1 Then
             Debug.Print "no Historical price"
             GoTo nohistoricalprice
          
          Else
               
               startpos = InStr(1, doc.body.innerText, "HistoricalQuotes")
               totallength = Len(doc.body.innerText) - startpos + 1
               
               myjson = "{" & Mid(doc.body.innerText, startpos - 1, totallength)
               
              ' Debug.Print myjson
               
               Set json = ParseJson(myjson)
               
               pricerow = 5
               For Each L1 In json("HistoricalQuotes")
                    Cells(pricerow, dc) = L1("Date")
                    Cells(pricerow, dc + 1) = L1("Open")
                    Cells(pricerow, dc + 2) = L1("High")
                    Cells(pricerow, dc + 3) = L1("Low")
                    Cells(pricerow, dc + 4) = L1("Close")
                    Cells(pricerow, dc + 5) = L1("Volume")
                    Cells(pricerow, dc + 6) = L1("PercentChange")
                    
                    pricerow = pricerow + 1
               Next L1
               
          
          End If
          
nohistoricalprice:
          
        '--- get split---
            '  https://globalhistorical.xignite.com/v3/xGlobalHistorical.json/GetSplitHistory?IdentifierType=Symbol&Identifier=aapl&AdjustmentMethod=All&StartDate=01/20/2020&EndDate=09/19/2021&IdentifierAsOfDate=&_callback=SA.Utils.SymbolData.clb16320617503056&_token=fd1f4a4e836e3c67486c178d82228ff7d616c6af732d8d8de92c602095f9e20e6986177b9c96cc13bbac6c5a6ea77d1d48e022f1&_token_userid=122&_=1632061454662
               
               url = "https://globalhistorical.xignite.com/v3/xGlobalHistorical.json/GetSplitHistory?IdentifierType=Symbol&Identifier=" & cd$ & "&AdjustmentMethod=All&StartDate=01/01/1991&EndDate=" & mydate & "&IdentifierAsOfDate=&_callback=SA.Utils.SymbolData.clb16320617503056&_token=" & mytoken & "&_token_userid=" & myid
               
             '  Debug.Print url
               
                Call ConnectXMLHTTP(url)
                
                
               If InStr(1, doc.body.innerText, "Splits") = 0 Or InStr(1, doc.body.innerText, "No data available") >= 1 Or InStr(1, doc.body.innerText, "No splits for") >= 1 Then
                         Debug.Print "no split"
                         GoTo nosplit
                      
                      Else
                           
                           startpos = InStr(1, doc.body.innerText, "Splits")
                           totallength = Len(doc.body.innerText) - startpos + 1
                           
                           myjson = "{" & Mid(doc.body.innerText, startpos - 1, totallength)
                           
                          ' Debug.Print myjson
                           
                           Set json = ParseJson(myjson)
                           
                           pricerow = Cells(100000, dc).End(xlUp).Row + 1
                           For Each L1 In json("Splits")
                                Cells(pricerow, dc) = L1("ExDate")
                                Cells(pricerow, dc + 1) = L1("SplitRatio")
                                Cells(pricerow, dc + 2) = "Split"
                                pricerow = pricerow + 1
                           Next L1
                           
                      
                      End If
                        
                             
nosplit:
 
                 
          '---get dividend---
          
            '  https://globalhistorical.xignite.com/v3/xGlobalHistorical.json/GetCashDividendHistory?IdentifierType=Symbol&Identifier=aapl&StartDate=01/20/2020&EndDate=09/19/2021&IdentifierAsOfDate=&CorporateActionsAdjusted=false&_callback=SA.Utils.SymbolData.clb16320617503067&_token=fd1f4a4e836e3c67486c178d82228ff7d616c6af732d8d8de92c602095f9e20e6986177b9c96cc13bbac6c5a6ea77d1d48e022f1&_token_userid=122&_=1632061454663
          
                url = "https://globalhistorical.xignite.com/v3/xGlobalHistorical.json/GetCashDividendHistory?IdentifierType=Symbol&Identifier=" & cd$ & "&StartDate=01/01/1991&EndDate=" & mydate & "&IdentifierAsOfDate=&CorporateActionsAdjusted=false&_callback=SA.Utils.SymbolData.clb16320617503067&_token=" & mytoken & "&_token_userid=" & myid
                ' Debug.Print url
                 
                Call ConnectXMLHTTP(url)
                
               If InStr(1, doc.body.innerText, "CashDividends") = 0 Or InStr(1, doc.body.innerText, "No data available") >= 1 Or InStr(1, doc.body.innerText, "No dividends") >= 1 Then
                         Debug.Print "no dividend"
                         GoTo nodividend
                      
                      Else
                           
                           startpos = InStr(1, doc.body.innerText, "CashDividends")
                           totallength = Len(doc.body.innerText) - startpos + 1
                           
                           myjson = "{" & Mid(doc.body.innerText, startpos - 1, totallength)
                           
                          ' Debug.Print myjson
                           
                           Set json = ParseJson(myjson)
                           
                           
                           pricerow = Cells(100000, dc).End(xlUp).Row + 1
                           For Each L1 In json("CashDividends")
                                Cells(pricerow, dc) = L1("ExDate")
                                Cells(pricerow, dc + 1) = L1("DividendAmount")
                                Cells(pricerow, dc + 2) = "Dividend"
                                
                                
                                pricerow = pricerow + 1
                           Next L1
                           
                      
                      End If
                        
                        
nodividend:
                   
                   Set json = Nothing
              

               '----clear and sort data-----------------------------------------------------------------------
              
                  endrow = Cells(100000, dc).End(xlUp).Row
                  
                  Range("BO4:BU" & endrow).Select
                  ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
                  ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("BO5:bo" & endrow), _
                      SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                  With ActiveWorkbook.ActiveSheet.Sort
                      .SetRange Range("BO4:BU" & endrow)
                      .Header = xlYes
                      .MatchCase = False
                      .Orientation = xlTopToBottom
                      .SortMethod = xlPinYin
                      .Apply
                  End With
                 
              '----end clear and sort data--------------------------------------------------------------------------
            
                                   
End If


If UCase([f12]) = "N" Then
dc = dc5
yrs$ = [n10] - 30
yre$ = [n10]
cd$ = yho([a2])

'Call ushis '__20_______yahoo____yahoo¤w¤£µ¹§ì
'Call ConnectXMLHTTP(url)

'        For Each tbl In doc.getElementsByTagName("table")
'            If tbl.className = "W(100%) M(0)" Then
'                ii = 0
'
'                              Set rng = Cells(3, dc)
'                                 For Each rw In tbl.Rows
'                                     For Each cl In rw.Cells
'                                            rng.Value = Trim(cl.innerText)
'                                            Set rng = rng.Offset(, 1)
'                                         ii = ii + 1 'column
'                                     Next cl
'                                     Set rng = rng.Offset(1, -ii)
'                                     ii = 0
'                                 Next rw
'
'            End If
'        Next tbl
'        Set doc = Nothing
     
  
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
url = "https://finance.yahoo.com/quote/" & cd$ & "/history?period1=573436800&period2=" & myday & "&interval=1mo&filter=history&frequency=1mo"
Cells(1, dc) = "5 / 9 Historical stock price " + url
Call Yahoo_Price_Dividend_Split(cd$, dc)

End If



CN1:


If wj = "Y" Then
    Application.StatusBar = "Company name + Profile + Stock price + Date  6 / 9"
    cd$ = wsj([a2])
    nd$ = "US"
    dc = dc6
   
      Cells(1, dc) = "6-1 / 9 Company name + Profile + Stock price + Date https://www.marketwatch.com/investing/stock/" & cd$ & "/company-profile"
      
      url = "https://www.marketwatch.com/investing/stock/" & cd$ & "/company-profile"
      

      Debug.Print url
      
      
      Call ConnectXMLHTTP(url)
     
      
              'companyname
              For Each L1 In doc.getElementsByTagName("h1")
                    If L1.className = "company__name" Then
                        companyname = L1.innerText
                        Exit For
                    
                    End If
          
              Next L1
              
              'price
              For Each L1 In doc.getElementsByTagName("h2")
                  If L1.className = "intraday__price " Then
                      stockprice = Replace(L1.innerText, "$", "")
                      
                      Exit For
                  End If
              Next L1
          
             'tickername, exchange
             For Each L0 In doc.getElementsByTagName("div")
                 If L0.className = "company__symbol" Then
                    For Each L1 In L0.getElementsByTagName("span")
                        If L1.className = "company__ticker" Then
                            tickername = L1.innerText
                              
                        ElseIf L1.className = "company__market" Then
                            exchange = L1.innerText
                            
                        End If
                        
                    Next
                    
                    
                    Exit For
                    
                 End If
             
             Next L0
          
                'date
                For Each L1 In doc.getElementsByTagName("span")
                
                       
                    If L1.className = "timestamp__time" Then
                        closedate = Replace(L1.innerText, "Last Updated:", "")
                        a = InStr(1, closedate, ",")
                        If a >= 1 Then
                           closedate = Trim(Left(closedate, a + 6))
                        
                        End If
                        
                        
                         
                        Exit For
                       
                    End If
                
                Next L1
                
               'industry, sector
               For Each L0 In doc.getElementsByTagName("li")
                                   
                   If L0.className = "kv__item w100" Then
                     
                          For Each L1 In L0.getElementsByTagName("small")
                                   
                                   If L1.className = "label" And L1.innerText = "Industry" Then
                                      
                                      industry = L0.getElementsByTagName("span")(0).innerText
                                    
                                   ElseIf L1.className = "label" And L1.innerText = "Sector" Then
                                      sector = L0.getElementsByTagName("span")(0).innerText
                                   
                                   End If
                                   
                              
                          Next L1
                
                   
                     End If
                  
               
               Next L0
                                    
                'profile
                For Each L1 In doc.getElementsByTagName("p")
                    If L1.className = "description__text" Then
                        myprofile = L1.innerText
                        Exit For
                    End If
               
                Next L1
         
                    
                [a1] = companyname & " " & tickername & " (" & exchange & ")  : " & sector & "*" & industry
                [i1] = closedate
                [q9] = stockprice
                [y16] = myprofile
                
        
   
   

End If

If wj = "N" Then
Application.StatusBar = "Company name + Profile + Stock price + Date  6-1 / 9"
cd$ = mkw([a2])
nd$ = "US"
dc = dc6
    Cells(1, dc) = "6-1 / 9 Company name + Profile + Stock price + Date https://www.marketwatch.com/investing/stock/" & cd$ & "/company-profile?mod=mw_quote_tab"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.marketwatch.com/investing/stock/" & cd$ & "/company-profile?mod=mw_quote_tab", Destination:=Cells(2, dc))
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
    End With 'by ___sºÖ
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    
   i = 2: Do Until Left(Cells(i, dc), 11) = "Watchlist C" Or i > 1000
    i = i + 1
   Loop
   If i < 999 Then
      [a1] = Cells(i - 1, dc) '___q¦W
      Cells(i - 1, dc - 1) = 1
      
      aa = InStr(1, Cells(i - 2, dc), ":")
      [y25] = Right(Cells(i - 2, dc), Len(Cells(i - 2, dc)) - aa - 1)
      [a1] = [a1] & " (" & [y25] & ")"
      Cells(i - 2, dc - 1) = 1
      
      aa = InStr(1, Cells(i + 4, dc), "dated")
      bb = InStr(1, Cells(i + 4, dc), ", 20")
      [i1] = Mid(Cells(i + 4, dc), aa + 7, bb - aa + 7) '__´Á
      Cells(i + 4, dc - 1) = 1
   
      [q9] = Cells(i + 5, dc) '__»ù
      Cells(i + 5, dc - 1) = 1
   End If

   Do Until Left(Cells(i, dc), 6) = "Sector" Or i > 1000
   i = i + 1
   Loop
   If i > 999 Then GoTo des
   [a1] = [a1] & "_" & Trim(Cells(i - 1, dc)))
   Cells(i, dc - 1) = 1
   
   [a1] = [a1] & "*" & Trim(Cells(i, dc))
   Cells(i - 1, dc - 1) = 1

   aa = InStr(1, [a1], "Sector")
   If aa > 0 Then [a1] = Left([a1], aa - 1) + Right([a1], Len([a1]) - aa - Len("Sector ") + 1)
   aa = InStr(1, [a1], "Industry")
   If aa > 0 Then [a1] = Left([a1], aa - 1) + Right([a1], Len([a1]) - aa - Len("Industry ") + 1)
   aa = InStr(1, [a1], "Companies on the")
   If aa > 0 Then [a1] = Left([a1], aa - 1) + Right([a1], Len([a1]) - aa - Len("Companies on the ") + 1)
   
des:

If InStr(1, [an2], "values") > 0 Then
[w25] = Mid(Cells(i, dc), InStr(1, [an2], "values") + 7, 3)
GoTo xbr
End If

Application.StatusBar = "Barron's All Values are in XXX Millions 6-2 / 9"
dc = 104
cd$ = mkw([a2])
If [y25] = "NYSE" Then nd$ = "xnys"
If [y25] = "Nasdaq" Then nd$ = "xnas"
If [y25] = "OTC" Then nd$ = "ootc"

    url = "https://www.barrons.com/quote/stock/us/" & nd$ & "/" & cd$ & "/financials"
    Cells(1, dc) = "6-2 / 9 All Values are in XXX Millions " & url
        
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.barrons.com/quote/stock/us/" & nd$ & "/" & cd$ & "/financials", Destination:=Cells(2, dc))
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
    End With 'by ___sºÖ
                      
For i = 1 To 500
If Left(Cells(i, dc), 3) = "Fis" Then
Cells(i, dc - 1) = 1
[w25] = Mid(Cells(i, dc), InStr(1, Cells(i, dc), "are in ") + 7, 3) '___v¥Î
Exit For
End If
Next i

xbr:


dc = 113
Application.StatusBar = "Barron's Company Description 6-3 / 9"
    url = "https://www.barrons.com/quote/stock/us/" & nd$ & "/" & cd$ & "/company-people"
    Cells(1, dc) = "6-3 / 9 Company Description " & url
        
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.barrons.com/quote/stock/us/" & nd$ & "/" & cd$ & "/company-people", Destination:=Cells(2, dc))
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
    End With 'by ___sºÖ
                     
i = 1
Do Until Trim(Cells(i, dc)) = "Company Description" Or i > 1000
    i = i + 1
Loop
   If Cells(i + 2, dc) <> "" Then
   [y16] = Cells(i + 2, dc)
   Cells(i + 2, dc - 1) = 1
   End If

End If
   
   
   


   
   If [a2] = "TLK" Then [w25] = "IDR" '____¹q«H
    If [w25] = "USD" Then GoTo cf
    Application.StatusBar = "Foreign Exchenge 7 / 9"
    dc = dc7
    
   
    
    url = "https://wise.com/zh-hk/currency-converter/usd-to-" & [w25] & "-rate" '_____yahooyahoo
    
    Cells(1, dc) = "7 / 9 Foreign exchange " & url
    

     
         Call ConnectWinHttp(url, 1)
         
                            
        
                            
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
Application.StatusBar = "Annual cashflow statement  8 / 9"
dc = dccf
cd$ = mkw([a2])
    url = "https://www.marketwatch.com/investing/Stock/" & cd$ & "/financials/cash-flow"
    Cells(1, dc) = "8 / 9 Annual cashflow statement " + url
    'Call ConnectMarketWatch(url, Cells(2, dc ), 2)
     'Call ConnectIE(url, 2, IEVisible)
     Call ConnectXMLHTTP(url)
     
          '-------------------------------------------------------
          'extract data
             ii = 0
         
             For Each tbl In doc.getElementsByTagName("table")
                  ' Debug.Print tbl.className
                   
                       If tbl.className = "table table--overflow align--right" Then
                             
                             endrow = Cells(10000, dc).End(xlUp).Row + 1
                             
                             Set rng = Range("ce" & endrow)
                             
                                For Each rw In tbl.Rows
                                    For Each cl In rw.Cells
                                           If InStr(1, cl.className, "overflow__heading") >= 1 Then
                                               rng.value = cl.getElementsByTagName("div")(0).innerText
                                           
                                           ElseIf cl.className = "overflow__cell fixed--column" Then
                                               
                                              rng.value = cl.getElementsByTagName("div")(0).innerText
                                              
                                           Else
                                           
                                              rng.value = Trim(cl.innerText)

                                           End If
                                    
                                    
                                        Set rng = rng.Offset(, 1)
                                        ii = ii + 1 'column
                                    Next cl
                                    Set rng = rng.Offset(1, -ii)
                                    ii = 0
                                Next rw
                     End If
               Next tbl
               
               'get fiscal year
               
               For Each L1 In doc.getElementsByTagName("small")
               
                   If L1.className = "small" And InStr(1, L1.innerText, "All value") >= 1 Then
                          
                         Range("ce2") = L1.innerText
                       
                         Exit For
                         
                   End If
                   
               Next L1
               
               
               
         '---------------------------------------------------------------
         Set doc = Nothing
        
    
         'If [zz1] = "" Or [zz1] Mod 3 = 0 Then ' "" means____, mod 3 from __m ¦¬ÂÃ
         '       IE.Quit
         '       Set IE = Nothing
         '       Call DelIE
         'End If
             
   '-----------------------------------------


For i = 2 To 5
For j = dc + 8 To dc Step -1
If Right(Trim(Cells(i, j)), 5) = "trend" Then GoTo cyyrr
Next j
Next i

cyyrr:
For k = 1 To 6
Cells(101 + k, 22) = ""
Next k

For k = 1 To 5
Cells(101 + k, 22) = Cells(i, dc + 6 - k)
Next k

For k = 1 To 5
If Right(Trim(Cells(101 + k, 22)), 5) = "trend" Then
Cells(101 + k, 22) = Cells(106, 22) - 1
Exit For
End If
Next k

dc = dcmc
Application.StatusBar = "Market capitalization  9 / 9"

If gr = "Y" Then
cd$ = mkw([a2])
        url = "https://www.gurufocus.com/stock/" & cd$ & "/guru-trades" ' RDS.B ____¿ù»~
        Cells(1, dc) = "9 / 9 Market capitalization " & url
        Call ConnectMarketWatch(url, Cells(2, dc), 2)
        
        i = 2
        Do Until UCase(Left(Trim(Cells(i, dc)), 10)) = "MARKET CAP" Or i > 900
        i = i + 1
        Loop
        [y23] = Cells(i, dc + 1)
        Cells(i, dc - 1) = 1
        
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
    End With 'by ___sºÖ
        
    i = 2
    Do Until UCase(Left(Trim(Cells(i, dc)), 10)) = "MARKET CAP" Or i > 900
    i = i + 1
    Loop
    unitt = Cells(i, dc + 1)
    [y23] = unitconversion(unitt)
    Cells(i, dc - 1) = 1
    
 End If
 
                         
err:
[b11] = ""

'----------------------------------------------------------------------------------------------

If [f12] = "Y" Then GoTo CN2

dc = dc5
For c = 2 To 500
If Cells(c, dc) = "Date" Or Cells(c, dc) = "__" Then Exit Foror
Next c
If c > 499 Then GoTo bf2
c = c + 1

Call highlow(dc, c, t, b, pnd)


CN2:

If [q9] = "" Or [q9] = 0 Or Not IsNumeric([q9]) Then
For i = 1 To 100
If Cells(i, dc + 4) <> "" And IsNumeric(Cells(i, dc + 4)) Then Exit For
Next i
[q9] = Cells(i, dc + 4) * pnd
End If

    If [q9] < [o10] Then [o10] = [q9]
    If [q9] > [p10] Then [p10] = [q9]
    If [o10] = "" Then [o10] = [q9]
    If [p10] = "" Then [p10] = [q9]
    

'--------------------------------------------------------------------------------------------------

bf2: If Cells(3, dc1) = "" Or Cells(3, dc1 + 9) = "" Or Cells(3, dc1 + 9 * 2) = "" Or Cells(3, dc1 + 9 * 3) = "" Then GoTo __¦¡


[aa22] = "Net Property, Plant & Equipment"
i = 2: Do Until i > 100
If Cells(i, dc1 + 9) = "Net Property, Plant & Equipments" Then
[aa22] = "Net Property, Plant & Equipments"
Exit Do
End If
i = i + 1
Loop

[aa13] = [ae3]
For dc = dc1 To dc1 + 9 * 3 Step 9 ' 31 40 49 58 ________³æ¦ìÂà´«
i = 2: Do Until Cells(i, dc) = "" And Cells(i + 1, dc) = "" And Cells(i + 2, dc) = ""
j = 13: Do Until Cells(j, 27) = ""
If Cells(i, dc) = Cells(j, 27) Then
For m = 1 To 5
If IsError(Cells(i, dc + m)) Then Cells(i, dc + m) = 0
unitt = Cells(i, dc + m)
Cells(i, dc + m) = unitconversion(unitt)
Next m
Range(Cells(i, dc + 1), Cells(i, dc + 5)).NumberFormatLocal = "#,##0_);(#,##0)"
End If
j = j + 1
Loop
i = i + 1
Loop
Next dc


dc = dccf ' _________r³æ¦ìÂà´«
i = 2: Do Until Cells(i, dc) = "" And Cells(i + 1, dc) = "" And Cells(i + 2, dc) = ""
j = 37: Do Until Cells(j, 27) = ""
If Cells(i, dc) = Cells(j, 27) Then
For m = 1 To 5
unitt = Cells(i, dc + m)
Cells(i, dc + m) = unitconversion(unitt)
Next m
End If
j = j + 1
Loop
i = i + 1
Loop

'-------------------------------------------------------------------------

__: [y24] = [q9]9]

    [h13].Select: cd$ = yho([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finance.yahoo.com/quote/" & cd$ & "/?p=" & cd$, _
    TextToDisplay:="Chart"
                   
    If [y25] = "NYSE" Then nd$ = "xnys"
    If [y25] = "Nasdaq" Then nd$ = "xnas"
    If [y25] = "OTC" Then nd$ = "ootc"

    [g13].Select: cd$ = mkw([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://www.morningstar.com/stocks/" & nd$ & "/" & cd$ & "/news", _
    TextToDisplay:="News"
    
    If [y25] = "NYSE" Then nd$ = "NYSE"
    If [y25] = "Nasdaq" Then nd$ = "NASDAQ"
    If [y25] = "OTC" Then nd$ = "OTCMKTS"
    
    [f13].Select: cd$ = mkw([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://wealth.esunbank.com.tw/usstock/esun/basic0001.xdjhtm?a=" & cd, _
    TextToDisplay:="Prfl-e"
    'Address:="https://www.moneydj.com/us/basic/basic0001/" & cd, _

    [e13].Select: cd$ = mkw([a2])
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://abxusa.com/" & cd$, _
    TextToDisplay:="Profile"
    'Address:="https://www.mg21.com/" & cd$ & ".html", _

    [d13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://finviz.com/screener.ashx?v=111&f=fa_pe_u15,fa_roe_o10,idx_sp500&ft=4", _
    TextToDisplay:="Screen"
    
    [c13].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://tradingeconomics.com/united-states/gdp-growth-annual/forecast", _
    TextToDisplay:="GDP"
    
    [h15].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="http://stocks.ddns.net/US.aspx", _
    TextToDisplay:="______- __Alex____"lex§d®á´£¨Ñ"
    
    [a70].Select
    Selection.Hyperlinks.Add Anchor:=Selection, _
    Address:="https://us.spindices.com/documents/additional-material/withholding-tax-index-values.pdf", _
    TextToDisplay:="ADR Div Tax"
   

If [b11] = "___the new financial report is not yet ready" Then GoTo y16y16

For m = 2 To 20
If Cells(m, dcmc) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___9 / 9 Latest market capitalization is missing."g."

For m = 2 To 20
If Cells(m, dccf) <> "" Then Exit For
Next m
If m > 19 Then
'[v42] = "___"¬y"
If [w25] <> "USD" Then [b11] = "___8 / 9 Cashflow statements are missing."g."
End If

If [q9] = 0 Or [q9] = "" Then [b11] = "___6 / 9 Stock prices are missing."g."


If UCase([f12]) <> "Y" Then
For m = 5 To 20
If Cells(m, dc5) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___5 / 9 Div, split & historic prices are missing."g."
End If

For m = 2 To 20
If Cells(m, dc1 + 9 * 3) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___4 / 9 Annual balance sheets are missing."g."

For m = 2 To 20
If Cells(m, dc1 + 9 * 2) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___3 / 9 Annual income statements are missing."g."

For m = 2 To 20
If Cells(m, dc1 + 9) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___2 / 9 Quarterly balance sheets are missing."g."

For m = 2 To 20
If Cells(m, dc1) <> "" Then Exit For
Next m
If m > 19 Then [b11] = "___1 / 9 Quarterly income statements are missing."g."

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
    
    Range(Cells(2, 2), Cells(54, 10)).ShrinkToFit = True '___Y¤p
    [f9].ShrinkToFit = False '____ÁY¤p
    [i12].ShrinkToFit = False '____ÁY¤p
    [e15].ShrinkToFit = False '____ÁY¤p
    [h15].ShrinkToFit = False '____ÁY¤p
    [g44].ShrinkToFit = False '____ÁY¤p
    [g45].ShrinkToFit = False '____ÁY¤p
    Range(Cells(18, 10), Cells(25, 10)).ShrinkToFit = False '____ÁY¤p
    Range(Cells(37, 10), Cells(47, 10)).ShrinkToFit = False '____ÁY¤p

   Range(Columns(dc1), Columns(dc1 + 100)).Select '___Y¤p
        With Selection
        .WrapText = False
        .ShrinkToFit = True
    End With
    
    Range(Cells(1, dc1), Cells(1, dc1 + 100)).ShrinkToFit = False '____ÁY¤p
    
    [j1].ColumnWidth = 8
    Cells(1, dc1).ColumnWidth = 20
    Cells(1, dc1 + 9).ColumnWidth = 20
    Cells(1, dc1 + 9 * 2).ColumnWidth = 20
    Cells(1, dc1 + 9 * 3).ColumnWidth = 20
    Cells(1, dc5).ColumnWidth = 8
    Cells(1, dccf).ColumnWidth = 30
    
    Range("k1,c13:i13, h16").Select
    With Selection.Font
        .Size = 12
        .Name = "Arial"
    End With
    
    Columns(dc5).NumberFormatLocal = "yyyy/m/d;@"
    Range("c3:e9, b20:b24, e20:j24, b29:b34, e29:j34, j35:j35,b39:f43, h39:i43, b48:f54, h48:i54").Select
    Selection.NumberFormatLocal = "#,##0_);(#,##0)"
    
    Range("e15").HorizontalAlignment = xlRight
    Range("f15").HorizontalAlignment = xlLeft
y16: Range("B11").ShrinkToFit = False
    [y16].Select
    ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    Application.StatusBar = "Done !"
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Application.Calculation = xlAutomatic           '_____}¼sºÖ
    
Exit Sub


'------------------------------------------------------------

Sub_ExtractData:
 
     
    endrow = Range("bx10000").End(xlUp).Row
    pricenow = "": priceclose = ""
    For irow = 2 To endrow
        
        If Range("bx" & irow) = "|" Then
           
            mymarket = Range("bx" & irow + 1)
            myname = Range("bx" & irow + 2)
            
        ElseIf InStr(1, Range("bx" & irow), "Last Updated:") >= 1 Then
            mydate = Range("bx" & irow)
            
            If Application.IsNumber(Range("bx" & irow + 1)) Then
                pricenow = Range("bx" & irow + 1)
            End If
        
        ElseIf InStr(1, Range("bx" & irow), "Close:") >= 1 Then
            
               priceclose = Range("bx" & irow)
               priceclose = Replace(priceclose, "Close:", "")
               
               
        ElseIf InStr(1, Range("bx" & irow), "Market Cap") >= 1 Then
           
            mycap = Range("bx" & irow)
            mycap = Replace(mycap, "Market Cap ", "")
            
        
        End If
        
    
    Next irow
    
    'Debug.Print "pricenow", pricenow, "priceclose", priceclose
    
    
    If pricenow <> "" Then
       myprice = pricenow
    ElseIf pricenow = "" And priceclose <> "" Then
       myprice = priceclose
    
    End If
     
     
     
    'below are xmlhttp method
    ' __³õ
      
    '  For Each l1 In doc.getElementsByTagName("div")
    '
    '      If l1.className = "company__symbol" Then
    '         ' Debug.Print L1.innerText
    '          mymarket = l1.innerText
    '
    '          Exit For
    '
    '      End If
      
     ' Next l1

     'companyname
     
     'For Each l1 In doc.getElementsByTagName("h1")
     '
     '    If l1.className = "company__name" Then
     '       myname = l1.innerText
     '       Exit For
     '    End If
     '
     '
     'Next l1
     
     'price-close
     'For Each l1 In doc.getElementsByTagName("h3")
     '
     '    If l1.className = "intraday__price " Then
     '
     '        myprice = l1.innerText
     '
     '        Exit For
     '    End If
    '
    ' Next l1
    '
    ' If myprice = "" Then '___price now, __price closedlosed
    '
    '    For Each l1 In doc.getElementsByTagName("table")
    '
    '        If l1.className = "table table--primary align--right" Then
    '
    '           For Each l2 In l1.getElementsByTagName("td")
    '               If l2.className = "table__cell u-semi" Then
    '                    myprice = l2.innerText
    '                  Exit For
    '               End If
    '
    '           Next l2
    '
    '         End If
    '    Next l1
    ' End If
     'date
    ' For Each l1 In doc.getElementsByTagName("span")
    '
    '     If l1.className = "timestamp__time" Then
    '
    '         mydate = l1.innerText
    '         Exit For
    '     End If
    '
    ' Next l1
    '
    ' '__­È
    ' For Each l1 In doc.getElementsByTagName("li")
    '      If l1.className = "kv__item" Then
    '           For Each l2 In l1.getElementsByTagName("small")
    '
    '               If l2.className = "kv__label" And l2.innerText = "Market Cap" Then
    '                    mycap = l1.getElementsByTagName("span")(0).innerText
    '                    Exit For
    '                   ' Debug.Print "market cap", mycap
    '                End If
    '           Next l2
    '
    '      End If
    '
    ' Next l1
    '
     
     If mymarket = "" Then mymarket = "Not Available"
     If myname = "" Then myname = "Not Available"
   '  If myprice = "" Then myprice = "Not Available"
     If mydate = "" Then mydate = "Not Available"
     
    

Return

'--------------------------------------------


'Errorhandler_IE:

'    If InStr(1, err.Description, "Automation") >= 1 Then
'        Debug.Print err.Number, err.Description
'        Call DelIE
        
'        GoTo ResetIE
'
'    Else
'
'       Debug.Print "not IE automation issue", err.Number, err.Description
'
'       Resume Next
'    End If
    

    
    
End Sub



Public Function CCRUMBT(rng As Range) '______®aÑÔ®á
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
    
  '  ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
  '  Delete_Pictures '__Alex_x§d
    
End Function

Public Function DCSV(url As String) '______®aÑÔ®á
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
    
   
    
End Function


Sub ConnectMarketWatch(ByVal url As String, dtrange As Range, ByVal UrlNbr)
     On Error Resume Next
    
    Application.EnableEvents = False
   
    With ActiveSheet.QueryTables.Add(Connection:="URL;" & url, Destination:=dtrange)
        .RefreshStyle = xlOverwriteCells
        .WebFormatting = xlWebFormattingNone
        If UrlNbr = 5 Then
            .WebSelectionType = xlEntirePage
            
        ElseIf Application.IsNumber(UrlNbr) = False Then
            .WebSelectionType = xlSpecifiedTables
            .WebTables = UrlNbr
        
        Else
             .WebSelectionType = xlAllTables
        End If
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    Application.EnableEvents = True
    
    err.Clear
    
   
End Sub


Public Sub ConnectXMLHTTP(ByVal myurl As String)
     
     On Error GoTo errorhandler
    
     ConnectErr = False
    
      Set doc = New HTMLDocument
      With CreateObject("MSXML2.XMLHTTP")
              
            .Open "GET", myurl, False
            .setRequestHeader "Content-Type", "text/plain"
            .setRequestHeader "Cache-Control", "no-cache" '__dicnasas
            .setRequestHeader "Pragma", "no-cache" '__dicnasas
            .setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT" '__dicnasas
            .send
            
            Application.Wait Now() + TimeValue("00:00:01") * 4
            doc.body.innerHTML = .responseText
        
            .abort
      End With
     
 Exit Sub
'---------
errorhandler:
    
    Debug.Print "Fail to connect Website", err.Number, err.Description, myurl
    ConnectErr = True
 
    err.Clear
    
    Resume Next
     
End Sub


Sub ConnectIE(ByVal myurl As String, waittime As Integer, myshow As Boolean)
         
         On Error GoTo IE_Errhandler:
         
         ie_starttime = Now()
         Set doc = New HTMLDocument
         With IE
                          
                    .Visible = myshow
             
                    .navigate myurl
                    
                     Do While IE.Busy Or IE.readyState <> READYSTATE_COMPLETE
           
                         DoEvents
                     Loop
                    
                     'Debug.Print IE.readyState, "IEStartTime:", ie_starttime, "IEReadyTime:", Now(), "waitTime:", waittime
                     Application.Wait Now() + TimeValue("00:00:01") * waittime
        
                     Set doc = IE.document
                        
          End With
    
Exit Sub
'------------------------------
IE_Errhandler:

        Debug.Print err.Number, err.Description
        
        err.Clear
        Resume Next

End Sub


Sub DelIE()

   ' Debug.Print "Start to Del IE"
    Shell "taskkill.exe /F /IM iexplore.exe /T", vbHide
    Application.StatusBar = "__Internet Explorer_...".."
    Application.Wait Now() + TimeValue("00:00:01")

End Sub

Private Sub ushis() '__20_______yahoo____yahoo¤w¤£µ¹§ì

Dim myurl(2)
myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=history&crumb="
myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=div&crumb="
myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=split&crumb="

Dim url As String, crumbrng As Range, crumbrng1 As Range
Dim serr As String, csvt As Variant, csv As Variant
Dim cii As Long, cjj As Long, ci As Long, cj As Long

On Error Resume Next
For iurl = 0 To 2
    url = myurl(iurl)
    If iurl = 0 Then 'price
        Cells(1, dc) = "5 / 9 ____ " & url url
        endrow = 3
    Else
        endrow = Range("bo100000").End(xlUp).Row
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
    endrow = Range("bo10000").End(xlUp).Row
    For i = endrow To 7 Step -1
        If Range("bo" & i) = "Date" Or Len(Range("bo" & i)) = 0 Then
             Range("BO" & i & ":BU" & i).Delete Shift:=xlUp
      
        End If
    Next i
    
    endrow = Range("bo10000").End(xlUp).Row
    
    Range("BO4:BU" & endrow).Select
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("bo5:bo" & endrow), _ _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("BO4:BU" & endrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
'----end clear and sort data--------------------------------------------------------------------------


End Sub


Public Sub highlow(dc, c, t, b, pnd)

For i = t To b
Cells(i, 20) = 1
Next i

i = c: D = t
Do Until Cells(i, dc) = "" Or i > 1000
If Cells(i, dc + 2) = "Dividend" Then
Cells(i, dc - 1) = Year(Cells(i, dc))
End If

If Cells(i, dc + 2) = "Split" Or Left(Cells(i, dc + 1), 1) = "_" Thenn
sp = Cells(i, dc + 1)
aa = InStr(1, sp, ":")
If aa > 0 Then Cells(i, dc + 1) = splitt(sp)
j = D
Do Until j > b
If Year(Cells(i, dc)) = Cells(j, 14) Then
Cells(j, 20) = Cells(j, 20) * Cells(i, dc + 1)
If Cells(c + 3, dc) > Cells(c + 5, dc) Then D = j
Exit Do
End If
j = j + 1
Loop
End If
i = i + 1
Loop

i = t: j = c: Cells(t, 21) = Cells(t, 20)
Do Until i > b
Cells(i, 15) = 999999999: Cells(i, 16) = 0
If Cells(c + 3, dc) < Cells(c + 5, dc) Then j = c
Do Until Cells(j, dc) = "" Or j > 1000
If Cells(j, dc + 4) = "" Then GoTo j15
If Cells(c + 3, dc) > Cells(c + 5, dc) Then
If Year(Cells(j, dc)) < Cells(i, 14) Then Exit Do
End If
If Cells(j, dc + 3) <> "" And Not (IsNumeric(Cells(j, dc + 3))) Then GoTo j15 'Isnumeric()=true, isnumber()=false
If Cells(j, dc + 2) <> "" And Not (IsNumeric(Cells(j, dc + 2))) Then GoTo j15
If Year(Cells(j, dc)) = Cells(i, 14) And Cells(j, dc + 3) <> "" And Cells(j, dc + 3) * pnd < Cells(i, 15) Then Cells(i, 15) = Cells(j, dc + 3) * pnd
If Year(Cells(j, dc)) = Cells(i, 14) And Cells(j, dc + 2) <> "" And Cells(j, dc + 2) * pnd > Cells(i, 16) Then Cells(i, 16) = Cells(j, dc + 2) * pnd
j15: j = j + 1
Loop

If Cells(i, 15) = 999999999 Then Cells(i, 15) = ""
If Cells(i, 16) = 0 Then Cells(i, 16) = ""
If i > t Then Cells(i, 21) = Cells(i - 1, 21) * Cells(i, 20) '____¤À³Î
If Cells(i, 15) <> "" Then Cells(i, 15) = Cells(i, 15) * Cells(i, 21)
If Cells(i, 16) <> "" Then Cells(i, 16) = Cells(i, 16) * Cells(i, 21)
i = i + 1
Loop

End Sub

Public Function unitconversion(unitt)

If unitt = "" Or unitt = "-" Then Exit Function

unitconversion = Mid(unitt, 1, Len(unitt) - 1)
If Left(unitt, 1) = "(" Then
unitconversion = -Mid(unitt, 2, Len(unitt) - 3)
If IsNumeric(Left(Right(unitt, 2), 1)) Then unitconversion = -Mid(unitt, 2, Len(unitt) - 2)
End If

If IsNumeric(unitconversion) Then
If unitt <> "" And IsNumeric(unitt) Then unitconversion = unitconversion * 1E-06
If Right(unitt, 1) = "K" Or Right(unitt, 2) = "K)" Then unitconversion = unitconversion * 0.001
If Right(unitt, 1) = "B" Or Right(unitt, 2) = "B)" Then unitconversion = unitconversion * 1000
If Right(unitt, 1) = "T" Or Right(unitt, 2) = "T)" Then unitconversion = unitconversion * 1000000
End If

End Function


Public Function splitt(sp)

splitt = 1
aa = InStr(1, sp, ":")
bb = InStr(1, sp, " ")
cc = InStr(1, sp, "_"))
dd = InStr(1, sp, ">")

If aa > 0 And bb > 0 And cc > 0 And dd > 0 Then
splitt = Mid(sp, aa + 1, cc - aa - 1) / Mid(sp, dd + 2, Len(sp) - 1 - dd - 1)
Else:
If aa > 0 And bb > 0 Then
splitt = Left(sp, aa - 1) / Mid(sp, aa + 1, bb - aa)
Else
splitt = 1
End If
End If

If bb = 0 And cc = 0 And dd = 0 Then splitt = Left(sp, aa - 1) / Right(sp, Len(sp) - aa)

End Function

Public Function mkw(cd$) As String
'marketwatch, monydj, quote123, google, morningstar, guru, seeking-alpha

mkw = cd$
If InStr(1, cd$, "-") > 0 Then mkw = Replace(cd$, "-", ".")
If InStr(1, cd$, "/") > 0 Then mkw = Replace(cd$, "/", ".")

End Function

Public Function yho(cd$) As String
'yahoo

yho = cd$
If InStr(1, cd$, ".") > 0 Then yho = Replace(cd$, ".", "-")
If InStr(1, cd$, "/") > 0 Then yho = Replace(cd$, "/", "-")

End Function

Public Function wsj(cd$) As String
'wsj

wsj = cd$
If InStr(1, cd$, ".") > 0 Then wsj = Replace(cd$, ".", "")
If InStr(1, cd$, "-") > 0 Then wsj = Replace(cd$, "-", "")
If InStr(1, cd$, "/") > 0 Then wsj = Replace(cd$, "/", "")

End Function



Sub Yahoo_Price_Dividend_Split_20241111(cd$, dc)
            ActiveSheet.Range("BN3:bu100000").ClearContents
            Dim myurl(2)
            myday = DateDiff("s", "1/1/1970", Now()) ' unix timestamp
            
            myurl(0) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=history&includeAdjustedClose=true"
            myurl(1) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=div&includeAdjustedClose=true"
            myurl(2) = "https://query1.finance.yahoo.com/v7/finance/download/" & cd$ & "?period1=57600&period2=" & myday & "&interval=1mo&events=split&includeAdjustedClose=true"
            
            
            
            Dim url As String, crumbrng As Range, crumbrng1 As Range
            Dim serr As String, csvt As Variant, csv As Variant
            Dim cii As Long, cjj As Long, ci As Long, cj As Long
        
            
            'On Error Resume Next
            For iurl = 0 To 2
                url = myurl(iurl)
                If iurl = 0 Then 'price
                    'Cells(1, dc) = "5 / 9 ____ " & url url
                    endrow = 3
                Else
                    endrow = Range("bo100000").End(xlUp).Row
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
            
                endrow = Range("bo10000").End(xlUp).Row
                For i = endrow To 7 Step -1
                    If Range("bo" & i) = "Date" Or Len(Range("bo" & i)) = 0 Then
                         Range("BO" & i & ":BU" & i).Delete Shift:=xlUp
                         
                  
                    End If
                Next i
                
            
            
                endrow = Range("bo10000").End(xlUp).Row
                
                Range("BO4:BU" & endrow).Select
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
                ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("BO5:bo" & endrow), _
                    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                With ActiveWorkbook.ActiveSheet.Sort
                    .SetRange Range("BO4:BU" & endrow)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
               
            '----end clear and sort data--------------------------------------------------------------------------
            
            
End Sub


Public Sub ConnectWinHttp(ByVal myurl As String, waittime As Integer)


     Dim oXMLHTTP As Object
     

      
      ConnectErr = False
      Set oXMLHTTP = CreateObject("winhttp.winhttpRequest.5.1")
    '  Set objStream = CreateObject("ADODB.stream")
      Set doc = New HTMLDocument
      With oXMLHTTP

            .Open "Get", myurl, False
            .Option(4) = 13056
          '  .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
            .send
                             
             Application.Wait Now() + TimeValue("00:00:01") * waittime
       
            
            doc.body.innerHTML = .responseText
            
             
           
            .abort
      End With
      
      Set oXMLHTTP = Nothing

     

End Sub

'Get Yahoo historic price, dividend and split
Sub Yahoo_Price_Dividend_Split(cd$, dc)
            
            'General variables
            Dim url As String, res As Variant, out_data() As Variant
            Dim endrow As Integer, datarng As Range
            
            'Historic price variables
            Dim timestamp As New Collection, quote As New Dictionary, openp As New Collection
            Dim high As New Collection, low As New Collection, closep As New Collection
            Dim adjclose As New Collection, volume As New Collection, historycol As Integer
            
            'div + split variables
            Dim div As New Dictionary, splits As New Dictionary, divsplitcol As Integer
            
            'Init vars
            today = DateDiff("s", "1/1/1970", Now()) 'today in unix timestamp
            url = "https://query1.finance.yahoo.com/v8/finance/chart/" & cd$ & "?period1=0&period2=" & today & "&interval=1mo&events=history|div|split&includeAdjustedClose=true"
            historycol = 7 'number of columns for historical prices
            divsplitcol = 3  'number of columns for dividends/splits
            
            ActiveSheet.Range(Cells(4, dc), Cells(100000, dc + historycol - 1)).ClearContents
            'ActiveSheet.Range(Cells(4, dc), Cells(100000, dc + historycol - 1)).ClearFormats
            
            '-----Request url-----
            For retry = 0 To 2
                'Request from yahoo
                With CreateObject("WinHttp.WinHttpRequest.5.1")
                    .Open "GET", url, False
                    .send
                    .waitForResponse
                    res = .responseText
                End With
                
                If InStr(1, res, "<!doctype html public") >= 1 Then
                    'Invalid response type
                    If retry >= 2 Then
                        'Max retry reached, exit early
                        Exit Sub
                    End If
                    'retry url
                    Application.Wait Now() + TimeValue("00:00:03")
                Else
                    'Valid response type, parse json and try next
                    Set res = ParseJson(res)
                    Exit For
                End If
            Next retry
            
            '-----Historic prices processing-----
            On Error Resume Next
                'Get historic prices as collection
                Set timestamp = res("chart")("result")(1)("timestamp")
                
                Set quote = res("chart")("result")(1)("indicators")("quote")(1)
                Set openp = quote("open")
                Set high = quote("high")
                Set low = quote("low")
                Set closep = quote("close")
                Set volume = quote("volume")
                
                Set adjclose = res("chart")("result")(1)("indicators")("adjclose")(1)("adjclose")
                
                If err <> 0 Or WorksheetFunction.Min(timestamp.Count, _
                                    openp.Count, _
                                    high.Count, _
                                    low.Count, _
                                    closep.Count, _
                                    volume.Count, _
                                    adjclose.Count) = 0 Then
                                    
                    'Error or one historical price have no entry, jump to div
                    Debug.Print "Historical price error"
                    On Error GoTo 0
                    GoTo div
                End If
            On Error GoTo 0
            
            'Fill array with processed prices
            ReDim out_data(1 To timestamp.Count, 1 To historycol)
            For i = 1 To timestamp.Count
                out_data(i, 1) = Format(DateAdd("s", CDbl(timestamp(i)), "1970/1/1"), "yyyy/mm/dd")
                out_data(i, 2) = HandleNulls(openp(i))
                out_data(i, 3) = HandleNulls(high(i))
                out_data(i, 4) = HandleNulls(low(i))
                out_data(i, 5) = HandleNulls(closep(i))
                out_data(i, 6) = HandleNulls(adjclose(i))
                out_data(i, 7) = HandleNulls(volume(i))
            Next i
            
            'Paste historic prices in sheet
            endrow = 4
            Set datarng = Cells(endrow + 1, dc)
            datarng.Resize(timestamp.Count, historycol).value = out_data
div:
            '-----Dividend processing-----
            On Error Resume Next
                'Get dividends as dictionary
                Set div = res("chart")("result")(1)("events")("dividends")
                
                If err <> 0 Or div.Count = 0 Then
                    'Error or no dividends, jump to split
                    Debug.Print "Dividend error"
                    On Error GoTo 0
                    GoTo split
                End If
            On Error GoTo 0
            
            'Fill array with processed dividends
            ReDim out_data(1 To div.Count, 1 To divsplitcol)
            i = 1
            For Each k In div.Keys
                out_data(i, 1) = Format(DateAdd("s", CDbl(div(k)("date")), "1970/1/1"), "yyyy/mm/dd")
                out_data(i, 2) = HandleNulls(div(k)("amount"))
                out_data(i, 3) = "Dividend"
                i = i + 1
            Next
            
            'Paste dividends in sheet
            endrow = Cells(100000, dc).End(xlUp).Row
            Set datarng = Cells(endrow + 1, dc)
            datarng.Resize(div.Count, divsplitcol).value = out_data
split:
            '-----Split processing-----
            On Error Resume Next
                'Get splits as dictionary
                Set splits = res("chart")("result")(1)("events")("splits")
                
                If err <> 0 Or splits.Count = 0 Then
                    'Error or no splits, jump to data sorting
                    Debug.Print "Split error"
                    On Error GoTo 0
                    GoTo data_sort
                End If
            On Error GoTo 0
            
            'Fill array with processed splits
            ReDim out_data(1 To splits.Count, 1 To divsplitcol)
            i = 1
            For Each k In splits.Keys
                out_data(i, 1) = Format(DateAdd("s", CDbl(splits(k)("date")), "1970/1/1"), "yyyy/mm/dd")
                out_data(i, 2) = "'" & splits(k)("splitRatio")
                out_data(i, 3) = "Split"
                i = i + 1
            Next
            
            'Paste splits in sheet
            endrow = Cells(100000, dc).End(xlUp).Row
            Set datarng = Cells(endrow + 1, dc)
            datarng.Resize(splits.Count, divsplitcol).value = out_data
data_sort:
            '-----Sort rows by date descending-----
            endrow = Cells(100000, dc).End(xlUp).Row
                
            ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
            ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range(Cells(5, dc), Cells(endrow, dc)), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            With ActiveWorkbook.ActiveSheet.Sort
                .SetRange Range(Cells(5, dc), Cells(endrow, dc + historycol - 1))
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            'Add column headers if there's valid data
            If Cells(5, dc).value <> "" Then
                Cells(4, dc).Resize(, historycol) = Array("Date", "Open", "High", "Low", "Close", "Adj Close", "Volume")
            End If
End Sub

' Modified from source: https://www.listendata.com/2021/02/excel-macro-historical-stock-data.html
Function HandleNulls(value As Variant) As Double
    If IsNumeric(value) Then
        HandleNulls = Round(CDbl(value), 6) ' Round to 6 d.p, consistent with known Yahoo format
    Else
        HandleNulls = 0 ' Return 0 if the value is null, empty, or not numeric
    End If
End Function





