Attribute VB_Name = "Module6"
Sub Macro10()
'
' mikeon _ 2015/10/11 _____的巨集
' ____台股

'
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.ScreenUpdating = True
[g8] = "___...".."
[g18] = "___...".."
Application.ScreenUpdating = False

 '____速度
    Application.MaxChange = 0.001
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
    
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Application.Calculation = xlManual
 
  Range("i:r,ba:cz,ha:hg").Clear
  Columns("ba:cz").ClearContents

'__市
Application.StatusBar = "________"姜禤々"
Application.ScreenUpdating = True
Application.ScreenUpdating = False

url = "https://www.twse.com.tw/exchangeReport/BWIBBU_d?response=html&selectType=ALL"

[i1] = url
Call ConnectMarketWatch(url, [ba1], 5)
    


    [f9] = [ba1] '
    [i3] = [n3]: [j3] = "__": [k3] = [p3]3]
    k = 1
    Do Until k > 3000
      If Cells(k, 53) = "1101" Then Exit Do
      k = k + 1
    Loop
    
    i = k: j = 4
    Do Until Cells(i, 53) = ""
     If IsNumeric(Cells(i, 53)) Then
        Cells(j, 9) = Cells(i, 53): Cells(j, 10) = Cells(i, 54)
        Cells(j, 11) = Cells(i, 57)
        If Cells(j, 11) = 0 Then Cells(j, 11) = ""
        j = j + 1
     End If
     i = i + 1
    Loop
    
    
'____指數


  url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&type=IND"
 
 Call ConnectMarketWatch(url, [HA1], 5)
 TSEIndex = ""
 For i = 1 To 1000
      If Range("HA" & i) = "_________" Then數" Then
          TSEIndex = Range("HB" & i)
          Columns("HA:HG").ClearContents
          Exit For
      End If
     
 Next i
 
 

[i2] = "____ (__ " & Format(TSEIndex, "#,##0") & Space(2) & (Left([f9], 3) + 1911) & "_" & Mid([f9], 5, 2) & "_" & Mid([f9], 8, 2) & "_)") & "日)"
[i3] = "": [k3] = ""
Call OTCC


[a19] = "": [f19] = "": [f9] = ""
Call fm6

End Sub

Sub OTCC()

'__ __Ale_Ale吳

Application.StatusBar = "________"d資料中"
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
      '____個股
      Range("ba:bf").Clear

      Dim url As String
       url = "https://www.tpex.org.tw/web/stock/aftertrading/peratio_analysis/pera_result.php?l=zh-tw&o=htm&c=&s=0,asc"
        [n1] = url
       'Call ConnectMarketWatch(url, [ba1], 5)
       
       Call ConnectXMLHTTP(url)
        GoSub Sub_ExtractData
             
       
       Range("n3:p" & Range("bc10000").End(xlUp).Row + 1).value = Range("ba2:bc" & Range("bc10000").End(xlUp).Row).value
       [f19] = Mid([ba1], InStr(1, [ba1], "____") + 5, 9) ' __' 日期
       Columns("ba:bf").ClearContents
      '____--數--
       url = "https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_index/st41_result.php?l=zh-tw&d=" & Left([f19], 6) & "&s=0,desc,1&o=htm"
       
       'Call ConnectMarketWatch(url, [ba1], 5)
       Call ConnectXMLHTTP(url)
        GoSub Sub_ExtractData
       
       
       [a19] = Application.VLookup([f19], Range("ba1:bf1000"), 5, False)
       [n2] = "____ (__ " & Format([a19], "#,##0") & Space(2) & (Left([f19], 3) + 1911) & "_" & Mid([f19], 5, 2) & "_" & Right([f19], 2) & "_)") & "日)"
       [n3] = "": [p3] = ""
Exit Sub
'---------------------------------------------------------------------------
Sub_ExtractData:

  ii = 0
    For Each tbl In doc.getElementsByTagName("table")
           Set rng = Range("ba1")
            
              For Each rw In tbl.Rows
                  For Each cl In rw.Cells
                         rng.value = Trim(cl.innerText)
                         Set rng = rng.Offset(, 1)
                      ii = ii + 1 'column
                  Next cl
                  Set rng = rng.Offset(1, -ii)
                  ii = 0
              Next rw
      Next tbl
                
 Return
      
       
End Sub


Sub Macro11()
'
' mikeon _ 2015/10/14 _____的巨集
' __S&P50000

'
On Error GoTo errorhandler

Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.ScreenUpdating = True
[g28] = "___...".."
Application.ScreenUpdating = False

 '____速度
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
        
   With Application '_____}廣福
       .MaxChange = 0.001
       .EnableCancelKey = xlInterrupt '_____\秀雯
       .Calculation = xlManual
    End With
  

 Range("s:x,ba:cz").Clear
 
 'S&P500 ___________}廣福桑撰寫
    For i = 1 To 501 Step 20
        ii = Round(i / 20) + 1
        iii$ = ii
        Application.StatusBar = "______  " & iii$ & " / 26" / 26"
        Application.ScreenUpdating = True
        Application.ScreenUpdating = False
        With ActiveSheet.QueryTables.Add(Connection:= _
            "URL;https://finviz.com/screener.ashx?v=111&f=idx_sp500&ft=4&r=" & CStr(i), Destination:=Range("bA1"))
            .Name = "screener.ashx?v=111&f=idx_sp500&ft=4&r=1"
            .BackgroundQuery = True
            .RefreshStyle = xlOverwriteCells
            .SaveData = True
            .AdjustColumnWidth = False
            .WebSelectionType = xlSpecifiedTables
            .WebFormatting = xlWebFormattingNone
            .WebTables = "22"
            .Refresh BackgroundQuery:=False
        End With

Range("bA1").QueryTable.Delete

        If i = 1 Then
Range("bA1:bK21").Copy Destination:=Range("bL1")
Range("bA1:bK21").ClearContents
        Else
            a = Range("bA1").End(xlDown).Row
            b = Range("bL1").End(xlDown).Row
Range("bA2:bK" & a).Copy Destination:=Range("bL" & b + 1)
Range("bA1:bK" & a).ClearContents
        End If
    Next i

    a = Range("bL1").End(xlDown).Row
Range("bL1:bV" & a).Copy Destination:=Range("bA1")
Range("bL1:bV" & a).ClearContents

'______謝徐桑

[s1] = "https://finviz.com/screener.ashx?v=111&f=idx_sp500&ft=4"
Cells(3, 19) = Cells(1, 54): Cells(3, 20) = Cells(1, 55): Cells(3, 21) = Cells(1, 57): Cells(3, 22) = Cells(1, 58): Cells(3, 23) = Cells(1, 60)

j = 4: i = 2
Do Until Trim(Cells(i, 54)) = ""
    Cells(j, 19) = Cells(i, 54): Cells(j, 20) = Cells(i, 55): Cells(j, 21) = Cells(i, 57): Cells(j, 22) = Cells(i, 58): Cells(j, 23) = Cells(i, 60)
    j = j + 1: i = i + 1
Loop

Columns("ba:cz").Clear

'__+__日期
'Application.StatusBar = "_SNP500___"々"
      
'      url_sp500 = "https://www.marketwatch.com/api/marketoverview/type/US"

'      Set doc = New HTMLDocument
'      With CreateObject("MSXML2.XMLHTTP")
            
'            .Open "GET", url_sp500, False
'            .send
'            Do: DoEvents: Loop Until .readyState = 4
'            Do: DoEvents: Loop Until .Status = 200
          
'            doc.body.innerHTML = .responseText
            
'            textbody = doc.getElementsByTagName("body")(0).innerText
'            [A29] = Mid(textbody, InStr(1, textbody, "S&P 500") + 8, (InStr((InStr(1, textbody, "S&P 500") + 9), textbody, " ")) - (InStr(1, textbody, "S&P 500") + 8))
                  
'            .abort
'            Set doc = Nothing
'      End With
   

'[s2] = "S&P500 " & " (__ " & Round([A29], 0) & "  " & [f29] & ")")"


Application.StatusBar = "_SNP500___"々"
Application.ScreenUpdating = True
Application.ScreenUpdating = False

url = "http://www.stockq.org/welcome.php"
Call ConnectMarketWatch(url, [ba1], 5)
For i = 1 To 300 '__數
If Trim(Cells(i, 66)) = "S&P 500" Then Exit For
Next i
[s3] = "": [w3] = ""
[s2] = "S&P 500 " & " (__ " & Format(Cells(i, 67), "#,##0") & "  " & Left([bn3], InStr(1, [bn3], " ") - 1) & ")")"
Call fm6


Exit Sub

errorhandler:

Debug.Print err.Number, err.Description

err.Clear
Resume Next


End Sub


Sub Macro12()
'
' mikeon _ 2015/10/21 _____的巨集
' ____日經
'
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.ScreenUpdating = True
[g38] = "___...".."
Application.ScreenUpdating = False

 '____速度
    Application.MaxChange = 0.001
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
    
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Application.Calculation = xlManual
 
Range("z:ac,ba:cz").Clear
 
Application.StatusBar = "______"禤々"
Application.ScreenUpdating = True
Application.ScreenUpdating = False
 
'__225 ___________悎}廣福桑撰寫

Dim myIE As InternetExplorer
Dim myIEdoc As HTMLDocument         '____物件
Dim thePars As IHTMLElementCollection

Dim theTBL As HTMLTable
Dim crRow As HTMLTableRow
Dim crCell As HTMLTableCell
Dim i As Long

Set myIE = New InternetExplorer '______InternetExplorer__orer物件

With myIE
    '__IE__視窗
    .Visible = False
    '__URLRL
    .navigate "https://www.investing.com/indices/japan-ni225-components"
    Do While .readyState <> 4: DoEvents: Loop
    t1 = Timer
    Do Until Timer > t1 + 5 '__5___________H改成等待几秒
        DoEvents
    Loop
    '______Alex___lex桑撰寫
    'Set resultClasses = .document.getElementsByTagName("a")
    'For Each resultclass In resultClasses
    '   If UCase(resultclass.className) = "NEWBTN TOGGLEBUTTON LIGHTGRAY LAST" Then
    '           resultclass.Click
    '          Exit For
    'End If
    'Next resultclass
    
    .document.getElementById("filter_fundamental").Focus
    .document.getElementById("filter_fundamental").Click
    
    
'______謝吳桑
End With

    t1 = Timer
    Do Until Timer > t1 + 5 '__5___________H改成等待几秒
DoEvents
    Loop

    Set PT = ThisWorkbook.Worksheets("__").Range("AZ1")")
    Set theTBL = myIE.document.getElementsByTagName("table")(1)
    If theTBL.Cells(0) <> PT.value Then
        For Each crRow In theTBL.Rows
            i = 0
            For Each crCell In crRow.Cells
                PT.Offset(0, i) = crCell.innerText
                i = i + 1
            Next
            Set PT = PT.Offset(1, 0)
        Next
    End If

myIE.Quit       '_____s覽器
Set myIE = Nothing

'______謝徐桑

[Z1] = "https://www.investing.com/indices/japan-ni225-components"
Cells(3, 26) = [ba1]: Cells(3, 27) = [be1]
j = 4: i = 2
Do Until Trim(Cells(i, 53)) = ""
    Cells(j, 26) = Cells(i, 53): Cells(j, 27) = Cells(i, 57)
    j = j + 1: i = i + 1
Loop
Columns("ba:cz").Clear

'__+__日期

Application.StatusBar = "___225___"數中"
Application.ScreenUpdating = True
Application.ScreenUpdating = False
url = "http://www.stockq.org/welcome.php"
Call ConnectMarketWatch(url, [ba1], 5)
For i = 1 To 300 '__數
If Trim(Cells(i, 53)) = "__225" Then Exit Foror
Next i
[z2] = "__225 " & " (__ " & Format(Cells(i, 54), "#,##0") & "  " & Left([bn3], InStr(1, [bn3], " ") - 1) & ")" ")"
[z3] = "": [aa3] = ""
Call fm6

End Sub


Sub Macro13()
'
' mikeon _ 2015/10/22 _____的巨集
' ____港股

'
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.ScreenUpdating = True
[g48] = "___...".."
Application.ScreenUpdating = False

 '____速度
    Application.MaxChange = 0.001
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
    
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Application.Calculation = xlManual

 Range("ad:ag,ba:cz").Clear
 Application.StatusBar = "______"禤々"
 Application.ScreenUpdating = True
 Application.ScreenUpdating = False

'____50 ___________‘悎}廣福桑撰寫

Dim myIE As InternetExplorer
Dim myIEdoc As HTMLDocument         '____物件
Dim thePars As IHTMLElementCollection

Dim theTBL As HTMLTable
Dim crRow As HTMLTableRow
Dim crCell As HTMLTableCell
Dim i As Long

Set myIE = New InternetExplorer '______InternetExplorer__orer物件

With myIE
    '__IE__視窗
    .Visible = False
    '__URLRL
    .navigate "https://www.investing.com/indices/hang-sen-40-components"
    Do While .readyState <> 4: DoEvents: Loop
    t1 = Timer
    Do Until Timer > t1 + 5 '__5___________H改成等待几秒
        DoEvents
    Loop
    '______Alex___lex桑撰寫
   ' Set resultClasses = .document.getElementsByTagName("a")
   ' For Each L0 In .document.getElementsByTagName("a")
       
   '    Debug.Print L0.ID, L0.className
       
       
       'If UCase(resultclass.className) = "NEWBTN TOGGLEBUTTON LIGHTGRAY LAST" Then
       '        resultclass.Click
       '       Exit For
       ' End If
   ' Next L0
     
    .document.getElementById("filter_fundamental").Focus
    .document.getElementById("filter_fundamental").Click
    
    
'______謝吳桑
End With

    t1 = Timer
    Do Until Timer > t1 + 5 '__5___________H改成等待几秒
DoEvents
    Loop
    
    Set PT = ThisWorkbook.Worksheets("__").Range("AZ1")")
    Set theTBL = myIE.document.getElementsByTagName("table")(1)
    If theTBL.Cells(0) <> PT.value Then
        For Each crRow In theTBL.Rows
            i = 0
            For Each crCell In crRow.Cells
                PT.Offset(0, i) = crCell.innerText
                i = i + 1
            Next
            Set PT = PT.Offset(1, 0)
        Next
'        PT.Worksheet.Columns.AutoFit
    End If

myIE.Quit       '_____s覽器
Set myIE = Nothing

'______謝徐桑

[ad1] = "https://www.investing.com/indices/hang-sen-40-components"

Cells(3, 30) = Cells(1, 53): Cells(3, 31) = Cells(1, 57)
j = 4: i = 2
Do Until Trim(Cells(i, 53)) = ""
    Cells(j, 30) = Cells(i, 53): Cells(j, 31) = Cells(i, 57)
    j = j + 1: i = i + 1
Loop

Columns("ba:cz").Clear

'__+__日期
Application.StatusBar = "___50___"數中"
Application.ScreenUpdating = True
Application.ScreenUpdating = False
url = "http://www.stockq.org/welcome.php"
Call ConnectMarketWatch(url, [ba1], 5)
    
For i = 1 To 300 '__數
If Trim(Cells(i, 53)) = "____" Then Exit For For
Next i
[ad2] = "____50 " & " (__ " & Format(Cells(i, 54), "#,##0") & "  " & Left([bn3], InStr(1, [bn3], " ") - 1) & ")" & ")"
[ad3] = "": [ae3] = ""
Call fm6

End Sub

Sub Macro14()
'
' mikeon _ 2016/8/22 _____的巨集
' ____中股

'
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.ScreenUpdating = True
[g58] = "___...".."
Application.ScreenUpdating = False

 '____速度
    Application.MaxChange = 0.001
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
    
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Application.Calculation = xlManual
Range("ah:aL,ba:cz").Clear
 
'____180 ___________‘悎}廣福桑撰寫

For i = 1 To 9
    Application.StatusBar = "______  " & i & " / 9"" / 9"
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.aastock.com/tc/cnhk/market/sh-connect.aspx?cat=2&t=1&s=1&o=1&p=4&page=" & CStr(i) _
        , Destination:=Range("BA1"))
            .BackgroundQuery = True
            .RefreshStyle = xlOverwriteCells
            .SaveData = True
            .AdjustColumnWidth = False
            .WebSelectionType = xlSpecifiedTables
            .WebFormatting = xlWebFormattingNone
            .WebTables = "14"
            .Refresh BackgroundQuery:=False
    End With
    Range("BA1").QueryTable.Delete
    If i = 1 Then
        a = Range("BA1").End(xlDown).Row
        Range("BA1:BK" & a).Copy Destination:=Range("BN1")
        Range("BA1:BK" & a).ClearContents
    Else
        a = Range("BA1").End(xlDown).Row
        b = Range("BN1").End(xlDown).Row
        Range("BA2:BK" & a).Copy Destination:=Range("BN" & b + 1)
        Range("BA1:BK" & a).ClearContents
    End If
Next i
a = Range("BL1").End(xlDown).Row
Range("BN1:BX" & a).Copy Destination:=Range("BA1")
Range("BN1:BX" & a).ClearContents

'______謝徐桑

[ah1] = "https://www.aastock.com/tc/cnhk/market/sh-connect.aspx?cat=2&t=1&s=1&o=1&p=4&page="

k = 1: i = 3
Do Until Cells(k, 53) = "" Or k > 500
If Right(Cells(k, 53), 1) = "H" Then
i = i + 1
Cells(i, 34) = Left(Cells(k, 53), 6)
Cells(i, 35) = Cells(k - 1, 53)
Cells(i, 36) = Cells(k - 1, 59)
End If
k = k + 1
Loop

[ah3] = "__": [ai3] = "__": [aj3] = "___"本益比"
Columns("ba:cz").Clear

'__+__日期
Application.StatusBar = "______"數中"
Application.ScreenUpdating = True
Application.ScreenUpdating = False
url = "http://www.stockq.org/welcome.php"
Call ConnectMarketWatch(url, [ba1], 5)

k = 1: Do Until Trim(Cells(k, 53)) = "____" Or k > 300 300
k = k + 1
Loop
[ah2] = "____180" & " (__ " & Format(Cells(k, 54), "#,##0") & "  " & Left([bn3], InStr(1, [bn3], " ") - 1) & ")" & ")"
[ah3] = "": [aj3] = ""
Call fm6

End Sub

Public Sub fm6()

Columns("ba:cz").Clear
Cells.Select
    Selection.RowHeight = 16
    Selection.ColumnWidth = 7.5
    With Selection.Font
        .Name = "____"體"
        .Name = "Arial"
        .FontStyle = "__""
        .Size = 10
    End With

Range("t:t, z:z, ad:ad").ColumnWidth = 18
Columns("u:u").ColumnWidth = 9

Range("a5:a5, a15:a15, a25:a25, a35:a35, a45:a45, a55:a55, g8:g8, g18:g18, g28:g28, g38:g38, g48:g48, g58:g58").Font.Color = -16776961

Range("k:k,p:p, w:w, aa:aa, ae:ae, aj:aj").NumberFormatLocal = "#,##0.0_);(#,##0.0)"
Range("g8:g8, g18:g18, g28:g28, g38:g38, g48:g48, g58:g58, i:aj").HorizontalAlignment = xlCenter
Range("a5:a5, a15:a15, a25:a25, a35:a35, a45:a45, a55:a55, i1:aj2, s:u, z:z, ad:ad").HorizontalAlignment = xlLeft

[g8] = "": [g18] = "": [g28] = "": [g38] = "": [g48] = "": [g58] = ""
[dz100].Select

Call ProtectSheet(ActiveSheet)
Application.Calculation = xlAutomatic           '_____}廣福
Application.StatusBar = "__""
Beep

End Sub


Sub stwid()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤台股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("i4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("i4:k10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6
End Sub



Sub stwpe()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤台股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("k4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("i4:k10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6
End Sub

Sub stwoid()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤上櫃
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("n4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("n4:p10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6


End Sub

Sub stwope()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤上櫃
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("p4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("n4:p10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6


End Sub





Sub susid()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤美股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("s4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("s4:w10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub




Sub suspe()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤美股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("w4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("s4:w10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub

Sub sjpid()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤日股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("z4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("z4:aa10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub


Sub sjppe()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤日股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("aa4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("z4:aa10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub


Sub shkid()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤港股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("ad4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("ad4:ae10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub



Sub shkpe()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤港股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("ae4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("ad4:ae10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub


Sub scnid()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤中股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("ah4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("ah4:aj10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub


Sub scnpe()

'
' mikeon _ 2016/5/5 _____的巨集
' __:____盤中股
'

Application.ScreenUpdating = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Sheets("__").Selectct
Call UnprotectSheet(ActiveSheet)
Application.Calculation = xlManual

ActiveWorkbook.Worksheets("__").Sort.SortFields.Clearar
    ActiveWorkbook.Worksheets("__").Sort.SortFields.Add Key:=Range("aj4"), SortOn _ _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("__").Sortrt
        .SetRange Range("ah4:aj10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call fm6

End Sub

