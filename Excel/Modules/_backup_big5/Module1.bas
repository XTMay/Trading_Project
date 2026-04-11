Attribute VB_Name = "Module1"
Sub Macro1()
'
' mikeon _ 2005/2/9 _____ªº¥¨¶°
'

    
    Sheets("__").Selectct
    Call UnprotectSheet(ActiveSheet)
    
    Dim urltype, urltype1  As String
    Dim url As String, rsrange As String
    Dim basicdata As Variant
    
    Application.MaxChange = 0.001 '____³t«×
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With

Application.EnableCancelKey = xlInterrupt '_____\¨q¶²
Application.Calculation = xlManual
On Error Resume Next
    
    Range("Y31").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-1]C="""","""",IF(ISNUMBER(YEAR(R[-1]C)),YEAR(R[-1]C)+11&""/""&MONTH(R[-1]C)&""/""&DAY(R[-1]C),LEFT(R[-1]C,3)+1911&""/""&MID(R[-1]C,5,2)&""/""&RIGHT(R[-1]C,2)))"
    
    [a18] = "______________________________________"À¡B¤½¦¡©M©úµP¦Ó¥¼µù©ú¥X³BªÌ§Y¬°¤p°½¡C"
    [a95] = "______________(Michael On)__"Michael On)©Ò¦³"
    [i14] = Left(ActiveWorkbook.Name, InStr(1, ActiveWorkbook.Name, ".") - 1)
    
    Range("A2").NumberFormatLocal = "@"
    cd$ = [a2]: [f17] = 1: [K12] = "": [i16] = "": [w32] = "": [w33] = ""
    Range("af:eq").Clear
    dc1 = 75 '____¸ê®Æ
    dc2 = 32 '______·l¯qªí
    dc6 = 87 '__§Q
    dc7 = 98 '__½X
    dc8 = 103 '___ë¸ê
    dc10 = 120 '__¦¬
    dc11 = 130 '____ªÑ»ù
    
    [y23] = 0: [y24] = 0
    Range("v47").ClearContents
    Range("v47").value = "EXCEL__:" & Application.Version & "_(" & cd$ & ")")"
    Range("E12").FormulaR1C1 = "=IFERROR(R[50]C[14],""_"")"""
    Range("k3").FormulaR1C1 = "=R[11]C[6]"
    [k10] = 12
    Range("K11").FormulaR1C1 = "=R[2]C[13]"
    Range("O10").FormulaR1C1 = "=VLOOKUP(R[5]C[-1]-ROUNDUP(RC[2],0)+1,R[5]C[-1]:R[42]C[2],2,FALSE)"
    Range("P10").FormulaR1C1 = "=VLOOKUP(R[5]C[-2]-ROUNDUP(RC[1],0)+1,R[5]C[-2]:R[42]C[1],4,FALSE)"
    Range("q10").FormulaR1C1 = "=YEAR(TODAY())-IF(R[13]C[-2]<>"""",R[13]C[-3],IF(R[12]C[-2]<>"""",R[12]C[-3],IF(R[11]C[-2]<>"""",R[11]C[-3],IF(R[10]C[-2]<>"""",R[10]C[-3],IF(R[9]C[-2]<>"""",R[9]C[-3],IF(R[8]C[-2]<>"""",R[8]C[-3],IF(R[7]C[-2]<>"""",R[7]C[-3],IF(R[6]C[-2]<>"""",R[6]C[-3],R[5]C[-3]-MONTH(TODAY())/12))))))))"

    urltype = RSURLTYPE
    
    '____¸ê®Æ
    Application.StatusBar = "_____  1 / 11"/ 11"
    
    dc = dc1
    url = "http://" & urltype & "/z/zc/zca/zca_" & cd$ & ".djhtm"
    Cells(1, dc + 1).value = "1 /11 ____ " & url url
    rsrange = TSECFSQT(url, Cells(2, dc), "1")
    ERRLOG "1 / 11 ____(" & rsrange & ")", err.Numbermber
    
    With Range(rsrange)
        Set c = .Find("___", LookIn:=xlValues)es)
        If Not c Is Nothing Then _
        [q14] = c.Offset(0, 1)
    
        Set c = .Find("____", LookIn:=xlValues)ues)
        If [a2] <> "1101" And Left(Trim(c), 2) = "__" Then GoTo errrr
    
        Set c = .Find("_____", LookIn:=xlValues)lues)
        If Not c Is Nothing Then
        [w30] = c.Offset(0, -1): [w31].Calculate: c.Offset(0, -1) = [w31] '____®É¶¡
        [w30] = c.Offset(1, -1): [w31].Calculate: c.Offset(1, -1) = [w31] '____®É¶¡
        End If
     
        Set c = .Find("___", LookIn:=xlValues)es)
        If Not c Is Nothing Then _
        [i1] = Year(Date) & "/" & Mid(c.Offset(-1, 0), 7, 5) '_____æ©ö¤é
    
        Set c = .Find("___", LookIn:=xlValues)es)
        If Not c Is Nothing Then _
        [y23] = c.Offset(0, 1)
    
        '________²§±`½Õ¾ã
        Set c1 = .Find("___", LookIn:=xlValues)es)
        Set c2 = .Find("___", LookIn:=xlValues)es)
        Set c3 = .Find("____", LookIn:=xlValues)ues)
        If Not (c1 Is Nothing Or c2 Is Nothing Or c3 Is Nothing) Then
            If Not c1.Offset(0, 1).Column >= c2.Offset(0, -1).Column Then
                basicdata = _
                Range(c2.Offset(0, -1).Address, Cells(c3.Row, c2.Offset(0, -1).Column + 6))
                Range(c2.Offset(0, -1).Address, Cells(c3.Row, c2.Offset(0, -1).Column + 6)).Clear
                Range(c1.Offset(0, 1).Address, Cells(c3.Row, c1.Offset(0, 1).Column + 6)) = basicdata
            End If
        End If
        
        Set c = Range(rsrange).Find("____", LookIn:=xlValues)ues)
        If Not c Is Nothing Then _
        [y16] = c.Offset(0, 1)  '____¤ñ­«
    End With
    Set c = Nothing
    
    For i = 1 To 200
    If Left(Cells(i, dc + 1), 2) = "__" Then Exit Foror
    Next i
    [x36] = Cells(i, dc + 2) '____®É¶¡

    
    '____®§¤é
    url = "http://" & urltype & "/z/zc/zci/zci_" & cd$ & ".djhtm"
    Cells(50, dc + 1).value = "____ " & url url
    rsrange = TSECFSQT(url, Cells(51, dc), "1")
    ERRLOG "___(" & rsrange & ")", err.Numberber
    
    Set c = Range(rsrange).Find("___", LookIn:=xlValues)es)
    If Not c Is Nothing Then
    [w30] = c.Offset(1, 1): [w31].Calculate: c.Offset(1, 1) = [w31]: [w32] = [w31] '___v¤é
    [w30] = c.Offset(1, 0): [w31].Calculate: c.Offset(1, 0) = [w31]: [w33] = [w31] '___§¤é
    End If
    Set c = Nothing

    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d
    
    '____¯qªí
    Application.StatusBar = "_______  2 / 11"2 / 11"
    dc = dc2
    url = "http://" & urltype & "/z/zc/zcq/zcq_" & cd$ & ".djhtm"
    Cells(1, dc).value = "2 / 11 ______ " & url & url
    
    Call TWReport(url, dc)
     
    If Cells(1000, dc).End(xlUp).Row < 10 Then '_____D¦X¨Ö
   
        Range(Cells(2, dc), Cells(1000, dc + 9)).Clear
        Application.StatusBar = "________  2 / 11" 2 / 11"
        url = "http://" & urltype & "/z/zc/zcq/zcq0_" & cd$ & ".djhtm"
        Cells(1, dc) = "2 / 11 _______ " & url" & url
        Call TWReport(url, dc)
    
    End If


     'rsrange = TSECFSQT(url, Cells(2, dc), "3")
   ' If Range(rsrange).Rows.Count > 6 And SEASONCU(dc + 1) = False Then
   '     ERRLOG "2 / 11 ______(" & rsrange & ")", err.NumberNumber
   ' Else
   '     Range(rsrange).Clear
   '     Application.StatusBar = "________  2 / 11" 2 / 11"
   '     url = "http://" & urltype & "/z/zc/zcq/zcq0_" & cd$ & ".djhtm"
   '     Cells(1, dc) = "2 / 11 _______ " & url" & url
   '     rsrange = TSECFSQT(url, Cells(2, dc), "3")
   '     ERRLOG "2 / 11 _______(" & rsrange & ")", err.Number.Number
   ' End If
    SEASONTF dc + 1
    
         [m1] = "": Dim ag As Date
         If Len(Sheets("__").[t3]) > 5 Thenen
         [m1] = Sheets("__").[t3]3]
         If Left([ag4], 1) = "1" Then ag = Year(Now()) & "/3/31"
         If Left([ag4], 1) = "2" Then ag = Year(Now()) & "/6/30"
         If Left([ag4], 1) = "3" Then ag = Year(Now()) & "/9/30"
         If Left([ag4], 1) = "4" Then ag = Year(Now()) - 1 & "/12/31"
         If DateDiff("d", [m1], ag) < 40 Then
         [m1] = "N"
         [b14] = "_______"¥¼¶i¨Ó"
         GoTo err
         End If
         End If
    
    '______­t¶Åªí
    Application.StatusBar = "_________  3 / 11"  3 / 11"
    url = "http://" & urltype & "/z/zc/zcp/zcpa/zcpa_" & cd$ & ".djhtm"
    Cells(1, dc + 11).value = "3 / 11 ________ " & url " & url
    
    
    Call TWReport(url, dc + 11)
   
    If Cells(1000, dc + 11).End(xlUp).Row < 10 Then '_____D¦X¨Ö
   
        Range(Cells(2, dc + 11), Cells(1000, dc + 11 + 9)).Clear
        Application.StatusBar = "__________  3 / 11"í  3 / 11"
        url = "http://" & urltype & "/z/zc/zcp/zcpa/zcpa0_" & cd$ & ".djhtm"
        Cells(1, dc + 11) = "3 / 11 _________ " & urlí " & url
        Call TWReport(url, dc + 11)
    
    End If
    
    
   ' rsrange = TSECFSQT(url, Cells(2, dc + 11), "3")
   ' If Range(rsrange).Rows.Count > 6 And SEASONCU(dc + 12) = False Then
   '     ERRLOG "3 / 11 ________(" & rsrange & ")", err.Numberr.Number
   ' Else
   '     Range(rsrange).Clear
   '     Application.StatusBar = "__________  3 / 11"í  3 / 11"
   '     url = "http://" & urltype & "/z/zc/zcp/zcpa/zcpa0_" & cd$ & ".djhtm"
   '     Cells(1, dc + 11) = "3 / 11 _________ " & urlí " & url
   '     rsrange = TSECFSQT(url, Cells(2, dc + 11), "3")
   '     ERRLOG "3/ 11 _________(" & rsrange & ")", err.Numberrr.Number
   ' End If
    SEASONTF dc + 12
    


    '____¯qªí
    Application.StatusBar = "_______  4 / 11"4 / 11"
    url = "http://" & urltype & "/z/zc/zcq/zcqa/zcqa_" & cd$ & ".djhtm"
    Cells(1, dc + 22).value = "4 / 11 ______ " & url & url
   
    Call TWReport(url, dc + 22)
   
    If Cells(1000, dc + 22).End(xlUp).Row < 10 Then '_____D¦X¨Ö
   
        Range(Cells(2, dc + 22), Cells(1000, dc + 22 + 9)).Clear
        Application.StatusBar = "________  4 / 11" 4 / 11"
        url = "http://" & urltype & "/z/zc/zcq/zcqa/zcqa0_" & cd$ & ".djhtm"
        Cells(1, dc + 22) = "4 / 11 _______ " & url" & url
        Call TWReport(url, dc + 22)
    
    End If
   
   
   ' rsrange = TSECFSQT(url, Cells(2, dc + 22), "2")
   ' If Range(rsrange).Rows.Count > 6 And YEARCU(dc + 23) = False Then
   '     ERRLOG "______(" & rsrange & ")", err.NumberNumber
   ' Else
   '     Range(rsrange).Clear
   '     Application.StatusBar = "________  4 / 11" 4 / 11"
   '     url = "http://" & urltype & "/z/zc/zcq/zcqa/zcqa0_" & cd$ & ".djhtm"
   '     Cells(1, dc + 22) = "4 / 11 _______ " & url" & url
   '     rsrange = TSECFSQT(url, Cells(2, dc + 22), "3")
   '     ERRLOG "4 / 11 _______(" & rsrange & ")", err.Number.Number
   ' End If
    YEARTF dc + 23
    
    '______­t¶Åªí
    Application.StatusBar = "_________  5 / 11"  5 / 11"
    url = "http://" & urltype & "/z/zc/zcp/zcpb/zcpb_" & cd$ & ".djhtm"
    Cells(1, dc + 33).value = "5 / 11 ________ " & url " & url
    
     Call TWReport(url, dc + 33)
   
    If Cells(1000, dc + 33).End(xlUp).Row < 10 Then '_____D¦X¨Ö
   
        Range(Cells(2, dc + 33), Cells(1000, dc + 33 + 9)).Clear
        Application.StatusBar = "__________  5 / 11"í  5 / 11"
        url = "http://" & urltype & "/z/zc/zcp/zcpb/zcpb0_" & cd$ & ".djhtm"
        Cells(1, dc + 33) = "5 / 11 _________ " & urlí " & url
       
        Call TWReport(url, dc + 33)
    
    End If
   
    
    
    'rsrange = TSECFSQT(url, Cells(2, dc + 33), "2")
    'If Range(rsrange).Rows.Count > 6 And YEARCU(dc + 34) = False Then
    '    ERRLOG "5 / 11 ________(" & rsrange & ")", err.Numberr.Number
    'Else
    '    Range(rsrange).Clear
    '    Application.StatusBar = "__________  5 / 11"í  5 / 11"
    '    url = "http://" & urltype & "/z/zc/zcp/zcpb/zcpb0_" & cd$ & ".djhtm"
    '    Cells(1, dc + 33) = "5 / 11 _________ " & urlí " & url
    '    rsrange = TSECFSQT(url, Cells(2, dc + 33), "3")
    '    ERRLOG "5 / 11 _________(" & rsrange & ")", err.Numberrr.Number
    'End If
    YEARTF dc + 34
    
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d
    
    For i = 0 To 3 '_________vlookup§Qvlookup
        For j = 1 To 200
            If Cells(j, dc + 11 * i) = "" And Cells(j + 1, dc + 11 * i) = "" Then Exit For
            If Not Cells(j, dc + 11 * i).value = Trim(Cells(j, dc + 11 * i).value) Then _
            Cells(j, dc + 11 * i).value = Trim(Cells(j, dc + 11 * i).value)
        Next j
    Next i
    
    For k = 1 To 200 '_____@­P¤Æ
        If Cells(k, dc) = "" And Cells(k + 1, dc) = "" Then Exit For
        If Left(Cells(k, dc).value, 6) = "______" Then Cells(k, dc).value = "______"»´Áµ|«á²b§Q"
        If Left(Cells(k, dc).value, 6) = "______" Then Cells(k, dc).value = "______"Ö¼ÆªÑÅv²b§Q"
    Next k
    
    For k = 1 To 200
    If Cells(k, dc + 22) = "" And Cells(k + 1, dc + 22) = "" Then Exit For
        If Left(Cells(k, dc + 22).value, 6) = "______" Then Cells(k, dc + 22).value = "______"»´Áµ|«á²b§Q"
        If Left(Cells(k, dc + 22).value, 6) = "______" Then Cells(k, dc + 22).value = "______"Ö¼ÆªÑÅv²b§Q"
    Next k
    
    '__§Q
    Application.StatusBar = "___  6 / 11"11"
    urltype1 = urltype
    If urltype = "justdata.yuanta.com.tw" Then urltype1 = "pscnetinvest.moneydj.com.tw" '___________¡ªø±o¤£¤@¼Ë
    dc = dc6
    url = "http://" & urltype1 & "/z/zc/zcc/zcc_" & cd$ & ".djhtm"
    Cells(1, dc).value = "6 / 11 __ " & urlrl
    rsrange = TSECFSQT(url, Cells(2, dc), "3")
    ERRLOG "6 / 11 __(" & rsrange & ")", err.Numberer
    
    For i = 3 To 300
    Cells(i, dc6 - 1) = Left(Cells(i, dc6), 4)
    Next i
    
    For i = 3 To 300
    If Cells(i + 1, dc6 - 1) = Cells(i, dc6) Then Cells(i, dc6 - 1) = ""
    Next i
                   
    '__½X
    Application.StatusBar = "___  7 / 11"11"
    dc = dc7
    url = "http://" & urltype & "/z/zc/zcj/zcj_" & cd$ & ".djhtm"
    Cells(1, dc).value = "7 / 11 __ " & urlrl
    rsrange = TSECFSQT(url, Cells(2, dc), "4")
    ERRLOG "7 / 11 __(" & rsrange & ")", err.Numberer
    
    i = 1: Do Until Cells(i, dc) = "____" Or i > 50 > 50
    i = i + 1
    Loop
    [K12] = Cells(i, dc + 3)
    
    '___ë¸ê
    Application.StatusBar = "____  8 / 11" 11"
    dc = dc8
    url = "http://" & urltype & "/z/zc/zcg/zcg_" & cd$ & ".djhtm"
    Cells(1, dc).value = "8 / 11 ___ " & urlurl
    rsrange = TSECFSQT(url, Cells(2, dc), "3")
    ERRLOG "8 / 11 ___(" & rsrange & ")", err.Numberber
    
    '__Åý
    Application.StatusBar = "___  9 / 11"11"
    url = "http://" & urltype & "/z/zc/zck/zck_" & cd$ & ".djhtm"
    Cells(1, dc + 9).value = "9 / 11 __ " & urlrl
    rsrange = TSECFSQTF(url, xlWebFormattingNone, xlEntirePage, "", Cells(2, dc + 8), "")
    ERRLOG "9 / 11 __(" & rsrange & ")", err.Numberer
    
    '__¦¬
    Application.StatusBar = "___  10 /11"11"
    dc = dc10
    url = "http://" & urltype & "/z/zc/zch/zch_" & cd$ & ".djhtm"
    Cells(1, dc).value = "10 / 11 __ " & urlrl
    rsrange = TSECFSQT(url, Cells(2, dc), """oMainTable""")
    ERRLOG "10 / 11 __(" & rsrange & ")", err.Numberer
    
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d
    
    j = 1: Do Until Cells(j, dc) = "_/_" Or j > 5050
    j = j + 1
    Loop
    
    For k = 1 To 12
        Cells(59 - k, 7) = Left(Cells(j + k, dc), Len(Cells(j + k, dc)) - 3) + 1911
        Cells(59 - k, 7) = Cells(59 - k, 7) & "/" & Right(Cells(j + k, dc), 2)
    Next k
    
    k = 1: Do Until Cells(j + k, dc) = "" And Cells(j + k + 1, dc) = ""
    Cells(j + k, dc) = Left(Cells(j + k, dc), Len(Cells(j + k, dc)) - 3) + 1911 & "/" & Right(Cells(j + k, dc), 2)
    k = k + 1
    Loop
    
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d
    
    '__¥«
    Application.StatusBar = "_____  11 / 11"/ 11"
    yr$ = [n15]: mo$ = Right([a58], 2)
    If mo$ = "01" Then mo$ = "02"
    [v35].NumberFormatLocal = "G/____"æ¦¡"
    If [v35] = 1 Then mo$ = "12" '_____¼¹L¦~
    dc = dc11
    url = "https://www.twse.com.tw/rwd/zh/afterTrading/FMSRFK?stockNo=" & cd$ & "&response=html"
    Cells(1, dc).value = "11 / 11 ______ " & url & url
    rsrange = TSECFSQT(url, Cells(2, dc), "1")
    
    If Not Cells(3, dc) = "" Then
    ERRLOG "11 / 11 ________(" & rsrange & ")", err.Numberr.Number
    url = "https://www.twse.com.tw/rwd/zh/afterTrading/FMNPTK?stockNo=" & cd$ & "&response=html"
    Cells(19, dc).value = "11 / 11 ______ " & url & url
    rsrange = TSECFSQT(url, Cells(20, dc), "1")
    ERRLOG "11 / 11 ________(" & rsrange & ")", err.Numberr.Number
    
    Else
        Range("ea:ez").Clear     'OTC__»ù
        Call OTC(cd$, dc)
        If IsError([c12]) Then
            If Cells(5, dc).value = "" Then GoTo his '_______"¥vªÑ»ù"
        Else
            If Cells(5, dc).value = "" And [c12] > 10 Then GoTo his '_______"¥vªÑ»ù"
        End If
    End If

ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
Delete_Pictures '__Alex_x§d

his:
    dc = dc11
    For k = 1 To 80 '____ªÑ»ù
    If k > 24 And Cells(k, dc) = "" Then Exit For
        If IsNumeric(Cells(k, dc).value) And Cells(k, dc).value > 1 Then Cells(k, dc).value = Cells(k, dc).value + 1911
    Next k
    
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d


[b14] = ""
err:    [y24] = [q14]
        With ActiveSheet
        .Hyperlinks.Add Anchor:=[h15], _
        Address:="https://tw.stock.yahoo.com/q/ta?s=" & cd$, TextToDisplay:="___"¹Ï"
    
        .Hyperlinks.Add Anchor:=[g15], _
        Address:="http://tw.stock.yahoo.com/q/q?s=" & cd$, TextToDisplay:="__"D"
    
        .Hyperlinks.Add Anchor:=[f15], _
        Address:="http://www.moneydj.com/KMDJ/Wiki/WikiSubjectList.aspx?a=TW." & cd$, TextToDisplay:="__"Ð"
        
        .Hyperlinks.Add Anchor:=[e15], _
        Address:="http://webretro.fortunengine.com.tw/stock/tool/diy_form.cfm", TextToDisplay:="___"¾¹"
    
        .Hyperlinks.Add Anchor:=[d15], _
        Address:="http://mikeon88.blogspot.tw/2012/02/blog-post_12.html", TextToDisplay:="GDP"
        
        .Hyperlinks.Add Anchor:=[k17], _
        Address:="http://mikeon88.imotor.com/forumdisplay.php?fid=2", TextToDisplay:="___2"Ï2"
        
         .Hyperlinks.Add Anchor:=[j17], _
        Address:="http://mikeon88.freebbs.tw/forumdisplay.php?fid=52", TextToDisplay:="___1"Ï1"
   
        .Hyperlinks.Add Anchor:=[c15], _
        Address:="https://tradingeconomics.com/taiwan/gdp-growth-annual", TextToDisplay:="GDP_"""
        
        .Hyperlinks.Add Anchor:=[j15], _
        Address:="https://mikeon88.666forum.com/f1-forum", TextToDisplay:="____"×°Ï"
    
        .Hyperlinks.Add Anchor:=[j16], _
        Address:="http://mikeon88.blogspot.com", TextToDisplay:="_______"³¡¸¨®æ"
    End With
    
For m = 2 To 20
If Cells(m, dc7) <> "" Then Exit For
Next m
If m > 19 Then [b14] = "7 / 11 ____________"Ñ¡A¨ä¾l¥¿±`"

For m = 2 To 20
If Cells(m, dc6) <> "" Then Exit For
Next m
If m > 19 Then [b14] = "6 / 11 _____"ªÑ§Q"

For m = 2 To 20
If Cells(m, dc1 + 1) <> "" Then Exit For
Next m
If m > 19 Then [b14] = "1 / 11 _____ "Ñ»ù "

For m = 2 To 20
If Cells(m, dc2) <> "" Then Exit For
Next m
If m > 19 Then [b14] = "2 / 11 _____________"³ø¡A§O¦A°Ý¤F"

    If [a2] = "5306" Then
    [a1] = "__" & " (" & "5306" & ")" & "_" & "____"°Ê²£«~"
    [a20] = "__"ù"
    End If

    ERRLOG "", err.Number
    'If InStr(Sheet1.Range("v47").Value, "__") > 0 Then _ _
    'Sheet1.Range("b14").Value = Sheet1.Range("b14").Value & "_" & Sheet1.Range("v47").Value
    
    With ActiveSheet.Cells
        .Font.Name = "____"úÅé"
        .Font.Name = "Arial"
        .Font.FontStyle = "__"Ç"
        .Font.Size = 10
        .RowHeight = 16
        .ColumnWidth = 7.5
    End With
    Range("a18:a19").Font.Size = 9
    
    Range(Cells(2, 2), Cells(93, 10)).ShrinkToFit = True '___Y¤p
    [f12].ShrinkToFit = False '____ÁY¤p
    [i14].ShrinkToFit = False '____ÁY¤p
    [j16].ShrinkToFit = False '____ÁY¤p
    [e17].ShrinkToFit = False '____ÁY¤p
    [g60].ShrinkToFit = False '____ÁY¤p
    [g80].ShrinkToFit = False '____ÁY¤p
    [g81].ShrinkToFit = False '____ÁY¤p
    Range(Cells(20, 11), Cells(25, 10)).ShrinkToFit = False '____ÁY¤p
    Range(Cells(72, 11), Cells(81, 10)).ShrinkToFit = False '____ÁY¤p
    
    Range(Columns(dc2), Columns(dc2 + 150)).Select '___Y¤p
        With Selection
        .WrapText = False
        .ShrinkToFit = True
    End With
    
    Range(Cells(1, dc2), Cells(1, dc2 + 150)).ShrinkToFit = False '____ÁY¤p
    Cells(50, dc1 + 1).ShrinkToFit = False '____ÁY¤p
    Range(Cells(26, dc1 + 2), Cells(27, dc1 + 2)).ShrinkToFit = False '____ÁY¤p
    Range(Cells(3, dc8 + 9), Cells(5, dc8 + 9)).ShrinkToFit = False '____ÁY¤p
    Cells(19, dc11).ShrinkToFit = False '____ÁY¤p
    Cells(41, dc8 + 9).ShrinkToFit = False '____ÁY¤p

    Cells(1, dc2).ColumnWidth = 20
    Cells(1, dc2 + 11).ColumnWidth = 20
    Cells(1, dc2 + 11 * 2).ColumnWidth = 20
    Cells(1, dc2 + 11 * 3).ColumnWidth = 20
    
    With Range("k1,c15:i16").Font
        .Size = 12
        .Name = "Arial"
    End With
    
    Range(Cells(18, dc1 + 2), Cells(19, dc1 + 2)).NumberFormatLocal = "yyyy/m/d;@"

    Range("c3:e12, b22:b29, e22:j29, b34:b42, e34:j42, j43:j43, b47:e54, b59:e67, h47:h58, j47:j58, b72:f79,h72:i79, b84:f93, h84:i93").Select
    Selection.NumberFormatLocal = "#,##0_);(#,##0)"
    
    Range("e17").HorizontalAlignment = xlRight
    Range("f17").HorizontalAlignment = xlLeft
    
    Range("B14").ShrinkToFit = False
    [y16].Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    Application.StatusBar = "__"¨"
    Application.Calculation = xlAutomatic           '_____}¼sºÖ
    
    If cd$ = "1723" Then
        [y38] = [i1]
        [y39] = [x39]
        [y40] = [w38]: [v39] = [q14]
    End If
    
End Sub


Public Sub OTC(cd$, dc)
    
    yr$ = [n15]
    Cells(1, dc) = "______ https://www.otc.org.tw/web/stock/statistics/monthly/st44.php?l=zh-tw"zh-tw"
   
    url = "https://www.tpex.org.tw//web/stock/statistics/monthly/st44.php?l=zh-tw"
     
   
           
       Set doc = New HTMLDocument
       
       
       With CreateObject("MSXML2.XMLHTTP")
           .Open "POST", url, False
           .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
           .setRequestHeader "Referer", url
           '.send "ajax=true&yy=" & yr$ & "&input_stock_code=" & cd$
           .send "yy=" & yr$ & "&input_stock_code=" & cd$
         
            doc.body.innerHTML = .responseText
               

            .abort
           
       End With
      
           
       Set rng = Cells(2, dc)
       
       For Each tbl In doc.getElementsByTagName("table")
            If tbl.className = "table table-bordered" Then
                 For Each r In tbl.Rows
                     For Each c In r.Cells
                          rng.value = c.innerText
                          Set rng = rng.Offset(, 1)
                          ii = ii + 1
                     Next c
                     Set rng = rng.Offset(1, -ii)
                     ii = 0
                 Next r
                 Exit For
           End If
       Next tbl
      Application.Wait Now() + TimeValue("00:00:01") * 2
       
   Cells(19, dc) = "______ https://www.otc.org.tw/web/stock/statistics/monthly/st42.php?l=zh-tw"zh-tw"
   
      url = "https://www.otc.org.tw/web/stock/statistics/monthly/st42.php?l=zh-tw"

      
      
      With CreateObject("MSXML2.XMLHTTP")
           .Open "GET", url, False
           .setRequestHeader "Connection", "keep-alive"
           .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
           .setRequestHeader "Referer", "https://www.tpex.org.tw//web/stock/statistics/monthly/st42.php?l=zh-tw"
      
           .send
           
            .abort
           
      End With
       
         Application.Wait Now() + TimeValue("00:00:01") * 1
      
          
      
       With CreateObject("MSXML2.XMLHTTP")
           .Open "POST", url, False
           .setRequestHeader "Connection", "keep-alive"
           .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
           .setRequestHeader "Referer", "https://www.tpex.org.tw//web/stock/statistics/monthly/st42.php?l=zh-tw"
       
    
           .send "input_stock_code=" & cd$
            doc.body.innerHTML = .responseText
           
            .abort
           
      End With
              
      Set rng = Cells(20, dc)
      
      For Each tbl In doc.getElementsByTagName("table")
            
          If tbl.className = "table table-bordered" Then
             
                For Each r In tbl.Rows
                    
                    For Each c In r.Cells
                         rng.value = c.innerText
                         Set rng = rng.Offset(, 1)
                         ii = ii + 1
                       
                    Next c
                    Set rng = rng.Offset(1, -ii)
                    ii = 0
                Next r
                Exit For
                
          End If
      Next tbl
      Set doc = Nothing
      
    
    

End Sub



Public Sub OTC_BAK(cd$, dc)
Dim otcp1 As Variant, otcp2 As Variant
yr$ = [n15]
    On Error Resume Next
    err.Clear
    Cells(1, dc) = "______ https://www.otc.org.tw/web/stock/statistics/monthly/st44.php?l=zh-tw"zh-tw"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.tpex.org.tw/web/stock/statistics/monthly/st44.php?l=zh-tw", Destination:=Cells(2, dc))
        .Name = "result_st44"
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SaveData = True
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "3"
        .PostText = "ajax=true&yy=" & yr$ & "&input_stock_code=" & cd$
        .Refresh BackgroundQuery:=False
        .ResultRange.UnMerge
        ERRLOG "________(" & .ResultRange.Address & ")", err.Numberr.Number
        ps1 = 1
        ps2 = .ResultRange.Rows.Count + 2
        If Not ps2 < 2 Then
            For p = ps2 To 1 Step -1
                If Cells(p, dc) = "_" Thenn
                    ps1 = p
                    psa = .ResultRange.Columns.Count
                    otcp1 = Cells(ps1, dc).Resize(ps2 - ps1, psa)
                    Exit For
                End If
            Next p
        End If
        .ResultRange.Clear
        .Delete
    End With
    
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d
    

    ERRLOG "", err.Number
    Cells(19, dc) = "______ https://www.otc.org.tw/web/stock/statistics/monthly/st42.php?l=zh-tw"zh-tw"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.tpex.org.tw/web/stock/statistics/monthly/st42.php?l=zh-tw", Destination:=Cells(20, dc))
        .Name = "result_st42"
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SaveData = True
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "3"
        .PostText = "ajax=true&input_stock_code=" & cd$
        .Refresh BackgroundQuery:=False
        .ResultRange.UnMerge
        ERRLOG "________(" & .ResultRange.Address & ")", err.Numberr.Number
        ps3 = 1
        ps4 = .ResultRange.Rows.Count + 20
        If Not ps4 < 2 Then
            For p = ps4 To 1 Step -1
                If Cells(p, dc) = "__" Thenen
                    ps3 = p
                    psb = .ResultRange.Columns.Count
                    otcp2 = Cells(ps3, dc).Resize(ps4 - ps3, psb)
                    Exit For
                End If
            Next p
        End If
        .ResultRange.Clear
        .Delete
    End With
    ERRLOG "", err.Number

    Cells(2, dc).Resize(ps2 - ps1, psa) = otcp1
    Cells(20, dc).Resize(ps4 - ps3, psb) = otcp2
    ERRLOG "", err.Number
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
   ' Delete_Pictures '__Alex_x§d
    
End Sub




Public Sub TWReport(url, dc)
       On Error Resume Next
       
       Set doc = New HTMLDocument
     
       With CreateObject("MSXML2.XMLHTTP")
           .Open "GET", url, False
           .send
            doc.body.innerHTML = .responseText
            .abort
     
       End With
     
     Cells(2, dc) = doc.getElementById("oScrollHead").innerText
     
     Cells(2, dc) = Replace(Cells(2, dc), "__ ____ __", "")ªí", "")
     
     ii = 0
     
     
     For Each tbl In doc.getElementsByTagName("table")
         
         If tbl.ID = "oMainTable" Then
                     nextrow = 4
                    
                     Set rng = Cells(nextrow, dc)
                     For Each L1 In tbl.getElementsByTagName("div")
                          
                          If L1.className = "table-row" Then
                           
                              For Each L2 In L1.getElementsByTagName("span")
                                   rng.value = Trim(L2.innerText)
                        
                                   Set rng = rng.Offset(, 1)
                                   ii = ii + 1 'column
                              Next L2
                              nextrow = nextrow + 1
                              Set rng = rng.Offset(1, -ii)
                              ii = 0
                          End If
                     
                     
                     Next L1
                                              
                Exit For
         
         End If
     
     
     Next tbl

  
      Set doc = Nothing
End Sub









Public Function TSECFSQT(url As String, dtrange As Range, wtable As String)
    On Error Resume Next
    ERRLOG "", err.Number
    
    
    
    With ActiveSheet.QueryTables.Add(Connection:="URL;" & url, Destination:=dtrange)
        .RefreshStyle = xlOverwriteCells
        .WebTables = wtable
        .Refresh BackgroundQuery:=False
        .ResultRange.UnMerge
        TSECFSQT = .ResultRange.Address
        .Delete
    End With
    'ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
 '   Delete_Pictures '__Alex_x§d
    
End Function

Public Function TSECFSQTF(url As String, wformat As String, wseltype As String, poststr As String, dtrange As Range, wtable As String)
    On Error Resume Next
    ERRLOG "", err.Number
    With ActiveSheet.QueryTables.Add(Connection:="URL;" & url, Destination:=dtrange)
        .RefreshStyle = xlOverwriteCells '******
        .WebFormatting = wformat
        .WebSelectionType = wseltype
        If Not wtable = "" Then .WebTables = wtable
        If Not poststr = "" Then .PostText = poststr
        .Refresh BackgroundQuery:=False
        TSECFSQTF = .ResultRange.Address
        .Delete
    End With
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________Ô´£¨Ñµ{¦¡
    Delete_Pictures '__Alex_x§d
End Function

Public Function ERRLOG(ST As String, EN As String)
    On Error Resume Next
    If Not ST = "" Then _
    ERRLOG = Replace("_" & ST, "$", "") & IIf(InStr(ST, ":") > 0, "", "____")±`")
    ERRLOG = ERRLOG & IIf(EN = "0", "", "____" & EN) EN)
    Sheet1.Range("v47").value = Sheet1.Range("v47").value & ERRLOG
    err.Clear
End Function

Public Function RSURLTYPE()
    Select Case (rnd() / rnd() * 10000 Mod 2)
        Case 0 '__¤@
            RSURLTYPE = "pscnetinvest.moneydj.com.tw"
        Case 1 '__¤@
            RSURLTYPE = "pscnetinvest.moneydj.com.tw"
        Case 2 '__»·
            RSURLTYPE = "just.honsec.com.tw"
    End Select
    
End Function

Public Function SEASONTF(dc As Integer)
    Dim i As Integer
    For i = dc To dc + 7
        If Cells(4, i) = "" Then GoTo xq
        Cells(1, dc + 7) = Left(Cells(4, i), Len(Cells(4, i)) - 3) + 0
        Cells(4, i) = Right(Cells(4, i), 2) + Right(Cells(1, dc + 7), 2)
xq: Next i
    Cells(1, dc + 7) = ""
End Function

Public Function YEARTF(dc As Integer)
    Dim i As Integer
    For i = dc To dc + 7
        If Cells(4, i) = "" Then GoTo xy
        Cells(4, i) = Cells(4, i) + 0
xy: Next i
End Function

Public Function SEASONCU(dc As Integer)
    '_______________6195÷§ï§ì«D¦X¨Ö6195
    On Error Resume Next
    SEASONCU = Year(Date) - (Val(Left(Cells(4, dc), Len(Cells(4, dc)) - 3)) + 1911) > 2
End Function

Public Function YEARCU(dc As Integer)
    '_______________6195÷§ï§ì«D¦X¨Ö6195
    On Error Resume Next
    YEARCU = Year(Date) - (Val(Cells(4, dc)) + 1911) > 2
End Function

Sub Macro8()
'
'  mikeon_  _ 2014/9/8 _____sªº¥¨¶°
'  __¤j

'
Application.ScreenUpdating = False
Range("A1:K16").Select: ActiveWindow.Zoom = True
[dz100].Select
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
Application.StatusBar = "__"¨"

End Sub

Sub Macro9()
'
' mikeon _ 2014/9/8 _____ªº¥¨¶°
' __¤p

'
Application.ScreenUpdating = False
Range("A1:R26").Select: ActiveWindow.Zoom = True
[dz100].Select
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
Application.StatusBar = "__"¨"
    
End Sub


















