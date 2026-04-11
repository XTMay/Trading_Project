Attribute VB_Name = "shared_modules"
Public theMarket As String
Public theIndustry As String
Public theMarketdate As String
Public Type baseInfo
    theCode As String
    theName As String
    Percent As Single
End Type
Public myWait As Boolean
Function findAll(theStr As String) As baseInfo
Dim theCode As String
Dim theName As String
Dim theArr
Dim c As Range
Dim theFlg As Boolean
Dim firstAddress As String
    '____, _________4___峖W稱至少4個字元
'[d14] = "___...___   "待   "
'Range("b18:L25,b30:L37,b59:f66,h59:j66,b70:f77,h70:j77").Select: Selection.ClearContents
'Application.ScreenUpdating = False      '________螢幕更新
    If LenB(theStr) < 4 Then
        Exit Function
    End If
    '_______________A_狳朣s碼工作表A欄
    With Worksheets("____").Columns(1)s(1)
        '__找
        Set c = .Find(theStr, LookIn:=xlValues, LookAt:=xlPart)
        '____?F嗎?
        If Not c Is Nothing Then
            '_________, ________________.與名稱中間以兩個空格分隔.
            '__: ________________1____1_______.間以1個空格及1個全形空格分隔.
            '          ____________________格全以半形空格取代了
            firstAddress = c.Address
            Do
                theArr = split(c.value, " ")
                '______票名稱
                theName = theArr(UBound(theArr))
                '______票代碼
                theCode = theArr(0)
                '_________________________?輸入的代碼或名稱是否符合?
                If theCode <> theStr And (Not theName Like theStr & "*") Then
                    Set c = .FindNext(c)
                Else
                    findAll.theCode = theCode
                    findAll.theName = theName
                    theIndustry = c.Offset(0, 3)
                    theFlg = True
                End If
             Loop While Not theFlg And c.Address <> firstAddress
        End If
    End With
    End Function
'_______威d編碼
Sub Stocklist()
Application.ScreenUpdating = False      '________螢幕更新
    With Worksheets("____")碼")
        .Columns("A:G").ClearContents
    End With
    With Worksheets("____").QueryTables.Add(Connection:= _:= _
        "URL;http://isin.twse.com.tw/isin/C_public.jsp?strMode=2", Destination:=Worksheets("____").Range("A1"))1"))
        .Name = "______"播s碼"
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SaveData = True
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "2"
        .Refresh BackgroundQuery:=False
    End With

    With Worksheets("____")碼")
        .[a1] = "_________"號及名稱"
        .[a1].QueryTable.Delete
        .Columns("F:G").ClearContents
    End With

    With Worksheets("____").Range("A:A"): a ")"
        Set a = .Find(What:="____(_)__", LookIn:=xlValues, LookAt:=xlPart)xlPart)
        Set b = .Find(What:="______", LookIn:=xlValues, LookAt:=xlPart)lPart)
        Set c = .Find(What:="____-_______", LookIn:=xlValues, LookAt:=xlPart)At:=xlPart)
        D = Worksheets("____").[a1].End(xlDown).Row.Row
        If Not a Is Nothing Then
            Worksheets("____").Range("A" & c.Row & ":E" & D).Delete Shift:=xlUpxlUp
            Worksheets("____").Range("A" & a.Row & ":E" & b.Row - 1).Delete Shift:=xlUpxlUp
        End If
        Set a = .Find(What:="______", LookIn:=xlValues, LookAt:=xlPart)lPart)
        b = Worksheets("____").[a1].End(xlDown).Row.Row
        With Worksheets("____")碼")
            .Range("E" & a.Row + 1) = "______"U憑證"
            .Range("E" & a.Row + 1).AutoFill Destination:=Worksheets("____").Range("E" & a.Row + 1 & ":E" & b), Type:=xlFillDefaultault
        End With
    End With
    
    a = Worksheets("____").[a1].End(xlDown).Row.Row
    b = 0
    For i = 1 To a
        If Not Worksheets("____").Cells(i, "A").Find("__ ", LookAt:=xlWhole) Is Nothing _hing _
        Or Not Worksheets("____").Cells(i, "A").Find("______(TDR) ", LookAt:=xlWhole) Is Nothing _ Nothing _
        Then
            Worksheets("____").Range("A" & i & ":E" & i).Delete Shift:=xlUpxlUp
            i = i - 1
            b = b + 1 'f__ㄕ
        End If
        If i > a - b Then Exit For 'i__f____h_Z行滌h出
    Next i
    
    a = Worksheets("____").Range("A1").End(xlDown).Row.Row
    With Worksheets("____").QueryTables.Add(Connection:= _:= _
        "URL;http://isin.twse.com.tw/isin/C_public.jsp?strMode=4", Destination:=Worksheets("____").Range("A" & a + 1)) 1))
        .Name = "______"播s碼"
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SaveData = True
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "2"
        .Refresh BackgroundQuery:=False
    End With
    
    With Worksheets("____")碼")
        .Range("A" & a + 1) = "_________"號及名稱"
        .Range("A" & a + 1).QueryTable.Delete
        .Columns("F:G").ClearContents
    End With

    With Worksheets("____").Range("A:A"): a ")"
        Set a = .Find(What:="____(_)__", LookIn:=xlValues, LookAt:=xlPart)xlPart)
        Set b = .Find(What:="__ ", LookIn:=xlValues, LookAt:=xlWhole)e)
        Set c = .Find(What:="______ ", LookIn:=xlValues, LookAt:=xlWhole)Whole)
        D = Worksheets("____").Range("A1").End(xlDown).Row.Row
        If Not a Is Nothing Then
            Worksheets("____").Range("E" & c.Row & ":E" & d) = "______"W存託憑證"
            Worksheets("____").Range("A" & a.Row & ":E" & b.Row).Delete Shift:=xlUpxlUp
        End If
        b = Worksheets("____").Range("A1").End(xlDown).Row.Row
        Set a = .Find(What:="____-______", LookIn:=xlValues, LookAt:=xlPart)t:=xlPart)
        If Not a Is Nothing Then
            Worksheets("____").Range("A" & a.Row & ":E" & b).Delete Shift:=xlUpxlUp
        End If
        
        a = Worksheets("____").Range("A1").End(xlDown).Row.Row
        b = 0
        For i = 1 To a
            If Not Worksheets("____").Cells(i, 1).Find("__ ", LookAt:=xlWhole) Is Nothing _hing _
            Or Not Worksheets("____").Cells(i, 1).Find("______ ", LookAt:=xlWhole) Is Nothing _ Nothing _
            Or Not Worksheets("____").Cells(i, 1).Find("____ ", LookAt:=xlWhole) Is Nothing Thening Then
                Worksheets("____").Range("A" & i & ":E" & i).Delete Shift:=xlUpxlUp
                i = i - 1
                b = b + 1 'f__ㄕ
            End If
            If i > a - b Then Exit For 'i__f____h_Z行滌h出
        Next i
        
        a = Worksheets("____").Range("A1").End(xlDown).Row.Row
        Worksheets("____").Range("C1:E" & a).Copy Destination:=Worksheets("____").Range("B1")ge("B1")
        .Columns("E").ClearContents
    End With
Application.ScreenUpdating = True       '______幕更新
End Sub
Sub enEvt()
Application.EnableEvents = True
End Sub

Sub Delete_Connection_XL2007()
    If ActiveWorkbook.Connections.Count > 0 Then
        For i = 1 To ActiveWorkbook.Connections.Count
            ActiveWorkbook.Connections.Item(1).Delete
        Next i
    End If
End Sub

Sub Delete_Connection_XL2007_2()
  Do Until ActiveWorkbook.Connections.Count = 0
        ActiveWorkbook.Connections(ActiveWorkbook.Connections.Count).Delete
    Loop
End Sub

Sub Delete_Connection_XL2003()

Dim qt As QueryTable
Dim oWS As Worksheet

For Each oWS In ActiveWorkbook.Sheets
    For Each qt In oWS.QueryTables
        qt.Delete
    Next qt
Next oWS

End Sub

Sub Delete_Connection_XL2003_2()
    If ActiveSheet.QueryTables.Count > 0 Then
        For i = 1 To ActiveSheet.QueryTables.Count
            ActiveSheet.QueryTables.Item(1).Delete
        Next i
    End If
End Sub

Sub ProtectSheet(mysheet As Worksheet)
'____Alex_ __吳 撰寫

   On Error Resume Next
     Dim cell As Range
     
        For Each cell In mysheet.Range("a1:ai1000")
            If cell.HasFormula Then cell.Locked = True
        Next cell
     
        If mysheet.Name = "__" Thenen
           Set unlockrange = mysheet.Range("e12, f17, o10:q10,k3,q14, k11")
        ElseIf mysheet.Name = "__" Thenen
             Set unlockrange = mysheet.Range("e9,f15, o5:q5,k3,q9, k11,bn1:bu1000")
        ElseIf mysheet.Name = "__" Thenen
             Set unlockrange = mysheet.Range("e9,f15,o2:q2,k3,q6, k11, cd1:ck1000")
        ElseIf mysheet.Name = "__" Thenen
             Set unlockrange = mysheet.Range("e9,f15,o2:q2,k3,q6, k11, bl1:bs1000")
        ElseIf mysheet.Name = "__" Thenen
             Set unlockrange = mysheet.Range("e9,f15,o2:q2,k3,q6, k11, bt1:ca1000")
        ElseIf mysheet.Name = "__" Thenen
             Set unlockrange = mysheet.Range("k1:as12")
        End If
        
        For Each cell In unlockrange
           cell.Locked = False
        Next cell
     
     mysheet.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
End Sub


Sub UnprotectSheet(mysheet As Worksheet)
'____Alex_ __吳 撰寫

     On Error Resume Next
     mysheet.Unprotect Password:=""
     Cells.Locked = False
End Sub

Sub Auto_open()

    Sheets("__").[zz1] = """"
    Sheets("__").[zz1] = """"


End Sub



