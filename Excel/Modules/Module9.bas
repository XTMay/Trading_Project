Attribute VB_Name = "Module9"
Sub tax()
'
' mikeon _ 23/5/7 _____的巨集
' __稅

'

Sheets("__").Selectct
Application.ScreenUpdating = True
[g3] = "___...".."
Range("g3").Font.Color = -16776961
Range("g3").HorizontalAlignment = xlCenter
        
Application.Wait Now() + TimeValue("00:00:02")

Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Application.Calculation = xlManual
Call UnprotectSheet(ActiveSheet)
Application.ScreenUpdating = False

Cells.Clear

Dim FileToOpen As Variant
Dim OpenBook As Workbook
FileToOpen = Application.GetOpenFilename(Title:="______")再表")
If FileToOpen <> False Then
Set OpenBook = Application.Workbooks.Open(FileToOpen)
OpenBook.Sheets("__").Activatete
Cells.Copy
    
ThisWorkbook.Worksheets("__").Activatete
  [a1].Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
       False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
       xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

OpenBook.Close False
End If

Call taxtb

[g3] = ""
Application.ScreenUpdating = True
[g3].Select
Application.Calculation = xlAutomatic           '_____}廣福
Application.StatusBar = "__""
Beep

End Sub


Sub exdiv()
'
' mikeon _ 23/5/1 _____的巨集
' __稅

'

Sheets("__").Selectct
Application.ScreenUpdating = True
[i3] = "___...".."
Range("i3").Font.Color = -16776961
Range("i3").HorizontalAlignment = xlCenter
        
Application.Wait Now() + TimeValue("00:00:02")

Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\秀雯
Application.Calculation = xlManual
Call UnprotectSheet(ActiveSheet)
Application.ScreenUpdating = False

m = 2: Do Until IsNumeric(Cells(m, 11)) And Cells(m, 11) <> ""
If m > 500 Then GoTo xx
m = m + 1
Loop
k = m

  Dim url As String, rsrange As String
  i = 1
  Do Until Cells(k, 11) = "" And Cells(k + 1, 11) = "" And Cells(k + 2, 11) = "" And Cells(k + 3, 11) = ""
  url = "http://pscnetinvest.moneydj.com.tw/z/zc/zci/zci_" & Cells(k, 11) & ".djhtm"
    rsrange = TSECFSQT(url, Cells(1, 21), "1")
    
    Application.StatusBar = "____ " & i & " _"" 支"
    Set c = Range(rsrange).Find("___", LookIn:=xlValues)es)
    If Not c Is Nothing Then
    End If
    Set c = Nothing
    ThisWorkbook.QueryTablesDelete ActiveSheet.Name '_________埭ㄗ捄{式
    Delete_Pictures '__Alex_x吳

For j = 12 To 19
If j = 16 Then j = j + 1
Cells(k, j) = ""
Next j
Cells(k, 12) = Left([v2], InStr(1, [v2], "(") - 1) '__名
Cells(k, 13) = Month([x4]) & "/" & Day([x4]) '___坐
Cells(k, 18) = Month([x11]) & "/" & Day([x11]) '___韙
If [x4] = "" Then
Cells(k, 13) = Month([aa4]) & "/" & Day([aa4])
Cells(k, 18) = Month([aa11]) & "/" & Day([aa11])
End If
Cells(k, 13).NumberFormat = "yyyy/m/d;@"
Cells(k, 18).NumberFormat = "yyyy/m/d;@"
If [x4] = "" And [aa4] = "" Then
Cells(k, 13) = ""
Cells(k, 18) = ""
End If
Cells(k, 15) = [aa10] '__資
If Cells(k, 15) <> "" Then Cells(k, 15) = [aa10] * 0.1

dc = 30
url = "http://pscnetinvest.moneydj.com.tw/z/zc/zcc/zcc_" & Cells(k, 11) & ".djhtm"
rsrange = TSECFSQT(url, Cells(2, dc), "3")
ERRLOG "6 / 11 __(" & rsrange & ")", err.Numberer
    
For i = 2 To 20
If Trim(Cells(i, dc + 3)) = "__" Then Exit Foror
Next i
Cells(k, 14) = [x10] '_/_股
If Year(Now()) = Cells(i + 1, dc) Then Cells(k, 14) = Cells(i + 1, dc + 3) '_/_股

Cells(k, 19) = Cells(k, 14) * Cells(k, 16) '__息

If Cells(k, 19) < [i10] Then
Cells(k, 17).FormulaR1C1 = "=(RC[-3]+RC[-2])*RC[-1]-R12C9"
End If

If Cells(k, 19) >= [i10] Then '______補充費
Cells(k, 17).FormulaR1C1 = "=RC[-3]*RC[-1]*(1-R11C9)+RC[-2]*RC[-1]-R12C9"

End If

If Cells(k, 13) = "" Or Now() < Cells(k, 13) Then
Cells(k, 17) = 0
Cells(k, 19) = 0
End If

k = k + 1: i = i + 1
Range(Columns(20), Columns(50)).Clear
Loop
n = k - 1

Range(Cells(n + 1, 11), Cells(n + 100, 19)).Clear

Cells(n + 2, 16) = "__"p"
Cells(n + 3, 16) = "KY"
Cells(n + 4, 16) = "_KY"""
Cells(n + 6, 16) = "KY__KY____"t減資"
Cells(n + 7, 16) = "(______)"繳稅)"

Cells(n + 3, 17).FormulaR1C1 = "="
Cells(n + 4, 17).FormulaR1C1 = "="
For j = m To n
If Not IsNumeric(Cells(j, 16)) Then Cells(j, 16) = ""
If Cells(j, 13) <> "" And Now() > Cells(j, 13) Then

If InStr(1, Cells(j, 12), "KY") > 0 And Cells(j, 19) < [i10] Then Cells(n + 3, 17).FormulaR1C1 = Cells(n + 3, 17).FormulaR1C1 & "+R[" & j - n - 3 & "]C[-3]*R[" & j - n - 3 & "]C[-1]-R12C9"
If InStr(1, Cells(j, 12), "KY") > 0 And Cells(j, 19) >= [i10] Then Cells(n + 3, 17).FormulaR1C1 = Cells(n + 3, 17).FormulaR1C1 & "+R[" & j - n - 3 & "]C[-3]*R[" & j - n - 3 & "]C[-1]*(1-R11C9)-R12C9"

If InStr(1, Cells(j, 12), "KY") = 0 And Cells(j, 19) < [i10] Then Cells(n + 4, 17).FormulaR1C1 = Cells(n + 4, 17).FormulaR1C1 & "+R[" & j - n - 4 & "]C[-3]*R[" & j - n - 4 & "]C[-1]-R12C9"
If InStr(1, Cells(j, 12), "KY") = 0 And Cells(j, 19) >= [i10] Then Cells(n + 4, 17).FormulaR1C1 = Cells(n + 4, 17).FormulaR1C1 & "+R[" & j - n - 4 & "]C[-3]*R[" & j - n - 4 & "]C[-1]*(1-R11C9)-R12C9" '______補充費
End If
Next j
Cells(n + 2, 17).FormulaR1C1 = "=SUM(R[" & m - n - 2 & "]C:R[-2]C)" '___賹
If Cells(n + 3, 17) = "=" Then Cells(n + 3, 17) = 0
If Cells(n + 4, 17) = "=" Then Cells(n + 4, 17) = 0

Columns("S:S").Clear

Columns("K:Q").Select
    Selection.ColumnWidth = 12.8
Columns("N:O").Select
    Selection.ColumnWidth = 9.13

Range(Cells(1, 11), Cells(n + 4, 17)).Select
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With

[m1] = "______________ [__] _______" [更新] 即補上其餘數字"

Cells(m - 1, 11) = "__""
Cells(m - 1, 12) = "__"W"
Cells(m - 1, 13) = "___"日"
Cells(m - 1, 14) = "_/_""
Cells(m - 1, 15) = "__""
Cells(m - 1, 16) = "__""
Cells(m - 1, 17) = "__"p"
Cells(m - 1, 18) = "___"日"
Cells(m - 1, 11).Select
    With Selection.Font
        .Color = -52429
        .TintAndShade = 0
    End With
    
xx:
Cells(m - 1, 16).Select
    With Selection.Font
       .Color = -52429
       .TintAndShade = 0
    End With

Range("K:O,R:R").Select
   With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Range(Columns(14), Columns(15)).Select
   Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
   With Selection
        .ShrinkToFit = True
End With

Range("P:Q, S:S").Select
   Selection.NumberFormatLocal = "#,##0_);[__](#,##0)")"
   With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
End With

Range(Cells(m - 1, 11), Cells(m - 1, 18)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

Range(Cells(n + 2, 16), Cells(n + 4, 16)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

Range(Cells(n + 2, 16), Cells(n + 4, 16)).Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.499984740745262
    End With
 
Range(Cells(n + 6, 16), Cells(n + 7, 16)).Select
With Selection.Font
        .Name = "Arial"
        .Size = 12
    End With
With Selection
        .ShrinkToFit = False
    End With
    
 Call taxtb
    
[i3] = ""
Application.ScreenUpdating = True
Cells(n + 2, 17).Select
Application.Calculation = xlAutomatic           '_____}廣福
Application.StatusBar = "____ " & n - m + 1 & " _" & "...__"..完成"
Beep

End Sub



Public Sub taxtb()

a = 1: c = 0
aa: b = 1
Do Until InStr(1, Cells(a, 1 + c), "__") > 0 0
If a > b + 500 Then GoTo cx
a = a + 1
Loop
b = a + 1

GL: Do Until InStr(1, Cells(b, 1 + c), "__") > 0 0
If InStr(1, Cells(b, 1 + c), "1042") > 0 Or b > a + 50 Then GoTo s1042
b = b + 1
Loop
Cells(b - 1, 4 + c).FormulaR1C1 = "=RC[-2]-RC[-1]"
Cells(b, 2 + c).FormulaR1C1 = "=R[-1]C*R" & a & "C" & 4 + c
Cells(b, 3 + c).FormulaR1C1 = "=R[-1]C*R" & a & "C" & 4 + c
Cells(b, 4 + c).FormulaR1C1 = "=R[-1]C*R" & a & "C" & 4 + c

b = b + 1
GoTo GL

s1042:
b = b + 1
ss:
b1 = b + 1
Do Until InStr(1, Cells(b, 1 + c), "__") > 0 0
If InStr(1, Cells(b, 1 + c), "KY") > 0 Or b > a + 50 Then GoTo KY
Cells(b, 4 + c).FormulaR1C1 = "=RC[-2]-RC[-1]"
b = b + 1
Loop

Cells(b, 4 + c).FormulaR1C1 = "=SUM(R[" & b1 - b - 1 & "]C:R[-1]C)*R" & a & "C" & 4 + c
b = b + 1
GoTo ss

KY:
Cells(b, 4 + c).FormulaR1C1 = "=RC[-2]"
Cells(b + 2, 4 + c).FormulaR1C1 = "=RC[-2]"
Cells(b + 3, 4 + c).FormulaR1C1 = "="

For t = a + 1 To b + 2
If InStr(1, Cells(t, 1 + c), "__") > 0 Or InStr(1, Cells(t, 1 + c), "KY") > 0 Thenen
Cells(b + 3, 4 + c).FormulaR1C1 = Cells(b + 3, 4 + c).FormulaR1C1 & "+R[" & t - b - 3 & "]C"
End If
Next t

a = b + 4
GoTo aa

cx:
If c = 5 Then GoTo xt
c = c + 5
a = 1
GoTo aa

xt:
With ActiveSheet.Cells
        .Font.Name = "____"體"
        .Font.Name = "Arial"
        .Font.FontStyle = "__""
        .Font.Size = 14
    End With

[i5] = "A-I___________"格填好之後"
[i6] = "_ [__] ______"算出結果"
[i7] = "_______________/__"A請自行複製/貼上"
Range("i5:i7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

[h9] = "_KY"""
[h10] = "______"坐j於"
[h11] = "_____"充費"
[h12] = "___"費"
Range("H9:H12").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

'[i10] = 20000
'[i11] = 0.0211
'[i12] = 10
Range("I10").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
Range("I11").NumberFormat = "0.00%"
Range("i10:i12").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    

End Sub



