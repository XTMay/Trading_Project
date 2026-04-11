Sub Macro30()
'
' mikeon _ 2018/4/6 _____™∫•®∂∞

'
'
' mikeon _ 2018/4/8 _____™∫•®∂∞
' __Æƒ

'
Sheets("__").Selectct

[g3] = "___...".."
[g3].Select
    With Selection.Font
        .Color = -16776961
    End With
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

    Application.MaxChange = 0.001
    With ActiveWorkbook
        .UpdateRemoteReferences = False
        .PrecisionAsDisplayed = False
    End With
Application.Wait Now() + TimeValue("00:00:01")
Application.Calculation = xlManual
Application.ScreenUpdating = False
Call UnprotectSheet(ActiveSheet)
Application.EnableCancelKey = xlInterrupt '_____\®q∂≤


dt: f = 10: Do Until Left(Cells(15, f), 1) <> "_" '(__R™Ì
f = f + 10
Loop
[d1] = f - 10

dfx = [d1] + 30 '__≤v
drr = dfx + 9 'XIRR
dbk = drr + 15 '____ÆÊßÔ

i = 11: Do Until UCase(Cells(1, i)) = "X" Or i > dfx
i = i + 10
Loop

If i < dfx Then
Columns(dbk + (i - 1) / 10).Delete Shift:=xlTokleft
Range(Columns(i - 1), Columns(i + 8)).Delete Shift:=xlTokleft
GoTo dt
End If '__)Ì)

Range(Columns(f), Columns(dbk)).Clear

f = 10: Do Until Left(Cells(15, f), 1) <> "_"""
Call pf(f) ' __™Ì
f = f + 10
Loop

Call pfn(f) ' __™Ì


f = 10: Do Until Left(Cells(15, f), 1) <> "_" ' (__[™Ì
f = f + 10
Loop

If Cells(16, f - 1) <> "" Then

Cells(1, f) = "_____X"kÆÊX"
Cells(15, f) = "___X"GX"
Cells(2, f + 1) = "__"·"
Cells(1, f + 2) = "__"º"
Cells(1, f + 3) = "__"˜"
Cells(1, f + 4) = "__": Cells(2, f + 4) = "JPY"Y"
Cells(2, f + 4).Select
    With Selection.Font
        .Color = -65536
        .TintAndShade = 0
    End With

Cells(1, f + 5) = "__"˜"
Cells(15, f + 1) = "__"¡"
Cells(14, f + 2) = "___": Cells(15, f + 2) = "XIRR"RR"
Cells(14, f + 3) = "___"ßQ"
Cells(14, f + 4) = "__"º"
Cells(14, f + 5) = "__"˜"
Cells(14, f + 6) = "__": Cells(15, f + 6) = "__"Ò®“"
Cells(14, f + 7) = "__"v"
Cells(14, f + 8) = "__+__"{™˜"
Cells(14, f + 9) = "__"˜"

Cells(2, f + 2).FormulaR1C1 = "=RC[2]"
Cells(2, f + 3).FormulaR1C1 = "=RC[1]"
Cells(2, f + 5).FormulaR1C1 = "=R[13]C[" & -6 - (f - 10) & "]"
Cells(15, f + 3).FormulaR1C1 = "=RC[" & -4 - (f - 10) & "]"
Cells(15, f + 4).FormulaR1C1 = "=RC[" & -5 - (f - 10) & "]"
Cells(15, f + 5).FormulaR1C1 = "=RC[" & -6 - (f - 10) & "]"
Cells(15, f + 7).FormulaR1C1 = "=R[-13]C[-3]"
Cells(15, f + 8).FormulaR1C1 = "=RC[" & -9 - (f - 10) & "]"
Cells(15, f + 9).FormulaR1C1 = "=RC[" & -10 - (f - 10) & "]"

Cells(1, f + 1) = "???"
Cells(1, f + 1).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Cells(14, f + 1).FormulaR1C1 = "=R[-13]C"
Cells(14, f + 1).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Cells(3, f + 1) = "A"
Cells(3, f + 1).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Cells(3, f + 2) = 0
Cells(3, f + 3) = 0
Range(Cells(3, f + 2), Cells(3, f + 3)).Select
    With Selection.Font
        .Color = -16776961
End With

Cells(3, f + 4).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    
Cells(3, f + 5) = 0
Cells(3, f + 5).Select
    With Selection.Font
        .Color = -16776961
    End With

End If ' __)Ì)

a = 11: Do Until Cells(1, a) = ""
If Left(Cells(1, a), 1) = "_" Or Left(Cells(1, a), 2) = "TW" Thenn
For b = 3 To 13
If InStr(1, Cells(b, a), "__") > 0 Then Cells(b, a) = """"
Next b

For b = 3 To 13
If Cells(b, a + 1) = "" Then Exit For
Next b
If b > 11 Then b = 11
Cells(b + 2, a) = "_________________(________)_________"µ|≠∂∞£Æß™Ì¡`≠p)§~¨O•øΩT™∫≤{™˜æl√B"
Cells(b + 2, a).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End If
    
a = a + 10
Loop

Cells(3, 7) = ""
[d1] = ""

i = 16: Do Until Cells(i, 3) = ""
i = i + 1
Loop
Cells(i - 1, 3).Select
Application.StatusBar = "__"®"
Application.Calculation = xlAutomatic
Beep

End Sub


Public Sub pf(f)

dfx = [d1] + 30 '__≤v
drr = dfx + 9 'XIRR
dbk = drr + 15 '____ÆÊßÔ


f1 = f
If Cells(3, f + 5) = 0 Or Cells(3, f + 5) = "" Then GoTo xx
Application.StatusBar = Cells(1, f + 1) & " - " & f / 10 & " / " & [d1] / 10


Call ddlet(f)


GoTo HL
For k = 1 To [d1] / 10
 Cells(5, dbk + k) = 1
 If [i15] <> Cells(2, 10 * k + 4) And Cells(3, 10 * k + 5) > 0 Then
 url = "https://finance.yahoo.com/quote/" & Cells(2, 10 * k + 4) & [i15] & "=x?ltr=1" '__Alex_xßd
        Cells(1, dfx + (k - 1) * 4) = "__ " & urlrl
        Call ConnectMarketWatch(url, Cells(2, dfx + (k - 1) * 4), 2)

        For i = 2 To 20
        If Cells(i, dfx + (k - 1) * 4) = "Previous Close" And Cells(i, dfx + (k - 1) * 4 + 1) <> "" And IsNumeric(Cells(i, dfx + (k - 1) * 4 + 1)) Then
        Cells(5, dbk + k) = Cells(i, dfx + (k - 1) * 4 + 1)
        Exit For
        End If
        If Cells(i, dfx + (k - 1) * 4) = "Open" And Cells(i, dfx + (k - 1) * 4 + 1) <> "" And IsNumeric(Cells(i, dfx + (k - 1) * 4 + 1)) Then Cells(5, dbk + k) = Cells(i, dfx + (k - 1) * 4 + 1)
        Next i
End If
Next k
    
'HL: Cells(5, dbk + f / 10) = 1
 If [i15] <> Cells(2, 10 * f / 10 + 4) And Cells(3, 10 * f / 10 + 5) > 0 Then
 k = 1: Do Until Cells(k, dfx) = ""
 If UCase(Trim(Cells(2, 10 * f / 10 + 4))) = UCase(Right(Cells(k, dfx), 3)) Then GoTo fxgot
 k = k + 1
 Loop
 
ffxx:
 'url = "https://transferwise.com/zh-hk/currency-converter/" & Cells(2, 10 * f / 10 + 4) & "-to-" & [i15] & "-rate" '_____yahooyahoo
 url = "https://wise.com/zh-hk/currency-converter/" & Cells(2, 10 * f / 10 + 4) & "-to-" & [i15] & "-rate" '_____yahooyahoo
    
    Cells(1, dfx + 5) = "__ " & urlrl
               
         Call ConnectWinHttp(url, 1)
         
         
           
                            
        If InStr(1, doc.body.innerHTML, "HTTP ERROR") >= 1 Then
        
            Debug.Print "No data"
        
         Else
                                   
        
                For Each kk In doc.getElementsByTagName("h3")
                    If kk.className = "cc__source-to-target" Then
                                           
                       ' Debug.Print kk.innerText
                                           
                        myfx = split(kk.innerText, "=")


                        fx1 = myfx(0)
                        fx2 = myfx(1)

                        For nn = 1 To Len(fx1)

                            If IsNumeric(Mid(fx1, nn, 1)) = True Then
                                Cells(k, dfx) = Trim(Mid(fx1, nn, Len(fx1) - nn + 1))

                               Exit For
                            End If

                        Next nn


                        For nn = 1 To Len(fx2)

                            If IsNumeric(Mid(fx2, nn, 1)) = True Then
                                Cells(k, dfx + 1) = Trim(Mid(fx2, nn, Len(fx2) - nn + 1))

                               Exit For
                            End If

                        Next nn
                        
                        
                        
'
'                        If InStr(1, myfx(0), "$") >= 1 Then
'
'                           Cells(k, dfx) = Split(myfx(0), "$")(1)
'                        Else
'                           Cells(k, dfx) = myfx(0)
'                        End If
'
'
'                        If InStr(1, myfx(1), "$") >= 1 Then
'                             Cells(k, dfx + 1) = Split(myfx(1), "$")(1)
'
'                        Else
'
'                         Cells(k, dfx + 1) = myfx(1)
'
'                        End If
'
'                        Cells(k, dfx + 1) = Split(myfx(1), "$")(1)
                          
                        Exit For
                        
                    End If
                Next kk
                

                Cells(k, dfx + 3) = Date

        End If

            

'Debug.Print "dbk &"; dbk, "k:" & k, "dfx:" & dfx

fxgot: Cells(5, dbk + f / 10) = Left(Cells(k, dfx + 1), Len(Cells(k, dfx + 1)) - 4) / Left(Cells(k, dfx), Len(Cells(k, dfx)) - 4)
End If
HL: Range(Columns(drr), Columns(dbk)).Clear


 i = 16
 Do Until Cells(i, f + 3) = ""
 i = i + 1
 Loop
 
 If Cells(i - 1, f + 1) = Date Then i = i - 1
     
     
 Cells(i, f + 1) = Date
 If Cells(i, f + 7) = "" Then
 Cells(i, f + 7) = Cells(5, dbk + f / 10)  '__≤v
 
 Else
 Cells(i - 9, dbk + f / 10) = Cells(i, f + 7)
 End If
 
 Cells(i, f + 4) = 0 '__≤º
 j = 3: Do Until Cells(j, f + 2) = "" Or j > 13
 Cells(i, f + 4) = Cells(i, f + 4) + Cells(j, f + 2)
 j = j + 1
 Loop
 Cells(i, f + 4) = Cells(i, f + 4) * Cells(i, f + 7)
 Cells(1, dbk + f / 10) = Cells(i, f + 4)

 If Cells(16, f + 4) = 0 Then
 Cells(18, f + 1) = "__________ 0"B§£±o¨∞ 0"
 If Cells(2, f + 4) = "USD" Then Cells(18, f + 1) = "The amount of the first stock cannot be zero."
 GoTo xx
 End If
 
 Cells(i, f + 5) = 0 '__™˜
 j = 3: Do Until Cells(j, f + 3) = "" Or j > 13
 Cells(i, f + 5) = Cells(i, f + 5) + Cells(j, f + 3)
 j = j + 1
 Loop
 Cells(i, f + 5) = Cells(i, f + 5) * Cells(i, f + 7)
 
 Cells(2, dbk + f / 10) = Cells(i, f + 5)
 
 Cells(i, f + 9) = Cells(3, f + 5) '__™˜
 Cells(3, dbk + f / 10) = Cells(i, f + 9)

'-----------------------------
k = 1: Do Until k > [d1] / 10
If Cells(7, dbk + k) <> "" Then Exit Do
k = k + 1
Loop

If k > [d1] / 10 Then
For j = 1 To [d1] / 10
i = 16: Do Until Cells(i, 9 + 10 * j) = ""
Cells(i - 9, dbk + j) = Cells(i, 7 + 10 * j)
i = i + 1
Loop

Do Until Cells(i - 9, dbk + j) = ""
Cells(i - 9, dbk + j) = ""
i = i + 1
Loop
Next j
End If
'----------------------------


For k = 1 To i + 1

If Cells(k + 15, f + 1) = "" Then Exit For

Cells(k + 1, drr + 1) = Cells(k + 15, f + 1) '__¥¡
If Cells(k + 6, dbk + f / 10) = "" Then Cells(k + 6, dbk + f / 10) = Cells(k + 15, f + 7)
Cells(k + 1, drr + 2) = Cells(k + 15, f + 4) * Cells(k + 15, f + 7) / Cells(k + 6, dbk + f / 10)   '__≤º
Cells(k + 15, f + 4) = Cells(k + 15, f + 4) * Cells(k + 15, f + 7) / Cells(k + 6, dbk + f / 10)
Cells(k + 1, drr + 3) = Cells(k + 15, f + 5) * Cells(k + 15, f + 7) / Cells(k + 6, dbk + f / 10)  '__™˜
Cells(k + 15, f + 5) = Cells(k + 15, f + 5) * Cells(k + 15, f + 7) / Cells(k + 6, dbk + f / 10)
Cells(k + 1, drr + 4) = Cells(k + 15, f + 9) '__™˜
Cells(k + 15, f + 9) = Cells(k + 15, f + 9)
If k > 1 Then  '__/____™˜ºW√B
Cells(k + 1, drr + 5) = Cells(k + 1, drr + 3) - Cells(k, drr + 3)
Cells(k + 1, drr + 6) = Cells(k + 1, drr + 4) - Cells(k, drr + 4)
Else
Cells(k + 1, drr + 5) = Cells(k + 1, drr + 3)
Cells(k + 1, drr + 6) = Cells(k + 1, drr + 4)
End If
Cells(k + 1, drr + 8) = Cells(k + 1, drr + 5) - Cells(k + 1, drr + 6) '____ºW√B
Cells(k + 1, drr + 7) = Cells(k + 1, drr + 8) + Cells(k + 1, drr + 2)

'If Cells(k + 15, f + 9) <= 0 Then GoTo kkk9
Cells(k + 1, drr).FormulaR1C1 = "=IFERROR(XIRR(R2C" & 86 + [d1] - 40 & ":RC[7],R2C" & 80 + [d1] - 40 & ":RC[1]),0)"
Cells(k + 15, f + 2) = Cells(k + 1, drr)
Cells(k + 1, drr + 7) = Cells(k + 1, drr + 8)
Cells(k + 1, drr + 8) = ""
Cells(k + 15, f + 3) = Cells(k + 15, f + 4) + Cells(k + 15, f + 5) - Cells(k + 15, f + 9) '___ÚßQ
Cells(4, dbk + f / 10) = Cells(k + 15, f + 3)
Cells(k + 15, f + 6) = Cells(k + 15, f + 4) / (Cells(k + 15, f + 4) + Cells(k + 15, f + 5)) '____§Ò®“
Cells(k + 15, f + 8) = Cells(k + 15, f + 4) + Cells(k + 15, f + 5)
'kkk9:
Next k

k = 1: Cells(k + 15, f + 2) = Cells(k + 15, f + 3) / Cells(k + 15, f + 9)
k = 2: Do Until Cells(k + 15, f + 1) = ""
If Cells(k + 15, f + 9) <= 0 Then GoTo k9
If DateDiff("d", Cells(16, f + 1), Cells(k + 15, f + 1)) < 365 Then Cells(k + 15, f + 2) = (1 + Cells(k + 15, f + 2)) ^ (DateDiff("d", Cells(16, f + 1), Cells(k + 15, f + 1)) / 365) - 1
k9: k = k + 1
Loop

Call cpbf(f)

For k = 1 To 4
If k = 1 Then m = 4 '__≤º
If k = 2 Then m = 5 '__™˜
If k = 3 Then m = 9 '__™˜
If k = 4 Then m = 3 '___ÚßQ
i = 16: Do Until Cells(i, m + f) = "" And IsNumeric(Cells(i - 1, m + f))
i = i + 1
Loop
Cells(k, dbk + f / 10) = Cells(i - 1, f + m)
If Cells(k, dbk + f / 10) = "" Then Cells(k, dbk + f / 10) = 0
Next k

Call fm8(f)
xx:
f = f1



End Sub


Public Sub ddlet(f)

dfx = [d1] + 30 '__≤v
drr = dfx + 9 'XIRR
dbk = drr + 15 '____ÆÊßÔ

i = 16
Do Until Cells(i, f + 3) = ""
i = i + 1
Loop

j = 1
Do Until Cells(i + j, f + 1) = ""
Cells(i + j, f + 1) = ""
j = j + 1
Loop

i = 16: p = 15: q = 0
Do Until Cells(i, f + 3) = ""
p = p + 1 '_____˝¥X§—
If UCase(Cells(i, f)) = "X" Then
q = 1

If f = 1 Then
r = f + 8
Else
r = f + 9
End If

For k = f To r
Cells(i, k) = ""
Next k
End If
i = i + 1
Loop

If q = 1 Then '__ßR
i = 16
Do Until Cells(i, f + 3) = "" '_______ iQßR≥B i
i = i + 1
Loop

j = i + 1
Do Until j > p
If Cells(j, f + 1) = "" Then GoTo jj
For k = f + 1 To r
Cells(i, k) = Cells(j, k)
Next k
i = i + 1
jj: j = j + 1
Loop

Do Until i > p '___hæl
For k = f + 1 To r
Cells(i, k) = ""
Next k
i = i + 1
Loop

'--------------------------
If f > 1 Then
For j = 1 To [d1] / 10
i = 16: Do Until Cells(i, 9 + 10 * j) = ""
Cells(i - 9, dbk + j) = Cells(i, 7 + 10 * j)
i = i + 1
Loop

Do Until Cells(i - 9, dbk + j) = ""
Cells(i - 9, dbk + j) = ""
i = i + 1
Loop
Next j
End If
'------------------

End If

If f > 1 Then
i = 16: Do Until Cells(i, f + 9) = ""
 i = i + 1
 Loop
 j = i
 Do Until j > p + 1
 Cells(j, f + 7) = "": Cells(j, f + 8) = ""
 j = j + 1
 Loop
End If


If f = 1 Then
j = 1
Do Until Cells(i + j, f + 1) = ""
Cells(i + j, f + 1) = ""
j = j + 1
Loop

GoTo tt
End If


If Cells(3, f + 5) = 0 Or Cells(3, f + 5) = "" Then
j = 1
Do Until Cells(i + j, f + 1) = ""
Cells(i + j, f + 1) = ""
j = j + 1
Loop

End If
tt:

End Sub



Public Sub pfn(f) ' ____¡`™Ì

f = 1
dfx = [d1] + 30 '__≤v
drr = dfx + 9 'XIRR
dbk = drr + 15 '____ÆÊßÔ

Application.StatusBar = "__ - " & [d1] / 10 & " / " & [d1] / 1010

Call ddlet(f)

 i = 16
 Do Until Cells(i, f + 3) = ""
 i = i + 1
 Loop

j = 1
Do Until Cells(i + j, f + 1) = ""
Cells(i + j, f + 1) = ""
j = j + 1
Loop

If Cells(i - 1, f + 1) = Date Then i = i - 1
 
Columns("drr:dbk").Clear
    
 Cells(i, 2) = Date '__™Ì
 Cells(i, 5) = 0
 Cells(i, 6) = 0
 Cells(i, 9) = 0
 
 For j = 1 To [d1] / 10
 Cells(i, 5) = Cells(i, 5) + Cells(1, dbk + j) '__≤º
 Cells(i, 6) = Cells(i, 6) + Cells(2, dbk + j) '__™˜
 Cells(i, 9) = Cells(i, 9) + Cells(3, dbk + j) '__™˜
 Next j
 
For k = 1 To i + 1

If Cells(k + 15, f + 1) = "" Then Exit For
Cells(k + 1, drr + 1) = Cells(k + 15, f + 1) '__¥¡
Cells(k + 1, drr + 2) = Cells(k + 15, f + 4) '__≤º
Cells(k + 1, drr + 3) = Cells(k + 15, f + 5) '__™˜
Cells(k + 1, drr + 4) = Cells(k + 15, f + 8) '__™˜

If k > 1 Then  '__/____™˜ºW√B
Cells(k + 1, drr + 5) = Cells(k + 1, drr + 3) - Cells(k, drr + 3)
Cells(k + 1, drr + 6) = Cells(k + 1, drr + 4) - Cells(k, drr + 4)
Else
Cells(k + 1, drr + 5) = Cells(k + 1, drr + 3)
Cells(k + 1, drr + 6) = Cells(k + 1, drr + 4)
End If
Cells(k + 1, drr + 8) = Cells(k + 1, drr + 5) - Cells(k + 1, drr + 6) '____ºW√B
Cells(k + 1, drr + 7) = Cells(k + 1, drr + 8) + Cells(k + 1, drr + 2)

'If Cells(k + 15, f + 8) <= 0 Then GoTo kkk8
Cells(k + 1, drr).FormulaR1C1 = "=IFERROR(XIRR(R2C" & 86 + [d1] - 40 & ":RC[7],R2C" & 80 + [d1] - 40 & ":RC[1]),0)"
Cells(k + 15, f + 2) = Cells(k + 1, drr)
Cells(k + 1, drr + 7) = Cells(k + 1, drr + 8)
Cells(k + 1, drr + 8) = ""
Cells(k + 15, f + 3) = Cells(k + 15, f + 4) + Cells(k + 15, f + 5) - Cells(k + 15, f + 8) '___ÚßQ
Cells(k + 15, f + 6) = Cells(k + 15, f + 4) / (Cells(k + 15, f + 4) + Cells(k + 15, f + 5)) '____§Ò®“
Cells(k + 15, 8) = Cells(k + 15, 5) + Cells(k + 15, 6)
'kkk8:
Next k

k = 1: Cells(k + 15, f + 2) = Cells(k + 15, f + 3) / Cells(k + 15, f + 8)
k = 2: Do Until Cells(k + 15, f + 1) = ""
If Cells(k + 15, f + 8) <= 0 Then GoTo k8
If DateDiff("d", Cells(16, f + 1), Cells(k + 15, f + 1)) < 365 Then Cells(k + 15, f + 2) = (1 + Cells(k + 15, f + 2)) ^ (DateDiff("d", Cells(16, f + 1), Cells(k + 15, f + 1)) / 365) - 1
k8: k = k + 1
Loop

Cells(k, drr + 7) = Cells(k, drr + 7) + Cells(k, drr + 2)
Cells(k, drr + 8) = "____+__"+™—≤º"


'--------------------------
For j = 1 To [d1] / 10
i = 16: Do Until Cells(i, 9 + 10 * j) = ""
Cells(i - 9, dbk + j) = Cells(i, 7 + 10 * j)
i = i + 1
Loop

Do Until Cells(i - 9, dbk + j) = ""
Cells(i - 9, dbk + j) = ""
i = i + 1
Loop
Next j
'------------------



xx: Call cpbf(f)
Call fm8(f)

End Sub



Private Sub cpbf(f)

dfx = [d1] + 30 '__≤v
drr = dfx + 9 'XIRR
dbk = drr + 15 '____ÆÊßÔ

If UCase(Trim(Cells(2, f + 4))) = UCase(Trim([i15])) Then GoTo xfx
If f = 1 Then GoTo xfx
 i = 16: Do Until Cells(i, f + 9) = ""
 i = i + 1
 Loop
 
 If i = 16 Then GoTo xfx
 Cells(i, f + 7) = Cells(16, f + 7) * Cells(16, f + 9) / Cells(i - 1, f + 9)
 j = 17: Do Until Cells(j, f + 9) = ""
 Cells(i, f + 7) = Cells(i, f + 7) + Cells(j, f + 7) * (Cells(j, f + 9) - Cells(j - 1, f + 9)) / Cells(i - 1, f + 9)
 j = j + 1
 Loop
 Cells(i, f + 8) = ": ______"°∂◊≤v"
 If Cells(2, f + 4) = "USD" Then Cells(i, f + 8) = ": Avg Fx rate of Principal"

xfx:
If f > 3 And Cells(3, f + 5) = 0 Then GoTo xbf

i = 16
Do Until Cells(i, f + 3) = ""
i = i + 1
Loop
 
If Cells(i - 1, f + 2) > 0 Then
Cells(i + 1, f + 5) = DateDiff("d", Cells(16, f + 1), Cells(i - 1, f + 1)) / 365
If Cells(i + 1, f + 5) < 1 Then Cells(i + 1, f + 5) = 1
[e10] = Cells(i + 1, f + 5)
[f10] = Cells(i - 1, f + 2)

Range("D5").Select
    ActiveCell.FormulaR1C1 = "=(1.2^(50-R[5]C[1]))^0.5"
Range("D6").Select
    ActiveCell.FormulaR1C1 = "=LOG(9100,1+R[4]C[2])"

If Cells(2, f + 4) <> "USD" Then

Range("G10").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<0,""__"",100*(1.2^(50-RC[-2]))^0.5*LOG(9100,1+RC[-1])/((1.2^(50-8))^0.5*LOG(9100,1+0.12)))")"
Range("H10").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]<0,""_____"",IF(RC[-1]>168,""_____"","""")&IF(AND(RC[-1]<=168,RC[-1]>100),""__"","""")&IF(AND(RC[-1]<=100,RC[-1]>21),""____"","""")&IF(RC[-1]<=21,""_____ !"",""""))"B¨¸§⁄µ·ØS !"",""""))"


Cells(i + 1, f + 1) = "_____ 50 ___ 20% = 1.2^50 = 9,100"= 9,100"
Cells(i + 2, f + 1) = "_ " & Format(Cells(i + 1, f + 5), "##,##0.0") & " ___ " & Format(Cells(i - 1, f + 2), "##,##0%") & " ______ " & Format([g10], "##,##0") & " (____)_" & [h10]V¶n)°A" & [h10]

Cells(i + 3, f + 1) = Format(Cells(i + 1, f + 5), "##,##0.0") & " ____A = 1.2^(50-" & Format(Cells(i + 1, f + 5), "##,##0.0") & ")_A^0.5 = " & Format([d5], "##,##0")##0")
Cells(i + 4, f + 1) = Format(Cells(i - 1, f + 2), "##,##0%") & " ___B (1+" & Format(Cells(i - 1, f + 2), "##,##0%") & ")^B = 9,100_B = " & Format([d6], "##,##0") & "_" "¶~"
Cells(i + 5, f + 1) = "______ " & Format([g10], "##,##0") & " = 100x(A^0.5xB)/3,701"3,701"
Cells(i + 6, f + 1) = "    __ 3,701 = (1.2^(50-8))^0.5xlog(9100,1+12%)_8___12%"ß°12%"
Else

Range("G10").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<0,""LOSS"",100*(1.2^(50-RC[-2]))^0.5*LOG(9100,1+RC[-1]))/((1.2^(50-8))^0.5*LOG(9100,1+0.12))"
 Range("H10").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]<0,""Review lecture transcript frequently"",IF(RC[-1]>168,""Review lecture transcript frequently"","""")&IF(AND(RC[-1]<=168,RC[-1]>100),""Go for it !"","""")&IF(AND(RC[-1]<=100,RC[-1]>21),""Excellent !"","""")&IF(RC[-1]<=21,""Comparable to Buffett !"",""""))"
Cells(i + 1, f + 1) = "Mr.Buffett's performance is a 50-year average of 20% = 1.2^50 = 9,100"
Cells(i + 2, f + 1) = "Your " & Format(Cells(i + 1, f + 5), "##,##0.0") & "-year average of " & Format(Cells(i - 1, f + 2), "##,##0%") & " distance from Buffett " & Format([g10], "##,##0") & " (The smaller the better) " & [h10]

Cells(i + 3, f + 1) = Format(Cells(i + 1, f + 5), "##,##0.0") & " years gap A = 1.2^(50-" & Format(Cells(i + 1, f + 5), "##,##0.0") & ")_A^0.5 = " & Format([d5], "##,##0"))
Cells(i + 4, f + 1) = Format(Cells(i - 1, f + 2), "##,##0%") & " gap B (1+" & Format(Cells(i - 1, f + 2), "##,##0%") & ")^B = 9,100_B = " & Format([d6], "##,##0") & " yr"""
Cells(i + 5, f + 1) = "Distance from Buffett " & Format([g10], "##,##0") & " = 100x(A^0.5xB)/3,701"
Cells(i + 6, f + 1) = "    where 3,701 = (1.2^(50-8)^0.5)xlog(9100,1+12%), 8-year average of 12%"
End If

Cells(i + 1, f + 5) = ""
Cells(i + 2, f + 5) = ""
End If

xbf:

[d5] = "": [d6] = ""
End Sub



Public Sub fm8(f)
    
dfx = [d1] + 30 '__≤v
drr = dfx + 9 'XIRR
dbk = drr + 15 '____ÆÊßÔ

Range(Columns(f), Columns(f + 1)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Columns(f + 1).NumberFormatLocal = "yyyy/m/d;@"

For i = 1 To 2
If i = 1 Then m = f + 2
If i = 2 Then m = f + 6
Columns(m).Select
    Selection.NumberFormatLocal = "0%"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With
Next i
    
For i = 1 To 5
If i = 1 Then m = f + 3
If i = 2 Then m = f + 4
If i = 3 Then m = f + 5
If i = 4 Then m = f + 8
If i = 5 Then m = f + 9
Columns(m).Select
Selection.NumberFormatLocal = "#,##0_);(#,##0)"
With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
End With
Next i

Range(Cells(3, f + 2), Cells(13, f + 5)).Select
Selection.NumberFormatLocal = "#,##0_);(#,##0)"
With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
End With

If f > 1 Then

If Cells(3, f + 5) > 0 Then
If UCase(Cells(2, f + 4)) = "USD" Then GoTo UUSD
i = 3: Do Until Cells(i, f + 2) = ""
Range(Cells(i, f + 2), Cells(i, f + 3)).Select
    With Selection.Font
        .Color = -16776961
    End With

Cells(i, f + 4).FormulaR1C1 = "=RC[-2]+RC[-1]"
Cells(i, f + 4).Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

i = i + 1
Loop
GoTo usfm

UUSD:
i = 3: Do Until Cells(i, f + 3) = ""
Range(Cells(i, f + 3), Cells(i, f + 4)).Select
    With Selection.Font
        .Color = -16776961
    End With

Cells(i, f + 2).FormulaR1C1 = "=RC[2]-RC[1]"
Cells(i, f + 2).Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
i = i + 1
Loop
End If

If Cells(1, f + 2) <> "Stock" Then
Cells(2, f + 1) = "Account"
Cells(1, f + 2) = "Stock"
Cells(1, f + 3) = "Cash"
Cells(1, f + 4) = "Sum"
Cells(1, f + 5) = "Principal"

Cells(15, f + 1) = "Date"
Cells(14, f + 2) = "Annual"
Cells(14, f + 3) = "T Profit"
Cells(14, f + 4) = "Stock"
Cells(14, f + 5) = "Cash"
Cells(14, f + 6) = "Holding": Cells(15, f + 6) = "Ratio"
Cells(14, f + 7) = "Fx"
Cells(14, f + 8) = "Stock+Cash"
Cells(14, f + 9) = "Principal"
End If
usfm:

Columns(f + 7).Select
Selection.NumberFormatLocal = "#,##0.00_);(#,##0.00)"
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With
End If

If f = 1 Then

Cells.Select
Selection.ColumnWidth = 10
With Selection.Font
    .Name = "____"˙≈È"
    .Name = "Arial"
    .FontStyle = "__"«"
    .Size = 10
End With
    With Selection
        .ShrinkToFit = True
End With

[a1] = "1. ___________x_16_______"±N16¶C©≥§UΩd®“ßR∞£"
[a2] = "2. ___________ [__]"A¶A´ˆ [ßÛ∑s]"
[a3] = "3. ____ " & [i15] & " ___ i15_________A70"•NΩX∞—æ\¨¸™—A70"
[a4] = "4. XIRR _________________ 1 _ (1+XIRR)^(__/365)-1"+XIRR)^(§—º∆/365)-1"
[a5] = "5. __ [__] [__] [__] [__] [__] _______ [__] __"ßÛßÔ°A¶A´ˆ [ßÛ∑s] ßY•i"
[a6] = "6. K _________________"•[§@±i°A±iº∆µL≠≠"
[a7] = "7. [__]_________________"§~ª›ºW¥Ó°A•≠Æ…§£≈‹"
[a8] = "    [__]_[__]_______"b§W≈„•‹∂Ò§J"
[a10] = "________ B"·´¸æ… B"

[e9] = "_"""
[f9] = "____"ZÆƒ"
[g9] = "______ (____)"∂V§p∂V¶n)"

Rows("1:2").Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Rows("14:15").Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Range("a1:a11").Select
With Selection
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = False
End With

i = 16: Do Until Cells(i, f + 3) = ""
i = i + 1
Loop
Range(Cells(i + 1, f + 1), Cells(i + 7, f + 1)).Select
    With Selection
         .HorizontalAlignment = xlLeft
        .ShrinkToFit = False
End With

For k = 10 To [d1] Step 10
i = 16: Do Until Cells(i, k + 3) = ""
i = i + 1
Loop
Range(Cells(i + 1, k + 1), Cells(i + 7, k + 1)).Select
    With Selection
         .HorizontalAlignment = xlLeft
        .ShrinkToFit = False
End With
Next k

For k = 10 To [d1] Step 10
i = 16: Do Until Left(Cells(i, k + 8), 1) = ":" Or Cells(i, k + 8) = ""
i = i + 1
Loop
Cells(i, k + 8).Select
    With Selection
         .HorizontalAlignment = xlLeft
        .ShrinkToFit = False
End With
Next k

Columns(10).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Range(Cells(1, dfx + 5), Cells(2, dfx + 5)).Select
    With Selection
         .HorizontalAlignment = xlLeft
        .ShrinkToFit = False
End With

Cells(1, drr) = "XIRR"
Cells(1, drr + 1) = "__"¡"
Columns(drr + 1).NumberFormatLocal = "yyyy/m/d;@"
Cells(1, drr + 2) = "__"º"
Cells(1, drr + 3) = "__"˜"
Cells(1, drr + 4) = "__"˜"
Cells(1, drr + 5) = "____"W√B"
Cells(1, drr + 6) = "____"W√B"
Cells(1, drr + 7) = "____ = ____ - ____" - •ª™˜ºW√B"

Range(Columns(drr), Columns(drr + 1)).Select
   With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
End With

Range(Columns(drr + 2), Columns(drr + 8)).Select
   Selection.NumberFormatLocal = "#,##0_);[__](#,##0)")"
   With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
End With

Range(Cells(1, drr), Cells(1, drr + 8)).Select
   With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
End With

Cells(1, drr + 7).Select
    With Selection
         .HorizontalAlignment = xlLeft
        .ShrinkToFit = False
End With

Cells(1, dbk) = "__"º"
Cells(2, dbk) = "__"˜"
Cells(3, dbk) = "__"˜"
Cells(4, dbk) = "___"ßQ"
Cells(5, dbk) = "Fx"

Range(Cells(1, dbk), Cells(5, dbk)).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Range(Columns(dbk + 1), Columns(dbk + [d1] / 10)).Select
Selection.NumberFormatLocal = "#,##0.00_);(#,##0.00)"
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Range(Cells(1, dbk + 1), Cells(4, dbk + [d1] / 10)).Select
   Selection.NumberFormatLocal = "#,##0_);[__](#,##0)")"
   With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
End With
    
    Range("E9:G10, h10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range("E10").Select
    Selection.NumberFormatLocal = "#,##0.0_);(#,##0.0)"
    Range("F10").Select
    Selection.NumberFormatLocal = "0%"
    Range("G10").Select
    Selection.NumberFormatLocal = "#,##0"
    
    Range("E10:F10").Select
    With Selection.Font
        .Color = -52429
        .TintAndShade = 0
    End With
    
    Range("G9,H10").Select
    With Selection
        .ShrinkToFit = False
    End With
    
    Range("G9").Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With

End If


End Sub



Sub Macro35()
'
' mikeon _ 2018/4/6 _____™∫•®∂∞
' ____©Ò§j

'
Sheets("__").Selectct
Application.ScreenUpdating = False
Range("A1:K16").Select: ActiveWindow.Zoom = True
i = 16: Do Until Cells(i, 2) = ""
i = i + 1
Loop

Cells(i - 1, 3).Select
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1

End Sub

Sub Macro36()
'
' mikeon _ 2018/4/6 _____™∫•®∂∞
'____¡Y§p

'
Sheets("__").Selectct
Application.ScreenUpdating = False
Range("A1:M16").Select: ActiveWindow.Zoom = True
i = 16: Do Until Cells(i, 2) = ""
i = i + 1
Loop

Cells(i - 1, 3).Select
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1

End Sub


Sub Macro37()
'
' mikeon _ 2019/10/5 _____™∫•®∂∞
' ______∂i¬¬¿…

'
Application.ScreenUpdating = True
Call UnprotectSheet(ActiveSheet)
[h3] = "___...".."
[h3].Select
    With Selection.Font
        .Color = -16776961
    End With
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With

Application.Wait Now() + TimeValue("00:00:01")
Application.StatusBar = False
Application.EnableCancelKey = xlInterrupt '_____\®q∂≤
Application.Calculation = xlManual
Application.ScreenUpdating = False

Cells.Clear

Dim FileToOpen As Variant
Dim OpenBook As Workbook
FileToOpen = Application.GetOpenFilename(Title:="______")¶A™Ì")
If FileToOpen <> False Then
Set OpenBook = Application.Workbooks.Open(FileToOpen)
OpenBook.Sheets("__").Activatete
Call UnprotectSheet(ActiveSheet)

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

[a1] = "1. _______16_______"C©≥§UΩd®“ßR∞£"
[a2] = "2. ___________ [__]"A¶A´ˆ [ßÛ∑s]"
[a3] = "3. ____ " & [i15] & " ___ i15_________A70"•NΩX∞—æ\¨¸™—A70"
[a4] = "4. XIRR _________________ 1 _ (1+XIRR)^(__/365)-1"+XIRR)^(§—º∆/365)-1"
[a5] = "5. __ [__] [__] [__] [__] [__] _______ [__] __"ßÛßÔ°A¶A´ˆ [ßÛ∑s] ßY•i"
[a6] = "6. K _________________"•[§@±i°A±iº∆µL≠≠"
[a7] = ""
[a8] = "________ B"·´¸æ… B"
[a9] = ""
[b12] = "": [b13] = ""

Range("D15").FormulaR1C1 = "=RC[5]"
Range("E15").FormulaR1C1 = "=RC[4]"
Range("F15").FormulaR1C1 = "=RC[3]"
Range("H15").FormulaR1C1 = "=RC[1]"


f = 10: Do Until Left(Cells(15, f), 1) <> "_"""
f = f + 10
Loop
[d1] = f - 10

dfx = [d1] + 30 '__≤v
drr = dfx + 9 'XIRR
dbk = drr + 15 '____ÆÊßÔ

For f = 10 To [d1] Step 10

Cells(1, f) = "_____X"kÆÊX"
Cells(2, f + 2).FormulaR1C1 = "=RC[2]"
Cells(2, f + 3).FormulaR1C1 = "=RC[1]"
Cells(2, f + 5).FormulaR1C1 = "=R[13]C[" & -6 - (f - 10) & "]"
Cells(15, f + 3).FormulaR1C1 = "=RC[" & -4 - (f - 10) & "]"
Cells(15, f + 4).FormulaR1C1 = "=RC[" & -5 - (f - 10) & "]"
Cells(15, f + 5).FormulaR1C1 = "=RC[" & -6 - (f - 10) & "]"
Cells(15, f + 7).FormulaR1C1 = "=R[-13]C[-3]"
Cells(15, f + 8).FormulaR1C1 = "=RC[" & -9 - (f - 10) & "]"
Cells(15, f + 9).FormulaR1C1 = "=RC[" & -10 - (f - 10) & "]"

If Cells(3, f + 5) > 0 Then
If UCase(Cells(2, f + 4)) = "USD" Then GoTo USDD
i = 3: Do Until Cells(i, f + 2) = "" And Cells(i, f + 3) = "" And i < 13   '_x
Cells(i, f + 4).FormulaR1C1 = "=RC[-2]+RC[-1]"
i = i + 1
Loop
GoTo ipt

USDD:
i = 3: Do Until Cells(i, f + 2) = "" And Cells(i, f + 4) = "" And i < 13   '_¸
Cells(i, f + 2).FormulaR1C1 = "=RC[2]-RC[1]"
i = i + 1
Loop
End If

ipt:
Next f

Application.ScreenUpdating = True
i = 16: Do Until Cells(i, 2) = ""
i = i + 1
Loop

Cells(i - 1, 3).Select
ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
[h3] = "": [d1] = ""
Application.StatusBar = "__"®"
Application.Calculation = xlAutomatic
Beep

End Sub


