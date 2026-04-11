' ============================================================
' [Windows] Auto Fetch Stock Data via Python + yfinance
' ============================================================
'
' Step 1: Paste PART 1 into Sheet1(Code) object
'         (double-click Sheet1 in VBA Project Explorer)
'
' Step 2: Insert -> Module, paste PART 2 into the module
'
' Step 3: Change PYTHON_PATH and SCRIPT_PATH to your paths
'
' Step 4: Save as .xlsm format
' ============================================================


' ============================================================
' PART 1: Paste into Sheet1(Code) object
' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$A$2" Then Exit Sub
    If Target.Value = "" Then Exit Sub
    Call FetchStockData
End Sub


' ============================================================
' PART 2: Paste into a new Module (Insert -> Module)
' ============================================================

' ***** CHANGE THESE TWO PATHS *****
Const PYTHON_PATH As String = "C:\StockTool\venv\Scripts\python.exe"
Const SCRIPT_PATH As String = "C:\StockTool\stock_fetcher.py"

Sub FetchStockData()

    Dim symbol As String
    Dim codeSheet As Worksheet
    Set codeSheet = ThisWorkbook.Sheets(1)

    symbol = Trim(CStr(codeSheet.Range("A2").Value))

    If symbol = "" Then
        MsgBox "Please enter a stock symbol in A2", vbExclamation
        Exit Sub
    End If

    ' Show status
    codeSheet.Range("B1").Value = "Status"
    codeSheet.Range("B2").Value = "Fetching..."
    Application.ScreenUpdating = True
    DoEvents

    ' Call Python to generate temp file
    Dim tempPath As String
    tempPath = Environ("TEMP") & "\stock_temp.xlsx"

    Dim cmd As String
    cmd = """" & PYTHON_PATH & """ """ & SCRIPT_PATH & """ " & symbol & " """ & tempPath & """"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c " & cmd, 0, True   ' 0=hidden window, True=wait for completion
    Set wsh = Nothing

    ' Check if temp file was created
    If Dir(tempPath) = "" Then
        codeSheet.Range("B2").Value = "Error"
        MsgBox "Fetch failed. Check Python and network.", vbExclamation
        Exit Sub
    End If

    ' Open temp file and copy data to this workbook
    Application.ScreenUpdating = False
    Dim tempWb As Workbook
    Set tempWb = Workbooks.Open(tempPath, ReadOnly:=True, UpdateLinks:=False)

    Dim srcSheet As Worksheet
    Dim dstSheet As Worksheet
    Dim i As Integer

    For i = 1 To tempWb.Sheets.Count

        Set srcSheet = tempWb.Sheets(i)

        If srcSheet.Name = "Income" Then

            ' --- Income -> Code sheet L2 ---
            Dim incomeArea As Range
            Set incomeArea = codeSheet.Range("L2").Resize( _
                srcSheet.UsedRange.Rows.Count, _
                srcSheet.UsedRange.Columns.Count)

            incomeArea.Clear
            srcSheet.UsedRange.Copy
            codeSheet.Range("L2").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False

        Else

            ' --- Other sheets -> separate sheet ---
            On Error Resume Next
            Set dstSheet = ThisWorkbook.Sheets(srcSheet.Name)
            On Error GoTo 0

            If dstSheet Is Nothing Then
                Set dstSheet = ThisWorkbook.Sheets.Add( _
                    After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                dstSheet.Name = srcSheet.Name
            Else
                dstSheet.Cells.Clear
            End If

            srcSheet.UsedRange.Copy
            dstSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False

            Set dstSheet = Nothing

        End If

    Next i

    ' Close and delete temp file
    tempWb.Close SaveChanges:=False
    Set tempWb = Nothing
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    ' Update status
    codeSheet.Activate
    codeSheet.Range("B2").Value = "Done"

    On Error Resume Next
    codeSheet.Range("C1").Value = "Company"
    codeSheet.Range("C2").Value = ThisWorkbook.Sheets("Info").Range("B3").Value
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox symbol & " data loaded! Check each sheet.", vbInformation

End Sub
