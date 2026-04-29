' ============================================================
' report_summary_yFinance.bas (EZ_table0227)
' Auto-fetch stock data into report_summary_yFinance.xlsm
'
' 使用方式：
'   在 A2 输入公司代号（如 DIOD、AAPL、2330.TW）后按 Enter
'   自动调用 Python 抓取所有财务数据，写入当前工作表
' ============================================================


' ============================================================
' PART 1：贴入 Sheet1(Code) 对象
' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address <> "$A$2" Then Exit Sub
    If Trim(CStr(Target.Value)) = "" Then Exit Sub
    Call FetchReportData
End Sub


' ============================================================
' PART 2：贴入新 Module（Insert -> Module）
' ============================================================

Const PYTHON_PATH As String = "C:\Github\Trading_Project\venv\Scripts\python.exe"
Const SCRIPT_PATH As String = "C:\Github\Trading_Project\EZ_table0227\generate_report_summary.py"

Sub FetchReportData()

    Dim ws As Worksheet
    Dim symbol As String
    Dim tempPath As String
    Dim runCmd As String
    Dim wsh As Object

    Set ws = ThisWorkbook.Sheets(1)
    symbol = Trim(CStr(ws.Range("A2").Value))

    If symbol = "" Then
        MsgBox "Please enter a stock symbol in A2 (e.g. DIOD, AAPL, 2330.TW)", vbExclamation, "Notice"
        Exit Sub
    End If

    Application.EnableEvents = False
    ws.Range("B1").Value = "Status"
    ws.Range("B2").Value = "Fetching " & symbol & "..."
    Application.ScreenUpdating = True
    DoEvents

    tempPath = Environ("TEMP") & "\report_temp_" & symbol & ".xlsx"
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    ' Call Python directly (no cmd /c) to avoid Windows quote-parsing issues
    ' Args: script.py  <symbol>  <currency=''>  <output_path>
    runCmd = Chr(34) & PYTHON_PATH & Chr(34) & " " _
           & Chr(34) & SCRIPT_PATH & Chr(34) & " " _
           & symbol & " " _
           & Chr(34) & Chr(34) & " " _
           & Chr(34) & tempPath & Chr(34)

    Set wsh = CreateObject("WScript.Shell")
    wsh.Run runCmd, 0, True
    Set wsh = Nothing

    If Dir(tempPath) = "" Then
        ws.Range("B2").Value = "Error"
        Application.EnableEvents = True
        MsgBox "Fetch failed for: " & symbol & vbCrLf & _
               "Check:" & vbCrLf & _
               "1. Python: " & PYTHON_PATH & vbCrLf & _
               "2. Network connection" & vbCrLf & _
               "3. Symbol is valid", vbExclamation, "Error"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim tempWb As Workbook
    Dim srcSheet As Worksheet
    Dim dstSheet As Worksheet

    Set tempWb = Workbooks.Open(tempPath, ReadOnly:=True, UpdateLinks:=False)
    Set srcSheet = tempWb.Sheets(1)

    Set dstSheet = ws
    dstSheet.Cells.ClearContents

    srcSheet.UsedRange.Copy
    dstSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    tempWb.Close SaveChanges:=False
    Set tempWb = Nothing

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    If Trim(CStr(dstSheet.Range("A2").Value)) = "" Then
        dstSheet.Range("A2").Value = symbol
    End If

    dstSheet.Range("B1").Value = "Status"
    dstSheet.Range("B2").Value = "Done - " & symbol & "  " & Format(Now, "HH:MM:SS")

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    dstSheet.Activate
    MsgBox symbol & " data loaded!", vbInformation, "Done"

End Sub

