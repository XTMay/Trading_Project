# -*- coding: utf-8 -*-
"""
setup_xlsm.py (EZ_table0227)
============================
Convert report_summary.xlsx to report_summary_yFinance.xlsm and inject VBA macros.

Usage:
    python setup_xlsm.py

Prerequisites:
    1. Microsoft Excel installed (Windows)
    2. Trust access to the VBA project object model is enabled:
       Excel -> File -> Options -> Trust Center -> Trust Center Settings
       -> Macro Settings -> check "Trust access to the VBA project object model"
"""

import io
import os
import shutil
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(HERE, "report_summary.xlsx")
XLSM_PATH = os.path.join(HERE, "report_summary_yFinance.xlsm")
PYTHON_PATH = r"C:\Github\Trading_Project\venv\Scripts\python.exe"
SCRIPT_PATH = r"C:\Github\Trading_Project\EZ_table0227\generate_report_summary.py"

SHEET_CODE = (
    'Private Sub Worksheet_Change(ByVal Target As Range)\r\n'
    '    If Target.Address <> "$A$2" Then Exit Sub\r\n'
    '    If Trim(CStr(Target.Value)) = "" Then Exit Sub\r\n'
    '    Call FetchReportData\r\n'
    'End Sub\r\n'
)

_py = PYTHON_PATH.replace("\\", "\\\\")
_scr = SCRIPT_PATH.replace("\\", "\\\\")

MODULE_CODE = (
    'Const PYTHON_PATH As String = "' + _py + '"\r\n'
    'Const SCRIPT_PATH As String = "' + _scr + '"\r\n'
    '\r\n'
    'Sub FetchReportData()\r\n'
    '\r\n'
    '    Dim ws As Worksheet\r\n'
    '    Dim symbol As String\r\n'
    '    Dim tempPath As String\r\n'
    '    Dim runCmd As String\r\n'
    '    Dim wsh As Object\r\n'
    '\r\n'
    '    Set ws = ThisWorkbook.Sheets(1)\r\n'
    '    symbol = Trim(CStr(ws.Range("A2").Value))\r\n'
    '\r\n'
    '    If symbol = "" Then\r\n'
    '        MsgBox "Please enter a stock symbol in A2 (e.g. DIOD, AAPL, 2330.TW)", vbExclamation, "Notice"\r\n'
    '        Exit Sub\r\n'
    '    End If\r\n'
    '\r\n'
    '    Application.EnableEvents = False\r\n'
    '    ws.Range("B1").Value = "Status"\r\n'
    '    ws.Range("B2").Value = "Fetching " & symbol & "..."\r\n'
    '    Application.ScreenUpdating = True\r\n'
    '    DoEvents\r\n'
    '\r\n'
    '    tempPath = Environ("TEMP") & "\\report_temp_" & symbol & ".xlsx"\r\n'
    '    On Error Resume Next\r\n'
    '    Kill tempPath\r\n'
    '    On Error GoTo 0\r\n'
    '\r\n'
    '    runCmd = Chr(34) & PYTHON_PATH & Chr(34) & " " _\r\n'
    '           & Chr(34) & SCRIPT_PATH & Chr(34) & " " _\r\n'
    '           & symbol & " " _\r\n'
    '           & Chr(34) & Chr(34) & " " _\r\n'
    '           & Chr(34) & tempPath & Chr(34)\r\n'
    '\r\n'
    '    Set wsh = CreateObject("WScript.Shell")\r\n'
    '    wsh.Run runCmd, 0, True\r\n'
    '    Set wsh = Nothing\r\n'
    '\r\n'
    '    If Dir(tempPath) = "" Then\r\n'
    '        ws.Range("B2").Value = "Error"\r\n'
    '        Application.EnableEvents = True\r\n'
    '        MsgBox "Fetch failed for: " & symbol & vbCrLf & _\r\n'
    '               "Check:" & vbCrLf & _\r\n'
    '               "1. Python: " & PYTHON_PATH & vbCrLf & _\r\n'
    '               "2. Network connection" & vbCrLf & _\r\n'
    '               "3. Symbol is valid", vbExclamation, "Error"\r\n'
    '        Exit Sub\r\n'
    '    End If\r\n'
    '\r\n'
    '    Application.ScreenUpdating = False\r\n'
    '\r\n'
    '    Dim tempWb As Workbook\r\n'
    '    Dim srcSheet As Worksheet\r\n'
    '    Dim dstSheet As Worksheet\r\n'
    '\r\n'
    '    Set tempWb = Workbooks.Open(tempPath, ReadOnly:=True, UpdateLinks:=False)\r\n'
    '    Set srcSheet = tempWb.Sheets(1)\r\n'
    '\r\n'
    '    Set dstSheet = ws\r\n'
    '    dstSheet.Cells.ClearContents\r\n'
    '\r\n'
    '    srcSheet.UsedRange.Copy\r\n'
    '    dstSheet.Range("A1").PasteSpecial Paste:=xlPasteValues\r\n'
    '    Application.CutCopyMode = False\r\n'
    '\r\n'
    '    tempWb.Close SaveChanges:=False\r\n'
    '    Set tempWb = Nothing\r\n'
    '\r\n'
    '    On Error Resume Next\r\n'
    '    Kill tempPath\r\n'
    '    On Error GoTo 0\r\n'
    '\r\n'
    '    If Trim(CStr(dstSheet.Range("A2").Value)) = "" Then\r\n'
    '        dstSheet.Range("A2").Value = symbol\r\n'
    '    End If\r\n'
    '\r\n'
    '    dstSheet.Range("B1").Value = "Status"\r\n'
    '    dstSheet.Range("B2").Value = "Done - " & symbol & "  " & Format(Now, "HH:MM:SS")\r\n'
    '\r\n'
    '    Application.ScreenUpdating = True\r\n'
    '    Application.EnableEvents = True\r\n'
    '    dstSheet.Activate\r\n'
    '    MsgBox symbol & " data loaded!", vbInformation, "Done"\r\n'
    '\r\n'
    'End Sub\r\n'
)


def inject_vba(wb):
    """Inject VBA code into Sheet1 and FetchModule."""
    try:
        _ = wb.VBProject
    except Exception:
        print("\n[ERROR] Cannot access VBA project object model.")
        print("Enable Excel Trust Center option:")
        print("  Trust access to the VBA project object model")
        raise

    # inject worksheet event code
    sheet_comp = wb.VBProject.VBComponents("Sheet1")
    count = sheet_comp.CodeModule.CountOfLines
    existing = sheet_comp.CodeModule.Lines(1, count) if count > 0 else ""
    if "Worksheet_Change" not in existing:
        sheet_comp.CodeModule.AddFromString(SHEET_CODE)
        print("[OK] Worksheet_Change injected into Sheet1")
    else:
        print("[SKIP] Sheet1 already contains Worksheet_Change")

    # replace FetchModule
    for comp in list(wb.VBProject.VBComponents):
        if comp.Name == "FetchModule":
            wb.VBProject.VBComponents.Remove(comp)
            break

    new_mod = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
    new_mod.Name = "FetchModule"
    new_mod.CodeModule.AddFromString(MODULE_CODE)
    print("[OK] FetchModule injected")


def main():
    if not os.path.exists(XLSX_PATH):
        print(f"[ERROR] Source file not found: {XLSX_PATH}")
        sys.exit(1)

    print(f"Source: {XLSX_PATH}")
    print(f"Target: {XLSM_PATH}")

    if os.path.exists(XLSM_PATH):
        backup = XLSM_PATH.replace(".xlsm", "_backup.xlsm")
        shutil.copy2(XLSM_PATH, backup)
        print(f"[OK] Backup created: {backup}")

    try:
        import win32com.client
    except ImportError:
        print("[ERROR] Missing pywin32. Install with:")
        print(f"  {PYTHON_PATH} -m pip install pywin32")
        sys.exit(1)

    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    try:
        # 1) xlsx -> xlsm
        wb = xl.Workbooks.Open(XLSX_PATH)
        wb.SaveAs(XLSM_PATH, FileFormat=52)  # xlOpenXMLMacroEnabled
        wb.Close()

        # 2) open xlsm and inject VBA
        wb = xl.Workbooks.Open(XLSM_PATH)
        inject_vba(wb)
        wb.Save()
        wb.Close()

        print("\n" + "=" * 56)
        print("Done.")
        print(f"  {XLSM_PATH}")
        print("How to use:")
        print("  1. Open report_summary_yFinance.xlsm")
        print("  2. Enter ticker in A2 (e.g. DIOD)")
        print("  3. Press Enter and wait for auto-fetch")
        print("=" * 56)

    finally:
        try:
            xl.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()

