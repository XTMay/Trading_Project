Sub FetchStockData()

    Dim symbol As String                ' 声明一个字符串变量, 存放股票代码
    Dim codeSheet As Worksheet          ' 声明一个工作表变量, 指向 Code sheet
    Set codeSheet = ThisWorkbook.Sheets(1) ' 将 codeSheet 设置为当前工作簿的第一个 Sheet

    symbol = Trim(CStr(codeSheet.Range("A2").Value)) ' 读取 A2 单元格的值, 去除空格, 转为字符串

    If symbol = "" Then                 ' 如果 A2 是空的
        MsgBox "Please enter a stock symbol in A2", vbExclamation ' 弹窗提示用户输入代码
        Exit Sub                        ' 退出函数
    End If

    codeSheet.Range("B1").Value = "Status"    ' B1 写入标题 "Status"
    codeSheet.Range("B2").Value = "Fetching..." ' B2 显示状态 "Fetching..."
    Application.ScreenUpdating = True   ' 强制刷新屏幕, 让用户看到状态变化
    DoEvents                            ' 让 Excel 处理待处理的事件 (保证屏幕刷新)

    Dim tempPath As String              ' 声明临时文件路径变量
    tempPath = "/tmp/stock_temp.xlsx"   ' 设置临时文件路径 (Python 会把数据写到这里)

    Dim result As String                ' 声明变量接收 AppleScript 返回值
    result = AppleScriptTask("FetchStock.scpt", "FetchStock", symbol) ' 调用 AppleScript -> 启动 Python 获取数据

    If Dir(tempPath) = "" Then          ' 检查临时文件是否生成成功
        codeSheet.Range("B2").Value = "Error" ' 如果文件不存在, 显示错误
        MsgBox "Fetch failed. Check Python and network.", vbExclamation ' 弹窗提示失败
        Exit Sub                        ' 退出函数
    End If

    Application.ScreenUpdating = False  ' 暂停屏幕刷新 (复制数据时避免屏幕闪烁)

    Dim tempWb As Workbook              ' 声明临时工作簿变量
    Set tempWb = Workbooks.Open(tempPath, ReadOnly:=True, UpdateLinks:=False) ' 打开临时 xlsx 文件 (只读模式)

    Dim srcSheet As Worksheet           ' 声明源 sheet 变量 (临时文件中的 sheet)
    Dim dstSheet As Worksheet           ' 声明目标 sheet 变量 (主工作簿中的 sheet)
    Dim i As Integer                    ' 循环计数器

    For i = 1 To tempWb.Sheets.Count   ' 遍历临时文件中的每一个 Sheet

        Set srcSheet = tempWb.Sheets(i) ' 获取当前源 sheet

        If srcSheet.Name = "Income" Then ' 如果是 Income (损益表), 特殊处理: 放到 Code sheet

            ' --- Income -> Code sheet L2 ---
            Dim incomeArea As Range     ' 声明目标区域变量
            Set incomeArea = codeSheet.Range("L2").Resize( _
                srcSheet.UsedRange.Rows.Count, _
                srcSheet.UsedRange.Columns.Count _
            )                           ' 计算目标区域大小 (从 L2 开始, 与源数据同样大小)

            incomeArea.Clear            ' 先清空目标区域的旧数据
            srcSheet.UsedRange.Copy     ' 复制源 sheet 的有效数据区域
            codeSheet.Range("L2").PasteSpecial Paste:=xlPasteValues ' 只粘贴值到 Code sheet 的 L2 位置
            Application.CutCopyMode = False ' 清除剪切板 (取消复制虚框)

        Else                            ' 其他 Sheet: 创建独立的 sheet 来存放

            ' --- Other sheets -> separate sheet ---
            On Error Resume Next        ' 忽略错误 (用于检查 sheet 是否已存在)
            Set dstSheet = ThisWorkbook.Sheets(srcSheet.Name) ' 尝试获取同名 sheet
            On Error GoTo 0             ' 恢复正常错误处理

            If dstSheet Is Nothing Then ' 如果目标 sheet 不存在
                Set dstSheet = ThisWorkbook.Sheets.Add( _
                    After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)) ' 在最后面新建一个 sheet
                dstSheet.Name = srcSheet.Name ' 设置 sheet 名称 (如 Info, History_ 等)
            Else                        ' 如果目标 sheet 已存在
                dstSheet.Cells.Clear    ' 清空所有旧数据
            End If

            srcSheet.UsedRange.Copy     ' 复制源 sheet 的有效数据区域
            dstSheet.Range("A1").PasteSpecial Paste:=xlPasteValues ' 只粘贴值到目标 sheet 的 A1
            Application.CutCopyMode = False ' 清除剪切板

            Set dstSheet = Nothing      ' 释放目标 sheet 引用

        End If

    Next i                              ' 继续下一个 sheet

    tempWb.Close SaveChanges:=False     ' 关闭临时文件 (不保存)

    On Error Resume Next                ' 忽略删除文件时可能的错误
    Kill tempPath                       ' 删除临时文件 /tmp/stock_temp.xlsx
    On Error GoTo 0                     ' 恢复正常错误处理

    codeSheet.Activate                  ' 切换回 Code sheet
    codeSheet.Range("B2").Value = "Done" ' 更新状态为 "Done"

    On Error Resume Next                ' 以下取公司名称, 可能失败所以忽略错误
    codeSheet.Range("C1").Value = "Company" ' C1 写入标题
    codeSheet.Range("C2").Value = ThisWorkbook.Sheets("Info").Range("B3").Value ' 从 Info sheet 读取公司名称写入 C2
    On Error GoTo 0                     ' 恢复正常错误处理

    Application.ScreenUpdating = True   ' 恢复屏幕刷新

    MsgBox symbol & " data loaded! Check each sheet.", vbInformation ' 弹窗提示完成

End Sub
