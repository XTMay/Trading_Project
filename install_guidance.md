  Step 1: 安装 Python

  - 下载 https://www.python.org/downloads/
  - 安装时 必须勾选 "Add Python to PATH"

  Step 2: 运行安装脚本

  - 把 StockTool_Windows.zip 解压到 C:\StockTool\
  - 双击 windows_setup.bat
  - 它会自动创建虚拟环境并安装 yfinance + openpyxl

  Step 3: 配置 Excel VBA

  1. 打开 output.xlsm
  2. 按 Alt+F11 打开 VBA 编辑器
  3. 左侧双击 Sheet1(Code) → 粘贴 vba_windows.bas 的 PART 1
  4. Insert → Module → 粘贴 vba_windows.bas 的 PART 2
  5. 修改这两行路径：
  Const PYTHON_PATH As String = "C:\StockTool\venv\Scripts\python.exe"
  Const SCRIPT_PATH As String = "C:\StockTool\stock_fetcher.py"
  6. 关闭 VBA 编辑器，保存

  Step 4: 使用

  在 Code sheet 的 A2 输入股票代码（如 AAPL），回车，等待数据自动填入。