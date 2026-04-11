on FetchStock(symbolArg)
	set pythonPath to "/Users/xiaotingzhou/Downloads/Trading Project/venv/bin/python"
	set scriptPath to "/Users/xiaotingzhou/Downloads/Trading Project/stock_fetcher.py"
	set cmd to quoted form of pythonPath & " " & quoted form of scriptPath & " " & symbolArg
	return do shell script cmd
end FetchStock
