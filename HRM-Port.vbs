Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d D:\HRM_Port && streamlit run app.py --server.headless true 2>nul", 0, False
WScript.Sleep 5000
WshShell.Run "http://localhost:8501"