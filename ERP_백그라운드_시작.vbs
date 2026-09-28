Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d ""G:\내 드라이브\시험농장"" && .venv\Scripts\python.exe -m streamlit run erp_ui.py --server.address=0.0.0.0", 0
Set WshShell = Nothing
