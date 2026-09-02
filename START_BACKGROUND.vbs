Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""cd 'C:\PYTHONV2'; .\.venv\Scripts\python.exe server.py""", 0, False
