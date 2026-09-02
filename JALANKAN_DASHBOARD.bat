@echo off
title Sales Dashboard Server
cd /d "%~dp0"
echo ===================================================
echo   SALES PERFORMANCE DASHBOARD SERVER
echo ===================================================
echo   Lokal PC ini  : http://localhost:5000/
echo   Dari PC Lain  : http://10.22.4.50:5000/
echo   Upload Data   : http://10.22.4.50:5000/upload
echo ===================================================
echo   Menjalankan server... (Jangan tutup jendela ini)
echo ===================================================
".venv\Scripts\python.exe" server.py
pause
