@echo off  
title CLOUDFLARE PERMANENT TUNNEL - ancol-analisa.my.id
color 0A  
echo ======================================================================  
echo       MENYALAKAN CLOUDFLARE PERMANENT TUNNEL (ancol-analisa.my.id)
echo ======================================================================  
echo.  
echo Sedang menghubungkan Flask port 5000 ke domain tetap Anda...  
echo Alamat Web Tetap: https://app.ancol-analisa.my.id
echo.  
cd /d %~dp0
cloudflared.exe tunnel run --token eyJhIjoiMTdmY2VhZTg5ZjI0ZmZlNDJjNzk5YWI0NjlkODhiNWYiLCJ0IjoiNjFkZjJjYWMtMzQxNS00MDRkLWE3Y2EtZTc1OWQzZWRjOGZhIiwicyI6Ik4yUm1Nemt5TUdJdE9EVmtZUzAwTkRKa0xXSXhOamd0WlRSaE9HVTVaV05pWkdRNSJ9
pause 
