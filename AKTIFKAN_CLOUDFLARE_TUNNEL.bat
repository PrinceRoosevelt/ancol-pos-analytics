@echo off  
title CLOUDFLARE LIVE TUNNEL - ANCOL SERVER  
color 0B  
echo ======================================================================  
echo             MENYALAKAN SERVER CLOUDFLARE TUNNEL  
echo ======================================================================  
echo.  
echo Sedang menghubungkan port 5000 ke internet global...  
echo LIHAT LINK DI DALAM KOTAK BESAR DI BAWAH INI (https://xxxx.trycloudflare.com)  
echo.  
cloudflared.exe tunnel --url http://127.0.0.1:5000  
pause 
