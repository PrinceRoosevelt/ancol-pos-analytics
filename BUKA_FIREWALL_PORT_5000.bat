@echo off
:: Script untuk membuka izin akses port 5000 di Windows Firewall
netsh advfirewall firewall add rule name="Sales Dashboard Port 5000" dir=in action=allow protocol=TCP localport=5000
echo ==========================================================
echo   Izin Port 5000 berhasil ditambahkan ke Windows Firewall!
echo   PC lain di jaringan kantor sekarang bisa membuka dashboard.
echo ==========================================================
pause
