@echo off
title Aktivasi Office - KMS

echo ==============================================
echo    Microsoft Office Activator - KMS Script
echo ==============================================
echo.

:: Matikan proteksi sementara (biasa dilakukan di aktivator ilegal)
:: net stop winmgmt /y
:: taskkill /f /im OfficeClickToRun.exe

:: Jalankan aktivasi via KMS
echo Menghubungkan ke server KMS...
cscript //nologo "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /sethst:kms.loli.best >nul
cscript //nologo "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /act >nul

:: Cek hasil aktivasi
echo.
echo Aktivasi selesai. Silakan cek Office Anda.
timeout /t 2 >nul

:: Buka halaman donasi/saweria
start https://www.instagram.com/stria_bgs/


exit
