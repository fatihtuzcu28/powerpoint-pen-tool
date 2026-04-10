@echo off
setlocal EnableExtensions EnableDelayedExpansion
:: PPTKalem COM Add-in Kaldırma Scripti
:: Yönetici olarak çalıştırın!

echo ============================================
echo   PPTKalem - PowerPoint Kalem Add-in
echo   Kaldirma Scripti
echo ============================================
echo.

:: Yönetici kontrolü
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [HATA] Bu scripti yonetici olarak calistirin!
    echo Sag tik ^> Yonetici olarak calistir
    pause
    exit /b 1
)

:: PowerPoint açıksa kapat
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if %errorlevel% equ 0 (
    echo [!] PowerPoint acik, kapatiliyor...
    taskkill /IM POWERPNT.EXE /F >nul 2>&1
    timeout /t 2 /nobreak >nul
    echo     PowerPoint kapatildi.
    echo.
)

set "DLL_NAME=PPTKalem.dll"
set "ADDINS_DIR=%APPDATA%\Microsoft\AddIns"

:: Office bit (32/64) otomatik algila - 32-bit icin Framework, 64-bit icin Framework64
set "PPT_PATH="
for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE" /ve 2^>nul ^| find /I "REG_SZ"') do set "PPT_PATH=%%B"
if not defined PPT_PATH for /f "tokens=2,*" %%A in ('reg query "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE" /ve 2^>nul ^| find /I "REG_SZ"') do set "PPT_PATH=%%B"

set "REGASM=C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
if defined PPT_PATH (
    echo [i] PowerPoint yolu: !PPT_PATH!
    echo(!PPT_PATH! | find /I "Program Files (x86)" >nul
    if errorlevel 1 (
        echo [i] Office 64-bit algilandi.
        set "REGASM=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
    ) else (
        echo [i] Office 32-bit algilandi.
        set "REGASM=C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    )
) else (
    echo [!] PowerPoint yolu bulunamadi. Varsayilan 32-bit RegAsm kullanilacak.
)

echo [i] RegAsm: %REGASM%

:: RegAsm ile COM kayıt silme
echo [1/2] COM kaydi siliniyor...
if exist "%ADDINS_DIR%\%DLL_NAME%" (
    "%REGASM%" "%ADDINS_DIR%\%DLL_NAME%" /unregister
) else (
    echo DLL AddIns klasorunde bulunamadi, registry temizleniyor...
    reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PPTKalem.Connect" /f 2>nul
)

:: DLL'i sil
echo [2/2] DLL siliniyor...
if exist "%ADDINS_DIR%\%DLL_NAME%" (
    del /F "%ADDINS_DIR%\%DLL_NAME%"
    echo      Silindi: %ADDINS_DIR%\%DLL_NAME%
) else (
    echo      DLL zaten mevcut degil.
)

echo.
echo ============================================
echo   Kaldirma tamamlandi!
echo   PowerPoint'i yeniden baslattiginizda
echo   add-in artik yuklenmeyecek.
echo ============================================
pause
