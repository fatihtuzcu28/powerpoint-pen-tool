@echo off
setlocal EnableExtensions EnableDelayedExpansion
:: PPTKalem COM Add-in Kurulum Scripti
:: Yönetici olarak çalıştırın!

echo ============================================
echo   PPTKalem - PowerPoint Kalem Add-in
echo   Kurulum Scripti
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

:: DLL yolunu belirle (önce Release, yoksa Debug)
set "SOURCE_DIR=%~dp0PPTKalem\bin\Release\net481"
if not exist "%SOURCE_DIR%\PPTKalem.dll" set "SOURCE_DIR=%~dp0PPTKalem\bin\Debug\net481"
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

:: DLL var mı kontrol et
if not exist "%SOURCE_DIR%\%DLL_NAME%" (
    echo [HATA] %DLL_NAME% bulunamadi!
    echo Once projeyi Visual Studio'da derleyin.
    echo Beklenen konum: %SOURCE_DIR%\%DLL_NAME%
    pause
    exit /b 1
)

:: AddIns klasörü yoksa oluştur
if not exist "%ADDINS_DIR%" mkdir "%ADDINS_DIR%"

:: DLL'i AddIns klasörüne kopyala
echo [1/2] DLL kopyalaniyor...
copy /Y "%SOURCE_DIR%\%DLL_NAME%" "%ADDINS_DIR%\%DLL_NAME%"
if %errorlevel% neq 0 (
    echo [HATA] Kopyalama basarisiz!
    pause
    exit /b 1
)
echo      %ADDINS_DIR%\%DLL_NAME%

:: RegAsm ile COM kayıt
echo [2/2] COM kaydediliyor...
"%REGASM%" "%ADDINS_DIR%\%DLL_NAME%" /codebase
if %errorlevel% neq 0 (
    echo [HATA] RegAsm kaydi basarisiz!
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Kurulum tamamlandi!
echo   PowerPoint'i acin, "Kalem Araclari"
echo   ribbon sekmesini goreceksiniz.
echo ============================================
pause
