@echo off
chcp 65001 >nul
cls
echo ======================================================
echo    DANG CAI DAT ADD-IN CHUAN HOA VAN BAN DANG ^& EXCEL
echo              Van phong Dang uy xa Thanh Phu
echo ======================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "WORD_SOURCE=%SCRIPT_DIR%ChuanHoa_VanBanDang.dotm"
set "EXCEL_SOURCE=%SCRIPT_DIR%ChuanHoa_Excel.xlam"

set "WORD_TARGET=%APPDATA%\Microsoft\Word\STARTUP"
set "EXCEL_TARGET=%APPDATA%\Microsoft\Excel\XLSTART"

:: 0. Word / Excel dang mo thi file add-in bi khoa, khong chep de duoc.
::    Khong tu tat chuong trinh (mat bai dang soan do) - nho nguoi dung luu va dong.
set "LAN_HOI=0"
:KIEM_TRA_DANG_MO
set "DANG_MO=0"
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul && set "DANG_MO=1"
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul && set "DANG_MO=1"
if "%DANG_MO%"=="1" if not "%LAN_HOI%"=="3" (
    set /a LAN_HOI+=1
    echo [!] Word hoac Excel dang mo. Hay LUU bai va DONG het cua so Word / Excel,
    echo     roi bam phim bat ky de tiep tuc...
    pause >nul
    goto KIEM_TRA_DANG_MO
)

set "WORD_OK=0"
set "EXCEL_OK=0"

:: 2. Cai dat Word Add-in
if exist "%WORD_SOURCE%" (
    if not exist "%WORD_TARGET%" mkdir "%WORD_TARGET%" 2>nul
    copy /Y "%WORD_SOURCE%" "%WORD_TARGET%\ChuanHoa_VanBanDang.dotm" >nul 2>&1
    if errorlevel 1 (
        echo [!] Khong chep duoc file add-in Word ^(Word con dang mo, hoac file bi phan mem diet virus giu^).
    ) else (
        powershell -NoProfile -Command "Unblock-File -LiteralPath '%WORD_TARGET%\ChuanHoa_VanBanDang.dotm'" >nul 2>&1
        echo [+] Da cai dat xong Add-in Word: ChuanHoa_VanBanDang.dotm
        set "WORD_OK=1"
    )
) else (
    echo [!] CANH BAO: Khong tim thay file ChuanHoa_VanBanDang.dotm trong thu muc giai nen!
)

:: 3. Cai dat Excel Add-in
if exist "%EXCEL_SOURCE%" (
    if not exist "%EXCEL_TARGET%" mkdir "%EXCEL_TARGET%" 2>nul
    copy /Y "%EXCEL_SOURCE%" "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" >nul 2>&1
    if errorlevel 1 (
        echo [!] Khong chep duoc file add-in Excel ^(Excel con dang mo, hoac file bi phan mem diet virus giu^).
    ) else (
        powershell -NoProfile -Command "Unblock-File -LiteralPath '%EXCEL_TARGET%\ChuanHoa_Excel.xlam'" >nul 2>&1
        echo [+] Da cai dat xong Add-in Excel: ChuanHoa_Excel.xlam
        set "EXCEL_OK=1"
    )
) else (
    echo [!] CANH BAO: Khong tim thay file ChuanHoa_Excel.xlam trong thu muc giai nen!
)

:: 4. Khong ghi Registry: Office da tin cay san thu muc Word\STARTUP va Excel\XLSTART (vi tri mac dinh).
::    Ghi khoa bao mat Office tu file .bat de bi phan mem diet virus giu lai, lam cua so dung im.

echo.
if "%WORD_OK%"=="1" if "%EXCEL_OK%"=="1" (
    echo ======================================================
    echo        CAI DAT THANH CONG TRON BO WORD ^& EXCEL!
    echo ======================================================
    echo.
    echo Hay mo lai Microsoft Word va Excel:
    echo - Word: the Home co nhom [THE THUC VAN BAN], nut [Chuan Hoa The Thuc].
    echo - Excel: the Home co nut [Chuan Hoa Trang In].
    echo Khong can mo truc tiep file .dotm / .xlam.
    echo ======================================================
    pause
    exit /b 0
)

echo ======================================================
echo [!] CHU Y: Cai dat chua du bo (Word=%WORD_OK%, Excel=%EXCEL_OK%).
echo - Neu thay dong "file contains a virus": phan mem diet virus chan nham.
echo   Mo Windows Security - Protection history, chon muc ChuanHoa - Restore,
echo   roi giai nen lai file .zip va chay lai bo cai.
echo - Neu khong: Giai Nen ^(Extract All^) toan bo file .zip ra thu muc
echo   truoc khi nhap dup vao file CaiDat_AddIn_Windows.bat!
echo ======================================================
pause
