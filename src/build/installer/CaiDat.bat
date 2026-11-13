@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion
title Cai dat bo Add-in NTU Thesis

set "SCRIPT_DIR=%~dp0"
set "WORD_TARGET=%APPDATA%\Microsoft\Word\STARTUP"
set "EXCEL_TARGET=%APPDATA%\Microsoft\Excel\XLSTART"

echo ==========================================================================
echo   BO CAI DAT ADD-IN NTU THESIS - Truong Dai hoc Nha Trang
echo ==========================================================================
echo.
echo   Gom 3 add-in:
echo     1. The thuc van ban Word       -^> tab [THE THUC VAN BAN]
echo     2. Chuan hoa luan an NTU      -^> tab [NTU Thesis]
echo     3. Chuan hoa trang in A4 Excel-^> tab [CHUAN HOA IN]
echo.
echo   Word : %WORD_TARGET%
echo   Excel: %EXCEL_TARGET%
echo.

for %%F in (GiaTwinH_NTU_Thesis_Word.dotm NTUThesis.dotm ChuanHoa_Excel.xlam) do (
    if not exist "%SCRIPT_DIR%%%F" goto THIEU_FILE
)

set "WORD_BUSY=0"
set "EXCEL_BUSY=0"
set "LAN_HOI=0"

:KIEM_TRA
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul && set "WORD_BUSY=1" || set "WORD_BUSY=0"
tasklist /FI "IMAGENAME eq EXCEL.EXE"  2>nul | find /I "EXCEL.EXE"  >nul && set "EXCEL_BUSY=1" || set "EXCEL_BUSY=0"

if "!WORD_BUSY!!EXCEL_BUSY!"=="00" goto CAI_DAT

set /a LAN_HOI+=1
if !LAN_HOI! GEQ 3 goto HOI_BO_QUA
echo   [!] Van con ung dung dang mo:
if "!WORD_BUSY!"=="1"  echo       - Microsoft Word
if "!EXCEL_BUSY!"=="1" echo       - Microsoft Excel
echo       Hay DONG het roi bam phim bat ky de thu lai... (lan !LAN_HOI!/3)
pause >nul
goto KIEM_TRA

:HOI_BO_QUA
echo.
echo   [x] Da thu 3 lan nhung ung dung van dang mo.
echo       Khong the ghi de add-in khi Word/Excel dang chay.
echo       Hay dong han Word va Excel, roi chay lai CaiDat.bat.
echo.
pause
exit /b 1

:CAI_DAT
echo   [1/5] Chuan bi thu muc...
if not exist "%WORD_TARGET%"  mkdir "%WORD_TARGET%"  >nul 2>&1
if not exist "%EXCEL_TARGET%" mkdir "%EXCEL_TARGET%" >nul 2>&1

echo   [2/5] Go ban cu neu co...
for %%F in (GiaTwinH_NTU_Thesis_Word.dotm NTUThesis.dotm ChuanHoa_VanBanDang.dotm) do (
    if exist "%WORD_TARGET%\%%F" del /F /Q "%WORD_TARGET%\%%F" >nul 2>&1
)
if exist "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" del /F /Q "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" >nul 2>&1

echo   [3/5] Sao chep add-in...
set "WORD_OK=0"
set "NTU_OK=0"
set "EXCEL_OK=0"
copy /Y "%SCRIPT_DIR%GiaTwinH_NTU_Thesis_Word.dotm" "%WORD_TARGET%\" >nul 2>&1 && set "WORD_OK=1"
copy /Y "%SCRIPT_DIR%NTUThesis.dotm"               "%WORD_TARGET%\" >nul 2>&1 && set "NTU_OK=1"
copy /Y "%SCRIPT_DIR%ChuanHoa_Excel.xlam"          "%EXCEL_TARGET%\" >nul 2>&1 && set "EXCEL_OK=1"

echo   [4/5] Go danh sach chan cua Office...
if exist "%SCRIPT_DIR%DonDepBlacklist.vbs" (
    cscript //nologo "%SCRIPT_DIR%DonDepBlacklist.vbs" >nul 2>&1
) else (
    echo         (khong tim thay DonDepBlacklist.vbs - bo qua buoc nay^)
)

echo   [5/5] Bo chan Mark-of-the-Web...
powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%WORD_TARGET%' -Filter '*.dotm' -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1
powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%EXCEL_TARGET%' -Filter '*.xlam' -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue" >nul 2>&1

echo.
echo ==========================================================================
echo   KET QUA
echo ==========================================================================
if "!WORD_OK!"=="1"  (echo   [OK] The thuc van ban Word        -^> tab [THE THUC VAN BAN]) else (echo   [!!] The thuc van ban Word        -^> CAI THAT BAI)
if "!NTU_OK!"=="1"   (echo   [OK] Chuan hoa luan an NTU       -^> tab [NTU Thesis])        else (echo   [!!] Chuan hoa luan an NTU       -^> CAI THAT BAI)
if "!EXCEL_OK!"=="1" (echo   [OK] Chuan hoa trang in A4 Excel -^> tab [CHUAN HOA IN])     else (echo   [!!] Chuan hoa trang in A4 Excel -^> CAI THAT BAI)
echo.
echo   MO LAI Word va Excel de su dung. Lan dau mo co the cham hon binh thuong.
echo.
echo   Neu Windows Security bao "file contains a virus" hoac file bi cach ly:
echo     Windows Security -^> Virus and threat protection
echo     -^> Protection history -^> chon file -^> Actions -^> Restore
echo     roi chay lai CaiDat.bat.
echo.
echo   Go cai dat : GoCaiDat.bat
echo   Huong dan  : HuongDan_CaiDat.txt
echo ==========================================================================
echo.
pause
exit /b 0

:THIEU_FILE
echo.
echo   [LOI] Thieu file add-in trong thu muc nay.
echo         Can du 3 file:
echo           GiaTwinH_NTU_Thesis_Word.dotm
echo           NTUThesis.dotm
echo           ChuanHoa_Excel.xlam
echo.
echo         Ban da giai nen TOAN BO file zip ra mot thu muc chua?
echo         Hay lam lai: chuot phai file zip -^> Extract All -^> chay CaiDat.bat
echo.
pause
exit /b 1