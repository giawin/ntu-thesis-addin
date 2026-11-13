@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion
title Go cai dat bo Add-in NTU Thesis

set "SCRIPT_DIR=%~dp0"
set "WORD_TARGET=%APPDATA%\Microsoft\Word\STARTUP"
set "EXCEL_TARGET=%APPDATA%\Microsoft\Excel\XLSTART"

echo ==========================================================================
echo   GO CAI DAT ADD-IN NTU THESIS
echo ==========================================================================
echo.

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul && (
    echo   [!] Word dang chay. Hay dong Word truoc khi go cai dat.
    pause
    exit /b 1
)
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul && (
    echo   [!] Excel dang chay. Hay dong Excel truoc khi go cai dat.
    pause
    exit /b 1
)

echo   Dang xoa add-in khoi thu muc khoi dong...
rem ChuanHoa_VanBanDang.dotm la ten add-in GOC cua tac gia - van xoa de don
rem sach neu nguoi dung da tung cai ban goc truoc day.
for %%F in (GiaTwinH_NTU_Thesis_Word.dotm NTUThesis.dotm ChuanHoa_VanBanDang.dotm) do (
    if exist "%WORD_TARGET%\%%F" (
        del /F /Q "%WORD_TARGET%\%%F" >nul 2>&1
        echo     - da xoa %WORD_TARGET%\%%F
    )
)
if exist "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" (
    del /F /Q "%EXCEL_TARGET%\ChuanHoa_Excel.xlam" >nul 2>&1
    echo     - da xoa %EXCEL_TARGET%\ChuanHoa_Excel.xlam
)

echo.
echo.
echo   Dang xoa danh sach chan cua Office (neu co)...
if exist "%SCRIPT_DIR%DonDepBlacklist.vbs" (
    cscript //nologo "%SCRIPT_DIR%DonDepBlacklist.vbs" >nul 2>&1
) else (
    echo     (khong tim thay DonDepBlacklist.vbs - bo qua buoc nay^)
)

echo   Da go cai dat. Mo lai Word/Excel de xac nhan cac tab khong con hien.
echo.
pause
exit /b 0