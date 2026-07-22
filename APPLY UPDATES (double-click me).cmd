@echo off
title SCU Label Tool - Apply Updates
cd /d "%~dp0"
echo ============================================================
echo    SCU LABEL TOOL  -  APPLY UPDATES
echo ============================================================
echo.
echo    This applies the latest code to the label workbook.
echo    (It runs Build-Release.vbs for you.)
echo.
echo    BEFORE YOU CONTINUE:
echo      1. Close MedicationDispensing.xlsm if it is open.
echo         (Otherwise the build can wipe the label logo.)
echo.
echo    Press any key to apply the updates,
echo    or close this window to cancel.
echo.
pause >nul
echo.
echo    Applying updates... follow the Excel pop-ups that appear.
start "" "Build-Release.vbs"
