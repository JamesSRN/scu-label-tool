@echo off
title SCU Label Tool
cd /d "%~dp0"
echo ============================================================
echo    SCU LABEL TOOL
echo ============================================================
echo.
echo    This opens the label tool with the latest code applied.
echo    Use this every time you start the tool.
echo.
echo    Make sure the label workbook is CLOSED first.
echo    (If it is already open, just use that window instead.)
echo.
echo    Press any key to open the tool, or close this window to cancel.
echo.
pause >nul
echo.
echo    Opening... follow the Excel pop-ups that appear.
start "" "Build-Release.vbs"
