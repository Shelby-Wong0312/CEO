@echo off
chcp 65001 >nul
title CEO - PPT Generation Only

echo ============================================================
echo    CEO Project - PPT Generation Only
echo ============================================================
echo.
echo  Use this when you've already enriched data or made manual
echo  edits, and only need to regenerate the PPT files.
echo.
echo ============================================================
echo.

:: Prompt user for row numbers
set /p ROWS="Enter row numbers (e.g., 2, 5-10, 15): "

:: Validate input
if "%ROWS%"=="" (
    echo Error: No row numbers provided.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Generating PowerPoint files...
echo ============================================================
echo.
echo (Auto-applying photo selections if photo_selections.json exists)
echo.

python src/generate_ppt.py --rows "%ROWS%"

if %ERRORLEVEL% neq 0 (
    echo.
    echo Warning: PowerPoint generation encountered issues.
    echo.
)

echo.
echo ============================================================
echo  Done!
echo.
echo  Output: output\ppt\[Category]\
echo ============================================================
echo.

pause
