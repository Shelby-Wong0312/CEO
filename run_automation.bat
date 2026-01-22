@echo off
chcp 65001 >nul
title CEO Automation Tool

echo ============================================================
echo    CEO Project - Automation Tool
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
echo  Step 1: Data Enrichment
echo ============================================================
echo.

python src/enrich_data.py --rows "%ROWS%"

if %ERRORLEVEL% neq 0 (
    echo.
    echo Warning: Data enrichment encountered issues.
    echo.
)

:: Check if photo review HTML exists
if exist "output\data\photo_review.html" (
    echo.
    echo ============================================================
    echo  Step 2: Photo Review (Optional)
    echo ============================================================
    echo.
    echo Photo review report: output\data\photo_review.html
    echo.
    echo [1] Open report, select photos, then continue
    echo [2] Skip photo review, generate PPT directly
    echo.
    set /p PHOTO_CHOICE="Enter choice (1 or 2): "

    if "%PHOTO_CHOICE%"=="1" (
        echo.
        echo Opening photo review report...
        start "" "output\data\photo_review.html"
        echo.
        echo ============================================================
        echo  After reviewing:
        echo   1. Click correct photos in browser
        echo   2. Click "Save Selections" to download JSON
        echo   3. Move JSON to: output\data\photo_selections.json
        echo   4. Press any key to continue
        echo ============================================================
        echo.
        pause
    )
)

echo.
echo ============================================================
echo  Step 3: PowerPoint Generation
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
echo  Output:
echo    - Excel: output\data\Standard_Example_Enriched.xlsx
echo    - PPT:   output\ppt\[Category]\
echo ============================================================
echo.

pause
