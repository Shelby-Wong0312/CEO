@echo off
chcp 65001 >nul
title CEO - Photo Search

echo ============================================================
echo    CEO Project - Photo Search
echo ============================================================
echo.
echo  This tool searches for executive photos and generates
echo  a review report. No API quota is used (DuckDuckGo only).
echo.
echo  NOTE: Please close Excel before running!
echo  TIP: Search 20-30 rows at a time to avoid rate limiting.
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
echo  Step 1: Searching for Photos...
echo ============================================================
echo.

python src/enrich_data.py --rows "%ROWS%" --photos-only

if %ERRORLEVEL% neq 0 (
    echo.
    echo Warning: Photo search encountered issues.
    echo.
)

:: Check if photo review HTML exists
if exist "output\data\photo_review.html" (
    echo.
    echo ============================================================
    echo  Step 2: Photo Review
    echo ============================================================
    echo.
    echo Photo review report: output\data\photo_review.html
    echo.
    echo [1] Open photo review report now
    echo [2] Skip (review later)
    echo.
    set /p OPEN_CHOICE="Enter choice (1 or 2): "

    if "%OPEN_CHOICE%"=="1" (
        echo.
        echo Opening photo review report...
        start "" "output\data\photo_review.html"
        echo.
        echo ============================================================
        echo  After reviewing:
        echo   1. Click correct photos in browser
        echo   2. Click "Save Selections" to download JSON
        echo   3. Move JSON to: output\data\photo_selections.json
        echo   4. Run generate_ppt.bat to create PPT with photos
        echo ============================================================
    )
)

echo.
echo ============================================================
echo  Done!
echo.
echo  Files:
echo    - Excel: output\data\Standard_Example_Enriched.xlsx
echo    - Review: output\data\photo_review.html
echo    - Candidates: output\data\photo_candidates.json
echo.
echo  Next step:
echo    After selecting photos, run generate_ppt.bat
echo ============================================================
echo.

pause
