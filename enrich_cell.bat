@echo off
chcp 65001 >nul
title CEO - 針對特定儲存格資料收集

echo ============================================================
echo    CEO Project - 針對特定儲存格資料收集
echo ============================================================
echo.
echo  此工具可針對特定欄位和列號收集資料，並生成 PPT。
echo.
echo  NOTE: 請先關閉 Excel！
echo.
echo ============================================================
echo  可用欄位:
echo ============================================================
echo.
echo   編號  欄位代號  欄位名稱
echo   ----  --------  ----------------
echo    1      C       年齡
echo    2      F       專業分類
echo    3      G       專業背景
echo    4      H       學歷
echo    5      I       主要經歷
echo    6      J       現職/任
echo    7      K       個人特質
echo    8      L       現擔任獨董家數(年)
echo    9      M       擔任獨董年資(年)
echo   10      N       電子郵件
echo   11      O       公司電話
echo   12      D       照片
echo.
echo ============================================================
echo  輸入格式範例:
echo ============================================================
echo.
echo   儲存格參照:  H26, H26-H30, H26,I27,J28
echo   欄位+列號:   學歷:26, 學歷:26-30
echo.
echo ============================================================
echo.

:: 選擇輸入模式
echo [1] 使用儲存格參照 (如 H26)
echo [2] 使用欄位名稱 + 列號
echo.
set /p MODE="請選擇輸入模式 (1 或 2): "

if "%MODE%"=="1" (
    echo.
    set /p CELLS="請輸入儲存格參照 (如 H26, H26-H30): "

    if "!CELLS!"=="" (
        echo 錯誤: 未輸入儲存格參照
        pause
        exit /b 1
    )

    echo.
    echo ============================================================
    echo  開始收集資料...
    echo ============================================================
    echo.

    python src/enrich_cell.py --cell "%CELLS%"

) else if "%MODE%"=="2" (
    echo.
    set /p FIELD="請輸入欄位編號或名稱 (如 4 或 學歷 或 H): "
    set /p ROWS="請輸入列號 (如 26, 26-30, 26,27,28): "

    if "!FIELD!"=="" (
        echo 錯誤: 未輸入欄位
        pause
        exit /b 1
    )
    if "!ROWS!"=="" (
        echo 錯誤: 未輸入列號
        pause
        exit /b 1
    )

    echo.
    echo ============================================================
    echo  開始收集資料...
    echo ============================================================
    echo.

    python src/enrich_cell.py --field "%FIELD%" --rows "%ROWS%"

) else (
    echo 錯誤: 無效的選擇
    pause
    exit /b 1
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo 資料收集遇到問題
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  是否要生成 PPT?
echo ============================================================
echo.
echo [1] 是，生成 PPT
echo [2] 否，稍後再生成
echo.
set /p PPT_CHOICE="請選擇 (1 或 2): "

if "%PPT_CHOICE%"=="1" (
    echo.
    echo ============================================================
    echo  正在生成 PPT...
    echo ============================================================
    echo.

    :: 從儲存格參照中提取列號
    if "%MODE%"=="1" (
        :: 使用 Python 提取列號
        for /f "tokens=*" %%i in ('python -c "import re; cells='%CELLS%'; rows=set(re.findall(r'\d+', cells)); print(','.join(sorted(rows, key=int)))"') do set EXTRACTED_ROWS=%%i
        python src/generate_ppt.py --rows "%EXTRACTED_ROWS%"
    ) else (
        python src/generate_ppt.py --rows "%ROWS%"
    )

    if %ERRORLEVEL% neq 0 (
        echo.
        echo PPT 生成遇到問題
    ) else (
        echo.
        echo PPT 生成完成！
    )
)

echo.
echo ============================================================
echo  完成！
echo.
echo  檔案位置:
echo    - Excel: output\data\Standard_Example_Enriched.xlsx
echo    - PPT:   output\ppt\[專業分類]\[姓名]_CV.pptx
echo ============================================================
echo.

pause
