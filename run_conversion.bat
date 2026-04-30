@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ╔════════════════════════════════════════════════════════════╗
echo ║           HTML 幻灯片批量转换工具                           ║
echo ╚════════════════════════════════════════════════════════════╝
echo.

if not exist "venv\Scripts\python.exe" (
    echo ❌ 虚拟环境不存在，请先运行: py -m venv venv
    pause & exit /b 1
)

venv\Scripts\playwright.exe install chromium >nul 2>&1

set "html_dir=01042026\01042026"
set "success=0"
set "fail=0"
set "total=0"

for %%f in ("%html_dir%\*.html") do set /a total+=1
echo 找到 %total% 个 HTML 文件，开始转换...
echo.

for %%f in ("%html_dir%\*.html") do (
    echo 正在转换: %%~nxf
    venv\Scripts\python.exe convert_slides.py "%%f"
    if !ERRORLEVEL! EQU 0 (
        set /a success+=1
    ) else (
        set /a fail+=1
        echo   ❌ 失败: %%~nxf
    )
)

echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo 完成：%success% 成功 / %fail% 失败 / %total% 总计
echo 输出目录：output\
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.
pause
