@echo off
chcp 65001 >nul
echo ========================================
echo 激活 Python 虚拟环境
echo ========================================
echo.
echo 虚拟环境路径: %CD%\venv
echo.
echo 激活后，你可以直接使用:
echo   - python (而不是 py)
echo   - pip (而不是 py -m pip)
echo.
echo 要退出虚拟环境，输入: deactivate
echo ========================================
echo.

call venv\Scripts\activate.bat
