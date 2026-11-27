@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ========================================================
echo   PPT-style 个人 PPT 风格工坊 · 一键安装程序 (Windows)
echo ========================================================
echo.

set "REAL_PY="

for %%P in (
    "%LOCALAPPDATA%\Python\bin\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "C:\Program Files\Python314\python.exe"
    "C:\Program Files\Python313\python.exe"
    "C:\Program Files\Python312\python.exe"
) do (
    if exist %%P (
        "%%~P" -c "import sys; assert sys.version_info[0]==3" >nul 2>&1
        if not errorlevel 1 (
            set "REAL_PY=%%~P"
            goto :found_py
        )
    )
)

for /f "delims=" %%I in ('where python 2^>nul') do (
    echo %%I | findstr /I "WindowsApps" >nul
    if errorlevel 1 (
        "%%I" -c "import sys; assert sys.version_info[0]==3" >nul 2>&1
        if not errorlevel 1 (
            set "REAL_PY=%%I"
            goto :found_py
        )
    )
)

echo [提示] 这台电脑上还没有真实可用的 Python 3 运行环境。
echo.
echo 请按下面两步快速准备：
echo   1. 在浏览器打开 https://www.python.org/downloads/ 下载安装包；
echo   2. 安装时务必勾选底部的「Add Python to PATH」，装完再双击本文件即可。
echo.
pause
exit /b 1

:found_py
echo [1/2] 正在使用运行环境：%REAL_PY%
"%REAL_PY%" -X utf8 "%~dp0install.py"
if errorlevel 1 (
    echo.
    echo [安装失败] 请截图以上信息联系提供该技能包的朋友协助排查。
    pause
    exit /b 1
)

echo.
echo [2/2] 安装已完成！
echo 请【新开一个对话窗口】（或重启你的 AI 软件），先发下面这句做 3 秒验活：
echo     检查 PPT-style 是否就绪
echo.
echo 看到【PPT-style ... 已激活就绪】回复后，再发：帮我建 PPT 风格工坊
echo 如果新开对话后反馈找不到技能，请让当前 AI 检查它实际读取的全局技能目录并同步。
echo.
pause
