@echo off
setlocal
chcp 65001 >nul

cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo [错误] 未检测到 Node.js。
  echo 请先安装 Node.js（LTS 版本），再双击本脚本启动系统。
  pause
  exit /b 1
)

echo 正在启动本地数据服务...
start "顺源物流本地服务" cmd /c "cd /d ""%~dp0"" && node local-server.js"

timeout /t 1 /nobreak >nul
start "" "http://127.0.0.1:5173"

echo.
echo 已启动：请在浏览器中确认顶部显示“本地文件模式”。
echo 数据将保存到当前目录的 data 文件夹内。
echo.
pause
