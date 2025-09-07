@echo off
title 文件着色分析工具 - Web版启动脚本

echo ==========================================
echo 文件着色分析工具 - Web版启动脚本
echo ==========================================

echo 检查依赖...
python -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Flask，请先安装Flask
    echo 运行命令: pip install flask
    pause
    exit /b 1
)

python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到pandas，请先安装pandas
    echo 运行命令: pip install pandas
    pause
    exit /b 1
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到openpyxl，请先安装openpyxl
    echo 运行命令: pip install openpyxl
    pause
    exit /b 1
)

echo 所有依赖检查通过

if not exist "web_app.py" (
    echo 错误: 未找到web_app.py文件
    pause
    exit /b 1
)

echo 正在启动Web应用...
echo 访问地址: http://127.0.0.1:9080
echo 日志文件: web_app.log
echo 按 Ctrl+C 停止应用
echo ==========================================

python web_app.py

echo 应用已停止
pause