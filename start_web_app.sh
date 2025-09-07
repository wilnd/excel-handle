#!/bin/bash

# 文件着色分析工具 - 本地启动脚本
# 作者: chenhao
# 日期: 2025-08-17

echo "=========================================="
echo "文件着色分析工具 - Web版启动脚本"
echo "=========================================="

# 检查是否已经安装了必要的依赖
echo "检查依赖..."
if ! python -c "import flask" &> /dev/null; then
    echo "错误: 未找到Flask，请先安装Flask"
    echo "运行命令: pip install flask"
    exit 1
fi

if ! python -c "import pandas" &> /dev/null; then
    echo "错误: 未找到pandas，请先安装pandas"
    echo "运行命令: pip install pandas"
    exit 1
fi

if ! python -c "import openpyxl" &> /dev/null; then
    echo "错误: 未找到openpyxl，请先安装openpyxl"
    echo "运行命令: pip install openpyxl"
    exit 1
fi

echo "所有依赖检查通过"

# 检查web_app.py文件是否存在
if [ ! -f "web_app.py" ]; then
    echo "错误: 未找到web_app.py文件"
    exit 1
fi

# 启动Web应用
echo "正在启动Web应用..."
echo "访问地址: http://127.0.0.1:9080"
echo "日志文件: web_app.log"
echo "按 Ctrl+C 停止应用"
echo "=========================================="

# 启动应用
python web_app.py

echo "应用已停止"
echo "日志已保存到 web_app.log 文件中"

echo "应用已停止"