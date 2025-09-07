#!/bin/bash

# 文件着色分析工具 - Web版启动脚本
# 作者: chenhao
# 日期: 2025-09-07
# 功能: 检查依赖、激活虚拟环境、启动Web应用并记录进程ID

# 定义常量
VENV_PATH="./venv/bin/activate"
APP_FILE="web_app.py"
LOG_FILE="web_app.log"
PID_FILE="web_app.pid"
PORT="9080"

# 检查虚拟环境
if [ ! -f "$VENV_PATH" ]; then
    echo "错误: 虚拟环境不存在，请先创建虚拟环境"
    echo "推荐命令: python -m venv venv"
    exit 1
fi

# 激活虚拟环境
source "$VENV_PATH"

echo "=========================================="
echo "文件着色分析工具 - Web版启动脚本"
echo "=========================================="

# 检查必要依赖
check_dependency() {
    local pkg_name=$1
    local install_cmd=$2
    if ! python -c "import $pkg_name" &> /dev/null; then
        echo "错误: 未找到依赖包 $pkg_name"
        echo "请运行: $install_cmd"
        exit 1
    fi
}

echo "检查依赖包..."
check_dependency "flask" "pip install flask"
check_dependency "pandas" "pip install pandas"
check_dependency "openpyxl" "pip install openpyxl"
echo "所有依赖检查通过"

# 检查应用文件
if [ ! -f "$APP_FILE" ]; then
    echo "错误: 应用文件 $APP_FILE 不存在"
    exit 1
fi

# 检查是否已在运行
if [ -f "$PID_FILE" ]; then
    local existing_pid=$(cat "$PID_FILE")
    if ps -p "$existing_pid" &> /dev/null; then
        echo "提示: 应用已在运行（PID: $existing_pid）"
        echo "访问地址: http://127.0.0.1:$PORT"
        exit 0
    else
        echo "警告: 发现残留PID文件，已自动清理"
        rm -f "$PID_FILE"
    fi
fi

# 启动应用
echo "正在启动Web应用..."
nohup python "$APP_FILE" > "$LOG_FILE" 2>&1 &
APP_PID=$!

# 验证启动是否成功
sleep 2
if ps -p "$APP_PID" &> /dev/null; then
    # 保存PID到文件
    echo "$APP_PID" > "$PID_FILE"
    echo "应用启动成功！"
    echo "=========================================="
    echo "进程ID: $APP_PID"
    echo "访问地址: http://127.0.0.1:$PORT"
    echo "日志文件: $LOG_FILE"
    echo "PID文件: $PID_FILE"
    echo "停止命令: ./stop_app.sh"
    echo "=========================================="
else
    echo "错误: 应用启动失败，请查看日志文件: $LOG_FILE"
    exit 1
fi
