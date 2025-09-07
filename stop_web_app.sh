#!/bin/bash

# 文件着色分析工具 - Web版停止脚本
# 作者: chenhao
# 日期: 2025-09-07
# 功能: 停止Web应用进程并清理PID文件

# 定义常量
PID_FILE="web_app.pid"
LOG_FILE="web_app.log"

echo "=========================================="
echo "文件着色分析工具 - Web版停止脚本"
echo "=========================================="

# 检查PID文件是否存在
if [ ! -f "$PID_FILE" ]; then
    echo "提示: 未找到PID文件 $PID_FILE"
    echo "应用可能未在运行"

    # 额外检查是否有残留进程
    local residual_pid=$(ps aux | grep "python web_app.py" | grep -v grep | awk '{print $2}')
    if [ -n "$residual_pid" ]; then
        echo "发现残留进程: $residual_pid"
        read -p "是否强制停止这些进程? (y/n): " confirm
        if [ "$confirm" = "y" ] || [ "$confirm" = "Y" ]; then
            kill -9 "$residual_pid"
            echo "已强制停止残留进程: $residual_pid"
        else
            echo "已取消操作"
            exit 0
        fi
    fi

    exit 0
fi

# 读取PID并检查进程状态
APP_PID=$(cat "$PID_FILE")
if ! ps -p "$APP_PID" &> /dev/null; then
    echo "警告: PID文件存在，但进程 $APP_PID 已不存在"
    echo "正在清理残留PID文件..."
    rm -f "$PID_FILE"
    exit 0
fi

# 停止进程
echo "正在停止应用进程 (PID: $APP_PID)..."
kill "$APP_PID"

# 等待进程退出
local wait_time=0
while ps -p "$APP_PID" &> /dev/null && [ $wait_time -lt 10 ]; do
    sleep 1
    wait_time=$((wait_time + 1))
    echo -n "."
done

# 检查停止结果
if ps -p "$APP_PID" &> /dev/null; then
    echo -e "\n警告: 进程未能正常退出，正在强制停止..."
    kill -9 "$APP_PID"
    sleep 1
fi

# 清理PID文件
if [ -f "$PID_FILE" ]; then
    rm -f "$PID_FILE"
    echo "已清理PID文件"
fi

echo "=========================================="
echo "应用已成功停止"
echo "最后运行日志: $LOG_FILE"
echo "启动命令: ./start_app.sh"
echo "=========================================="
