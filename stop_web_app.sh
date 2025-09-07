#!/bin/bash

# file_coloring_gui_enhanced.py - GUI停止脚本
# 作者: chenhao
# 日期: 2025-09-07
# 功能: 停止GUI进程、清理PID文件和临时记录

# 配置参数（与启动脚本保持一致）
PID_FILE="gui_app.pid"
LOG_FILE="gui_app.log"
GUI_SCRIPT="file_coloring_gui_enhanced.py"

echo "=========================================="
echo "file_coloring_gui_enhanced.py - 停止脚本"
echo "=========================================="

# 1. 检查PID文件是否存在
if [ ! -f "$PID_FILE" ]; then
    echo "提示: 未找到PID文件 $PID_FILE"

    # 检查是否有残留的GUI进程
    RESIDUAL_PIDS=$(ps aux | grep "python $GUI_SCRIPT" | grep -v grep | awk '{print $2}')
    if [ -n "$RESIDUAL_PIDS" ]; then
        echo "发现残留GUI进程: $RESIDUAL_PIDS"
        read -p "是否强制停止这些进程? (y/n): " CONFIRM
        if [ "$CONFIRM" = "y" ] || [ "$CONFIRM" = "Y" ]; then
            kill -9 $RESIDUAL_PIDS
            echo "已强制停止残留进程"
        else
            echo "操作取消，残留进程未处理"
            exit 0
        fi
    else
        echo "未检测到运行中的GUI应用"
        exit 0
    fi
fi

# 2. 读取PID并验证进程状态
APP_PID=$(cat "$PID_FILE")
if ! ps -p "$APP_PID" &> /dev/null; then
    echo "警告: PID文件记录的进程 $APP_PID 已不存在"
    echo "正在清理残留PID文件..."
    rm -f "$PID_FILE"
    exit 0
fi

# 3. 尝试优雅停止进程
echo "正在停止GUI应用（PID: $APP_PID）..."
kill "$APP_PID"

# 等待进程退出（最多等待10秒）
WAIT_TIME=0
while ps -p "$APP_PID" &> /dev/null && [ $WAIT_TIME -lt 10 ]; do
    sleep 1
    WAIT_TIME=$((WAIT_TIME + 1))
    echo -n "."
done

# 4. 处理未正常退出的情况
if ps -p "$APP_PID" &> /dev/null; then
    echo -e "\n警告: 进程未能优雅退出，正在强制停止..."
    kill -9 "$APP_PID"
    sleep 1
fi

# 5. 清理文件
if [ -f "$PID_FILE" ]; then
    rm -f "$PID_FILE"
    echo "已清理PID文件"
fi

echo "=========================================="
echo "GUI应用已成功停止"
echo "最后运行日志: $LOG_FILE（如需查看请执行 cat $LOG_FILE）"
echo "启动命令: ./start_gui.sh"
echo "=========================================="
