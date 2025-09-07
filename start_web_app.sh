#!/bin/bash

# file_coloring_gui_enhanced.py - GUI启动脚本
# 作者: chenhao
# 日期: 2025-09-07
# 功能: 检查依赖、启动GUI脚本并记录进程信息

# 配置参数
VENV_PATH="./venv"  # 虚拟环境路径
GUI_SCRIPT="file_coloring_gui_enhanced.py"  # GUI脚本文件名
PID_FILE="gui_app.pid"  # 进程ID存储文件
LOG_FILE="gui_app.log"  # 运行日志文件

echo "=========================================="
echo "file_coloring_gui_enhanced.py - 启动脚本"
echo "=========================================="

# 1. 检查GUI脚本是否存在
if [ ! -f "$GUI_SCRIPT" ]; then
    echo "错误: 未找到脚本文件 $GUI_SCRIPT"
    echo "请确保脚本在当前目录下"
    exit 1
fi

# 2. 检查是否已在运行
if [ -f "$PID_FILE" ]; then
    EXIST_PID=$(cat "$PID_FILE")
    if ps -p "$EXIST_PID" &> /dev/null; then
        echo "提示: GUI应用已在运行（PID: $EXIST_PID）"
        echo "如需重启，请先执行 ./stop_gui.sh 停止现有进程"
        exit 0
    else
        echo "警告: 发现残留PID文件，正在清理..."
        rm -f "$PID_FILE"
    fi
fi

# 3. 激活虚拟环境（若存在）
if [ -d "$VENV_PATH" ]; then
    echo "激活虚拟环境: $VENV_PATH"
    source "$VENV_PATH/bin/activate"
    if [ $? -ne 0 ]; then
        echo "错误: 虚拟环境激活失败"
        exit 1
    fi
else
    echo "提示: 未找到虚拟环境 $VENV_PATH，将使用系统Python环境"
fi

# 4. 检查必要依赖
echo "检查必要依赖..."
REQUIRED_DEPS=("tkinter" "pandas" "openpyxl")
DEP_MISSING=0

for dep in "${REQUIRED_DEPS[@]}"; do
    if ! python -c "import $dep" &> /dev/null; then
        echo "错误: 未找到依赖包 $dep"
        DEP_MISSING=1
    fi
done

if [ $DEP_MISSING -eq 1 ]; then
    echo "请安装缺失的依赖包:"
    echo "pip install pandas openpyxl"
    # tkinter需单独安装（系统差异）
    if ! python -c "import tkinter" &> /dev/null; then
        echo "提示: tkinter安装方式（根据系统选择）:"
        echo "  Ubuntu/Debian: sudo apt-get install python3-tk"
        echo "  CentOS/RHEL: sudo yum install python3-tkinter"
        echo "  macOS: 已包含在Python安装中（若缺失需重新安装Python）"
    fi
    exit 1
fi

echo "所有依赖检查通过"

# 5. 启动GUI脚本
echo "正在启动GUI应用: $GUI_SCRIPT"
echo "运行日志: $LOG_FILE"
echo "按 Ctrl+C 可停止应用（或执行 ./stop_gui.sh）"
echo "=========================================="

# 启动并记录PID和日志
nohup python "$GUI_SCRIPT" > "$LOG_FILE" 2>&1 &
APP_PID=$!

# 验证启动结果
sleep 2
if ps -p "$APP_PID" &> /dev/null; then
    echo "GUI应用启动成功（PID: $APP_PID）"
    echo "$APP_PID" > "$PID_FILE"  # 保存PID到文件
else
    echo "错误: GUI应用启动失败"
    echo "查看日志获取详情: cat $LOG_FILE"
    exit 1
fi
