#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$APP_DIR"

ENV_FILE="$APP_DIR/runtime.env"
[ -f "$ENV_FILE" ] && source "$ENV_FILE" || true

PORT="${PORT:-9090}"
WEB_WORKERS="${WEB_WORKERS:-2}"
BIND_ADDR="${BIND_ADDR:-0.0.0.0}"

export PORT MAX_CONTENT_LENGTH_MB SECRET_KEY

VENV_DIR="$APP_DIR/venv"
PID_FILE="$APP_DIR/gunicorn.pid"
LOG_DIR="$APP_DIR/logs"
ACCESS_LOG="$LOG_DIR/gunicorn_access.log"
ERROR_LOG="$LOG_DIR/gunicorn_error.log"

mkdir -p "$LOG_DIR" "$APP_DIR/uploads"

if [ ! -d "$VENV_DIR" ]; then
  python3 -m venv "$VENV_DIR"
fi

# shellcheck disable=SC1090
source "$VENV_DIR/bin/activate"

pip install --upgrade pip >/dev/null
pip install -r requirements.txt

# 若已在跑，先退出
if [ -f "$PID_FILE" ] && ps -p "$(cat "$PID_FILE")" > /dev/null 2>&1; then
  echo "Gunicorn 已在运行 (PID $(cat "$PID_FILE"))，无需重复启动。"
  exit 0
fi

echo "启动 Web 应用：Gunicorn（端口 $PORT，workers=$WEB_WORKERS）"

# --daemon 后台运行；--pid 写 PID 文件
gunicorn \
  --daemon \
  --workers "$WEB_WORKERS" \
  --bind "$BIND_ADDR:$PORT" \
  --pid "$PID_FILE" \
  --access-logfile "$ACCESS_LOG" \
  --error-logfile "$ERROR_LOG" \
  --timeout 600 \
  app:app

sleep 1
if [ -f "$PID_FILE" ] && ps -p "$(cat "$PID_FILE")" > /dev/null 2>&1; then
  echo "✅ 已启动。PID=$(cat "$PID_FILE")"
  echo "访问地址: http://$BIND_ADDR:$PORT/ （若有反代请用域名）"
else
  echo "❌ 启动失败，请查看日志：$ERROR_LOG"
  exit 1
fi
