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
WHEELS_DIR="$APP_DIR/.wheels"

mkdir -p "$LOG_DIR" "$APP_DIR/uploads" "$WHEELS_DIR"

if [ ! -d "$VENV_DIR" ]; then
  python3 -m venv "$VENV_DIR"
fi

# shellcheck disable=SC1090
source "$VENV_DIR/bin/activate"

# ---------- PIP 加速配置 ----------
RETRIES="${PIP_RETRIES:-10}"
TIMEOUT="${PIP_TIMEOUT:-300}"

# 镜像优先队列（可在 runtime.env 自定义 PIP_INDEX_URLS，用空格分隔）
if [ -n "${PIP_INDEX_URLS:-}" ]; then
  read -r -a MIRRORS <<<"$PIP_INDEX_URLS"
else
  MIRRORS=(
    "https://pypi.tuna.tsinghua.edu.cn/simple"
    "https://mirrors.aliyun.com/pypi/simple"
    "https://mirrors.cloud.tencent.com/pypi/simple"
    "https://repo.huaweicloud.com/repository/pypi/simple"
  )
fi

# 可选：走代理（按需在 runtime.env 里设定）
# export HTTP_PROXY=${HTTP_PROXY:-http://127.0.0.1:7890}
# export HTTPS_PROXY=${HTTPS_PROXY:-http://127.0.0.1:7890}

echo "升级 pip..."
python -m pip install --upgrade pip >/dev/null 2>&1 || true

try_install_from_mirrors() {
  local ok=0
  for url in "${MIRRORS[@]}"; do
    echo "尝试使用镜像: $url"
    if pip install -r requirements.txt \
        -i "$url" \
        --retries "$RETRIES" \
        --timeout "$TIMEOUT" \
        --prefer-binary; then
      ok=1
      break
    else
      echo "从镜像 $url 安装失败，继续尝试下一个镜像..."
    fi
  done
  return $ok
}

offline_fallback() {
  echo "尝试『预下载 → 离线安装』兜底方案..."
  local ok=0
  for url in "${MIRRORS[@]}"; do
    echo "预下载（镜像: $url）到 $WHEELS_DIR ..."
    if pip download -r requirements.txt \
         -d "$WHEELS_DIR" \
         -i "$url" \
         --retries "$RETRIES" \
         --timeout "$TIMEOUT" \
         --prefer-binary; then
      ok=1
      break
    else
      echo "从镜像 $url 预下载失败，尝试下一个镜像..."
    fi
  done

  if [ "$ok" -eq 1 ]; then
    echo "离线安装（使用本地 wheels 缓存）..."
    pip install --no-index --find-links "$WHEELS_DIR" -r requirements.txt --prefer-binary
  else
    echo "预下载未成功，尝试直接使用官方源（可能较慢）..."
    pip install -r requirements.txt --retries "$RETRIES" --timeout "$TIMEOUT" --prefer-binary
  fi
}

echo "安装依赖（带镜像加速/重试/兜底）..."
if ! try_install_from_mirrors; then
  echo "镜像安装失败，进入兜底流程。"
  offline_fallback
fi

# ---------- 如果已在跑，直接退出 ----------
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
