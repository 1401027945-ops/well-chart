#!/usr/bin/env bash
# 一键启动脚本（Linux / macOS）：自动创建虚拟环境并安装依赖
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
    echo "[1/2] 正在创建虚拟环境并安装依赖..."
    python3 -m venv .venv
    .venv/bin/python -m pip install --upgrade pip
    .venv/bin/python -m pip install -r requirements.txt
fi

echo "[2/2] 正在启动 Web 应用，请稍候..."
exec .venv/bin/python -m streamlit run app.py
