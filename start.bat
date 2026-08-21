@echo off
chcp 65001 >nul
cd /d %~dp0

REM 一键启动：自动创建虚拟环境并安装依赖（首次运行较慢）
if not exist .venv (
    echo [1/2] 正在创建虚拟环境并安装依赖...
    python -m venv .venv
    .venv\Scripts\python -m pip install --upgrade pip
    .venv\Scripts\python -m pip install -r requirements.txt
)

echo [2/2] 正在启动 Web 应用，请稍候...
.venv\Scripts\python -m streamlit run app.py
pause
