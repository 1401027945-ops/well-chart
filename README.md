# 单井历史数据曲线图自动生成工具

基于 Python + Streamlit 的网页应用：以“模板”形式组织作图功能。当前上线
「单井生产曲线模板」：上传单井历史数据 Excel 文件，自动完成数据清洗、
插值修复、曲线图绘制，并导出包含三个子表的 Excel 文件。

后续将陆续上线「日期叠合曲线」和「天数叠合曲线模板」两个模板（网页已预留入口）。

## 功能特性

- 支持 .xls / .xlsx 文件，自动识别表头（兼容标题行 + 表头行格式）和井号；
- 自动清洗：负值替换为 NaN、油压/套压固定值段识别与线性插值修复、
  瞬时气量零值保留（真实关井）；
- 双纵轴曲线图（左：压力 MPa；右：瞬时气量 万方/天）：油压红 / 套压蓝 /
  气量橙，图例居中且只占一行，横轴固定 6 个等时间间隔刻度，无网格线、无图名；
- 输出 Excel 包含三个子表：
  - 「原始数据」：上传文件的原始数据（未清洗）；
  - 「处理后的数据」：清洗/插值后的数据 + 插值标记；
  - 「图片」：曲线图图片 + 原生可编辑 Excel 图表（双击可编辑、修改数据后自动更新）；
- 数据量不足、时间列缺失、全空数据等边缘情况均有提示。

## 快速启动

### Windows

双击 `start.bat`，或命令行执行：

```bat
start.bat
```

### Linux / macOS

```bash
chmod +x start.sh
./start.sh
```

### 手动启动

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt   # Windows
# 或 .venv/bin/pip install -r requirements.txt   # Linux/macOS
.venv\Scripts\python -m streamlit run app.py
```

浏览器会自动打开 http://localhost:8501 。

## Docker 部署（可选）

```bash
docker build -t well-chart .
docker run -p 8501:8501 well-chart
```

## 输入文件格式

| 列 | 说明 |
| --- | --- |
| 日期 | 时间列，格式不统一可自动解析（如 2026-01-01 18:00:00、2025/07/30 19:50） |
| 油压 | 数值，MPa |
| 套压 | 数值，MPa |
| 瞬时气量 | 数值，万方/天（零值保留，视为关井） |

示例文件仅保留在本地 `sample_data/` 目录中用于测试，不随代码上传到公开仓库。

## 数据处理规则摘要

- 油压/套压/瞬时气量为负 → NaN；
- 油压/套压连续不变 ≥6 个点或 ≥24 小时 → 无效段，线性插值修复；
- 瞬时气量全为 0 → 纵轴 0~1 并标注“关井”；
- 时间范围超过 1 年 → 按季度刻度；小于 1 天 → 每 2 小时刻度。

## 运行单元测试

```bash
.venv\Scripts\python tests\test_all.py
```

## 部署为公网网址（可选）

应用本身就是一个网页，部署后任何人通过网址访问即可使用。

### 方式一：Streamlit Community Cloud（免费，推荐）

1. 把本项目推送到 GitHub 仓库；
2. 打开 https://share.streamlit.io 并用 GitHub 账号登录；
3. 点击 “New app”，选择本仓库、分支与入口文件 `app.py`；
4. 部署完成后会得到类似 `https://你的名字-项目名.streamlit.app` 的网址。

完整的图文操作步骤见 [部署手册.md](部署手册.md)。

### 方式二：Hugging Face Spaces

1. 在 https://huggingface.co 新建 Space，选择 SDK 为 Streamlit；
2. 把项目文件推送到该 Space 仓库；
3. 自动部署完成，得到 `https://huggingface.co/spaces/你的名字/空间名`。

### 方式三：自己的服务器 / 内网

```bash
docker build -t well-chart .
docker run -d -p 8501:8501 well-chart
```

访问 `http://服务器IP:8501`。如需公网，可配合 Nginx 反向代理并绑定域名。

## 目录结构

```text
app.py                 # Streamlit 主程序
well_chart/
  loader.py            # 文件解析与表头/井号识别
  cleaning.py          # 数据清洗与插值
  plotting.py          # 曲线图绘制与刻度计算
  excel_export.py      # Excel 导出
  config.py            # 全局配置
tests/                 # 单元测试
requirements.txt       # 依赖清单
start.bat / start.sh   # 一键启动脚本
Dockerfile             # 容器部署（可选）
```
