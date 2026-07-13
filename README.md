# ^60Co γ 射线累积辐照电学特性数据分析

[![Release](https://img.shields.io/github/v/release/dude1wudv/radiation-data-analysis?display_name=tag)](https://github.com/dude1wudv/radiation-data-analysis/releases/latest)
[![Python](https://img.shields.io/badge/Python-3.9-blue)](https://www.python.org/)

面向 ^60Co γ 射线 **0–60 krad(Si)** 累积总电离剂量实验的可追溯数据处理仓库。项目覆盖原始 Excel 抽取、质量控制、转移/输出参数统一重算、Origin 制图以及中文技术报告交付。

## 正式报告

- [下载最新版 PDF（GitHub Release）](https://github.com/dude1wudv/radiation-data-analysis/releases/latest/download/report.pdf)
- [在线查看 Markdown 报告](deliverables/04_manuscript/report.md)
- [仓库内 PDF](deliverables/04_manuscript/report.pdf)
- [可编辑 DOCX](deliverables/04_manuscript/report.docx)

> 本仓库提供的是完整技术报告及可复现分析材料，不代表已经完成同行评审或正式期刊发表。

## 数据与结果概览

| 项目 | 数量或结果 |
| --- | --- |
| 原始工作簿 | 28 个 |
| 数据工作表 | 97 个（转移 42、输出 55） |
| 原始数据点 | 转移 14,322 行、输出 8,305 行 |
| 剂量点 | 0、10、20、30、40、50、60 krad(Si) |
| 跟踪器件 | S1–S3、T1–T3 |
| 正式图件 | 13 组，提供 PNG/PDF/EMF |

60 krad(Si) 时的主要观测结果：

- S、T 组阈值漂移分别为 **−6.120±0.045 V** 和 **−9.969±0.337 V**；
- 两组峰值跨导相对基线均下降约 **60%**；
- S、T 组 30 V 漏电相对基线平均增加约 **1840%** 和 **60.2%**；
- S 组 13 条后续加压扫描的 30 V 漏电降低 **70.65%–96.19%**，表现出测量历史相关恢复现象。

这些结果仅描述本批纵向跟踪器件。S/T 是样品组标签，仓库不披露或反推器件结构、材料、尺寸、制造工艺和封装差异。

## 数据处理流程

```mermaid
flowchart LR
    A[28 个原始 Excel] --> B[只读抽取与来源校验]
    B --> C[质量控制]
    C --> D[统一参数重算]
    D --> E[汇总 CSV]
    E --> F[Origin 图件]
    E --> G[中文技术报告]
```

原始 Excel 的 GM、VT 列包含 42 个错误标记，因此正式结果不直接使用这些单元格，而是由 `DrainI–GateV` 原始列统一重算：

- 漏极电流采用 11 点二阶 Savitzky–Golay 平滑；
- 峰值跨导由平滑电流对栅压的数值梯度提取；
- 阈值电压采用峰值跨导切线外推法；
- 亚阈值摆幅采用 21 点滑动线性拟合，并要求 `R² ≥ 0.98`；
- 输出曲线提取 10/20/30 V 电流、低场电导和指定电流对应电压；
- 约 10 mA 合规限流区不作为剂量响应指标。

## 仓库结构

```text
.
├── analysis_config.json          # 实验条件与参数提取配置
├── scripts/
│   ├── extract_excel.py          # Excel COM 只读抽取与质控
│   ├── analyze_data.py           # 参数重算、汇总和恢复分析
│   └── build_origin.py           # Origin 工程及图件生成
├── S转移特性曲线/                 # S 组原始转移特性
├── S输出特性曲线/                 # S 组原始输出特性
├── T转移特性曲线/                 # T 组原始转移特性
├── T输出特性曲线/                 # T 组原始输出特性
└── deliverables/
    ├── 01_origin/                # Origin 工程
    ├── 02_data/                  # 原始抽取、参数和汇总 CSV
    ├── 03_figures/               # PNG/PDF/EMF 图件
    ├── 04_manuscript/            # Markdown/TeX/DOCX/PDF 报告
    └── 05_manifest/              # 来源、图件与质控清单
```

## 复现分析

基础分析环境：Windows、Python 3.9。安装 Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

按顺序运行：

```powershell
python scripts\extract_excel.py
python scripts\analyze_data.py
python scripts\build_origin.py
```

运行边界：

- `extract_excel.py` 使用 `pywin32` 调用本机 Microsoft Excel；
- `analyze_data.py` 可在已生成 `deliverables/02_data/raw_*.csv` 后独立运行；
- `build_origin.py` 需要本机安装可用的 Origin；
- 脚本会更新 `deliverables/` 下的正式产物，建议运行前确认工作区状态。

## 数据质量与解释边界

详细记录见 [数据质量报告](deliverables/05_manifest/data_quality_report.md) 和 [来源清单](deliverables/05_manifest/source_manifest.csv)。每组仅有 3 个纵向器件，统计量用于描述本批样品，不支持总体显著性推断。恢复现象只表示连续扫描后漏电降低，不单独证明具体微观机制。