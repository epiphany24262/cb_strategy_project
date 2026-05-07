# 可转债多因子选券研究

A 股可转债横截面多因子选券策略研究，包含完整的研究底稿、Markdown 报告和 Word 报告。

## 策略概要

| 项目 | 口径 |
|---|---|
| 数据区间 | 2018-01-02 至 2024-12-26 |
| 因子组合 | `bond_prem + dblow + alpha_pct_chg_5` |
| 可交易性约束 | `amount` 日截面底部 20% 剔除 |
| 调仓频率 | BW 双周 |
| 选券缓冲 | Top10 买入，跌出 Top15 卖出 |
| 权重方案 | 持仓内按综合评分排名加权 |
| 手续费率 | 0.002 |
| 基准 | 研究样本等权（不计手续费） |

**核心结果**（缓冲排名加权 BW 主策略，`fee_rate=0.002`）：

| 指标 | 数值 |
|---|---|
| 年化收益 | 36.86% |
| 夏普比率 | 1.76 |
| 最大回撤 | -18.25% |
| 超额累计收益 | 348.29% |
| 信息比率 | 1.44 |

## 文件结构

```
.
├── README.md
├── requirements.txt             # Python 依赖清单
├── data/
│   └── cb_data.pq               # 原始数据（需自行获取，见下方说明）
├── notebooks/
│   ├── convertible_bond_factor_report.ipynb          # 主研究底稿
│   └── convertible_bond_factor_report_fee_robust.ipynb  # 补充实验底稿
├── report/
│   ├── report.md                # Markdown 版研究报告
│   ├── report.docx              # Word 版研究报告（黑白灰排版）
│   ├── build_report_docx.py     # Word 报告生成脚本
│   └── assets/                  # 报告插图（共 11 张）
└── ref/
    ├── student_demo_convertible_bond.py   # 课程脚手架代码（仅供学习参考）
    ├── 可转债多因子策略探究.pdf           # 参考材料
    └── 可转债量化交易.pptx                # 参考材料
```

## 环境配置

```bash
pip install -r requirements.txt
```

依赖清单：

| 包 | 用途 |
|---|---|
| `numpy`, `pandas` | 数据处理与矩阵运算 |
| `matplotlib` | 图表绘制 |
| `pyarrow` | Parquet 文件读写 |
| `nbformat` | Notebook 文件解析（Word 报告导出用） |
| `python-docx` | Word 文档生成 |
| `jupyter` | Notebook 运行环境 |

### 中文字体

研究底稿已配置 matplotlib 中文字体，默认使用 `Microsoft YaHei / SimHei / SimSun`。macOS / Linux 用户需要将 notebook 中的字体列表替换为系统可用的中文字体（如 `WenQuanYi Micro Hei` 或 `Noto Sans CJK`）。

## 数据说明

原始数据文件为 `data/cb_data.pq`（Parquet 格式），包含 2018-01-02 至 2024-12-26 的 A 股可转债日频数据，共约 56 万行、80 列，`code + trade_date` 为唯一键。

主要字段：
- 行情数据：`close`、`amount`、`pct_chg`、`turnover_5` 等
- 转债指标：`conv_prem`、`bond_prem`、`dblow`、`theory_bias` 等
- 正股联动：`alpha_pct_chg_5`
- 条款相关：剩余期限、转股起始日等

> 由于数据文件超过 GitHub 100MB 限制，如仅需浏览报告和代码逻辑，项目中已包含完整的研究底稿和 Word 报告。

## 复现步骤

1. 确保 `data/cb_data.pq` 放在项目根目录的 `data/` 文件夹中
2. 激活环境：`conda activate QuantEnv`
3. 在项目根目录打开并运行主研究底稿：
   ```
   notebooks/convertible_bond_factor_report.ipynb
   ```
4. 如需重新生成 Word 报告：
   ```powershell
   python report/build_report_docx.py
   ```

所有报告图片均来自研究底稿实际输出，不在报告脚本中重复绘制。

## 报告结构

报告按以下逻辑展开：

1. **数据与样本过滤** — 样本池描述与可交易性约束
2. **单因子检验** — 分组回测、IC/RankIC/IR（dblow、bond_prem、alpha_pct_chg_5）
3. **多因子合成** — 因子标准化、相关性分析与组合比较
4. **主策略构建** — 缓冲排名加权 BW 回测与绩效指标
5. **稳健性分析** — 调仓频率、TopN、手续费敏感性、年度/分阶段表现、执行层对比、滚动验证
6. **风险与局限性**
7. **结论**

## 定位

本项目是一份量化策略课程作业，展示了从单因子检验到多因子合成再到稳健性验证的完整研究流程。策略结果基于历史回测，不等同于实盘表现。
