# Mathematical Modeling: Microgrid Scheduling

2026 高教社杯数学建模 C 题：微网与外部电网电力调控策略。

## 当前内容

- `src/q1_reference_check.py`：Q1 的独立参照核对脚本。它读取附件1，计算无储能基准，并以 MILP 验证购电计划、SOC 约束和指定时段结果。
- `requirements.txt`：当前脚本所需 Python 依赖。

该脚本用于数据对齐和结果交叉核验，不替代项目当前的 Q1 主求解接口。Q2、Q3 的正式代码将在完成验证后加入仓库。

## 本地数据

为避免重复分发题目附件，原始数据和官方模板不提交到仓库。请将题目附件放到：

```text
data/附件1.xlsx
```

## 运行

```bash
python -m pip install -r requirements.txt
python src/q1_reference_check.py
```

也可指定附件路径：

```bash
python src/q1_reference_check.py --input "D:/数模/附件1.xlsx"
```

## 协作约定

- 代码、可复现运行说明和小型配置文件提交到仓库。
- 原始附件、官方结果模板和生成结果保留在本地 `data/`、`outputs/`，除非团队另行决定上传。
- 每个问题的正式实现应提供运行入口、参数说明、输入文件约定和验收输出。
