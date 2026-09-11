# outputs/：运行产物（本地，不入库）

填好的官方结果模板与审计产物放这里，**不提交**（已在 `.gitignore` 排除，仅保留本说明）。

建议命名：

```text
outputs/result1.xlsx              Q1 结果（按位置回填）
outputs/result2.xlsx              Q2 结果
outputs/result3.xlsx              Q3 结果
outputs/result4-2.xlsx            Q4 实时电价下 Q2 策略
outputs/result4-3.xlsx            Q4 实时电价下 Q3 策略
outputs/audit_q2_daily.csv        逐日指标（购电量/紧急/费用/SOC 端点）
outputs/audit_q2_scenarios.csv    场景审计（条数、权重、残差来源日）
outputs/audit_q2_executor.csv     执行器对比（解析响应 vs DP）
outputs/run_summary.md            运行摘要（参数、耗时、验收项实测值）
```

**不要覆盖仓库外的官方模板原件。** 复制一份到本目录再回填。

模板的 144 个 10 分钟列标签整体偏移了 10 分钟，位置映射与修正版模板说明见
`docs/结果模板_标签修正说明.md`（`templates/` 下有修正标签版，仅 `modeler` 分支）。
