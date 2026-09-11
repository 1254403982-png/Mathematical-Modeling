# data/：题目附件（本地，不入库）

附件体积大且属竞赛原始材料，**不提交到仓库**。请在本机放到本目录：

```text
data/附件1.xlsx    电价 + 典型日负荷/光伏（144 点）
data/附件2.xlsx    全年 365 天 × 144 槽实际负荷与光伏
data/附件3.xlsx    光伏整点预报（每日 0/6/12/18 时发布）
data/附件4.xlsx    实时电价
```

官方结果模板 `result1.xlsx`、`result2.xlsx`、`result3.xlsx`、`result4-2.xlsx`、`result4-3.xlsx`
同样不入库，请自行放到本地。

运行脚本时可用 `--input` 指定任意路径，不强制放在本目录：

```bash
python src/q1/reference_check.py --input "D:/数模/附件1.xlsx"
```
