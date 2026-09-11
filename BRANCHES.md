# 分支与协作规则

仓库由三个角色共用：**模型手**（口径与文档）、**编程手**（实现）、**论文手**（只读）。用分支把三者隔开，避免"两边各改一版、谁也不知道哪份是准的"。

## 一、分支一览

| 分支 | 类型 | 归属 | 放什么 | 不放什么 |
|---|---|---|---|---|
| `main` | 长期 | 模型手维护 | `docs/`、`paper/`、`src/` 骨架、`README`、`BRANCHES` | 开发中的实现、临时脚本、内部对答案记录 |
| `modeler` | 长期 | 模型手 | `main` 全部 + `templates/`（模板标签修正版）+ `tools/` + 内部分析 | 正式实现代码 |
| `programmer` | 长期 | 编程手 | `main` 全部 + `src/qN/` 正式实现 + `outputs/` 结构 | 改 `docs/`（只读） |
| `feat/q1` … `feat/q4` | 短分支 | 编程手 | 单问实现，从 `programmer` 分出、完成即合回 | 跨问的公共改动 |

## 二、谁写哪里（防冲突的第一原则）

| 路径 | 唯一写者 | 其他人 |
|---|---|---|
| `docs/**` | 模型手 | 编程手 / 论文手 **只读** |
| `paper/**` | 模型手（论文手后续可提分支） | 只读 |
| `src/qN/**` | 编程手（`qN` 归 `feat/qN` 分支） | 模型手只读 |
| `templates/**`、`tools/**` | 模型手（仅存在于 `modeler`） | — |
| `README.md`、`BRANCHES.md` | 模型手 | 只读 |

**口径只有一个来源。** `docs/` 是发布版；脚本里的常量必须与 `docs/AGENTS.md` 决策台账一致。发现不一致 → 停下来问模型手，**不要自行改口径后继续跑**。

## 三、合并方向

```text
feat/q1 ─┐
feat/q2 ─┼─► programmer ──(验收通过)──► main
feat/q3 ─┤                                   ▲
feat/q4 ─┘                                   │
                        modeler ─────────────┘
```

- 编程手：`feat/qN` → `programmer`。**只合自己的目录**，`docs/` 冲突一律以 `main` 为准。
- 模型手：`modeler` → `main`。文档更新后同步到 `main`，再通知编程手 `git merge main`。
- `main` 由模型手合入，编程手不直接推 `main`。

## 四、编程手开一条新分支的标准流程

```bash
git checkout programmer && git pull
git checkout -b feat/q2
# ... 在 src/q2/ 下实现；不改 docs/
# 跑完 src/q2/README.md 里的验收项，输出审计产物
git add src/q2 && git commit -m "Q2: 实现两阶段随机规划 + 因果执行器"
git checkout programmer && git merge --no-ff feat/q2
git push origin programmer
```

**做完一问先回报，不要连着往下做。** Q1 是全链条标尺（单位、时间映射、约束都在它上面验证），Q1 数字对上了 Q2–Q4 才不会白跑。

## 五、提交信息约定

```text
Q<问号>: <做了什么>

<关键口径/公式变化；验收锚点实测值；与 docs 的差异(若有)>
```

例：`Q1: LP 主模型 + DP 互证` / `Q3: 附件3 改为 Bates-Granger 组合，S0–S4 重算`。

## 六、环境约定

- 用仓库内 `requirements.txt` 的依赖（numpy / openpyxl / scipy）。
- **求解器用 `scipy.optimize.linprog(method="highs")`**；不引入商业求解器，不用遗传算法 / 粒子群等启发式。
- 时间索引、位置回填、单位换算等铁律见 `docs/编程手交接.md`。
