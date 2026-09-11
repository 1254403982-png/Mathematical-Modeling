# -*- coding: utf-8 -*-
"""生成论文图 1：全年电价 / 小区负荷 / 光伏发电的时空分布

数据源（官方附件，本机路径）：
  附件2.xlsx  sheet「小区负载」「光伏发电实际功率」  365 天 × 144 时段
  附件4.xlsx  sheet「Sheet1」                      365 天 × 144 时段（电价）
时间轴口径（AGENTS.md）：附件第 i 个数据（零基）↔ 模型时段 t=i+1 ↔ 物理区间 [i*10,(i+1)*10) 分钟

输出：
  paper/figures/图1_全年时空分布.png
  （图 1(d) 计划购电 需 Q2 结果复现后补，本脚本暂不生成）
"""
import os
import numpy as np
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

DESK = r"C:\Users\youra\Desktop"
OUT_DIR = r"C:\Users\youra\WorkBuddy\2026-09-12-01-15-09\Opt-Solver-211\paper\figures"
os.makedirs(OUT_DIR, exist_ok=True)

START = "2025-02-01"   # 考核期起点：1 月为预热期（AGENTS.md）


def load_matrix(fname, sheet):
    wb = openpyxl.load_workbook(os.path.join(DESK, fname), read_only=True, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    dates, mat = [], []
    for r in rows[1:]:
        if r[0] is None:
            continue
        d = r[0]
        key = "%04d-%02d-%02d" % (d.year, d.month, d.day)
        dates.append(key)
        mat.append([float(v) if v is not None else np.nan for v in r[1:145]])
    return dates, np.array(mat, dtype=float)


d_load, M_load = load_matrix("附件2.xlsx", "小区负载")
d_pv, M_pv = load_matrix("附件2.xlsx", "光伏发电实际功率")
d_price, M_price = load_matrix("附件4.xlsx", "Sheet1")

assert d_load == d_pv == d_price, "三份附件日期不一致"
sel = [i for i, d in enumerate(d_load) if d >= START]
dates = [d_load[i] for i in sel]
PL, PV, PR = M_load[sel], M_pv[sel], M_price[sel]
nD = len(sel)
print("考核期天数 =", nD, dates[0], "→", dates[-1])

# ---------------- 数据特征（供论文 2.1 引用） ----------------
print("\n[附件4 电价] min=%.4f max=%.4f mean=%.4f 元/kWh" % (PR.min(), PR.max(), PR.mean()))
print("[附件2 小区负荷] min=%.2f max=%.2f mean=%.2f kW" % (PL.min(), PL.max(), PL.mean()))
print("[附件2 光伏实际] min=%.2f max=%.2f mean=%.2f kW" % (PV.min(), PV.max(), PV.mean()))

prof_p = PR.mean(axis=0)                     # 日内平均电价曲线（144）
prof_l = PL.mean(axis=0)
prof_v = PV.mean(axis=0)
order = np.argsort(prof_p)
print("\n电价最低 6 个时段(t):", [(int(t) + 1, round(prof_p[t], 3)) for t in order[:6]])
print("电价最高 6 个时段(t):", [(int(t) + 1, round(prof_p[t], 3)) for t in order[-6:][::-1]])
print("电价 08:00 附近 t=49:", round(prof_p[48], 3), " 12:00 附近 t=73:", round(prof_p[72], 3),
      " 19:00 附近 t=115:", round(prof_p[114], 3))
print("负荷 峰值时段 t=%d (%.1f kW)，谷值时段 t=%d (%.1f kW)"
      % (int(np.argmax(prof_l)) + 1, prof_l.max(), int(np.argmin(prof_l)) + 1, prof_l.min()))

import datetime
wd = np.array([datetime.date(int(d[:4]), int(d[5:7]), int(d[8:10])).weekday() for d in dates])
names = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
print("\n负荷的周内效应（日均总电量 kWh）:")
for k in range(7):
    print("   %s  %.0f" % (names[k], PL[wd == k].sum() / max((wd == k).sum(), 1)))

mon = np.array([int(d[5:7]) for d in dates])
print("\n光伏的月度均值（kW）:")
print("   " + "  ".join("%d月:%.0f" % (m, PV[mon == m].mean()) for m in range(2, 13)))
print("电价的月度均值（元/kWh）:")
print("   " + "  ".join("%d月:%.3f" % (m, PR[mon == m].mean()) for m in range(2, 13)))

# ---------------- 绘图 ----------------
Y = np.arange(145) / 6.0            # 纵轴：一天 0~24 小时
panels = [(PR, "(a) 外网电价", "YlOrRd", "元/kWh"),
          (PL, "(b) 小区负荷", "viridis", "kW"),
          (PV, "(c) 光伏发电实际功率", "YlGnBu", "kW")]

fig, axes = plt.subplots(3, 1, figsize=(11.5, 9.6), sharex=True,
                         gridspec_kw=dict(hspace=0.28))
for ax, (M, title, cmap, unit) in zip(axes, panels):
    im = ax.imshow(M.T, aspect="auto", origin="lower", cmap=cmap,
                   extent=[0, nD, 0, 24], interpolation="nearest")
    ax.set_ylabel("时刻 / h", fontsize=11)
    ax.set_yticks([0, 4, 8, 12, 16, 20, 24])
    ax.set_title(title, fontsize=12.5, loc="left", pad=6)
    cb = fig.colorbar(im, ax=ax, pad=0.012, fraction=0.03)
    cb.set_label(unit, fontsize=9.5)
    cb.ax.tick_params(labelsize=9)

ticks, labels = [], []
for i, d in enumerate(dates):
    if i == 0 or d[5:7] != dates[i - 1][5:7]:
        ticks.append(i)
        labels.append("%d月" % int(d[5:7]))
axes[-1].set_xticks(ticks)
axes[-1].set_xticklabels(labels, fontsize=10)
axes[-1].set_xlabel("报告日期序（2025-02-01 → 2025-12-31，共 %d 天）" % nD, fontsize=11)

fig.suptitle("图 1  全年电价、小区负荷与光伏发电的时空分布", fontsize=14, y=0.985)
out = os.path.join(OUT_DIR, "图1_全年时空分布.png")
fig.savefig(out, dpi=200, bbox_inches="tight")
print("\n已保存:", out)
