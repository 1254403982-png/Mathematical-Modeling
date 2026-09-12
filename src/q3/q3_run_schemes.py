"""并行启动五个Q3方案进程（先检查prep产物），全部完成后自动执行assemble。

用法：python scr/q3_run_schemes.py
也可以手动开终端分别运行：python scr/q3.py --scheme all|-6|-12|-18|0
"""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

SCHEMES = ("0", "all", "-6", "-12", "-18")


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    q3_script = Path(__file__).resolve().parent / "q3.py"
    out = root / "output" / "q3"
    out.mkdir(parents=True, exist_ok=True)
    required = (out / "q3_scenarios.npz", out / "q3_event_index.csv")
    missing = [path.name for path in required if not path.is_file()]
    if missing:
        raise SystemExit(
            f"缺少prep产物：{missing}\n请先运行：python scr/q3.py --prep"
        )

    processes: list[tuple[str, subprocess.Popen]] = []
    for scheme in SCHEMES:
        log = (out / f"q3_{scheme}_console.log").open("w", encoding="utf-8")
        process = subprocess.Popen(
            [sys.executable, str(q3_script), "--scheme", scheme],
            cwd=str(root),
            stdout=log,
            stderr=subprocess.STDOUT,
        )
        processes.append((scheme, process))
        print(f"[runner] 已启动方案进程 {scheme} (pid={process.pid})", flush=True)

    failed: list[str] = []
    while processes:
        finished = []
        for i, (scheme, process) in enumerate(processes):
            code = process.poll()
            if code is not None:
                finished.append(i)
                if code == 0:
                    print(f"[runner] {scheme} 完成 (exit 0)", flush=True)
                else:
                    print(
                        f"[runner] {scheme} 失败 (exit {code})，"
                        f"详见 output/q3/q3_{scheme}_console.log",
                        flush=True,
                    )
                    failed.append(scheme)
        for i in reversed(finished):
            processes.pop(i)
        if processes:
            time.sleep(5)

    if failed:
        raise SystemExit(f"方案失败：{failed}，不执行assemble")
    print("[runner] 全部方案完成，开始assemble", flush=True)
    result = subprocess.run(
        [sys.executable, str(q3_script), "--assemble"], cwd=str(root)
    )
    raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
