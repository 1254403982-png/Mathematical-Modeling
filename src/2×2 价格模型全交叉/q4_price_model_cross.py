"""Q4 价格水平模型 2×2 交叉验证。

比较 4-2/4-3 两种调控分支在 net/AR 两种价格水平模型下的全年结果。
正式 ``output/q4`` 始终只读；默认复用其中已验收的 net 两格，并在隔离目录
重建 AR 预测/配对场景后运行 42-ar、43-ar。

建议运行（在 code 目录）：
  python scr/q4_price_model_cross.py --self-test
  python scr/q4_price_model_cross.py --all --max-days 2 \
      --output-dir output/q4_price_model_cross_smoke
  python scr/q4_price_model_cross.py --all --resume
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import sys
import time
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Iterable, Sequence

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent))
import q4  # noqa: E402

EXPERIMENT_VERSION = 1
FAMILIES = ("net", "ar")
CELLS = ("42-net", "43-net", "42-ar", "43-ar")
AR_CELLS = ("42-ar", "43-ar")

CELL_CONFIGS = {
    "42-net": q4.RunConfig(
        "42-net", "4-2", "overnight", q4.ISSUE_HOURS, False, True, True,
        "4-2正式控制结构；价格水平模型强制为net（交叉验证）",
    ),
    "43-net": q4.RunConfig(
        "43-net", "4-3", "overnight", q4.ISSUE_HOURS, True, True, True,
        "4-3正式控制结构；价格水平模型强制为net（交叉验证）",
    ),
    "42-ar": q4.RunConfig(
        "42-ar", "4-2", "overnight", q4.ISSUE_HOURS, False, True, True,
        "4-2正式控制结构；价格水平模型强制为AR（交叉验证）",
    ),
    "43-ar": q4.RunConfig(
        "43-ar", "4-3", "overnight", q4.ISSUE_HOURS, True, True, True,
        "4-3正式控制结构；价格水平模型强制为AR（交叉验证）",
    ),
}

OFFICIAL_IMMUTABLE_FILES = (
    "q4_prep_summary.json",
    "q4_run_summary.json",
    "q4_42-main_summary.json",
    "q4_43-main_summary.json",
    "q4-2_daily_metrics.csv",
    "q4-3_daily_metrics.csv",
    "q4-2_main_arrays.npz",
    "q4-3_main_arrays.npz",
    "result4-2.xlsx",
    "result4-3.xlsx",
)

SELECTION_FIELDS = (
    "stage", "candidate", "branch", "event_count", "position_count",
    "mae_yuan_per_kwh", "rmse_yuan_per_kwh",
    "bias_forecast_minus_actual", "selected_by_original_rule",
    "used_in_cross_run", "note",
)

CROSS_FIELDS = (
    "cell", "branch", "price_family", "short_rho", "source_type",
    "source_path", "formal_days", "total_cost_yuan", "plan_cost_yuan",
    "effective_base_cost_yuan", "adjustment_transaction_yuan",
    "adjustment_surcharge_yuan", "emergency_cost_yuan",
    "emergency_energy_kwh", "adjust_abs_energy_kwh", "terminal_soc_kwh",
    "price_forecast_mae_yuan_per_kwh", "price_forecast_rmse_yuan_per_kwh",
    "price_forecast_bias_yuan_per_kwh", "all_planning_statuses_optimal",
    "max_billing_identity_residual_yuan", "max_cross_day_soc_residual_kwh",
    "max_dp_balance_residual_kwh", "max_dp_soc_residual_kwh",
    "max_planning_balance_residual_kwh", "max_planning_soc_residual_kwh",
    "validation_passed",
)

EFFECT_FIELDS = (
    "effect_type", "scope", "formula", "value_yuan",
    "reference_yuan", "percent_of_reference",
)


def parse_args() -> argparse.Namespace:
    root = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(description="Q4价格模型×调控分支2×2交叉验证")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--all", action="store_true", help="准备、运行并组装（默认）")
    mode.add_argument("--prep", action="store_true", help="只构造强制模型缓存")
    mode.add_argument("--run-config", choices=AR_CELLS, help="只运行一个AR单元格")
    mode.add_argument("--assemble", action="store_true", help="只组装已有四格结果")
    mode.add_argument("--self-test", action="store_true", help="运行本脚本的单元级检查")
    parser.add_argument("--root", type=Path, default=root)
    parser.add_argument("--output-dir", type=Path, default=None)
    parser.add_argument("--official-output-dir", type=Path, default=None)
    parser.add_argument("--max-days", type=int, default=None)
    parser.add_argument("--progress-every", type=int, default=20)
    parser.add_argument("--dp-grid-step", type=float, default=q4.DP_GRID_STEP_KWH)
    parser.add_argument("--resume", action="store_true", help="只复用严格匹配的实验缓存/单元格")
    parser.add_argument(
        "--rerun-net", action="store_true",
        help="正式334天也在隔离目录重跑net两格；默认复用已验收正式结果",
    )
    return parser.parse_args()


def output_dir(args: argparse.Namespace) -> Path:
    path = args.output_dir or args.root / "output" / "q4_price_model_cross"
    path = path.resolve()
    path.mkdir(parents=True, exist_ok=True)
    return path


def official_dir(args: argparse.Namespace) -> Path:
    return (args.official_output_dir or args.root / "output" / "q4").resolve()


def family_dir(args: argparse.Namespace, family: str) -> Path:
    path = output_dir(args) / family
    path.mkdir(parents=True, exist_ok=True)
    return path


def q4_args(
    args: argparse.Namespace,
    target: Path,
    *,
    resume: bool = False,
) -> SimpleNamespace:
    return SimpleNamespace(
        root=args.root.resolve(),
        output_dir=target.resolve(),
        resume=resume,
        dp_grid_step=float(args.dp_grid_step),
        progress_every=int(args.progress_every),
    )


def count_days(args: argparse.Namespace) -> int:
    if args.max_days is None:
        return q4.FORMAL_DAYS
    if not 1 <= args.max_days <= q4.FORMAL_DAYS:
        raise ValueError("--max-days必须位于1..334")
    return int(args.max_days)


def read_json(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as file:
        value = json.load(file)
    if not isinstance(value, dict):
        raise ValueError(f"JSON顶层不是对象：{path}")
    return value


def write_json(path: Path, value: dict[str, Any]) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2, allow_nan=False) + "\n",
        encoding="utf-8",
    )


def input_hashes(args: argparse.Namespace) -> dict[str, str]:
    helper_args = q4_args(args, output_dir(args))
    return {
        key: q4.q2.sha256_file(q4.input_path(helper_args, key))
        for key in ("annual", "pv_forecast", "prices")
    }


def immutable_snapshot(args: argparse.Namespace) -> dict[str, str]:
    result = {"code/scr/q4.py": q4.q2.sha256_file(Path(q4.__file__).resolve())}
    base = official_dir(args)
    for name in OFFICIAL_IMMUTABLE_FILES:
        path = base / name
        result[f"official/{name}"] = (
            q4.q2.sha256_file(path) if path.is_file() else "<missing>"
        )
    return result


def assert_immutable_unchanged(
    before: dict[str, str],
    args: argparse.Namespace,
) -> None:
    after = immutable_snapshot(args)
    changed = [key for key in before if before[key] != after.get(key)]
    if changed:
        raise RuntimeError(f"交叉验证意外修改了正式Q4文件：{changed}")


def expected_cache_identity(
    args: argparse.Namespace,
    family: str,
    count: int,
) -> dict[str, Any]:
    return {
        "experiment": "q4_price_model_cross",
        "experiment_version": EXPERIMENT_VERSION,
        "formal_days": count,
        "forced_price_model_family": family,
        "q4_source_sha256": q4.q2.sha256_file(Path(q4.__file__).resolve()),
        "input_sha256": input_hashes(args),
    }


def cache_matches(
    args: argparse.Namespace,
    family: str,
    count: int,
) -> bool:
    target = family_dir(args, family)
    metadata_path = target / q4.CACHE_FILES["metadata"]
    if not metadata_path.is_file():
        return False
    try:
        metadata = read_json(metadata_path)
    except (OSError, ValueError, json.JSONDecodeError):
        return False
    expected = expected_cache_identity(args, family, count)
    return (
        all(metadata.get(key) == value for key, value in expected.items())
        and metadata.get("selected_price_model_family") == family
        and all((target / name).is_file() for name in q4.CACHE_FILES.values())
    )


def cross_selection_rows(
    model_rows: Sequence[dict[str, Any]],
    rho_rows: Sequence[dict[str, Any]],
    family: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for source in model_rows:
        row = {key: source.get(key, "") for key in SELECTION_FIELDS[:8]}
        row.update(
            {
                "selected_by_original_rule": bool(source["selected"]),
                "used_in_cross_run": source["candidate"] == family,
                "note": f"模型族强制为{family}；原规则选择结果单独保留",
            }
        )
        rows.append(row)
    for source in rho_rows:
        row = {key: source.get(key, "") for key in SELECTION_FIELDS[:8]}
        row.update(
            {
                "selected_by_original_rule": bool(source["selected"]),
                "used_in_cross_run": bool(source["selected"]),
                "note": f"rho在强制family={family}条件下按原规则选择",
            }
        )
        rows.append(row)
    return rows


def prepare_family(
    args: argparse.Namespace,
    family: str,
    count: int,
    *,
    allow_resume: bool,
) -> dict[str, Any]:
    if family not in FAMILIES:
        raise ValueError(f"未知价格模型族：{family}")
    target = family_dir(args, family)
    if allow_resume and cache_matches(args, family, count):
        metadata = read_json(target / q4.CACHE_FILES["metadata"])
        print(f"[prep-{family}] 复用严格匹配缓存", flush=True)
        return metadata

    started = time.perf_counter()
    helper_args = q4_args(args, target)
    annual = q4.q2.read_annual_data(q4.input_path(helper_args, "annual"))
    prices = q4.read_price_data(q4.input_path(helper_args, "prices"))
    q4.validate_inputs(annual, prices)
    attachment3 = q4.q3.read_attachment3(q4.input_path(helper_args, "pv_forecast"))
    raw, _, _ = q4.build_raw_resource_events(annual, attachment3)
    bg = q4.calibrate_bg_weight(raw["4-3"], annual)
    resources = q4.finalize_resource_events(raw, annual, float(bg["w"]))

    natural_family, model_rows = q4.select_price_model(resources, prices)
    short_rho, rho_rows = q4.select_short_rho(family, resources, prices)
    selection = cross_selection_rows(model_rows, rho_rows, family)
    q4.write_csv(
        target / "q4_price_model_cross_selection.csv",
        selection,
        SELECTION_FIELDS,
    )

    forecasts = q4.build_price_forecasts(family, short_rho, resources, prices)
    q4.write_price_forecast_audit(
        target / "q4_price_forecast_audit.csv", forecasts, resources, prices
    )
    metric_rows = q4.price_forecast_metric_rows(forecasts, resources, prices)
    q4.write_csv(
        target / "q4_price_forecast_metrics.csv",
        metric_rows,
        tuple(metric_rows[0]),
    )
    scenario_summary = q4.build_scenario_cache(
        target, count, resources, forecasts, annual, prices
    )
    shape_max_error = max(
        abs(float(np.mean(forecast.shape_window)) - 1.0)
        for branch in q4.BRANCHES
        for forecast in forecasts[branch].values()
        if forecast.issue_hour == 0
    )
    metadata = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        **expected_cache_identity(args, family, count),
        "selected_price_model_family": family,
        "natural_selected_price_model_family": natural_family,
        "selected_short_rho": short_rho,
        "selection_rule": (
            "family forced for 2x2 validation; rho selected conditionally by the "
            "original two-branch January MAE/RMSE/lower-rho rule"
        ),
        "price_history_days": q4.PRICE_HISTORY_DAYS,
        "minimum_regression_samples": q4.MIN_REGRESSION_SAMPLES,
        "price_epsilon": q4.PRICE_EPSILON,
        "net_load_scale_kwh": q4.NET_LOAD_SCALE_KWH,
        "bates_granger": bg,
        "scenario": scenario_summary,
        "max_zero_hour_shape_mean_error": shape_max_error,
        "runtime_seconds": time.perf_counter() - started,
    }
    write_json(target / q4.CACHE_FILES["metadata"], metadata)
    print(
        f"[prep-{family}] 完成：强制模型={family}, 自然选择={natural_family}, "
        f"rho_p={short_rho}, 事件={scenario_summary['event_count']}, "
        f"耗时={metadata['runtime_seconds']:.1f}s",
        flush=True,
    )
    return metadata


def cell_provenance_path(target: Path, cell: str) -> Path:
    return target / f"q4_{cell}_provenance.json"


def expected_cell_provenance(
    args: argparse.Namespace,
    cell: str,
    count: int,
) -> dict[str, Any]:
    config = CELL_CONFIGS[cell]
    target = family_dir(args, cell.split("-")[1])
    metadata_path = target / q4.CACHE_FILES["metadata"]
    if not metadata_path.is_file():
        raise FileNotFoundError(f"缺少实验缓存元数据：{metadata_path}")
    return {
        "experiment": "q4_price_model_cross",
        "experiment_version": EXPERIMENT_VERSION,
        "cell": cell,
        "formal_days": count,
        "config": json.loads(json.dumps(asdict(config), ensure_ascii=False)),
        "dp_grid_step_kwh": float(args.dp_grid_step),
        "cache_metadata_sha256": q4.q2.sha256_file(metadata_path),
        "q4_source_sha256": q4.q2.sha256_file(Path(q4.__file__).resolve()),
        "runner_source_sha256": q4.q2.sha256_file(Path(__file__).resolve()),
        "input_sha256": input_hashes(args),
    }


def reusable_cell(
    args: argparse.Namespace,
    cell: str,
    count: int,
) -> bool:
    config = CELL_CONFIGS[cell]
    target = family_dir(args, cell.split("-")[1])
    if not q4.configuration_complete(target, config, count):
        return False
    provenance_path = cell_provenance_path(target, cell)
    if not provenance_path.is_file():
        return False
    provenance = read_json(provenance_path)
    provenance.pop("generated_at", None)
    return provenance == expected_cell_provenance(args, cell, count)


def run_cell(
    args: argparse.Namespace,
    cell: str,
    count: int,
    *,
    allow_resume: bool,
) -> dict[str, Any]:
    family = cell.split("-")[1]
    target = family_dir(args, family)
    if not cache_matches(args, family, count):
        prepare_family(args, family, count, allow_resume=True)
    if allow_resume and reusable_cell(args, cell, count):
        summary_path = q4.config_paths(target, CELL_CONFIGS[cell])["summary"]
        print(f"[{cell}] 复用严格匹配的全年结果", flush=True)
        return read_json(summary_path)

    helper_args = q4_args(args, target, resume=False)
    summary = q4.run_configuration(helper_args, CELL_CONFIGS[cell], count)
    provenance = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        **expected_cell_provenance(args, cell, count),
    }
    write_json(cell_provenance_path(target, cell), provenance)
    return summary


def validate_official_net(
    args: argparse.Namespace,
) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    target = official_dir(args)
    prep_path = target / q4.CACHE_FILES["metadata"]
    run_path = target / "q4_run_summary.json"
    if not prep_path.is_file() or not run_path.is_file():
        raise FileNotFoundError("缺少正式Q4 prep/run汇总，不能复用net基线")
    prep = read_json(prep_path)
    run = read_json(run_path)
    if prep.get("selected_price_model_family") != "net":
        raise RuntimeError("正式Q4缓存不是net模型，不能作为2×2的net基线")
    if prep.get("formal_days") != q4.FORMAL_DAYS:
        raise RuntimeError("正式Q4 net缓存不是334天")
    if prep.get("input_sha256") != input_hashes(args):
        raise RuntimeError("正式Q4 net缓存的输入哈希与当前附件不一致")
    if not (
        run.get("run_type") == "full"
        and run.get("formal_days") == q4.FORMAL_DAYS
        and run.get("status") == "success"
        and run.get("validation", {}).get("all_checks_passed") is True
    ):
        raise RuntimeError("正式Q4 net结果未通过完整验收")

    result: dict[str, dict[str, Any]] = {}
    for cell, main_name in (("42-net", "42-main"), ("43-net", "43-main")):
        config = q4.RUN_CONFIGS[main_name]
        if not q4.configuration_complete(target, config, q4.FORMAL_DAYS):
            raise RuntimeError(f"正式net单元格不完整：{cell}")
        summary = read_json(q4.config_paths(target, config)["summary"])
        branch = config.branch
        reported = float(run["main_results"][branch]["total_cost_yuan"])
        actual = float(summary["annual"]["total_cost_yuan"])
        if not math.isclose(reported, actual, rel_tol=0.0, abs_tol=1.0e-6):
            raise RuntimeError(f"正式net汇总之间费用不一致：{cell}")
        result[cell] = summary
    return result, prep


def load_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def forecast_metrics(path: Path, branch: str) -> dict[str, float]:
    selected = [row for row in load_csv_rows(path) if row["branch"] == branch]
    if not selected:
        raise RuntimeError(f"价格预测指标缺少分支{branch}：{path}")
    counts = np.asarray([float(row["position_count"]) for row in selected])
    total = float(np.sum(counts))
    mae = np.asarray([float(row["mae_yuan_per_kwh"]) for row in selected])
    rmse = np.asarray([float(row["rmse_yuan_per_kwh"]) for row in selected])
    bias = np.asarray([float(row["bias_forecast_minus_actual"]) for row in selected])
    return {
        "mae": float(np.dot(counts, mae) / total),
        "rmse": float(np.sqrt(np.dot(counts, np.square(rmse)) / total)),
        "bias": float(np.dot(counts, bias) / total),
    }


def summary_validation_passes(summary: dict[str, Any]) -> bool:
    validation = summary.get("validation", {})
    tolerance_checks = (
        ("max_billing_identity_residual_yuan", q4.CHECK_TOLERANCE),
        ("max_cross_day_soc_residual_kwh", q4.CHECK_TOLERANCE),
        ("max_dp_balance_residual_kwh", q4.CHECK_TOLERANCE),
        ("max_dp_soc_residual_kwh", q4.CHECK_TOLERANCE),
        ("max_dp_simultaneous_product_kwh2", q4.SIMULTANEOUS_PRODUCT_TOLERANCE),
        ("max_dp_value_convexity_violation", 1.0e-7),
        ("max_planning_balance_residual_kwh", q4.CHECK_TOLERANCE),
        ("max_planning_soc_residual_kwh", q4.CHECK_TOLERANCE),
        ("max_planning_simultaneous_product_kwh2", q4.SIMULTANEOUS_PRODUCT_TOLERANCE),
        ("max_planning_delta_product_kwh2", q4.SIMULTANEOUS_PRODUCT_TOLERANCE),
    )
    if validation.get("all_planning_statuses_optimal") is not True:
        return False
    for key, tolerance in tolerance_checks:
        value = float(validation.get(key, math.inf))
        if not math.isfinite(value) or value > tolerance:
            return False
    annual = summary.get("annual", {})
    return all(
        math.isfinite(float(annual.get(key, math.nan)))
        for key in (
            "total_cost_yuan", "plan_cost_yuan", "effective_base_cost_yuan",
            "adjustment_transaction_yuan", "adjustment_surcharge_yuan",
            "emergency_cost_yuan", "emergency_energy_kwh",
            "adjust_abs_energy_kwh", "terminal_soc_kwh",
        )
    )


def isolated_summary(
    args: argparse.Namespace,
    cell: str,
    count: int,
) -> tuple[dict[str, Any], dict[str, Any], Path]:
    family = cell.split("-")[1]
    target = family_dir(args, family)
    config = CELL_CONFIGS[cell]
    if not reusable_cell(args, cell, count):
        raise RuntimeError(f"实验单元格缺失或来源不匹配：{cell}")
    return (
        read_json(q4.config_paths(target, config)["summary"]),
        read_json(target / q4.CACHE_FILES["metadata"]),
        target,
    )


def cell_row(
    cell: str,
    summary: dict[str, Any],
    prep: dict[str, Any],
    source_type: str,
    source_path: Path,
) -> dict[str, Any]:
    branch = CELL_CONFIGS[cell].branch
    metric = forecast_metrics(source_path / "q4_price_forecast_metrics.csv", branch)
    annual = summary["annual"]
    validation = summary["validation"]
    passed = summary_validation_passes(summary)
    return {
        "cell": cell,
        "branch": branch,
        "price_family": cell.split("-")[1],
        "short_rho": prep["selected_short_rho"],
        "source_type": source_type,
        "source_path": str(source_path),
        "formal_days": summary["formal_days"],
        "total_cost_yuan": annual["total_cost_yuan"],
        "plan_cost_yuan": annual["plan_cost_yuan"],
        "effective_base_cost_yuan": annual["effective_base_cost_yuan"],
        "adjustment_transaction_yuan": annual["adjustment_transaction_yuan"],
        "adjustment_surcharge_yuan": annual["adjustment_surcharge_yuan"],
        "emergency_cost_yuan": annual["emergency_cost_yuan"],
        "emergency_energy_kwh": annual["emergency_energy_kwh"],
        "adjust_abs_energy_kwh": annual["adjust_abs_energy_kwh"],
        "terminal_soc_kwh": annual["terminal_soc_kwh"],
        "price_forecast_mae_yuan_per_kwh": metric["mae"],
        "price_forecast_rmse_yuan_per_kwh": metric["rmse"],
        "price_forecast_bias_yuan_per_kwh": metric["bias"],
        "all_planning_statuses_optimal": validation["all_planning_statuses_optimal"],
        "max_billing_identity_residual_yuan": validation["max_billing_identity_residual_yuan"],
        "max_cross_day_soc_residual_kwh": validation["max_cross_day_soc_residual_kwh"],
        "max_dp_balance_residual_kwh": validation["max_dp_balance_residual_kwh"],
        "max_dp_soc_residual_kwh": validation["max_dp_soc_residual_kwh"],
        "max_planning_balance_residual_kwh": validation["max_planning_balance_residual_kwh"],
        "max_planning_soc_residual_kwh": validation["max_planning_soc_residual_kwh"],
        "validation_passed": passed,
    }


def effect_rows(rows: Sequence[dict[str, Any]]) -> list[dict[str, Any]]:
    costs = {row["cell"]: float(row["total_cost_yuan"]) for row in rows}
    if set(costs) != set(CELLS):
        raise RuntimeError(f"2×2结果单元格不完整：{sorted(costs)}")

    def make(
        effect_type: str,
        scope: str,
        formula: str,
        value: float,
        reference: float,
    ) -> dict[str, Any]:
        return {
            "effect_type": effect_type,
            "scope": scope,
            "formula": formula,
            "value_yuan": value,
            "reference_yuan": reference,
            "percent_of_reference": 100.0 * value / reference if reference else 0.0,
        }

    family_42 = costs["42-ar"] - costs["42-net"]
    family_43 = costs["43-ar"] - costs["43-net"]
    saving_net = costs["42-net"] - costs["43-net"]
    saving_ar = costs["42-ar"] - costs["43-ar"]
    return [
        make("family_effect_ar_minus_net", "4-2", "42-ar - 42-net", family_42, costs["42-net"]),
        make("family_effect_ar_minus_net", "4-3", "43-ar - 43-net", family_43, costs["43-net"]),
        make("strategy_saving_42_minus_43", "net", "42-net - 43-net", saving_net, costs["42-net"]),
        make("strategy_saving_42_minus_43", "ar", "42-ar - 43-ar", saving_ar, costs["42-ar"]),
        make(
            "interaction_strategy_saving_ar_minus_net",
            "2x2",
            "(42-ar - 43-ar) - (42-net - 43-net)",
            saving_ar - saving_net,
            saving_net,
        ),
    ]


def assemble_cross(
    args: argparse.Namespace,
    count: int,
    immutable_before: dict[str, str],
) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    family_metadata: dict[str, Any] = {}
    use_official_net = count == q4.FORMAL_DAYS and not args.rerun_net

    if use_official_net:
        net_summaries, net_prep = validate_official_net(args)
        family_metadata["net"] = net_prep
        for cell in ("42-net", "43-net"):
            rows.append(
                cell_row(
                    cell, net_summaries[cell], net_prep,
                    "reused_verified_official", official_dir(args),
                )
            )
    else:
        for cell in ("42-net", "43-net"):
            summary, prep, source = isolated_summary(args, cell, count)
            family_metadata["net"] = prep
            rows.append(cell_row(cell, summary, prep, "isolated_fresh_run", source))

    for cell in ("42-ar", "43-ar"):
        summary, prep, source = isolated_summary(args, cell, count)
        family_metadata["ar"] = prep
        rows.append(cell_row(cell, summary, prep, "isolated_fresh_run", source))

    order = {cell: index for index, cell in enumerate(CELLS)}
    rows.sort(key=lambda row: order[str(row["cell"])])
    if not all(bool(row["validation_passed"]) for row in rows):
        failed = [row["cell"] for row in rows if not row["validation_passed"]]
        raise RuntimeError(f"交叉实验验收失败：{failed}")
    effects = effect_rows(rows)
    out = output_dir(args)
    q4.write_csv(out / "q4_price_model_cross.csv", rows, CROSS_FIELDS)
    q4.write_csv(out / "q4_price_model_effects.csv", effects, EFFECT_FIELDS)
    assert_immutable_unchanged(immutable_before, args)

    report = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "experiment": "q4_price_model_cross",
        "experiment_version": EXPERIMENT_VERSION,
        "run_type": "full" if count == q4.FORMAL_DAYS else "partial_smoke_test",
        "formal_days": count,
        "status": "success",
        "errors": [],
        "design": {
            "factors": {
                "control_branch": list(q4.BRANCHES),
                "price_level_family": list(FAMILIES),
            },
            "cells": list(CELLS),
            "all_non_price_settings_match_formal_q4": True,
            "net_source": "verified official Q4" if use_official_net else "isolated rerun",
            "ar_source": "isolated forced-family rerun",
            "formal_q4_outputs_modified": False,
        },
        "provenance": {
            "q4_source_sha256": immutable_before["code/scr/q4.py"],
            "official_immutable_sha256": {
                key: value for key, value in immutable_before.items() if key.startswith("official/")
            },
            "input_sha256": input_hashes(args),
            "family_metadata": family_metadata,
        },
        "cells": rows,
        "effects": effects,
        "validation": {
            "cell_count": len(rows),
            "exact_cell_set": set(row["cell"] for row in rows) == set(CELLS),
            "all_cells_passed": all(bool(row["validation_passed"]) for row in rows),
            "all_planning_statuses_optimal": all(
                bool(row["all_planning_statuses_optimal"]) for row in rows
            ),
            "official_files_unchanged": True,
        },
        "outputs": {},
    }
    report["outputs"] = {
        name: {
            "path": str(out / name),
            "sha256": q4.q2.sha256_file(out / name),
        }
        for name in ("q4_price_model_cross.csv", "q4_price_model_effects.csv")
    }
    write_json(out / "q4_price_model_cross_summary.json", report)
    print("[assemble] 2×2价格模型交叉验证组装并验收通过", flush=True)
    for row in rows:
        print(
            f"  {row['cell']}: cost={float(row['total_cost_yuan']):.2f}, "
            f"rho={float(row['short_rho']):g}, SOC={float(row['terminal_soc_kwh']):.3f}",
            flush=True,
        )
    return report


def self_test() -> None:
    synthetic = [
        {"cell": "42-net", "total_cost_yuan": 100.0},
        {"cell": "43-net", "total_cost_yuan": 80.0},
        {"cell": "42-ar", "total_cost_yuan": 110.0},
        {"cell": "43-ar", "total_cost_yuan": 85.0},
    ]
    effects = effect_rows(synthetic)
    values = {row["formula"]: row["value_yuan"] for row in effects}
    assert values["42-ar - 42-net"] == 10.0
    assert values["43-ar - 43-net"] == 5.0
    assert values["42-net - 43-net"] == 20.0
    assert values["42-ar - 43-ar"] == 25.0
    assert values["(42-ar - 43-ar) - (42-net - 43-net)"] == 5.0
    assert set(CELL_CONFIGS) == set(CELLS)
    assert all(config.terminal_mode == "overnight" for config in CELL_CONFIGS.values())
    assert CELL_CONFIGS["42-ar"].adjust_purchases is False
    assert CELL_CONFIGS["43-ar"].adjust_purchases is True
    print("Q4价格模型2×2交叉验证脚本self-test通过", flush=True)


def main() -> None:
    args = parse_args()
    args.root = args.root.resolve()
    count = count_days(args)
    before = immutable_snapshot(args)
    try:
        if args.self_test:
            self_test()
            return
        if args.assemble:
            assemble_cross(args, count, before)
            return
        if args.prep:
            families = FAMILIES if count < q4.FORMAL_DAYS or args.rerun_net else ("ar",)
            for family in families:
                prepare_family(args, family, count, allow_resume=args.resume)
            return
        if args.run_config is not None:
            family = args.run_config.split("-")[1]
            prepare_family(args, family, count, allow_resume=True)
            run_cell(args, args.run_config, count, allow_resume=args.resume)
            return

        # 默认行为与--all相同。正式运行复用已验收net；smoke必须实际跑满四格。
        if count == q4.FORMAL_DAYS and not args.rerun_net:
            validate_official_net(args)
            families: Sequence[str] = ("ar",)
            cells: Sequence[str] = AR_CELLS
        else:
            families = FAMILIES
            cells = CELLS
        for family in families:
            prepare_family(args, family, count, allow_resume=args.resume)
        for cell in cells:
            run_cell(args, cell, count, allow_resume=args.resume)
        assemble_cross(args, count, before)
    finally:
        assert_immutable_unchanged(before, args)


if __name__ == "__main__":
    main()
