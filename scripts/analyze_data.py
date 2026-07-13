from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.signal import savgol_filter
from scipy.stats import linregress

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "deliverables" / "02_data"
MANIFEST_DIR = ROOT / "deliverables" / "05_manifest"
CONFIG = json.loads((ROOT / "analysis_config.json").read_text(encoding="utf-8"))

KEYS = ["device_group", "device_no", "device_id", "dose_krad_si", "scan_no"]


def interpolate(x: np.ndarray, y: np.ndarray, point: float) -> float:
    order = np.argsort(x)
    return float(np.interp(point, x[order], y[order]))


def smooth_current(values: np.ndarray) -> np.ndarray:
    cfg = CONFIG["transfer"]
    size = len(values)
    window = min(int(cfg["gm_savgol_window"]), size if size % 2 else size - 1)
    if window < 5:
        return values.copy()
    return savgol_filter(values, window, int(cfg["gm_savgol_order"]), mode="interp")


def extract_ss(vgs: np.ndarray, current: np.ndarray) -> tuple[float, float, float, float]:
    cfg = CONFIG["transfer"]
    absolute = np.abs(current)
    positive = absolute[absolute > 0]
    if len(positive) < int(cfg["ss_window_points"]):
        return np.nan, np.nan, np.nan, np.nan
    noise = max(float(np.nanpercentile(positive, 10)), 1e-14)
    ceiling = float(np.nanmax(positive)) * 0.1
    mask = (absolute >= noise * 3) & (absolute <= ceiling)
    x = vgs[mask]
    y = np.log10(absolute[mask])
    window = int(cfg["ss_window_points"])
    if len(x) < window:
        return np.nan, np.nan, np.nan, np.nan
    best = None
    for start in range(len(x) - window + 1):
        result = linregress(x[start : start + window], y[start : start + window])
        if result.slope <= 0:
            continue
        score = result.rvalue**2
        if best is None or score > best[0]:
            best = (score, result.slope, x[start], x[start + window - 1])
    if best is None or best[0] < float(cfg["ss_min_r2"]):
        return np.nan, best[0] if best else np.nan, np.nan, np.nan
    return 1000.0 / best[1], best[0], best[2], best[3]


def transfer_parameters(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows = []
    curves = []
    for key, frame in raw.groupby(KEYS, sort=True):
        frame = frame.dropna(subset=["GateV", "DrainI"]).sort_values("GateV")
        vgs = frame["GateV"].to_numpy(float)
        drain = frame["DrainI"].to_numpy(float)
        smoothed = smooth_current(drain)
        gm = np.gradient(smoothed, vgs)
        peak = int(np.nanargmax(gm))
        gm_max = float(gm[peak])
        vt = float(vgs[peak] - smoothed[peak] / gm_max) if gm_max != 0 else np.nan
        ss, ss_r2, ss_start, ss_end = extract_ss(vgs, drain)
        record = dict(zip(KEYS, key))
        record.update(
            {
                "vgs_min_v": float(vgs.min()),
                "vgs_max_v": float(vgs.max()),
                "points": len(vgs),
                "vt_tangent_v": vt,
                "gm_max_s": gm_max,
                "gm_peak_vgs_v": float(vgs[peak]),
                "ss_mv_dec": ss,
                "ss_r2": ss_r2,
                "ss_vgs_start_v": ss_start,
                "ss_vgs_end_v": ss_end,
            }
        )
        for point in CONFIG["transfer"]["common_vgs_points_v"]:
            label = str(point).replace("-", "m").replace(".", "p")
            record[f"abs_id_at_{label}v_a"] = abs(interpolate(vgs, drain, float(point)))
        rows.append(record)
        curve = frame[KEYS + ["point_no", "GateV", "DrainI"]].copy()
        curve["DrainI_smooth"] = smoothed
        curve["GM_recalc"] = gm
        curves.append(curve)
    result = pd.DataFrame(rows).sort_values(KEYS)
    baseline = result[result["dose_krad_si"] == 0][
        ["device_group", "device_no", "vt_tangent_v", "gm_max_s"]
    ].rename(columns={"vt_tangent_v": "vt_0_v", "gm_max_s": "gm_0_s"})
    result = result.merge(baseline, on=["device_group", "device_no"], how="left")
    result["delta_vt_v"] = result["vt_tangent_v"] - result["vt_0_v"]
    result["gm_change_pct"] = (result["gm_max_s"] / result["gm_0_s"] - 1) * 100
    return result, pd.concat(curves, ignore_index=True)


def first_crossing(vds: np.ndarray, current: np.ndarray, threshold: float) -> float:
    matches = np.flatnonzero(np.abs(current) >= threshold)
    return float(vds[matches[0]]) if len(matches) else np.nan


def output_parameters(raw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    cfg = CONFIG["output"]
    for key, frame in raw.groupby(KEYS, sort=True):
        frame = frame.dropna(subset=["DrainV", "DrainI"]).sort_values("DrainV")
        vds = frame["DrainV"].to_numpy(float)
        drain = frame["DrainI"].to_numpy(float)
        gate = frame["GateI"].to_numpy(float) if "GateI" in frame else np.full_like(drain, np.nan)
        record = dict(zip(KEYS, key))
        record.update(
            {
                "vds_min_v": float(vds.min()),
                "vds_max_v": float(vds.max()),
                "points": len(vds),
                "max_abs_gate_i_a": float(np.nanmax(np.abs(gate))),
                "compliance_onset_v": first_crossing(vds, drain, float(cfg["compliance_current_a"])),
            }
        )
        for point in cfg["vds_points_v"]:
            label = str(point).replace(".", "p")
            record[f"abs_id_at_{label}v_a"] = abs(interpolate(vds, drain, float(point)))
            record[f"abs_ig_at_{label}v_a"] = abs(interpolate(vds, gate, float(point)))
        for threshold in cfg["crossing_currents_a"]:
            exponent = int(round(-np.log10(float(threshold))))
            record[f"v_at_1e_m{exponent}a_v"] = first_crossing(vds, drain, float(threshold))
        low = vds <= float(cfg["low_field_fit_max_v"])
        fit = linregress(vds[low], drain[low])
        record["low_field_conductance_s"] = fit.slope
        record["low_field_fit_r2"] = fit.rvalue**2
        rows.append(record)
    result = pd.DataFrame(rows).sort_values(KEYS)
    baseline = result[(result["dose_krad_si"] == 0) & (result["scan_no"] == 1)]
    baseline = baseline[["device_group", "device_no", "abs_id_at_30p0v_a"]].rename(
        columns={"abs_id_at_30p0v_a": "abs_id_30v_0_a"}
    )
    result = result.merge(baseline, on=["device_group", "device_no"], how="left")
    result["id_30v_change_vs_0_pct"] = (
        result["abs_id_at_30p0v_a"] / result["abs_id_30v_0_a"] - 1
    ) * 100
    return result


def recovery_parameters(output: pd.DataFrame) -> pd.DataFrame:
    primary = output[output["scan_no"] == 1].copy()
    compare = output[output["scan_no"] > 1].copy()
    if compare.empty:
        return compare
    columns = ["abs_id_at_10p0v_a", "abs_id_at_20p0v_a", "abs_id_at_30p0v_a"]
    base = primary[["device_group", "device_no", "dose_krad_si", *columns]].rename(
        columns={column: f"first_{column}" for column in columns}
    )
    compare = compare.merge(base, on=["device_group", "device_no", "dose_krad_si"], how="left")
    for column in columns:
        first = f"first_{column}"
        suffix = column.removeprefix("abs_id_at_").removesuffix("_a")
        compare[f"leakage_reduction_{suffix}_pct"] = (compare[first] - compare[column]) / compare[first] * 100
        compare[f"log10_ratio_{suffix}"] = np.log10(compare[column] / compare[first])
    return compare


def summarize(frame: pd.DataFrame, group_cols: list[str], metrics: list[str]) -> pd.DataFrame:
    primary = frame[frame["scan_no"] == 1]
    summary = primary.groupby(group_cols)[metrics].agg(["count", "mean", "std", "min", "max"])
    summary.columns = [f"{metric}_{stat}" for metric, stat in summary.columns]
    return summary.reset_index()


def write_report(qc: pd.DataFrame, transfer: pd.DataFrame, output: pd.DataFrame, recovery: pd.DataFrame) -> None:
    errors = int(qc["error_cells"].sum())
    non_monotonic = int((~qc["x_strictly_increasing"].astype(bool)).sum())
    repeated = output[output["scan_no"] > 1]
    text = f"""# 数据质控与参数提取摘要

## 数据完整性

- 原始工作簿：28 个；原始文件保持只读。
- 数据工作表：{len(qc)} 个，其中转移 {int((qc['sweep'] == 'transfer').sum())} 个、输出 {int((qc['sweep'] == 'output').sum())} 个。
- 含错误标记的单元格总数：{errors}；这些值未直接用于参数计算。
- 横轴非严格递增的数据表：{non_monotonic} 个。
- 输出连续加压复测表：{len(repeated)} 个，均作为同一器件配对序列，不增加样本量。

## 统一计算

- 转移特性使用原始 `DrainI–GateV` 重算平滑电流、跨导和最大跨导切线阈值电压。
- `SS` 仅在对数电流滑动窗口满足 `R² ≥ {CONFIG['transfer']['ss_min_r2']}` 时保留。
- 输出特性提取 10/20/30 V 共同安全偏压电流、栅电流和 0–{CONFIG['output']['low_field_fit_max_v']} V 线性电导。
- 150 V 处大量曲线达到约 10 mA 合规限流，因此不作为剂量响应指标；高压行为改用达到指定电流及合规限流的起始电压描述。
- 所有剂量变化均按相同物理器件相对 0 krad(Si) 的基线计算。

## 数据边界

- S/T 仅作为样品组标签，缺少结构差异信息，不进行结构优劣归因。
- 输出扫描固定 `VGS=-20 V`，不能生成或解释多栅压输出曲线族。
- 加压后漏电变化只支持“恢复现象”描述；空间电荷重分布、去俘获、热效应和仪器稳定均保留为候选解释。

## 产出规模

- 转移参数记录：{len(transfer)} 条。
- 输出参数记录：{len(output)} 条。
- 加压恢复记录：{len(recovery)} 条。
"""
    (MANIFEST_DIR / "data_quality_report.md").write_text(text, encoding="utf-8")


def analyze() -> None:
    transfer_raw = pd.read_csv(DATA_DIR / "raw_transfer.csv")
    output_raw = pd.read_csv(DATA_DIR / "raw_output.csv")
    qc = pd.read_csv(DATA_DIR / "quality_control.csv")

    transfer, transfer_curves = transfer_parameters(transfer_raw)
    output = output_parameters(output_raw)
    recovery = recovery_parameters(output)

    transfer.to_csv(DATA_DIR / "transfer_parameters.csv", index=False, encoding="utf-8-sig")
    transfer_curves.to_csv(DATA_DIR / "transfer_curves_recalculated.csv", index=False, encoding="utf-8-sig")
    output.to_csv(DATA_DIR / "output_parameters.csv", index=False, encoding="utf-8-sig")
    recovery.to_csv(DATA_DIR / "recovery_parameters.csv", index=False, encoding="utf-8-sig")

    transfer_metrics = [
        "vt_tangent_v",
        "delta_vt_v",
        "gm_max_s",
        "gm_change_pct",
        "ss_mv_dec",
        "abs_id_at_m10p0v_a",
        "abs_id_at_m5p0v_a",
    ]
    output_metrics = [
        "abs_id_at_10p0v_a",
        "abs_id_at_20p0v_a",
        "abs_id_at_30p0v_a",
        "abs_ig_at_30p0v_a",
        "low_field_conductance_s",
        "id_30v_change_vs_0_pct",
        "v_at_1e_m6a_v",
        "v_at_1e_m3a_v",
        "compliance_onset_v",
    ]
    summarize(transfer, ["device_group", "dose_krad_si"], transfer_metrics).to_csv(
        DATA_DIR / "transfer_summary.csv", index=False, encoding="utf-8-sig"
    )
    summarize(output, ["device_group", "dose_krad_si"], output_metrics).to_csv(
        DATA_DIR / "output_summary.csv", index=False, encoding="utf-8-sig"
    )
    write_report(qc, transfer, output, recovery)
    print(
        json.dumps(
            {
                "transfer_parameters": len(transfer),
                "output_parameters": len(output),
                "recovery_parameters": len(recovery),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    analyze()