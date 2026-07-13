from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "deliverables" / "02_data"
ORIGIN_DIR = ROOT / "deliverables" / "01_origin"
FIGURE_DIR = ROOT / "deliverables" / "03_figures"
MANIFEST_DIR = ROOT / "deliverables" / "05_manifest"
DOSES = [0, 10, 20, 30, 40, 50, 60]

# Wong color-blind-safe palette. Line styles provide a second channel for
# readers using grayscale printouts or with impaired color perception.
DOSE_STYLES = [
    {"color": "#000000", "line": 0, "symbol": 3},
    {"color": "#0072B2", "line": 1, "symbol": 4},
    {"color": "#E69F00", "line": 2, "symbol": 5},
    {"color": "#009E73", "line": 3, "symbol": 6},
    {"color": "#D55E00", "line": 4, "symbol": 7},
    {"color": "#CC79A7", "line": 5, "symbol": 8},
    {"color": "#56B4E9", "line": 6, "symbol": 9},
]
GROUP_STYLES = [
    {"color": "#0072B2", "line": 0, "symbol": 3},
    {"color": "#D55E00", "line": 1, "symbol": 4},
]
RECOVERY_STYLES = [
    {"color": "#0072B2", "line": 0, "symbol": 3},
    {"color": "#D55E00", "line": 1, "symbol": 4},
    {"color": "#009E73", "line": 2, "symbol": 5},
]


def pad_columns(columns: list[np.ndarray]) -> np.ndarray:
    rows = max(len(column) for column in columns)
    matrix = np.full((rows, len(columns)), np.nan)
    for index, column in enumerate(columns):
        matrix[: len(column), index] = column
    return matrix


def mean_curve(raw: pd.DataFrame, group: str, x: str, y: str) -> tuple[np.ndarray, list[dict[str, Any]]]:
    source = raw[(raw["device_group"] == group) & (raw["scan_no"] == 1)]
    columns: list[np.ndarray] = []
    metadata: list[dict[str, Any]] = []
    for dose in DOSES:
        curve = source[source["dose_krad_si"] == dose].groupby(x, as_index=False)[y].mean().sort_values(x)
        columns.extend([curve[x].to_numpy(float), np.abs(curve[y].to_numpy(float))])
        metadata.extend(
            [
                {"long_name": f"{dose} krad(Si) {x}", "units": "V", "type": 3},
                {"long_name": f"{dose} krad(Si)", "units": "A", "type": 0},
            ]
        )
    return pad_columns(columns), metadata


def summary_matrix(summary: pd.DataFrame, metric: str) -> tuple[np.ndarray, list[dict[str, Any]]]:
    columns: list[np.ndarray] = []
    metadata: list[dict[str, Any]] = []
    for group in ["S", "T"]:
        frame = summary[summary["device_group"] == group].sort_values("dose_krad_si")
        columns.extend(
            [
                frame["dose_krad_si"].to_numpy(float),
                frame[f"{metric}_mean"].to_numpy(float),
                frame[f"{metric}_std"].fillna(0).to_numpy(float),
            ]
        )
        metadata.extend(
            [
                {"long_name": f"{group} dose", "units": "krad(Si)", "type": 3},
                {"long_name": f"{group} mean", "units": "", "type": 0},
                {"long_name": f"{group} SD", "units": "", "type": 2},
            ]
        )
    return pad_columns(columns), metadata


def recovery_matrix(recovery: pd.DataFrame) -> tuple[np.ndarray, list[dict[str, Any]]]:
    columns: list[np.ndarray] = []
    metadata: list[dict[str, Any]] = []
    for scan in sorted(recovery["scan_no"].unique()):
        frame = (
            recovery[recovery["scan_no"] == scan]
            .groupby("dose_krad_si", as_index=False)["leakage_reduction_30p0v_pct"]
            .agg(["mean", "std"])
            .reset_index()
            .sort_values("dose_krad_si")
        )
        columns.extend(
            [
                frame["dose_krad_si"].to_numpy(float),
                frame["mean"].to_numpy(float),
                frame["std"].fillna(0).to_numpy(float),
            ]
        )
        metadata.extend(
            [
                {"long_name": f"Scan {int(scan)} dose", "units": "krad(Si)", "type": 3},
                {"long_name": f"Scan {int(scan)} vs first", "units": "%", "type": 0},
                {"long_name": f"Scan {int(scan)} SD", "units": "%", "type": 2},
            ]
        )
    return pad_columns(columns), metadata


def create_workbook(app: Any, name: str, long_name: str, data: np.ndarray, columns: list[dict[str, Any]]) -> Any:
    page_name = app.CreatePage(2, name, "Origin")
    worksheet = app.FindWorksheet(page_name)
    worksheet.Name = name
    worksheet.LongName = long_name
    worksheet.Cols = data.shape[1]
    worksheet.SetData(data.tolist(), 0, 0)
    for index, info in enumerate(columns):
        column = worksheet.Columns.Item(index)
        column.LongName = info["long_name"]
        column.Units = info.get("units", "")
        column.Type = info["type"]
    return worksheet


def style_plot(
    layer: Any,
    plot_index: int,
    style: dict[str, Any],
    width: float,
    show_symbol: bool,
    force_solid: bool = False,
) -> None:
    color = style["color"]
    commands = [
        f"range rp = 1!{plot_index}",
        f'set rp -cl color("{color}")',
        f'set rp -cse color("{color}")',
        f'set rp -csf color("{color}")',
        f"set rp -wp {width}",
        f"set rp -d {0 if force_solid else style['line']}",
        "set rp -l 1",
    ]
    if show_symbol:
        commands.extend(
            [
                f"set rp -k {style['symbol']}",
                "set rp -z 7",
                "set rp -kh 1",
                "set rp -erw 1",
                "set rp -erwc 8",
            ]
        )
    layer.Execute("; ".join(commands) + ";")


def create_graph(
    app: Any,
    name: str,
    long_name: str,
    worksheet: Any,
    series: list[tuple[int, int, int | None, str]],
    x_title: str,
    y_title: str,
    log_y: bool = False,
    line_plot: bool = True,
    style_family: str = "dose",
    force_solid: bool = False,
    wide: bool = False,
) -> str:
    graph_name = app.CreatePage(3, name, "Origin")
    layer = app.FindGraphLayer(graph_name)
    layer.LongName = long_name
    if wide:
        layer.Execute(
            "page.width=page.width*1.35; "
            "layer.unit=1; layer.left=12; layer.top=16; "
            "layer.width=80; layer.height=62;"
        )
    plots = layer.DataPlots
    styles = {
        "dose": DOSE_STYLES,
        "group": GROUP_STYLES,
        "recovery": RECOVERY_STYLES,
    }[style_family]
    show_symbol = not line_plot
    for index, (x_col, y_col, error_col, label) in enumerate(series):
        data_range = app.NewDataRange()
        data_range.Add("X", worksheet, 0, x_col, -1, x_col)
        data_range.Add("Y", worksheet, 0, y_col, -1, y_col)
        if error_col is not None:
            data_range.Add("ED", worksheet, 0, error_col, -1, error_col)
        plot = plots.Add(data_range, 202 if show_symbol else 200)
        plot.LongName = label
        style_plot(
            layer,
            index + 1,
            styles[index % len(styles)],
            1.5 if line_plot else 1.2,
            show_symbol,
            force_solid,
        )
    layer.Execute("rescale;")
    layer.Execute(f'xb.text$ = "{x_title}";')
    layer.Execute(f'yl.text$ = "{y_title}";')
    layer.Execute("range ll = !; ll.x2.showAxes=3; ll.y2.showAxes=3;")
    layer.Execute("legend -s;")
    layer.Execute(
        "int pubfont=font(Arial); "
        "xb.font=pubfont; yl.font=pubfont; legend.font=pubfont; "
        "layer.x.label.font=pubfont; layer.y.label.font=pubfont; "
        "xb.fsize=22; yl.fsize=22; legend.fsize=18; "
        "layer.x.label.pt=18; layer.y.label.pt=18;"
    )
    if log_y:
        layer.Execute("layer.y.type=2; rescale;")
    return graph_name


def export_graph(app: Any, graph_name: str, base_name: str, width_cm: float = 16) -> dict[str, str]:
    app.Execute(f"win -a {graph_name};")
    outputs: dict[str, str] = {}
    formats = [
        (
            "pdf",
            "pdf",
            "pdf",
            "tr.Margin:=1 tr.Advanced.Resolution:=0 tr.Advanced.DPI:=600 "
            "tr2.PDF.PDF.ColorTranslation:=0 tr2.PDF.Fonts.Embed:=1 "
            "tr2.PDF.Fonts.TrueType:=1",
        ),
        ("emf", "emf", "emf", "tr.Margin:=1 tr.Advanced.Resolution:=0 tr.Advanced.DPI:=600"),
        (
            "png",
            "png",
            "png",
            f'tr.Margin:=1 tr1.Unit:=1 tr1.Rescaling:=0 tr1.Width:={width_cm} '
            'tr2.PNG.dotsperinch:=600 tr2.PNG.bitsperpixel:="24-bit Color"',
        ),
    ]
    for extension, export_type, subdir, settings in formats:
        folder = FIGURE_DIR / subdir
        folder.mkdir(parents=True, exist_ok=True)
        app.Execute(
            f'expGraph type:={export_type} filename:="{base_name}" path:="{folder.as_posix()}" '
            f"{settings} overwrite:=replace;"
        )
        path = folder / f"{base_name}.{extension}"
        outputs[extension] = str(path.relative_to(ROOT))
    return outputs


def build() -> None:
    ORIGIN_DIR.mkdir(parents=True, exist_ok=True)
    MANIFEST_DIR.mkdir(parents=True, exist_ok=True)
    transfer_raw = pd.read_csv(DATA_DIR / "raw_transfer.csv")
    output_raw = pd.read_csv(DATA_DIR / "raw_output.csv")
    transfer_summary = pd.read_csv(DATA_DIR / "transfer_summary.csv")
    output_summary = pd.read_csv(DATA_DIR / "output_summary.csv")
    recovery = pd.read_csv(DATA_DIR / "recovery_parameters.csv")

    app = win32com.client.gencache.EnsureDispatch("Origin.ApplicationSI")
    app.NewProject()
    manifest: list[dict[str, Any]] = []
    try:
        for group in ["S", "T"]:
            transfer_data, transfer_columns = mean_curve(transfer_raw, group, "GateV", "DrainI")
            transfer_sheet = create_workbook(
                app, f"TrMean{group}", f"{group} transfer mean curves", transfer_data, transfer_columns
            )
            series = [(2 * i, 2 * i + 1, None, f"{dose} krad(Si)") for i, dose in enumerate(DOSES)]
            linear = create_graph(
                app,
                f"GTrLin{group}",
                f"{group} transfer characteristics",
                transfer_sheet,
                series,
                "Gate voltage, V\\-(GS) (V)",
                "Drain current, |I\\-(D)| (A)",
            )
            log = create_graph(
                app,
                f"GTrLog{group}",
                f"{group} transfer characteristics, semilog",
                transfer_sheet,
                series,
                "Gate voltage, V\\-(GS) (V)",
                "Drain current, |I\\-(D)| (A)",
                log_y=True,
            )
            manifest.append({"graph": linear, "base_name": f"transfer_linear_{group}", **export_graph(app, linear, f"transfer_linear_{group}")})
            manifest.append({"graph": log, "base_name": f"transfer_semilog_{group}", **export_graph(app, log, f"transfer_semilog_{group}")})

            output_data, output_columns = mean_curve(output_raw, group, "DrainV", "DrainI")
            output_sheet = create_workbook(
                app, f"OutMean{group}", f"{group} output mean curves", output_data, output_columns
            )
            output_graph = create_graph(
                app,
                f"GOutLog{group}",
                f"{group} output characteristics, semilog",
                output_sheet,
                series,
                "Drain voltage, V\\-(DS) (V)",
                "Drain current, |I\\-(D)| (A)",
                log_y=True,
                force_solid=True,
                wide=True,
            )

            manifest.append(
                {
                    "graph": output_graph,
                    "base_name": f"output_semilog_{group}",
                    **export_graph(app, output_graph, f"output_semilog_{group}", width_cm=22),
                }
            )

        parameter_specs = [
            (transfer_summary, "vt_tangent_v", "ParVth", "GParVth", "threshold_voltage", "Threshold voltage, V\\-(th) (V)", False),
            (transfer_summary, "delta_vt_v", "ParDVth", "GParDVth", "threshold_shift", "Threshold-voltage shift (V)", False),
            (transfer_summary, "gm_change_pct", "ParGm", "GParGm", "gm_change", "Peak transconductance change (%)", False),
            (transfer_summary, "ss_mv_dec", "ParSS", "GParSS", "subthreshold_swing", "Subthreshold swing (mV/decade)", False),
            (output_summary, "abs_id_at_30p0v_a", "ParI30", "GParI30", "output_current_30v", "Drain current at 30 V, |I\\-(D)| (A)", True),
            (output_summary, "v_at_1e_m6a_v", "ParV1u", "GParV1u", "voltage_at_1ua", "Voltage at |I\\-(D)| = 1e-6 A (V)", False),
        ]
        for summary, metric, book, graph, base, y_title, log_y in parameter_specs:
            data, columns = summary_matrix(summary, metric)
            worksheet = create_workbook(app, book, base, data, columns)
            parameter_graph = create_graph(
                app,
                graph,
                base,
                worksheet,
                [(0, 1, 2, "S"), (3, 4, 5, "T")],
                "Total ionizing dose (krad(Si))",
                y_title,
                log_y=log_y,
                line_plot=False,
                style_family="group",
            )
            manifest.append({"graph": parameter_graph, "base_name": base, **export_graph(app, parameter_graph, base)})

        recovery_data, recovery_columns = recovery_matrix(recovery)
        recovery_sheet = create_workbook(
            app, "Recovery", "S electrical-stress recovery at 30 V", recovery_data, recovery_columns
        )
        recovery_series = [
            (3 * i, 3 * i + 1, 3 * i + 2, f"Scan {int(scan)} vs first")
            for i, scan in enumerate(sorted(recovery["scan_no"].unique()))
        ]
        recovery_graph = create_graph(
            app,
            "GRecovery",
            "S electrical-stress recovery",
            recovery_sheet,
            recovery_series,
            "Total ionizing dose (krad(Si))",
            "Leakage-current reduction at 30 V (%)",
            line_plot=False,
            style_family="recovery",
        )
        manifest.append({"graph": recovery_graph, "base_name": "stress_recovery_30v", **export_graph(app, recovery_graph, "stress_recovery_30v")})

        project = ORIGIN_DIR / "radiation_analysis.opju"
        app.Save(str(project))
        (MANIFEST_DIR / "figure_manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(json.dumps({"project": str(project), "figures": len(manifest)}, ensure_ascii=False))
    finally:
        app.Exit()


if __name__ == "__main__":
    build()