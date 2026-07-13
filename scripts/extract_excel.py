from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path
from typing import Any

import pandas as pd
import win32com.client

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "deliverables" / "02_data"
MANIFEST_DIR = ROOT / "deliverables" / "05_manifest"

GROUP_DIRS = {
    "S_transfer": ROOT / "S转移特性曲线" / "S转移特性曲线",
    "T_transfer": ROOT / "T转移特性曲线",
    "S_output": ROOT / "S输出特性曲线",
    "T_output": ROOT / "T输出特性曲线",
}

FILE_RE = re.compile(r"v(?P<sweep>gs|ds)-id#(?P<group>[ST])-(?P<dose>\d+)K\.xlsx$", re.I)
SHEET_RE = re.compile(
    r"^(?P<group>[ST])(?P<device>[123])-(?P<dose>\d+)K(?:-(?P<scan>\d+))?$", re.I
)
IGNORED_SHEETS = {"calc", "settings"}
ERROR_PREFIX = "#"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def matrix(value: Any) -> list[list[Any]]:
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def numeric(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, str) and value.startswith(ERROR_PREFIX):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_settings(ws: Any, data_sheet_names: set[str], source: str) -> list[dict[str, Any]]:
    values = matrix(ws.UsedRange.Value2)
    current_sheet = None
    rows: list[dict[str, Any]] = []
    lookup = {name.lower(): name for name in data_sheet_names}
    for row_no, row in enumerate(values, start=1):
        cells = list(row) + [None] * (5 - len(row))
        first = str(cells[0]).strip() if cells[0] is not None else ""
        if first.lower() in lookup:
            current_sheet = lookup[first.lower()]
            continue
        if current_sheet and first and not set(first) <= {"=", " "}:
            rows.append(
                {
                    "source_file": source,
                    "sheet": current_sheet,
                    "row": row_no,
                    "key": first,
                    "value_1": cells[1],
                    "value_2": cells[2],
                    "value_3": cells[3],
                    "value_4": cells[4],
                }
            )
    return rows


def qc_record(meta: dict[str, Any], frame: pd.DataFrame) -> dict[str, Any]:
    sweep = meta["sweep"]
    x_col = "GateV" if sweep == "transfer" else "DrainV"
    id_values = pd.to_numeric(frame.get("DrainI"), errors="coerce")
    x_values = pd.to_numeric(frame.get(x_col), errors="coerce")
    diffs = x_values.diff().dropna()
    error_cells = sum(
        value.startswith(ERROR_PREFIX)
        for col in frame.columns
        for value in frame[col].dropna().astype(str)
    )
    record = {
        **meta,
        "rows": len(frame),
        "columns": ",".join(map(str, frame.columns)),
        "x_min": x_values.min(),
        "x_max": x_values.max(),
        "x_step_median": diffs.median(),
        "x_strictly_increasing": bool((diffs > 0).all()),
        "drain_i_min_a": id_values.min(),
        "drain_i_max_a": id_values.max(),
        "negative_drain_i_points": int((id_values < 0).sum()),
        "missing_drain_i_points": int(id_values.isna().sum()),
        "error_cells": int(error_cells),
    }
    if "GateI" in frame:
        gate_i = pd.to_numeric(frame["GateI"], errors="coerce")
        record["max_abs_gate_i_a"] = gate_i.abs().max()
    return record


def extract() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    MANIFEST_DIR.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    transfer_rows: list[dict[str, Any]] = []
    output_rows: list[dict[str, Any]] = []
    settings_rows: list[dict[str, Any]] = []
    workbook_rows: list[dict[str, Any]] = []
    qc_rows: list[dict[str, Any]] = []

    try:
        for group_name, folder in GROUP_DIRS.items():
            for path in sorted(folder.glob("*.xlsx")):
                file_match = FILE_RE.match(path.name)
                if not file_match:
                    continue
                file_meta = file_match.groupdict()
                sweep = "transfer" if file_meta["sweep"].lower() == "gs" else "output"
                workbook = excel.Workbooks.Open(str(path), 0, True)
                try:
                    data_names = {
                        ws.Name for ws in workbook.Worksheets if ws.Name.lower() not in IGNORED_SHEETS
                    }
                    workbook_rows.append(
                        {
                            "group_folder": group_name,
                            "source_file": str(path.relative_to(ROOT)),
                            "device_group": file_meta["group"].upper(),
                            "sweep": sweep,
                            "dose_krad_si": int(file_meta["dose"]),
                            "data_sheets": len(data_names),
                            "sheet_names": ",".join(sorted(data_names)),
                            "bytes": path.stat().st_size,
                            "modified_time": path.stat().st_mtime,
                            "sha256": sha256(path),
                        }
                    )
                    if "Settings" in [ws.Name for ws in workbook.Worksheets]:
                        settings_rows.extend(
                            parse_settings(workbook.Worksheets("Settings"), data_names, str(path.relative_to(ROOT)))
                        )
                    for sheet_name in sorted(data_names):
                        match = SHEET_RE.match(sheet_name)
                        if not match:
                            workbook_rows[-1]["unparsed_sheet"] = sheet_name
                            continue
                        sheet_meta = match.groupdict()
                        ws = workbook.Worksheets(sheet_name)
                        values = matrix(ws.UsedRange.Value2)
                        if len(values) < 2:
                            continue
                        headers = [str(value).strip() for value in values[0]]
                        frame = pd.DataFrame(values[1:], columns=headers)
                        meta = {
                            "source_file": str(path.relative_to(ROOT)),
                            "sheet": sheet_name,
                            "device_group": sheet_meta["group"].upper(),
                            "device_no": int(sheet_meta["device"]),
                            "device_id": f"{sheet_meta['group'].upper()}{sheet_meta['device']}",
                            "dose_krad_si": int(sheet_meta["dose"]),
                            "scan_no": int(sheet_meta["scan"] or 1),
                            "sweep": sweep,
                        }
                        qc_rows.append(qc_record(meta, frame))
                        target = transfer_rows if sweep == "transfer" else output_rows
                        for point_no, row in frame.iterrows():
                            record = {**meta, "point_no": int(point_no) + 1}
                            for header, value in row.items():
                                record[header] = numeric(value)
                            target.append(record)
                finally:
                    workbook.Close(False)
    finally:
        excel.Quit()

    pd.DataFrame(transfer_rows).to_csv(DATA_DIR / "raw_transfer.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(output_rows).to_csv(DATA_DIR / "raw_output.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(settings_rows).to_csv(DATA_DIR / "settings_rows.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(qc_rows).to_csv(DATA_DIR / "quality_control.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(workbook_rows).to_csv(MANIFEST_DIR / "source_manifest.csv", index=False, encoding="utf-8-sig")

    summary = {
        "workbooks": len(workbook_rows),
        "transfer_rows": len(transfer_rows),
        "output_rows": len(output_rows),
        "quality_records": len(qc_rows),
        "settings_records": len(settings_rows),
    }
    (MANIFEST_DIR / "extraction_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    extract()