from __future__ import annotations

import argparse
import base64
import datetime as dt
import html
import json
import os
import re
import statistics
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook


WORKBOOK_NAME_HINT = "PPT presentation source data"
TIMEZONE_NAME = "Africa/Johannesburg"
DEFAULT_OUTPUT = "dashboard_data.json"
LIVE_DATA_MAX_ROW = 100
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build dashboard JSON from the PPT presentation source workbook.")
    parser.add_argument("--workbook", help="Local workbook path (.xlsx/.xlsm).")
    parser.add_argument("--workbook-url", help="Public share or direct download URL for the workbook.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Where to write the dashboard JSON.")
    return parser.parse_args()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value).replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> int | None:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def is_whole_number(value: float) -> bool:
    return abs(value - round(value)) < 1e-9


def format_number(value: float | None, digits: int = 0) -> str:
    if value is None:
        return "-"
    if digits == 0 and is_whole_number(value):
        return f"{int(round(value)):,}"
    return f"{value:,.{digits}f}"


def format_percent(value: float | None, digits: int = 0) -> str:
    if value is None:
        return "-"
    return f"{value * 100:.{digits}f}%"


def format_month_label(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    normalized = text[:3].title()
    if normalized == "Sep":
        return "Sep"
    return normalized


def mean(values: list[float]) -> float | None:
    if not values:
        return None
    return statistics.fmean(values)


def tone_from_percent(value: float | None, *, good: float, warn: float, higher_is_better: bool = True) -> str:
    if value is None:
        return "quiet"
    if higher_is_better:
        if value >= good:
            return "good"
        if value >= warn:
            return "warn"
        return "bad"
    if value <= good:
        return "good"
    if value <= warn:
        return "warn"
    return "bad"


def tone_housekeeping(value: float | None) -> str:
    if value is None:
        return "quiet"
    if value < 0.50:
        return "bad"
    if value < 0.80:
        return "warn"
    return "good"


def tone_accuracy(value: float | None) -> str:
    if value is None:
        return "quiet"
    if value <= 0.97:
        return "bad"
    if value < 0.99:
        return "warn"
    return "good"


def tone_urgent_orders(value: float | None) -> str:
    if value is None:
        return "quiet"
    if value <= 0.05:
        return "good"
    if value <= 0.10:
        return "warn"
    return "bad"


def tone_from_rank(rank: int | None) -> str:
    if rank is None:
        return "quiet"
    if rank == 1:
        return "good"
    if rank == 2:
        return "warn"
    return "bad"


def tone_assembly_fill_rate(value: float | None) -> str:
    if value is None:
        return "quiet"
    if value < 0.70:
        return "bad"
    if value < 0.85:
        return "warn"
    return "good"


def latest_non_null(rows: list[dict[str, Any]], key: str) -> dict[str, Any] | None:
    for row in reversed(rows):
        if row.get(key) is not None:
            return row
    return None


def max_row(rows: list[dict[str, Any]], key: str) -> dict[str, Any] | None:
    with_values = [row for row in rows if row.get(key) is not None]
    if not with_values:
        return None
    return max(with_values, key=lambda row: row[key])


def min_row(rows: list[dict[str, Any]], key: str) -> dict[str, Any] | None:
    with_values = [row for row in rows if row.get(key) is not None]
    if not with_values:
        return None
    return min(with_values, key=lambda row: row[key])


def build_trendline(values: list[float | None]) -> list[float | None]:
    points = [(index, value) for index, value in enumerate(values) if value is not None]
    if len(points) < 2:
        return [None for _ in values]

    x_values = [point[0] for point in points]
    y_values = [point[1] for point in points]
    x_mean = statistics.fmean(x_values)
    y_mean = statistics.fmean(y_values)
    denominator = sum((x - x_mean) ** 2 for x in x_values)
    slope = 0 if denominator == 0 else sum((x - x_mean) * (y - y_mean) for x, y in points) / denominator
    intercept = y_mean - slope * x_mean
    return [intercept + slope * index for index in range(len(values))]


def constant_series(length: int, value: float | None) -> list[float | None]:
    if value is None:
        return [None for _ in range(length)]
    return [value for _ in range(length)]


def add_target_line(
    dataset: dict[str, Any],
    *,
    value: float | None,
    label: str,
    format_type: str,
    color: str = "#39ff88",
) -> None:
    if value is None:
        return
    labels = dataset["chart"].get("labels", [])
    dataset["chart"]["series"].append(
        {
            "name": label,
            "format": format_type,
            "color": color,
            "values": constant_series(len(labels), value),
            "style": "dashed",
            "showDots": False,
            "strokeWidth": 2,
        }
    )


def add_trendline(
    dataset: dict[str, Any],
    *,
    source_key: str,
    label: str,
    color: str = "#f8fafc",
) -> None:
    source = next((item for item in dataset["chart"]["series"] if item.get("key") == source_key), None)
    if not source:
        return
    dataset["chart"]["series"].append(
        {
            "name": label,
            "format": source["format"],
            "color": color,
            "values": build_trendline(source["values"]),
            "style": "dotted",
            "showDots": False,
            "strokeWidth": 2,
        }
    )


def share_token(url: str) -> str:
    raw = base64.b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return "u!" + raw.replace("/", "_").replace("+", "-")


def request_json(url: str, headers: dict[str, str] | None = None, data: bytes | None = None) -> dict[str, Any]:
    request = urllib.request.Request(url, headers=headers or {}, data=data)
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8"))


def request_bytes(url: str, headers: dict[str, str] | None = None) -> bytes:
    request = urllib.request.Request(url, headers=headers or {})
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read()


def write_temp_workbook(payload: bytes, suffix: str) -> Path:
    fd, temp_path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    path = Path(temp_path)
    path.write_bytes(payload)
    return path


def create_excel_snapshot(workbook_path: Path) -> Path:
    helper_script = Path(__file__).with_name("save_excel_snapshot.ps1")
    if not helper_script.exists():
        raise FileNotFoundError(f"Snapshot helper not found: {helper_script}")

    fd, temp_path = tempfile.mkstemp(suffix=workbook_path.suffix or ".xlsx")
    os.close(fd)
    snapshot_path = Path(temp_path)

    startupinfo = None
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0

    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(helper_script),
        "-SourcePath",
        str(workbook_path),
        "-TargetPath",
        str(snapshot_path),
    ]
    last_message = "Excel could not create a readable snapshot."
    for attempt in range(4):
        result = subprocess.run(
            command,
            check=False,
            text=True,
            capture_output=True,
            creationflags=CREATE_NO_WINDOW,
            startupinfo=startupinfo,
        )
        if result.returncode == 0:
            return snapshot_path

        last_message = result.stderr.strip() or result.stdout.strip() or last_message
        snapshot_path.unlink(missing_ok=True)
        if "0x800AC472" not in last_message or attempt == 3:
            break
        time.sleep(1.5)

    raise RuntimeError(last_message)


def load_dashboard_workbook(workbook_path: Path) -> tuple[Any, Path | None]:
    try:
        return load_workbook(workbook_path, data_only=True), None
    except PermissionError:
        snapshot_path = create_excel_snapshot(workbook_path)
        return load_workbook(snapshot_path, data_only=True), snapshot_path


def cleanup_temp_file(path: Path | None) -> None:
    if path is None:
        return
    for _ in range(6):
        try:
            path.unlink(missing_ok=True)
            return
        except PermissionError:
            time.sleep(0.5)


def find_download_url(payload: Any) -> str | None:
    if isinstance(payload, dict):
        for key, value in payload.items():
            if "downloadurl" in key.lower() and isinstance(value, str) and value.startswith("http"):
                return value
            found = find_download_url(value)
            if found:
                return found
    if isinstance(payload, list):
        for item in payload:
            found = find_download_url(item)
            if found:
                return found
    return None


def extract_download_url_from_html(page_html: str) -> str | None:
    patterns = [
        r'"downloadUrl":"([^"]+)"',
        r'"@content\.downloadUrl":"([^"]+)"',
        r'"downloadUrl\\":\\"([^"]+)\\"',
    ]
    for pattern in patterns:
        match = re.search(pattern, page_html)
        if not match:
            continue
        candidate = match.group(1)
        candidate = candidate.replace("\\u0026", "&").replace("\\/", "/").replace('\\"', '"')
        return html.unescape(candidate)
    return None


def with_download_hint(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    keys = {key.lower() for key, _ in query}
    if "download" not in keys:
        query.append(("download", "1"))
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, urllib.parse.urlencode(query), parsed.fragment))


def candidate_urls(url: str) -> list[str]:
    hinted = with_download_hint(url)
    return [url] if hinted == url else [url, hinted]


def ensure_excel_file(path: Path) -> None:
    if not zipfile.is_zipfile(path):
        raise ValueError(f"Downloaded file is not a valid Excel workbook: {path.name}")


def looks_like_excel_payload(payload: bytes) -> bool:
    return payload[:4] == b"PK\x03\x04"


def onedrive_badger_headers() -> dict[str, str]:
    token_payload = request_json(
        "https://api-badgerp.svc.ms/v1.0/token",
        headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
        data=json.dumps({"appId": "5cbed6ac-a083-4e14-b191-b4ba07653de2"}).encode("utf-8"),
    )
    token = token_payload.get("token")
    if not token:
        raise RuntimeError("Could not get OneDrive public access token.")
    return {
        "Authorization": f"Badger {token}",
        "Prefer": "autoredeem",
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, */*",
    }


def download_onedrive_share(url: str) -> tuple[Path, str]:
    token = share_token(url)
    headers = onedrive_badger_headers()
    metadata = request_json(f"https://api.onedrive.com/v1.0/shares/{token}/driveItem", headers=headers)
    download_url = find_download_url(metadata)
    if download_url:
        data = request_bytes(download_url, headers={"User-Agent": "Mozilla/5.0"})
        filename = metadata.get("name") or Path(urllib.parse.urlsplit(download_url).path).name or "dashboard_source.xlsx"
        suffix = Path(filename).suffix or ".xlsx"
        target = write_temp_workbook(data, suffix)
        ensure_excel_file(target)
        return target, filename

    raise RuntimeError("OneDrive share metadata loaded, but no downloadable workbook URL was available.")


def download_workbook(url: str) -> tuple[Path, str]:
    headers = {"User-Agent": "Mozilla/5.0 Codex Dashboard Refresher"}
    last_error: Exception | None = None

    for candidate in candidate_urls(url):
        try:
            request = urllib.request.Request(candidate, headers=headers)
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read()
                final_url = response.geturl()
                content_type = response.headers.get("Content-Type", "")
                if looks_like_excel_payload(payload) or "spreadsheet" in content_type.lower():
                    suffix = Path(urllib.parse.urlsplit(final_url).path).suffix or ".xlsx"
                    target = write_temp_workbook(payload, suffix)
                    ensure_excel_file(target)
                    filename = Path(urllib.parse.urlsplit(final_url).path).name or "dashboard_source.xlsx"
                    return target, filename
                page_html = payload.decode("utf-8", errors="ignore")
                nested_url = extract_download_url_from_html(page_html)
                if nested_url:
                    nested_payload = request_bytes(nested_url, headers=headers)
                    suffix = Path(urllib.parse.urlsplit(nested_url).path).suffix or ".xlsx"
                    target = write_temp_workbook(nested_payload, suffix)
                    ensure_excel_file(target)
                    filename = Path(urllib.parse.urlsplit(nested_url).path).name or "dashboard_source.xlsx"
                    return target, filename
        except Exception as exc:  # noqa: BLE001
            last_error = exc

    if "onedrive" in url.lower() or "1drv.ms" in url.lower() or "sharepoint" in url.lower():
        return download_onedrive_share(url)

    raise RuntimeError(f"Could not download workbook from URL: {last_error}")


def find_default_workbook(bundle_dir: Path) -> Path | None:
    search_roots = [bundle_dir.parent, bundle_dir.parent.parent]
    for root in search_roots:
        if not root.exists():
            continue
        preferred = sorted(root.glob(f"*{WORKBOOK_NAME_HINT}*.xlsx"))
        if preferred:
            return preferred[0]
        for pattern in ("*.xlsx", "*.xlsm"):
            matches = sorted(root.glob(pattern))
            if matches:
                return matches[0]
    return None


def series_dataset(
    *,
    key: str,
    label: str,
    category: str,
    description: str,
    label_key: str,
    label_title: str,
    rows: list[dict[str, Any]],
    series: list[dict[str, Any]],
    facts: list[dict[str, str]],
    headline_value: str,
    headline_detail: str,
    tone: str,
    note: str,
    chart_kind: str = "line",
) -> dict[str, Any]:
    columns = [{"key": label_key, "label": label_title, "format": "text"}] + [
        {"key": item["key"], "label": item["label"], "format": item["format"]}
        for item in series
    ]
    chart_rows = [row for row in rows if any(row.get(item["key"]) is not None for item in series)]
    return {
        "key": key,
        "label": label,
        "category": category,
        "description": description,
        "headlineValue": headline_value,
        "headlineDetail": headline_detail,
        "tone": tone,
        "note": note,
        "facts": facts,
        "table": {
            "columns": columns,
            "rows": rows,
        },
        "chart": {
            "kind": chart_kind,
            "labels": [row[label_key] for row in chart_rows],
            "series": [
                {
                    "name": item["label"],
                    "key": item["key"],
                    "format": item["format"],
                    "color": item["color"],
                    "values": [row.get(item["key"]) for row in chart_rows],
                    "style": item.get("style", "solid"),
                    "showDots": item.get("showDots", True),
                    "strokeWidth": item.get("strokeWidth", 3),
                }
                for item in series
            ],
        },
    }


def parse_housekeeping(ws) -> dict[str, Any]:
    rows = []
    average_92m = to_float(ws["B3"].value)
    average_12m = to_float(ws["C3"].value)
    rows.append({"period": "2025 Avg", "score92m": average_92m, "score12m": average_12m})
    for row_num in range(4, LIVE_DATA_MAX_ROW + 1):
        week = to_int(ws[f"A{row_num}"].value)
        score_92m = to_float(ws[f"B{row_num}"].value)
        score_12m = to_float(ws[f"C{row_num}"].value)
        if week is None and score_92m is None and score_12m is None:
            continue
        rows.append({"period": f"Week {week}", "score92m": score_92m, "score12m": score_12m})

    weekly_rows = [row for row in rows if row["period"].startswith("Week")]
    latest = latest_non_null(weekly_rows, "score92m")
    best = max_row(weekly_rows, "score92m")
    dataset = series_dataset(
        key="housekeeping",
        label="Housekeeping",
        category="Quality",
        description="Weekly housekeeping performance with both the 92M and 12M score streams.",
        label_key="period",
        label_title="Period",
        rows=rows,
        series=[
            {"key": "score92m", "label": "92M", "format": "percent", "color": "#00cfff"},
            {"key": "score12m", "label": "12M", "format": "percent", "color": "#ffb703"},
        ],
        facts=[
            {"label": "Latest", "value": format_percent(latest["score92m"], 0) if latest else "-"},
            {"label": "Target", "value": format_percent(0.80, 0)},
            {"label": "Average", "value": format_percent(average_92m, 0)},
            {"label": "Best Week", "value": f"{best['period']} - {format_percent(best['score92m'], 0)}" if best else "-"},
        ],
        headline_value=format_percent(latest["score92m"], 0) if latest else "-",
        headline_detail=f"{latest['period']} 92M reading" if latest else "No reading yet",
        tone=tone_housekeeping(latest["score92m"] if latest else None),
        note="Shows weekly housekeeping performance against the green 80% target line and the 2025 average benchmark.",
    )
    add_target_line(dataset, value=0.80, label="92M Target", format_type="percent")
    add_trendline(dataset, source_key="score92m", label="92M Trend")
    return dataset


def parse_single_series_block(
    ws,
    *,
    key: str,
    label: str,
    category: str,
    description: str,
    label_key: str,
    label_title: str,
    label_prefix: str,
    label_col: str,
    value_col: str,
    start_row: int,
    end_row: int,
    series_key: str,
    series_label: str,
    series_format: str,
    color: str,
    target_value: float,
    tone_resolver: Callable[[float | None], str],
    higher_is_better: bool = True,
    note: str = "",
) -> dict[str, Any]:
    rows = []
    for row_num in range(start_row, end_row + 1):
        raw_label = to_int(ws[f"{label_col}{row_num}"].value)
        raw_value = to_float(ws[f"{value_col}{row_num}"].value)
        if raw_label is None and raw_value is None:
            continue
        rows.append({label_key: f"{label_prefix} {raw_label}", series_key: raw_value})

    latest = latest_non_null(rows, series_key)
    best = max_row(rows, series_key) if higher_is_better else min_row(rows, series_key)
    values = [row[series_key] for row in rows if row.get(series_key) is not None]
    average_value = mean(values)
    dataset = series_dataset(
        key=key,
        label=label,
        category=category,
        description=description,
        label_key=label_key,
        label_title=label_title,
        rows=rows,
        series=[{"key": series_key, "label": series_label, "format": series_format, "color": color}],
        facts=[
            {"label": "Latest", "value": format_percent(latest[series_key], 1) if latest else "-"},
            {"label": "Target", "value": format_percent(target_value, 1)},
            {"label": "Average", "value": format_percent(average_value, 1)},
            {"label": "Best", "value": f"{best[label_key]} - {format_percent(best[series_key], 1)}" if best else "-"},
            {"label": "Readings", "value": str(len(values))},
        ],
        headline_value=format_percent(latest[series_key], 1) if latest else "-",
        headline_detail=f"{latest[label_key]} reading" if latest else "No reading yet",
        tone=tone_resolver(latest[series_key] if latest else None),
        note=note,
    )
    add_target_line(dataset, value=target_value, label="Target", format_type=series_format)
    add_trendline(dataset, source_key=series_key, label="Trend")
    return dataset


def parse_sku_share(ws) -> dict[str, Any]:
    rows = []
    for row_num in range(3, LIVE_DATA_MAX_ROW + 1):
        picker = clean_text(ws[f"T{row_num}"].value)
        count = to_float(ws[f"U{row_num}"].value)
        share = to_float(ws[f"V{row_num}"].value)
        if not picker:
            continue
        rows.append({"picker": picker, "count": count, "share": share})

    top_picker = max_row(rows, "count")
    total_units = sum(row["count"] for row in rows if row.get("count") is not None)
    return {
        "key": "sku-share",
        "label": "SKU Picked by Picker",
        "category": "Ranking",
        "description": "Team contribution split across the current SKU picker roster.",
        "headlineValue": format_number(top_picker["count"], 0) if top_picker else "-",
        "headlineDetail": f"Top picker: {top_picker['picker']}" if top_picker else "No picker data yet",
        "tone": "good",
        "note": "Donut view shows share concentration while the table keeps the exact counts visible.",
        "facts": [
            {"label": "Top Picker", "value": f"{top_picker['picker']} - {format_number(top_picker['count'], 0)}" if top_picker else "-"},
            {"label": "Top Share", "value": format_percent(top_picker["share"], 1) if top_picker else "-"},
            {"label": "Total Picked", "value": format_number(total_units, 0)},
            {"label": "Roster Size", "value": str(len(rows))},
        ],
        "table": {
            "columns": [
                {"key": "picker", "label": "Picker", "format": "text"},
                {"key": "count", "label": "SKU Count", "format": "integer"},
                {"key": "share", "label": "Share", "format": "percent"},
            ],
            "rows": rows,
        },
        "chart": {
            "kind": "donut",
            "series": [
                {"name": row["picker"], "value": row["count"], "share": row["share"], "color": color}
                for row, color in zip(
                    rows,
                    ["#ff8a5b", "#26d0ce", "#ffc857", "#8ddf6e", "#8a7dff"],
                    strict=False,
                )
            ],
        },
    }


def parse_monthly_sku(ws) -> dict[str, Any]:
    rows = []
    for row_num in range(3, LIVE_DATA_MAX_ROW + 1):
        month = format_month_label(ws[f"X{row_num}"].value)
        if not month:
            continue
        rows.append(
            {
                "month": month,
                "y2024": to_float(ws[f"Y{row_num}"].value),
                "y2025": to_float(ws[f"Z{row_num}"].value),
                "y2026": to_float(ws[f"AA{row_num}"].value),
            }
        )

    ytd_2026 = sum(row["y2026"] for row in rows if row.get("y2026") is not None)
    live_rows = [row for row in rows if row.get("y2026") is not None]
    ytd_2024 = sum(row["y2024"] for row in live_rows if row.get("y2024") is not None)
    ytd_2025 = sum(row["y2025"] for row in live_rows if row.get("y2025") is not None)
    ytd_values = sorted([ytd_2024, ytd_2025, ytd_2026], reverse=True)
    ytd_rank = ytd_values.index(ytd_2026) + 1 if ytd_values else None
    best_2025 = max_row(rows, "y2025")
    latest_2026 = latest_non_null(rows, "y2026")
    return {
        "key": "sku-monthly",
        "label": "SKU Picked by Month",
        "category": "Volume",
        "description": "Monthly SKU volume split by year so you can compare the current run rate with prior years.",
        "headlineValue": format_number(ytd_2026, 0),
        "headlineDetail": "2026 YTD picked volume",
        "tone": tone_from_rank(ytd_rank),
        "note": "The chart view now uses layered line work so 2026 can be read against the 2025 benchmark and its own trend.",
        "facts": [
            {"label": "2026 YTD", "value": format_number(ytd_2026, 0)},
            {"label": "Latest 2026 Month", "value": f"{latest_2026['month']} - {format_number(latest_2026['y2026'], 0)}" if latest_2026 else "-"},
            {"label": "Best 2025 Month", "value": f"{best_2025['month']} - {format_number(best_2025['y2025'], 0)}" if best_2025 else "-"},
            {"label": "YTD Rank", "value": f"{ytd_rank} of 3" if ytd_rank else "-"},
        ],
        "table": {
            "columns": [
                {"key": "month", "label": "Month", "format": "text"},
                {"key": "y2024", "label": "2024", "format": "integer"},
                {"key": "y2025", "label": "2025", "format": "integer"},
                {"key": "y2026", "label": "2026", "format": "integer"},
            ],
            "rows": rows,
        },
        "chart": {
            "kind": "line",
            "labels": [row["month"] for row in rows],
            "series": [
                {"name": "2024", "key": "y2024", "format": "integer", "color": "#3b82f6", "values": [row["y2024"] for row in rows], "style": "solid", "showDots": True, "strokeWidth": 2},
                {"name": "2025 Target", "key": "y2025", "format": "integer", "color": "#39ff88", "values": [row["y2025"] for row in rows], "style": "dashed", "showDots": False, "strokeWidth": 2},
                {"name": "2026 Actual", "key": "y2026", "format": "integer", "color": "#00cfff", "values": [row["y2026"] for row in rows], "style": "solid", "showDots": True, "strokeWidth": 4},
                {"name": "2026 Trend", "format": "integer", "color": "#f8fafc", "values": build_trendline([row["y2026"] for row in rows]), "style": "dotted", "showDots": False, "strokeWidth": 2},
            ],
        },
    }


def parse_assembly_backorders(ws) -> dict[str, Any]:
    rows = []
    for row_num in range(3, LIVE_DATA_MAX_ROW + 1):
        month = format_month_label(ws[f"AC{row_num}"].value)
        if not month:
            continue
        assembled = to_float(ws[f"AD{row_num}"].value)
        backorders = to_float(ws[f"AE{row_num}"].value)
        fill_rate = to_float(ws[f"AF{row_num}"].value)
        if fill_rate is None and assembled not in (None, 0) and backorders is not None:
            fill_rate = 1 - (backorders / assembled)
        rows.append(
            {
                "month": month,
                "assembled": assembled,
                "backorders": backorders,
                "fillRate": fill_rate,
            }
        )

    latest = latest_non_null(rows, "fillRate")
    best = max_row(rows, "fillRate")
    average_value = mean([row["fillRate"] for row in rows if row.get("fillRate") is not None])
    dataset = {
        "key": "assembly-backorders",
        "label": "Assembly vs Backorders",
        "category": "Production",
        "description": "Monthly Liseo assembly output against backorders, with fill rate calculated from the live workbook.",
        "headlineValue": format_percent(latest["fillRate"], 1) if latest else "-",
        "headlineDetail": f"{latest['month']} fill rate" if latest else "No reading yet",
        "tone": tone_assembly_fill_rate(latest["fillRate"] if latest else None),
        "note": "Table view keeps the assembled and backorder counts, while chart view focuses on the monthly fill-rate line against the green 85% target.",
        "facts": [
            {"label": "Latest", "value": format_percent(latest["fillRate"], 1) if latest else "-"},
            {"label": "Target", "value": format_percent(0.85, 0)},
            {"label": "Average", "value": format_percent(average_value, 1)},
            {"label": "Best", "value": f"{best['month']} - {format_percent(best['fillRate'], 1)}" if best else "-"},
        ],
        "table": {
            "columns": [
                {"key": "month", "label": "Month", "format": "text"},
                {"key": "assembled", "label": "Assembled", "format": "integer"},
                {"key": "backorders", "label": "Backorders", "format": "integer"},
                {"key": "fillRate", "label": "Fill Rate", "format": "percent"},
            ],
            "rows": rows,
        },
        "chart": {
            "kind": "line",
            "labels": [row["month"] for row in rows if row.get("fillRate") is not None],
            "series": [
                {
                    "name": "Fill Rate",
                    "key": "fillRate",
                    "format": "percent",
                    "color": "#8b5cf6",
                    "values": [row["fillRate"] for row in rows if row.get("fillRate") is not None],
                    "style": "solid",
                    "showDots": True,
                    "strokeWidth": 3,
                }
            ],
        },
    }
    add_target_line(dataset, value=0.85, label="Target", format_type="percent")
    add_trendline(dataset, source_key="fillRate", label="Trend")
    return dataset


def parse_points_yoy(ws) -> dict[str, Any]:
    rows = []
    for row_num in range(19, 31):
        month = format_month_label(ws[f"A{row_num}"].value)
        if not month:
            continue
        rows.append(
            {
                "month": month,
                "y2021": to_float(ws[f"B{row_num}"].value),
                "y2022": to_float(ws[f"C{row_num}"].value),
                "y2023": to_float(ws[f"D{row_num}"].value),
                "y2024": to_float(ws[f"E{row_num}"].value),
                "y2025": to_float(ws[f"F{row_num}"].value),
                "y2026": to_float(ws[f"G{row_num}"].value),
                "total": to_float(ws[f"H{row_num}"].value),
            }
        )

    total_2026 = to_float(ws["G31"].value)
    total_2025 = to_float(ws["F31"].value)
    peak_month = None
    for year_key in ("y2021", "y2022", "y2023", "y2024", "y2025", "y2026"):
        candidate = max_row(rows, year_key)
        if not candidate:
            continue
        if peak_month is None or candidate[year_key] > peak_month["value"]:
            peak_month = {"month": candidate["month"], "year": year_key[1:], "value": candidate[year_key]}

    return {
        "key": "points-yoy",
        "label": "Liseo Points YOY",
        "category": "Year over Year",
        "description": "Monthly points history across six years with a total-per-month column kept in the table view.",
        "headlineValue": format_number(total_2026, 1),
        "headlineDetail": "2026 year-to-date points",
        "tone": "good",
        "note": "The line view exposes long-term seasonality and where 2026 is still building up.",
        "facts": [
            {"label": "2026 YTD", "value": format_number(total_2026, 1)},
            {"label": "2025 Total", "value": format_number(total_2025, 0)},
            {"label": "Peak Month", "value": f"{peak_month['month']} {peak_month['year']} - {format_number(peak_month['value'], 1)}" if peak_month else "-"},
            {"label": "Benchmark", "value": "2025 line"},
        ],
        "table": {
            "columns": [
                {"key": "month", "label": "Month", "format": "text"},
                {"key": "y2021", "label": "2021", "format": "decimal1"},
                {"key": "y2022", "label": "2022", "format": "decimal1"},
                {"key": "y2023", "label": "2023", "format": "decimal1"},
                {"key": "y2024", "label": "2024", "format": "decimal1"},
                {"key": "y2025", "label": "2025", "format": "decimal1"},
                {"key": "y2026", "label": "2026", "format": "decimal1"},
                {"key": "total", "label": "Total", "format": "decimal1"},
            ],
            "rows": rows + [
                {
                    "month": "Total",
                    "y2021": to_float(ws["B31"].value),
                    "y2022": to_float(ws["C31"].value),
                    "y2023": to_float(ws["D31"].value),
                    "y2024": to_float(ws["E31"].value),
                    "y2025": to_float(ws["F31"].value),
                    "y2026": total_2026,
                    "total": to_float(ws["H31"].value),
                }
            ],
        },
        "chart": {
            "kind": "line",
            "labels": [row["month"] for row in rows],
            "series": [
                {"name": "2021", "key": "y2021", "format": "decimal1", "color": "#6c7385", "values": [row["y2021"] for row in rows], "style": "solid", "showDots": False, "strokeWidth": 1.5},
                {"name": "2022", "key": "y2022", "format": "decimal1", "color": "#ff8a5b", "values": [row["y2022"] for row in rows], "style": "solid", "showDots": False, "strokeWidth": 1.5},
                {"name": "2023", "key": "y2023", "format": "decimal1", "color": "#ffd54f", "values": [row["y2023"] for row in rows], "style": "solid", "showDots": False, "strokeWidth": 1.5},
                {"name": "2024", "key": "y2024", "format": "decimal1", "color": "#7ac7b1", "values": [row["y2024"] for row in rows], "style": "solid", "showDots": False, "strokeWidth": 1.5},
                {"name": "2025 Target", "key": "y2025", "format": "decimal1", "color": "#00cfff", "values": [row["y2025"] for row in rows], "style": "dashed", "showDots": False, "strokeWidth": 2},
                {"name": "2026 Actual", "key": "y2026", "format": "decimal1", "color": "#ff6b74", "values": [row["y2026"] for row in rows], "style": "solid", "showDots": True, "strokeWidth": 3.5},
                {"name": "2026 Trend", "format": "decimal1", "color": "#f5efeb", "values": build_trendline([row["y2026"] for row in rows]), "style": "dotted", "showDots": False, "strokeWidth": 2},
            ],
        },
    }


def build_dashboard(workbook_path: Path) -> dict[str, Any]:
    workbook = None
    snapshot_path: Path | None = None
    try:
        workbook, snapshot_path = load_dashboard_workbook(workbook_path)
        ws = workbook["DATA"]
        source_mtime = dt.datetime.fromtimestamp(workbook_path.stat().st_mtime, dt.timezone.utc)
        generated_at = dt.datetime.now(dt.timezone.utc)

        housekeeping = parse_housekeeping(ws)
        jhb = parse_single_series_block(
            ws,
            key="container-jhb",
            label="Container Accuracy JHB",
            category="Accuracy",
            description="Load-by-load accuracy readings for the Johannesburg container stream.",
            label_key="load",
            label_title="Load",
            label_prefix="Load",
            label_col="F",
            value_col="G",
            start_row=3,
            end_row=LIVE_DATA_MAX_ROW,
            series_key="accuracy",
            series_label="Accuracy",
            series_format="percent",
            color="#00cfff",
            target_value=0.99,
            tone_resolver=tone_accuracy,
            note="Higher is better. Missing loads are left blank so the chart stays honest.",
        )
        george = parse_single_series_block(
            ws,
            key="container-george",
            label="Container Accuracy George",
            category="Accuracy",
            description="Load-by-load accuracy readings for the George container stream.",
            label_key="load",
            label_title="Load",
            label_prefix="Load",
            label_col="I",
            value_col="J",
            start_row=3,
            end_row=LIVE_DATA_MAX_ROW,
            series_key="accuracy",
            series_label="Accuracy",
            series_format="percent",
            color="#3b82f6",
            target_value=0.99,
            tone_resolver=tone_accuracy,
            note="A shorter series today, but the same live refresh loop will grow it automatically.",
        )
        urgent = parse_single_series_block(
            ws,
            key="urgent-orders",
            label="Wholesaler Urgent Orders",
            category="Exceptions",
            description="Weekly urgent-order rate. Lower is healthier here.",
            label_key="week",
            label_title="Week",
            label_prefix="Week",
            label_col="M",
            value_col="N",
            start_row=3,
            end_row=LIVE_DATA_MAX_ROW,
            series_key="rate",
            series_label="Urgent Rate",
            series_format="percent",
            color="#ffb703",
            target_value=0.05,
            tone_resolver=tone_urgent_orders,
            higher_is_better=False,
            note="This one flips the interpretation: the closer to zero, the better.",
        )
        dispatch = parse_single_series_block(
            ws,
            key="dispatch-accuracy",
            label="Dispatch Accuracy",
            category="Accuracy",
            description="Weekly dispatch accuracy performance.",
            label_key="week",
            label_title="Week",
            label_prefix="Week",
            label_col="Q",
            value_col="R",
            start_row=3,
            end_row=LIVE_DATA_MAX_ROW,
            series_key="accuracy",
            series_label="Accuracy",
            series_format="percent",
            color="#ff5d73",
            target_value=0.99,
            tone_resolver=tone_accuracy,
            note="A near-perfect trend, so the chart leans on fine-grained percentage labels.",
        )
        sku_share = parse_sku_share(ws)
        monthly_sku = parse_monthly_sku(ws)
        assemblies = parse_assembly_backorders(ws)
        return {
            "title": "Operations Live Dashboard",
            "subtitle": "PPT Presentation Source",
            "sourceName": workbook_path.name,
            "generatedAt": generated_at.isoformat(),
            "sourceModifiedAt": source_mtime.isoformat(),
            "refreshSeconds": 60,
            "summaryCards": [
                {
                    "label": "Housekeeping",
                    "value": housekeeping["headlineValue"],
                    "detail": housekeeping["headlineDetail"],
                    "tone": housekeeping["tone"],
                },
                {
                    "label": "JHB Accuracy",
                    "value": jhb["headlineValue"],
                    "detail": jhb["headlineDetail"],
                    "tone": jhb["tone"],
                },
                {
                    "label": "George Accuracy",
                    "value": george["headlineValue"],
                    "detail": george["headlineDetail"],
                    "tone": george["tone"],
                },
                {
                    "label": "Dispatch Accuracy",
                    "value": dispatch["headlineValue"],
                    "detail": dispatch["headlineDetail"],
                    "tone": dispatch["tone"],
                },
                {
                    "label": "Urgent Orders",
                    "value": urgent["headlineValue"],
                    "detail": urgent["headlineDetail"],
                    "tone": urgent["tone"],
                },
                {
                    "label": "Top SKU Picker",
                    "value": sku_share["headlineValue"],
                    "detail": sku_share["headlineDetail"],
                    "tone": sku_share["tone"],
                },
                {
                    "label": "SKU 2026 YTD",
                    "value": monthly_sku["headlineValue"],
                    "detail": monthly_sku["headlineDetail"],
                    "tone": monthly_sku["tone"],
                },
                {
                    "label": "Assembly vs BO",
                    "value": assemblies["headlineValue"],
                    "detail": assemblies["headlineDetail"],
                    "tone": assemblies["tone"],
                },
            ],
            "datasets": [
                housekeeping,
                jhb,
                george,
                urgent,
                dispatch,
                sku_share,
                monthly_sku,
                assemblies,
            ],
        }
    finally:
        if workbook is not None:
            workbook.close()
        if snapshot_path is not None:
            cleanup_temp_file(snapshot_path)


def resolve_workbook(args: argparse.Namespace, bundle_dir: Path) -> tuple[Path, str]:
    if args.workbook:
        path = Path(args.workbook).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        return path, path.name

    if args.workbook_url:
        temp_path, filename = download_workbook(args.workbook_url)
        return temp_path, filename

    path = find_default_workbook(bundle_dir)
    if path is None:
        raise FileNotFoundError("Could not find a local workbook to build the dashboard from.")
    return path.resolve(), path.name


def refresh_dashboard_data(*, workbook: str | None = None, workbook_url: str | None = None, output: str | Path = DEFAULT_OUTPUT) -> Path:
    script_dir = Path(__file__).resolve().parent
    bundle_dir = script_dir.parent
    args = argparse.Namespace(workbook=workbook, workbook_url=workbook_url, output=output)
    workbook_path, _ = resolve_workbook(args, bundle_dir)

    output_path = Path(output).expanduser()
    if not output_path.is_absolute():
        output_path = (bundle_dir / output_path).resolve()

    payload = build_dashboard(workbook_path)
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return output_path


def main() -> None:
    args = parse_args()
    refresh_dashboard_data(
        workbook=args.workbook,
        workbook_url=args.workbook_url or os.environ.get("WORKBOOK_URL"),
        output=args.output,
    )


if __name__ == "__main__":
    main()
