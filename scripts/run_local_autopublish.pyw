from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
import traceback
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")
sys.dont_write_bytecode = True


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
LOG_PATH = SCRIPT_DIR / "local_autopublish.log"
CONFIG_PATH = SCRIPT_DIR / "workbook_source.local.json"
sys.path.insert(0, str(SCRIPT_DIR))

from refresh_dashboard_data import choose_preferred_source, path_is_excel  # noqa: E402
from publish_dashboard_data import find_default_workbook, push_dashboard  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the local dashboard auto-publish without opening a console window.")
    parser.add_argument("--workbook", help="Absolute path to the local workbook.")
    parser.add_argument("--workbook-url", help="Optional public share or direct-download URL for the live workbook.")
    return parser.parse_args()


def log(message: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        stamp = dt.datetime.now().isoformat(timespec="seconds")
        handle.write(f"[{stamp}] {message.rstrip()}\n")


def load_source_config() -> dict[str, str]:
    if not CONFIG_PATH.exists():
        return {}

    try:
        payload = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        log(f"Could not read source config {CONFIG_PATH}: {exc}")
        return {}

    if not isinstance(payload, dict):
        log(f"Ignoring non-object source config at {CONFIG_PATH}")
        return {}

    return {str(key): str(value) for key, value in payload.items() if value not in (None, "")}


def resolve_workbook_url(candidate: str | None, config: dict[str, str]) -> str | None:
    for value in (
        candidate,
        os.environ.get("WORKBOOK_URL"),
        config.get("workbookUrl"),
        config.get("workbook_url"),
    ):
        if value and value.strip():
            return value.strip()
    return None


def resolve_configured_workbook_path(config: dict[str, str]) -> str | None:
    for value in (config.get("workbookPath"), config.get("workbook_path")):
        if value and value.strip():
            return value.strip()
    return None


def resolve_workbook_path(candidate: str | None) -> Path:
    if candidate:
        preferred = Path(candidate).expanduser().resolve()
        if preferred.exists():
            return preferred
        log(f"Preferred workbook was missing, falling back to auto-detect: {preferred}")

    fallback = find_default_workbook(BUNDLE_DIR)
    if fallback and fallback.exists():
        return fallback.resolve()
    source_path = choose_preferred_source(None, BUNDLE_DIR)
    if source_path and source_path.exists():
        return source_path.resolve()
    raise FileNotFoundError("No local workbook could be found for auto-publish.")


def main() -> int:
    args = parse_args()
    config = load_source_config()
    workbook_url = resolve_workbook_url(args.workbook_url, config)
    workbook_hint = args.workbook or resolve_configured_workbook_path(config)
    workbook_path = resolve_workbook_path(workbook_hint)
    source_path = workbook_path
    log(f"Auto-publish start. Workbook candidate: {workbook_path}")

    if path_is_excel(workbook_path):
        preferred_source = choose_preferred_source(workbook_path, BUNDLE_DIR)
        if preferred_source and preferred_source != workbook_path:
            source_path = preferred_source
            log(f"Found fresher source for publishing: {preferred_source}")
    else:
        source_path = choose_preferred_source(workbook_path, BUNDLE_DIR) or workbook_path

    if workbook_url:
        log("Workbook URL configured. Cloud workbook will be preferred when reachable.")

    try:
        published = push_dashboard(
            workbook_path=source_path,
            workbook_url=workbook_url,
            output_path=BUNDLE_DIR / "dashboard_data.json",
            commit_message="Refresh operations dashboard data",
        )
        source_mode = "cloud" if workbook_url else "local"
    except Exception as exc:
        if not workbook_url:
            raise

        log(f"Cloud refresh failed, falling back to local workbook. Error: {exc}")
        published = push_dashboard(
            workbook_path=source_path,
            workbook_url=None,
            output_path=BUNDLE_DIR / "dashboard_data.json",
            commit_message="Refresh operations dashboard data",
        )
        source_mode = "local-fallback"

    log(f"Auto-publish completed. Published={published}. Source used: {source_path}. Mode={source_mode}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:  # noqa: BLE001
        log(traceback.format_exc())
        raise SystemExit(1)
