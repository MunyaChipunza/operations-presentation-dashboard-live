from __future__ import annotations

import argparse
import os
import sys
import traceback
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")
sys.dont_write_bytecode = True


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
LOG_PATH = SCRIPT_DIR / "local_autopublish.log"
sys.path.insert(0, str(SCRIPT_DIR))

from publish_dashboard_data import find_default_workbook, push_dashboard  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the local dashboard auto-publish without opening a console window.")
    parser.add_argument("--workbook", help="Absolute path to the local workbook.")
    return parser.parse_args()


def log(message: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(message.rstrip() + "\n")


def resolve_workbook_path(candidate: str | None) -> Path:
    if candidate:
        preferred = Path(candidate).expanduser().resolve()
        if preferred.exists():
            return preferred
        log(f"Preferred workbook was missing, falling back to auto-detect: {preferred}")

    fallback = find_default_workbook(BUNDLE_DIR)
    if fallback and fallback.exists():
        return fallback.resolve()
    raise FileNotFoundError("No local workbook could be found for auto-publish.")


def main() -> int:
    args = parse_args()
    workbook_path = resolve_workbook_path(args.workbook)

    push_dashboard(
        workbook_path=workbook_path,
        workbook_url=None,
        output_path=BUNDLE_DIR / "dashboard_data.json",
        commit_message="Refresh operations dashboard data",
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:  # noqa: BLE001
        log(traceback.format_exc())
        raise SystemExit(1)
