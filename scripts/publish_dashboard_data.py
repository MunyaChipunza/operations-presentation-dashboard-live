from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from refresh_dashboard_data import refresh_dashboard_data  # noqa: E402


CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Refresh dashboard_data.json and optionally push it to GitHub.")
    parser.add_argument("--workbook", help="Optional local workbook path.")
    parser.add_argument("--workbook-url", help="Optional public workbook URL for cloud refreshes.")
    parser.add_argument("--output", default="dashboard_data.json", help="Relative or absolute output path for the dashboard JSON.")
    parser.add_argument("--commit-message", default="Refresh operations dashboard data", help="Git commit message to use when the JSON changes.")
    return parser.parse_args()


def find_default_workbook(bundle_dir: Path) -> Path | None:
    search_roots = [bundle_dir.parent, bundle_dir.parent.parent]
    for root in search_roots:
        if not root.exists():
            continue
        preferred = sorted(root.glob("*PPT presentation source data*.xlsx"))
        if preferred:
            return preferred[0]
        for pattern in ("*.xlsx", "*.xlsm"):
            matches = sorted(root.glob(pattern))
            if matches:
                return matches[0]
    return None


def git_executable() -> str:
    git_path = shutil.which("git")
    if git_path:
        return git_path

    candidates = [
        Path("C:/Program Files/Git/cmd/git.exe"),
        Path("C:/Program Files/Git/bin/git.exe"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    raise FileNotFoundError("Git executable not found. Install Git for Windows first.")


def run_git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    command = [git_executable(), "-C", str(BUNDLE_DIR), *args]
    startupinfo = None
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0

    return subprocess.run(
        command,
        check=check,
        text=True,
        capture_output=True,
        creationflags=CREATE_NO_WINDOW,
        startupinfo=startupinfo,
    )


def is_git_repo() -> bool:
    result = run_git("rev-parse", "--is-inside-work-tree", check=False)
    return result.returncode == 0 and result.stdout.strip() == "true"


def has_origin() -> bool:
    result = run_git("remote", "get-url", "origin", check=False)
    return result.returncode == 0 and bool(result.stdout.strip())


def ensure_identity() -> None:
    name = run_git("config", "--get", "user.name", check=False)
    email = run_git("config", "--get", "user.email", check=False)
    if name.returncode == 0 and email.returncode == 0 and name.stdout.strip() and email.stdout.strip():
        return

    run_git("config", "user.name", "Operations Dashboard Sync")
    run_git("config", "user.email", "dashboard-sync@local")


def sync_repo() -> None:
    fetch = run_git("fetch", "origin", "main", check=False)
    if fetch.returncode != 0:
        raise RuntimeError(fetch.stderr.strip() or "Could not fetch origin/main before publishing.")

    rebase = run_git("rebase", "origin/main", check=False)
    if rebase.returncode != 0:
        raise RuntimeError(rebase.stderr.strip() or rebase.stdout.strip() or "Could not rebase onto origin/main before publishing.")


def has_dashboard_changes(output_path: Path) -> bool:
    rel_output = output_path.relative_to(BUNDLE_DIR)
    status = run_git("status", "--short", "--", str(rel_output), check=False)
    return bool(status.stdout.strip())


def push_dashboard(workbook_path: Path | None, workbook_url: str | None, output_path: Path, commit_message: str) -> bool:
    if not is_git_repo():
        refresh_dashboard_data(workbook=str(workbook_path) if workbook_path else None, workbook_url=workbook_url, output=output_path)
        print("Dashboard data refreshed locally. No Git repository detected, so nothing was pushed.")
        return False

    if not has_origin():
        refresh_dashboard_data(workbook=str(workbook_path) if workbook_path else None, workbook_url=workbook_url, output=output_path)
        print("Dashboard data refreshed locally. No origin remote is configured yet.")
        return False

    sync_repo()
    refresh_dashboard_data(workbook=str(workbook_path) if workbook_path else None, workbook_url=workbook_url, output=output_path)

    if not has_dashboard_changes(output_path):
        print("Dashboard data is already up to date.")
        return False

    ensure_identity()
    rel_output = output_path.relative_to(BUNDLE_DIR)
    run_git("add", "--", str(rel_output))
    run_git("commit", "-m", commit_message)
    run_git("push", "-u", "origin", "main")
    print("Dashboard data refreshed and pushed.")
    return True


def main() -> None:
    args = parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve() if args.workbook else find_default_workbook(BUNDLE_DIR)
    if args.workbook and (workbook_path is None or not workbook_path.exists()):
        raise FileNotFoundError("Could not find the local workbook to publish.")

    output_path = Path(args.output).expanduser()
    if not output_path.is_absolute():
        output_path = (BUNDLE_DIR / output_path).resolve()

    push_dashboard(
        workbook_path=workbook_path if workbook_path and workbook_path.exists() else None,
        workbook_url=args.workbook_url or os.environ.get("WORKBOOK_URL"),
        output_path=output_path,
        commit_message=args.commit_message,
    )


if __name__ == "__main__":
    main()
