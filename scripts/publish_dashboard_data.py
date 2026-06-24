from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from refresh_dashboard_data import refresh_dashboard_data  # noqa: E402


CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
WORKBOOK_NAME_HINTS = ("PPT presentation source data", "Operations Data")
TRANSIENT_GIT_ERRORS = (
    "Could not resolve host",
    "Failed to connect",
    "Connection was reset",
    "Recv failure",
    "Operation timed out",
    "TLS",
    "SSL",
)


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
        for hint in WORKBOOK_NAME_HINTS:
            for pattern in (f"*{hint}*.xlsx", f"*{hint}*.xlsm"):
                preferred = sorted(root.glob(pattern))
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


def run_git(*args: str, check: bool = True, repo_dir: Path | None = BUNDLE_DIR) -> subprocess.CompletedProcess[str]:
    command = [git_executable()]
    if repo_dir is not None:
        command.extend(["-C", str(repo_dir)])
    command.extend(args)
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


def git_error_text(result: subprocess.CompletedProcess[str]) -> str:
    return (result.stderr or result.stdout or "").strip()


def is_transient_git_failure(message: str) -> bool:
    lowered = message.lower()
    return any(fragment.lower() in lowered for fragment in TRANSIENT_GIT_ERRORS)


def run_git_with_retry(
    *args: str,
    attempts: int = 4,
    delay_seconds: float = 5.0,
    repo_dir: Path | None = BUNDLE_DIR,
) -> subprocess.CompletedProcess[str]:
    last_result: subprocess.CompletedProcess[str] | None = None
    for attempt in range(1, attempts + 1):
        result = run_git(*args, check=False, repo_dir=repo_dir)
        if result.returncode == 0:
            return result

        last_result = result
        if attempt == attempts or not is_transient_git_failure(git_error_text(result)):
            break
        time.sleep(delay_seconds * attempt)

    assert last_result is not None
    return last_result


def is_git_repo() -> bool:
    result = run_git("rev-parse", "--is-inside-work-tree", check=False)
    return result.returncode == 0 and result.stdout.strip() == "true"


def has_origin() -> bool:
    result = run_git("remote", "get-url", "origin", check=False)
    return result.returncode == 0 and bool(result.stdout.strip())


def origin_url() -> str | None:
    result = run_git("remote", "get-url", "origin", check=False)
    if result.returncode != 0:
        return None
    value = result.stdout.strip()
    return value or None


def ensure_identity() -> None:
    name = run_git("config", "--get", "user.name", check=False)
    email = run_git("config", "--get", "user.email", check=False)
    if name.returncode == 0 and email.returncode == 0 and name.stdout.strip() and email.stdout.strip():
        return

    run_git("config", "user.name", "Operations Dashboard Sync")
    run_git("config", "user.email", "dashboard-sync@local")


def local_ahead_count() -> int:
    result = run_git("rev-list", "--count", "origin/main..HEAD", check=False)
    if result.returncode != 0:
        return 0
    try:
        return int(result.stdout.strip() or "0")
    except ValueError:
        return 0


def sync_repo() -> None:
    fetch = run_git_with_retry("fetch", "origin", "main")
    if fetch.returncode != 0:
        raise RuntimeError(git_error_text(fetch) or "Could not fetch origin/main before publishing.")

    rebase = run_git("rebase", "origin/main", check=False)
    if rebase.returncode != 0:
        raise RuntimeError(git_error_text(rebase) or "Could not rebase onto origin/main before publishing.")


def has_dashboard_changes(output_path: Path) -> bool:
    rel_output = output_path.relative_to(BUNDLE_DIR)
    status = run_git("status", "--short", "--", str(rel_output), check=False)
    return bool(status.stdout.strip())


def write_local_output(source_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_path, output_path)


def ensure_identity_for_repo(repo_dir: Path) -> None:
    run_git("config", "user.name", "Operations Dashboard Sync", repo_dir=repo_dir)
    run_git("config", "user.email", "dashboard-sync@local", repo_dir=repo_dir)


def push_output_from_clean_clone(output_path: Path, rel_output: Path, commit_message: str) -> bool:
    remote_url = origin_url()
    if not remote_url:
        print("Dashboard data refreshed locally. No origin remote is configured yet.")
        return False

    clone_dir = Path(tempfile.mkdtemp(prefix="ops-dashboard-publish-"))
    try:
        clone = run_git_with_retry(
            "clone",
            "--branch",
            "main",
            "--single-branch",
            remote_url,
            str(clone_dir),
            repo_dir=None,
        )
        if clone.returncode != 0:
            print(
                "Dashboard data refreshed locally. Git clone skipped: "
                f"{git_error_text(clone) or 'Could not clone origin/main for publishing.'}"
            )
            return False

        clone_output = clone_dir / rel_output
        write_local_output(output_path, clone_output)

        status = run_git("status", "--short", "--", rel_output.as_posix(), check=False, repo_dir=clone_dir)
        if not status.stdout.strip():
            print("Dashboard data is already up to date.")
            return False

        ensure_identity_for_repo(clone_dir)
        run_git("add", "--", rel_output.as_posix(), repo_dir=clone_dir)
        commit = run_git("commit", "-m", commit_message, check=False, repo_dir=clone_dir)
        if commit.returncode != 0:
            print(
                "Dashboard data refreshed locally. Git commit skipped: "
                f"{git_error_text(commit) or 'Could not commit dashboard data.'}"
            )
            return False

        push = run_git_with_retry("push", "origin", "main", repo_dir=clone_dir)
        if push.returncode != 0:
            print(
                "Dashboard data refreshed locally. Git push skipped: "
                f"{git_error_text(push) or 'Could not push dashboard data to origin/main.'}"
            )
            return False

        print("Dashboard data refreshed and pushed.")
        return True
    finally:
        shutil.rmtree(clone_dir, ignore_errors=True)


def dirty_paths_excluding_output(output_path: Path) -> list[str]:
    rel_output = output_path.relative_to(BUNDLE_DIR).as_posix()
    result = run_git("status", "--porcelain", check=False)
    if result.returncode != 0:
        return ["<git-status-unavailable>"]

    dirty_paths: list[str] = []
    for line in result.stdout.splitlines():
        if len(line) < 4:
            continue
        candidate = line[3:].strip().replace("\\", "/")
        if candidate == rel_output:
            continue
        dirty_paths.append(candidate)
    return dirty_paths


def same_file_content(left: Path, right: Path) -> bool:
    if not left.exists() or not right.exists():
        return False
    return left.read_bytes() == right.read_bytes()


def push_dashboard(workbook_path: Path | None, workbook_url: str | None, output_path: Path, commit_message: str) -> bool:
    fd, temp_output_raw = tempfile.mkstemp(suffix=output_path.suffix or ".json")
    os.close(fd)
    temp_output = Path(temp_output_raw)
    try:
        refresh_dashboard_data(workbook=str(workbook_path) if workbook_path else None, workbook_url=workbook_url, output=temp_output)
        dashboard_changed = not same_file_content(temp_output, output_path)
        if dashboard_changed:
            write_local_output(temp_output, output_path)

        if not is_git_repo():
            print("Dashboard data refreshed locally. No Git repository detected, so nothing was pushed.")
            return False

        if not has_origin():
            print("Dashboard data refreshed locally. No origin remote is configured yet.")
            return False

        if not dashboard_changed:
            print("Dashboard data is already up to date.")
            return False

        try:
            rel_output = output_path.relative_to(BUNDLE_DIR)
        except ValueError:
            print("Dashboard data refreshed locally. The output file is outside the dashboard repo, so nothing was pushed.")
            return False

        return push_output_from_clean_clone(output_path, rel_output, commit_message)
    finally:
        temp_output.unlink(missing_ok=True)


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
