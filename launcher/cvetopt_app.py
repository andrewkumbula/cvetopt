"""
cvetopt Windows launcher (built to cvetopt.exe).

1) Start uvicorn from project .venv (hidden)
2) Wait until http://127.0.0.1:8000 answers
3) Open Edge/Chrome --app with a dedicated profile (so we can wait for close)
4) On window close - stop the server
"""
from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

APP_URL = "http://127.0.0.1:8000/"
HEALTH_URL = "http://127.0.0.1:8000/api/state"
START_TIMEOUT_SEC = 120
CREATE_NO_WINDOW = 0x08000000


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def message_box(text: str, *, title: str = "cvetopt", error: bool = False) -> None:
    if sys.platform != "win32":
        print(f"{title}: {text}", file=sys.stderr)
        return
    style = 0x10 if error else 0x30
    ctypes.windll.user32.MessageBoxW(0, text, title, style)


def is_server_up() -> bool:
    try:
        with urllib.request.urlopen(HEALTH_URL, timeout=2) as response:
            return int(response.status) == 200
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return False


def wait_for_server() -> bool:
    for _ in range(START_TIMEOUT_SEC):
        if is_server_up():
            return True
        time.sleep(1)
    return False


def start_server(root: Path) -> subprocess.Popen[bytes] | None:
    """Start uvicorn via .venv python; fall back to cvetopt.bat."""
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PLAYWRIGHT_BROWSERS_PATH"] = str(root / "ms-playwright")
    env["CVETOPT_HIDDEN"] = "1"
    env["CVETOPT_NO_BROWSER"] = "1"

    log_dir = root / "data"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "launcher-server.log"
    log_f = open(log_path, "w", encoding="utf-8", errors="replace")

    venv_py = root / ".venv" / "Scripts" / "python.exe"
    if venv_py.is_file():
        log_f.write(f"starting: {venv_py}\n")
        log_f.flush()
        return subprocess.Popen(
            [
                str(venv_py),
                "-u",
                "-m",
                "uvicorn",
                "cvetopt.app:app",
                "--host",
                "127.0.0.1",
                "--port",
                "8000",
                "--app-dir",
                "src",
                "--log-level",
                "info",
            ],
            cwd=root,
            env=env,
            stdout=log_f,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW,
        )

    bat = root / "cvetopt.bat"
    if bat.is_file():
        log_f.write(f"starting bat: {bat}\n")
        log_f.flush()
        return subprocess.Popen(
            ["cmd", "/c", str(bat)],
            cwd=root,
            env=env,
            stdout=log_f,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW,
        )

    log_f.write("no .venv python and no cvetopt.bat\n")
    log_f.close()
    return None


def browser_candidates() -> list[Path]:
    roots = [
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
    ]
    names = (
        r"Microsoft\Edge\Application\msedge.exe",
        r"Google\Chrome\Application\chrome.exe",
    )
    found: list[Path] = []
    seen: set[str] = set()
    for base in roots:
        for name in names:
            path = Path(base) / name
            key = str(path).lower()
            if key in seen:
                continue
            seen.add(key)
            if path.is_file():
                found.append(path)
    return found


def open_app_and_wait(root: Path, url: str) -> None:
    """
    Dedicated --user-data-dir so Edge/Chrome is a separate process.
    Without it, --app often returns immediately (shared browser process)
    and the launcher would kill the server right away.
    """
    profile = root / "data" / "edge-app-profile"
    profile.mkdir(parents=True, exist_ok=True)

    for browser in browser_candidates():
        subprocess.run(
            [
                str(browser),
                f"--user-data-dir={profile}",
                "--no-first-run",
                "--no-default-browser-check",
                f"--app={url}",
            ],
            check=False,
        )
        return

    subprocess.run(
        ["cmd", "/c", "start", "/wait", "", url],
        check=False,
        creationflags=CREATE_NO_WINDOW,
    )


def stop_server(root: Path, proc: subprocess.Popen[bytes] | None) -> None:
    if proc is not None and proc.poll() is None:
        try:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
        except OSError:
            pass

    stop_bat = root / "cvetopt-stop.bat"
    if stop_bat.is_file():
        env = os.environ.copy()
        env["CVETOPT_QUIET"] = "1"
        subprocess.run(
            ["cmd", "/c", str(stop_bat)],
            cwd=root,
            env=env,
            creationflags=CREATE_NO_WINDOW,
            check=False,
        )


def main() -> int:
    if sys.platform != "win32":
        message_box("cvetopt is Windows-only.", error=True)
        return 1

    root = project_root()
    os.chdir(root)

    if not (root / ".venv" / "Scripts" / "python.exe").is_file() and not (
        root / "cvetopt.bat"
    ).is_file():
        message_box(
            f"Missing .venv and cvetopt.bat in:\n{root}",
            error=True,
        )
        return 1

    owned_proc: subprocess.Popen[bytes] | None = None
    started_here = False

    if is_server_up():
        # Reuse existing server; do not kill it on exit (another session may own it).
        owned_proc = None
        started_here = False
    else:
        owned_proc = start_server(root)
        started_here = True
        if owned_proc is None:
            message_box("Could not start server (.venv / cvetopt.bat).", error=True)
            return 1
        if not wait_for_server():
            stop_server(root, owned_proc)
            log_hint = root / "data" / "launcher-server.log"
            message_box(
                f"Server did not answer in {START_TIMEOUT_SEC}s.\n\n"
                f"See log:\n{log_hint}",
                error=True,
            )
            return 1

    try:
        open_app_and_wait(root, APP_URL)
    finally:
        if started_here:
            stop_server(root, owned_proc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
