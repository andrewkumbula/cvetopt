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
import traceback
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


def append_log(root: Path, text: str) -> None:
    try:
        log_dir = root / "data"
        log_dir.mkdir(parents=True, exist_ok=True)
        with open(log_dir / "launcher.log", "a", encoding="utf-8", errors="replace") as f:
            f.write(text)
            if not text.endswith("\n"):
                f.write("\n")
    except OSError:
        pass


def message_box(text: str, *, title: str = "cvetopt", error: bool = False) -> None:
    if sys.platform != "win32":
        print(f"{title}: {text}", file=sys.stderr)
        return
    style = 0x10 if error else 0x30
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, style)
    except Exception:
        pass


def is_server_up() -> bool:
    try:
        with urllib.request.urlopen(HEALTH_URL, timeout=2) as response:
            return int(response.status) == 200
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return False


def wait_for_server(root: Path) -> bool:
    for i in range(START_TIMEOUT_SEC):
        if is_server_up():
            append_log(root, f"server up after {i}s")
            return True
        if i in (5, 15, 30, 60):
            append_log(root, f"waiting for server... {i}s")
        time.sleep(1)
    return False


def start_server(root: Path) -> tuple[subprocess.Popen | None, object | None]:
    """Start uvicorn via .venv python; fall back to cvetopt.bat. Returns (proc, log_handle)."""
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
    creation = CREATE_NO_WINDOW if sys.platform == "win32" else 0

    if venv_py.is_file():
        cmd = [
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
        ]
        append_log(root, f"start cmd: {' '.join(cmd)}")
        log_f.write(f"cmd: {' '.join(cmd)}\n")
        log_f.flush()
        proc = subprocess.Popen(
            cmd,
            cwd=str(root),
            env=env,
            stdout=log_f,
            stderr=subprocess.STDOUT,
            creationflags=creation,
        )
        append_log(root, f"uvicorn pid={proc.pid}")
        return proc, log_f

    bat = root / "cvetopt.bat"
    if bat.is_file():
        append_log(root, f"start bat: {bat}")
        log_f.write(f"bat: {bat}\n")
        log_f.flush()
        proc = subprocess.Popen(
            ["cmd", "/c", str(bat)],
            cwd=str(root),
            env=env,
            stdout=log_f,
            stderr=subprocess.STDOUT,
            creationflags=creation,
        )
        return proc, log_f

    log_f.write("no .venv python and no cvetopt.bat\n")
    log_f.close()
    return None, None


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


def browser_pids_for_profile(profile_marker: str) -> list[int]:
    """PIDs whose command line mentions our Edge/Chrome profile folder."""
    # Escape single quotes for PowerShell string
    marker = profile_marker.replace("'", "''")
    ps = (
        "Get-CimInstance Win32_Process | "
        f"Where-Object {{ $_.CommandLine -and $_.CommandLine -like '*{marker}*' }} | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=45,
            creationflags=CREATE_NO_WINDOW,
        )
    except (OSError, subprocess.TimeoutExpired) as e:
        return []
    pids: list[int] = []
    for line in (r.stdout or "").splitlines():
        line = line.strip()
        if line.isdigit():
            pids.append(int(line))
    return pids


def open_app_and_wait(root: Path, url: str) -> None:
    """
    Start Edge/Chrome --app with a dedicated profile, then wait until those
    processes exit. Edge's stub process often exits immediately; we must poll
    by command-line marker or the launcher kills the server right away.
    """
    profile = root / "data" / "edge-app-profile"
    profile.mkdir(parents=True, exist_ok=True)
    marker = "edge-app-profile"

    browsers = browser_candidates()
    append_log(root, f"browsers found: {[str(b) for b in browsers]}")
    if not browsers:
        append_log(root, "no Edge/Chrome - using start /wait")
        subprocess.run(
            ["cmd", "/c", "start", "/wait", "", url],
            check=False,
            creationflags=CREATE_NO_WINDOW,
        )
        return

    browser = browsers[0]
    cmd = [
        str(browser),
        f"--user-data-dir={profile}",
        "--no-first-run",
        "--no-default-browser-check",
        f"--app={url}",
    ]
    append_log(root, f"open app: {' '.join(cmd)}")
    subprocess.Popen(cmd, creationflags=CREATE_NO_WINDOW)

    # Wait until browser processes with our profile appear, then until they leave.
    appeared = False
    idle_rounds = 0
    for tick in range(3600):  # up to ~1 hour (2s * 1800 would be better; use 2s sleep)
        time.sleep(2)
        pids = browser_pids_for_profile(marker)
        if pids:
            if not appeared:
                append_log(root, f"app window processes: {pids}")
            appeared = True
            idle_rounds = 0
            continue
        if appeared:
            idle_rounds += 1
            # Require a couple empty polls so we don't race process restarts.
            if idle_rounds >= 2:
                append_log(root, "app window closed")
                return
        elif tick >= 15:
            # 30s and still no profile process — fall through to keep server? Better warn.
            append_log(root, "browser profile process not seen - keep waiting anyway")
            appeared = True  # avoid infinite "not seen"; wait for close or user kill
    append_log(root, "app wait timed out")


def stop_server(root: Path, proc: subprocess.Popen | None) -> None:
    append_log(root, "stopping server...")
    if proc is not None and proc.poll() is None:
        try:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
        except OSError as e:
            append_log(root, f"terminate failed: {e}")

    stop_bat = root / "cvetopt-stop.bat"
    if stop_bat.is_file():
        env = os.environ.copy()
        env["CVETOPT_QUIET"] = "1"
        subprocess.run(
            ["cmd", "/c", str(stop_bat)],
            cwd=str(root),
            env=env,
            creationflags=CREATE_NO_WINDOW,
            check=False,
        )
    append_log(root, "stop done")


def run() -> int:
    root = project_root()
    os.chdir(root)
    append_log(root, f"=== launcher start root={root} frozen={getattr(sys, 'frozen', False)}")

    venv_py = root / ".venv" / "Scripts" / "python.exe"
    bat = root / "cvetopt.bat"
    if not venv_py.is_file() and not bat.is_file():
        message_box(f"Missing .venv and cvetopt.bat in:\n{root}", error=True)
        return 1

    owned_proc: subprocess.Popen | None = None
    log_handle = None
    started_here = False

    if is_server_up():
        append_log(root, "server already up - reuse")
        started_here = False
    else:
        owned_proc, log_handle = start_server(root)
        started_here = True
        if owned_proc is None:
            message_box("Could not start server (.venv / cvetopt.bat).", error=True)
            return 1
        if not wait_for_server(root):
            # Capture early crash output
            time.sleep(0.5)
            dead = owned_proc.poll()
            append_log(root, f"server timeout; proc exit={dead}")
            stop_server(root, owned_proc)
            if log_handle is not None:
                try:
                    log_handle.close()
                except OSError:
                    pass
            msg = (
                f"Server did not answer in {START_TIMEOUT_SEC}s.\n\n"
                f"Log:\n{root / 'data' / 'launcher-server.log'}\n"
                f"{root / 'data' / 'launcher.log'}"
            )
            message_box(msg, error=True)
            return 1

    try:
        open_app_and_wait(root, APP_URL)
    finally:
        if started_here:
            stop_server(root, owned_proc)
        if log_handle is not None:
            try:
                log_handle.close()
            except OSError:
                pass
        append_log(root, "=== launcher exit")

    return 0


def main() -> int:
    if sys.platform != "win32":
        message_box("cvetopt is Windows-only.", error=True)
        return 1
    try:
        return run()
    except Exception as e:
        root = project_root()
        tb = traceback.format_exc()
        append_log(root, tb)
        message_box(f"Launcher crash:\n{e}\n\nSee data\\launcher.log", error=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
