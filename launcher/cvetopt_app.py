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


def log_tail(path: Path, lines: int = 15) -> str:
    """Последние строки лога — чтобы показать причину сразу в окне."""
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return "(log not available)"
    tail = [ln for ln in text.splitlines() if ln.strip()][-lines:]
    return "\n".join(tail) if tail else "(log is empty)"


def is_server_up() -> bool:
    try:
        with urllib.request.urlopen(HEALTH_URL, timeout=2) as response:
            return int(response.status) == 200
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return False


def wait_for_server(root: Path, proc: subprocess.Popen | None = None) -> bool:
    for i in range(START_TIMEOUT_SEC):
        if is_server_up():
            append_log(root, f"server up after {i}s")
            return True
        if proc is not None and proc.poll() is not None:
            append_log(root, f"server process died after {i}s code={proc.returncode}")
            return False
        if i in (5, 15, 30, 60):
            append_log(root, f"waiting for server... {i}s")
        time.sleep(1)
    return False


def _uvicorn_cmd(root: Path) -> list[str] | None:
    venv_py = root / ".venv" / "Scripts" / "python.exe"
    if not venv_py.is_file():
        return None
    return [
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


def _do_update_pull(root: Path, log_f: object) -> None:
    """git pull + sync after admin update (exit code 42)."""
    import shutil

    def _w(msg: str) -> None:
        append_log(root, msg)
        try:
            log_f.write(msg + "\n")
            log_f.flush()
        except OSError:
            pass

    _w("update: git pull --ff-only")
    if (root / ".git").is_dir():
        subprocess.run(
            ["git", "pull", "--ff-only"],
            cwd=str(root),
            stdout=log_f,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW,
            check=False,
        )
    uv = shutil.which("uv")
    if uv:
        _w("update: uv sync")
        subprocess.run(
            [uv, "sync"],
            cwd=str(root),
            stdout=log_f,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW,
            check=False,
        )
    else:
        _w("update: uv not in PATH - sync skipped (using existing .venv)")


def start_server(root: Path) -> tuple[subprocess.Popen | None, object | None]:
    """
    Start server. Prefer cvetopt.bat (has exit-42 update loop).
    Fallback: .venv uvicorn with our own 42-restart watcher.
    """
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
    creation = CREATE_NO_WINDOW if sys.platform == "win32" else 0

    bat = root / "cvetopt.bat"
    if bat.is_file():
        # Bat owns the update/restart loop (exit code 42).
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
        append_log(root, f"bat pid={proc.pid}")
        return proc, log_f

    cmd = _uvicorn_cmd(root)
    if cmd is None:
        log_f.write("no cvetopt.bat and no .venv python\n")
        log_f.close()
        return None, None

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

    def _watch_exit42() -> None:
        nonlocal proc
        while True:
            code = proc.wait()
            append_log(root, f"uvicorn exited code={code}")
            if code != 42:
                return
            _do_update_pull(root, log_f)
            append_log(root, "restarting uvicorn after update")
            proc = subprocess.Popen(
                cmd,
                cwd=str(root),
                env=env,
                stdout=log_f,
                stderr=subprocess.STDOUT,
                creationflags=creation,
            )
            append_log(root, f"uvicorn pid={proc.pid}")

    import threading

    threading.Thread(target=_watch_exit42, daemon=True).start()
    return proc, log_f


def browser_candidates() -> list[Path]:
    """Prefer Edge (more stable --app process) over Chrome."""
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    ordered = [
        Path(pf) / r"Microsoft\Edge\Application\msedge.exe",
        Path(pf86) / r"Microsoft\Edge\Application\msedge.exe",
        Path(pf) / r"Google\Chrome\Application\chrome.exe",
        Path(pf86) / r"Google\Chrome\Application\chrome.exe",
    ]
    found: list[Path] = []
    seen: set[str] = set()
    for path in ordered:
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        if path.is_file():
            found.append(path)
    return found


def browser_pids_for_profile(profile_marker: str) -> list[int]:
    """PIDs whose command line mentions our Edge/Chrome profile folder."""
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
    except (OSError, subprocess.TimeoutExpired):
        return []
    pids: list[int] = []
    for line in (r.stdout or "").splitlines():
        line = line.strip()
        if line.isdigit():
            pids.append(int(line))
    return pids


def open_app_and_wait(root: Path, url: str) -> None:
    """
    Open Edge/Chrome --app, then block until the user finishes.

    Chrome often drops profile processes quickly (merges into main browser),
    so close-detection alone is unreliable. We:
    1) try to wait until profile processes disappear (after they appeared),
    2) always fall back to a MessageBox — OK stops the server.
    """
    profile = root / "data" / "edge-app-profile"
    profile.mkdir(parents=True, exist_ok=True)
    marker = "edge-app-profile"

    browsers = browser_candidates()
    append_log(root, f"browsers found: {[str(b) for b in browsers]}")
    if browsers:
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
    else:
        append_log(root, "no Edge/Chrome - start URL")
        subprocess.Popen(
            ["cmd", "/c", "start", "", url],
            creationflags=CREATE_NO_WINDOW,
        )

    # Give the window a few seconds to appear; optional close-detect.
    appeared_at: float | None = None
    idle_rounds = 0
    for tick in range(45):  # ~90 seconds max of polling before MessageBox
        time.sleep(2)
        pids = browser_pids_for_profile(marker)
        if pids:
            if appeared_at is None:
                appeared_at = time.time()
                append_log(root, f"app window processes: {pids}")
            idle_rounds = 0
            continue
        if appeared_at is not None and (time.time() - appeared_at) >= 10:
            idle_rounds += 1
            if idle_rounds >= 5:
                append_log(root, "app window closed (process poll)")
                return

    append_log(root, "MessageBox wait (OK = stop server)")
    message_box(
        "cvetopt is running.\n\n"
        "Work in the app window.\n"
        "When finished, click OK — the server will stop."
    )


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
        if not wait_for_server(root, owned_proc):
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
            server_log = root / "data" / "launcher-server.log"
            msg = (
                f"Server did not answer in {START_TIMEOUT_SEC}s "
                f"(server process exit={dead}).\n\n"
                f"Last lines of {server_log.name}:\n{log_tail(server_log)}\n\n"
                f"Full logs:\n{server_log}\n{root / 'data' / 'launcher.log'}"
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
