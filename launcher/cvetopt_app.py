"""
cvetopt — лаунчер Windows-приложения (собирается в cvetopt.exe).

Стартует сервер в фоне, открывает окно --app, при закрытии окна останавливает сервер.
Только стандартная библиотека — для компактного .exe через PyInstaller.
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
START_TIMEOUT_SEC = 90
CREATE_NO_WINDOW = 0x08000000


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def message_box(text: str, *, title: str = "cvetopt", error: bool = False) -> None:
    if sys.platform != "win32":
        print(f"{title}: {text}", file=sys.stderr)
        return
    style = 0x10 if error else 0x30  # MB_ICONERROR / MB_ICONWARNING
    ctypes.windll.user32.MessageBoxW(0, text, title, style)


def is_server_up() -> bool:
    try:
        with urllib.request.urlopen(HEALTH_URL, timeout=2) as response:
            return response.status == 200
    except (urllib.error.URLError, TimeoutError, OSError):
        return False


def start_server(root: Path) -> None:
    bat = root / "cvetopt.bat"
    env = os.environ.copy()
    env["CVETOPT_HIDDEN"] = "1"
    env["CVETOPT_NO_BROWSER"] = "1"
    subprocess.Popen(
        ["cmd", "/c", str(bat)],
        cwd=root,
        env=env,
        creationflags=CREATE_NO_WINDOW,
    )


def wait_for_server() -> bool:
    for _ in range(START_TIMEOUT_SEC):
        if is_server_up():
            return True
        time.sleep(1)
    return False


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
    for root in roots:
        for name in names:
            path = Path(root) / name
            key = str(path).lower()
            if key in seen:
                continue
            seen.add(key)
            if path.is_file():
                found.append(path)
    return found


def open_app_and_wait(url: str) -> None:
    arg = f"--app={url}"
    for browser in browser_candidates():
        subprocess.run([str(browser), arg], check=False)
        return

    # Запасной путь: браузер по умолчанию, ждём закрытия окна.
    subprocess.run(
        ["cmd", "/c", "start", "/wait", "", url],
        check=False,
        creationflags=CREATE_NO_WINDOW,
    )


def stop_server(root: Path) -> None:
    stop_bat = root / "cvetopt-stop.bat"
    if not stop_bat.is_file():
        return
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
        message_box("cvetopt доступен только на Windows.", error=True)
        return 1

    root = project_root()
    bat = root / "cvetopt.bat"
    if not bat.is_file():
        message_box(
            f"Не найден cvetopt.bat рядом с программой:\n{bat}",
            error=True,
        )
        return 1

    os.chdir(root)

    if not is_server_up():
        start_server(root)
        if not wait_for_server():
            message_box(
                f"Сервер не ответил за {START_TIMEOUT_SEC} с.\n\n"
                f"Проверьте вручную:\n{bat}",
                error=True,
            )
            return 1

    try:
        open_app_and_wait(APP_URL)
    finally:
        stop_server(root)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
