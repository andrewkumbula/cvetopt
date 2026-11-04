"""VPN для браузера скраперов: sing-box, который программа поднимает сама.

Включается строкой VPN_SINGBOX=<папка> в .env (обычно VPN_SINGBOX=vpn — папка
vpn внутри программы, она в .gitignore); в папке лежат sing-box(.exe)
и config.json (собирается на маке: vpnctl, render_client.py --mode server).
Без этой строки ничего не запускается — работает прежний PLAYWRIGHT_PROXY (Happ).

Один процесс на всех: первая задача запускает, последняя гасит. PID пишется в
data/state/sing-box.pid, чтобы после падения программы не остался висеть
старый sing-box, держащий порт.
"""

from __future__ import annotations

import atexit
import json
import os
import shutil
import socket
import subprocess
import threading
import time
import urllib.request
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from cvetopt.core.settings import EnvSettings

# По http: внутри туннеля это всё равно зашифровано, а фейковому sing-box в
# тестах не нужно уметь CONNECT.
IP_CHECK_URL = "http://api.ipify.org"
START_TIMEOUT_SEC = 15.0
TUNNEL_TIMEOUT_SEC = 30.0

_EXE_NAME = "sing-box.exe" if os.name == "nt" else "sing-box"

_lock = threading.Lock()
_users = 0
_proc: subprocess.Popen | None = None
_proxy: dict[str, str] | None = None


def configured(env: EnvSettings) -> bool:
    return bool((env.vpn_singbox or "").strip())


def _paths(env: EnvSettings) -> tuple[Path, Path]:
    folder = Path(env.vpn_singbox.strip())
    if not folder.is_absolute():
        folder = env.project_root / folder  # VPN_SINGBOX=vpn — папка vpn в программе
    exe = folder / _EXE_NAME
    config = folder / "config.json"
    if not exe.is_file():
        found = shutil.which("sing-box") if os.name != "nt" else None
        if not found:
            raise RuntimeError(f"VPN: нет {exe} (VPN_SINGBOX в .env)")
        exe = Path(found)
    if not config.is_file():
        raise RuntimeError(f"VPN: нет {config} — соберите его в vpnctl (--mode server)")
    return exe, config


def _listen_port(config: Path) -> int:
    try:
        data = json.loads(config.read_text(encoding="utf-8"))
        for inbound in data.get("inbounds", []):
            if inbound.get("type") in ("mixed", "http"):
                return int(inbound["listen_port"])
    except (OSError, ValueError, KeyError) as e:
        raise RuntimeError(f"VPN: не читается {config}: {e}") from None
    raise RuntimeError(f"VPN: в {config} нет входа mixed/http — это не конфиг --mode server")


def _state_dir(env: EnvSettings) -> Path:
    d = env.project_root / "data" / "state"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _port_busy(port: int) -> bool:
    try:
        socket.create_connection(("127.0.0.1", port), timeout=1).close()
        return True
    except OSError:
        return False


def _is_singbox(pid: int) -> bool:
    """Чтобы по устаревшему PID-файлу не убить чужой процесс с тем же номером."""
    try:
        if os.name == "nt":
            out = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            ).stdout
        else:
            out = subprocess.run(
                ["ps", "-p", str(pid), "-o", "command="],
                capture_output=True, text=True, timeout=10,
            ).stdout
    except (OSError, subprocess.SubprocessError):
        return False
    return "sing-box" in out.lower()


def _kill_leftover(pid_file: Path) -> None:
    try:
        pid = int(pid_file.read_text().strip())
    except (OSError, ValueError):
        return
    if _is_singbox(pid):
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        else:
            try:
                os.kill(pid, 9)
            except OSError:
                pass
        time.sleep(0.5)
    pid_file.unlink(missing_ok=True)


def _log_tail(log_file: Path, lines: int = 5) -> str:
    try:
        text = log_file.read_text(encoding="utf-8", errors="replace").strip()
    except OSError:
        return ""
    return " | ".join(text.splitlines()[-lines:])


def _tunnel_ip(proxy_server: str) -> str:
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": proxy_server, "https": proxy_server})
    )
    with opener.open(IP_CHECK_URL, timeout=10) as resp:
        return resp.read(64).decode("ascii", errors="replace").strip()


def _start(env: EnvSettings) -> list[str]:
    global _proc, _proxy
    exe, config = _paths(env)
    port = _listen_port(config)
    state = _state_dir(env)
    pid_file = state / "sing-box.pid"
    log_file = state / "sing-box.log"

    _kill_leftover(pid_file)
    if _port_busy(port):
        raise RuntimeError(
            f"VPN: порт {port} занят другой программой — "
            f"netstat -ano | findstr {port} покажет, какой"
        )

    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    with log_file.open("w", encoding="utf-8") as log:
        try:
            proc = subprocess.Popen(
                [str(exe), "run", "-c", str(config)],
                stdout=log, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                creationflags=flags,
            )
        except OSError as e:
            # Так выглядит и блокировка Defender/SmartScreen.
            raise RuntimeError(f"VPN: не удалось запустить {exe}: {e}") from None
    pid_file.write_text(str(proc.pid))
    _proc = proc
    proxy_server = f"http://127.0.0.1:{port}"

    deadline = time.monotonic() + START_TIMEOUT_SEC
    while not _port_busy(port):
        if proc.poll() is not None:
            _stop(env)
            raise RuntimeError(f"VPN: sing-box не запустился: {_log_tail(log_file)}")
        if time.monotonic() > deadline:
            _stop(env)
            raise RuntimeError(f"VPN: sing-box не открыл порт {port} за {START_TIMEOUT_SEC:.0f} с")
        time.sleep(0.2)

    # Порт открыт ещё не значит «туннель работает»: проверяем настоящим запросом.
    deadline = time.monotonic() + TUNNEL_TIMEOUT_SEC
    last_error = ""
    while True:
        try:
            ip = _tunnel_ip(proxy_server)
            break
        except OSError as e:
            last_error = str(e)
        if time.monotonic() > deadline:
            _stop(env)
            raise RuntimeError(
                f"VPN: туннель не отвечает ({last_error}). Узлы недоступны или доступ "
                f"cvetopt-server отозван — проверьте с мака: ./vpn статус"
            )
        time.sleep(2)

    _proxy = {"server": proxy_server}
    return [f"VPN: sing-box запущен, порт {port}", f"VPN: туннель работает, внешний IP {ip}"]


def _stop(env: EnvSettings | None = None) -> None:
    global _proc, _proxy
    proc, _proc, _proxy = _proc, None, None
    if proc is not None and proc.poll() is None:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=5)
    if env is not None:
        (_state_dir(env) / "sing-box.pid").unlink(missing_ok=True)


def acquire(env: EnvSettings) -> tuple[dict[str, str], list[str]]:
    """Поднять туннель (или подключиться к уже поднятому). Парный release() обязателен."""
    global _users
    with _lock:
        if _proc is not None and _proc.poll() is not None:
            _stop(env)  # умер сам, пока им пользовались — перезапустим
        notes = _start(env) if _proc is None else ["VPN: sing-box уже запущен другой задачей"]
        _users += 1
        assert _proxy is not None
        return dict(_proxy), notes


def release(env: EnvSettings) -> None:
    global _users
    with _lock:
        _users = max(0, _users - 1)
        if _users == 0:
            _stop(env)


@contextmanager
def session(env: EnvSettings) -> Iterator[tuple[dict[str, str], list[str]]]:
    proxy, notes = acquire(env)
    try:
        yield proxy, notes
    finally:
        release(env)


atexit.register(_stop)
