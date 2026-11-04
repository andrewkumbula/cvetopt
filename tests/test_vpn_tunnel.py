"""sing-box, который программа поднимает сама (core/vpn_tunnel).

Вместо настоящего sing-box — маленький скрипт с тем же именем: читает
config.json, слушает порт и отвечает на любой запрос «внешним IP». Поведение
задаётся полем "fake" в конфиге: crash — падает сразу, dead — порт открыт,
а туннель не работает (502).
"""

import asyncio
import json
import os
import socket
import subprocess
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from cvetopt.core import vpn_tunnel
from cvetopt.core.playwright_proxy import browser_proxy

FAKE_SINGBOX = """#!/usr/bin/env python3
import json, sys
from http.server import BaseHTTPRequestHandler, HTTPServer

cfg = json.load(open(sys.argv[sys.argv.index("-c") + 1]))
mode = cfg.get("fake", "ok")
if mode == "crash":
    print("FATAL decode config: bad uuid")
    sys.exit(1)

class H(BaseHTTPRequestHandler):
    def do_GET(self):
        code, body = (502, b"") if mode == "dead" else (200, b"203.0.113.7")
        self.send_response(code)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
    def log_message(self, *a):
        pass

print("started", flush=True)
HTTPServer(("127.0.0.1", cfg["inbounds"][0]["listen_port"]), H).serve_forever()
"""


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    # Зомби (умер, но ещё не прибран) тоже считаем мёртвым.
    out = subprocess.run(["ps", "-p", str(pid), "-o", "stat="], capture_output=True, text=True).stdout
    return bool(out.strip()) and not out.strip().startswith("Z")


@unittest.skipIf(os.name == "nt", "фейковый sing-box — shell-скрипт")
class VpnTunnelTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.vpn_dir = self.root / "vpn"
        self.vpn_dir.mkdir()
        exe = self.vpn_dir / "sing-box"
        exe.write_text(FAKE_SINGBOX)
        exe.chmod(0o755)
        self.port = _free_port()
        self.write_config()
        self.env = SimpleNamespace(
            vpn_singbox=str(self.vpn_dir),
            project_root=self.root,
            playwright_proxy="",
            playwright_proxy_username="",
            playwright_proxy_password="",
        )
        self.pid_file = self.root / "data" / "state" / "sing-box.pid"
        for name, value in (("TUNNEL_TIMEOUT_SEC", 1.0), ("START_TIMEOUT_SEC", 5.0)):
            patcher = mock.patch.object(vpn_tunnel, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        patcher = mock.patch.object(vpn_tunnel, "IP_CHECK_URL", "http://ip.example/")
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self._reset_module)

    def _reset_module(self):
        vpn_tunnel._stop(self.env)
        vpn_tunnel._users = 0

    def write_config(self, fake: str = "ok"):
        cfg = {
            "fake": fake,
            "inbounds": [{"type": "mixed", "listen": "127.0.0.1", "listen_port": self.port}],
        }
        (self.vpn_dir / "config.json").write_text(json.dumps(cfg))

    def test_start_check_stop(self):
        with vpn_tunnel.session(self.env) as (proxy, notes):
            self.assertEqual(proxy, {"server": f"http://127.0.0.1:{self.port}"})
            self.assertIn("203.0.113.7", notes[-1])
            pid = int(self.pid_file.read_text())
            self.assertTrue(_alive(pid))
        self.assertFalse(_alive(pid))
        self.assertFalse(self.pid_file.exists())
        self.assertFalse(vpn_tunnel._port_busy(self.port))

    def test_relative_folder_is_inside_program(self):
        self.env.vpn_singbox = "vpn"
        with vpn_tunnel.session(self.env) as (proxy, _):
            self.assertEqual(proxy, {"server": f"http://127.0.0.1:{self.port}"})

    def test_two_jobs_share_one_process_last_one_stops_it(self):
        with vpn_tunnel.session(self.env):
            pid = int(self.pid_file.read_text())
            with vpn_tunnel.session(self.env) as (_, notes):
                self.assertIn("уже запущен", notes[0])
                self.assertEqual(int(self.pid_file.read_text()), pid)
            self.assertTrue(_alive(pid), "первая задача ещё работает — VPN гасить рано")
        self.assertFalse(_alive(pid))

    def test_crash_reports_log_tail(self):
        self.write_config("crash")
        with self.assertRaises(RuntimeError) as cm:
            vpn_tunnel.acquire(self.env)
        self.assertIn("не запустился", str(cm.exception))
        self.assertIn("bad uuid", str(cm.exception))
        self.assertEqual(vpn_tunnel._users, 0)
        self.assertFalse(self.pid_file.exists())

    def test_dead_tunnel_is_an_error_and_process_is_stopped(self):
        self.write_config("dead")
        with self.assertRaises(RuntimeError) as cm:
            vpn_tunnel.acquire(self.env)
        self.assertIn("туннель не отвечает", str(cm.exception))
        self.assertIsNone(vpn_tunnel._proc)
        self.assertFalse(vpn_tunnel._port_busy(self.port))

    def test_leftover_singbox_from_crashed_program_is_killed(self):
        old = subprocess.Popen(
            [str(self.vpn_dir / "sing-box"), "run", "-c", str(self.vpn_dir / "config.json")],
            stdout=subprocess.DEVNULL,
        )
        self.addCleanup(old.kill)
        deadline = time.monotonic() + 5
        while not vpn_tunnel._port_busy(self.port) and time.monotonic() < deadline:
            time.sleep(0.1)
        self.pid_file.parent.mkdir(parents=True, exist_ok=True)
        self.pid_file.write_text(str(old.pid))

        with vpn_tunnel.session(self.env):
            self.assertIsNotNone(old.poll(), "старый sing-box должен быть убит")
            self.assertNotEqual(int(self.pid_file.read_text()), old.pid)

    def test_port_taken_by_another_program(self):
        with socket.socket() as s:
            s.bind(("127.0.0.1", self.port))
            s.listen()
            with self.assertRaises(RuntimeError) as cm:
                vpn_tunnel.acquire(self.env)
        self.assertIn(f"порт {self.port} занят", str(cm.exception))

    def test_missing_config(self):
        (self.vpn_dir / "config.json").unlink()
        with self.assertRaises(RuntimeError) as cm:
            vpn_tunnel.acquire(self.env)
        self.assertIn("config.json", str(cm.exception))


class BrowserProxyTests(unittest.TestCase):
    """Без VPN_SINGBOX всё как раньше: PLAYWRIGHT_PROXY, sing-box не трогаем."""

    def _run(self, env, needed=True):
        log: list[str] = []

        async def lg(msg: str) -> None:
            log.append(msg)

        async def go():
            async with browser_proxy(env, lg, needed=needed) as proxy:
                return proxy

        with mock.patch.object(vpn_tunnel, "acquire", side_effect=AssertionError("не должен")):
            return asyncio.run(go()), log

    def _env(self, **kw):
        base = dict(vpn_singbox="", playwright_proxy="", playwright_proxy_username="",
                    playwright_proxy_password="")
        base.update(kw)
        return SimpleNamespace(**base)

    def test_happ_proxy_when_singbox_not_configured(self):
        proxy, log = self._run(self._env(playwright_proxy="http://127.0.0.1:10809"))
        self.assertEqual(proxy, {"server": "http://127.0.0.1:10809"})
        self.assertEqual(log, [])

    def test_not_needed_means_no_vpn_even_if_configured(self):
        proxy, _ = self._run(self._env(vpn_singbox="/some/dir"), needed=False)
        self.assertIsNone(proxy)


if __name__ == "__main__":
    unittest.main()
