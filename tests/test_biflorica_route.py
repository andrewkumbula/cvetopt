"""Biflorica: .ru не отвечает → .com через VPN → .ru через VPN (настоящий браузер, сайты локальные).

Адреса ec.term.biflorica.ru/.com направляем в браузере на локальные порты: «живой» порт
отдаёт страницу входа, «мёртвый» закрыт. «VPN» — локальный прокси, который отвечает на всё.
"""

import asyncio
import socket
import threading
import unittest
from contextlib import asynccontextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from unittest import mock

from playwright.async_api import async_playwright

from cvetopt.core.settings import BifloricaPortalConfig
from cvetopt.scrapers import biflorica_route
from cvetopt.scrapers.biflorica_route import open_biflorica, routes

RU = "ec.term.biflorica.ru"
COM = "ec.term.biflorica.com"


class _Page(BaseHTTPRequestHandler):
    def do_GET(self):
        body = f"<title>Login</title><p>{self.headers.get('Host')}</p>".encode()
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *_args):
        pass


def _serve():
    server = ThreadingHTTPServer(("127.0.0.1", 0), _Page)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server


def _dead_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class RoutesOrderTests(unittest.TestCase):
    def test_order(self):
        portal = BifloricaPortalConfig(login_url=f"https://{RU}/marketPlaceNew/balance")
        self.assertEqual(routes(portal), [(RU, False), (COM, True), (RU, True)])

    def test_order_when_vpn_is_on(self):
        portal = BifloricaPortalConfig(login_url=f"https://{RU}/x", use_proxy=True)
        self.assertEqual(routes(portal), [(RU, True), (COM, True), (RU, False)])


class OpenBifloricaTests(unittest.TestCase):
    def setUp(self):
        self.site = _serve()
        self.vpn = _serve()
        self.addCleanup(self.site.shutdown)
        self.addCleanup(self.vpn.shutdown)
        self.dead = _dead_port()
        self.portal = BifloricaPortalConfig(
            login_url=f"http://{RU}/marketPlaceNew/balance",
            orders_url=f"http://{RU}/marketPlaceNew/default/orders",
        )

    def run_open(self, alive: set[str], vpn: bool):
        rules = ", ".join(
            f"MAP {h} 127.0.0.1:{self.site.server_port if h in alive else self.dead}" for h in (RU, COM)
        )
        proxy = {"server": f"http://127.0.0.1:{self.vpn.server_port}"} if vpn else None

        @asynccontextmanager
        async def fake_proxy(_env, _log, *, needed=True):
            yield proxy if needed else None

        lines = []

        async def log(msg):
            lines.append(msg)

        async def go():
            async with async_playwright() as p, open_biflorica(
                mock.Mock(),
                self.portal,
                p,
                launch_kwargs={"headless": True, "args": [f"--host-resolver-rules={rules}"]},
                context_opts={},
                navigation_timeout_ms=10_000,
                log=log,
            ) as (page, portal):
                return await page.content(), portal

        with mock.patch.object(biflorica_route, "browser_proxy", fake_proxy):
            content, portal = asyncio.run(go())
        return content, portal, lines

    def test_ru_alive_stays_on_ru(self):
        content, portal, lines = self.run_open({RU}, vpn=False)
        self.assertIn(RU, content)
        self.assertEqual(portal.orders_url, f"http://{RU}/marketPlaceNew/default/orders")
        self.assertNotIn("не отвечает", "\n".join(lines))

    def test_ru_dead_goes_to_com_through_vpn(self):
        content, portal, lines = self.run_open(set(), vpn=True)
        self.assertIn(COM, content)  # .com — только через VPN
        self.assertEqual(portal.orders_url, f"http://{COM}/marketPlaceNew/default/orders")
        self.assertIn(f"Biflorica: {RU} напрямую не отвечает", lines)
        self.assertIn(f"Biflorica: открываю {COM} через VPN", lines)
        self.assertNotIn(f"Biflorica: открываю {COM} напрямую", lines)

    def test_com_alive_but_no_vpn_is_not_used_directly(self):
        with self.assertRaisesRegex(Exception, "ERR_"):
            self.run_open({COM}, vpn=False)

    def test_nothing_works_raises_site_error(self):
        with self.assertRaisesRegex(Exception, "ERR_"):
            self.run_open(set(), vpn=False)


if __name__ == "__main__":
    unittest.main()
