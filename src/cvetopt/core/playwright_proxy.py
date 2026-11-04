"""Прокси для Playwright Chromium.

Два источника:
  VPN_SINGBOX — программа сама поднимает sing-box на время задачи (vpn_tunnel);
  PLAYWRIGHT_PROXY — уже работающий прокси, обычно локальный HTTP-порт Happ.
Системный прокси Windows / TUN-режим Happ Playwright НЕ использует.
"""

from __future__ import annotations

import asyncio
import os
from collections.abc import AsyncIterator, Awaitable, Callable
from contextlib import asynccontextmanager
from urllib.parse import unquote, urlparse

from cvetopt.core import vpn_tunnel
from cvetopt.core.settings import EnvSettings


def playwright_proxy_dict(env: EnvSettings | None = None) -> dict[str, str] | None:
    """
    Возвращает dict для chromium.launch(proxy=...) или None.

    Поддерживает:
      PLAYWRIGHT_PROXY=http://127.0.0.1:10809
      PLAYWRIGHT_PROXY=socks5://127.0.0.1:10808
      PLAYWRIGHT_PROXY=http://user:pass@127.0.0.1:10809
      + PLAYWRIGHT_PROXY_USERNAME / PLAYWRIGHT_PROXY_PASSWORD
    """
    server = ""
    user = ""
    pwd = ""
    if env is not None:
        server = (env.playwright_proxy or "").strip()
        user = (env.playwright_proxy_username or "").strip()
        pwd = (env.playwright_proxy_password or "").strip()
    if not server:
        server = (
            os.getenv("PLAYWRIGHT_PROXY", "").strip()
            or os.getenv("HTTPS_PROXY", "").strip()
            or os.getenv("HTTP_PROXY", "").strip()
        )
    if not server:
        return None
    if not user:
        user = os.getenv("PLAYWRIGHT_PROXY_USERNAME", "").strip()
    if not pwd:
        pwd = os.getenv("PLAYWRIGHT_PROXY_PASSWORD", "").strip()

    server, user, pwd = _normalize_proxy(server, user, pwd)
    proxy: dict[str, str] = {"server": server}
    if user:
        proxy["username"] = user
        proxy["password"] = pwd
    return proxy


def _normalize_proxy(server: str, user: str, pwd: str) -> tuple[str, str, str]:
    """Выделить user:pass из URL и добавить схему, если её нет."""
    raw = server.strip()
    if "://" not in raw:
        # Без схемы Chromium часто ломается — по умолчанию HTTP (порт Happ).
        raw = f"http://{raw}"

    parsed = urlparse(raw)
    if parsed.username and not user:
        user = unquote(parsed.username)
    if parsed.password and not pwd:
        pwd = unquote(parsed.password)

    host = parsed.hostname or ""
    if not host:
        return raw, user, pwd

    port = f":{parsed.port}" if parsed.port else ""
    scheme = (parsed.scheme or "http").lower()
    # Chromium надёжнее с http:// локального Happ, чем с socks5://
    # (socks5 с auth в Chromium не работает вовсе).
    clean = f"{scheme}://{host}{port}"
    return clean, user, pwd


@asynccontextmanager
async def browser_proxy(
    env: EnvSettings,
    log: Callable[[str], Awaitable[None]],
    *,
    needed: bool = True,
) -> AsyncIterator[dict[str, str] | None]:
    """Прокси для chromium.launch(proxy=...) на время блока, или None.

    С VPN_SINGBOX поднимает sing-box и гасит его на выходе из блока."""
    if not needed:
        yield None
        return
    if not vpn_tunnel.configured(env):
        yield playwright_proxy_dict(env)
        return
    await log("VPN: запускаю sing-box…")
    proxy, notes = await asyncio.to_thread(vpn_tunnel.acquire, env)
    try:
        for note in notes:
            await log(note)
        yield proxy
    finally:
        await asyncio.to_thread(vpn_tunnel.release, env)
