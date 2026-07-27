"""Прокси для Playwright Chromium.

Системный прокси Windows / TUN-режим Happ Playwright НЕ использует —
нужен явный PLAYWRIGHT_PROXY в .env (обычно локальный HTTP-порт Happ).
"""

from __future__ import annotations

import os
from typing import Any
from urllib.parse import unquote, urlparse

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


def apply_playwright_proxy(
    launch_kwargs: dict[str, Any],
    env: EnvSettings | None = None,
) -> str | None:
    """Добавляет proxy в launch_kwargs. Возвращает server для лога или None."""
    proxy = playwright_proxy_dict(env)
    if not proxy:
        return None
    launch_kwargs["proxy"] = proxy
    return proxy["server"]
