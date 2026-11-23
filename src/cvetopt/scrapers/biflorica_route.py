"""
Как открыть Biflorica: .ru напрямую → .com через VPN → .ru через VPN (.com напрямую не пробуем).

Переключаемся только если не открылась первая страница (до входа), поэтому повторный
заход ничего не скачивает дважды. 23.11.2026 ec.term.biflorica.ru не отвечал, а .com работал.
"""

from __future__ import annotations

from collections.abc import AsyncIterator, Awaitable, Callable
from contextlib import AsyncExitStack, asynccontextmanager
from urllib.parse import urlparse

from playwright.async_api import Error as PlaywrightError
from playwright.async_api import Page, Playwright

from cvetopt.core.playwright_proxy import browser_proxy
from cvetopt.core.settings import BifloricaPortalConfig, EnvSettings

LogFn = Callable[[str], Awaitable[None]]

# Сайт не ответил — пробуем следующий путь. Остальные ошибки (страница открылась,
# но что-то не так) пробрасываем как есть.
_UNREACHABLE = (
    "err_connection",
    "err_timed_out",
    "err_name_not_resolved",
    "err_address_unreachable",
    "err_internet_disconnected",
    "err_ssl",
    "err_proxy",
    "err_tunnel",
    "err_empty_response",
    "timeout",
)


def is_unreachable(error: BaseException) -> bool:
    text = str(error).casefold()
    return any(marker in text for marker in _UNREACHABLE)


def with_host(url: str, host: str) -> str:
    parsed = urlparse(url)
    return parsed._replace(netloc=host).geturl() if parsed.hostname else url


def mirror_host(host: str) -> str | None:
    """ec.term.biflorica.ru ↔ ec.term.biflorica.com"""
    if host.endswith(".ru"):
        return host[:-3] + ".com"
    if host.endswith(".com"):
        return host[:-4] + ".ru"
    return None


def routes(portal: BifloricaPortalConfig) -> list[tuple[str, bool]]:
    """(хост, через VPN) в порядке попыток. .com — только через VPN."""
    host = urlparse(portal.login_url).hostname or ""
    hosts = [host] + ([m] if (m := mirror_host(host)) else [])
    vpn_order = [True, False] if portal.use_proxy else [False, True]
    result = []
    for via_vpn in vpn_order:
        for h in hosts:
            route = (h, via_vpn or h.endswith(".com"))
            if route not in result:
                result.append(route)
    return result


def portal_for_host(portal: BifloricaPortalConfig, host: str) -> BifloricaPortalConfig:
    return portal.model_copy(
        update={
            "login_url": with_host(portal.login_url, host),
            "orders_url": with_host(portal.orders_url, host),
        }
    )


def _route_name(host: str, via_vpn: bool) -> str:
    return f"{host} {'через VPN' if via_vpn else 'напрямую'}"


@asynccontextmanager
async def open_biflorica(
    env: EnvSettings,
    portal: BifloricaPortalConfig,
    pw: Playwright,
    *,
    launch_kwargs: dict,
    context_opts: dict,
    navigation_timeout_ms: int,
    log: LogFn,
) -> AsyncIterator[tuple[Page, BifloricaPortalConfig]]:
    """Открывает страницу входа первым живым путём. Отдаёт страницу и адреса этого зеркала."""
    last_error: BaseException | None = None
    for host, via_vpn in routes(portal):
        route_portal = portal_for_host(portal, host)
        async with AsyncExitStack() as stack:
            try:
                proxy = await stack.enter_async_context(browser_proxy(env, log, needed=via_vpn))
            except RuntimeError as e:
                await log(f"Biflorica: VPN не включился — {e}")
                last_error = e
                continue
            if via_vpn and proxy is None:
                await log("Biflorica: VPN не настроен — пропускаю путь через VPN")
                continue
            kwargs = dict(launch_kwargs)
            if proxy:
                kwargs["proxy"] = proxy
            await log(f"Biflorica: открываю {_route_name(host, via_vpn)}")
            browser = await pw.chromium.launch(**kwargs)
            stack.push_async_callback(browser.close)
            context = await browser.new_context(**context_opts)
            stack.push_async_callback(context.close)
            context.set_default_navigation_timeout(navigation_timeout_ms)
            page = await context.new_page()
            await log(f"Открываю {route_portal.login_url}")
            try:
                await page.goto(route_portal.login_url, wait_until="domcontentloaded")
            except PlaywrightError as e:
                if not is_unreachable(e):
                    raise
                await log(f"Biflorica: {_route_name(host, via_vpn)} не отвечает")
                last_error = e
                continue
            yield page, route_portal
            return
    if last_error is not None:
        raise last_error
    raise RuntimeError("Biflorica: не осталось путей для подключения")
