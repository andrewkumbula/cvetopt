from __future__ import annotations

import asyncio
import getpass
import queue
import random
import re
import subprocess
import sys
from collections.abc import Awaitable, Callable
from datetime import date, timedelta
from pathlib import Path

from dateutil import parser as date_parser
from loguru import logger
from playwright.async_api import Browser, Download, Locator, Page, async_playwright

from cvetopt.core.job_manager import job_log, job_manager, raise_if_cancelled
from cvetopt.core.models import Order
from cvetopt.core.registry import DownloadRegistry
from cvetopt.core.runtime_settings import (
    archive_biflorica_download_dir,
    biflorica_download_filename,
    load_runtime_settings,
    resolve_biflorica_archive_dir,
    resolve_biflorica_download_dir,
)
from cvetopt.core.settings import (
    AppYamlConfig,
    BifloricaPortalConfig,
    EnvSettings,
    SelectionConfig,
    SelectionOverride,
    _resolve_selection,
    merged_playwright,
)


LogFn = Callable[[str], Awaitable[None]]


def _effective_selection(
    yaml_cfg: AppYamlConfig,
    override: SelectionOverride | None = None,
) -> SelectionConfig:
    """Окно для джоба заказов: переопределение portals.biflorica.selection, иначе глобальная selection."""
    base = _resolve_selection(yaml_cfg.selection, yaml_cfg.portals.biflorica.selection)
    return _resolve_selection(base, override)


def _flight_window_for_ui(
    today: date,
    yaml_cfg: AppYamlConfig,
    override: SelectionOverride | None = None,
) -> tuple[date, date]:
    sel = _effective_selection(yaml_cfg, override)
    buf = sel.list_buffer_days
    lo = today - timedelta(days=sel.max_age_days + buf)
    hi = today - timedelta(days=sel.min_age_days - buf)
    return lo, hi


def _parse_cell_date(text: str) -> date | None:
    t = (text or "").strip()
    if not t:
        return None
    try:
        dt = date_parser.parse(t, dayfirst=False, fuzzy=True)
        return dt.date()
    except (ValueError, OverflowError, TypeError):
        return None


def _parse_order_id(text: str) -> str | None:
    t = (text or "").strip().replace("\u00a0", " ")
    if t.isdigit():
        return t
    m = re.search(r"\b(\d{6,})\b", t)
    return m.group(1) if m else None


def _age_days(today: date, flight: date) -> int:
    return (today - flight).days


def _norm_header_cell(text: str) -> str:
    return " ".join((text or "").split()).strip()


def _resolve_id_flight_columns(headers: list[str]) -> tuple[int, int]:
    """Индексы колонок ID и «Дата вылета» по тексту заголовков (как в bf-table)."""
    hid = hflight = None
    for i, raw in enumerate(headers):
        h = _norm_header_cell(raw)
        low = h.lower()
        if low == "id" or low.endswith(" id") or re.fullmatch(r"id|id заказа|№", low):
            hid = i
        if "дата вылета" in low or "вылета" in low or "departure" in low:
            hflight = i
    if hid is None:
        hid = 0
    if hflight is None:
        hflight = 4
    return hid, hflight


async def _click_first_working(page: Page, selectors: list[str], log: LogFn) -> bool:
    for sel in selectors:
        if not sel:
            continue
        try:
            loc = page.locator(sel).first
            if await loc.count() and await loc.is_visible():
                await loc.click()
                await log(f"Клик по селектору: {sel!r}")
                return True
        except Exception as e:
            await log(f"Селектор {sel!r} не сработал: {e}")
    return False


async def _ensure_tab_all(page: Page, portal: BifloricaPortalConfig, log: LogFn) -> None:
    if portal.tab != "all":
        return
    try:
        all_inp = page.locator("#orderController #all-requests").first
        if await all_inp.count() and await all_inp.is_checked():
            await log("Вкладка «Все» уже выбрана — клик не нужен")
            return
    except Exception:
        pass
    ok = await _click_first_working(page, [portal.selectors.tab_all], log)
    if ok:
        await page.wait_for_timeout(500)
        return
    try:
        tab = page.get_by_role("tab", name="Все")
        if await tab.count():
            await tab.first.click()
            await log("Вкладка «Все» (ARIA tab, fallback)")
            await page.wait_for_timeout(500)
            return
    except Exception as e:
        await log(f"Вкладка «Все» (role): {e}")
    await log("Вкладка «Все» не найдена — проверьте selectors.tab_all в config.yaml")
    await page.wait_for_timeout(400)


async def _await_orders_table(page: Page, portal: BifloricaPortalConfig, log: LogFn) -> None:
    sel = portal.selectors.orders_table_ready
    try:
        await page.wait_for_selector(sel, state="visible", timeout=120_000)
    except Exception as e:
        await log(f"Таблица заказов не появилась ({sel}): {e}")
        raise


async def _collect_orders_one_page(
    page: Page,
    portal: BifloricaPortalConfig,
    today: date,
    sel_cfg: AppYamlConfig,
    selection_override: SelectionOverride | None,
    log: LogFn,
) -> list[Order]:
    s = portal.selectors
    headers: list[str] = []
    try:
        hr = page.locator(s.header_row).first
        if await hr.count():
            cells = hr.locator("th")
            n = await cells.count()
            for i in range(n):
                headers.append((await cells.nth(i).inner_text()).strip())
    except Exception as e:
        await log(f"Заголовок таблицы: {e}")

    if not headers:
        await log("Заголовки не найдены, используем fallback-колонки 0 (ID) и 4 (Дата вылета).")

    id_col, flight_col = _resolve_id_flight_columns(headers) if headers else (0, 4)

    rows = page.locator(s.orders_table_row)
    count = await rows.count()
    out: list[Order] = []
    for i in range(count):
        row = rows.nth(i)
        cells = row.locator("td")
        nc = await cells.count()
        if nc <= max(id_col, flight_col):
            continue
        id_text = (await cells.nth(id_col).inner_text()).strip()
        flight_text = (await cells.nth(flight_col).inner_text()).strip()
        oid = _parse_order_id(id_text)
        fd = _parse_cell_date(flight_text)
        if not oid or not fd:
            continue
        age = _age_days(today, fd)
        eff = _effective_selection(sel_cfg, selection_override)
        lo, hi = eff.min_age_days, eff.max_age_days
        if lo <= age <= hi:
            out.append(
                Order(
                    portal_id="biflorica",
                    order_id=oid,
                    flight_date=fd,
                    name=None,
                    cargo=None,
                )
            )
    return out


async def _goto_next_page(page: Page, portal: BifloricaPortalConfig) -> bool:
    nxt = page.locator(portal.selectors.next_page).first
    if not await nxt.count():
        return False
    try:
        if not await nxt.is_visible():
            return False
    except Exception:
        return False
    try:
        await nxt.click()
    except Exception:
        return False
    await page.wait_for_timeout(900)
    return True


async def _rewind_to_first_page(page: Page, portal: BifloricaPortalConfig, log: LogFn) -> None:
    """Возврат на первую страницу пагинации (.bf-pagination)."""
    prev_primary = (portal.selectors.prev_page or "").strip()
    prev_fallback = [
        prev_primary,
        'button:has-text("Previous")',
        '[aria-label="Previous Page"]',
    ]
    guard = 0
    while guard < 50:
        guard += 1
        moved = False
        for sel in prev_fallback:
            if not sel:
                continue
            try:
                btn = page.locator(sel).first
                if not await btn.count():
                    continue
                if not await btn.is_visible():
                    continue
                await btn.click()
                await page.wait_for_timeout(500)
                moved = True
                break
            except Exception:
                continue
        if not moved:
            break
    if guard > 1:
        await log(f"Прокрутка пагинации назад: ~{guard} шаг(ов)")


async def _collect_all_orders_paginated(
    page: Page,
    portal: BifloricaPortalConfig,
    sel_cfg: AppYamlConfig,
    today: date,
    selection_override: SelectionOverride | None,
    log: LogFn,
) -> list[Order]:
    await _rewind_to_first_page(page, portal, log)
    seen: dict[str, Order] = {}
    page_no = 0
    while True:
        page_no += 1
        batch = await _collect_orders_one_page(
            page, portal, today, sel_cfg, selection_override, log
        )
        await log(f"Страница {page_no}: заказов в диапазоне возраста на странице: {len(batch)}")
        for o in batch:
            seen[o.order_id] = o
        if not await _goto_next_page(page, portal):
            break
    return list(seen.values())


async def _find_order_row(
    page: Page,
    portal: BifloricaPortalConfig,
    order_id: str,
    log: LogFn,
) -> Locator | None:
    """Ищет строку заказа, обходя страницы с начала."""
    await page.goto(portal.orders_url, wait_until="domcontentloaded")
    try:
        await page.wait_for_load_state("networkidle", timeout=120_000)
    except Exception:
        pass
    await _await_orders_table(page, portal, log)
    await _ensure_tab_all(page, portal, log)
    await _rewind_to_first_page(page, portal, log)

    s = portal.selectors
    while True:
        row = page.locator(s.orders_table_row).filter(
            has=page.get_by_text(order_id, exact=True)
        )
        if await row.count():
            return row.first
        if not await _goto_next_page(page, portal):
            return None


async def _download_order_report(
    page: Page,
    portal: BifloricaPortalConfig,
    order_id: str,
    dest_path: Path,
) -> None:
    """Отмечает заказ и жмёт «Отчет по сделкам». Input скрыт — клик по label; без scroll_into_view на input (Angular)."""
    s = portal.selectors
    lid = f"list-orders-{order_id}"
    lbl = page.locator(f'#orderController label[for="{lid}"]')
    inp = page.locator(f"#orderController #{lid}")

    await inp.first.wait_for(state="attached", timeout=30_000)
    await lbl.first.wait_for(state="attached", timeout=10_000)
    await page.wait_for_timeout(300)

    try:
        await inp.first.evaluate(
            """el => {
                const tr = el && el.closest('tr');
                if (tr) tr.scrollIntoView({ block: 'center', inline: 'nearest' });
            }"""
        )
    except Exception:
        pass

    clicked = False
    for force in (False, True):
        try:
            await lbl.first.click(timeout=20_000, force=force)
            clicked = True
            break
        except Exception:
            await page.wait_for_timeout(250)
    if not clicked:
        await inp.first.click(timeout=20_000, force=True)

    await page.wait_for_timeout(350)
    btn = page.locator(s.deal_report_button).first
    for _ in range(40):
        try:
            if await btn.is_enabled():
                break
        except Exception:
            break
        await page.wait_for_timeout(100)
    try:
        async with page.expect_download(timeout=120_000) as dl_info:
            await page.locator(s.deal_report_button).first.click()
        dl: Download = await dl_info.value
        await dl.save_as(str(dest_path))
    finally:
        try:
            if await inp.count() and await inp.first.is_checked():
                try:
                    await lbl.first.click(timeout=10_000, force=True)
                except Exception:
                    await inp.first.click(timeout=10_000, force=True)
        except Exception:
            pass


async def run_biflorica_job(
    job_id: str,
    env: EnvSettings,
    selection_override: SelectionOverride | None = None,
) -> None:
    yaml_cfg = env.yaml_config()
    portal = yaml_cfg.portals.biflorica
    if not portal.enabled:
        await job_log(job_id, "Портал biflorica отключён в config.yaml")
        return

    root = env.project_root
    session_path = root / "data" / "sessions" / "biflorica.json"
    registry_path = root / "data" / "state" / "biflorica_downloaded.json"
    today = date.today()
    runtime = load_runtime_settings(env)
    download_dir = resolve_biflorica_download_dir(env, runtime.biflorica_download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    archive_base = resolve_biflorica_archive_dir(
        env,
        runtime.biflorica_archive_dir,
        download_dir,
        runtime=runtime,
    )
    archive_base.mkdir(parents=True, exist_ok=True)

    registry = DownloadRegistry(registry_path)
    downloaded_ids = registry.load()
    session_downloaded_paths: set[Path] = set()

    await raise_if_cancelled(job_id)

    if sys.platform == "win32":
        await job_log(
            job_id,
            f"Архив: Windows-пользователь процесса «{getpass.getuser()}» "
            f"(запускайте cvetopt.bat под тем же пользователем, что работает с C:\\Invoice)",
        )
    await job_log(job_id, f"Архив Biflorica: {archive_base}")
    await _archive_biflorica_reports(
        job_id,
        download_dir,
        archive_base,
        keep_order_ids=downloaded_ids,
        policy="unregistered_only",
        label="Архив (до скачивания)",
    )
    await job_log(job_id, f"Папка скачивания: {download_dir}")
    await job_log(job_id, f"Уже в реестре заказов: {len(downloaded_ids)}")

    pw_cfg = merged_playwright(env, yaml_cfg)
    email = env.biflorica_email
    password = env.biflorica_password
    if not email or not password:
        raise RuntimeError("Задайте BIFLORICA_EMAIL и BIFLORICA_PASSWORD в .env")

    async def lg(msg: str) -> None:
        await job_log(job_id, msg)

    async with async_playwright() as p:
        browser: Browser = await p.chromium.launch(
            headless=pw_cfg.headless,
            slow_mo=pw_cfg.slow_mo_ms or None,
        )
        context_opts: dict = {
            "accept_downloads": True,
            "viewport": {"width": 1400, "height": 900},
        }
        if session_path.exists():
            context_opts["storage_state"] = str(session_path)
            await lg("Загружена сохранённая сессия")
        context = await browser.new_context(**context_opts)
        context.set_default_navigation_timeout(pw_cfg.navigation_timeout_ms)
        page = await context.new_page()

        try:
            await lg(f"Открываю {portal.login_url}")
            await page.goto(portal.login_url, wait_until="domcontentloaded")

            try:
                await page.locator(portal.selectors.password).first.wait_for(
                    state="visible",
                    timeout=12_000,
                )
                need_login = True
            except Exception:
                need_login = False

            if need_login:
                await lg("Выполняю вход…")
                await page.locator(portal.selectors.email).first.fill(email)
                await page.locator(portal.selectors.password).first.fill(password)
                await page.locator(portal.selectors.login_submit).first.click()
                await page.wait_for_load_state("networkidle", timeout=120_000)
                session_path.parent.mkdir(parents=True, exist_ok=True)
                await context.storage_state(path=str(session_path))
                await lg("Сессия сохранена")

            await lg(f"Перехожу к заказам: {portal.orders_url}")
            await page.goto(portal.orders_url, wait_until="domcontentloaded")
            try:
                await page.wait_for_load_state("networkidle", timeout=120_000)
            except Exception:
                pass
            await _await_orders_table(page, portal, lg)

            try:
                await page.locator(portal.selectors.password).first.wait_for(
                    state="visible",
                    timeout=4000,
                )
                await lg("Снова форма логина — повторный вход")
                await page.locator(portal.selectors.email).first.fill(email)
                await page.locator(portal.selectors.password).first.fill(password)
                await page.locator(portal.selectors.login_submit).first.click()
                await page.wait_for_load_state("networkidle", timeout=120_000)
                await context.storage_state(path=str(session_path))
            except Exception:
                pass

            await _ensure_tab_all(page, portal, lg)

            lo, hi = _flight_window_for_ui(today, yaml_cfg, selection_override)
            eff = _effective_selection(yaml_cfg, selection_override)
            await lg(
                f"Ожидаемые даты вылета в выборке: {lo.isoformat()} … {hi.isoformat()} "
                f"(возраст {eff.min_age_days}–{eff.max_age_days} дн.)",
            )
            await lg(
                "Если в таблице не видно нужных заказов, выставьте на сайте фильтр дат "
                "по дате вылета в этом диапазоне и перезапустите.",
            )

            orders = await _collect_all_orders_paginated(
                page, portal, yaml_cfg, today, selection_override, lg
            )
            orders.sort(key=lambda o: o.flight_date)
            await lg(f"Всего уникальных заказов в диапазоне возраста: {len(orders)}")

            for order in orders:
                await raise_if_cancelled(job_id)
                if order.order_id in downloaded_ids:
                    await lg(f"Пропуск (уже в реестре): {order.order_id}")
                    continue
                dest = download_dir / biflorica_download_filename(
                    order.order_id, order.flight_date
                )
                legacy_dest = download_dir / (
                    f"{order.order_id}__{order.flight_date.isoformat()}.xlsx"
                )
                existing = dest if dest.exists() else legacy_dest
                if existing.exists() and existing.stat().st_size > 0:
                    await lg(f"Файл уже есть, добавляю в реестр: {existing.name}")
                    registry.add(order.order_id)
                    downloaded_ids.add(order.order_id)
                    await job_manager.add_downloaded(job_id, str(existing))
                    continue

                await lg(f"Скачиваю отчёт: заказ {order.order_id}, вылет {order.flight_date}")
                try:
                    row = await _find_order_row(page, portal, order.order_id, lg)
                    if row is None:
                        await lg(f"Строка заказа {order.order_id} не найдена при обходе страниц")
                        continue
                    await _download_order_report(page, portal, order.order_id, dest)
                except Exception as e:
                    await lg(f"Ошибка скачивания {order.order_id}: {e}")
                    logger.exception("download failed")
                    continue

                registry.add(order.order_id)
                downloaded_ids.add(order.order_id)
                await job_manager.add_downloaded(job_id, str(dest))
                session_downloaded_paths.add(dest.resolve())
                await lg(f"Сохранено: {dest}")
                await _archive_biflorica_reports(
                    job_id,
                    download_dir,
                    archive_base,
                    keep_order_ids=downloaded_ids,
                    keep_paths=session_downloaded_paths,
                    policy="stale_registered",
                    label="Архив (старые отчёты)",
                )
                ecuador_auto = env.yaml_config().ecuador_create.auto_after_biflorica
                if not ecuador_auto:
                    pass
                elif sys.platform != "win32":
                    await lg("Эквадор: пропуск (нужен Windows + Excel)")
                else:
                    await raise_if_cancelled(job_id)
                    await lg("Эквадор: запуск обработки (Excel)…")
                    try:
                        from cvetopt.invoice.ecuador_create import (
                            create_ecuador_file_from_biflorica,
                        )

                        ecuador_log_q: queue.Queue[str] = queue.Queue()

                        def _ecuador_log(msg: str) -> None:
                            ecuador_log_q.put(msg)
                            logger.info("[job {}] {}", job_id, msg)

                        async def _drain_ecuador_logs() -> None:
                            while True:
                                try:
                                    while True:
                                        await lg(ecuador_log_q.get_nowait())
                                except queue.Empty:
                                    pass
                                await asyncio.sleep(0.2)

                        drain_task = asyncio.create_task(_drain_ecuador_logs())
                        ecuador_timeout_sec = 600
                        try:
                            out = await asyncio.wait_for(
                                asyncio.to_thread(
                                    create_ecuador_file_from_biflorica,
                                    dest,
                                    env,
                                    log=_ecuador_log,
                                ),
                                timeout=ecuador_timeout_sec,
                            )
                        except TimeoutError:
                            subprocess.run(
                                ["taskkill", "/im", "EXCEL.EXE", "/f"],
                                capture_output=True,
                                check=False,
                            )
                            raise RuntimeError(
                                f"Эквадор: таймаут {ecuador_timeout_sec} с — Excel, вероятно, "
                                "ждёт диалог. Зайдите на сервер под BananaMan, один раз откройте "
                                "Excel и шаблон «Прием товара Эквадор-4.xlsm», закройте все "
                                "окна; или задайте ECUADOR_EXCEL_VISIBLE=1 в .env и перезапустите "
                                "cvetopt.bat, чтобы увидеть диалог."
                            ) from None
                        finally:
                            drain_task.cancel()
                            try:
                                await drain_task
                            except asyncio.CancelledError:
                                pass
                            try:
                                while True:
                                    await lg(ecuador_log_q.get_nowait())
                            except queue.Empty:
                                pass

                        await lg(f"Эквадор: готово → {out}")
                        await job_manager.add_downloaded(job_id, str(out))
                    except Exception as e:
                        await lg(f"Эквадор (не создан): {e}")
                        logger.exception("ecuador create failed")
                await asyncio.sleep(random.uniform(1.0, 3.0))

            if session_downloaded_paths:
                await _archive_biflorica_reports(
                    job_id,
                    download_dir,
                    archive_base,
                    keep_order_ids=downloaded_ids,
                    keep_paths=session_downloaded_paths,
                    policy="stale_registered",
                    label="Архив (после скачивания)",
                )
            await lg("Готово.")
        finally:
            await context.close()
            await browser.close()


async def _archive_biflorica_reports(
    job_id: str,
    download_dir: Path,
    archive_base: Path,
    *,
    keep_order_ids: set[str],
    keep_paths: set[Path] | None = None,
    policy: str,
    label: str,
) -> None:
    try:
        archive_dir, archived_names, archive_warnings, kept_names = (
            archive_biflorica_download_dir(
                download_dir,
                archive_base,
                keep_order_ids=keep_order_ids,
                keep_paths=keep_paths,
                policy=policy,
            )
        )
        if kept_names:
            await job_log(job_id, f"{label}: оставлено в папке: {len(kept_names)}")
        if archive_dir is not None:
            await job_log(
                job_id,
                f"{label}: перенесено {len(archived_names)} → {archive_dir}",
            )
            for warn in archive_warnings:
                await job_log(job_id, f"{label}: {warn}")
        elif policy == "stale_registered":
            await job_log(job_id, f"{label}: старых отчётов нет")
    except Exception as e:
        await job_log(job_id, f"{label}: пропущено — {e}")
