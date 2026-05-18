from __future__ import annotations

import re
from collections.abc import Awaitable, Callable
from datetime import date

from dateutil import parser as date_parser
from playwright.async_api import Browser, Page, async_playwright

from cvetopt.auto_new_xls import FlightFillRow, apply_balance_flights
from cvetopt.core.job_manager import job_log, job_manager
from cvetopt.core.settings import (
    AppYamlConfig,
    BalanceAutoConfig,
    BifloricaPortalConfig,
    EnvSettings,
    _resolve_selection,
    merged_playwright,
)

LogFn = Callable[[str], Awaitable[None]]

_DATE_IN_DESC = re.compile(r"(?:от|from)\s*(\d{4}-\d{2}-\d{2})", re.I)
_FALLBACK_DATE = re.compile(r"\b(\d{4}-\d{2}-\d{2})\b")
_WEIGHT = re.compile(r"Вес[:\s]+([\d.,]+)", re.I)
_AWB = re.compile(r"AWB\s*#?\s*(\d+)", re.I)


def _age_days(today: date, flight: date) -> int:
    return (today - flight).days


def _parse_money(text: str) -> float:
    """Разбирает сумму из ячейки (в т.ч. «- $2 090.70», unicode-минус, неразрывные пробелы)."""
    t = (text or "").replace("\xa0", " ").replace("\u2212", "-").strip()
    t = re.sub(r"[^\d.,\-]", "", t)
    if not t or t in ("-", ".", ",", "-.", "-,"):
        return 0.0
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _parse_flight_date(desc: str) -> date | None:
    m = _DATE_IN_DESC.search(desc or "")
    if not m:
        m = _FALLBACK_DATE.search(desc or "")
    if not m:
        return None
    try:
        return date_parser.parse(m.group(1), dayfirst=False).date()
    except (ValueError, OverflowError, TypeError):
        return None


def _parse_weight(desc: str) -> float | None:
    m = _WEIGHT.search(desc or "")
    if not m:
        return None
    raw = m.group(1).replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return None


def _parse_awb(desc: str) -> str | None:
    m = _AWB.search(desc or "")
    return m.group(1) if m else None


def _platform(desc: str) -> str | None:
    low = (desc or "").lower()
    if "colombia" in low or "колумб" in low:
        return "colombia"
    if "ecuador" in low or "эквадор" in low:
        return "ecuador"
    return None


def _norm_header(s: str) -> str:
    return " ".join((s or "").split()).strip().lower()


def _find_col(headers: list[str], predicate) -> int | None:
    for i, h in enumerate(headers):
        if predicate(_norm_header(h)):
            return i
    return None


async def _maybe_relogin_on_page(
    page: Page,
    portal: BifloricaPortalConfig,
    env: EnvSettings,
    log: LogFn,
) -> None:
    try:
        await page.locator(portal.selectors.password).first.wait_for(
            state="visible",
            timeout=5000,
        )
    except Exception:
        return
    await log("Снова форма логина — повторный вход")
    session_path = env.project_root / "data" / "sessions" / "biflorica.json"
    await page.locator(portal.selectors.email).first.fill(env.biflorica_email)
    await page.locator(portal.selectors.password).first.fill(env.biflorica_password)
    await page.locator(portal.selectors.login_submit).first.click()
    try:
        await page.wait_for_load_state("networkidle", timeout=120_000)
    except Exception:
        pass
    session_path.parent.mkdir(parents=True, exist_ok=True)
    await page.context.storage_state(path=str(session_path))


async def _pick_balance_table(page: Page, table_sel: str, log: LogFn):
    loc = page.locator(table_sel)
    try:
        await loc.first.wait_for(state="attached", timeout=90_000)
    except Exception as e:
        await log(f"Таблица баланса ({table_sel}): {e}")
        raise
    n = await loc.count()
    for i in range(n):
        t = loc.nth(i)
        try:
            txt = await t.inner_text()
        except Exception:
            continue
        low = txt.lower()
        if "операция" in low and "описание" in low:
            await log(f"Таблица движений: вариант #{i + 1} из {n}")
            return t
    await log("Заголовки «Операция»+«Описание» не найдены — беру первую таблицу по селектору")
    return loc.first


async def _best_header_row(table, log: LogFn) -> int:
    rows = table.locator("tr")
    count = await rows.count()
    best_i = 0
    best = -1
    scan = min(count, 12)
    for i in range(scan):
        cells = rows.nth(i).locator("th, td")
        k = await cells.count()
        texts: list[str] = []
        for j in range(k):
            texts.append((await cells.nth(j).inner_text()).strip())
        joined = " ".join(_norm_header(x) for x in texts)
        score = 0
        if "операция" in joined:
            score += 3
        if "описание" in joined:
            score += 3
        if "списание" in joined:
            score += 1
        if score > best:
            best = score
            best_i = i
    await log(f"Строка заголовка таблицы баланса (индекс tr): {best_i}")
    return best_i


async def scrape_balance_flights(
    page: Page,
    cfg: BalanceAutoConfig,
    sel_cfg: AppYamlConfig,
    portal: BifloricaPortalConfig,
    env: EnvSettings,
    today: date,
    log: LogFn,
) -> list[FlightFillRow]:
    await page.goto(cfg.balance_url, wait_until="domcontentloaded")
    try:
        await page.wait_for_load_state("networkidle", timeout=120_000)
    except Exception:
        pass

    await _maybe_relogin_on_page(page, portal, env, log)

    table = await _pick_balance_table(page, cfg.selectors.table, log)
    hi = await _best_header_row(table, log)
    rows = table.locator("tr")
    n = await rows.count()

    header_cells = rows.nth(hi).locator("th, td")
    hn = await header_cells.count()
    headers: list[str] = []
    for j in range(hn):
        headers.append((await header_cells.nth(j).inner_text()).strip())

    i_op = _find_col(headers, lambda low: low == "операция" or low.startswith("операция"))
    i_desc = _find_col(headers, lambda low: "описание" in low)
    i_debit = _find_col(
        headers,
        lambda low: "списание" in low and "пополн" not in low,
    )

    if i_op is None or i_desc is None:
        await log(f"Заголовки: {headers!r}")
        raise RuntimeError("Не удалось найти колонки «Операция» / «Описание» — проверьте balance_auto.selectors.table")

    if i_debit is None:
        i_debit = _find_col(headers, lambda low: "списание" in low)
    if i_debit is None:
        raise RuntimeError("Не найдена колонка «Списание» (сумма)")

    await log(
        f"Колонки: операция={i_op}, описание={i_desc}, списание={i_debit}",
    )

    eff = _resolve_selection(sel_cfg.selection, cfg.selection)
    min_age = eff.min_age_days
    max_age = eff.max_age_days
    await log(f"Возрастное окно (balance_auto): {min_age}..{max_age} дн.")
    buffer: list[tuple[date, FlightFillRow]] = []

    for ri in range(hi + 1, n):
        row = rows.nth(ri)
        cells = row.locator("th, td")
        nc = await cells.count()
        if nc <= max(i_op, i_desc, i_debit):
            continue
        op = (await cells.nth(i_op).inner_text()).strip()
        desc = (await cells.nth(i_desc).inner_text()).strip()
        debit_txt = (await cells.nth(i_debit).inner_text()).strip()

        op_n = _norm_header(op)
        if op_n != "списание" and not op_n.startswith("списание "):
            continue
        if "списание за перелет" not in (desc or "").lower():
            continue

        fd = _parse_flight_date(desc)
        if not fd:
            await log(f"Пропуск (нет даты в описании): {desc[:120]!r}")
            continue
        age = _age_days(today, fd)
        if not (min_age <= age <= max_age):
            continue

        plat = _platform(desc)
        if not plat:
            await log(f"Пропуск (платформа не распознана): {desc[:120]!r}")
            continue

        w = _parse_weight(desc)
        awb = _parse_awb(desc)
        if w is None or not awb:
            await log(f"Пропуск (нет веса или AWB): {desc[:120]!r}")
            continue

        raw = _parse_money(debit_txt)
        # На балансе списание часто показывается отрицательным числом — в таблицу пишем положительную сумму.
        price = abs(raw)
        if price <= 0:
            await log(f"Пропуск (не удалось разобрать сумму списания): {debit_txt!r}")
            continue

        buffer.append((fd, FlightFillRow(platform=plat, weight=w, awb=awb, price=price)))

    buffer.sort(key=lambda t: t[0])
    out = [row for _, row in buffer]
    await log(f"Отобрано строк перелёта: {len(out)}")
    return out


async def run_balance_auto_job(job_id: str, env: EnvSettings) -> None:
    yaml_cfg = env.yaml_config()
    bcfg = yaml_cfg.balance_auto
    portal = yaml_cfg.portals.biflorica

    if not bcfg.enabled:
        await job_log(job_id, "balance_auto отключён в config.yaml")
        return

    root = env.project_root
    wb_path = (root / bcfg.workbook_path).resolve()

    async def lg(msg: str) -> None:
        await job_log(job_id, msg)

    pw_cfg = merged_playwright(env, yaml_cfg)
    email = env.biflorica_email
    password = env.biflorica_password
    if not email or not password:
        raise RuntimeError("Задайте BIFLORICA_EMAIL и BIFLORICA_PASSWORD в .env")

    session_path = root / "data" / "sessions" / "biflorica.json"
    today = date.today()

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
            await lg("Загружена сохранённая сессия Biflorica")
        context = await browser.new_context(**context_opts)
        context.set_default_navigation_timeout(pw_cfg.navigation_timeout_ms)
        page = await context.new_page()

        try:
            flights = await _login_and_scrape(
                page, portal, bcfg, yaml_cfg, env, today, lg,
            )
            cleared, we, wc, overflow, xls_notes = apply_balance_flights(wb_path, bcfg, flights)
            for line in xls_notes:
                await lg(line)
            await lg(
                f"Excel: очищено ячеек (без формул): {cleared}; "
                f"записано Эквадор: {we}, Колумбия: {wc}; не влезло строк: {overflow}",
            )
            await job_manager.add_downloaded(job_id, str(wb_path))
            await lg(f"Файл сохранён: {wb_path}")
            await lg("Готово.")
        finally:
            await context.close()
            await browser.close()


async def _login_and_scrape(
    page: Page,
    portal: BifloricaPortalConfig,
    bcfg: BalanceAutoConfig,
    yaml_cfg: AppYamlConfig,
    env: EnvSettings,
    today: date,
    log: LogFn,
) -> list[FlightFillRow]:
    email = env.biflorica_email
    password = env.biflorica_password

    await log(f"Открываю {portal.login_url}")
    await page.goto(portal.login_url, wait_until="domcontentloaded")

    try:
        await page.locator(portal.selectors.password).first.wait_for(
            state="visible",
            timeout=12_000,
        )
        need_login = True
    except Exception:
        need_login = False

    session_path = env.project_root / "data" / "sessions" / "biflorica.json"

    if need_login:
        await log("Выполняю вход…")
        await page.locator(portal.selectors.email).first.fill(email)
        await page.locator(portal.selectors.password).first.fill(password)
        await page.locator(portal.selectors.login_submit).first.click()
        await page.wait_for_load_state("networkidle", timeout=120_000)
        session_path.parent.mkdir(parents=True, exist_ok=True)
        await page.context.storage_state(path=str(session_path))
        await log("Сессия сохранена")

    await log(f"Перехожу на баланс: {bcfg.balance_url}")
    flights = await scrape_balance_flights(page, bcfg, yaml_cfg, portal, env, today, log)
    return flights
