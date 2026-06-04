from __future__ import annotations

import re
from collections.abc import Awaitable, Callable
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin, urlparse

from dateutil import parser as date_parser
from loguru import logger
from playwright.async_api import Browser, Locator, Page, async_playwright

from cvetopt.auto_new_xls import apply_transport_costs
from cvetopt.core.job_manager import job_log, job_manager
from cvetopt.core.settings import (
    AppYamlConfig,
    DelmirConfig,
    EnvSettings,
    merged_playwright,
)

LogFn = Callable[[str], Awaitable[None]]

_IMP = re.compile(r"\b(IMP-?\d{3,})\b", re.I)
_DATE_DOT = re.compile(r"\b(\d{2})[.\-/](\d{2})[.\-/](\d{2,4})\b")
_ISO_DATE = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")


def _url_host(url: str) -> str:
    try:
        return (urlparse(url).netloc or "").lower()
    except Exception:
        return ""


def _is_browser_error_url(url: str) -> bool:
    u = (url or "").lower()
    return u.startswith(
        ("chrome-error:", "about:", "devtools:", "chrome://", "edge://", "view-source:")
    )


def _is_same_site_host(url: str, site_origin: str) -> bool:
    """True, если url на том же хосте, что и site_origin (учитываем www)."""
    exp = _url_host(site_origin)
    act = _url_host(url)
    if not exp or not act:
        return False
    if act == exp:
        return True
    # www.del-mir.com vs del-mir.com
    def _strip_www(h: str) -> str:
        return h[4:] if h.startswith("www.") else h

    return _strip_www(act) == _strip_www(exp)


def _is_delmir_logged_in_url(url: str, cfg: DelmirConfig) -> bool:
    """Успешный вход: реальный сайт del-mir, не /login, не страница ошибки браузера."""
    if _is_browser_error_url(url):
        return False
    if "/login" in url.lower():
        return False
    return _is_same_site_host(url, cfg.site_origin)


def _digits_only(s: str | None) -> str:
    return re.sub(r"\D+", "", s or "")


def _parse_money(text: str) -> float:
    t = (text or "").replace("\xa0", " ").replace("\u2212", "-").strip()
    t = re.sub(r"[^\d.,\-]", "", t)
    if not t or t in ("-", ".", ",", "-.", "-,"):
        return 0.0
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return abs(float(t))
    except ValueError:
        return 0.0


def _parse_date_any(text: str) -> date | None:
    """Парсит ISO 2026-05-06 или 06.05.26 / 06.05.2026."""
    m = _ISO_DATE.search(text or "")
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = _DATE_DOT.search(text or "")
    if m:
        try:
            return date_parser.parse(m.group(0), dayfirst=True).date()
        except (ValueError, OverflowError, TypeError):
            pass
    return None


async def _clear_delmir_session(
    page: Page,
    sess_path: Path,
    log: LogFn,
    *,
    site_origin: str,
) -> None:
    """Стирает cookies, localStorage и файл storage_state (иначе /login → /personal без формы)."""
    ctx = page.context
    try:
        await ctx.clear_cookies()
    except Exception as e:
        await log(f"clear_cookies: {e}")
    try:
        await page.goto(site_origin.rstrip("/") + "/", wait_until="domcontentloaded", timeout=20_000)
        await page.evaluate(
            """() => {
                try { localStorage.clear(); } catch (_) {}
                try { sessionStorage.clear(); } catch (_) {}
            }"""
        )
    except Exception as e:
        await log(f"clear_storage: {e}")
    try:
        if sess_path.exists():
            sess_path.unlink()
    except Exception as e:
        await log(f"unlink session file: {e}")
    await log("Сессия del-mir сброшена (cookies, localStorage, файл).")


async def _login(
    page: Page,
    cfg: DelmirConfig,
    env: EnvSettings,
    log: LogFn,
    *,
    force: bool = False,
    sess_path: Path | None = None,
) -> None:
    sel = cfg.selectors
    email = env.delmir_email
    password = env.delmir_password
    if not email or not password:
        raise RuntimeError("Задайте DELMIR_EMAIL и DELMIR_PASSWORD в .env")

    if force:
        await _clear_delmir_session(
            page, sess_path or Path("_missing_delmir_session.json"), log, site_origin=cfg.site_origin
        )

    suffix = " (повторный вход — старая сессия не валидна)" if force else ""
    await log(f"Открываю {cfg.login_url} (email={email!r}){suffix}")
    await page.goto(cfg.login_url, wait_until="domcontentloaded")
    try:
        await page.wait_for_load_state("load", timeout=12_000)
    except Exception:
        pass

    # SPA: с валидной сессией /login сразу уводит в /personal; при force — ждём форму после очистки.
    for _ in range(24):
        url = page.url
        if not force and _is_delmir_logged_in_url(url, cfg) and "/login" not in url.lower():
            await log(f"Сессия del-mir уже активна (URL={url}).")
            return
        try:
            if await page.locator(sel.password).first.is_visible(timeout=400):
                break
        except Exception:
            pass
        await page.wait_for_timeout(500)

    try:
        await page.locator(sel.password).first.wait_for(state="visible", timeout=15_000)
    except Exception as e:
        url = page.url
        if not force and _is_delmir_logged_in_url(url, cfg) and "/login" not in url.lower():
            await log(f"Форма входа не показана, но URL кабинета ({url}) — считаем сессию активной.")
            return
        await log(f"Поле пароля не появилось: {e} (URL={url!r})")
        await _dump_debug(page, log, "login_no_password_field")
        raise

    email_loc = page.locator(sel.email).first
    pwd_loc = page.locator(sel.password).first
    try:
        await email_loc.fill("")
        await email_loc.type(email, delay=8)
        await pwd_loc.fill("")
        await pwd_loc.type(password, delay=8)
    except Exception as e:
        await log(f"Не удалось заполнить форму: {e}")
        await _dump_debug(page, log, "login_fill_failed")
        raise

    try:
        actual_email = await email_loc.input_value()
        actual_pwd_len = len(await pwd_loc.input_value())
    except Exception:
        actual_email, actual_pwd_len = "?", -1
    await log(f"Форма заполнена: email={actual_email!r}, длина пароля={actual_pwd_len}")
    if actual_email != email or actual_pwd_len != len(password):
        await log("ВНИМАНИЕ: значения в полях не совпадают с переданными. Проверьте селекторы.")

    # Слушаем все ответы — нам важны коды auth-запросов, чтобы по логу было понятно, что ответил сервер.
    auth_log: list[tuple[int, str]] = []

    def on_response(resp) -> None:
        try:
            u = resp.url.lower()
            if any(k in u for k in ("login", "auth", "sign", "session", "token")):
                auth_log.append((resp.status, resp.url))
        except Exception:
            pass

    page.on("response", on_response)
    try:
        try:
            await page.locator(sel.login_submit).first.click()
        except Exception as e:
            await log(f"Клик по кнопке «Войти» не удался: {e}; пробую Enter.")

        try:
            await page.wait_for_url(lambda url: _is_delmir_logged_in_url(url, cfg), timeout=15_000)
        except Exception:
            await log("После клика URL не сменился за 15с — добиваю Enter в поле пароля.")
            try:
                await pwd_loc.focus()
                await page.keyboard.press("Enter")
            except Exception as e:
                await log(f"Не удалось нажать Enter: {e}")
            try:
                await page.wait_for_url(lambda url: _is_delmir_logged_in_url(url, cfg), timeout=45_000)
            except Exception:
                pass

        try:
            await page.wait_for_load_state("load", timeout=12_000)
        except Exception:
            pass
    finally:
        try:
            page.remove_listener("response", on_response)
        except Exception:
            pass

    if auth_log:
        for status, u in auth_log[-10:]:
            await log(f"auth-ответ: {status} {u}")

    if "/login" in page.url.lower():
        msgs: list[str] = []
        for err_sel in (
            ".ant-form-item-explain-error",
            ".ant-message-error",
            ".ant-notification-notice-message",
            ".ant-alert-error",
        ):
            try:
                cnt = await page.locator(err_sel).count()
            except Exception:
                cnt = 0
            for i in range(cnt):
                try:
                    txt = (await page.locator(err_sel).nth(i).inner_text()).strip()
                except Exception:
                    txt = ""
                if txt:
                    msgs.append(txt)
        await _dump_debug(page, log, "login_failed")
        if any(status in (401, 403, 422) for status, _ in auth_log):
            hint = "сервер отверг учётку (401/403/422) — проверьте DELMIR_EMAIL/DELMIR_PASSWORD."
        else:
            hint = "форма отправилась, но мы остались на /login (вероятно, неверные логин/пароль или скрытая проверка)."
        msg = "; ".join(msgs) or hint
        raise RuntimeError(f"del-mir.com: вход не выполнен. {msg}")

    if not _is_delmir_logged_in_url(page.url, cfg):
        await _dump_debug(page, log, "login_bad_redirect")
        hint = (
            f"После входа ожидался сайт {cfg.site_origin}, получено: {page.url!r}. "
            "Если это chrome-error:// — сбой сети, DNS, TLS или прокси на сервере; "
            "проверьте открытие https://www.del-mir.com/login вручную в том же Chromium."
        )
        raise RuntimeError(f"del-mir.com: вход не выполнен. {hint}")

    await log(f"del-mir.com: вход выполнен (URL={page.url}).")


async def _dump_debug(page: Page, log: LogFn, tag: str) -> None:
    try:
        url = page.url
    except Exception:
        url = "?"
    await log(f"DEBUG[{tag}]: текущий URL: {url}")
    try:
        title = await page.title()
        await log(f"DEBUG[{tag}]: title: {title!r}")
    except Exception:
        pass
    try:
        body = await page.locator("body").inner_text()
        await log(f"DEBUG[{tag}]: первые 500 символов body: {body[:500]!r}")
    except Exception:
        pass
    try:
        screenshot_dir = Path("data/state")
        screenshot_dir.mkdir(parents=True, exist_ok=True)
        path = screenshot_dir / f"delmir_{tag}.png"
        await page.screenshot(path=str(path), full_page=True)
        await log(f"DEBUG[{tag}]: скриншот сохранён: {path}")
    except Exception as e:
        await log(f"DEBUG[{tag}]: не удалось сохранить скриншот: {e}")


async def _scroll_infinite_list(page: Page, log: LogFn, max_steps: int = 30) -> int:
    """Скроллит таблицу/основной контейнер вниз, пока подгружаются новые строки.
    Возвращает финальное число строк по селектору tr.ant-table-row."""
    last_count = -1
    stable = 0
    for step in range(max_steps):
        try:
            count = await page.locator("tr.ant-table-row, .ant-table-tbody > tr").count()
        except Exception:
            count = 0
        if count == last_count:
            stable += 1
            if stable >= 3:
                break
        else:
            stable = 0
            last_count = count
        try:
            await page.evaluate(
                """() => {
                    const sels = ['.ant-table-body', '.ant-table-content', '.ant-table-container', '.flex-1.overflow-auto', 'main'];
                    for (const sel of sels) {
                        const el = document.querySelector(sel);
                        if (el && el.scrollHeight > el.clientHeight) {
                            el.scrollTo({top: el.scrollHeight});
                            return;
                        }
                    }
                    window.scrollTo({top: document.body.scrollHeight});
                }"""
            )
        except Exception:
            pass
        await page.wait_for_timeout(500)
    await log(f"Доскролл /personal/balance: {last_count} строк после {step + 1} шагов")
    return max(last_count, 0)


async def _open_balance_or_relogin(
    page: Page,
    cfg: DelmirConfig,
    env: EnvSettings,
    log: LogFn,
    *,
    on_relogin: Callable[[], Awaitable[None]] | None = None,
    sess_path: Path | None = None,
) -> None:
    """
    Открывает /personal/balance. Если сервер выкинул на /login (cookies протухли),
    очищает контекст через relogin-callback (если задан) и логинится заново, потом снова
    открывает /personal/balance. Кидает RuntimeError, если и после повторного входа /login.
    """
    sel = cfg.selectors
    await page.goto(cfg.balance_url, wait_until="domcontentloaded")
    try:
        await page.wait_for_load_state("load", timeout=15_000)
    except Exception:
        pass

    if "/login" in page.url.lower() or _is_browser_error_url(page.url):
        await log(
            f"После перехода на /personal/balance мы оказались на {page.url!r} — "
            "сохранённая сессия del-mir просрочилась, делаю повторный вход."
        )
        if on_relogin is not None:
            try:
                await on_relogin()
            except Exception as e:
                await log(f"Не удалось очистить старую сессию: {e}")
        await _login(page, cfg, env, log, force=True, sess_path=sess_path)
        await page.goto(cfg.balance_url, wait_until="domcontentloaded")
        try:
            await page.wait_for_load_state("load", timeout=15_000)
        except Exception:
            pass
        if "/login" in page.url.lower() or _is_browser_error_url(page.url):
            await _dump_debug(page, log, "balance_after_relogin")
            raise RuntimeError(
                f"del-mir.com: после повторного входа /personal/balance всё ещё не открывается (URL={page.url!r})."
            )

    try:
        await page.locator(sel.balance_table).first.wait_for(state="attached", timeout=30_000)
    except Exception as e:
        await log(f"Таблица /personal/balance не найдена ({sel.balance_table}): {e}")
        await _dump_debug(page, log, "balance_no_table")
        raise


async def _gather_imp_rows(
    page: Page,
    cfg: DelmirConfig,
    env: EnvSettings,
    today: date,
    log: LogFn,
    *,
    on_relogin: Callable[[], Awaitable[None]] | None = None,
    sess_path: Path | None = None,
) -> list[dict]:
    """
    Возвращает список словарей {imp_id, date, href, raw_text} — записи /personal/balance,
    где описание содержит «IMP-…» и дата в пределах последних N дней.
    """
    sel = cfg.selectors
    await _open_balance_or_relogin(page, cfg, env, log, on_relogin=on_relogin, sess_path=sess_path)

    # Подождать первой строки или сообщения «Нет данных»
    try:
        await page.locator(
            "tr.ant-table-row, .ant-table-tbody > tr, .ant-table-placeholder"
        ).first.wait_for(state="attached", timeout=30_000)
    except Exception:
        await log("Не дождались строк таблицы за 30с — возможно, фильтр валюты ещё не применён.")

    await _scroll_infinite_list(page, log)

    rows = page.locator(sel.balance_row)
    n = await rows.count()
    await log(f"Строк в Ant-таблице: {n}")
    if n == 0:
        await _dump_debug(page, log, "balance_empty")

    cutoff = today - timedelta(days=cfg.lookback_days)
    seen_hrefs: set[str] = set()
    out: list[dict] = []

    for i in range(n):
        row = rows.nth(i)
        try:
            text = (await row.inner_text()).strip()
        except Exception:
            continue
        m = _IMP.search(text)
        if not m:
            continue
        imp_id = m.group(1).upper().replace(" ", "")
        d = _parse_date_any(text)
        if d is None:
            await log(f"{imp_id}: дату не распознал в строке — {text[:80]!r}")
            continue
        if d < cutoff or d > today:
            continue
        href: str | None = None
        try:
            link = row.locator(sel.row_link).filter(has_text=_IMP).first
            if not await link.count():
                link = row.locator(sel.row_link).first
            if await link.count():
                href = await link.get_attribute("href")
        except Exception:
            pass
        if not href:
            await log(f"{imp_id}: ссылка не найдена, пропуск")
            continue
        full = href if href.startswith("http") else urljoin(cfg.site_origin, href)
        if full in seen_hrefs:
            continue
        seen_hrefs.add(full)
        out.append({"imp_id": imp_id, "date": d.isoformat(), "href": full, "raw_text": text[:160]})

    await log(f"Подходящих IMP-записей за {cfg.lookback_days} дн.: {len(out)}")
    return out


async def _read_value_after_label(page: Page, label_selector: str) -> str | None:
    """Возвращает текст ближайшего следующего sibling-элемента после найденной метки."""
    loc = page.locator(label_selector).first
    if not await loc.count():
        return None
    try:
        val = await loc.evaluate(
            """el => {
                let n = el.nextElementSibling;
                while (n && (!n.textContent || !n.textContent.trim())) n = n.nextElementSibling;
                return n ? n.textContent.trim() : null;
            }"""
        )
        return val if val else None
    except Exception:
        return None


async def _read_siblings_after_label(page: Page, label_selector: str, limit: int = 6) -> list[str]:
    """Возвращает текст всех непустых следующих sibling-ов после первого совпадения метки."""
    loc = page.locator(label_selector).first
    if not await loc.count():
        return []
    try:
        vals = await loc.evaluate(
            """(el, limit) => {
                const out = [];
                let n = el.nextElementSibling;
                while (n && out.length < limit) {
                    const t = (n.textContent || '').trim();
                    if (t) out.push(t);
                    n = n.nextElementSibling;
                }
                return out;
            }""",
            limit,
        )
        if isinstance(vals, list):
            return [str(v) for v in vals]
    except Exception:
        return []
    return []


def _pick_final_amount(siblings: list[str]) -> float:
    """Выбирает итоговую «1638.75 $» среди соседних строк после метки «Стоимость доставки»."""
    cur_re = re.compile(r"^\s*\d[\d\s.,]*\s*[\$€₽£]\s*$")
    for s in siblings:
        if cur_re.match(s):
            return _parse_money(s)
    # Fallback: формула «$1638.75 = ...» — берём число сразу после знака валюты.
    for s in siblings:
        m = re.match(r"^\s*[\$€₽£]?\s*(\d[\d\s.,]*)", s)
        if m:
            v = _parse_money(m.group(1))
            if v > 0:
                return v
    return 0.0


async def _scrape_detail(page: Page, cfg: DelmirConfig, url: str, log: LogFn) -> tuple[str | None, float | None]:
    sel = cfg.selectors
    await page.goto(url, wait_until="domcontentloaded")
    try:
        await page.wait_for_load_state("load", timeout=15_000)
    except Exception:
        pass

    # SPA отрисовывает карточку после load — ждём сам контент, а не таймер.
    try:
        await page.locator(sel.detail_awb_label).first.wait_for(state="attached", timeout=30_000)
    except Exception:
        pass
    try:
        await page.locator(sel.detail_shipping_label).first.wait_for(
            state="attached", timeout=15_000
        )
    except Exception:
        pass

    awb_text = await _read_value_after_label(page, sel.detail_awb_label)
    awb_digits = _digits_only(awb_text) if awb_text else None

    siblings = await _read_siblings_after_label(page, sel.detail_shipping_label, limit=6)
    cost = _pick_final_amount(siblings) if siblings else 0.0

    if cost == 0:
        try:
            body = await page.locator("body").inner_text()
        except Exception:
            body = ""
        m = re.search(r"Стоимость\s+доставки[\s\S]{0,400}?(\d[\d\s.,]*)\s*[\$€₽£]", body)
        if m:
            cost = _parse_money(m.group(1))

    if not awb_digits:
        await log(f"Не удалось извлечь AWB на {url}")
    if cost <= 0:
        await log(f"Не удалось извлечь «Стоимость доставки» на {url}; рядом было: {siblings!r}")
    return awb_digits, (cost if cost > 0 else None)


async def collect_awb_to_cost(
    page: Page,
    cfg: DelmirConfig,
    env: EnvSettings,
    today: date,
    log: LogFn,
    *,
    on_relogin: Callable[[], Awaitable[None]] | None = None,
    sess_path: Path | None = None,
) -> dict[str, float]:
    rows = await _gather_imp_rows(
        page, cfg, env, today, log, on_relogin=on_relogin, sess_path=sess_path
    )
    result: dict[str, float] = {}
    for r in rows:
        await log(f"Открываю {r['imp_id']} (дата {r['date']}) → {r['href']}")
        awb, cost = await _scrape_detail(page, cfg, r["href"], log)
        if awb and cost is not None and cost > 0:
            prev = result.get(awb)
            if prev is None or prev == 0:
                result[awb] = cost
                await log(f"{r['imp_id']}: AWB={awb}, стоимость={cost}")
            else:
                await log(
                    f"AWB {awb}: уже было {prev}, новое {cost} — оставляю первое.",
                )
        else:
            await log(f"{r['imp_id']}: пропуск (AWB={awb!r}, cost={cost!r})")
    await log(f"Готовая карта AWB→cost: {len(result)} записей")
    return result


async def run_delmir_transport_job(
    job_id: str,
    env: EnvSettings,
    lookback_days_override: int | None = None,
) -> None:
    yaml_cfg = env.yaml_config()
    cfg = yaml_cfg.delmir
    bcfg = yaml_cfg.balance_auto

    if not cfg.enabled:
        await job_log(job_id, "del-mir отключён в config.yaml: delmir.enabled=false")
        return
    if lookback_days_override is not None:
        cfg = cfg.model_copy(update={"lookback_days": lookback_days_override})
        await job_log(
            job_id,
            f"Переопределён период для del-mir: {cfg.lookback_days} дн.",
        )

    async def lg(msg: str) -> None:
        await job_log(job_id, msg)

    from cvetopt.core.runtime_settings import (
        effective_auto_new_workbook_raw,
        load_runtime_settings,
        resolve_auto_new_workbook,
    )

    runtime = load_runtime_settings(env)
    wb_raw = effective_auto_new_workbook_raw(
        runtime,
        yaml_auto1=yaml_cfg.auto1_pipeline.workbook_path,
        yaml_balance=bcfg.workbook_path,
    )
    wb_path = resolve_auto_new_workbook(env, wb_raw)
    if not wb_path.exists():
        raise FileNotFoundError(wb_path)

    pw_cfg = merged_playwright(env, yaml_cfg)
    today = date.today()

    async with async_playwright() as p:
        browser: Browser = await p.chromium.launch(
            headless=pw_cfg.headless,
            slow_mo=pw_cfg.slow_mo_ms or None,
        )
        context_opts: dict = {
            "viewport": {"width": 1400, "height": 900},
        }
        sess_path = env.project_root / "data" / "sessions" / "delmir.json"
        if sess_path.exists():
            context_opts["storage_state"] = str(sess_path)
        context = await browser.new_context(**context_opts)
        context.set_default_navigation_timeout(pw_cfg.navigation_timeout_ms)
        page = await context.new_page()

        async def _on_relogin() -> None:
            await _clear_delmir_session(page, sess_path, lg, site_origin=cfg.site_origin)

        try:
            await _login(page, cfg, env, lg, sess_path=sess_path)
            sess_path.parent.mkdir(parents=True, exist_ok=True)
            try:
                await context.storage_state(path=str(sess_path))
            except Exception:
                pass

            awb_to_cost = await collect_awb_to_cost(
                page, cfg, env, today, lg, on_relogin=_on_relogin, sess_path=sess_path
            )
            try:
                await context.storage_state(path=str(sess_path))
            except Exception:
                pass
            if not awb_to_cost:
                await lg("Нет данных с del-mir.com — выходим без изменений файла.")
                return

            written, missing, notes = apply_transport_costs(wb_path, bcfg, awb_to_cost)
            for line in notes:
                await lg(line)
            if missing:
                await lg(f"AWB без совпадения на del-mir: {sorted(missing)}")
            await lg(f"Записано «Транспорт трак»: {written} ячеек. Файл: {wb_path}")
            await job_manager.add_downloaded(job_id, str(wb_path))
        except Exception as e:
            logger.exception("delmir job failed")
            await lg(f"Ошибка: {e}")
            raise
        finally:
            await context.close()
            await browser.close()
