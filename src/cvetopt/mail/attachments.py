from __future__ import annotations

import email
import hashlib
import imaplib
import json
import os
import re
import sys
from collections.abc import Callable, Sequence
from datetime import date, datetime, timedelta
from email.header import decode_header
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Any

from cvetopt.core.runtime_settings import (
    MailOutputLayout,
    RuntimeSettings,
    _archive_one_entry,
    _archive_target_path,
    load_runtime_settings,
    mail_destination_dir,
    resolve_mail_archive_dir,
    resolve_mail_output_layout,
)
from cvetopt.core.settings import EnvSettings, MailConfig
from cvetopt.invoice.xlsx_read import read_xlsx_grid
from cvetopt.mail.short_postprocess import clear_price_total_columns

LogFn = Callable[[str], None]

_REGISTRY_NAME = "mail_attachments_downloaded.json"
# Номер заказа/инвойса в имени вложения (2621491, 260602 …)
_MAIL_BUSINESS_REF_RE = re.compile(r"\d{6,}")


def _decode_mime_header(value: str | None) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for chunk, enc in decode_header(value):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(enc or "utf-8", errors="replace"))
        else:
            parts.append(str(chunk))
    return "".join(parts).strip()


def _sanitize_filename(name: str) -> str:
    name = name.replace("\x00", "").strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = name.strip(" .")
    return name or "attachment.bin"


def _collect_known_filenames(
    registry_keys: set[str],
    registry_entries: list[dict[str, Any]],
    *out_dirs: Path,
) -> set[str]:
    """Имена файлов, уже сохранённых или записанных в реестре (включая старый формат ключа)."""
    names: set[str] = set()
    for entry in registry_entries:
        fn = entry.get("filename")
        if fn:
            names.add(str(fn))
    for key in registry_keys:
        if "|" in key:
            names.add(key.split("|", 1)[1])
        else:
            names.add(key)
    for out_dir in out_dirs:
        if out_dir.is_dir():
            for path in out_dir.iterdir():
                if path.is_file():
                    names.add(path.name)
    return names


def _load_registry(path: Path) -> set[str]:
    if not path.exists():
        return set()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return set()
    keys: set[str] = set()
    if isinstance(raw, list):
        for item in raw:
            if isinstance(item, dict) and item.get("key"):
                keys.add(str(item["key"]))
            elif isinstance(item, str):
                keys.add(item)
    elif isinstance(raw, dict) and "keys" in raw:
        keys.update(str(k) for k in raw.get("keys", []))
    return keys


def _save_registry(path: Path, keys: set[str], entries: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def _registry_key(message_id: str, filename: str) -> str:
    return f"{message_id}|{filename}"


def _from_matches(from_hdr: str, cfg: MailConfig) -> bool:
    if not cfg.from_contains:
        return True
    low_hdr = from_hdr.lower()
    addrs = [a.lower() for _, a in getaddresses([from_hdr]) if a]
    for pattern in cfg.from_contains:
        pl = pattern.lower().strip()
        if not pl:
            continue
        if pl in low_hdr:
            return True
        if any(pl in addr or addr == pl for addr in addrs):
            return True
    return False


def _matches_filters(
    from_hdr: str,
    subject: str,
    cfg: MailConfig,
) -> bool:
    if not _from_matches(from_hdr, cfg):
        return False
    if cfg.subject_contains:
        low = subject.lower()
        if not any(s.lower() in low for s in cfg.subject_contains):
            return False
    return True


def _imap_search_criteria(cfg: MailConfig, since_imap: str) -> str:
    parts: list[str] = []
    if cfg.only_unread:
        parts.append("UNSEEN")
    parts.append(f"SINCE {since_imap}")
    from_emails = [s.strip() for s in cfg.from_contains if "@" in s]
    if len(from_emails) == 1:
        parts.append(f'FROM "{from_emails[0]}"')
    return "(" + " ".join(parts) + ")"


def _extension_ok(filename: str, cfg: MailConfig) -> bool:
    if not cfg.allowed_extensions:
        return True
    low = filename.lower()
    return any(low.endswith(ext.lower()) for ext in cfg.allowed_extensions)


def _mail_files_in_dir(out_dir: Path, cfg: MailConfig) -> list[Path]:
    if not out_dir.is_dir():
        return []
    exts = tuple(ext.lower() for ext in cfg.allowed_extensions)
    files = [
        p
        for p in out_dir.iterdir()
        if p.is_file() and (not exts or p.name.lower().endswith(exts))
    ]
    return sorted(files, key=lambda p: p.name.lower())


def _norm_cell_value(value: object) -> str:
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return repr(value)
    return str(value).strip()


def _spreadsheet_fingerprint(path: Path) -> str:
    """Сводка ячеек таблицы (без метаданных Excel) для сравнения «одинаковых» файлов."""
    raw = path.read_bytes()
    if len(raw) >= 2 and raw[:2] == b"PK":
        grid = read_xlsx_grid(path)
        lines = [f"{ref}\t{grid[ref]}" for ref in sorted(grid) if str(grid[ref]).strip()]
        return "\n".join(lines)

    import xlrd

    wb = xlrd.open_workbook(file_contents=raw)
    lines: list[str] = []
    for sheet_idx in range(wb.nsheets):
        sheet = wb.sheet_by_index(sheet_idx)
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                text = _norm_cell_value(sheet.cell_value(row, col))
                if text:
                    lines.append(f"{sheet_idx}\t{row}\t{col}\t{text}")
    return "\n".join(lines)


def _spreadsheet_data_digest(path: Path) -> str | None:
    try:
        fp = _spreadsheet_fingerprint(path)
        return hashlib.sha256(fp.encode("utf-8")).hexdigest()
    except Exception:
        return None


def _mail_business_ref_key(path: Path) -> str | None:
    """
    Ключ «тот же инвойс» из имени файла: самый длинный номер из 6+ цифр
    (2621491 в PRO FORMA и в 20-05-47-2621491…).
    """
    nums = _MAIL_BUSINESS_REF_RE.findall(path.stem)
    if not nums:
        return None
    return max(nums, key=len)


def dedupe_mail_downloads_by_content(
    out_dirs: Sequence[Path],
    cfg: MailConfig,
    log: LogFn | None = None,
) -> list[Path]:
    """
    Удаляет лишние копии с разными именами:
    сначала байт-в-байт, затем по данным ячеек (.xls/.xlsx).
    Оставляет файл с самым ранним временем изменения в папке.
    """
    def _log(msg: str) -> None:
        if log is not None:
            log(msg)

    if not cfg.dedupe_same_content:
        return []

    files: list[Path] = []
    for out_dir in out_dirs:
        files.extend(_mail_files_in_dir(out_dir, cfg))
    if len(files) < 2:
        return []

    removed: list[Path] = []

    def _remove_dupes(groups: dict[str, list[Path]], label: str) -> None:
        nonlocal removed
        for group in groups.values():
            if len(group) < 2:
                continue
            group.sort(key=lambda p: p.stat().st_mtime)
            keep = group[0]
            for dup in group[1:]:
                if not dup.exists():
                    continue
                dup.unlink(missing_ok=True)
                removed.append(dup)
                _log(f"Дубликат ({label}): удалён {dup.name} (оставлен {keep.name})")

    by_bytes: dict[str, list[Path]] = {}
    for path in files:
        if not path.exists():
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        by_bytes.setdefault(digest, []).append(path)
    _remove_dupes(by_bytes, "файл байт-в-байт")

    remaining = [p for p in files if p.exists()]
    if len(remaining) < 2:
        if removed:
            _log(f"Постобработка: удалено дубликатов {len(removed)}")
        return removed

    by_sheet: dict[str, list[Path]] = {}
    for path in remaining:
        digest = _spreadsheet_data_digest(path)
        if digest is None:
            _log(f"Постобработка: не удалось прочитать таблицу — {path.name}")
            continue
        by_sheet.setdefault(digest, []).append(path)
    _remove_dupes(by_sheet, "данные таблицы")

    remaining = [p for p in files if p.exists()]
    by_ref: dict[str, list[Path]] = {}
    for path in remaining:
        ref = _mail_business_ref_key(path)
        if ref:
            by_ref.setdefault(ref, []).append(path)
    for ref, group in by_ref.items():
        if len(group) < 2:
            continue
        group.sort(key=lambda p: p.stat().st_mtime)
        keep = group[0]
        for dup in group[1:]:
            if not dup.exists():
                continue
            dup.unlink(missing_ok=True)
            removed.append(dup)
            _log(
                f"Дубликат (номер {ref} в имени): удалён {dup.name} "
                f"(оставлен {keep.name})"
            )

    if removed:
        _log(f"Постобработка: удалено дубликатов {len(removed)}")
    return removed


def archive_stale_mail_attachments(
    layout: MailOutputLayout,
    archive_dir: Path,
    *,
    keep_paths: Sequence[Path],
    cfg: MailConfig,
    log: LogFn | None = None,
) -> tuple[Path | None, list[str], list[str]]:
    """
    После скачивания новых вложений переносит остальные .xls/.xlsx
    из папок почты 1 и 2 в архив (новые файлы остаются).
    """
    def _log(msg: str) -> None:
        if log is not None:
            log(msg)

    archive_dir = archive_dir.resolve()
    short_dir = layout.short_dir.resolve()
    long_dir = layout.long_dir.resolve()
    if archive_dir in (short_dir, long_dir):
        raise ValueError("Папка архива почты не может совпадать с папками 1 или 2.")

    keep = {p.resolve() for p in keep_paths}
    candidates: list[Path] = []
    for out_dir in (short_dir, long_dir):
        candidates.extend(_mail_files_in_dir(out_dir, cfg))

    to_move = [p for p in candidates if p.resolve() not in keep]
    if not to_move:
        _log("Почта: старых файлов для архива нет")
        return None, [], []

    archive_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    moved: list[str] = []
    warnings: list[str] = []

    if sys.platform == "win32" and not os.access(archive_dir, os.W_OK):
        import getpass

        warnings.append(
            f"Папка архива {archive_dir}: нет записи для «{getpass.getuser()}»."
        )

    for src in sorted(to_move, key=lambda p: p.name.lower()):
        target = _archive_target_path(archive_dir, src, stamp)
        if not os.access(src, os.R_OK):
            warnings.append(f"{src.name}: пропуск (нет чтения)")
            continue
        try:
            warn = _archive_one_entry(src, target)
            moved.append(src.name)
            if warn:
                warnings.append(warn)
        except OSError as e:
            warnings.append(f"{src.name}: не удалось архивировать — {e}")

    if moved:
        _log(f"Почта: в архив {len(moved)} → {archive_dir}")
    for warn in warnings:
        _log(f"Почта (архив): {warn}")

    return (archive_dir if moved else None), moved, warnings


def collect_mail_attachments(
    cfg: MailConfig,
    env: EnvSettings,
    log: LogFn | None = None,
    runtime: RuntimeSettings | None = None,
) -> tuple[list[Path], list[str]]:
    logs: list[str] = []

    def _log(msg: str) -> None:
        logs.append(msg)
        if log is not None:
            log(msg)
    """
    Подключается к IMAP, обходит письма за lookback_days.
    Короткие имена файлов → папка 1, длинные → папка 2 (см. настройки почты).
    """
    mail_email = (env.mail_email or "").strip()
    mail_password = (env.mail_password or "").strip()
    if not mail_email or not mail_password:
        raise RuntimeError(
            "Задайте MAIL_EMAIL и MAIL_PASSWORD в .env. "
            "Яндекс: включите IMAP в настройках почты и при двухфакторной аутентификации "
            "создайте пароль приложения — https://id.yandex.ru/security/app-passwords"
        )

    host = env.mail_imap_host or cfg.imap_host
    port = cfg.imap_port

    root = env.project_root
    runtime = runtime or load_runtime_settings(env)
    layout = resolve_mail_output_layout(env, runtime, cfg)
    layout.short_dir.mkdir(parents=True, exist_ok=True)
    layout.long_dir.mkdir(parents=True, exist_ok=True)
    registry_path = root / "data" / "state" / _REGISTRY_NAME
    registry_keys = _load_registry(registry_path)
    registry_entries: list[dict[str, Any]] = []
    if registry_path.exists():
        try:
            raw = json.loads(registry_path.read_text(encoding="utf-8"))
            if isinstance(raw, list):
                registry_entries = [e for e in raw if isinstance(e, dict)]
        except (json.JSONDecodeError, OSError):
            registry_entries = []

    _log(f"Почта: подключаюсь к {host}:{port}, ящик {mail_email!r}, папка {cfg.folder!r}")
    _log(f"Пароль в .env задан (длина {len(mail_password)} символов)")
    _log(
        f"Почта: папка 1 (имя ≤{layout.short_max_len} симв.): {layout.short_dir}"
    )
    _log(f"Почта: папка 2 (длинные имена): {layout.long_dir}")

    if cfg.from_contains:
        _log(f"Почта: только от отправителя: {', '.join(cfg.from_contains)}")
    if cfg.subject_contains:
        _log(f"Почта: фильтр темы: {', '.join(cfg.subject_contains)}")

    since = date.today() - timedelta(days=cfg.lookback_days)
    since_imap = since.strftime("%d-%b-%Y")

    if cfg.use_ssl:
        imap = imaplib.IMAP4_SSL(host, port)
    else:
        imap = imaplib.IMAP4(host, port)

    saved: list[Path] = []
    try:
        try:
            imap.login(mail_email, mail_password)
        except imaplib.IMAP4.error as e:
            if e.args and isinstance(e.args[0], bytes):
                err = e.args[0].decode("utf-8", errors="replace")
            else:
                err = str(e)
            raise RuntimeError(
                "Яндекс IMAP: вход не выполнен (AUTHENTICATIONFAILED). "
                "Проверьте по порядку:\n"
                "1) IMAP включён (галочка imap.yandex.ru) — у вас это уже есть.\n"
                "2) Если в настройках выбрано «Пароли приложений и OAuth-токены» "
                "(как на типичной странице «Почтовые программы»), в MAIL_PASSWORD "
                "нужен только пароль приложения, НЕ пароль от входа на yandex.ru.\n"
                "   Создать: Почта → ссылка «Пароли приложений» или "
                "https://id.yandex.ru/security/app-passwords → Почта/IMAP.\n"
                f"3) Логин в .env — полный адрес: {mail_email!r}\n"
                "4) После смены пароля в .env перезапустите uvicorn.\n"
                "5) Если в пароле есть # или пробелы — возьмите значение в кавычки в .env.\n"
                f"Ответ сервера: {err}"
            ) from e
        status, _ = imap.select(cfg.folder, readonly=not cfg.mark_as_seen)
        if status != "OK":
            raise RuntimeError(f"Не удалось открыть папку {cfg.folder!r}")

        criteria = _imap_search_criteria(cfg, since_imap)
        status, data = imap.search(None, criteria)
        if status != "OK":
            raise RuntimeError(f"IMAP SEARCH не удался: {criteria}")

        ids = (data[0] or b"").split()
        _log(f"Писем по IMAP {criteria} за {cfg.lookback_days} дн.: {len(ids)}")
        known_filenames = _collect_known_filenames(
            registry_keys, registry_entries, layout.short_dir, layout.long_dir
        )
        if cfg.allowed_extensions:
            _log(f"Типы вложений: {', '.join(cfg.allowed_extensions)}")

        for num in ids:
            status, msg_data = imap.fetch(num, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            if not isinstance(raw, bytes):
                continue
            msg = email.message_from_bytes(raw)
            message_id = (msg.get("Message-ID") or f"uid-{num.decode()}").strip()
            subject = _decode_mime_header(msg.get("Subject"))
            from_hdr = _decode_mime_header(msg.get("From"))
            try:
                msg_date = parsedate_to_datetime(msg.get("Date") or "").date()
            except (ValueError, TypeError, OverflowError):
                msg_date = None

            if not _matches_filters(from_hdr, subject, cfg):
                continue

            att_count = 0
            for part in msg.walk():
                disp = (part.get_content_disposition() or "").lower()
                if disp != "attachment":
                    continue
                raw_name = part.get_filename()
                if not raw_name:
                    continue
                filename = _sanitize_filename(_decode_mime_header(raw_name))
                if not _extension_ok(filename, cfg):
                    _log(f"Пропуск вложения (расширение): {filename!r} ← {subject[:60]!r}")
                    continue

                key = _registry_key(message_id, filename)
                if key in registry_keys:
                    _log(f"Уже скачано ранее: {filename!r} ({subject[:50]!r})")
                    continue

                if cfg.skip_if_filename_exists and filename in known_filenames:
                    existing = layout.short_dir / filename
                    if not existing.is_file():
                        existing = layout.long_dir / filename
                    _log(
                        f"Пропуск (файл с таким именем уже есть): {filename!r} "
                        f"← «{subject[:50]}»"
                    )
                    registry_keys.add(key)
                    registry_entries.append(
                        {
                            "key": key,
                            "message_id": message_id,
                            "filename": filename,
                            "path": str(existing) if existing.is_file() else None,
                            "subject": subject[:200],
                            "from": from_hdr[:200],
                            "date": msg_date.isoformat() if msg_date else None,
                            "downloaded_at": datetime.utcnow().isoformat() + "Z",
                            "skipped_duplicate_name": True,
                        }
                    )
                    continue

                dest_dir = mail_destination_dir(filename, layout)
                dest = dest_dir / filename
                slot = "1" if dest_dir == layout.short_dir else "2"

                payload = part.get_payload(decode=True)
                if not payload:
                    _log(f"Пустое вложение: {filename!r}")
                    continue

                dest.write_bytes(payload)
                if slot == "1" and cfg.clear_price_total_in_short:
                    try:
                        if clear_price_total_columns(dest):
                            _log(
                                f"Почта: очищены price и total (включая заголовки) в {dest.name}"
                            )
                        else:
                            _log(
                                f"Почта: столбцы price/total не найдены в {dest.name} "
                                "(файл без изменений)"
                            )
                    except Exception as e:
                        _log(f"Почта: не удалось очистить price/total в {dest.name}: {e}")
                att_count += 1
                saved.append(dest)
                known_filenames.add(filename)
                registry_keys.add(key)
                registry_entries.append(
                    {
                        "key": key,
                        "message_id": message_id,
                        "filename": filename,
                        "path": str(dest),
                        "mail_slot": slot,
                        "subject": subject[:200],
                        "from": from_hdr[:200],
                        "date": msg_date.isoformat() if msg_date else None,
                        "downloaded_at": datetime.utcnow().isoformat() + "Z",
                    }
                )
                _log(
                    f"Сохранено в папку {slot}: {dest.name} ({len(payload)} байт) "
                    f"← «{subject[:70]}» от {from_hdr[:50]}"
                )

            if att_count == 0 and _matches_filters(from_hdr, subject, cfg):
                if cfg.from_contains or cfg.subject_contains:
                    _log(f"Письмо без подходящих вложений: «{subject[:70]}»")

            if cfg.mark_as_seen:
                imap.store(num, "+FLAGS", "\\Seen")

    finally:
        try:
            imap.logout()
        except Exception:
            pass

    _save_registry(registry_path, registry_keys, registry_entries)
    dedupe_mail_downloads_by_content(
        [layout.short_dir, layout.long_dir], cfg, log=_log
    )
    if cfg.archive_previous_on_download and saved:
        try:
            archive_dir = resolve_mail_archive_dir(env, runtime, layout, cfg)
            archive_stale_mail_attachments(
                layout,
                archive_dir,
                keep_paths=saved,
                cfg=cfg,
                log=_log,
            )
        except Exception as e:
            _log(f"Почта: архив пропущен — {e}")
    elif cfg.archive_previous_on_download and not saved:
        _log("Почта: архив не нужен — новых вложений не скачано")
    _log(
        f"Готово: новых файлов {len(saved)}; каталоги: {layout.short_dir} и {layout.long_dir}"
    )
    return saved, logs
