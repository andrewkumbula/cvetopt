from __future__ import annotations

import email
import imaplib
import json
import re
from collections.abc import Callable
from datetime import date, datetime, timedelta
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any

from cvetopt.core.settings import EnvSettings, MailConfig

LogFn = Callable[[str], None]

_REGISTRY_NAME = "mail_attachments_downloaded.json"


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


def _unique_path(dest_dir: Path, filename: str) -> Path:
    base = dest_dir / filename
    if not base.exists():
        return base
    stem = base.stem
    suffix = base.suffix
    for i in range(2, 10_000):
        candidate = dest_dir / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
    return dest_dir / f"{stem}_{datetime.now().strftime('%H%M%S')}{suffix}"


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


def _matches_filters(
    from_hdr: str,
    subject: str,
    cfg: MailConfig,
) -> bool:
    if cfg.from_contains:
        low = from_hdr.lower()
        if not any(s.lower() in low for s in cfg.from_contains):
            return False
    if cfg.subject_contains:
        low = subject.lower()
        if not any(s.lower() in low for s in cfg.subject_contains):
            return False
    return True


def _extension_ok(filename: str, cfg: MailConfig) -> bool:
    if not cfg.allowed_extensions:
        return True
    low = filename.lower()
    return any(low.endswith(ext.lower()) for ext in cfg.allowed_extensions)


def collect_mail_attachments(
    cfg: MailConfig,
    env: EnvSettings,
    log: LogFn | None = None,
) -> tuple[list[Path], list[str]]:
    logs: list[str] = []

    def _log(msg: str) -> None:
        logs.append(msg)
        if log is not None:
            log(msg)
    """
    Подключается к IMAP, обходит письма за lookback_days, сохраняет вложения в output_dir.
    Возвращает список путей к новым файлам.
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
    out_dir = (root / cfg.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
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
    _log(f"Каталог для вложений: {out_dir}")

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

        criteria = f'(SINCE {since_imap})'
        if cfg.only_unread:
            criteria = f'(UNSEEN SINCE {since_imap})'
        status, data = imap.search(None, criteria)
        if status != "OK":
            raise RuntimeError(f"IMAP SEARCH не удался: {criteria}")

        ids = (data[0] or b"").split()
        _log(f"Писем за последние {cfg.lookback_days} дн. (с {since.isoformat()}): {len(ids)}")

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

                dest = _unique_path(out_dir, filename)
                if cfg.skip_if_filename_exists and dest.name != filename:
                    # _unique_path picked alternate — original name existed
                    _log(f"Файл уже есть на диске, сохраняю как {dest.name}")

                payload = part.get_payload(decode=True)
                if not payload:
                    _log(f"Пустое вложение: {filename!r}")
                    continue

                dest.write_bytes(payload)
                att_count += 1
                saved.append(dest)
                registry_keys.add(key)
                registry_entries.append(
                    {
                        "key": key,
                        "message_id": message_id,
                        "filename": filename,
                        "path": str(dest),
                        "subject": subject[:200],
                        "from": from_hdr[:200],
                        "date": msg_date.isoformat() if msg_date else None,
                        "downloaded_at": datetime.utcnow().isoformat() + "Z",
                    }
                )
                _log(
                    f"Сохранено: {dest.name} ({len(payload)} байт) "
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
    _log(f"Готово: новых файлов {len(saved)}, всего в каталоге см. {out_dir}")
    return saved, logs
