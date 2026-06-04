from __future__ import annotations

from pathlib import Path

from cvetopt.core.runtime_settings import (
    _is_biflorica_report_to_archive,
    load_runtime_settings,
    resolve_biflorica_download_dir,
    resolve_mail_output_layout,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.mail.attachments import _REGISTRY_NAME as MAIL_REGISTRY_NAME

BIFLORICA_REGISTRY = "data/state/biflorica_downloaded.json"


def reset_testing_state(
    env: EnvSettings,
    *,
    biflorica_registry: bool = True,
    mail_registry: bool = True,
    biflorica_files: bool = False,
    mail_files: bool = False,
) -> list[str]:
    """
    Сброс реестров «уже скачано» для повторного тестирования.
    Опционально удаляет файлы из папок скачивания (не трогает runtime_settings.json).
    """
    lines: list[str] = []
    root = env.project_root
    runtime = load_runtime_settings(env)
    yaml_cfg = env.yaml_config()

    if biflorica_registry:
        path = root / BIFLORICA_REGISTRY
        if path.exists():
            path.unlink()
            lines.append(f"Реестр Biflorica удалён: {path}")
        else:
            lines.append("Реестр Biflorica: уже пуст")

    if mail_registry:
        path = root / "data" / "state" / MAIL_REGISTRY_NAME
        if path.exists():
            path.unlink()
            lines.append(f"Реестр почты удалён: {path}")
        else:
            lines.append("Реестр почты: уже пуст")

    if biflorica_files:
        try:
            dl_dir = resolve_biflorica_download_dir(env, runtime.biflorica_download_dir)
            if dl_dir.is_dir():
                removed = 0
                for entry in dl_dir.iterdir():
                    if _is_biflorica_report_to_archive(entry):
                        entry.unlink(missing_ok=True)
                        removed += 1
                lines.append(f"Файлы Biflorica: удалено отчётов {removed} из {dl_dir}")
            else:
                lines.append(f"Файлы Biflorica: папка не найдена — {dl_dir}")
        except (OSError, ValueError) as e:
            lines.append(f"Файлы Biflorica: пропуск — {e}")

    if mail_files:
        mail_cfg = yaml_cfg.mail
        layout = resolve_mail_output_layout(env, runtime, mail_cfg)
        exts = tuple(ext.lower() for ext in mail_cfg.allowed_extensions)
        removed = 0
        for out_dir in (layout.short_dir, layout.long_dir):
            if not out_dir.is_dir():
                lines.append(f"Вложения почты: папка не найдена — {out_dir}")
                continue
            for entry in out_dir.iterdir():
                if entry.is_file() and (
                    not exts or entry.name.lower().endswith(exts)
                ):
                    entry.unlink(missing_ok=True)
                    removed += 1
        lines.append(
            f"Вложения почты: удалено файлов {removed} "
            f"({layout.short_dir}, {layout.long_dir})"
        )

    if not lines:
        lines.append("Ничего не выбрано для сброса.")
    return lines
