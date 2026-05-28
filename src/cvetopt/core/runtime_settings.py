from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel

from cvetopt.core.settings import EnvSettings, _resolve_selection


class RuntimeSettings(BaseModel):
    biflorica_min_age_days: int
    biflorica_max_age_days: int
    delmir_lookback_days: int
    mail_lookback_days: int


def _settings_path(env: EnvSettings) -> Path:
    return env.project_root / "data" / "state" / "runtime_settings.json"


def default_runtime_settings(env: EnvSettings) -> RuntimeSettings:
    yaml_cfg = env.yaml_config()
    bif_sel = _resolve_selection(yaml_cfg.selection, yaml_cfg.portals.biflorica.selection)
    return RuntimeSettings(
        biflorica_min_age_days=bif_sel.min_age_days,
        biflorica_max_age_days=bif_sel.max_age_days,
        delmir_lookback_days=yaml_cfg.delmir.lookback_days,
        mail_lookback_days=yaml_cfg.mail.lookback_days,
    )


def load_runtime_settings(env: EnvSettings) -> RuntimeSettings:
    defaults = default_runtime_settings(env)
    path = _settings_path(env)
    if not path.exists():
        return defaults
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return defaults
    if not isinstance(raw, dict):
        return defaults
    return defaults.model_copy(update=raw)


def save_runtime_settings(env: EnvSettings, settings: RuntimeSettings) -> RuntimeSettings:
    path = _settings_path(env)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(settings.model_dump_json(indent=2), encoding="utf-8")
    tmp.replace(path)
    return settings
