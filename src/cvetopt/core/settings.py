from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import yaml
from pydantic import BaseModel, Field
from pydantic_settings import BaseSettings, SettingsConfigDict

_PROJECT_ROOT = Path(__file__).resolve().parents[3]


class SelectionConfig(BaseModel):
    min_age_days: int = 6
    max_age_days: int = 12
    list_buffer_days: int = 3


class SelectionOverride(BaseModel):
    """Опциональное переопределение возрастного окна на уровне конкретного джоба."""

    min_age_days: int | None = None
    max_age_days: int | None = None
    list_buffer_days: int | None = None


def _resolve_selection(base: SelectionConfig, override: SelectionOverride | None) -> SelectionConfig:
    """Возвращает эффективные значения окна, накладывая override на дефолты."""
    if override is None:
        return base
    return SelectionConfig(
        min_age_days=override.min_age_days if override.min_age_days is not None else base.min_age_days,
        max_age_days=override.max_age_days if override.max_age_days is not None else base.max_age_days,
        list_buffer_days=(
            override.list_buffer_days if override.list_buffer_days is not None else base.list_buffer_days
        ),
    )


class PlaywrightConfig(BaseModel):
    headless: bool = True
    slow_mo_ms: int = 0
    navigation_timeout_ms: int = 60_000


class BifloricaSelectors(BaseModel):
    """Значения по умолчанию — по сохранённому образцу «Заказы - Biflorica.com.html»."""

    email: str = '#LoginForm_username, input[name="LoginForm[username]"], input[type="email"]'
    password: str = '#LoginForm_password, input[name="LoginForm[password]"], input[type="password"]'
    login_submit: str = '#loginButton, button:has-text("Log in"), button:has-text("Login")'
    # Вкладка «Все» — Bootstrap label с radio #all-requests (не ARIA tab).
    tab_all: str = "#orderController label:has(#all-requests)"
    orders_table_ready: str = "#orderController table.bf-table-with-grey-border"
    # Две строки заголовка в tbody; данные — tr без .bf-header
    orders_table_row: str = (
        "#orderController table.bf-table-with-grey-border tbody > tr:not(.bf-header)"
    )
    header_row: str = (
        "#orderController table.bf-table-with-grey-border "
        "tbody tr.second-header-row.bf-header"
    )
    checkbox_in_row: str = "input.order-checkbox"
    deal_report_button: str = (
        '#orderController button.bf-btn-copy-ico.group-orders-btn[title="Отчет по сделкам"]'
    )
    next_page: str = (
        '#orderController .bf-pagination li.label-next a[aria-label="Next"]'
    )
    prev_page: str = (
        '#orderController .bf-pagination li.label-prev a[aria-label="Previous"]'
    )


class BifloricaPortalConfig(BaseModel):
    enabled: bool = True
    login_url: str = "https://ec.term.biflorica.ru/marketPlaceNew/balance"
    orders_url: str = "https://ec.term.biflorica.ru/marketPlaceNew/default/orders"
    tab: str = "all"
    selectors: BifloricaSelectors = Field(default_factory=BifloricaSelectors)
    # Возрастное окно дат вылета для скачивания «Отчёт по сделкам»:
    # по умолчанию 3..7 дней (переопределяет глобальную selection.*).
    selection: SelectionOverride | None = Field(
        default_factory=lambda: SelectionOverride(min_age_days=3, max_age_days=7)
    )


class PortalsConfig(BaseModel):
    biflorica: BifloricaPortalConfig = Field(default_factory=BifloricaPortalConfig)


class BalanceAutoSelectors(BaseModel):
    """Таблица движений по балансу (подстройте под DOM, если строка не находится)."""

    table: str = "table.bf-table-with-grey-border"
    tbody_rows: str = "tbody tr"


class BalanceAutoBlockConfig(BaseModel):
    """Блок строк на листе «АВБ перелеты»: первая/последняя строка данных и колонки (буквы Excel)."""

    first_data_row_excel: int = 15
    last_data_row_excel: int = 19
    weight_col: str = "K"
    awb_col: str = "L"
    price_col: str = "O"
    # Колонка «Транспорт трак» (заполняется с del-mir.com). Если None — блок не учитывается в delmir.
    transport_col: str | None = None
    # В строках first..last эти колонки очищаются всегда, в т.ч. поверх формул (иначе остаётся старый «Транспорт трак»).
    force_clear_cols: list[str] = Field(default_factory=list)


class BalanceAutoClearRange(BaseModel):
    """Прямоугольник очистки в нотации Excel (включительно). Ячейки с формулами не затираются."""

    top_left: str = "K15"
    bottom_right: str = "P19"


class BalanceAutoConfig(BaseModel):
    enabled: bool = True
    workbook_path: str = "Auto_new.xls"
    sheet_name: str = "АВБ перелеты"
    balance_url: str = "https://ec.term.biflorica.ru/marketPlaceNew/balance"
    # auto: Excel через xlwings (сохраняет VBA), иначе xlwt (макросы в .xls пропадают).
    excel_engine: Literal["auto", "xlwings", "xlwt"] = "auto"
    backup_before_save: bool = True
    backup_suffix: str = ".bak"
    # Если задано — переопределяет глобальную selection.* только для этой задачи.
    selection: SelectionOverride | None = None
    selectors: BalanceAutoSelectors = Field(default_factory=BalanceAutoSelectors)
    clear_ranges: list[BalanceAutoClearRange] = Field(
        default_factory=lambda: [
            BalanceAutoClearRange(top_left="K15", bottom_right="P19"),
            BalanceAutoClearRange(top_left="A28", bottom_right="F34"),
        ]
    )
    ecuador: BalanceAutoBlockConfig = Field(
        default_factory=lambda: BalanceAutoBlockConfig(
            first_data_row_excel=15,
            last_data_row_excel=19,
            weight_col="K",
            awb_col="L",
            price_col="N",
            transport_col="M",
            force_clear_cols=["M"],
        )
    )
    colombia: BalanceAutoBlockConfig = Field(
        default_factory=lambda: BalanceAutoBlockConfig(
            first_data_row_excel=28,
            last_data_row_excel=34,
            weight_col="A",
            awb_col="B",
            price_col="F",
            transport_col="C",
            force_clear_cols=["C"],
        )
    )


class Auto1PipelineConfig(BaseModel):
    """Цепочка макросов на листе auto1 (Scan → … → for sklad) через Excel на Windows."""

    enabled: bool = True
    workbook_path: str = "Auto_new.xls"
    sheet_name: str = "auto1"
    backup_before_run: bool = True
    backup_suffix: str = ".bak"


class HollandTranslateConfig(BaseModel):
    """Перевод Description в Голландия_1_*.xlsx по Словарь.xls (B → C)."""

    enabled: bool = True
    sklad_output_dir: str = r"C:\Инвойсы склад"
    dictionary_path: str = "Invoice/Словарь.xls"


class DelmirSelectors(BaseModel):
    """Селекторы del-mir.com — сверены с сохранённой страницей «Мой баланс»."""

    # Форма входа Ant Design Form v4. Имя формы зависит от страницы, потому ищем по типу/имени поля.
    email: str = (
        '#edit_car_username, input[type="email"], input[name="username"], '
        'input[name="email"], input[name="login"], input[id$="_username"], '
        'input[id$="_email"], input[id$="_login"]'
    )
    password: str = (
        '#edit_car_password, input[type="password"], input[name="password"], '
        'input[id$="_password"]'
    )
    login_submit: str = 'button[type="submit"]:has-text("Войти"), button:has-text("Войти"), button[type="submit"]'
    # /personal/balance — Ant Table v4. У <table> класса нет; ant-table — на обёртке.
    balance_table: str = ".ant-table-wrapper, .ant-table, .ant-table-content, table"
    balance_row: str = "tr.ant-table-row, .ant-table-tbody > tr, tbody > tr"
    # Описание содержит ссылку на /personal/<id>?id=... с текстом «IMP-…».
    row_link: str = 'a[href*="/personal/"]'
    # На детальной странице — лейблы «№ AWB» и «Стоимость доставки» (см. [id]-*.js).
    detail_awb_label: str = 'p:has-text("№ AWB")'
    detail_shipping_label: str = 'p:has-text("Стоимость доставки")'


class MailConfig(BaseModel):
    """Скачивание вложений из почты по IMAP (Яндекс, Mail.ru, Gmail и др.)."""

    enabled: bool = True
    imap_host: str = "imap.yandex.ru"
    imap_port: int = 993
    use_ssl: bool = True
    folder: str = "INBOX"
    lookback_days: int = 14
    # Устаревшее: одна папка; если в UI не заданы short/long — берётся как родитель.
    output_dir: str = "data/downloads/mail"
    output_dir_short: str = "data/downloads/mail/1"
    output_dir_long: str = "data/downloads/mail/2"
    # Имя файла (с расширением) не длиннее — в папку «1», иначе в папку «2».
    filename_short_max_len: int = 35
    only_unread: bool = False
    mark_as_seen: bool = False
    # Не сохранять второй раз файл с тем же именем (в т.ч. из другого письма).
    skip_if_filename_exists: bool = True
    # После загрузки удалить копии: одинаковые байты или одинаковые ячейки в .xls/.xlsx.
    dedupe_same_content: bool = True
    # В папке 1 (короткие имена): очистить столбцы price и total после сохранения.
    clear_price_total_in_short: bool = True
    allowed_extensions: list[str] = Field(default_factory=lambda: [".xlsx", ".xls"])
    # Пустой список = не фильтровать. Иначе — хотя бы одна подстрока должна совпасть.
    from_contains: list[str] = Field(default_factory=list)
    subject_contains: list[str] = Field(default_factory=list)


class DelmirConfig(BaseModel):
    enabled: bool = False
    # Главная редиректит на /login, но открываем явный URL формы.
    login_url: str = "https://www.del-mir.com/login"
    balance_url: str = "https://www.del-mir.com/personal/balance"
    site_origin: str = "https://www.del-mir.com"
    lookback_days: int = 3
    description_prefix: str = "IMP"
    selectors: DelmirSelectors = Field(default_factory=DelmirSelectors)


class AppYamlConfig(BaseModel):
    selection: SelectionConfig = Field(default_factory=SelectionConfig)
    playwright: PlaywrightConfig = Field(default_factory=PlaywrightConfig)
    portals: PortalsConfig = Field(default_factory=PortalsConfig)
    balance_auto: BalanceAutoConfig = Field(default_factory=BalanceAutoConfig)
    auto1_pipeline: Auto1PipelineConfig = Field(default_factory=Auto1PipelineConfig)
    holland_translate: HollandTranslateConfig = Field(default_factory=HollandTranslateConfig)
    delmir: DelmirConfig = Field(default_factory=DelmirConfig)
    mail: MailConfig = Field(default_factory=MailConfig)


def load_yaml_config(path: Path) -> AppYamlConfig:
    if not path.exists():
        return AppYamlConfig()
    raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return AppYamlConfig.model_validate(raw)


class EnvSettings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=str(_PROJECT_ROOT / ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    biflorica_email: str = ""
    biflorica_password: str = ""

    delmir_email: str = ""
    delmir_password: str = ""

    # Почта (IMAP). Яндекс: полный адрес + пароль (при 2FA — пароль приложения).
    mail_email: str = ""
    mail_password: str = ""
    mail_imap_host: str = ""

    app_host: str = "127.0.0.1"
    app_port: int = 8000
    log_level: str = "INFO"

    playwright_headless: bool = True

    @property
    def project_root(self) -> Path:
        return _PROJECT_ROOT

    def yaml_config(self) -> AppYamlConfig:
        return load_yaml_config(self.project_root / "config.yaml")


def merged_playwright(env: EnvSettings, yaml_cfg: AppYamlConfig) -> PlaywrightConfig:
    pw = yaml_cfg.playwright.model_copy()
    pw.headless = env.playwright_headless
    return pw
