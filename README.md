# cvetopt

Скачивание xlsx «Отчёт по сделкам» с портала Biflorica по правилам из `ТЗ.md`.

## Требования

- Python 3.11+
- Chromium для Playwright

## Установка

```bash
cd "/path/to/cvetopt for myx"
/opt/homebrew/bin/python3.11 -m venv .venv   # или другой Python ≥ 3.11
source .venv/bin/activate
pip install -e .
playwright install chromium
```

Скопируйте `.env.example` в `.env` и укажите `BIFLORICA_EMAIL` и `BIFLORICA_PASSWORD`.

Для отладки с видимым браузером в `.env` задайте `PLAYWRIGHT_HEADLESS=false`.

## Запуск

```bash
uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000 --app-dir src --reload
```

Флаг `--reload` подхватывает правки в `src/` без ручного перезапуска uvicorn.

Или:

```bash
cd src && uvicorn cvetopt.app:app --host 127.0.0.1 --port 8000
```

Откройте http://127.0.0.1:8000 и нажмите «Скачать отчёты (Biflorica)».

## Где лежат файлы

| Путь | Назначение |
|------|------------|
| `data/sessions/biflorica.json` | Сессия браузера после входа |
| `data/state/biflorica_downloaded.json` | Реестр `order_id`, по которым отчёт уже скачан |
| `data/downloads/biflorica/<дата>/` | Скачанные xlsx |

## Если что-то не кликается

Селекторы в `config.yaml` сверены с сохранённым образцом страницы в папке `пример страницы заказы/` (Bootstrap + Angular 1.x, таблица `bf-table-with-grey-border`, пагинация `.bf-pagination`).

При смене вёрстки портала правьте блок `portals.biflorica.selectors`. Удобнее сначала запустить с `PLAYWRIGHT_HEADLESS=false` и смотреть, на чём скрипт спотыкается.

## Сброс «уже скачано»

Удалите или отредактируйте `data/state/biflorica_downloaded.json` (ключ `order_ids`).
