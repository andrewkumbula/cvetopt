# Разворот на Windows Server — текущий статус

Обновлено: 2026-05-27

## Что уже успешно сделано

- [x] Подтверждено, что `winget` на сервере недоступен (`"winget" не является ...`).
- [x] Выбран путь установки вручную (без `winget`).
- [x] Установлен Git через официальный installer:
  - `Git-2.54.0-64-bit.exe`
- [x] Проверено, что Git установлен физически:
  - работает команда `"C:\Program Files\Git\cmd\git.exe" --version`
- [x] Диагностирована причина ошибки `git --version`:
  - Git не добавлен в `PATH` текущей сессии/пользователя.

## Что уже закрыто дополнительно

- [x] Git добавлен/проверен, следующий блокер — не найден Python в PATH.
- [x] Подтверждено по консоли: `python --version` и `pip --version` не выполняются (`не является ... командой`).
- [x] Подтверждено, что Python установлен и доступен через launcher:
  - `py --version`
  - `py -m pip --version`
- [x] Установлен `uv` и подтверждена версия:
  - `uv --version` → `uv 0.11.16`
- [x] Выявлена причина ошибки клонирования:
  - в `git clone` был использован плейсхолдер `<URL-репозитория>`, а не реальный URL.
- [x] Зафиксировано следствие:
  - папка `C:\Apps\cvetopt` не создана, поэтому `uv sync` запускался вне проекта и падал с `No pyproject.toml found`.
- [x] Выявлен новый блокер доступа к репозиторию:
  - `401 Unauthorized` при `git clone https://github.com/andrewkumbula/cvetopt.git`
  - GitHub по HTTPS не принимает пароль, нужна авторизация через PAT или SSH.

## После этого (разворот приложения)

- [x] Установить Python 3.11 x64 (installer):
  - `python-3.11.9-amd64.exe`
- [x] Проверить Python:

```cmd
py --version
py -m pip --version
```

- [x] Установить `uv` (выбранным способом: скрипт или zip).
- [ ] Клонировать проект:

```cmd
mkdir C:\Apps
cd C:\Apps
git clone https://github.com/andrewkumbula/cvetopt.git cvetopt
cd cvetopt
```

- [ ] Если репозиторий приватный — настроить доступ:
  - вариант A (быстрый): HTTPS + Personal Access Token (PAT)
  - вариант B: SSH-ключ и `git@github.com:andrewkumbula/cvetopt.git`

- [ ] Поднять зависимости:

```cmd
uv sync
uv run playwright install chromium
```

- [ ] Создать `.env` на основе `.env.example` и заполнить логины/пароли.
- [ ] Положить `Auto_new.xls` в корень проекта.
- [ ] Запустить:

```cmd
cvetopt.bat
```

## Примечания

- Если после `setx` Git все еще не виден, проверить:

```cmd
"C:\Program Files\Git\cmd\git.exe" --version
```

- При необходимости добавить в PATH вручную через GUI:
  - System Properties → Environment Variables → User `Path` → Add `C:\Program Files\Git\cmd`.
