# cvetopt на Windows Server — установка и обновление

Это инструкция для развёртывания приложения на машине с Windows Server и установленным Microsoft Excel. После установки запуск — один двойной клик по `cvetopt.bat`, обновление — кнопка «🔄 Обновить программу» прямо в веб-интерфейсе.

## 0. Требования

- Windows Server 2019 или новее.
- Установленный **Microsoft Excel** (любая desktop-версия, не Excel Online). Нужен для `xlwings` — иначе VBA-макросы в `Auto_new.xls` будут потеряны.
- Учётная запись Windows, под которой будет крутиться приложение, должна иметь право интерактивного входа в систему (см. п. 4).

## 1. Поставить системные зависимости

В PowerShell от админа:

```powershell
winget install astral-sh.uv
winget install Git.Git
```

Закройте и снова откройте PowerShell, чтобы обновился `PATH`.

## 2. Развернуть проект

```powershell
cd C:\Apps
git clone <URL-вашего-репозитория> cvetopt
cd cvetopt
uv sync
uv run playwright install chromium
```

`playwright install chromium` скачает встроенный Chromium в `%USERPROFILE%\AppData\Local\ms-playwright\` — это нужно сделать **под той же учёткой**, под которой будет запускаться `cvetopt.bat` (иначе Playwright не найдёт браузер).

## 3. Конфиг и креды

Скопируйте `.env.example` (если есть) или создайте `.env` в корне проекта:

```ini
BIFLORICA_EMAIL=...
BIFLORICA_PASSWORD=...
DELMIR_EMAIL=...
DELMIR_PASSWORD=...
MAIL_EMAIL=...
MAIL_PASSWORD=...
PLAYWRIGHT_HEADLESS=true
```

Установите NTFS-права: чтение `.env` — только у владельца. Через свойства файла → Security → удалить «Users», оставить только сервисную учётку.

`Auto_new.xls` положите в корень проекта (`C:\Apps\cvetopt\Auto_new.xls`). Перед каждой записью приложение делает копию `Auto_new.xls.bak`.

## 4. Excel + xlwings: грабли Windows Server

Без этих шагов `xlwings` упадёт при попытке открыть книгу.

### 4.1. Создать «системные» Desktop-папки

В PowerShell от админа:

```powershell
New-Item -ItemType Directory -Force `
  -Path "C:\Windows\System32\config\systemprofile\Desktop"
New-Item -ItemType Directory -Force `
  -Path "C:\Windows\SysWOW64\config\systemprofile\Desktop"
```

Это классическая «лечилка» для COM-Excel на Server-системах. Без этих папок Excel падает при попытке открытия книги из автоматизации.

### 4.2. Запустить Excel руками первый раз

Войдите под сервисной учёткой, запустите Excel → согласитесь со всеми диалогами активации/Privacy → закройте. Это нужно, чтобы Excel инициализировал свои настройки в профиле пользователя.

### 4.3. Дать DCOM-права (если приложение крутится не от того же пользователя, что Excel)

`dcomcnfg` → Component Services → Computers → My Computer → DCOM Config → найти **Microsoft Excel Application** → Properties → вкладка Security → Launch and Activation Permissions → Customize → Add → добавить сервисную учётку с правами Local Launch / Local Activation.

В большинстве случаев пункт не нужен (если приложение запускается под тем же пользователем). Делайте только если xlwings ругается на `0x80070005 / Access Denied`.

## 5. Авто-запуск при старте сервера

Excel COM требует **интерактивную сессию пользователя**. Поэтому НЕ оборачивайте приложение в Windows Service из-под `LocalSystem` — это не сработает.

Рабочая схема:

1. **Auto-logon сервисной учётки.** Скачайте [Sysinternals Autologon](https://learn.microsoft.com/en-us/sysinternals/downloads/autologon), запустите, введите имя/пароль сервисной учётки → нажмите Enable. После следующей перезагрузки система автоматически логинится под этой учёткой.
2. **Task Scheduler:**
   - Триггер: «At log on of <сервисная учётка>».
   - Action: Start a program → `C:\Apps\cvetopt\cvetopt.bat`.
   - Settings → ☑ «If the task is already running… → Do not start a new instance».
   - Settings → ☑ «Run task as soon as possible after a missed start».
   - Conditions → снять «Stop if computer switches to battery power» (для серверов не критично).
3. После перезагрузки сервер сам логинится → запускается `cvetopt.bat` → uvicorn слушает 127.0.0.1:8000.

## 6. Доступ снаружи

Не публикуйте 8000 порт в интернет напрямую. Варианты:

- **Tailscale** (рекомендую): `winget install tailscale.tailscale`, авторизация, и сервер виден из любой точки только тем, кто в твоей tailnet.
- **Cloudflare Tunnel**: `winget install --id Cloudflare.cloudflared`, `cloudflared tunnel login`, привязать поддомен.
- Локальная сеть + RDP — если не нужен внешний доступ.

При желании добавьте Basic Auth — раскомментируйте middleware в `app.py` (если нет — просите добавить).

## 7. Запуск и обновление

- **Запуск:** двойной клик по `cvetopt.bat`. Через 3 секунды откроется браузер на `http://127.0.0.1:8000`. **Окно консоли не закрывайте**, пока работаете в браузере — в нём крутится сервер; закрытие = остановка программы. Выход: `Ctrl+C` в этом окне.
- **Обновление:** в веб-UI кнопка «🔄 Обновить программу». Она просит подтверждение, потом сервер сам перезапускается (`git pull --ff-only && uv sync`) и страница автоматически перезагрузится на новой версии.
- В верхней строке UI видна текущая версия (`commit ... (branch) от YYYY-MM-DD`) и состояние («готово к запуску» / «идёт прогон…»).

Если обновление сломалось (например, `uv sync` упал из-за нового пакета) — лаунчер выйдет из цикла и закроется. Откатить руками:

```powershell
cd C:\Apps\cvetopt
git log --oneline -10
git reset --hard <предыдущий-коммит>
uv sync
.\cvetopt.bat
```

## 8. Диагностика

- Логи uvicorn — в окне `cvetopt.bat`.
- Скриншоты failed-страниц (Playwright) — `data\state\*.png`.
- Бэкапы Excel — `Auto_new.xls.bak` рядом с файлом.
- Сессии (cookies) — `data\sessions\biflorica.json`, `data\sessions\delmir.json`. Удалите, если приложение перестало «помнить» вход.
- Если Excel зависает — `taskkill /im EXCEL.EXE /f` (можно повесить на ежедневный Scheduled Task в качестве страховки).
