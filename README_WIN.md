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

### Папка `Invoice`

В корне проекта создайте (или скопируйте) каталог **`Invoice`** — это рабочая зона на сервере: сюда попадают скачанные отчёты и другие xlsx для дальнейшей обработки. Папка **не коммитится** в git (см. `.gitignore`).

В веб-интерфейсе → **Настройки** укажите, например:

- Biflorica, папка на ПК: `C:\Apps\cvetopt\Invoice` (или подпапка внутри)
- Biflorica, папка архива: `C:\Apps\cvetopt\Invoice\архив`

Пути сохраняются в `data\state\runtime_settings.json`.

После скачивания отчёта Biflorica (только Windows + Excel) автоматически создаётся файл **«Эквадор … .xlsm»** в папке выгрузки (по умолчанию `D:\Склад ОБмен\Инвойсы Склад`). Шаблон и папку можно сменить в **Настройки** → «Эквадор, шаблон» / «папка Создать файл».

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

Войдите под сервисной учёткой (той же, что в логе cvetopt, например BananaMan), запустите Excel → согласитесь со всеми диалогами активации/Privacy → закройте. Затем **один раз** откройте шаблон `C:\Invoice\3\Обработка\Прием товара Эквадор-4.xlsm` → «Включить содержимое» / доверие макросам → закройте без сохранения. Иначе автоматизация может **зависнуть** на шаге «Эквадор: запуск Excel…».

Если зависло: в `.env` добавьте `ECUADOR_EXCEL_VISIBLE=1`, перезапустите `cvetopt.bat` — появится окно Excel и будет виден диалог. После настройки строку можно убрать.

### 4.3. Дать DCOM-права (если приложение крутится не от того же пользователя, что Excel)

`dcomcnfg` → Component Services → Computers → My Computer → DCOM Config → найти **Microsoft Excel Application** → Properties → вкладка Security → Launch and Activation Permissions → Customize → Add → добавить сервисную учётку с правами Local Launch / Local Activation.

В большинстве случаев пункт не нужен (если приложение запускается под тем же пользователем). Делайте только если xlwings ругается на `0x80070005 / Access Denied`.

### 4.4. Права на `C:\Invoice` и архив Biflorica (WinError 5)

Если в логе: `BiFlorica-….xlsx: не удалось архивировать — [WinError 5] Отказано в доступе`, это **не «другой тип» учётки**, а несовпадение **пользователя процесса** и **владельца файлов/папок**.

В начале прогона Biflorica в логе есть строка  
`Архив: Windows-пользователь процесса «…»` — это учётка, под которой реально крутится uvicorn. Она должна совпадать с тем, кто заходит по RDP и кладёт файлы в `C:\Invoice\3`.

**Частые причины:**

| Ситуация | Что сделать |
|----------|-------------|
| `cvetopt.bat` в Планировщике с «Выполнять вне зависимости от входа» / другой пользователь | Триггер **«При входе в систему»** для **той же** учётки, что и RDP; не SYSTEM |
| Файлы создавались под `UserA`, сервер запущен под `UserB` | Запускать `cvetopt.bat` только под `UserA` или выдать права `UserB` |
| Нет записи в папку архива (`C:\Invoice\3\архив` или отдельный путь в настройках) | Права **Modify** на папку архива и скачивания |
| Excel держит xlsx открытым | Закрыть книгу; иногда помогает `taskkill /im EXCEL.EXE /f` перед прогоном |

**Проверка в cmd под той же учёткой, что в логе cvetopt** (подставьте своего пользователя `SERVER\invoice`):

```cmd
whoami
icacls "C:\Invoice\3"
icacls "C:\Invoice\3\архив"
```

**Выдать права сервисной учётке** (cmd **от администратора**, замените `SERVER\invoice`):

```cmd
icacls "C:\Invoice" /grant "SERVER\invoice:(OI)(CI)M" /T
```

`(OI)(CI)M` — изменение файлов и подпапок (достаточно для копирования в архив и удаления старых xlsx). Полный контроль не обязателен.

После `icacls` перезапустите `cvetopt.bat` и снова запустите Biflorica. Если ошибка останется — пришлите из лога строку с `Windows-пользователь процесса` и вывод `icacls "C:\Invoice\3"` для одного проблемного файла.

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

### Для пользователя (рекомендуется)

**`cvetopt-launcher.vbs`** — как обычная программа Windows:

1. Поднимает сервер **в фоне** (без чёрного окна).
2. Открывает окно **cvetopt** (Edge/Chrome в режиме `--app`).
3. Пока окно открыто — можно жать кнопки на вкладке «Запуск».
4. **Закрыли окно — сервер останавливается.** Не нужно держать процесс в фоне сутками.

Создать ярлык на рабочий стол:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\create-desktop-shortcut.ps1
```

Или вручную: ярлык → объект  
`wscript.exe "C:\Apps\cvetopt\cvetopt-launcher.vbs"`.

Если сделали PWA из браузера — **замените цель ярлыка** на лаунчер выше (иначе сервер не стартует и не остановится вместе с окном).

Принудительно остановить сервер (если завис): `cvetopt-stop.bat`.

### Для администратора / отладки

- **`cvetopt.bat`** — видимое окно консоли с логами uvicorn. Закрытие окна = остановка сервера. Выход: `Ctrl+C`.
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
