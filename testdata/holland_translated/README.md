# Локальный плацдарм: «Переведено»

Без Windows/Excel: проверяется только **перевод Description + ростовка у роз**.
Шаг Auto1 (Scan…Sort) на Mac не гоняется — его смотрите на сервере.

## Структура

| Путь | Назначение |
|------|------------|
| `pristine/1`, `pristine/2` | Исходные почтовые Excel (не трогать) |
| `work/1`, `work/2` | Рабочие копии — сюда пишет прогон |
| `Словарь.xls` | Тестовый словарь (B→C) |

## Команды

Из корня репозитория:

```bash
# полный цикл: сброс → перевод+ростовка → проверка
.venv/bin/python scripts/holland_translated_local.py

# только восстановить work/ из pristine/
.venv/bin/python scripts/holland_translated_local.py --reset-only

# только проверить уже обработанные work/
.venv/bin/python scripts/holland_translated_local.py --verify-only
```

Ожидание после полного прогона: `VERIFY OK`, у роз в названии длина (`… 60` / `… 70`), у хризантем — без длины.

## Через UI (опционально)

1. Запустить uvicorn.
2. В **Настройки** временно указать:
   - почта, папка 1 → `testdata/holland_translated/work/1`
   - почта, папка 2 → `testdata/holland_translated/work/2`
   - словарь → `testdata/holland_translated/Словарь.xls`
3. Сначала `--reset-only`, потом кнопка **«Переведено…»**.
4. Auto1 на Mac пропустится с сообщением в логе — это нормально.

После теста верните в Настройках серверные пути `C:\Invoice\…`.
