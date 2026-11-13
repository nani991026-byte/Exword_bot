# Client Schedule Bot

Telegram-бот для ведения клиентов и расписания с выгрузкой в Excel.

## Возможности
- добавление записи: имя клиента, Telegram, услуга, дата, время, комментарий;
- быстрый выбор даты и времени кнопками;
- просмотр расписания по дням;
- изменение даты, времени, услуги и комментария;
- удаление записи;
- список клиентов;
- повторное использование клиента по Telegram-нику;
- экспорт расписания и клиентов в Excel;
- SQLite база данных.

## 1. Создать бота
В Telegram откройте @BotFather → `/newbot` → получите BOT_TOKEN.

## 2. Установка
Требуется Python 3.10+.

```bash
python -m venv .venv
```

Windows:
```bash
.venv\\Scripts\\activate
```

Linux/macOS/Termux:
```bash
source .venv/bin/activate
```

```bash
pip install -r requirements.txt
```

## 3. Настроить токен
Скопируйте `.env.example` в `.env` и вставьте токен:

```env
BOT_TOKEN=123456:ABCDEF...
```

`.env` не загружать в GitHub.

## 4. Запуск

```bash
python main.py
```

## 5. GitHub
Загрузите в репозиторий только:

- `main.py`
- `requirements.txt`
- `.env.example`
- `.gitignore`
- `README.md`

Файл `.env` и `data.sqlite3` в GitHub не загружать.
