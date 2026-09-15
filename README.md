# NEMAZING RP — V2

Исправленная версия бота.

## Что исправлено
- В заявлении теперь: `Введите звание в игре`, без «желаемого звания».
- Кнопка `ОТПРАВИТЬ` заявления и рапорта обрабатывается без жёсткой привязки к FSM-состоянию, поэтому не зависает из-за состояния.
- Рапорт: выбор адресата — Начальник ДПС, Начальник ВЧ, Начальник ФСБ, Начальник УМВД, Начальник ЕСС.
- Отправителя можно взять из профиля или ввести вручную.
- Рапорт получает номер `#1`, `#2`, `#3` и т.д. по ID базы.
- Рапорт сначала попадает владельцу на проверку.
- После принятия рапорта он автоматически отправляется в Telegram-чат организации, из которой пришёл сотрудник. Для этого владелец должен один раз настроить chat_id организации через `/setchat`.
- Владелец может добавлять критерии повышения через `Админ-панель → Критерии`.
- Критерии сохраняются в БД и отображаются сотрудникам в разделе `Критерии`.
- Excel содержит личный состав, заявления и рапорты, включая организацию-источник.
- Баннер автоматически ищется в `баннеры/` и `banners/` по PNG/JPG/JPEG/WEBP.

## Render
Build:
```bash
python3 -m venv .venv && .venv/bin/pip install --upgrade pip && .venv/bin/pip install -r requirements.txt
```

Start:
```bash
.venv/bin/python -u -m app.main
```

Python: `3.12.10`

## Настройка чатов организаций
Владелец пишет боту:
```text
/setchat ВЧ -1001234567890
/setchat ФСБ -1001234567890
/setchat ЕСС -1001234567890
/setchat УМВД -1001234567890
```
После этого при принятии рапорта от сотрудника ВЧ, ФСБ, ЕСС или УМВД бот автоматически отправит утверждённый рапорт в соответствующий чат.


## V4 database compatibility fix

This build includes a migration for older Render PostgreSQL databases where `organizations.title` existed as NOT NULL. The migration fills legacy titles and makes the unused legacy column safe for current inserts. It also removes the deprecated `datetime.utcnow()` warning. Do not delete the existing PostgreSQL database.
