# PROTEP Reports App v2

Расширенная версия веб-приложения для ВКР.

## Логины

Администратор: `admin / admin123`

Пользователь: `user / user123`

## Что добавлено

К основным возможностям системы можно отнести авторизацию пользователей, разграничение прав доступа, ведение объектов недвижимости, хранение договоров, регистрацию заявок на обслуживание, формирование расширенного Excel-отчёта, загрузку отчёта в Supabase Storage и просмотр ранее созданных отчётов через веб-интерфейс.

## Запуск локально

```bash
pip install -r requirements.txt
python app.py
```

## Render

Build Command:

```bash
pip install -r requirements.txt
```

Start Command:

```bash
gunicorn app:app --bind 0.0.0.0:$PORT
```

## Supabase

Выполните SQL из `supabase_schema_v2.sql`.
