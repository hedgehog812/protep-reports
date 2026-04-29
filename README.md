# PROTEP Reports App v3

Версия с Kanban-доской заявок, ролями `admin/user`, редактированием и закрытием заявок исполнителем, а также удалением отчётов администратором.

Логины:
- admin / admin123
- user / user123

Перед обновлением выполните `supabase_schema_v3.sql` в Supabase SQL Editor.

Для Render:
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`
