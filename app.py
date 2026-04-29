import os
from datetime import datetime
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, flash, redirect, render_template, request, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from supabase import create_client

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "dev-secret-key-change-later")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "reports")

TEMP_DIR = Path("generated_reports")
TEMP_DIR.mkdir(exist_ok=True)

USERS = {
    "admin": {"password": "admin123", "role": "admin", "display_name": "Администратор"},
    "user": {"password": "user123", "role": "user", "display_name": "Пользователь"},
}

def db():
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("Не заполнены SUPABASE_URL и SUPABASE_KEY.")
    return create_client(SUPABASE_URL, SUPABASE_KEY)

def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login"))
        return func(*args, **kwargs)
    return wrapper

def admin_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            flash("Недостаточно прав для выполнения операции.", "error")
            return redirect(url_for("index"))
        return func(*args, **kwargs)
    return wrapper

def current_user():
    return {
        "username": session.get("username"),
        "role": session.get("role"),
        "display_name": session.get("display_name"),
    }

def fetch_table(name):
    return db().table(name).select("*").order("created_at", desc=True).execute().data

def fetch_contracts():
    return db().table("contracts").select("*, objects(name, address)").order("created_at", desc=True).execute().data

def fetch_requests():
    return db().table("service_requests").select("*, objects(name, address)").order("created_at", desc=True).execute().data

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        user = USERS.get(username)
        if not user or user["password"] != password:
            flash("Неверный логин или пароль.", "error")
            return redirect(url_for("login"))
        session["username"] = username
        session["role"] = user["role"]
        session["display_name"] = user["display_name"]
        return redirect(url_for("index"))
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def index():
    try:
        objects = fetch_table("objects")
        contracts = fetch_contracts()
        requests_data = fetch_requests()
        reports = fetch_table("reports")
    except Exception as error:
        objects, contracts, requests_data, reports = [], [], [], []
        flash(str(error), "error")
    return render_template(
        "index.html",
        objects=objects,
        contracts=contracts,
        service_requests=requests_data,
        reports=reports,
        user=current_user()
    )

@app.route("/objects", methods=["POST"])
@admin_required
def add_object():
    try:
        db().table("objects").insert({
            "name": request.form["name"].strip(),
            "address": request.form["address"].strip(),
            "area": float(request.form["area"].replace(",", ".")),
            "status": request.form["status"],
            "responsible_person": request.form.get("responsible_person", "").strip(),
        }).execute()
        flash("Объект недвижимости добавлен.", "success")
    except Exception as error:
        flash(f"Ошибка при добавлении объекта: {error}", "error")
    return redirect(url_for("index"))

@app.route("/contracts", methods=["POST"])
@admin_required
def add_contract():
    try:
        db().table("contracts").insert({
            "object_id": int(request.form["object_id"]),
            "tenant_name": request.form["tenant_name"].strip(),
            "contract_number": request.form["contract_number"].strip(),
            "monthly_payment": float(request.form["monthly_payment"].replace(",", ".")),
            "start_date": request.form["start_date"],
            "end_date": request.form["end_date"],
            "status": request.form["contract_status"],
        }).execute()
        flash("Договор добавлен.", "success")
    except Exception as error:
        flash(f"Ошибка при добавлении договора: {error}", "error")
    return redirect(url_for("index"))

@app.route("/requests", methods=["POST"])
@admin_required
def add_request():
    try:
        db().table("service_requests").insert({
            "object_id": int(request.form["request_object_id"]),
            "title": request.form["title"].strip(),
            "description": request.form.get("description", "").strip(),
            "priority": request.form["priority"],
            "status": request.form["request_status"],
        }).execute()
        flash("Заявка на обслуживание добавлена.", "success")
    except Exception as error:
        flash(f"Ошибка при добавлении заявки: {error}", "error")
    return redirect(url_for("index"))

def style_sheet(ws):
    fill = PatternFill("solid", fgColor="D9EAF7")
    side = Side(style="thin", color="BFBFBF")
    border = Border(left=side, right=side, top=side, bottom=side)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.row >= 4:
                cell.border = border
    for cell in ws[4]:
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def create_excel_report(objects, contracts, service_requests):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"protep_extended_report_{timestamp}.xlsx"
    file_path = TEMP_DIR / file_name

    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"

    total_area = sum(float(o.get("area") or 0) for o in objects)
    total_payment = sum(float(c.get("monthly_payment") or 0) for c in contracts)
    active_contracts = len([c for c in contracts if c.get("status") == "Действует"])
    open_requests = len([r for r in service_requests if r.get("status") != "Закрыта"])

    ws["A1"] = "Сводный отчёт по управлению недвижимым имуществом АО «ПРОТЭП»"
    ws["A1"].font = Font(size=14, bold=True)
    ws.merge_cells("A1:D1")

    rows = [
        ["Дата формирования", datetime.now().strftime("%d.%m.%Y %H:%M")],
        ["Количество объектов", len(objects)],
        ["Общая площадь, м²", total_area],
        ["Количество договоров", len(contracts)],
        ["Действующих договоров", active_contracts],
        ["Плановый месячный доход, руб.", total_payment],
        ["Количество заявок", len(service_requests)],
        ["Открытых заявок", open_requests],
    ]

    for idx, row in enumerate(rows, 3):
        ws[f"A{idx}"] = row[0]
        ws[f"B{idx}"] = row[1]

    ws_o = wb.create_sheet("Объекты")
    ws_o.append(["№", "Название", "Адрес", "Площадь, м²", "Статус", "Ответственный", "Дата добавления"])
    ws_o.insert_rows(1, 3)
    ws_o["A1"] = "Объекты недвижимости"
    ws_o["A1"].font = Font(size=14, bold=True)
    ws_o.merge_cells("A1:G1")

    for i, o in enumerate(objects, 1):
        ws_o.append([
            i,
            o.get("name",""),
            o.get("address",""),
            float(o.get("area") or 0),
            o.get("status",""),
            o.get("responsible_person",""),
            (o.get("created_at") or "")[:10]
        ])

    ws_c = wb.create_sheet("Договоры")
    ws_c.append(["№", "Объект", "Арендатор", "№ договора", "Платёж, руб.", "Дата начала", "Дата окончания", "Статус"])
    ws_c.insert_rows(1, 3)
    ws_c["A1"] = "Договоры"
    ws_c["A1"].font = Font(size=14, bold=True)
    ws_c.merge_cells("A1:H1")

    for i, c in enumerate(contracts, 1):
        obj = c.get("objects") or {}
        ws_c.append([
            i,
            obj.get("name",""),
            c.get("tenant_name",""),
            c.get("contract_number",""),
            float(c.get("monthly_payment") or 0),
            c.get("start_date",""),
            c.get("end_date",""),
            c.get("status","")
        ])

    ws_r = wb.create_sheet("Заявки")
    ws_r.append(["№", "Объект", "Тема", "Описание", "Приоритет", "Статус", "Дата создания"])
    ws_r.insert_rows(1, 3)
    ws_r["A1"] = "Заявки на обслуживание"
    ws_r["A1"].font = Font(size=14, bold=True)
    ws_r.merge_cells("A1:G1")

    for i, r in enumerate(service_requests, 1):
        obj = r.get("objects") or {}
        ws_r.append([
            i,
            obj.get("name",""),
            r.get("title",""),
            r.get("description",""),
            r.get("priority",""),
            r.get("status",""),
            (r.get("created_at") or "")[:10]
        ])

    for sheet in [ws, ws_o, ws_c, ws_r]:
        style_sheet(sheet)
        for col_idx in range(1, sheet.max_column + 1):
            letter = get_column_letter(col_idx)
            sheet.column_dimensions[letter].width = 24

    wb.save(file_path)
    return file_name, file_path

def upload_report(file_name, file_path):
    client = db()
    storage_path = f"reports/{file_name}"
    with open(file_path, "rb") as f:
        client.storage.from_(SUPABASE_BUCKET).upload(
            path=storage_path,
            file=f.read(),
            file_options={
                "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upsert": "true"
            },
        )
    return client.storage.from_(SUPABASE_BUCKET).get_public_url(storage_path)

@app.route("/reports/generate", methods=["POST"])
@admin_required
def generate_report():
    try:
        objects = fetch_table("objects")
        contracts = fetch_contracts()
        requests_data = fetch_requests()

        if not objects:
            flash("Нельзя сформировать отчёт: список объектов пуст.", "error")
            return redirect(url_for("index"))

        file_name, file_path = create_excel_report(objects, contracts, requests_data)
        file_url = upload_report(file_name, file_path)

        db().table("reports").insert({
            "file_name": file_name,
            "file_url": file_url,
            "report_type": "Расширенный отчёт",
            "created_by": session.get("username"),
        }).execute()

        flash("Расширенный Excel-отчёт сформирован и загружен в облако.", "success")

    except Exception as error:
        flash(f"Ошибка при формировании отчёта: {error}", "error")

    return redirect(url_for("index"))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
