import os
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, redirect, render_template, request, url_for, flash
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from supabase import create_client, Client


load_dotenv()

app = Flask(__name__)
app.secret_key = "dev-secret-key-change-later"

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "reports")

TEMP_DIR = Path("generated_reports")
TEMP_DIR.mkdir(exist_ok=True)


def get_supabase_client() -> Client:
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError(
            "Не заполнены SUPABASE_URL и SUPABASE_KEY. "
            "Создайте файл .env на основе .env.example."
        )
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def fetch_objects():
    supabase = get_supabase_client()
    result = (
        supabase.table("objects")
        .select("*")
        .order("created_at", desc=True)
        .execute()
    )
    return result.data


def fetch_reports():
    supabase = get_supabase_client()
    result = (
        supabase.table("reports")
        .select("*")
        .order("created_at", desc=True)
        .execute()
    )
    return result.data


def create_excel_report(objects):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"protep_objects_report_{timestamp}.xlsx"
    file_path = TEMP_DIR / file_name

    wb = Workbook()
    ws = wb.active
    ws.title = "Отчёт по объектам"

    ws.merge_cells("A1:F1")
    title_cell = ws["A1"]
    title_cell.value = "Отчёт по объектам недвижимости АО «ПРОТЭП»"
    title_cell.font = Font(size=14, bold=True)
    title_cell.alignment = Alignment(horizontal="center")

    ws["A2"] = f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws.merge_cells("A2:F2")

    headers = ["№", "Название объекта", "Адрес", "Площадь, м²", "Статус", "Дата добавления"]
    ws.append([])
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[4]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for index, obj in enumerate(objects, start=1):
        created_at = obj.get("created_at", "")
        ws.append([
            index,
            obj.get("name", ""),
            obj.get("address", ""),
            float(obj.get("area", 0) or 0),
            obj.get("status", ""),
            created_at[:10] if created_at else "",
        ])

    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    widths = {
        "A": 6,
        "B": 28,
        "C": 42,
        "D": 14,
        "E": 18,
        "F": 18,
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    wb.save(file_path)
    return file_name, file_path


def upload_report_to_storage(file_name, file_path):
    supabase = get_supabase_client()

    with open(file_path, "rb") as file:
        file_bytes = file.read()

    storage_path = f"reports/{file_name}"

    supabase.storage.from_(SUPABASE_BUCKET).upload(
        path=storage_path,
        file=file_bytes,
        file_options={
            "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "upsert": "true",
        },
    )

    public_url = supabase.storage.from_(SUPABASE_BUCKET).get_public_url(storage_path)
    return public_url


@app.route("/")
def index():
    try:
        objects = fetch_objects()
        reports = fetch_reports()
    except Exception as error:
        objects = []
        reports = []
        flash(str(error), "error")

    return render_template("index.html", objects=objects, reports=reports)


@app.route("/objects", methods=["POST"])
def add_object():
    name = request.form.get("name", "").strip()
    address = request.form.get("address", "").strip()
    area = request.form.get("area", "").strip()
    status = request.form.get("status", "").strip()

    if not name or not address or not area or not status:
        flash("Заполните все поля объекта.", "error")
        return redirect(url_for("index"))

    try:
        area_value = float(area.replace(",", "."))
        supabase = get_supabase_client()
        supabase.table("objects").insert({
            "name": name,
            "address": address,
            "area": area_value,
            "status": status,
        }).execute()
        flash("Объект недвижимости добавлен.", "success")
    except Exception as error:
        flash(f"Ошибка при добавлении объекта: {error}", "error")

    return redirect(url_for("index"))


@app.route("/reports/generate", methods=["POST"])
def generate_report():
    try:
        objects = fetch_objects()

        if not objects:
            flash("Нельзя сформировать отчёт: список объектов пуст.", "error")
            return redirect(url_for("index"))

        file_name, file_path = create_excel_report(objects)
        file_url = upload_report_to_storage(file_name, file_path)

        supabase = get_supabase_client()
        supabase.table("reports").insert({
            "file_name": file_name,
            "file_url": file_url,
        }).execute()

        flash("Excel-отчёт сформирован и загружен в облако.", "success")
    except Exception as error:
        flash(f"Ошибка при формировании отчёта: {error}", "error")

    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
