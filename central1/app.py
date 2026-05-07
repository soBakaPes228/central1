from fastapi import FastAPI, Request, Form, Query
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from datetime import datetime
import os
from database.init_db import get_connection, init_db

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# --- При старте создаём БД ---
@app.on_event("startup")
def startup():
    init_db()

# --- Главная — сотрудник (заглушка вместо бота) ---
@app.get("/", response_class=HTMLResponse)
def employee_page(request: Request):
    conn = get_connection()
    users = conn.execute("SELECT * FROM users").fetchall()
    projects = conn.execute("SELECT * FROM projects").fetchall()
    conn.close()
    return templates.TemplateResponse("employee.html", {
        "request": request,
        "users": users,
        "projects": projects,
        "message": None
    })

@app.post("/employee/start", response_class=HTMLResponse)
def employee_start(
    request: Request,
    user_id: int = Form(...),
    project_id: int = Form(...)
):
    conn = get_connection()
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    time = now.strftime("%H:%M:%S")

    # Проверка на незакрытую запись
    open_record = conn.execute(
        "SELECT * FROM time_records WHERE user_id = ? AND end_time IS NULL AND date = ?",
        (user_id, today)
    ).fetchone()

    users = conn.execute("SELECT * FROM users").fetchall()
    projects = conn.execute("SELECT * FROM projects").fetchall()

    if open_record:
        conn.close()
        return templates.TemplateResponse("employee.html", {
            "request": request,
            "users": users,
            "projects": projects,
            "message": "❌ У вас уже есть незавершённая запись. Сначала завершите её."
        })

    conn.execute(
        "INSERT INTO time_records (user_id, project_id, date, start_time) VALUES (?, ?, ?, ?)",
        (user_id, project_id, today, time)
    )
    conn.commit()
    conn.close()

    return templates.TemplateResponse("employee.html", {
        "request": request,
        "users": users,
        "projects": projects,
        "message": f"✅ Начало работы зафиксировано: {time}"
    })

@app.post("/employee/stop", response_class=HTMLResponse)
def employee_stop(
    request: Request,
    user_id: int = Form(...)
):
    conn = get_connection()
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    time = now.strftime("%H:%M:%S")

    open_record = conn.execute(
        "SELECT * FROM time_records WHERE user_id = ? AND end_time IS NULL AND date = ?",
        (user_id, today)
    ).fetchone()

    users = conn.execute("SELECT * FROM users").fetchall()
    projects = conn.execute("SELECT * FROM projects").fetchall()

    if not open_record:
        conn.close()
        return templates.TemplateResponse("employee.html", {
            "request": request,
            "users": users,
            "projects": projects,
            "message": "❌ Нечего завершать. Вы ещё не начинали работу."
        })

    conn.execute(
        "UPDATE time_records SET end_time = ? WHERE record_id = ?",
        (time, open_record["record_id"])
    )
    conn.commit()
    conn.close()

    start = open_record["start_time"]
    return templates.TemplateResponse("employee.html", {
        "request": request,
        "users": users,
        "projects": projects,
        "message": f"✅ Работа завершена. Начало: {start}, Конец: {time}"
    })

# --- Вход для руководителя ---
@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})

# --- Дашборд руководителя ---
@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(
    request: Request,
    project_id: int = Query(None),
    date_from: str = Query(None),
    date_to: str = Query(None)
):
    conn = get_connection()

    query = """
        SELECT u.full_name, p.name as project, r.date, r.start_time, r.end_time
        FROM time_records r
        JOIN users u ON r.user_id = u.user_id
        LEFT JOIN projects p ON r.project_id = p.project_id
        WHERE 1=1
    """
    params = []

    if project_id:
        query += " AND r.project_id = ?"
        params.append(project_id)
    if date_from:
        query += " AND r.date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND r.date <= ?"
        params.append(date_to)

    query += " ORDER BY r.date DESC, r.start_time DESC"

    records = conn.execute(query, params).fetchall()
    projects = conn.execute("SELECT * FROM projects").fetchall()
    conn.close()

    return templates.TemplateResponse("dashboard.html", {
        "request": request,
        "records": records,
        "projects": projects,
        "selected_project": project_id,
        "date_from": date_from or "",
        "date_to": date_to or ""
    })

# --- Экспорт в Excel ---
@app.get("/export")
def export_excel(
    project_id: int = Query(None),
    date_from: str = Query(None),
    date_to: str = Query(None)
):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment

    conn = get_connection()
    query = """
        SELECT u.full_name, p.name as project, r.date, r.start_time, r.end_time
        FROM time_records r
        JOIN users u ON r.user_id = u.user_id
        LEFT JOIN projects p ON r.project_id = p.project_id
        WHERE 1=1
    """
    params = []

    if project_id:
        query += " AND r.project_id = ?"
        params.append(project_id)
    if date_from:
        query += " AND r.date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND r.date <= ?"
        params.append(date_to)

    query += " ORDER BY r.date DESC"
    records = conn.execute(query, params).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Отчёт"

    headers = ["ФИО", "Проект", "Дата", "Начало", "Конец"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)

    for row_idx, r in enumerate(records, 2):
        ws.cell(row=row_idx, column=1, value=r["full_name"])
        ws.cell(row=row_idx, column=2, value=r["project"] or "—")
        ws.cell(row=row_idx, column=3, value=r["date"])
        ws.cell(row=row_idx, column=4, value=r["start_time"] or "—")
        ws.cell(row=row_idx, column=5, value=r["end_time"] or "не завершено")

    filename = f"otchet_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    filepath = os.path.join(os.path.dirname(__file__), filename)
    wb.save(filepath)

    return FileResponse(filepath, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")