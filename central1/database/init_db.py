import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "tracker.db")

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.executescript("""
        DROP TABLE IF EXISTS time_records;
        DROP TABLE IF EXISTS users;
        DROP TABLE IF EXISTS projects;

        CREATE TABLE users (
            user_id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now', '+3 hours'))
        );

        CREATE TABLE projects (
            project_id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL
        );

        CREATE TABLE time_records (
            record_id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            project_id INTEGER,
            date TEXT NOT NULL,
            start_time TEXT,
            end_time TEXT,
            FOREIGN KEY (user_id) REFERENCES users(user_id),
            FOREIGN KEY (project_id) REFERENCES projects(project_id)
        );
    """)

    # Тестовые сотрудники
    users = ["Иванов И.И.", "Петров П.П."]
    for u in users:
        cursor.execute("INSERT INTO users (full_name) VALUES (?)", (u,))

    # Тестовые проекты
    projects = ["Разработка бота", "Вёрстка админки", "Тестирование", "Совещание"]
    for p in projects:
        cursor.execute("INSERT INTO projects (name) VALUES (?)", (p,))

    conn.commit()
    conn.close()
    print("БД пересоздана: таблицы, 2 сотрудника, 4 проекта.")

if __name__ == "__main__":
    init_db()