import sqlite3
from pathlib import Path
DB = Path(__file__).parent / 'data' / 'excelflow.db'

def save_run(summary, rows):
    DB.parent.mkdir(exist_ok=True)
    with sqlite3.connect(DB) as conn:
        conn.execute('CREATE TABLE IF NOT EXISTS report_runs (id INTEGER PRIMARY KEY, created_at TEXT DEFAULT CURRENT_TIMESTAMP, rows INTEGER NOT NULL, revenue REAL NOT NULL, gross_profit REAL NOT NULL)')
        conn.execute('INSERT INTO report_runs(rows,revenue,gross_profit) VALUES (?,?,?)', (rows,summary['revenue'],summary['gross_profit']))
