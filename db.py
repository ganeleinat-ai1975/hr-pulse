import os
import sqlite3
import pandas as pd


def get_db_path() -> str:
    data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, "hr_pulse.db")


def init_db():
    """Create tables if they don't exist."""
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            department TEXT NOT NULL,
            manager_name TEXT NOT NULL,
            hire_date TEXT,
            salary_satisfaction REAL,
            development_satisfaction REAL,
            manager_satisfaction REAL,
            balance_satisfaction REAL,
            turnover_intent INTEGER,
            enps INTEGER,
            open_feedback TEXT,
            flight_risk_score REAL,
            risk_level TEXT
        )
    """)
    conn.commit()
    conn.close()


def get_all_employees() -> pd.DataFrame:
    conn = sqlite3.connect(get_db_path())
    df = pd.read_sql_query("SELECT * FROM employees ORDER BY id", conn)
    conn.close()
    return df


def get_employees_by_department(dept: str) -> pd.DataFrame:
    conn = sqlite3.connect(get_db_path())
    df = pd.read_sql_query(
        "SELECT * FROM employees WHERE department = ? ORDER BY id",
        conn, params=(dept,)
    )
    conn.close()
    return df


def get_employees_by_manager(manager: str) -> pd.DataFrame:
    conn = sqlite3.connect(get_db_path())
    df = pd.read_sql_query(
        "SELECT * FROM employees WHERE manager_name = ? ORDER BY id",
        conn, params=(manager,)
    )
    conn.close()
    return df


def get_manager_stats() -> pd.DataFrame:
    conn = sqlite3.connect(get_db_path())
    df = pd.read_sql_query("""
        SELECT
            manager_name,
            department,
            COUNT(*) as employee_count,
            ROUND(AVG(salary_satisfaction), 2) as avg_salary,
            ROUND(AVG(development_satisfaction), 2) as avg_development,
            ROUND(AVG(manager_satisfaction), 2) as avg_manager,
            ROUND(AVG(balance_satisfaction), 2) as avg_balance,
            ROUND(AVG(enps), 1) as avg_enps,
            ROUND(AVG(flight_risk_score), 1) as avg_flight_risk,
            ROUND(AVG((salary_satisfaction + development_satisfaction + manager_satisfaction + balance_satisfaction) / 4.0), 2) as avg_overall,
            SUM(CASE WHEN risk_level = 'גבוה' THEN 1 ELSE 0 END) as high_risk_count,
            ROUND(100.0 * SUM(CASE WHEN risk_level = 'גבוה' THEN 1 ELSE 0 END) / COUNT(*), 1) as high_risk_pct,
            ROUND(
                (AVG(manager_satisfaction) * 0.5 + AVG(development_satisfaction) * 0.3 + AVG(balance_satisfaction) * 0.2),
                2
            ) as engagement_score
        FROM employees
        GROUP BY manager_name
        ORDER BY engagement_score DESC
    """, conn)
    conn.close()
    return df


def get_department_stats() -> pd.DataFrame:
    conn = sqlite3.connect(get_db_path())
    df = pd.read_sql_query("""
        SELECT
            department,
            COUNT(*) as employee_count,
            ROUND(AVG(salary_satisfaction), 2) as avg_salary,
            ROUND(AVG(development_satisfaction), 2) as avg_development,
            ROUND(AVG(manager_satisfaction), 2) as avg_manager,
            ROUND(AVG(balance_satisfaction), 2) as avg_balance,
            ROUND(AVG(enps), 1) as avg_enps,
            ROUND(AVG(flight_risk_score), 1) as avg_flight_risk,
            ROUND(AVG((salary_satisfaction + development_satisfaction + manager_satisfaction + balance_satisfaction) / 4.0), 2) as avg_overall,
            SUM(CASE WHEN risk_level = 'גבוה' THEN 1 ELSE 0 END) as high_risk_count,
            ROUND(100.0 * SUM(CASE WHEN risk_level = 'גבוה' THEN 1 ELSE 0 END) / COUNT(*), 1) as high_risk_pct,
            SUM(CASE WHEN risk_level = 'בינוני' THEN 1 ELSE 0 END) as mid_risk_count,
            SUM(CASE WHEN risk_level = 'נמוך' THEN 1 ELSE 0 END) as low_risk_count
        FROM employees
        GROUP BY department
        ORDER BY avg_overall DESC
    """, conn)
    conn.close()
    return df


def import_from_dataframe(df: pd.DataFrame):
    """Import employees from a dataframe (e.g., from Excel upload)."""
    from models.flight_risk import calculate_flight_risk_score, get_risk_level
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor()
    for _, row in df.iterrows():
        row_dict = row.to_dict()
        # Calculate flight risk if not present
        try:
            score = calculate_flight_risk_score(row_dict)
            level = get_risk_level(score)
        except Exception:
            score = None
            level = None
        c.execute("""
            INSERT INTO employees
            (name, department, manager_name, hire_date,
             salary_satisfaction, development_satisfaction, manager_satisfaction, balance_satisfaction,
             turnover_intent, enps, open_feedback, flight_risk_score, risk_level)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            row_dict.get('name'), row_dict.get('department'), row_dict.get('manager_name'),
            row_dict.get('hire_date'),
            row_dict.get('salary_satisfaction'), row_dict.get('development_satisfaction'),
            row_dict.get('manager_satisfaction'), row_dict.get('balance_satisfaction'),
            row_dict.get('turnover_intent'), row_dict.get('enps'),
            row_dict.get('open_feedback'), score, level
        ))
    conn.commit()
    conn.close()


def update_employee(emp_id: int, data: dict):
    """Update a single employee record."""
    from models.flight_risk import calculate_flight_risk_score, get_risk_level
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor()
    # Recalculate flight risk
    try:
        score = calculate_flight_risk_score(data)
        level = get_risk_level(score)
    except Exception:
        score = data.get('flight_risk_score')
        level = data.get('risk_level')
    c.execute("""
        UPDATE employees SET
            name=?, department=?, manager_name=?, hire_date=?,
            salary_satisfaction=?, development_satisfaction=?, manager_satisfaction=?,
            balance_satisfaction=?, turnover_intent=?, enps=?, open_feedback=?,
            flight_risk_score=?, risk_level=?
        WHERE id=?
    """, (
        data.get('name'), data.get('department'), data.get('manager_name'), data.get('hire_date'),
        data.get('salary_satisfaction'), data.get('development_satisfaction'),
        data.get('manager_satisfaction'), data.get('balance_satisfaction'),
        data.get('turnover_intent'), data.get('enps'), data.get('open_feedback'),
        score, level, emp_id
    ))
    conn.commit()
    conn.close()


def delete_employee(emp_id: int):
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor()
    c.execute("DELETE FROM employees WHERE id=?", (emp_id,))
    conn.commit()
    conn.close()


def get_summary_stats() -> dict:
    """Return KPIs for the overview dashboard."""
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM employees")
    total = c.fetchone()[0]

    c.execute("""
        SELECT
            AVG((salary_satisfaction + development_satisfaction + manager_satisfaction + balance_satisfaction) / 4.0),
            AVG(enps),
            SUM(CASE WHEN risk_level='גבוה' THEN 1 ELSE 0 END),
            AVG(salary_satisfaction),
            AVG(development_satisfaction),
            AVG(manager_satisfaction),
            AVG(balance_satisfaction)
        FROM employees
    """)
    row = c.fetchone()
    conn.close()
    high_risk_count = row[2] or 0
    return {
        "total_employees": total,
        "avg_overall": round(row[0] or 0, 2),
        "avg_enps": round(row[1] or 0, 1),
        "high_risk_count": int(high_risk_count),
        "high_risk_pct": round(100.0 * high_risk_count / total, 1) if total > 0 else 0,
        "avg_salary": round(row[3] or 0, 2),
        "avg_development": round(row[4] or 0, 2),
        "avg_manager": round(row[5] or 0, 2),
        "avg_balance": round(row[6] or 0, 2),
    }


def get_db_meta() -> dict:
    """Return metadata about the DB."""
    import os
    path = get_db_path()
    exists = os.path.exists(path)
    mtime = None
    if exists:
        import datetime
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
    conn = sqlite3.connect(path)
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM employees")
    count = c.fetchone()[0]
    conn.close()
    return {"path": path, "last_updated": mtime, "total_records": count}
