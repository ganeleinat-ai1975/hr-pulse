# HR Pulse — Developer Skill

Use this skill when working on the HR Pulse project. It gives Claude full context to help you develop, fix, and extend the app.

## הפעלה
פשוט כתבי `/hr-pulse-dev` בצ'אט עם Claude Code כשאת עובדת על הפרויקט.

---

## Context שנטען אוטומטית

### Project
- **שם:** HR Pulse — מערכת ניתוח שביעות רצון עובדים
- **Live:** https://hr-pulse-7lhc9hqsoefettnf8mikhe.streamlit.app
- **GitHub:** https://github.com/ganeleinat-ai1975/hr-pulse
- **Stack:** Python + Streamlit + SQLite + Plotly + pandas + python-docx

### קבצים מרכזיים
| קובץ | תפקיד |
|------|--------|
| `app.py` | כל האפליקציה — sidebar nav + 8 דפים כ-functions |
| `db.py` | CRUD + aggregations ל-SQLite |
| `generate_demo.py` | יצירת 1010 עובדי דמו (idempotent) |
| `models/flight_risk.py` | ניקוד 0-100 לכל עובד |
| `exporters/word_report.py` | דוח Word RTL |
| `exporters/coaching.py` | מכתבי coaching ל-Bottom 5 מנהלים |

### DB Schema — טבלת `employees`
```sql
id, name, department, manager_name, hire_date,
salary_satisfaction REAL,      -- 1-5
development_satisfaction REAL,  -- 1-5
manager_satisfaction REAL,      -- 1-5
balance_satisfaction REAL,      -- 1-5
turnover_intent INTEGER,        -- 1-5
enps INTEGER,                   -- -100 to 100
open_feedback TEXT,
flight_risk_score REAL,         -- 0-100
risk_level TEXT                 -- גבוה/בינוני/נמוך
```

---

## הרצה מקומית
```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python generate_demo.py
.venv/bin/python -m streamlit run app.py
# → http://localhost:8501
```

## Deploy לאחר שינויים
```bash
git add -A
git commit -m "תיאור השינוי"
git push origin main
# Streamlit Cloud מתעדכן אוטומטית תוך ~2 דקות
```

## הוספת דף חדש
1. כתבי `def page_שם():` ב-`app.py`
2. הוסיפי לרשימת `PAGES` ב-sidebar
3. הוסיפי `elif page == "שם": page_שם()`

## כללים חשובים
- ה-UI כולו בעברית RTL — אל תוסיפי טקסט באנגלית גלוי למשתמש
- DB path תמיד דרך `database.get_db_path()` — לא hard-coded
- אחרי שינויים ב-`app.py`: `python -c "import ast; ast.parse(open('app.py').read())"`
- אל תוסיפי תלויות כבדות (scikit-learn, transformers וכד')
