# HR Pulse — AI Agent Context

## WHAT — Project Overview

**Project:** HR Pulse — מערכת ניתוח שביעות רצון עובדים  
**Purpose:** לוח בקרה אינטראקטיבי לניתוח סקרי עובדים, זיהוי סיכוני עזיבה, וניתוח מנהלים  
**Live URL:** https://hr-pulse-7lhc9hqsoefettnf8mikhe.streamlit.app  
**GitHub:** https://github.com/ganeleinat-ai1975/hr-pulse  

### Tech Stack
- **Python 3.11** + **Streamlit ≥1.32** — UI ו-server
- **SQLite** (via `sqlite3`) — DB מקומי, קובץ `data/hr_pulse.db`
- **Pandas ≥2.0** — עיבוד נתונים
- **Plotly ≥5.18** — גרפים אינטראקטיביים
- **python-docx ≥1.1** — יצוא דוח Word RTL
- **openpyxl ≥3.1** — ייבוא/יצוא Excel
- **NumPy ≥1.26** — חישובי Flight Risk

### Repository Structure
```
hr-pulse/
├── app.py                  ← נקודת כניסה, כל הדפים כ-functions
├── db.py                   ← כל פעולות SQLite (CRUD + aggregations)
├── generate_demo.py        ← יצירת ~1010 עובדי דמו (idempotent)
├── requirements.txt
├── models/
│   └── flight_risk.py      ← ניקוד סיכון 0-100
├── exporters/
│   ├── word_report.py      ← דוח Word RTL לדירקטוריון
│   └── coaching.py         ← מכתבי Coaching ל-Bottom 5 מנהלים
└── data/
    └── hr_pulse.db         ← SQLite (מחוץ ל-git)
```

### Key Data Schema — `employees` table
| עמודה | טיפוס | הערה |
|-------|--------|------|
| salary/development/manager/balance_satisfaction | REAL | סקלה 1–5 |
| turnover_intent | INTEGER | 1–5 (5 = בוודאות עוזב) |
| enps | INTEGER | -100 עד 100 |
| flight_risk_score | REAL | 0–100, מחושב |
| risk_level | TEXT | גבוה / בינוני / נמוך |

---

## WHY — Architecture Decisions

- **Single-file app.py** במקום pages/ — מאפשר שיתוף state בין דפים ללא `st.session_state` מורכב
- **SQLite ולא PostgreSQL** — מתאים ל-localhost + Streamlit Cloud; קל לגיבוי ולשיתוף DB
- **Auto-seed בהפעלה** — `generate_demo.py` רץ אוטומטית אם הטבלה ריקה (חשוב ל-Streamlit Cloud שאינו שומר קבצים)
- **RTL CSS גלובלי** — מוזרק דרך `st.markdown` עם Heebo font מ-Google Fonts
- **Flight Risk Score** — ניקוד משוקלל (turnover_intent 40%, eNPS 20%, שכר 15%, פיתוח 15%, מנהל 10%)
- **ללא AI/ML מורכב** — ניתוח טקסט מבוסס keyword-matching ולא NLP, להשאיר פשוט ויציב

### Anti-patterns to Avoid
- אל תשנה את מסלול ה-DB — תמיד `os.path.join(os.path.dirname(__file__), "data", "hr_pulse.db")`
- אל תשתמש ב-`st.experimental_rerun()` — deprecated, השתמש ב-`st.rerun()`
- אל תוסיף תלויות כבדות (scikit-learn, transformers) — לא נדרש ויאט deploy

---

## HOW — Workflows

### הרצה מקומית
```bash
cd "/Users/einatganel/Documents/claude projects/hr-pulse"
.venv/bin/python generate_demo.py          # פעם אחת
.venv/bin/python -m streamlit run app.py
```

### התקנת סביבה חדשה
```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python generate_demo.py
.venv/bin/python -m streamlit run app.py
```

### Deploy ל-Streamlit Cloud
```bash
git add -A && git commit -m "תיאור השינוי"
git push origin main
# Streamlit Cloud מתעדכן אוטומטית תוך ~2 דקות
```

### הוספת דף חדש ל-app.py
1. הוסף פונקציה `def page_xxx():`
2. הוסף לרשימת `PAGES` ב-sidebar
3. הוסף `elif page == "שם הדף": page_xxx()`

### בדיקת syntax לפני push
```bash
.venv/bin/python -c "import ast; ast.parse(open('app.py').read()); print('OK')"
```
