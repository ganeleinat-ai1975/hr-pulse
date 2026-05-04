"""
Generate ~1000 demo employee records into the SQLite DB.
Idempotent: skips if >100 rows already exist (unless --force flag passed).
Usage:
    python generate_demo.py
    python generate_demo.py --force
"""

import sys
import os
import sqlite3
import random
import numpy as np
from datetime import date, timedelta

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import db as database
from models.flight_risk import calculate_flight_risk_score, get_risk_level

# ── Seed ──────────────────────────────────────────────────────────────────────
np.random.seed(42)
random.seed(42)

# ── Department / Manager map ──────────────────────────────────────────────────
DEPARTMENTS = {
    "פיתוח מוצר": ["יואב כהן", "שירה לוי"],
    "שיווק": ["רועי גולן", "נועה ברק"],
    "מכירות": ["אסף מזרחי", "הדס שפירא", "עמית דוד"],
    "פעילות": ["גלית חדד", "רן אברהם"],
    "כספים": ["מיכל פרידמן"],
    "משאבי אנוש": ["תמר גרינברג"],
    "שירות לקוחות": ["עידו שמחון", "לימור כץ"],
    "מחקר ופיתוח": ["בן שלום", "אורן רייך"],
}

DEPT_SIZES = {
    "פיתוח מוצר": 180,
    "שיווק": 90,
    "מכירות": 200,
    "פעילות": 150,
    "כספים": 70,
    "משאבי אנוש": 40,
    "שירות לקוחות": 180,
    "מחקר ופיתוח": 100,
}

# ── Name pools ────────────────────────────────────────────────────────────────
FIRST_NAMES = [
    "אביגיל", "אביב", "אבינועם", "אדם", "אהרון", "אודיה", "אור", "אורי", "אורית", "אורן",
    "אורנה", "אז", "איל", "אילן", "אילנה", "אייל", "אימן", "אינה", "אלה", "אלון",
    "אלונה", "אלי", "אליאב", "אליהו", "אלינה", "אלכסנדר", "אמיר", "אמירה", "אנה", "אסף",
    "אפרת", "ארז", "בן", "בנימין", "גל", "גלי", "גלית", "דוד", "דינה", "הדס",
    "הילה", "ויקטוריה", "זיו", "חן", "חנה", "טל", "יאיר", "יובל", "יוסף", "יחיאל",
    "ינון", "יעל", "יפית", "יצחק", "יקיר", "לאה", "לי", "ליאור", "לימור", "מאיר",
    "מיה", "מיכל", "מירי", "משה", "נדב", "נועה", "נועם", "נטע", "נילי", "ניר",
    "נעמה", "עדי", "עומר", "עמית", "עמליה", "פלג", "צחי", "רוני", "רועי", "רחל",
    "רינה", "רן", "שגיא", "שי", "שיר", "שירה", "שלום", "שמואל", "שרון", "תמר",
]

LAST_NAMES = [
    "אברהם", "אדלר", "אוחיון", "אלון", "אלוני", "אפרימוב", "ארביב", "בוזגלו", "בן-דוד", "ברק",
    "גוטמן", "גולן", "גרינברג", "דוד", "דויטש", "הורוביץ", "וינשטיין", "זכריה", "חדד", "כהן",
    "כץ", "כרמי", "לוי", "לוין", "מזרחי", "מלכה", "מנשה", "מרקוביץ", "משה", "נחמני",
    "סבן", "עמר", "פרידמן", "פרץ", "צברי", "ראובן", "רוזנברג", "רייך", "שושן", "שחר",
    "שטיין", "שלום", "שמחון", "שפירא", "תבורי",
]

# ── Feedback pools ────────────────────────────────────────────────────────────
POSITIVE_FEEDBACK = [
    "אני נהנה מאוד מהעבודה עם הצוות שלי, יש כאן אנשים מדהימים",
    "המנהל שלי תמיד זמין ומקשיב, מרגיש שאכפת לו",
    "יש הזדמנויות התפתחות טובות ואני מרגיש שאני גדל",
    "תרבות ארגונית טובה, שקיפות מהנהלה",
    "אני גאה להיות חלק מהארגון הזה",
    "האיזון בין עבודה לחיים אישיים מצוין כאן",
    "קיבלתי העלאה הוגנת השנה, מרגיש מוערך",
    "הפרויקטים מעניינים ומאתגרים, לא משעמם",
]

NEGATIVE_FEEDBACK = [
    "השכר שלי לא תואם לשוק, בדקתי והפרש של 20-30 אחוז",
    "המנהל לא מקשיב להצעות שלנו, פוגש אותנו פעם בחודש",
    "אין מסלול קידום ברור, לא יודע לאן אני הולך",
    "עומס עבודה לא סביר, עובד עד 10 בלילה כל יום",
    "חוסר הערכה לעבודה שאני עושה, אף פעם לא מקבל פידבק",
    "שקיפות נמוכה מאוד מהנהלה בכירה",
    "הכשרות לא רלוונטיות ומיושנות",
    "תרבות של האשמות ולא פתרון בעיות",
    "שיתוף פעולה ירוד בין מחלקות",
    "המנהל מקיים פגישות צוות רק בלחץ, אין תקשורת שוטפת",
    "מרגיש תקוע בתפקיד, אין אתגר",
    "שמועות על פיטורים גורמות לחוסר ביטחון",
    "יחס לא שוויוני בין עובדים",
    "מחשבות על עזיבה הולכות וגוברות לאחרונה",
    "תגמול על ביצועים מעולים לא קיים פה",
]

NEUTRAL_FEEDBACK = [
    "בסך הכל בסדר, יש עליות וירידות",
    "יש דברים טובים ויש מקום לשיפור",
    "עובד כאן כבר 3 שנים, מרגיש ממוצע",
    "לא רע אבל גם לא מדהים",
    "מחפש שינוי אבל לא בדחיפות",
    "הצוות טוב, הנהלה פחות",
]

# ── Department bias (mean offset from baseline) ───────────────────────────────
# "Problem" dept = שירות לקוחות, "Star" = מחקר ופיתוח
DEPT_BIAS = {
    "פיתוח מוצר": {"salary": 0.1, "development": 0.2, "manager": 0.1, "balance": 0.0},
    "שיווק": {"salary": 0.0, "development": 0.0, "manager": 0.0, "balance": -0.1},
    "מכירות": {"salary": 0.2, "development": -0.1, "manager": 0.0, "balance": -0.3},
    "פעילות": {"salary": -0.1, "development": -0.1, "manager": -0.1, "balance": -0.2},
    "כספים": {"salary": 0.1, "development": 0.0, "manager": 0.1, "balance": 0.1},
    "משאבי אנוש": {"salary": 0.0, "development": 0.1, "manager": 0.2, "balance": 0.2},
    "שירות לקוחות": {"salary": -0.4, "development": -0.3, "manager": -0.3, "balance": -0.5},  # problem dept
    "מחקר ופיתוח": {"salary": 0.3, "development": 0.5, "manager": 0.3, "balance": 0.2},    # star dept
}

# ── Manager bias (on top of dept bias) ───────────────────────────────────────
# Deliberately make some managers "bad" (bottom-5 candidates)
MANAGER_BIAS = {
    # Good managers
    "בן שלום": {"salary": 0.1, "development": 0.3, "manager": 0.5, "balance": 0.2},
    "אורן רייך": {"salary": 0.1, "development": 0.2, "manager": 0.4, "balance": 0.1},
    "תמר גרינברג": {"salary": 0.0, "development": 0.2, "manager": 0.4, "balance": 0.3},
    "מיכל פרידמן": {"salary": 0.1, "development": 0.1, "manager": 0.3, "balance": 0.1},
    # Neutral
    "יואב כהן": {"salary": 0.0, "development": 0.1, "manager": 0.1, "balance": 0.0},
    "שירה לוי": {"salary": 0.0, "development": 0.0, "manager": 0.0, "balance": 0.0},
    "רועי גולן": {"salary": 0.0, "development": 0.0, "manager": 0.1, "balance": 0.0},
    "נועה ברק": {"salary": 0.1, "development": 0.1, "manager": 0.1, "balance": 0.1},
    "הדס שפירא": {"salary": 0.1, "development": 0.0, "manager": 0.0, "balance": -0.1},
    "עמית דוד": {"salary": 0.1, "development": -0.1, "manager": -0.1, "balance": -0.2},
    "גלית חדד": {"salary": -0.1, "development": -0.1, "manager": -0.1, "balance": -0.1},
    "רן אברהם": {"salary": -0.1, "development": -0.2, "manager": -0.2, "balance": -0.1},
    # Bad managers (bottom-5 candidates)
    "עידו שמחון": {"salary": -0.5, "development": -0.4, "manager": -0.8, "balance": -0.6},
    "לימור כץ": {"salary": -0.4, "development": -0.3, "manager": -0.7, "balance": -0.5},
    "אסף מזרחי": {"salary": -0.2, "development": -0.3, "manager": -0.6, "balance": -0.3},
}


def random_name(used: set) -> str:
    """Generate a unique Hebrew full name."""
    for _ in range(200):
        name = f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"
        if name not in used:
            used.add(name)
            return name
    return f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"


def clamp(v, lo, hi):
    return max(lo, min(hi, v))


def random_date(start_year=2015, end_year=2024) -> str:
    start = date(start_year, 1, 1)
    end = date(end_year, 12, 31)
    delta = (end - start).days
    return (start + timedelta(days=random.randint(0, delta))).isoformat()


def generate_employee(dept: str, manager: str, used_names: set) -> dict:
    db = DEPT_BIAS[dept]
    mb = MANAGER_BIAS.get(manager, {k: 0 for k in ["salary", "development", "manager", "balance"]})

    baseline_sat = 3.2
    baseline_turnover = 2.5
    noise = 0.6  # std dev

    salary = clamp(np.random.normal(baseline_sat + db["salary"] + mb["salary"], noise), 1.0, 5.0)
    development = clamp(np.random.normal(baseline_sat + db["development"] + mb["development"], noise), 1.0, 5.0)
    manager_sat = clamp(np.random.normal(baseline_sat + db["manager"] + mb["manager"], noise), 1.0, 5.0)
    balance = clamp(np.random.normal(baseline_sat + db["balance"] + mb["balance"], noise), 1.0, 5.0)

    # Turnover intent inversely correlated with avg satisfaction
    avg_sat = (salary + development + manager_sat + balance) / 4
    turnover_mean = 6.0 - avg_sat  # higher satisfaction → lower intent
    turnover = clamp(round(np.random.normal(turnover_mean, 0.8)), 1, 5)

    # eNPS correlated with avg satisfaction
    enps_mean = (avg_sat - 3.0) * 40  # map 1-5 roughly to -80..+80
    enps = clamp(round(np.random.normal(enps_mean, 25)), -100, 100)

    # Feedback: higher risk → more negative
    risk_proxy = turnover  # 1-5
    if risk_proxy >= 4:
        feedback = random.choice(NEGATIVE_FEEDBACK)
    elif risk_proxy <= 2:
        feedback = random.choice(POSITIVE_FEEDBACK)
    else:
        pool = NEGATIVE_FEEDBACK + NEUTRAL_FEEDBACK + POSITIVE_FEEDBACK
        weights = [0.2] * len(NEGATIVE_FEEDBACK) + [0.5] * len(NEUTRAL_FEEDBACK) + [0.3] * len(POSITIVE_FEEDBACK)
        feedback = random.choices(pool, weights=weights, k=1)[0]

    row = {
        "name": random_name(used_names),
        "department": dept,
        "manager_name": manager,
        "hire_date": random_date(),
        "salary_satisfaction": round(salary, 2),
        "development_satisfaction": round(development, 2),
        "manager_satisfaction": round(manager_sat, 2),
        "balance_satisfaction": round(balance, 2),
        "turnover_intent": int(turnover),
        "enps": int(enps),
        "open_feedback": feedback,
    }
    from models.flight_risk import calculate_flight_risk_score, get_risk_level
    row["flight_risk_score"] = calculate_flight_risk_score(row)
    row["risk_level"] = get_risk_level(row["flight_risk_score"])
    return row


def main(force: bool = False):
    database.init_db()

    # Check if already populated
    conn = sqlite3.connect(database.get_db_path())
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM employees")
    existing = c.fetchone()[0]
    conn.close()

    if existing > 100 and not force:
        print(f"DB already has {existing} rows. Use --force to regenerate.")
        return

    if force:
        conn = sqlite3.connect(database.get_db_path())
        conn.execute("DELETE FROM employees")
        conn.commit()
        conn.close()
        print("Cleared existing data.")

    used_names: set = set()
    rows = []

    for dept, managers in DEPARTMENTS.items():
        total = DEPT_SIZES[dept]
        per_manager = total // len(managers)
        remainder = total % len(managers)

        for i, mgr in enumerate(managers):
            count = per_manager + (1 if i < remainder else 0)
            for _ in range(count):
                rows.append(generate_employee(dept, mgr, used_names))

    # Insert all rows
    conn = sqlite3.connect(database.get_db_path())
    c = conn.cursor()
    for row in rows:
        c.execute("""
            INSERT INTO employees
            (name, department, manager_name, hire_date,
             salary_satisfaction, development_satisfaction, manager_satisfaction, balance_satisfaction,
             turnover_intent, enps, open_feedback, flight_risk_score, risk_level)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            row["name"], row["department"], row["manager_name"], row["hire_date"],
            row["salary_satisfaction"], row["development_satisfaction"],
            row["manager_satisfaction"], row["balance_satisfaction"],
            row["turnover_intent"], row["enps"], row["open_feedback"],
            row["flight_risk_score"], row["risk_level"],
        ))
    conn.commit()
    conn.close()

    print(f"Generated {len(rows)} employee records.")
    # Print a quick summary
    dept_counts = {}
    risk_counts = {"גבוה": 0, "בינוני": 0, "נמוך": 0}
    for r in rows:
        dept_counts[r["department"]] = dept_counts.get(r["department"], 0) + 1
        risk_counts[r["risk_level"]] += 1
    print("\nDistribution by department:")
    for d, cnt in dept_counts.items():
        print(f"  {d}: {cnt}")
    print(f"\nRisk levels: גבוה={risk_counts['גבוה']}, בינוני={risk_counts['בינוני']}, נמוך={risk_counts['נמוך']}")


if __name__ == "__main__":
    force = "--force" in sys.argv
    main(force=force)
