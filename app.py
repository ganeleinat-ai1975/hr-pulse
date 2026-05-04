import streamlit as st

st.set_page_config(
    page_title="HR Pulse — ניתוח עובדים",
    page_icon="💡",
    layout="wide",
    initial_sidebar_state="expanded"
)

import os
import sys
import io
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db as database

# ── RTL + Heebo CSS ──────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;600;700&display=swap');
* { font-family: 'Heebo', Arial, sans-serif !important; direction: rtl !important; }
.stApp { direction: rtl; }
[data-testid="stSidebar"] { direction: rtl; }
[data-testid="stSidebarNav"] { direction: rtl; }
div[data-testid="metric-container"] { direction: rtl; text-align: right; }
.stSelectbox label, .stMultiSelect label, .stSlider label { direction: rtl; text-align: right; }
.stDataFrame { direction: rtl; }

.kpi-card {
    background: #f8f9fa;
    border-right: 4px solid #4CAF50;
    padding: 16px 20px;
    border-radius: 8px;
    text-align: center;
    margin: 4px;
}
.kpi-number { font-size: 2rem; font-weight: 700; color: #1a73e8; }
.kpi-label { font-size: 0.9rem; color: #666; margin-top: 4px; }

.traffic-green { color: #2ecc71; font-size: 24px; }
.traffic-yellow { color: #f39c12; font-size: 24px; }
.traffic-red { color: #e74c3c; font-size: 24px; }

.rec-card {
    background: #fff;
    border-right: 4px solid #1a73e8;
    padding: 14px 18px;
    border-radius: 8px;
    margin-bottom: 12px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}
.rec-card-high { border-right-color: #ea4335; }
.rec-card-mid  { border-right-color: #fbbc04; }
.rec-card-low  { border-right-color: #34a853; }
.rec-title { font-weight: 700; font-size: 1.05rem; margin-bottom: 6px; }
.rec-detail { font-size: 0.92rem; color: #444; }
.rec-meta { font-size: 0.82rem; color: #888; margin-top: 6px; }

.risk-high { color: #ea4335; font-weight: 700; }
.risk-mid  { color: #fbbc04; font-weight: 700; }
.risk-low  { color: #34a853; font-weight: 700; }

h1, h2, h3 { direction: rtl; text-align: right; }
.block-container { padding-top: 1.5rem; }
</style>
""", unsafe_allow_html=True)

# ── Color palette ─────────────────────────────────────────────────────────────
C_BLUE   = "#1a73e8"
C_GREEN  = "#34a853"
C_YELLOW = "#fbbc04"
C_RED    = "#ea4335"
C_BG     = "#f8f9fa"

PLOTLY_FONT = dict(family="Heebo, Arial", size=13)

# ── Init DB ───────────────────────────────────────────────────────────────────
database.init_db()

# Auto-seed demo data on first run (e.g. Streamlit Cloud fresh deployment)
_stats = database.get_summary_stats()
if _stats["total_employees"] == 0:
    import generate_demo as _gd
    _gd.main(force=False)

# ── Sidebar navigation ────────────────────────────────────────────────────────
PAGES = [
    "🏠 תמונת מצב כוללת",
    "🏢 ניתוח מחלקות",
    "👤 ניתוח מנהלים",
    "⚠️ סיכון עזיבה",
    "💬 ניתוח תשובות פתוחות",
    "💡 המלצות אסטרטגיות",
    "📊 יצוא דוחות",
    "⚙️ ניהול נתונים",
]

with st.sidebar:
    st.markdown("## 💡 HR Pulse")
    st.markdown("---")
    page = st.radio("ניווט", PAGES, index=0, label_visibility="collapsed")
    st.markdown("---")
    st.caption("גרסה 1.0 | כל הזכויות שמורות")

# ── Helper functions ─────────────────────────────────────────────────────────

def risk_color(level: str) -> str:
    return {"גבוה": C_RED, "בינוני": C_YELLOW, "נמוך": C_GREEN}.get(level, "#999")


def score_color(score: float) -> str:
    if score >= 65: return C_RED
    if score >= 40: return C_YELLOW
    return C_GREEN


def kpi_card(label: str, value, suffix=""):
    return f"""
    <div class="kpi-card">
        <div class="kpi-number">{value}{suffix}</div>
        <div class="kpi-label">{label}</div>
    </div>"""


def traffic_light(avg: float):
    if avg >= 3.5:
        return "✅", "ירוק", "מצב תקין", "traffic-green"
    elif avg >= 3.0:
        return "⚠️", "צהוב", "דורש תשומת לב", "traffic-yellow"
    else:
        return "🔴", "אדום", "מצב דורש טיפול מיידי", "traffic-red"


def make_bar_chart(df, x_col, y_col, title, color_col=None, color_map=None, orientation='v'):
    fig = px.bar(
        df, x=x_col, y=y_col, title=title,
        color=color_col, color_discrete_map=color_map,
        orientation=orientation,
        color_discrete_sequence=[C_BLUE],
    )
    fig.update_layout(
        font=PLOTLY_FONT,
        plot_bgcolor=C_BG,
        paper_bgcolor="white",
        title_x=1.0,
        title_xanchor="right",
    )
    return fig


# ════════════════════════════════════════════════════════════════════════════
# PAGE 1: Overview
# ════════════════════════════════════════════════════════════════════════════
if page == "🏠 תמונת מצב כוללת":
    st.title("🏠 תמונת מצב כוללת")

    stats = database.get_summary_stats()
    total = stats['total_employees']

    if total == 0:
        st.warning("אין נתוני עובדים. עבור לדף ניהול נתונים וייבא נתונים או הפעל את generate_demo.py.")
        st.stop()

    # ── KPI row ───────────────────────────────────────────────────────────
    cols = st.columns(4)
    with cols[0]:
        st.markdown(kpi_card("סה\"כ עובדים", total), unsafe_allow_html=True)
    with cols[1]:
        st.markdown(kpi_card("ציון מחוברות כולל", stats['avg_overall'], "/5"), unsafe_allow_html=True)
    with cols[2]:
        st.markdown(kpi_card("eNPS ממוצע", stats['avg_enps']), unsafe_allow_html=True)
    with cols[3]:
        st.markdown(kpi_card("% בסיכון עזיבה גבוה", stats['high_risk_pct'], "%"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Traffic light ─────────────────────────────────────────────────────
    icon, color_name, label, css_class = traffic_light(stats['avg_overall'])
    st.markdown(f"""
    <div style="background:{C_BG}; border-radius:10px; padding:16px 24px; display:flex; align-items:center; gap:16px; margin-bottom:16px;">
        <span style="font-size:2.5rem;">{icon}</span>
        <div>
            <div style="font-size:1.2rem; font-weight:700;">רמזור מצב ארגוני: {label}</div>
            <div style="color:#666;">ציון מחוברות כולל: {stats['avg_overall']}/5</div>
        </div>
    </div>""", unsafe_allow_html=True)

    # ── Charts row ────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        sat_data = pd.DataFrame({
            "קטגוריה": ["שכר", "פיתוח", "מנהל", "איזון"],
            "ציון ממוצע": [stats['avg_salary'], stats['avg_development'],
                           stats['avg_manager'], stats['avg_balance']],
        })
        fig = px.bar(sat_data, x="קטגוריה", y="ציון ממוצע",
                     title="ציוני שביעות רצון ממוצעים לפי קטגוריה",
                     color="ציון ממוצע",
                     color_continuous_scale=[[0, C_RED], [0.5, C_YELLOW], [1, C_GREEN]],
                     range_color=[1, 5])
        fig.update_layout(font=PLOTLY_FONT, plot_bgcolor=C_BG, paper_bgcolor="white",
                          title_x=1.0, title_xanchor="right", yaxis_range=[0, 5])
        fig.update_coloraxes(showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        df = database.get_all_employees()
        risk_counts = df['risk_level'].value_counts().reset_index()
        risk_counts.columns = ['רמת סיכון', 'מספר עובדים']
        color_map = {"גבוה": C_RED, "בינוני": C_YELLOW, "נמוך": C_GREEN}
        fig2 = px.pie(risk_counts, names='רמת סיכון', values='מספר עובדים',
                      title="פילוח רמות סיכון עזיבה",
                      color='רמת סיכון', color_discrete_map=color_map,
                      hole=0.4)
        fig2.update_layout(font=PLOTLY_FONT, paper_bgcolor="white",
                           title_x=1.0, title_xanchor="right")
        st.plotly_chart(fig2, use_container_width=True)

    # ── Department comparison ─────────────────────────────────────────────
    dept_df = database.get_department_stats()
    if not dept_df.empty:
        dept_df_sorted = dept_df.sort_values("avg_overall", ascending=True)
        colors = [C_RED if v < 3.0 else C_YELLOW if v < 3.5 else C_GREEN
                  for v in dept_df_sorted["avg_overall"]]
        fig3 = go.Figure(go.Bar(
            x=dept_df_sorted["avg_overall"],
            y=dept_df_sorted["department"],
            orientation='h',
            marker_color=colors,
            text=dept_df_sorted["avg_overall"],
            textposition='outside',
        ))
        fig3.update_layout(
            title="ציון כולל לפי מחלקה",
            font=PLOTLY_FONT, plot_bgcolor=C_BG, paper_bgcolor="white",
            xaxis_range=[0, 5],
            title_x=1.0, title_xanchor="right",
        )
        st.plotly_chart(fig3, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 2: Department analysis
# ════════════════════════════════════════════════════════════════════════════
elif page == "🏢 ניתוח מחלקות":
    st.title("🏢 ניתוח מחלקות")

    dept_df = database.get_department_stats()
    if dept_df.empty:
        st.warning("אין נתונים.")
        st.stop()

    all_depts = dept_df['department'].tolist()
    selected = st.multiselect("בחר מחלקות", all_depts, default=all_depts)
    if not selected:
        selected = all_depts

    filtered = dept_df[dept_df['department'].isin(selected)]

    # ── Summary table ─────────────────────────────────────────────────────
    st.subheader("סיכום מחלקות")
    display_cols = {
        "department": "מחלקה",
        "employee_count": "עובדים",
        "avg_overall": "ציון כולל",
        "avg_enps": "eNPS",
        "high_risk_pct": "% סיכון גבוה",
        "avg_salary": "שכר",
        "avg_development": "פיתוח",
        "avg_manager": "מנהל",
        "avg_balance": "איזון",
    }
    show_cols = [c for c in display_cols if c in filtered.columns]
    tbl = filtered[show_cols].rename(columns=display_cols)
    st.dataframe(tbl, use_container_width=True, hide_index=True)

    # ── Side-by-side comparison bar charts ────────────────────────────────
    st.subheader("השוואת מחלקות — ממדי שביעות רצון")
    dims = {"avg_salary": "שכר", "avg_development": "פיתוח",
            "avg_manager": "מנהל", "avg_balance": "איזון"}

    melted_rows = []
    for col, label in dims.items():
        if col in filtered.columns:
            for _, row in filtered.iterrows():
                melted_rows.append({"מחלקה": row["department"], "ממד": label, "ציון": row[col]})
    melted = pd.DataFrame(melted_rows)

    if not melted.empty:
        fig = px.bar(melted, x="ממד", y="ציון", color="מחלקה", barmode="group",
                     title="השוואת ממדי שביעות רצון לפי מחלקה",
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_layout(font=PLOTLY_FONT, plot_bgcolor=C_BG, paper_bgcolor="white",
                          title_x=1.0, title_xanchor="right", yaxis_range=[0, 5])
        st.plotly_chart(fig, use_container_width=True)

    # ── Radar charts per department ───────────────────────────────────────
    st.subheader("פרופיל מחלקה — גרף עכביש")
    radar_dept = st.selectbox("בחר מחלקה לגרף עכביש", selected)
    radar_row = filtered[filtered["department"] == radar_dept]
    if not radar_row.empty:
        r = radar_row.iloc[0]
        categories = ["שכר", "פיתוח", "מנהל", "איזון"]
        vals = [r.get("avg_salary", 0), r.get("avg_development", 0),
                r.get("avg_manager", 0), r.get("avg_balance", 0)]
        fig_radar = go.Figure(go.Scatterpolar(
            r=vals + [vals[0]],
            theta=categories + [categories[0]],
            fill='toself',
            line_color=C_BLUE,
            fillcolor=f"rgba(26,115,232,0.2)",
            name=radar_dept,
        ))
        fig_radar.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
            title=f"פרופיל {radar_dept}",
            font=PLOTLY_FONT, paper_bgcolor="white",
            title_x=0.5,
        )
        st.plotly_chart(fig_radar, use_container_width=True)

    # ── Gaps analysis ─────────────────────────────────────────────────────
    st.subheader("פערים בולטים — מחלקות מתחת לממוצע")
    if "avg_overall" in filtered.columns:
        overall_mean = filtered["avg_overall"].mean()
        gaps = filtered[filtered["avg_overall"] < (overall_mean - 0.5)]
        if not gaps.empty:
            for _, row in gaps.iterrows():
                delta = round(overall_mean - row["avg_overall"], 2)
                st.markdown(f"""
                <div style="background:#fff3f3; border-right:4px solid {C_RED};
                     padding:10px 16px; border-radius:6px; margin-bottom:8px;">
                    <strong>{row['department']}</strong> — ציון כולל: {row['avg_overall']}
                    (פער של {delta} מתחת לממוצע {round(overall_mean,2)})
                </div>""", unsafe_allow_html=True)
        else:
            st.success("אין פערים בולטים במחלקות שנבחרו.")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 3: Manager analysis
# ════════════════════════════════════════════════════════════════════════════
elif page == "👤 ניתוח מנהלים":
    st.title("👤 ניתוח מנהלים")

    mgr_df = database.get_manager_stats()
    if mgr_df.empty:
        st.warning("אין נתונים.")
        st.stop()

    emp_df = database.get_all_employees()

    mgr_sorted = mgr_df.sort_values("engagement_score", ascending=False).reset_index(drop=True)

    # ── Full table ────────────────────────────────────────────────────────
    st.subheader("כל המנהלים")
    display_cols = {
        "manager_name": "שם מנהל",
        "department": "מחלקה",
        "employee_count": "עובדים",
        "engagement_score": "ציון מעורבות",
        "avg_manager": "שביעות ממנהל",
        "avg_development": "פיתוח",
        "avg_balance": "איזון",
        "avg_enps": "eNPS",
        "high_risk_pct": "% סיכון גבוה",
    }
    show_cols = [c for c in display_cols if c in mgr_sorted.columns]
    tbl = mgr_sorted[show_cols].rename(columns=display_cols)
    st.dataframe(tbl, use_container_width=True, hide_index=True)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🏆 Top 5 מנהלים")
        top5 = mgr_sorted.head(5)
        for i, (_, row) in enumerate(top5.iterrows(), 1):
            st.markdown(f"""
            <div style="background:#f0fff4; border-right:4px solid {C_GREEN};
                 padding:10px 16px; border-radius:6px; margin-bottom:8px;">
                <strong>🏆 {i}. {row['manager_name']}</strong> | {row.get('department','')}
                <br><small>ציון מעורבות: {row.get('engagement_score','N/A')} |
                עובדים: {row.get('employee_count','N/A')} |
                % סיכון גבוה: {row.get('high_risk_pct','N/A')}%</small>
            </div>""", unsafe_allow_html=True)

    with col2:
        st.subheader("⚠️ Bottom 5 מנהלים")
        bot5 = mgr_sorted.tail(5).iloc[::-1]
        for i, (_, row) in enumerate(bot5.iterrows(), 1):
            st.markdown(f"""
            <div style="background:#fff8f8; border-right:4px solid {C_RED};
                 padding:10px 16px; border-radius:6px; margin-bottom:8px;">
                <strong>⚠️ {row['manager_name']}</strong> | {row.get('department','')}
                <br><small>ציון מעורבות: {row.get('engagement_score','N/A')} |
                עובדים: {row.get('employee_count','N/A')} |
                % סיכון גבוה: {row.get('high_risk_pct','N/A')}%</small>
            </div>""", unsafe_allow_html=True)

    # ── Coaching letters preview ───────────────────────────────────────────
    st.markdown("---")
    st.subheader("מכתבי Coaching — Bottom 5 מנהלים")

    bot5_list = mgr_sorted.tail(5)
    for _, mgr_row in bot5_list.iterrows():
        mgr_name = mgr_row['manager_name']
        dept = mgr_row.get('department', '')
        emp_count = int(mgr_row.get('employee_count', 0))

        with st.expander(f"📝 מכתב Coaching — {mgr_name} ({dept})"):
            mgr_sat = float(mgr_row.get('avg_manager', 0))
            dev_sat = float(mgr_row.get('avg_development', 0))
            bal_sat = float(mgr_row.get('avg_balance', 0))
            sal_sat = float(mgr_row.get('avg_salary', 0))
            avg_enps = float(mgr_row.get('avg_enps', 0))
            high_risk_pct = float(mgr_row.get('high_risk_pct', 0))

            # Sample quotes
            mgr_emps = emp_df[emp_df['manager_name'] == mgr_name]
            quotes = []
            if not mgr_emps.empty and 'open_feedback' in mgr_emps.columns:
                valid_feedback = mgr_emps['open_feedback'].dropna().tolist()
                import random
                n = min(3, len(valid_feedback))
                if n > 0:
                    quotes = random.sample(valid_feedback, n)

            st.markdown(f"""
**לכבוד {mgr_name},**

בעקבות סקר שביעות רצון העובדים, ברצוננו לשתף אותך בממצאים מצוות {dept} ({emp_count} עובדים).

**ציוני הצוות שלך:**
- שביעות רצון ממנהל: **{mgr_sat:.1f}/5**
- שביעות רצון מפיתוח: **{dev_sat:.1f}/5**
- שביעות רצון מאיזון עבודה-חיים: **{bal_sat:.1f}/5**
- שביעות רצון שכר: **{sal_sat:.1f}/5**
- eNPS ממוצע: **{avg_enps:.0f}**
- % בסיכון עזיבה גבוה: **{high_risk_pct:.0f}%**
""")
            if quotes:
                st.markdown("**נושאים עיקריים שעלו ממשוב הצוות:**")
                for q in quotes:
                    st.markdown(f'> *"{q}"*')

            scores_dict = {
                "שביעות רצון ממנהל": mgr_sat,
                "שביעות רצון מפיתוח": dev_sat,
                "שביעות רצון מאיזון עבודה-חיים": bal_sat,
                "שביעות רצון שכר": sal_sat,
            }
            weak = sorted([(k, v) for k, v in scores_dict.items() if v < 3.5], key=lambda x: x[1])
            if weak:
                st.markdown("**תחומים לשיפור עיקריים:**")
                for k, v in weak[:3]:
                    st.markdown(f"- ⚠️ {k}: {v:.1f}/5")

            st.markdown("""
**המלצות לפעולה:**
1. לקיים פגישות אחד-על-אחד שבועיות עם כל חבר צוות
2. לספק פידבק בונה ומעודד באופן שוטף
3. להשתתף בסדנת מנהיגות ארגונית

---
*אנו כאן לתמוך בך בתהליך,*
**צוות משאבי אנוש**
""")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 4: Flight Risk
# ════════════════════════════════════════════════════════════════════════════
elif page == "⚠️ סיכון עזיבה":
    st.title("⚠️ סיכון עזיבה")

    df = database.get_all_employees()
    if df.empty:
        st.warning("אין נתונים.")
        st.stop()

    high_risk = df[df['risk_level'] == 'גבוה']
    mid_risk  = df[df['risk_level'] == 'בינוני']
    low_risk  = df[df['risk_level'] == 'נמוך']
    total = len(df)

    # ── Summary cards ─────────────────────────────────────────────────────
    cols = st.columns(3)
    with cols[0]:
        st.markdown(kpi_card("🔴 סיכון גבוה",
                             f"{len(high_risk)} ({round(100*len(high_risk)/total,1)}%)"),
                    unsafe_allow_html=True)
    with cols[1]:
        st.markdown(kpi_card("🟡 סיכון בינוני",
                             f"{len(mid_risk)} ({round(100*len(mid_risk)/total,1)}%)"),
                    unsafe_allow_html=True)
    with cols[2]:
        st.markdown(kpi_card("🟢 סיכון נמוך",
                             f"{len(low_risk)} ({round(100*len(low_risk)/total,1)}%)"),
                    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Stacked bar by department ─────────────────────────────────────────
    dept_risk = df.groupby(['department', 'risk_level']).size().reset_index(name='count')
    fig = px.bar(dept_risk, x='department', y='count', color='risk_level',
                 title='רמות סיכון לפי מחלקה',
                 color_discrete_map={'גבוה': C_RED, 'בינוני': C_YELLOW, 'נמוך': C_GREEN},
                 barmode='stack')
    fig.update_layout(font=PLOTLY_FONT, plot_bgcolor=C_BG, paper_bgcolor="white",
                      title_x=1.0, title_xanchor="right")
    st.plotly_chart(fig, use_container_width=True)

    # ── High risk table (filterable) ──────────────────────────────────────
    st.subheader("עובדים בסיכון גבוה")

    dept_filter = st.multiselect("סנן לפי מחלקה", df['department'].unique().tolist(), key="risk_dept")
    mgr_filter  = st.multiselect("סנן לפי מנהל", df['manager_name'].unique().tolist(), key="risk_mgr")

    filtered = high_risk.copy()
    if dept_filter:
        filtered = filtered[filtered['department'].isin(dept_filter)]
    if mgr_filter:
        filtered = filtered[filtered['manager_name'].isin(mgr_filter)]

    show_cols = ['name', 'department', 'manager_name', 'flight_risk_score', 'risk_level',
                 'turnover_intent', 'enps',
                 'salary_satisfaction', 'development_satisfaction', 'manager_satisfaction', 'balance_satisfaction']
    show_cols = [c for c in show_cols if c in filtered.columns]
    col_labels = {
        'name': 'שם', 'department': 'מחלקה', 'manager_name': 'מנהל',
        'flight_risk_score': 'ציון סיכון', 'risk_level': 'רמה',
        'turnover_intent': 'כוונת עזיבה', 'enps': 'eNPS',
        'salary_satisfaction': 'שכר', 'development_satisfaction': 'פיתוח',
        'manager_satisfaction': 'מנהל', 'balance_satisfaction': 'איזון',
    }
    tbl = filtered[show_cols].rename(columns=col_labels)
    st.dataframe(tbl, use_container_width=True, hide_index=True)

    # Export high-risk to Excel
    if not filtered.empty:
        buf = io.BytesIO()
        filtered.to_excel(buf, index=False, engine='openpyxl')
        buf.seek(0)
        st.download_button(
            label="⬇️ הורד רשימת עובדים בסיכון גבוה (Excel)",
            data=buf,
            file_name="high_risk_employees.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── Risk driver analysis ──────────────────────────────────────────────
    st.subheader("מה מניע את הסיכון? — ניתוח קורלציות")
    if not df.empty:
        corr_cols = ['salary_satisfaction', 'development_satisfaction',
                     'manager_satisfaction', 'balance_satisfaction', 'enps', 'turnover_intent']
        corr_cols = [c for c in corr_cols if c in df.columns]
        corr_labels = {
            'salary_satisfaction': 'שכר',
            'development_satisfaction': 'פיתוח',
            'manager_satisfaction': 'מנהל',
            'balance_satisfaction': 'איזון',
            'enps': 'eNPS',
            'turnover_intent': 'כוונת עזיבה',
        }
        corr_data = []
        for col in corr_cols:
            if col in df.columns and 'flight_risk_score' in df.columns:
                corr_val = df[col].corr(df['flight_risk_score'])
                corr_data.append({"גורם": corr_labels.get(col, col), "קורלציה עם סיכון": round(corr_val, 3)})
        if corr_data:
            corr_df = pd.DataFrame(corr_data).sort_values("קורלציה עם סיכון", ascending=True)
            fig_corr = px.bar(corr_df, x="קורלציה עם סיכון", y="גורם",
                              orientation='h', title="קורלציה של גורמים עם ציון סיכון עזיבה",
                              color="קורלציה עם סיכון",
                              color_continuous_scale=[[0, C_GREEN], [0.5, C_YELLOW], [1, C_RED]])
            fig_corr.update_layout(font=PLOTLY_FONT, plot_bgcolor=C_BG, paper_bgcolor="white",
                                   title_x=1.0, title_xanchor="right")
            st.plotly_chart(fig_corr, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 5: Open text analysis
# ════════════════════════════════════════════════════════════════════════════
elif page == "💬 ניתוח תשובות פתוחות":
    st.title("💬 ניתוח תשובות פתוחות")

    df = database.get_all_employees()
    if df.empty:
        st.warning("אין נתונים.")
        st.stop()

    TOPICS = {
        "שכר ותגמול":         ["שכר", "משכורת", "תגמול", "העלאה", "בונוס"],
        "פיתוח מקצועי":       ["קידום", "פיתוח", "הכשרה", "למידה", "גדילה"],
        "מנהיגות":            ["מנהל", "מנהלת", "הנהלה", "מקשיב", "פידבק"],
        "איזון עבודה-חיים":   ["עומס", "שעות", "לילה", "חיים", "איזון"],
        "תרבות ארגונית":      ["תרבות", "שקיפות", "אמון", "ערכים"],
        "ביטחון תעסוקתי":     ["פיטורים", "חוסר ביטחון", "עזיבה", "לעזוב"],
    }

    # ── Filters ───────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        dept_filter = st.multiselect("סנן לפי מחלקה", df['department'].unique().tolist(), key="open_dept")
    with col2:
        risk_filter = st.multiselect("סנן לפי רמת סיכון", ["גבוה", "בינוני", "נמוך"], key="open_risk")

    filtered = df.copy()
    if dept_filter:
        filtered = filtered[filtered['department'].isin(dept_filter)]
    if risk_filter:
        filtered = filtered[filtered['risk_level'].isin(risk_filter)]

    feedback_texts = filtered['open_feedback'].dropna().tolist()

    # ── Sentiment split ───────────────────────────────────────────────────
    POSITIVE_KEYWORDS = ["נהנה", "מדהים", "מקשיב", "גדל", "שקיפות", "גאה", "מצוין", "מוערך", "מעניינים", "הוגנת"]
    NEGATIVE_KEYWORDS = ["לא תואם", "לא מקשיב", "אין מסלול", "לא סביר", "חוסר הערכה",
                         "נמוכה", "מיושנות", "האשמות", "ירוד", "בלחץ", "תקוע",
                         "שמועות", "לא שוויוני", "עזיבה הולכות", "לא קיים"]

    pos, neg, neu = 0, 0, 0
    for text in feedback_texts:
        t_lower = text
        if any(kw in t_lower for kw in POSITIVE_KEYWORDS):
            pos += 1
        elif any(kw in t_lower for kw in NEGATIVE_KEYWORDS):
            neg += 1
        else:
            neu += 1

    total_f = pos + neg + neu
    if total_f > 0:
        st.subheader("פילוח סנטימנט")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(kpi_card("😊 חיובי", f"{pos} ({round(100*pos/total_f)}%)"), unsafe_allow_html=True)
        with c2:
            st.markdown(kpi_card("😐 ניטרלי", f"{neu} ({round(100*neu/total_f)}%)"), unsafe_allow_html=True)
        with c3:
            st.markdown(kpi_card("😟 שלילי", f"{neg} ({round(100*neg/total_f)}%)"), unsafe_allow_html=True)

    # ── Topic frequency ───────────────────────────────────────────────────
    st.subheader("נושאים שעולים בתשובות")
    topic_counts = []
    for topic, keywords in TOPICS.items():
        count = sum(1 for text in feedback_texts if any(kw in text for kw in keywords))
        topic_counts.append({"נושא": topic, "מספר אזכורים": count})

    topic_df = pd.DataFrame(topic_counts).sort_values("מספר אזכורים", ascending=True)
    fig = px.bar(topic_df, x="מספר אזכורים", y="נושא", orientation='h',
                 title="נושאים שעולים מהמשוב הפתוח",
                 color="מספר אזכורים",
                 color_continuous_scale=[[0, C_BLUE], [1, C_RED]])
    fig.update_layout(font=PLOTLY_FONT, plot_bgcolor=C_BG, paper_bgcolor="white",
                      title_x=1.0, title_xanchor="right")
    st.plotly_chart(fig, use_container_width=True)

    # ── Sample quotes per topic ───────────────────────────────────────────
    st.subheader("ציטוטים לפי נושא")
    selected_topic = st.selectbox("בחר נושא לציטוטים", list(TOPICS.keys()))
    keywords = TOPICS[selected_topic]
    matching = [t for t in feedback_texts if any(kw in t for kw in keywords)]
    import random
    samples = random.sample(matching, min(3, len(matching))) if matching else []
    if samples:
        for s in samples:
            st.markdown(f"""
            <div style="background:#f8f9fa; border-right:3px solid {C_BLUE};
                 padding:10px 16px; border-radius:6px; margin-bottom:8px; font-style:italic;">
                "{s}"
            </div>""", unsafe_allow_html=True)
    else:
        st.info(f"לא נמצאו ציטוטים לנושא '{selected_topic}' עם הפילטרים הנוכחיים.")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 6: Strategic recommendations
# ════════════════════════════════════════════════════════════════════════════
elif page == "💡 המלצות אסטרטגיות":
    st.title("💡 המלצות אסטרטגיות")

    stats = database.get_summary_stats()
    dept_df = database.get_department_stats()

    if stats['total_employees'] == 0:
        st.warning("אין נתונים.")
        st.stop()

    recs = []

    if stats.get('avg_salary', 5) < 3.2:
        recs.append({
            "title": "שכר ותגמול — סקר שוק",
            "detail": "ביצוע סקר שכר חיצוני ופנימי. התאמת מסגרות שכר לשוק בתוך 60 יום. "
                      "שקילת מנגנוני תגמול משתנה לעובדים מצטיינים.",
            "priority": "גבוה", "roi": "הפחתת עזיבות ב-15-20%",
            "class": "rec-card-high",
        })

    if stats.get('high_risk_pct', 0) > 20:
        recs.append({
            "title": "שימור עובדים — תוכנית Retention",
            "detail": "הקמת תוכנית שימור ממוקדת לעובדים בסיכון גבוה: שיחות אישיות, "
                      "תוכניות פיתוח אישיות, ומענקי שימור. מינוי מנהלי רווחה מחלקתיים.",
            "priority": "גבוה", "roi": "חיסכון ממוצע 50-100K ₪ לכל עובד נשמר",
            "class": "rec-card-high",
        })

    if stats.get('avg_development', 5) < 3.0:
        recs.append({
            "title": "פיתוח מקצועי — מסלולי קידום",
            "detail": "הגדרת מסלולי קידום ברורים בכל מחלקה. תוכנית הכשרות עדכנית ורלוונטית. "
                      "הקמת מנטורינג פנים-ארגוני. תקציב הכשרה אישי לכל עובד.",
            "priority": "בינוני", "roi": "שיפור מחוברות ב-10-15%",
            "class": "rec-card-mid",
        })

    if stats.get('avg_manager', 5) < 3.2:
        recs.append({
            "title": "מנהיגות — פיתוח מנהלים",
            "detail": "הדרכות מנהיגות לכלל המנהלים. Coaching ממוקד לBottom 5 מנהלים. "
                      "שילוב KPIs של מחוברות צוות בהערכת מנהלים.",
            "priority": "גבוה", "roi": "שיפור ציון מנהל ב-0.3-0.5 נקודות",
            "class": "rec-card-high",
        })

    if stats.get('avg_balance', 5) < 3.0:
        recs.append({
            "title": "איזון עבודה-חיים — מדיניות גמישות",
            "detail": "הטמעת מדיניות עבודה גמישה. הגבלת שעות עבודה. "
                      "מדידת עומס עבודה ברמת צוות ותיקון חוסר איזון.",
            "priority": "בינוני", "roi": "הפחתת burnout ושיפור פריון",
            "class": "rec-card-mid",
        })

    if not dept_df.empty and 'high_risk_pct' in dept_df.columns:
        overall_risk = stats.get('high_risk_pct', 0)
        worst = dept_df.nlargest(1, 'high_risk_pct').iloc[0]
        if worst['high_risk_pct'] > overall_risk * 1.5 and worst['high_risk_pct'] > 25:
            recs.append({
                "title": f"מחלקת {worst['department']} — התייחסות מיידית",
                "detail": f"שיעור סיכון עזיבה של {worst['high_risk_pct']}% — "
                          f"פי {round(worst['high_risk_pct']/max(overall_risk,1),1)} מהממוצע. "
                          "נדרשת תוכנית מיוחדת עבור המחלקה כולל שיחות אישיות עם כלל העובדים.",
                "priority": "גבוה", "roi": "מניעת גל עזיבות מרוכז",
                "class": "rec-card-high",
            })

    if not recs:
        recs.append({
            "title": "שמירה על מגמה חיובית",
            "detail": "המדדים הארגוניים תקינים. המשך מעקב רבעוני. "
                      "חיזוק נקודות חוזק ארגוניות ושיתוף best practices.",
            "priority": "נמוך", "roi": "שמירה על רמת מחוברות קיימת",
            "class": "rec-card-low",
        })

    priority_icon = {"גבוה": "🔴", "בינוני": "🟡", "נמוך": "🟢"}

    for i, rec in enumerate(recs, 1):
        st.markdown(f"""
        <div class="rec-card {rec['class']}">
            <div class="rec-title">{priority_icon.get(rec['priority'], '')} {i}. {rec['title']}</div>
            <div class="rec-detail">{rec['detail']}</div>
            <div class="rec-meta">עדיפות: <strong>{rec['priority']}</strong> &nbsp;|&nbsp;
                 השפעה משוערת: {rec['roi']}</div>
        </div>""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PAGE 7: Export reports
# ════════════════════════════════════════════════════════════════════════════
elif page == "📊 יצוא דוחות":
    st.title("📊 יצוא דוחות")

    stats = database.get_summary_stats()
    dept_df = database.get_department_stats()
    mgr_df = database.get_manager_stats()
    emp_df = database.get_all_employees()

    if stats['total_employees'] == 0:
        st.warning("אין נתונים לייצא.")
        st.stop()

    st.markdown("### 📄 דוחות Word")
    col1, col2 = st.columns(2)

    with col1:
        if st.button("📊 הורד דוח Word לדירקטוריון", use_container_width=True):
            with st.spinner("מייצר דוח..."):
                from exporters.word_report import generate_word_report
                docx_bytes = generate_word_report(stats, dept_df, mgr_df, emp_df)
            st.download_button(
                label="⬇️ לחץ להורדה — דוח לדירקטוריון.docx",
                data=docx_bytes,
                file_name="hr_pulse_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_report",
            )

    with col2:
        if st.button("✉️ הורד מכתבי Coaching (Bottom 5 מנהלים)", use_container_width=True):
            with st.spinner("מייצר מכתבי coaching..."):
                from exporters.coaching import generate_coaching_letters
                coaching_bytes = generate_coaching_letters(mgr_df, emp_df)
            st.download_button(
                label="⬇️ לחץ להורדה — מכתבי_coaching.docx",
                data=coaching_bytes,
                file_name="coaching_letters.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_coaching",
            )

    st.markdown("### 📊 דוחות Excel")
    col3, col4 = st.columns(2)

    with col3:
        if not emp_df.empty:
            high_risk = emp_df[emp_df['risk_level'] == 'גבוה']
            buf = io.BytesIO()
            high_risk.to_excel(buf, index=False, engine='openpyxl')
            buf.seek(0)
            st.download_button(
                label="⬇️ רשימת עובדים בסיכון גבוה (Excel)",
                data=buf,
                file_name="high_risk_employees.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with col4:
        if not emp_df.empty:
            buf2 = io.BytesIO()
            emp_df.to_excel(buf2, index=False, engine='openpyxl')
            buf2.seek(0)
            st.download_button(
                label="⬇️ נתונים גולמיים — כל העובדים (Excel)",
                data=buf2,
                file_name="all_employees.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.markdown("---")
    st.info("גרפים אינטראקטיביים זמינים בדשבורד בלבד. הדוח כולל טבלאות מספריות עם אותם הנתונים.")


# ════════════════════════════════════════════════════════════════════════════
# PAGE 8: Data management
# ════════════════════════════════════════════════════════════════════════════
elif page == "⚙️ ניהול נתונים":
    st.title("⚙️ ניהול נתונים")

    meta = database.get_db_meta()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("סה\"כ רשומות", meta['total_records'])
    with col2:
        st.metric("עדכון אחרון", meta.get('last_updated', 'N/A'))
    with col3:
        st.metric("מסד נתונים", "SQLite")

    st.markdown("---")

    # ── Upload Excel ──────────────────────────────────────────────────────
    st.subheader("📥 ייבוא נתונים מ-Excel")
    uploaded = st.file_uploader("העלה קובץ Excel", type=["xlsx", "xls"])

    if uploaded:
        raw_df = pd.read_excel(uploaded, engine='openpyxl')
        st.write("**עמודות שנמצאו בקובץ:**", list(raw_df.columns))
        st.dataframe(raw_df.head(5), use_container_width=True)

        DB_FIELDS = ['name', 'department', 'manager_name', 'hire_date',
                     'salary_satisfaction', 'development_satisfaction',
                     'manager_satisfaction', 'balance_satisfaction',
                     'turnover_intent', 'enps', 'open_feedback']
        FIELD_LABELS = {
            'name': 'שם עובד', 'department': 'מחלקה', 'manager_name': 'מנהל',
            'hire_date': 'תאריך קליטה', 'salary_satisfaction': 'שביעות שכר',
            'development_satisfaction': 'שביעות פיתוח', 'manager_satisfaction': 'שביעות מנהל',
            'balance_satisfaction': 'שביעות איזון', 'turnover_intent': 'כוונת עזיבה',
            'enps': 'eNPS', 'open_feedback': 'משוב פתוח',
        }

        st.markdown("**מיפוי עמודות:**")
        mapping = {}
        file_cols = ["— לא למפות —"] + list(raw_df.columns)
        for field in DB_FIELDS:
            default_idx = 0
            for i, c in enumerate(raw_df.columns):
                if c.lower() == field.lower() or c == FIELD_LABELS.get(field, ''):
                    default_idx = i + 1
                    break
            mapping[field] = st.selectbox(
                f"{FIELD_LABELS.get(field, field)}", file_cols,
                index=default_idx, key=f"map_{field}"
            )

        if st.button("ייבא נתונים"):
            mapped_df = pd.DataFrame()
            for field, col in mapping.items():
                if col != "— לא למפות —":
                    mapped_df[field] = raw_df[col]
            if 'name' not in mapped_df.columns or mapped_df['name'].isna().all():
                st.error("חובה למפות את עמודת שם העובד.")
            else:
                database.import_from_dataframe(mapped_df)
                st.success(f"יובאו {len(mapped_df)} רשומות בהצלחה!")
                st.rerun()

    st.markdown("---")

    # ── View / search employees ───────────────────────────────────────────
    st.subheader("📋 כל העובדים")
    df = database.get_all_employees()
    if not df.empty:
        search = st.text_input("חיפוש לפי שם / מחלקה / מנהל")
        if search:
            mask = (df['name'].str.contains(search, na=False) |
                    df['department'].str.contains(search, na=False) |
                    df['manager_name'].str.contains(search, na=False))
            df = df[mask]

        # Paginate
        page_size = 50
        total_pages = max(1, (len(df) - 1) // page_size + 1)
        cur_page = st.number_input("עמוד", min_value=1, max_value=total_pages, value=1, step=1)
        start = (cur_page - 1) * page_size
        end = start + page_size
        st.dataframe(df.iloc[start:end], use_container_width=True, hide_index=True)
        st.caption(f"מציג {start+1}–{min(end, len(df))} מתוך {len(df)} רשומות")

        # ── Edit employee ──────────────────────────────────────────────────
        st.markdown("---")
        st.subheader("✏️ עריכת עובד")
        edit_id = st.number_input("מזהה עובד לעריכה (ID)", min_value=1, step=1, value=1)
        emp_row = database.get_all_employees()
        emp_row = emp_row[emp_row['id'] == edit_id]
        if not emp_row.empty:
            emp = emp_row.iloc[0].to_dict()
            with st.form(f"edit_{edit_id}"):
                name = st.text_input("שם", value=str(emp.get('name', '')))
                dept = st.text_input("מחלקה", value=str(emp.get('department', '')))
                mgr  = st.text_input("מנהל", value=str(emp.get('manager_name', '')))
                hire = st.text_input("תאריך קליטה", value=str(emp.get('hire_date', '')))

                c1, c2 = st.columns(2)
                with c1:
                    sal  = st.slider("שביעות שכר", 1.0, 5.0, float(emp.get('salary_satisfaction', 3.0)), 0.1)
                    dev  = st.slider("שביעות פיתוח", 1.0, 5.0, float(emp.get('development_satisfaction', 3.0)), 0.1)
                with c2:
                    man  = st.slider("שביעות מנהל", 1.0, 5.0, float(emp.get('manager_satisfaction', 3.0)), 0.1)
                    bal  = st.slider("שביעות איזון", 1.0, 5.0, float(emp.get('balance_satisfaction', 3.0)), 0.1)

                turn = st.slider("כוונת עזיבה", 1, 5, int(emp.get('turnover_intent', 3)))
                enps = st.slider("eNPS", -100, 100, int(emp.get('enps', 0)))
                feedback = st.text_area("משוב פתוח", value=str(emp.get('open_feedback', '')))

                submitted = st.form_submit_button("שמור שינויים")
                if submitted:
                    database.update_employee(edit_id, {
                        'name': name, 'department': dept, 'manager_name': mgr,
                        'hire_date': hire, 'salary_satisfaction': sal,
                        'development_satisfaction': dev, 'manager_satisfaction': man,
                        'balance_satisfaction': bal, 'turnover_intent': turn,
                        'enps': enps, 'open_feedback': feedback,
                    })
                    st.success("הרשומה עודכנה בהצלחה!")
                    st.rerun()
        else:
            st.info("לא נמצא עובד עם מזהה זה.")

        # ── Delete employee ────────────────────────────────────────────────
        st.markdown("---")
        st.subheader("🗑️ מחיקת עובד")
        del_id = st.number_input("מזהה עובד למחיקה (ID)", min_value=1, step=1, value=1, key="del_id")
        if st.button("מחק עובד", type="secondary"):
            database.delete_employee(del_id)
            st.success(f"עובד {del_id} נמחק.")
            st.rerun()

    else:
        st.info("אין עובדים במסד הנתונים.")

    # ── Reset to demo ─────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🔄 איפוס לנתוני דמו")
    st.warning("פעולה זו תמחק את כל הנתונים הקיימים ותייצר מחדש ~1000 עובדי דמו.")
    if st.button("איפוס לנתוני דמו", type="primary"):
        import subprocess, sys as _sys
        result = subprocess.run(
            [_sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), "generate_demo.py"), "--force"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            st.success("נתוני הדמו נוצרו מחדש בהצלחה!")
            st.code(result.stdout)
        else:
            st.error("שגיאה בייצור נתוני דמו:")
            st.code(result.stderr)
        st.rerun()
