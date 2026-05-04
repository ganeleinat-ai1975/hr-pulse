import pandas as pd


def calculate_flight_risk_score(row) -> float:
    """Calculate 0-100 flight risk score based on survey dimensions."""
    score = 0.0
    # Turnover intent (0-40 points)
    score += (row['turnover_intent'] - 1) / 4 * 40
    # eNPS (0-20 points, inverted — low eNPS = high risk)
    score += (1 - (row['enps'] + 100) / 200) * 20
    # Salary satisfaction (0-15 points, inverted)
    score += (1 - (row['salary_satisfaction'] - 1) / 4) * 15
    # Development satisfaction (0-15 points, inverted)
    score += (1 - (row['development_satisfaction'] - 1) / 4) * 15
    # Manager satisfaction (0-10 points, inverted)
    score += (1 - (row['manager_satisfaction'] - 1) / 4) * 10
    return min(100.0, max(0.0, round(score, 1)))


def get_risk_level(score: float) -> str:
    if score >= 65:
        return "גבוה"
    if score >= 40:
        return "בינוני"
    return "נמוך"


def recalculate_all(df: pd.DataFrame) -> pd.DataFrame:
    """Recalculate risk scores for entire dataframe and return updated df."""
    df = df.copy()
    df['flight_risk_score'] = df.apply(calculate_flight_risk_score, axis=1)
    df['risk_level'] = df['flight_risk_score'].apply(get_risk_level)
    return df
