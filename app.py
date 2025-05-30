import streamlit as st
import pandas as pd
import base64

# === CONFIG ===
ASCENDING_CATEGORIES = ['ERA', 'WHIP']
DESCENDING_CATEGORIES = ['R', 'HR', 'RBI', 'SB', 'AVG', 'K', 'W', 'SV']
CATEGORIES = DESCENDING_CATEGORIES + ASCENDING_CATEGORIES

# === ROTO POINTS ===
def calculate_roto_points(series, ascending):
    ranked = series.rank(ascending=ascending, method='min')
    points = ranked.copy()
    value_counts = ranked.value_counts()
    for rank, count in value_counts.items():
        total = sum(11 - r for r in range(int(rank), int(rank + count)))
        points[ranked == rank] = total / count
    return points

# === PARSE ESPN STRUCTURED EXPORT ===
def parse_structured_csv(uploaded_file):
    df = pd.read_csv(uploaded_file, header=None, encoding='latin1')

    # Team names: rows 4–15 (Excel row 5–16)
    team_names = df.iloc[3:15, 1].dropna().tolist()

    # Stat headers in row 17
    stat_headers = df.iloc[16].dropna().tolist()

    # Category totals: rows 18–27
    stat_rows = df.iloc[17:27, 0:len(stat_headers)].applymap(lambda x: float(str(x).replace(',', '')))
    stat_rows.columns = stat_headers

    # Merge
    final_df = pd.DataFrame({'Team': team_names})
    final_df = pd.concat([final_df.reset_index(drop=True), stat_rows.reset_index(drop=True)], axis=1)
    return final_df

# === STREAMLIT APP ===
st.set_page_config(page_title="NFBC Roto Standings", layout="wide")
st.title("🏆 NFBC Roto Standings Viewer")

st.markdown("""
### 📝 Instructions:
1. Download the sample file: [espn_roto_standings.csv](https://github.com/schinnith/baseballmins_roto/raw/main/espn_roto_standings.csv)
2. Open in Excel or Google Sheets  
3. Paste over the cells with your ESPN standings **without changing the structure**
4. Save as CSV and re-upload it here
""")


# File upload
uploaded_file = st.file_uploader("📤 Upload your ESPN CSV export", type="csv")

if uploaded_file:
    df = parse_structured_csv(uploaded_file)
    roto = df[['Team']].copy()

    for cat in DESCENDING_CATEGORIES:
        roto[f"{cat}_Pts"] = calculate_roto_points(df[cat], ascending=False)
    for cat in ASCENDING_CATEGORIES:
        roto[f"{cat}_Pts"] = calculate_roto_points(df[cat], ascending=True)

    point_cols = [col for col in roto.columns if col.endswith('_Pts')]
    roto['Total_Points'] = roto[point_cols].sum(axis=1)
    roto_sorted = roto.sort_values(by='Total_Points', ascending=False).reset_index(drop=True)

    st.subheader("🔢 Roto Standings")
    st.dataframe(roto_sorted.style.format(precision=1))

    st.subheader("📊 Raw Category Totals")
    st.dataframe(df.style.format(precision=3))
