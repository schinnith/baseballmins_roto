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

    # Team names: rows 3–12 (index), column 1
    team_names = df.iloc[3:13, 1].tolist()

    # Stat headers: row 14 (index 14)
    stat_headers = df.iloc[14].dropna().tolist()

    # Category totals: rows 15–24 (index), columns 0:len(stat_headers)
    stat_rows = df.iloc[15:25, 0:len(stat_headers)].applymap(
        lambda x: float(str(x).replace(',', '')) if pd.notnull(x) and str(x).replace('.', '', 1).replace('-', '', 1).isdigit() else x
    )
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
1. Go to our ESPN league standings page.
2. Highlight from **“Season Stats”** to the final **“Moves”** cell of the last team.
3. **Copy** that selection and **paste as plain text** (or “paste values only”) into the downloaded  
   [espn_roto_standings.csv](https://github.com/schinnith/baseballmins_roto/raw/main/espn_roto_standings.csv), starting at **cell A1**.
4. Be sure it replaces the existing content without adding rows or columns.
5. Save the file as CSV and upload it below.
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

    # Format raw stats: only AVG, ERA, WHIP need decimals
    decimal_cols = ['AVG', 'ERA', 'WHIP']
    format_dict = {col: '{:.3f}' for col in decimal_cols}
    for col in df.columns:
        if col not in format_dict:
            format_dict[col] = '{:.0f}'

    st.subheader("📊 Raw Category Totals")
    st.dataframe(df.style.format(format_dict))
