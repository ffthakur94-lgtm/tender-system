import os
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px


st.set_page_config(
    page_title="GBP Infraprojects Dashboard",
    page_icon="🏗️",
    layout="wide"
)
st.markdown("""
<style>

.main {
    background-color: #0E1117;
}

div[data-testid="metric-container"] {
    background-color: #1E1E1E;
    border: 1px solid #2E2E2E;
    padding: 15px;
    border-radius: 15px;
}

div[data-testid="metric-container"] label {
    color: white !important;
}

div[data-testid="metric-container"] div {
    color: white !important;
}

h1, h2, h3 {
    color: white;
}

</style>
""", unsafe_allow_html=True)


MASTER_FILE = "all_captured_1cr_tenders.xlsx"
BEST_FILE = "best_tenders.xlsx"
REVIEW_FILE = "review_tenders.xlsx"
BACKUP_FILE = "backup_1cr_tenders.xlsx"


CATEGORY_KEYWORDS = {

    "Construction": [
        "building",
        "residential",
        "school",
        "college",
        "hostel",
        "hospital",
        "quarters",
        "housing",
        "office building",
        "administrative building",
        "industrial building",
        "factory building",
        "warehouse"
    ],

    "Prefab / PEB": [
        "prefab",
        "prefabricated",
        "pre fabricated",
        "peb",
        "pre engineered building",
        "pre-engineered building",
        "steel structure",
        "structural steel",
        "modular building",
        "shed"
    ]
}

# Anything matching these is forced to "Other" even if it also matches a
# category keyword above (e.g. "office building at hydro power station").
# Checked first in detect_category, so exclusion always wins.
EXCLUDE_KEYWORDS = [
    "road", "rcc road", "pmgsy", "widening", "strengthening",
    "hydro", "hydel", "dam", "barrage", "canal", "spillway", "tunnel",
    "water supply", "pipeline", "sewerage", "stp", "sewage treatment",
    "bridge", "culvert", "railway",
    "bro", "barrack", "hangar", "defence", "jammu",
]

# Scope limits from the project brief: reject anything outside this band.
MIN_TENDER_VALUE = 1_00_00_000      # ₹1 Cr
MAX_TENDER_VALUE = 20_00_00_000     # ₹20 Cr


def load_excel(file_name):
    if not os.path.exists(file_name):
        return pd.DataFrame()

    try:
        return pd.read_excel(file_name)
    except Exception:
        return pd.DataFrame()




def dataframe_to_csv(df):
    return df.to_csv(index=False).encode("utf-8")

def clean_number(value):
    try:
        if pd.isna(value):
            return 0
        return int(float(str(value).replace(",", "").replace("₹", "").strip()))
    except Exception:
        return 0


def prepare_data(df):
    if df.empty:
        return df

    df = df.copy()

    if "Tender Value Number" not in df.columns:
        df["Tender Value Number"] = df.get("Tender Value", 0).apply(clean_number)

    if "Score" in df.columns:
        df["Score"] = pd.to_numeric(df["Score"], errors="coerce").fillna(0)

    if "Closing Date" in df.columns:
        df["Closing Date Parsed"] = pd.to_datetime(df["Closing Date"], errors="coerce", dayfirst=True)

    df["Search Text"] = (
        df.get("Tender Title / Ref / ID", "").astype(str) + " " +
        df.get("Organisation", "").astype(str) + " " +
        df.get("Portal", "").astype(str)
    ).str.lower()

    return df


def detect_category(text):
    """Return the matching category, or 'Other' if out of scope.

    Returns a single category (not a joined list) so it lines up exactly
    with the sidebar's category filter options. Exclusion keywords are
    checked first and always win, so a tender that mentions both an
    excluded sector (e.g. hydro, road) and a category keyword (e.g.
    "office building") is still correctly dropped.
    """
    text = str(text).lower()

    if any(keyword in text for keyword in EXCLUDE_KEYWORDS):
        return "Other"

    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(keyword in text for keyword in keywords):
            return category

    return "Other"


def make_clickable_link(url):
    if pd.isna(url) or str(url).strip() == "":
        return ""
    return f'<a href="{url}" target="_blank">Open Tender</a>'


def show_table(df, title):
    st.subheader(title)

    if df.empty:
        st.warning("No data found.")
        return

    show_cols = [
        "Tender Type",
        "Portal",
        "Score",
        "Published Date",
        "Closing Date",
        "Tender Title / Ref / ID",
        "Organisation",
        "Tender Value",
        "Direct Link"
    ]

    available_cols = [col for col in show_cols if col in df.columns]
    display_df = df[available_cols].copy()

    if "Direct Link" in display_df.columns:
        display_df["Direct Link"] = display_df["Direct Link"].apply(make_clickable_link)

    st.write(
        display_df.to_html(escape=False, index=False),
        unsafe_allow_html=True
    )


def apply_scope_filter(df):
    """Enforce the project's actual scope at the data layer.

    Drops anything categorized as 'Other' (roads, hydro, bridges, defence,
    etc.) and anything outside the ₹1 Cr - ₹20 Cr band. This used to only
    happen if a user manually set sidebar filters; now it happens before
    the data ever reaches the UI, KPIs, or charts.

    Returns (kept_df, excluded_count) so the caller can show a transparency
    note about how many rows were dropped and why.
    """
    if df.empty:
        return df, 0

    before = len(df)

    in_scope = (
        (df["Category"] != "Other")
        & (df["Tender Value Number"] >= MIN_TENDER_VALUE)
        & (df["Tender Value Number"] <= MAX_TENDER_VALUE)
    )

    kept = df[in_scope].copy()
    excluded_count = before - len(kept)

    return kept, excluded_count


def filter_data(df, search_text, portal_filter, type_filter, category_filter):
    if df.empty:
        return df

    filtered = df.copy()

    if search_text:
        search_text = search_text.lower()
        filtered = filtered[filtered["Search Text"].str.contains(search_text, na=False)]

    if portal_filter and "All" not in portal_filter:
        filtered = filtered[filtered["Portal"].isin(portal_filter)]

    if type_filter and "All" not in type_filter and "Tender Type" in filtered.columns:
        filtered = filtered[filtered["Tender Type"].isin(type_filter)]

    if category_filter and "All" not in category_filter:
        filtered = filtered[filtered["Category"].isin(category_filter)]

    return filtered


master_df = prepare_data(load_excel(MASTER_FILE))
best_df = prepare_data(load_excel(BEST_FILE))
review_df = prepare_data(load_excel(REVIEW_FILE))
backup_df = prepare_data(load_excel(BACKUP_FILE))

if not master_df.empty:
    master_df["Category"] = master_df["Search Text"].apply(detect_category)

if not best_df.empty:
    best_df["Category"] = best_df["Search Text"].apply(detect_category)

if not review_df.empty:
    review_df["Category"] = review_df["Search Text"].apply(detect_category)

if not backup_df.empty:
    backup_df["Category"] = backup_df["Search Text"].apply(detect_category)

# Enforce scope (Construction / Prefab-PEB only, ₹1 Cr - ₹20 Cr) on every
# dataframe, not just on whatever the sidebar happens to have selected.
master_df, master_excluded = apply_scope_filter(master_df)
best_df, best_excluded = apply_scope_filter(best_df)
review_df, review_excluded = apply_scope_filter(review_df)
backup_df, backup_excluded = apply_scope_filter(backup_df)

total_excluded = master_excluded + best_excluded + review_excluded + backup_excluded


st.title("🏗️ GBP Infraprojects")
st.subheader("Tender Intelligence Dashboard V7")
st.caption("Construction + Prefab Opportunity Monitoring | ₹1 Cr – ₹20 Cr")

if master_df.empty:
    st.error("No Excel data found. Keep this dashboard.py file in the same folder where scraper Excel files are saved.")
    st.stop()

if total_excluded > 0:
    st.info(
        f"ℹ️ {total_excluded} tender(s) excluded across all sheets — outside "
        f"Construction/Prefab-PEB scope or outside ₹1 Cr–₹20 Cr range."
    )


total_count = len(master_df)
best_count = len(best_df)
review_count = len(review_df)
backup_count = len(backup_df)

col1, col2, col3, col4 = st.columns(4)

col1.metric("📋 Total Opportunities", total_count)
col2.metric("⭐ Best Opportunities", best_count)
col3.metric("🟡 Review Opportunities", review_count)
col4.metric("📦 Backup Opportunities", backup_count)



st.divider()
st.subheader("📥 Download Center")

d1, d2, d3, d4 = st.columns(4)

with d1:
    st.download_button("Full Data", dataframe_to_csv(master_df), "all_tenders.csv")

with d2:
    st.download_button("Best", dataframe_to_csv(best_df), "best_tenders.csv")

with d3:
    st.download_button("Review", dataframe_to_csv(review_df), "review_tenders.csv")

with d4:
    st.download_button("Backup", dataframe_to_csv(backup_df), "backup_tenders.csv")

with st.sidebar:

    st.header("Filters")

    quick_keyword = st.selectbox(
        "⚡ Quick Filter",
        [
    "None",
    "Construction",
    "PEB",
    "Prefab",
    "Warehouse",
    "Industrial",
    "School",
    "Hospital"
])

    search_text = st.text_input("Search tender", "")

    portal_options = ["All"] + sorted(master_df["Portal"].dropna().unique().tolist())
    portal_filter = st.multiselect("Portal", portal_options, default=["All"])

    type_options = ["All"]
    if "Tender Type" in master_df.columns:
        type_options += sorted(master_df["Tender Type"].dropna().unique().tolist())

    type_filter = st.multiselect("Tender Type", type_options, default=["All"])

    category_options = ["All"] + sorted(master_df["Category"].dropna().unique().tolist())
    category_filter = st.multiselect("Category", category_options, default=["All"])


filtered_df = filter_data(master_df, search_text, portal_filter, type_filter, category_filter)

if quick_keyword != "None":
    filtered_df = filtered_df[filtered_df["Search Text"].str.contains(quick_keyword.lower(), na=False)]



tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Overview",
    "⭐ Best",
    "🟡 Review",
    "💰 Highest Value",
    "⏰ Closing Priority",
    "📋 All Data"
])


with tab1:
    st.subheader("Portal-wise Tender Count")

    portal_count = filtered_df["Portal"].value_counts().reset_index()
    portal_count.columns = ["Portal", "Count"]

    if not portal_count.empty:
        fig = px.bar(
            portal_count,
            x="Portal",
            y="Count",
            title="Portal-wise Opportunities",
            text="Count"
        )

        fig.update_layout(template="plotly_dark")

        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Category-wise Tender Count")

    category_count = filtered_df["Category"].value_counts().reset_index()
    category_count.columns = ["Category", "Count"]

    if not category_count.empty:
        fig2 = px.pie(
            category_count,
            names="Category",
            values="Count",
            title="Category Distribution"
        )

        fig2.update_layout(template="plotly_dark")

        st.plotly_chart(fig2, use_container_width=True)


with tab2:
    show_table(best_df.sort_values(by=["Score", "Tender Value Number"], ascending=False), "⭐ Best Tenders")


with tab3:
    show_table(review_df.sort_values(by=["Score", "Tender Value Number"], ascending=False), "📋 Review Tenders")


with tab4:
    st.subheader("💰 Highest Value Projects")

    top_choice = st.selectbox("Select Top Projects", [10, 25, 50], index=2)

    top_value = filtered_df.sort_values(
        by="Tender Value Number",
        ascending=False
    ).head(top_choice)

    show_table(top_value, f"Top {top_choice} Highest Value Tenders")

    fig3 = px.bar(
        top_value.head(10),
        x="Tender Value Number",
        y="Tender Title / Ref / ID",
        orientation="h",
        title="Top 10 Highest Value Opportunities"
    )

    fig3.update_layout(
        template="plotly_dark"
    )

    st.plotly_chart(fig3, use_container_width=True)


with tab5:
    st.subheader("⏰ Closing Priority")

    if "Closing Date Parsed" not in filtered_df.columns:
        st.warning("Closing Date column not found.")
    else:
        today = pd.Timestamp(datetime.today().date())

        closing_df = filtered_df.dropna(subset=["Closing Date Parsed"]).copy()
        closing_df["Days Left"] = (closing_df["Closing Date Parsed"] - today).dt.days

        urgent = closing_df[(closing_df["Days Left"] >= 0) & (closing_df["Days Left"] <= 7)]
        urgent = urgent.sort_values(by=["Days Left", "Tender Value Number"], ascending=[True, False])

        today_df = urgent[urgent["Days Left"] == 0]
        three_df = urgent[urgent["Days Left"] <= 3]
        seven_df = urgent[urgent["Days Left"] <= 7]

        show_table(today_df, "🔴 Closing Today")
        show_table(three_df, "🟠 Closing Within 3 Days")
        show_table(seven_df, "🟡 Closing Within 7 Days")


with tab6:
    st.subheader("🔎 Filtered Data")
    show_table(filtered_df.sort_values(by=["Score", "Tender Value Number"], ascending=False), "All Filtered Tenders")


st.divider()
st.caption("GBP Infraprojects | Tender Intelligence Dashboard V7")