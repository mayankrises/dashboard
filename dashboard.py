"""
BPCL Uran - Document Control Index (DCI) Dashboard
Run with:  streamlit run dashboard.py
"""

import io
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(page_title="DCI Dashboard - BPCL Uran", layout="wide")

REV_BLOCK_START_COLS = [19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59]  # LATEST REV. col per block

APPROVED_CODES = {"1", "1 WITH COMMENTS"}
RETAINED_CODES = {"R"}
CODE2_CODES = {"2"}
CODE3_CODES = {"3"}
VOID_CODES = {"V"}

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

AGING_BUCKET_ORDER = ["Not due yet", "0-7 days", "8-14 days", "15-21 days", ">21 days"]

# Palette matched to the reference "Monthly Submission Chart" (purple / coral / mint),
# cycled if there are more than 3 years in the data.
CHART_PALETTE = ["#B4A7D6", "#F1948A", "#8EE4D0", "#f5c26b", "#8ab4f8", "#c9a0dc"]

RECORD_COLUMNS = [
    "sr_no", "doc_no", "title", "discipline", "poc_submission",
    "latest_rev", "latest_uploaded_date", "eil_review_status", "latest_code",
    "latest_from_eil_date", "category", "l4_marked", "schedule_submission_date",
    "status", "poc", "eil_mandays", "eil_manhours", "stspl_mandays",
    "stspl_manhours", "revision_cycle", "expected_manhours",
]

HISTORY_COLUMNS = ["doc_no", "rev_no", "uploaded_date", "code", "from_eil_date"]

NUMERIC_COLUMNS = [
    "eil_mandays", "eil_manhours", "stspl_mandays",
    "stspl_manhours", "revision_cycle", "expected_manhours",
]

ROOT_CAUSE_NOTES = [
    "Change in Tank Anchor Plate Design",
    "Dome Roof Rafter Section & Numbers and Compression plate thickness",
    "Suspended deck Design and Instrument Nozzle and Stillwell Size",
    "Intank Pump Selection and its document finalization",
    "Product Bottom Inlet Pipe Size reduction and Vapour Disengagement Design",
    "VMS Selection and its document finalization",
    "Tank Shell Insulation Arrangement change",
    "Tank Cool Down Sensor Arrangement change",
    "Change in Nos. of PSV & VSV and Change in their Platform arrangement",
    "Change in the Pump Roof Top maintenance platform arrangement",
    "Change in ROV arrangement at Liquid Platform and Multilevel Pipe Rack",
    "BOG line Orientation and Supporting on Shell and Trestle Design",
]


# ----------------------------------------------------------------------------
# DATA LOADING & PARSING
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_workbook_bytes(file_bytes):
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "DCI" not in wb.sheetnames:
        raise ValueError("This workbook does not contain a sheet named 'DCI'.")
    ws = wb["DCI"]

    project_name = ws.cell(row=2, column=2).value
    contractor = ws.cell(row=3, column=2).value

    records, history = [], []

    for r in range(9, ws.max_row + 1):
        doc_no = ws.cell(row=r, column=2).value
        if doc_no is None or str(doc_no).strip() == "":
            continue

        rec = {
            "sr_no": ws.cell(row=r, column=1).value,
            "doc_no": str(doc_no).strip(),
            "title": ws.cell(row=r, column=3).value,
            "discipline": ws.cell(row=r, column=4).value,
            "poc_submission": ws.cell(row=r, column=5).value,
            "latest_rev": ws.cell(row=r, column=10).value,
            "latest_uploaded_date": ws.cell(row=r, column=11).value,
            "eil_review_status": ws.cell(row=r, column=12).value,
            "latest_code": ws.cell(row=r, column=14).value,
            "latest_from_eil_date": ws.cell(row=r, column=15).value,
            "category": ws.cell(row=r, column=16).value,
            "l4_marked": ws.cell(row=r, column=17).value,
            "schedule_submission_date": ws.cell(row=r, column=18).value,
            "status": ws.cell(row=r, column=66).value,
            "poc": ws.cell(row=r, column=76).value,
            "eil_mandays": ws.cell(row=r, column=77).value,
            "eil_manhours": ws.cell(row=r, column=78).value,
            "stspl_mandays": ws.cell(row=r, column=79).value,
            "stspl_manhours": ws.cell(row=r, column=80).value,
            "revision_cycle": ws.cell(row=r, column=81).value,
            "expected_manhours": ws.cell(row=r, column=82).value,
        }
        records.append(rec)

        for start in REV_BLOCK_START_COLS:
            rev_no = ws.cell(row=r, column=start).value
            uploaded = ws.cell(row=r, column=start + 1).value
            code = ws.cell(row=r, column=start + 2).value
            from_eil = ws.cell(row=r, column=start + 3).value
            if rev_no is None and uploaded is None and code is None:
                continue
            history.append(
                {
                    "doc_no": str(doc_no).strip(),
                    "rev_no": rev_no,
                    "uploaded_date": uploaded,
                    "code": code,
                    "from_eil_date": from_eil,
                }
            )

    # NOTE: always pass explicit columns. If a workbook has no revision-history
    # cells populated at all, `history` would be an empty list and
    # pd.DataFrame([]) would have NO columns, which later crashes clean_dates()
    # with a KeyError. Explicit columns guarantee the shape is always correct.
    df = pd.DataFrame(records, columns=RECORD_COLUMNS)
    hist = pd.DataFrame(history, columns=HISTORY_COLUMNS)
    return df, hist, project_name, contractor


def clean_dates(df, cols):
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NaT
        else:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def normalize_code(code):
    """Collapse int/float/str/whitespace variations of an approval code to one canonical string."""
    if pd.isna(code):
        return ""
    if isinstance(code, (int, float)) and float(code).is_integer():
        return str(int(code))
    return " ".join(str(code).strip().upper().split())


def classify_code(code):
    value = normalize_code(code)
    if value in APPROVED_CODES:
        return "Approved (Code 1)"
    if value in RETAINED_CODES:
        return "Retained for Record (Code R)"
    if value in CODE2_CODES:
        return "Code 2 - Comments/Resubmit"
    if value in CODE3_CODES:
        return "Code 3 - Rejected/Resubmit"
    if value in VOID_CODES:
        return "Void"
    return "No code yet"


def normalize_status(value):
    """Normalize free-text status values and fold known variants onto a canonical label."""
    if pd.isna(value):
        return ""
    status = " ".join(str(value).strip().upper().split())
    aliases = {
        "WITH EIL FOR REVIEW": "WITH EIL",
        "PENDING WITH EIL": "WITH EIL",
        "FRESH SUBMISSION PENDING": "FRESH SUBMISSION",
        "CODE 1": "APPROVED",
    }
    return aliases.get(status, status)


def status_fallback(row):
    status = normalize_status(row.get("status"))
    if status:
        return status
    # Use pd.isna(), not `is None` -- pandas represents missing Excel cells as
    # NaN, and `np.nan is None` is False, which previously misclassified
    # blank-code rows as "WITH EIL" instead of "FRESH SUBMISSION".
    if pd.isna(row.get("latest_uploaded_date")):
        return "FRESH SUBMISSION"
    if row["code_bucket"] in ("Approved (Code 1)", "Retained for Record (Code R)"):
        return "APPROVED"
    if pd.notna(row.get("latest_from_eil_date")):
        return "RETURNED BY EIL"
    return "WITH EIL"


def aging_bucket(age):
    if pd.isna(age):
        return None
    if age < 0:
        return "Not due yet"
    if age <= 7:
        return "0-7 days"
    if age <= 14:
        return "8-14 days"
    if age <= 21:
        return "15-21 days"
    return ">21 days"


def derive_fields(df, review_period_days, today):
    df = df.copy()
    df = clean_dates(df, ["latest_uploaded_date", "latest_from_eil_date", "schedule_submission_date"])

    for col in NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["code_bucket"] = df["latest_code"].apply(classify_code)
    df["is_void"] = df["code_bucket"] == "Void"

    df["status_clean"] = df.apply(status_fallback, axis=1)

    df["is_submitted"] = df["latest_uploaded_date"].notna() | (df["status_clean"] != "FRESH SUBMISSION")
    df["is_reviewed"] = df["latest_from_eil_date"].notna()
    df["is_approved_final"] = df["code_bucket"].isin(
        ["Approved (Code 1)", "Retained for Record (Code R)"]
    )

    # Two distinct delay concepts that were previously conflated under one
    # "review period" threshold:
    #   - submission_delay: contractor is late submitting (vs schedule date)
    #   - review_age: EIL has had the document longer than the allowed review period
    df["submission_delay_days"] = np.where(
        df["latest_uploaded_date"].isna() & df["schedule_submission_date"].notna(),
        (today - df["schedule_submission_date"]).dt.days,
        np.nan,
    )
    df["review_age_days"] = np.where(
        df["latest_uploaded_date"].notna()
        & df["latest_from_eil_date"].isna()
        & ~df["is_approved_final"]
        & ~df["is_void"],
        (today - df["latest_uploaded_date"]).dt.days,
        np.nan,
    )
    df["is_submission_overdue"] = df["submission_delay_days"] > 0
    df["is_review_overdue"] = df["review_age_days"] > review_period_days

    # Legacy combined "age" used for the aging-bucket chart (whichever clock is running
    # for an open document: schedule delay if not yet submitted, else review age).
    def compute_age(row):
        if row["is_void"] or row["is_approved_final"]:
            return np.nan
        if row["status_clean"] == "FRESH SUBMISSION":
            base = row["schedule_submission_date"]
        else:
            base = row["latest_uploaded_date"]
        if pd.isna(base):
            return np.nan
        return (today - base).days

    df["age_days"] = df.apply(compute_age, axis=1)
    df["aging_bucket"] = df["age_days"].apply(aging_bucket)
    df["is_overdue"] = df["is_submission_overdue"].fillna(False) | df["is_review_overdue"].fillna(False)

    df["turnaround_days"] = (df["latest_from_eil_date"] - df["latest_uploaded_date"]).dt.days

    for c in ["latest_uploaded_date", "latest_from_eil_date", "schedule_submission_date"]:
        df[c + "_month"] = df[c].dt.to_period("M").astype(str)

    # Display-safe revision label: mixed numeric/text revisions (0, 1, "A", "IFC", ...)
    # cannot be sorted directly with Python's `sorted()`, so we normalize to strings
    # up front and sort those instead.
    df["latest_rev_display"] = df["latest_rev"].apply(
        lambda v: "" if pd.isna(v) else str(v).strip()
    )

    return df


def derive_history(hist):
    hist = hist.copy()
    hist = clean_dates(hist, ["uploaded_date", "from_eil_date"])
    hist["code_bucket"] = hist["code"].apply(classify_code)
    hist["turnaround_days"] = (hist["from_eil_date"] - hist["uploaded_date"]).dt.days

    invalid_turnaround = hist["turnaround_days"] < 0
    invalid_count = int(invalid_turnaround.sum())
    if invalid_count:
        hist.loc[invalid_turnaround, "turnaround_days"] = np.nan
    st.session_state["_invalid_turnaround_count"] = invalid_count

    hist["uploaded_month"] = hist["uploaded_date"].dt.to_period("M").astype(str)
    hist["returned_month"] = hist["from_eil_date"].dt.to_period("M").astype(str)
    return hist


# ----------------------------------------------------------------------------
# SIDEBAR - FILE UPLOAD
# ----------------------------------------------------------------------------
st.sidebar.title("DCI Dashboard")
uploaded_file = st.sidebar.file_uploader("Upload the DCI workbook (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.title("BPCL Uran - Document Control Index Dashboard")
    st.info("Upload a DCI workbook (.xlsx) from the sidebar to get started.")
    st.stop()

try:
    raw_df, raw_hist, project_name, contractor = load_workbook_bytes(uploaded_file.getvalue())
except Exception as e:
    st.error(f"Could not read this workbook: {e}")
    st.stop()

if raw_df.empty:
    st.warning("No document rows were found in the 'DCI' sheet.")
    st.stop()

# `data_only=True` only returns the last-cached formula result. If the workbook
# was produced/edited without Excel recalculating and saving it, formula-backed
# columns can silently come back blank. Give a visible warning rather than a
# dashboard that quietly shows zeros everywhere.
if raw_df["status"].isna().all():
    st.warning(
        "The Status column is empty for every row. If this workbook uses formulas, "
        "make sure it was recalculated and saved in Excel before uploading "
        "(cached formula results, not formulas themselves, are what gets read)."
    )

# ----------------------------------------------------------------------------
# SIDEBAR - SETTINGS & FILTERS
# ----------------------------------------------------------------------------
st.sidebar.markdown("---")
today = pd.Timestamp(datetime.today().date())
review_period_days = st.sidebar.number_input(
    "Allowed EIL review period (days)", min_value=1, max_value=90, value=21
)

df = derive_fields(raw_df, review_period_days, today)
hist = derive_history(raw_hist)

# "Reviewed" can mean either "the current/latest revision came back from EIL"
# (is_reviewed) or "this document has ever, at any revision, been reviewed"
# (ever_reviewed). These are different KPIs, so both are kept, separately named.
ever_reviewed_docs = set(hist.loc[hist["from_eil_date"].notna(), "doc_no"])
df["ever_reviewed"] = df["doc_no"].isin(ever_reviewed_docs)

st.sidebar.markdown("---")
st.sidebar.subheader("Filters")

disciplines = sorted([d for d in df["discipline"].dropna().unique()])
vendors = sorted([v for v in df["poc"].dropna().unique()])
statuses = sorted([s for s in df["status_clean"].dropna().unique()])
code_buckets = sorted([c for c in df["code_bucket"].dropna().unique()])
revisions = sorted(
    [r for r in df["latest_rev_display"].unique() if r != ""],
    key=str.casefold,
)
aging_buckets = AGING_BUCKET_ORDER

f_discipline = st.sidebar.multiselect("Discipline", disciplines)
f_vendor = st.sidebar.multiselect("Vendor / Consultant (POC)", vendors)
f_status = st.sidebar.multiselect("Status", statuses)
f_code = st.sidebar.multiselect("Approval code", code_buckets)
f_revision = st.sidebar.multiselect("Revision", revisions)
f_aging = st.sidebar.multiselect("Aging bucket (open docs only)", aging_buckets)
search = st.sidebar.text_input("Search document number / title")

date_cols_available = df["latest_uploaded_date"].dropna()
if not date_cols_available.empty:
    min_d, max_d = date_cols_available.min().date(), date_cols_available.max().date()
    f_date_range = st.sidebar.date_input(
        "Uploaded date range", value=(min_d, max_d), min_value=min_d, max_value=max_d
    )
    include_undated = st.sidebar.checkbox(
        "Include documents without an uploaded date", value=True
    )
else:
    f_date_range = None
    include_undated = True

filtered = df.copy()
if f_discipline:
    filtered = filtered[filtered["discipline"].isin(f_discipline)]
if f_vendor:
    filtered = filtered[filtered["poc"].isin(f_vendor)]
if f_status:
    filtered = filtered[filtered["status_clean"].isin(f_status)]
if f_code:
    filtered = filtered[filtered["code_bucket"].isin(f_code)]
if f_revision:
    filtered = filtered[filtered["latest_rev_display"].isin(f_revision)]
if f_aging:
    filtered = filtered[filtered["aging_bucket"].isin(f_aging)]
if search:
    s = search.lower()
    filtered = filtered[
        filtered["doc_no"].str.lower().str.contains(s, na=False)
        | filtered["title"].fillna("").str.lower().str.contains(s, na=False)
    ]
if f_date_range and isinstance(f_date_range, tuple) and len(f_date_range) == 2:
    start_d, end_d = pd.Timestamp(f_date_range[0]), pd.Timestamp(f_date_range[1])
    # Strict by default: only documents uploaded inside the chosen range match.
    # Undated documents are included only if the user opts in, so an "undated"
    # document can no longer silently survive an explicit date filter.
    date_match = filtered["latest_uploaded_date"].between(start_d, end_d)
    if include_undated:
        date_match = date_match | filtered["latest_uploaded_date"].isna()
    filtered = filtered[date_match]

# History rows must follow the same document-level filters, otherwise picking a
# discipline/vendor/status/search updates the document charts but the
# month-by-month history charts (submissions, reviews, turnaround) silently
# keep showing every document.
filtered_doc_numbers = set(filtered["doc_no"])
filtered_hist = hist[hist["doc_no"].isin(filtered_doc_numbers)].copy()

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
as_of_candidates = pd.concat(
    [df["latest_uploaded_date"], df["latest_from_eil_date"]]
).dropna()
as_of_date = as_of_candidates.max().strftime("%d %b %Y") if not as_of_candidates.empty else "N/A"

st.title(project_name or "Document Control Index Dashboard")
st.caption(f"Contractor: {contractor or 'N/A'}  |  Data as of: {as_of_date}  |  Total documents: {len(df)}")

tabs = st.tabs(
    [
        "Executive Summary",
        "Submission Progress",
        "Review Progress",
        "Approval Status",
        "Discipline Analysis",
        "Vendor Analysis",
        "Delays & Aging",
        "Manhours",
        "Document Explorer",
    ]
)

try:
    # ----------------------------------------------------------------------------
    # TAB 1: EXECUTIVE SUMMARY
    # ----------------------------------------------------------------------------
    with tabs[0]:
        total = len(filtered)
        submitted = int(filtered["is_submitted"].sum())
        reviewed = int(filtered["is_reviewed"].sum())
        approved1 = int((filtered["code_bucket"] == "Approved (Code 1)").sum())
        retained = int((filtered["code_bucket"] == "Retained for Record (Code R)").sum())
        under_review = int((filtered["status_clean"] == "WITH EIL").sum())
        pending_first_sub = int((filtered["status_clean"] == "FRESH SUBMISSION").sum())
        overdue = int(filtered["is_overdue"].sum())
    
        r1 = st.columns(5)
        r1[0].metric("Total Documents", total)
        r1[1].metric("Submitted", submitted, f"{submitted/total*100:.0f}%" if total else None)
        r1[2].metric("Reviewed (latest rev)", reviewed, f"{reviewed/total*100:.0f}%" if total else None)
        r1[3].metric("Approved (Code 1)", approved1, f"{approved1/total*100:.0f}%" if total else None)
        r1[4].metric("Retained for Record", retained)
    
        r2 = st.columns(5)
        r2[0].metric("Under Review (with EIL)", under_review)
        r2[1].metric("Pending First Submission", pending_first_sub)
        r2[2].metric("Code 2 (comments)", int((filtered["code_bucket"] == "Code 2 - Comments/Resubmit").sum()))
        r2[3].metric("Void", int(filtered["is_void"].sum()))
        r2[4].metric("Overdue (submission or review)", overdue, delta_color="inverse")
    
        st.markdown("---")
        c1, c2 = st.columns([2, 1])
        with c1:
            breakdown = filtered["code_bucket"].value_counts().reset_index()
            breakdown.columns = ["Code", "Count"]
            fig = px.pie(breakdown, names="Code", values="Count", title="Approval Status Breakdown", hole=0.4)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.subheader("Root Cause - Manhour Overshoot")
            st.caption("Carried over from the previous Excel dashboard (view-only).")
            for i, note in enumerate(ROOT_CAUSE_NOTES, 1):
                st.markdown(f"**{i}.** {note}")
    
    # ----------------------------------------------------------------------------
    # TAB 2: SUBMISSION PROGRESS
    # ----------------------------------------------------------------------------
    with tabs[1]:
        sub_hist = filtered_hist.dropna(subset=["uploaded_date"]).copy()
    
        if sub_hist.empty:
            st.subheader("Monthly Submission Chart")
            st.info("No submission history available for the current filters.")
        else:
            sub_hist["year"] = sub_hist["uploaded_date"].dt.year
            sub_hist["month_num"] = sub_hist["uploaded_date"].dt.month
    
            years = sorted(sub_hist["year"].unique(), reverse=True)
            grid_index = pd.MultiIndex.from_product([years, range(1, 13)], names=["year", "month_num"])
            grid = (
                sub_hist.groupby(["year", "month_num"]).size()
                .reindex(grid_index, fill_value=0)
                .reset_index(name="Submissions")
            )
            grid["month_name"] = grid["month_num"].apply(lambda m: MONTH_ORDER[m - 1])
            grid["year"] = grid["year"].astype(str)
    
            year_labels = [str(y) for y in years]
            color_map = {y: CHART_PALETTE[i % len(CHART_PALETTE)] for i, y in enumerate(year_labels)}
    
            with st.container(border=True):
                fig = px.bar(
                    grid,
                    x="month_name",
                    y="Submissions",
                    color="year",
                    barmode="group",
                    category_orders={"month_name": MONTH_ORDER, "year": year_labels},
                    color_discrete_map=color_map,
                )
                fig.update_traces(marker_line_color="rgba(255,255,255,0.6)", marker_line_width=1)
                fig.update_layout(
                    title=dict(
                        text="Monthly Submission Chart",
                        x=0.5,
                        xanchor="center",
                        font=dict(family="Arial Black, Arial, sans-serif", size=22, color="#595959"),
                    ),
                    yaxis_title="No. of Deliverables",
                    xaxis_title="",
                    yaxis=dict(gridcolor="#e6e6e6", zeroline=False),
                    xaxis=dict(showgrid=False),
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    bargap=0.28,
                    bargroupgap=0.04,
                    legend=dict(
                        orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5,
                        title_text="", font=dict(size=13),
                    ),
                    margin=dict(t=70, b=10, l=10, r=10),
                )
                st.plotly_chart(fig, use_container_width=True)
    
                # Data table styled to echo the reference image: a small colour swatch
                # next to each year label, bordered cells, months across the top.
                header_cells = "".join(
                    f"<th style='padding:5px 10px;border:1px solid #d9d9d9;text-align:center;'>{m}</th>"
                    for m in MONTH_ORDER
                )
                row_html = []
                for y in year_labels:
                    color = color_map[y]
                    series = grid.loc[grid["year"] == y].set_index("month_name").reindex(MONTH_ORDER)["Submissions"]
                    value_cells = "".join(
                        f"<td style='padding:5px 10px;border:1px solid #d9d9d9;text-align:center;'>{int(v)}</td>"
                        for v in series
                    )
                    row_html.append(
                        "<tr>"
                        f"<td style='padding:5px 10px;border:1px solid #d9d9d9;white-space:nowrap;'>"
                        f"<span style='display:inline-block;width:10px;height:10px;background:{color};"
                        f"border:1px solid #999;margin-right:6px;'></span>{y}</td>"
                        f"{value_cells}</tr>"
                    )
                table_html = f"""
                <table style='border-collapse:collapse;width:100%;font-size:14px;color:#333;margin-top:4px;'>
                    <tr><th style='padding:5px 10px;border:1px solid #d9d9d9;'></th>{header_cells}</tr>
                    {''.join(row_html)}
                </table>
                """
                st.markdown(table_html, unsafe_allow_html=True)
    
        st.markdown("---")
        st.subheader("Planned vs Actual First Submission")
    
        # Actual first-submission month must come from the earliest revision-history
        # upload date per document, not `latest_uploaded_date` (which reflects the
        # most recent revision and would push a document's "first submission" out
        # to whenever it was last revised).
        first_submission = (
            filtered_hist.dropna(subset=["uploaded_date"])
            .groupby("doc_no", as_index=False)["uploaded_date"].min()
            .rename(columns={"uploaded_date": "first_uploaded_date"})
        )
    
        planned = filtered.dropna(subset=["schedule_submission_date"]).copy()
        planned["planned_month"] = planned["schedule_submission_date"].dt.to_period("M").astype(str)
        planned_counts = planned.groupby("planned_month").size().reset_index(name="Planned")
    
        actual = filtered.merge(first_submission, on="doc_no", how="left").dropna(subset=["first_uploaded_date"])
        actual["actual_month"] = actual["first_uploaded_date"].dt.to_period("M").astype(str)
        actual_counts = actual.groupby("actual_month").size().reset_index(name="Actual")
    
        merged = pd.merge(
            planned_counts, actual_counts, left_on="planned_month", right_on="actual_month", how="outer"
        )
        merged["month"] = merged["planned_month"].combine_first(merged["actual_month"])
        merged = merged[["month", "Planned", "Actual"]].fillna(0).sort_values("month")
        fig2 = px.line(
            merged, x="month", y=["Planned", "Actual"], markers=True,
            title="Planned vs Actual First Submission by Month",
        )
        st.plotly_chart(fig2, use_container_width=True)
    
        cum = merged.copy()
        cum["Planned Cum"] = cum["Planned"].cumsum()
        cum["Actual Cum"] = cum["Actual"].cumsum()
        fig3 = px.line(cum, x="month", y=["Planned Cum", "Actual Cum"], markers=True, title="Cumulative Submission Progress")
        st.plotly_chart(fig3, use_container_width=True)
    
    # ----------------------------------------------------------------------------
    # TAB 3: REVIEW PROGRESS
    # ----------------------------------------------------------------------------
    with tabs[2]:
        st.subheader("Monthly Reviews Completed (EIL return)")
        rev_hist = filtered_hist.dropna(subset=["from_eil_date"])
        monthly_rev = rev_hist.groupby("returned_month").size().reset_index(name="Reviews").sort_values("returned_month")
        fig = px.bar(monthly_rev, x="returned_month", y="Reviews", title="Documents reviewed per month (all revisions)")
        st.plotly_chart(fig, use_container_width=True)
    
        st.subheader("Submission-to-Review Turnaround")
        invalid_count = st.session_state.get("_invalid_turnaround_count", 0)
        if invalid_count:
            st.warning(
                f"{invalid_count} revision record(s) had an EIL return date earlier than their "
                "upload date and were excluded from turnaround calculations."
            )
        turn = filtered_hist.dropna(subset=["turnaround_days"])
        if not turn.empty:
            avg_turn = turn.groupby("returned_month")["turnaround_days"].mean().reset_index()
            fig2 = px.line(avg_turn, x="returned_month", y="turnaround_days", markers=True,
                            title="Average Turnaround Time (days) by Month")
            st.plotly_chart(fig2, use_container_width=True)
            st.metric("Average turnaround (all revisions, days)", f"{turn['turnaround_days'].mean():.1f}")
        else:
            st.info("No turnaround data available yet.")
    
    # ----------------------------------------------------------------------------
    # TAB 4: APPROVAL STATUS
    # ----------------------------------------------------------------------------
    with tabs[3]:
        st.subheader("Approval Code Breakdown (current / latest revision)")
        code_counts = filtered["code_bucket"].value_counts().reset_index()
        code_counts.columns = ["Code", "Count"]
        fig = px.bar(code_counts, x="Code", y="Count", title="Documents by Approval Code", text="Count")
        st.plotly_chart(fig, use_container_width=True)
    
        st.subheader("Approval Code by Revision (full history)")
        rc = filtered_hist.dropna(subset=["rev_no"]).groupby(["rev_no", "code_bucket"]).size().reset_index(name="Count")
        fig2 = px.bar(rc, x="rev_no", y="Count", color="code_bucket", title="Code outcome at each revision number", barmode="stack")
        st.plotly_chart(fig2, use_container_width=True)
    
        st.subheader("Revision Distribution")
        rev_dist = (
            filtered.loc[filtered["latest_rev_display"] != "", "latest_rev_display"]
            .value_counts()
            .reset_index()
        )
        rev_dist.columns = ["Latest Revision", "Count"]
        rev_dist = rev_dist.sort_values("Latest Revision", key=lambda s: s.str.casefold())
        fig3 = px.bar(rev_dist, x="Latest Revision", y="Count", title="Documents by Current Revision Number")
        st.plotly_chart(fig3, use_container_width=True)
    
    # ----------------------------------------------------------------------------
    # TAB 5: DISCIPLINE ANALYSIS
    # ----------------------------------------------------------------------------
    with tabs[4]:
        st.subheader("Documents by Discipline")
        disc_counts = filtered["discipline"].value_counts().reset_index()
        disc_counts.columns = ["Discipline", "Count"]
        fig = px.bar(disc_counts.sort_values("Count"), x="Count", y="Discipline", orientation="h",
                     title="Total Documents by Discipline", height=600)
        st.plotly_chart(fig, use_container_width=True)
    
        st.subheader("Pending Documents by Discipline")
        pending = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
        pend_disc = pending["discipline"].value_counts().reset_index()
        pend_disc.columns = ["Discipline", "Pending"]
        fig2 = px.bar(pend_disc.sort_values("Pending"), x="Pending", y="Discipline", orientation="h",
                      title="Pending Documents by Discipline", height=600)
        st.plotly_chart(fig2, use_container_width=True)
    
        st.subheader("Overdue Documents by Discipline")
        overdue_df = filtered[filtered["is_overdue"] == True]
        od_disc = overdue_df["discipline"].value_counts().reset_index()
        od_disc.columns = ["Discipline", "Overdue"]
        if not od_disc.empty:
            fig3 = px.bar(od_disc.sort_values("Overdue"), x="Overdue", y="Discipline", orientation="h",
                          title="Overdue Documents by Discipline", height=500)
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("No overdue documents for the current filter.")
    
    # ----------------------------------------------------------------------------
    # TAB 6: VENDOR ANALYSIS
    # ----------------------------------------------------------------------------
    with tabs[5]:
        st.subheader("Documents by Vendor / Consultant")
        vend_counts = filtered["poc"].value_counts().reset_index()
        vend_counts.columns = ["Vendor / POC", "Count"]
        fig = px.bar(vend_counts.sort_values("Count").tail(20), x="Count", y="Vendor / POC", orientation="h",
                     title="Documents by Vendor (Top 20)", height=600)
        st.plotly_chart(fig, use_container_width=True)
    
        st.subheader("Top Vendors with Pending Documents")
        pending = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
        pend_vend = pending["poc"].value_counts().reset_index().head(15)
        pend_vend.columns = ["Vendor / POC", "Pending"]
        fig2 = px.bar(pend_vend.sort_values("Pending"), x="Pending", y="Vendor / POC", orientation="h",
                      title="Top 15 Vendors by Pending Documents", height=500)
        st.plotly_chart(fig2, use_container_width=True)
    
    # ----------------------------------------------------------------------------
    # TAB 7: DELAYS & AGING
    # ----------------------------------------------------------------------------
    with tabs[6]:
        open_docs = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
        st.subheader("Aging Buckets (open documents)")
        bucket_counts = open_docs["aging_bucket"].value_counts().reindex(AGING_BUCKET_ORDER).fillna(0).reset_index()
        bucket_counts.columns = ["Aging Bucket", "Count"]
        fig = px.bar(bucket_counts, x="Aging Bucket", y="Count", text="Count", title="Open Documents by Aging Bucket")
        st.plotly_chart(fig, use_container_width=True)
    
        m = st.columns(3)
        m[0].metric("Total open (pending) documents", len(open_docs))
        m[1].metric("Late First Submissions", int(filtered["is_submission_overdue"].fillna(False).sum()))
        m[2].metric(f"EIL Reviews Overdue (> {review_period_days} days)", int(filtered["is_review_overdue"].fillna(False).sum()))
        st.caption(
            "\"Late First Submissions\" compares today against the scheduled submission date for "
            "documents not yet uploaded. \"EIL Reviews Overdue\" compares today against the upload "
            "date for documents still awaiting EIL return. These were previously combined under a "
            "single threshold, which mixed two different kinds of delay."
        )
    
        st.subheader("Overdue Documents")
        overdue_table = open_docs[open_docs["is_overdue"] == True][
            ["doc_no", "title", "discipline", "poc", "status_clean", "age_days", "schedule_submission_date", "latest_uploaded_date"]
        ].sort_values("age_days", ascending=False).copy()
        for date_col in ["schedule_submission_date", "latest_uploaded_date"]:
            overdue_table[date_col] = overdue_table[date_col].dt.strftime("%d-%b-%Y").fillna("")
        st.dataframe(overdue_table, use_container_width=True, hide_index=True)
    
    # ----------------------------------------------------------------------------
    # TAB 8: MANHOURS
    # ----------------------------------------------------------------------------
    with tabs[7]:
        total_expected = filtered["expected_manhours"].fillna(0).sum()
        total_actual = (filtered["eil_manhours"].fillna(0) + filtered["stspl_manhours"].fillna(0)).sum()
        variance = total_actual - total_expected
    
        c = st.columns(3)
        c[0].metric("Planned Manhours (Expected)", f"{total_expected:,.0f}")
        c[1].metric("Actual Manhours (EIL + STSPL)", f"{total_actual:,.0f}")
        c[2].metric("Variance", f"{variance:,.0f}", delta_color="inverse")
    
        st.subheader("Manhours by Discipline")
        mh_disc = filtered.groupby("discipline")[["expected_manhours", "eil_manhours", "stspl_manhours"]].sum().reset_index()
        mh_disc["actual_manhours"] = mh_disc["eil_manhours"].fillna(0) + mh_disc["stspl_manhours"].fillna(0)
        fig = px.bar(mh_disc, x="discipline", y=["expected_manhours", "actual_manhours"], barmode="group",
                     title="Planned vs Actual Manhours by Discipline")
        st.plotly_chart(fig, use_container_width=True)
    
        st.subheader("Manhours by Vendor")
        mh_vend = filtered.groupby("poc")[["expected_manhours", "eil_manhours", "stspl_manhours"]].sum().reset_index()
        mh_vend["actual_manhours"] = mh_vend["eil_manhours"].fillna(0) + mh_vend["stspl_manhours"].fillna(0)
        mh_vend = mh_vend.sort_values("actual_manhours", ascending=False).head(15)
        fig2 = px.bar(mh_vend, x="poc", y=["expected_manhours", "actual_manhours"], barmode="group",
                      title="Planned vs Actual Manhours - Top 15 Vendors")
        st.plotly_chart(fig2, use_container_width=True)
    
        st.markdown("---")
        st.subheader("Root Cause Analysis for Manhour Overshoot")
        st.caption("Carried over from the original Excel dashboard (view-only reference list).")
        for i, note in enumerate(ROOT_CAUSE_NOTES, 1):
            st.markdown(f"**{i}.** {note}")
    
    # ----------------------------------------------------------------------------
    # TAB 9: DOCUMENT EXPLORER
    # ----------------------------------------------------------------------------
    with tabs[8]:
        st.subheader(f"Document Explorer ({len(filtered)} documents matching filters)")
    
        display_cols = [
            "sr_no", "doc_no", "title", "discipline", "poc", "latest_rev_display", "status_clean",
            "code_bucket", "schedule_submission_date", "latest_uploaded_date", "latest_from_eil_date",
            "age_days", "aging_bucket", "is_overdue",
        ]
        explorer_df = filtered[display_cols].rename(columns={"latest_rev_display": "latest_rev"}).copy()
    
        # NOTE: a row-wise pandas Styler (`.style.apply(..., axis=1)`) on a
        # few-hundred-row table, re-built on every single sidebar interaction,
        # is a well-known source of native Arrow-serialization crashes in
        # Streamlit ("Python quit unexpectedly" rather than a normal error page).
        # A plain, unstyled dataframe with an explicit "Overdue" flag column is
        # just as informative and avoids that failure path entirely.
        explorer_df["Overdue"] = explorer_df["is_overdue"].map({True: "⚠️ Yes", False: ""})
        explorer_df = explorer_df.drop(columns=["is_overdue"]).sort_values(
            ["Overdue", "doc_no"], ascending=[False, True]
        )
    
        # Dates render as plain strings rather than raw Timestamp/NaT objects,
        # which keeps the frame a simple, uniformly-typed table for Arrow to
        # serialize (Timestamp columns mixed with NaT + object dtypes elsewhere
        # in the same table is the other common trigger for this kind of crash).
        for date_col in ["schedule_submission_date", "latest_uploaded_date", "latest_from_eil_date"]:
            explorer_df[date_col] = explorer_df[date_col].dt.strftime("%d-%b-%Y").fillna("")
    
        st.dataframe(
            explorer_df,
            use_container_width=True,
            hide_index=True,
            height=600,
        )
    
        csv_bytes = explorer_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download filtered rows as CSV", data=csv_bytes, file_name="dci_filtered_export.csv", mime="text/csv"
        )
except Exception as e:
    st.error(
        "Something went wrong while rendering the dashboard tabs. "
        "The details below should help pin down the cause."
    )
    st.exception(e)