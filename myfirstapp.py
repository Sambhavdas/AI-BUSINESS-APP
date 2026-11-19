
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
import hashlib
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="AI Business Intelligence",
    page_icon="📊",
    layout="wide"
)

# =====================================================
# SESSION STATE
# =====================================================

STATE_DEFAULTS = {
    "source_key": None,
    "source_df": None,
    "cleaned_df": None,
    "cleaning_preview_df": None,
    "cleaning_log": [],
    "cleaning_approved": False,
    "review_approved": False,
    "correction_preview_df": None,
    "correction_description": "",
    "correction_log": [],
    "pending_feedback": ""
}

for key, value in STATE_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


def reset_workflow():
    for key, value in STATE_DEFAULTS.items():
        st.session_state[key] = value


def read_uploaded_file(file_name, file_bytes, sheet_name=None):
    buffer = BytesIO(file_bytes)

    if file_name.lower().endswith(".csv"):
        return pd.read_csv(buffer)

    return pd.read_excel(buffer, sheet_name=sheet_name)


def get_excel_sheet_names(file_name, file_bytes):
    if file_name.lower().endswith(".csv"):
        return ["CSV Data"]

    return pd.ExcelFile(BytesIO(file_bytes)).sheet_names


def quality_report(df):
    return pd.DataFrame({
        "Column": df.columns.astype(str),
        "Data Type": df.dtypes.astype(str).values,
        "Missing Values": df.isna().sum().values,
        "Unique Values": df.nunique().values
    })


def format_number(value):
    if pd.isna(value):
        return "N/A"
    return f"{value:,.2f}"


# =====================================================
# HEADER
# =====================================================

st.title("📊 AI Business Intelligence Platform")
st.caption(
    "One platform for data quality, cleaning, "
    "business analytics, and reporting."
)

with st.sidebar:
    st.header("Workflow")
    st.write("1. Upload Multiple Files / Sheets")
    st.write("2. Append / Merge and Data Quality")
    st.write("3. Approve Cleaning")
    st.write("4. Review and Correct")
    st.write("5. Business Dashboard")
    st.write("6. Advanced Analysis")
    st.write("7. Business Insights")
    st.write("8. Data Explorer and Export")


# =====================================================
# STEP 1: UPLOAD
# =====================================================
# STEP 1: MULTI-FILE AND MULTI-SHEET UPLOAD
# =====================================================

st.header("Step 1 — Upload Data")

uploaded_files = st.file_uploader(
    "Choose one or more CSV or Excel files",
    type=["csv", "xlsx"],
    accept_multiple_files=True,
    help="Upload multiple workbooks and select the sheets you want to analyse."
)

if not uploaded_files:
    st.info("Upload one or more files to start the analysis workflow.")
    col1, col2, col3 = st.columns(3)
    col1.metric("Data Quality", "Ready")
    col2.metric("Business Dashboard", "Ready")
    col3.metric("Report Export", "Ready")
    st.stop()

file_sheet_options = {}
file_bytes_map = {}

try:
    for uploaded in uploaded_files:
        raw_bytes = uploaded.getvalue()
        file_bytes_map[uploaded.name] = raw_bytes
        file_sheet_options[uploaded.name] = get_excel_sheet_names(
            uploaded.name, raw_bytes
        )
except Exception as error:
    st.error(f"Could not inspect the uploaded files: {error}")
    st.stop()

st.subheader("Select Sheets")
selected_tables = []

for uploaded in uploaded_files:
    filename = uploaded.name
    available_sheets = file_sheet_options[filename]

    if filename.lower().endswith(".csv"):
        chosen_sheets = ["CSV Data"]
        st.caption(f"{filename} — CSV data")
    else:
        chosen_sheets = st.multiselect(
            f"Sheets to include from {filename}",
            options=available_sheets,
            default=available_sheets,
            key=f"sheets_{hashlib.md5(filename.encode()).hexdigest()}"
        )

    for sheet_name in chosen_sheets:
        try:
            frame = read_uploaded_file(
                filename,
                file_bytes_map[filename],
                None if sheet_name == "CSV Data" else sheet_name
            )
            if frame is not None:
                frame = frame.copy()
                frame["_Source_File"] = filename
                frame["_Source_Sheet"] = sheet_name
                selected_tables.append({
                    "label": f"{filename} | {sheet_name}",
                    "data": frame
                })
        except Exception as error:
            st.error(f"Could not read '{sheet_name}' from {filename}: {error}")
            st.stop()

if not selected_tables:
    st.warning("Select at least one sheet to continue.")
    st.stop()

st.subheader("Combine Selected Data")
combine_mode = st.radio(
    "How should the selected sheets be combined?",
    ["Append Rows", "Merge Tables"],
    horizontal=True,
    help=(
        "Append stacks similar tables vertically. Merge joins different "
        "tables using a shared key such as Customer ID or Product ID."
    )
)

merge_key = None
join_type = "inner"

if combine_mode == "Merge Tables":
    common_columns = set(selected_tables[0]["data"].columns)
    for table in selected_tables[1:]:
        common_columns &= set(table["data"].columns)

    common_columns = sorted(
        col for col in common_columns
        if not str(col).startswith("_Source_")
    )

    if len(selected_tables) < 2:
        st.warning("Merge requires at least two selected sheets.")
        st.stop()

    if not common_columns:
        st.error(
            "No common column was found across all selected sheets. "
            "Choose Append Rows or ensure the tables share a key column."
        )
        st.stop()

    merge_key = st.selectbox(
        "Select the common key column",
        common_columns,
        help="Choose a matching ID column, such as Order ID or Product ID."
    )
    join_type = st.selectbox(
        "Merge type",
        ["left", "inner", "outer", "right"],
        index=0
    )

st.caption(f"{len(selected_tables)} sheet(s) selected for combining.")

input_signature = tuple(
    (
        uploaded.name,
        hashlib.sha256(file_bytes_map[uploaded.name]).hexdigest()
    )
    for uploaded in uploaded_files
)

source_key = (
    input_signature,
    tuple(table["label"] for table in selected_tables),
    combine_mode,
    merge_key,
    join_type
)

try:
    if combine_mode == "Append Rows":
        combined_df = pd.concat(
            [table["data"] for table in selected_tables],
            ignore_index=True,
            sort=False
        )
    else:
        combined_df = selected_tables[0]["data"].copy()
        for index, table in enumerate(selected_tables[1:], start=1):
            combined_df = combined_df.merge(
                table["data"],
                on=merge_key,
                how=join_type,
                suffixes=("", f"_table{index + 1}")
            )
except Exception as error:
    st.error(f"Could not combine the selected sheets: {error}")
    st.stop()

if st.session_state.source_key != source_key:
    reset_workflow()
    st.session_state.source_key = source_key
    st.session_state.source_df = combined_df.copy()

source_df = st.session_state.source_df.copy()
source_label = ", ".join(uploaded.name for uploaded in uploaded_files)

if source_df.empty:
    st.warning("The selected files and sheets contain no data rows.")
    st.stop()

st.success(
    f"Combined dataset ready: {len(selected_tables)} sheet(s), "
    f"{len(source_df):,} rows and {len(source_df.columns):,} columns."
)

with st.expander("Selected source sheets", expanded=False):
    st.write([table["label"] for table in selected_tables])

st.subheader("Combined Data Preview")
st.dataframe(source_df.head(100), use_container_width=True)


# =====================================================
# STEP 3: CLEANING PREVIEW AND APPROVAL
# =====================================================

st.header("Step 3 — Cleaning Preview and Approval")

if not st.session_state.cleaning_approved:

    st.write(
        "Select the cleaning actions you want to preview. "
        "The original uploaded data will remain unchanged."
    )

    remove_duplicates = st.checkbox(
        "Remove duplicate rows",
        value=True,
        key="clean_duplicates"
    )

    fill_numeric = st.checkbox(
        "Fill missing numeric values with the column median",
        value=False,
        key="clean_numeric"
    )

    fill_text = st.checkbox(
        "Fill missing text values with the most frequent value",
        value=False,
        key="clean_text"
    )

    drop_missing_rows = st.checkbox(
        "Remove rows containing missing values",
        value=False,
        key="clean_drop_missing"
    )

    if st.button("Generate Cleaning Preview", type="primary"):

        proposed_df = source_df.copy()
        log = []

        if remove_duplicates:
            before = len(proposed_df)
            proposed_df = proposed_df.drop_duplicates()
            removed = before - len(proposed_df)
            log.append(f"Duplicate rows removed: {removed}")

        if drop_missing_rows:
            before = len(proposed_df)
            proposed_df = proposed_df.dropna()
            removed = before - len(proposed_df)
            log.append(f"Rows with missing values removed: {removed}")

        if fill_numeric:
            numeric_cols = proposed_df.select_dtypes(
                include=np.number
            ).columns

            for col in numeric_cols:
                missing_count = int(proposed_df[col].isna().sum())
                median_value = proposed_df[col].median()

                if missing_count and pd.notna(median_value):
                    proposed_df[col] = proposed_df[col].fillna(
                        median_value
                    )
                    log.append(
                        f"Filled {missing_count} missing value(s) "
                        f"in '{col}' using the median."
                    )

        if fill_text:
            text_cols = proposed_df.select_dtypes(
                include=["object", "string", "category"]
            ).columns

            for col in text_cols:
                missing_count = int(proposed_df[col].isna().sum())
                modes = proposed_df[col].mode(dropna=True)

                if missing_count and not modes.empty:
                    proposed_df[col] = proposed_df[col].fillna(
                        modes.iloc[0]
                    )
                    log.append(
                        f"Filled {missing_count} missing value(s) "
                        f"in '{col}' using the most frequent value."
                    )

        st.session_state.cleaning_preview_df = proposed_df
        st.session_state.cleaning_log = log

    if st.session_state.cleaning_preview_df is not None:

        proposed_df = st.session_state.cleaning_preview_df

        st.subheader("Proposed Cleaning Results")

        p1, p2, p3 = st.columns(3)
        p1.metric("Original Rows", f"{len(source_df):,}")
        p2.metric("Proposed Rows", f"{len(proposed_df):,}")
        p3.metric(
            "Proposed Missing Cells",
            f"{int(proposed_df.isna().sum().sum()):,}"
        )

        if st.session_state.cleaning_log:
            for item in st.session_state.cleaning_log:
                st.write(f"- {item}")
        else:
            st.info("No changes are needed for the selected actions.")

        st.dataframe(
            proposed_df.head(100),
            use_container_width=True
        )

        approve_col, cancel_col = st.columns(2)

        with approve_col:
            if st.button("Approve and Apply Cleaning", type="primary"):
                st.session_state.cleaned_df = proposed_df.copy()
                st.session_state.cleaning_approved = True
                st.session_state.review_approved = False
                st.session_state.correction_preview_df = None
                st.rerun()

        with cancel_col:
            if st.button("Discard Cleaning Preview"):
                st.session_state.cleaning_preview_df = None
                st.session_state.cleaning_log = []
                st.rerun()

    st.stop()


# =====================================================
# STEP 4: REVIEW AND CORRECTION
# =====================================================

st.header("Step 4 — Review Cleaned Data")

cleaned_df = st.session_state.cleaned_df.copy()

st.success("Cleaning has been applied after your approval.")

r1, r2, r3 = st.columns(3)
r1.metric("Original Rows", f"{len(source_df):,}")
r2.metric("Cleaned Rows", f"{len(cleaned_df):,}")
r3.metric(
    "Remaining Missing Cells",
    f"{int(cleaned_df.isna().sum().sum()):,}"
)

with st.expander("View Cleaning Summary", expanded=True):
    for item in st.session_state.cleaning_log:
        st.write(f"- {item}")

st.subheader("Cleaned Data Preview")
st.dataframe(cleaned_df.head(100), use_container_width=True)

st.subheader("Remaining Data Quality Issues")
st.dataframe(quality_report(cleaned_df), use_container_width=True)

st.subheader("Report a Problem")

st.text_area(
    "Describe the problem in English",
    placeholder=(
        "Example: The Sales column contains negative values "
        "that should be reviewed."
    ),
    key="pending_feedback"
)

columns = cleaned_df.columns.tolist()

correction_column = st.selectbox(
    "Select the column to correct",
    columns,
    key="correction_column"
)

correction_action = st.selectbox(
    "Select a proposed correction",
    [
        "Fill missing values",
        "Trim extra spaces",
        "Convert column to numeric",
        "Convert column to date",
        "Replace negative numeric values with zero",
        "Remove rows with missing values in this column",
        "Remove duplicate rows from the dataset"
    ],
    key="correction_action"
)

st.caption(
    "Your problem description is recorded as feedback. "
    "Choose a specific correction action below; the app will "
    "preview the result rather than guessing what you meant."
)

if st.button("Generate Correction Preview"):

    proposed = cleaned_df.copy()
    description = ""

    try:
        if correction_action == "Fill missing values":
            if pd.api.types.is_numeric_dtype(proposed[correction_column]):
                median_value = proposed[correction_column].median()

                if pd.isna(median_value):
                    raise ValueError(
                        "No valid numeric values are available "
                        "to calculate a median."
                    )

                proposed[correction_column] = proposed[
                    correction_column
                ].fillna(median_value)

                description = (
                    f"Fill missing values in '{correction_column}' "
                    "using the column median."
                )
            else:
                modes = proposed[correction_column].mode(dropna=True)

                if modes.empty:
                    raise ValueError(
                        "No existing values are available "
                        "to fill missing values."
                    )

                proposed[correction_column] = proposed[
                    correction_column
                ].fillna(modes.iloc[0])

                description = (
                    f"Fill missing values in '{correction_column}' "
                    "using the most frequent value."
                )

        elif correction_action == "Trim extra spaces":
            if not (
                pd.api.types.is_object_dtype(proposed[correction_column])
                or pd.api.types.is_string_dtype(proposed[correction_column])
            ):
                raise ValueError(
                    "Select a text column for this action."
                )

            proposed[correction_column] = proposed[
                correction_column
            ].map(
                lambda value: value.strip()
                if isinstance(value, str) else value
            )

            description = (
                f"Trim leading and trailing spaces in "
                f"'{correction_column}'."
            )

        elif correction_action == "Convert column to numeric":
            converted = pd.to_numeric(
                proposed[correction_column],
                errors="coerce"
            )

            newly_missing = int(
                (
                    proposed[correction_column].notna()
                    & converted.isna()
                ).sum()
            )

            if newly_missing:
                st.warning(
                    f"{newly_missing} value(s) could not be converted "
                    "and would become missing. Review the preview carefully."
                )

            proposed[correction_column] = converted
            description = (
                f"Convert '{correction_column}' to numeric values."
            )

        elif correction_action == "Convert column to date":
            converted = pd.to_datetime(
                proposed[correction_column],
                errors="coerce"
            )

            newly_missing = int(
                (
                    proposed[correction_column].notna()
                    & converted.isna()
                ).sum()
            )

            if newly_missing:
                st.warning(
                    f"{newly_missing} value(s) could not be parsed "
                    "and would become missing."
                )

            proposed[correction_column] = converted
            description = (
                f"Convert '{correction_column}' to dates."
            )

        elif correction_action == (
            "Replace negative numeric values with zero"
        ):
            if not pd.api.types.is_numeric_dtype(
                proposed[correction_column]
            ):
                raise ValueError(
                    "Select a numeric column for this action."
                )

            negative_count = int(
                (proposed[correction_column] < 0).sum()
            )

            proposed.loc[
                proposed[correction_column] < 0,
                correction_column
            ] = 0

            description = (
                f"Replace {negative_count} negative value(s) "
                f"in '{correction_column}' with zero."
            )

        elif correction_action == (
            "Remove rows with missing values in this column"
        ):
            before = len(proposed)
            proposed = proposed.dropna(
                subset=[correction_column]
            )

            description = (
                f"Remove {before - len(proposed)} row(s) with "
                f"missing values in '{correction_column}'."
            )

        elif correction_action == (
            "Remove duplicate rows from the dataset"
        ):
            before = len(proposed)
            proposed = proposed.drop_duplicates()

            description = (
                f"Remove {before - len(proposed)} duplicate row(s)."
            )

        st.session_state.correction_preview_df = proposed
        st.session_state.correction_description = description

    except Exception as error:
        st.error(f"Could not prepare correction: {error}")


if st.session_state.correction_preview_df is not None:

    proposed = st.session_state.correction_preview_df

    st.subheader("Proposed Correction")
    st.write(st.session_state.correction_description)

    if st.session_state.pending_feedback.strip():
        st.write(
            "**Your problem report:**",
            st.session_state.pending_feedback
        )

    st.write("Review the proposed data before applying it.")

    st.dataframe(proposed.head(100), use_container_width=True)

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Approve and Apply Correction", type="primary"):
            st.session_state.cleaned_df = proposed.copy()
            st.session_state.correction_log.append(
                st.session_state.correction_description
            )
            st.session_state.correction_preview_df = None
            st.session_state.correction_description = ""
            st.session_state.review_approved = False
            st.rerun()

    with c2:
        if st.button("Reject Correction"):
            st.session_state.correction_preview_df = None
            st.session_state.correction_description = ""
            st.rerun()


if st.session_state.correction_log:
    with st.expander("Correction History"):
        for item in st.session_state.correction_log:
            st.write(f"- {item}")


st.divider()

if not st.session_state.review_approved:
    st.write(
        "Review the cleaned data. Approve only when you are "
        "satisfied with the results."
    )

    if st.button("Approve Data and Continue to Dashboard", type="primary"):
        st.session_state.review_approved = True
        st.rerun()

else:
    st.success("Data review approved. Analytics are now unlocked.")


# =====================================================
# STEP 5: BUSINESS DASHBOARD
# =====================================================

if st.session_state.review_approved:

    dashboard_source = st.session_state.cleaned_df.copy()

    st.header("Step 5 — Business Dashboard")

    all_columns = dashboard_source.columns.tolist()

    numeric_columns = dashboard_source.select_dtypes(
        include=np.number
    ).columns.tolist()

    if not numeric_columns:
        st.warning(
            "No numeric columns were found. Some business metrics "
            "and charts require numeric data."
        )

    def find_column(words, available):
        for col in available:
            if any(word in str(col).lower() for word in words):
                return col
        return None

    sales_candidate = find_column(
        ["sales", "revenue", "amount", "total"],
        numeric_columns
    )

    profit_candidate = find_column(
        ["profit", "net income", "earnings"],
        numeric_columns
    )

    order_candidate = find_column(
        ["order id", "order_id", "transaction id", "invoice id"],
        all_columns
    )

    quantity_candidate = find_column(
        ["quantity", "qty", "units", "items"],
        numeric_columns
    )

    date_candidate = find_column(
        ["date", "order date", "transaction date"],
        all_columns
    )

    category_candidates = [
        col for col in all_columns
        if (
            not str(col).startswith("_Source_")
            and not pd.api.types.is_numeric_dtype(dashboard_source[col])
            and dashboard_source[col].nunique(dropna=True) <= 100
        )
    ]

    sales_options = ["Not selected"] + numeric_columns
    profit_options = ["Not selected"] + numeric_columns
    order_options = ["Not selected"] + all_columns
    quantity_options = ["Not selected"] + numeric_columns

    sales_index = (
        sales_options.index(sales_candidate)
        if sales_candidate in sales_options else 0
    )

    profit_index = (
        profit_options.index(profit_candidate)
        if profit_candidate in profit_options else 0
    )

    order_index = (
        order_options.index(order_candidate)
        if order_candidate in order_options else 0
    )

    quantity_index = (
        quantity_options.index(quantity_candidate)
        if quantity_candidate in quantity_options else 0
    )

    metric1, metric2, metric3, metric4 = st.columns(4)

    with metric1:
        sales_column = st.selectbox(
            "Sales / Revenue Column",
            sales_options,
            index=sales_index
        )

    with metric2:
        profit_column = st.selectbox(
            "Profit Column",
            profit_options,
            index=profit_index
        )

    with metric3:
        order_column = st.selectbox(
            "Order ID Column",
            order_options,
            index=order_index
        )

    with metric4:
        quantity_column = st.selectbox(
            "Quantity Column",
            quantity_options,
            index=quantity_index
        )

    # -------------------------------------------------
    # FILTERS
    # -------------------------------------------------

    st.subheader("Dashboard Filters")

    filtered_df = dashboard_source.copy()

    filter1, filter2, filter3 = st.columns(3)

    with filter1:
        if date_candidate is not None:
            use_date_filter = st.checkbox(
                f"Filter by {date_candidate}",
                value=False
            )

            if use_date_filter:
                parsed_date = pd.to_datetime(
                    filtered_df[date_candidate],
                    errors="coerce"
                )
                valid_dates = parsed_date.dropna()

                if not valid_dates.empty:
                    min_date = valid_dates.min().date()
                    max_date = valid_dates.max().date()

                    date_range = st.date_input(
                        "Select date range",
                        value=(min_date, max_date),
                        min_value=min_date,
                        max_value=max_date
                    )

                    if (
                        isinstance(date_range, (tuple, list))
                        and len(date_range) == 2
                    ):
                        start_date, end_date = date_range
                        filtered_df = filtered_df.loc[
                            parsed_date.notna()
                            & (parsed_date.dt.date >= start_date)
                            & (parsed_date.dt.date <= end_date)
                        ]

    with filter2:
        if category_candidates:
            category_column = st.selectbox(
                "Primary Filter",
                ["Not selected"] + category_candidates,
                key="dashboard_primary_filter"
            )

            if category_column != "Not selected":
                category_values = sorted(
                    filtered_df[category_column]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
                selected_categories = st.multiselect(
                    f"Select {category_column}",
                    category_values,
                    default=category_values,
                    key="dashboard_primary_values"
                )
                filtered_df = filtered_df[
                    filtered_df[category_column].astype(str).isin(
                        selected_categories
                    )
                ]

    with filter3:
        secondary_options = [
            col for col in category_candidates
            if col != category_column
        ] if category_candidates else []

        if secondary_options:
            secondary_column = st.selectbox(
                "Secondary Filter",
                ["Not selected"] + secondary_options,
                key="dashboard_secondary_filter"
            )

            if secondary_column != "Not selected":
                secondary_values = sorted(
                    filtered_df[secondary_column]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
                selected_secondary = st.multiselect(
                    f"Select {secondary_column}",
                    secondary_values,
                    default=secondary_values,
                    key="dashboard_secondary_values"
                )
                filtered_df = filtered_df[
                    filtered_df[secondary_column].astype(str).isin(
                        selected_secondary
                    )
                ]

    # -------------------------------------------------
    # KPI CARDS
    # -------------------------------------------------

    st.subheader("Key Performance Indicators")

    kpis = st.columns(5)

    kpis[0].metric("Filtered Records", f"{len(filtered_df):,}")

    if sales_column != "Not selected":
        sales_values = pd.to_numeric(filtered_df[sales_column], errors="coerce")
        kpis[1].metric("Total Sales", format_number(sales_values.sum()))
        kpis[2].metric("Average Sales", format_number(sales_values.mean()))
    else:
        kpis[1].metric("Total Sales", "Select a column")
        kpis[2].metric("Average Sales", "Select a column")

    if profit_column != "Not selected":
        profit_values = pd.to_numeric(filtered_df[profit_column], errors="coerce")
        kpis[3].metric("Total Profit", format_number(profit_values.sum()))
    else:
        kpis[3].metric("Total Profit", "Select a column")

    if quantity_column != "Not selected":
        quantity_values = pd.to_numeric(
            filtered_df[quantity_column], errors="coerce"
        )
        kpis[4].metric("Total Quantity", format_number(quantity_values.sum()))
    else:
        kpis[4].metric("Total Quantity", "Select a column")

    if order_column != "Not selected":
        order_count = filtered_df[order_column].nunique(dropna=True)
        st.metric("Unique Orders / IDs", f"{order_count:,}")
    else:
        st.caption(
            "Select an Order ID column to calculate unique orders. "
            "Otherwise, use Filtered Records as the record count."
        )

    # -------------------------------------------------
    # PHASE 3: ADVANCED CHARTS
    # -------------------------------------------------

    st.header("Step 6 — Advanced Data Analysis")

    if sales_column != "Not selected" and category_candidates:

        chart_category = st.selectbox(
            "Choose a category for sales analysis",
            category_candidates
        )

        category_summary = (
            filtered_df.groupby(
                chart_category,
                dropna=False
            )[sales_column]
            .sum()
            .reset_index()
            .sort_values(sales_column, ascending=False)
        )

        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            fig_bar = px.bar(
                category_summary,
                x=chart_category,
                y=sales_column,
                title=f"Sales by {chart_category}"
            )

            st.plotly_chart(fig_bar, use_container_width=True)

        with chart_col2:
            fig_pie = px.pie(
                category_summary,
                names=chart_category,
                values=sales_column,
                title=f"Sales Share by {chart_category}"
            )

            st.plotly_chart(fig_pie, use_container_width=True)

    if sales_column != "Not selected" and date_candidate is not None:

        date_series = pd.to_datetime(
            filtered_df[date_candidate],
            errors="coerce"
        )

        line_df = filtered_df.copy()
        line_df["_AnalysisDate"] = date_series
        line_df = line_df.dropna(subset=["_AnalysisDate"])

        if not line_df.empty:
            line_df["_Period"] = (
                line_df["_AnalysisDate"].dt.to_period("M").astype(str)
            )

            trend_df = (
                line_df.groupby("_Period")[sales_column]
                .sum()
                .reset_index()
            )

            fig_line = px.line(
                trend_df,
                x="_Period",
                y=sales_column,
                markers=True,
                title="Sales Trend by Month"
            )

            st.plotly_chart(fig_line, use_container_width=True)

    if numeric_columns:
        st.subheader("Numeric Data Analysis")

        selected_numeric = st.selectbox(
            "Select a numeric column for distribution analysis",
            numeric_columns
        )

        fig_hist = px.histogram(
            filtered_df,
            x=selected_numeric,
            title=f"Distribution of {selected_numeric}"
        )

        st.plotly_chart(fig_hist, use_container_width=True)

        if len(numeric_columns) >= 2:
            with st.expander("Numeric Correlation Analysis"):
                correlation = filtered_df[
                    numeric_columns
                ].corr(numeric_only=True)

                fig_corr = px.imshow(
                    correlation,
                    text_auto=".2f",
                    aspect="auto",
                    title="Numeric Column Correlation"
                )

                st.plotly_chart(fig_corr, use_container_width=True)

    # -------------------------------------------------
    # PHASE 4: AUTOMATED BUSINESS INSIGHTS
    # -------------------------------------------------

    st.header("Step 7 — Automated Business Insights")

    st.caption(
        "These are rule-based observations from your data, "
        "not AI-generated advice from a language model."
    )

    if filtered_df.empty:
        st.warning(
            "No records match the current filters. "
            "Adjust the filters to generate insights."
        )
    else:
        st.write(f"- The current filtered dataset contains {len(filtered_df):,} record(s).")

        missing_count = int(filtered_df.isna().sum().sum())
        duplicate_count = int(filtered_df.duplicated().sum())

        if missing_count:
            st.write(
                f"- There are {missing_count:,} missing cell(s) "
                "in the filtered data. Review the Data Quality Report."
            )
        else:
            st.write("- No missing cells were found in the filtered data.")

        if duplicate_count:
            st.write(
                f"- There are {duplicate_count:,} duplicate row(s) "
                "in the filtered data."
            )
        else:
            st.write("- No duplicate rows were found in the filtered data.")

        if sales_column != "Not selected":
            sales_values = filtered_df[sales_column].dropna()

            if not sales_values.empty:
                if (sales_values < 0).any():
                    st.write(
                        "- Negative sales values were found. "
                        "Check whether these represent refunds or data errors."
                    )

                if len(sales_values) >= 2:
                    st.write(
                        "- Review the Sales Distribution chart to identify "
                        "unusually high or low values."
                    )

        if profit_column != "Not selected":
            profit_values = filtered_df[profit_column].dropna()

            if not profit_values.empty:
                total_profit = profit_values.sum()

                if total_profit < 0:
                    st.warning(
                        "Total profit is negative for the selected data. "
                        "Check costs, refunds, and the selected column."
                    )
                elif total_profit > 0:
                    st.write(
                        "- Total profit is positive for the selected data. "
                        "Compare profit across categories and periods."
                    )

    # -------------------------------------------------
    # PHASE 5: DATA EXPLORER AND EXPORT
    # -------------------------------------------------

    
# ============================================================
# STEP 8 — DATA EXPLORER AND COMPLETE BUSINESS REPORT EXPORT
# ============================================================

st.header("Step 8 — Data Explorer and Export")
# Make sure filtered_df exists before Step 8
if "filtered_df" not in locals():
    filtered_df = st.session_state.cleaned_df.copy()

search_text = st.text_input(
    "Search text across the filtered dataset"
)

explorer_df = filtered_df.copy()

if search_text.strip():
    search_mask = explorer_df.astype(str).apply(
        lambda column: column.str.contains(
            search_text,
            case=False,
            regex=False,
            na=False
        )
    ).any(axis=1)

    explorer_df = explorer_df[search_mask]

st.write(f"Records displayed: {len(explorer_df):,}")

st.dataframe(
    explorer_df,
    use_container_width=True
)

# ------------------------------------------------------------
# 1. CSV DOWNLOAD
# ------------------------------------------------------------

csv_export = explorer_df.to_csv(index=False).encode("utf-8-sig")

st.download_button(
    "Download Filtered CSV",
    data=csv_export,
    file_name="business_analysis.csv",
    mime="text/csv",
    key="step8_csv_download"
)

# ------------------------------------------------------------
# 2. COMPLETE EXCEL WORKBOOK
# ------------------------------------------------------------

st.subheader("Download Complete Business Report")

st.write(
    "Export your data, dashboard charts, business summary, "
    "data quality report, and business insights into one Excel file."
)

if not explorer_df.empty:

    excel_buffer = BytesIO()

    # Find columns using common business column names
    def find_export_column(dataframe, keywords):
        for column in dataframe.columns:
            name = str(column).lower().strip()
            if any(keyword in name for keyword in keywords):
                return column
        return None

    sales_col = find_export_column(
        explorer_df, ["sales", "revenue", "amount", "total"]
    )
    profit_col = find_export_column(
        explorer_df, ["profit", "earnings", "net income"]
    )
    order_col = find_export_column(
        explorer_df, ["order id", "order number", "invoice id"]
    )
    date_col = find_export_column(
        explorer_df, ["order date", "transaction date", "date"]
    )
    category_col = find_export_column(
        explorer_df, ["category", "product", "item"]
    )
    region_col = find_export_column(
        explorer_df, ["region", "location", "city", "state"]
    )

    sales = (
        pd.to_numeric(explorer_df[sales_col], errors="coerce")
        if sales_col is not None
        else pd.Series(dtype="float64")
    )

    profit = (
        pd.to_numeric(explorer_df[profit_col], errors="coerce")
        if profit_col is not None
        else pd.Series(dtype="float64")
    )

    total_sales = float(sales.sum()) if sales_col is not None else 0
    total_profit = float(profit.sum()) if profit_col is not None else 0
    average_sales = float(sales.mean()) if sales_col is not None else 0

    total_orders = (
        int(explorer_df[order_col].nunique())
        if order_col is not None
        else len(explorer_df)
    )

    missing_cells = int(explorer_df.isna().sum().sum())
    duplicate_rows = int(explorer_df.duplicated().sum())

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:

        # SHEET 1: DATA
        explorer_df.to_excel(
            writer,
            index=False,
            sheet_name="Data"
        )

        # SHEET 2: SUMMARY
        summary_df = pd.DataFrame([
            ["BUSINESS PERFORMANCE SUMMARY", "Value"],
            ["Total Sales", total_sales],
            ["Total Profit", total_profit],
            ["Total Orders", total_orders],
            ["Average Sales", average_sales],
            ["Total Rows", len(explorer_df)],
            ["Total Columns", len(explorer_df.columns)],
            ["Missing Cells", missing_cells],
            ["Duplicate Rows", duplicate_rows],
            ["Source File", source_label],
        ])

        summary_df.to_excel(
            writer,
            index=False,
            header=False,
            sheet_name="Summary"
        )

        # SHEET 3: DASHBOARD
        dashboard_df = pd.DataFrame([
            ["BUSINESS INTELLIGENCE DASHBOARD", ""],
            ["Key Performance Indicator", "Value"],
            ["Total Sales", total_sales],
            ["Total Profit", total_profit],
            ["Total Orders", total_orders],
            ["Average Sales", average_sales],
        ])

        dashboard_df.to_excel(
            writer,
            index=False,
            header=False,
            sheet_name="Dashboard"
        )

        # SHEET 4: BUSINESS REPORT
        report_rows = [
            ["BUSINESS ANALYSIS REPORT", "Details"],
            ["Source File", source_label],
            ["Records in Report", len(explorer_df)],
            ["Total Sales", total_sales if sales_col is not None else "Sales column not found"],
            ["Total Profit", total_profit if profit_col is not None else "Profit column not found"],
            ["Total Orders", total_orders],
            ["Average Sales", average_sales if sales_col is not None else "Not available"],
            ["Missing Cells", missing_cells],
            ["Duplicate Rows", duplicate_rows],
        ]

        if sales_col is not None and profit_col is not None:
            if total_sales != 0:
                profit_margin = total_profit / total_sales * 100
                report_rows.append([
                    "Profit Margin",
                    f"{profit_margin:.2f}%"
                ])
            else:
                report_rows.append([
                    "Profit Margin",
                    "Not calculated because total sales is zero"
                ])

            if total_profit > 0:
                insight = "Overall profit is positive."
            elif total_profit < 0:
                insight = "Overall profit is negative. Review costs, returns, and sales."
            else:
                insight = "Overall profit is zero."

            report_rows.append(["Performance Insight", insight])

        if category_col is not None and sales_col is not None:
            category_sales = explorer_df.copy()
            category_sales["_sales_value"] = pd.to_numeric(
                category_sales[sales_col], errors="coerce"
            )

            category_totals = category_sales.groupby(
                category_col, dropna=False
            )["_sales_value"].sum()

            if not category_totals.empty:
                report_rows.append([
                    "Top Category by Sales",
                    str(category_totals.idxmax())
                ])

        if region_col is not None and sales_col is not None:
            region_sales = explorer_df.copy()
            region_sales["_sales_value"] = pd.to_numeric(
                region_sales[sales_col], errors="coerce"
            )

            region_totals = region_sales.groupby(
                region_col, dropna=False
            )["_sales_value"].sum()

            if not region_totals.empty:
                report_rows.append([
                    "Top Region by Sales",
                    str(region_totals.idxmax())
                ])

        report_rows.append([
            "Recommendation",
            "Review category performance, regional results, and monthly trends "
            "before making business decisions."
        ])

        pd.DataFrame(
            report_rows
        ).to_excel(
            writer,
            index=False,
            header=False,
            sheet_name="Report"
        )

        # SHEET 5: DATA QUALITY
        quality_df = quality_report(explorer_df)

        quality_df.to_excel(
            writer,
            index=False,
            sheet_name="Data Quality"
        )

        # ----------------------------------------------------
        # ADD EXCEL CHARTS TO THE DASHBOARD
        # ----------------------------------------------------

        from openpyxl.chart import BarChart, PieChart, LineChart, Reference
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        workbook = writer.book
        dashboard_ws = workbook["Dashboard"]

        # Category sales table + bar and pie charts
        next_row = 9

        if category_col is not None and sales_col is not None:
            chart_df = explorer_df.copy()

            chart_df["_sales_value"] = pd.to_numeric(
                chart_df[sales_col], errors="coerce"
            )

            category_table = (
                chart_df.groupby(category_col, dropna=False)["_sales_value"]
                .sum()
                .sort_values(ascending=False)
                .head(15)
                .reset_index()
            )

            category_table.columns = ["Category", "Sales"]

            dashboard_ws.cell(next_row, 1, "Sales by Category")
            dashboard_ws.cell(next_row + 1, 1, "Category")
            dashboard_ws.cell(next_row + 1, 2, "Sales")

            for row_number, values in enumerate(
                category_table.itertuples(index=False, name=None),
                start=next_row + 2
            ):
                dashboard_ws.cell(row_number, 1, str(values[0]))
                dashboard_ws.cell(row_number, 2, float(values[1]))

            if not category_table.empty:
                data_ref = Reference(
                    dashboard_ws,
                    min_col=2,
                    min_row=next_row + 1,
                    max_row=next_row + 1 + len(category_table)
                )
                category_ref = Reference(
                    dashboard_ws,
                    min_col=1,
                    min_row=next_row + 2,
                    max_row=next_row + 1 + len(category_table)
                )

                bar_chart = BarChart()
                bar_chart.type = "bar"
                bar_chart.style = 10
                bar_chart.title = "Sales by Category"
                bar_chart.x_axis.title = "Sales"
                bar_chart.y_axis.title = "Category"
                bar_chart.add_data(data_ref, titles_from_data=True)
                bar_chart.set_categories(category_ref)
                bar_chart.height = 8
                bar_chart.width = 15
                dashboard_ws.add_chart(bar_chart, "D2")

                pie_chart = PieChart()
                pie_chart.title = "Sales Share by Category"
                pie_chart.add_data(data_ref, titles_from_data=True)
                pie_chart.set_categories(category_ref)
                pie_chart.height = 8
                pie_chart.width = 12
                dashboard_ws.add_chart(pie_chart, "D18")

            next_row += len(category_table) + 4

        # Sales by region chart
        if region_col is not None and sales_col is not None:
            region_df = explorer_df.copy()

            region_df["_sales_value"] = pd.to_numeric(
                region_df[sales_col], errors="coerce"
            )

            region_table = (
                region_df.groupby(region_col, dropna=False)["_sales_value"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )

            region_table.columns = ["Region", "Sales"]

            dashboard_ws.cell(next_row, 1, "Sales by Region")
            dashboard_ws.cell(next_row + 1, 1, "Region")
            dashboard_ws.cell(next_row + 1, 2, "Sales")

            for row_number, values in enumerate(
                region_table.itertuples(index=False, name=None),
                start=next_row + 2
            ):
                dashboard_ws.cell(row_number, 1, str(values[0]))
                dashboard_ws.cell(row_number, 2, float(values[1]))

            if not region_table.empty:
                region_chart = BarChart()
                region_chart.type = "col"
                region_chart.style = 11
                region_chart.title = "Sales by Region"
                region_chart.y_axis.title = "Sales"
                region_chart.x_axis.title = "Region"

                region_data = Reference(
                    dashboard_ws,
                    min_col=2,
                    min_row=next_row + 1,
                    max_row=next_row + 1 + len(region_table)
                )
                region_categories = Reference(
                    dashboard_ws,
                    min_col=1,
                    min_row=next_row + 2,
                    max_row=next_row + 1 + len(region_table)
                )

                region_chart.add_data(
                    region_data, titles_from_data=True
                )
                region_chart.set_categories(region_categories)
                region_chart.height = 8
                region_chart.width = 15
                dashboard_ws.add_chart(region_chart, "D34")

            next_row += len(region_table) + 4

        # Monthly sales trend
        if date_col is not None and sales_col is not None:
            trend_df = explorer_df.copy()

            trend_df["_date_value"] = pd.to_datetime(
                trend_df[date_col], errors="coerce"
            )
            trend_df["_sales_value"] = pd.to_numeric(
                trend_df[sales_col], errors="coerce"
            )

            trend_df = trend_df.dropna(
                subset=["_date_value", "_sales_value"]
            )

            if not trend_df.empty:
                trend_df["_month"] = (
                    trend_df["_date_value"].dt.to_period("M").astype(str)
                )

                monthly_table = (
                    trend_df.groupby("_month")["_sales_value"]
                    .sum()
                    .reset_index()
                )

                monthly_table.columns = ["Month", "Sales"]

                dashboard_ws.cell(next_row, 1, "Monthly Sales Trend")
                dashboard_ws.cell(next_row + 1, 1, "Month")
                dashboard_ws.cell(next_row + 1, 2, "Sales")

                for row_number, values in enumerate(
                    monthly_table.itertuples(index=False, name=None),
                    start=next_row + 2
                ):
                    dashboard_ws.cell(row_number, 1, str(values[0]))
                    dashboard_ws.cell(row_number, 2, float(values[1]))

                monthly_chart = LineChart()
                monthly_chart.style = 13
                monthly_chart.title = "Monthly Sales Trend"
                monthly_chart.y_axis.title = "Sales"
                monthly_chart.x_axis.title = "Month"

                monthly_data = Reference(
                    dashboard_ws,
                    min_col=2,
                    min_row=next_row + 1,
                    max_row=next_row + 1 + len(monthly_table)
                )
                monthly_categories = Reference(
                    dashboard_ws,
                    min_col=1,
                    min_row=next_row + 2,
                    max_row=next_row + 1 + len(monthly_table)
                )

                monthly_chart.add_data(
                    monthly_data, titles_from_data=True
                )
                monthly_chart.set_categories(monthly_categories)
                monthly_chart.height = 8
                monthly_chart.width = 15
                dashboard_ws.add_chart(monthly_chart, "D50")

        # ----------------------------------------------------
        # PROFESSIONAL EXCEL FORMATTING
        # ----------------------------------------------------

        dark_blue = "17365D"
        medium_blue = "1F4E78"
        light_blue = "D9EAF7"

        header_fill = PatternFill(
            fill_type="solid",
            fgColor=dark_blue
        )
        title_fill = PatternFill(
            fill_type="solid",
            fgColor=medium_blue
        )
        header_font = Font(
            bold=True,
            color="FFFFFF",
            size=11
        )
        thin_border = Border(
            bottom=Side(style="thin", color="D9E2F3")
        )

        for ws in workbook.worksheets:
            ws.sheet_view.showGridLines = False

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(
                        vertical="center"
                    )

            # Style the first row
            for cell in ws[1]:
                cell.fill = title_fill
                cell.font = Font(
                    bold=True,
                    color="FFFFFF",
                    size=14
                )

            ws.row_dimensions[1].height = 27

            # Format Data headers
            if ws.title == "Data" and ws.max_row >= 1:
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

                ws.freeze_panes = "A2"
                if ws.max_row > 1 and ws.max_column > 0:
                    ws.auto_filter.ref = ws.dimensions

            # Style table headers
            if ws.title in ["Summary", "Report", "Data Quality"]:
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font

            # Auto-fit column widths with a reasonable maximum
            for column_cells in ws.columns:
                column_letter = get_column_letter(
                    column_cells[0].column
                )

                # Charts may extend the dashboard's used area
                if ws.title == "Dashboard" and column_cells[0].column > 8:
                    continue

                max_length = 0
                for cell in column_cells:
                    if cell.value is not None:
                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )

                ws.column_dimensions[column_letter].width = min(
                    max(max_length + 2, 12),
                    35
                )

        # Dashboard title
        dashboard_ws.merge_cells("A1:B1")
        dashboard_ws["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        dashboard_ws["A1"].font = Font(
            bold=True,
            color="FFFFFF",
            size=16
        )

        # Highlight KPI labels
        for row_number in range(2, 7):
            dashboard_ws.cell(row_number, 1).font = Font(
                bold=True,
                color=medium_blue,
                size=11
            )
            dashboard_ws.cell(row_number, 1).fill = PatternFill(
                fill_type="solid",
                fgColor=light_blue
            )

        # Currency-style display for sales/profit/average
        for sheet_name in ["Summary", "Dashboard"]:
            ws = workbook[sheet_name]

            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        if cell.row in [2, 3, 4, 5, 6]:
                            cell.number_format = '#,##0.00;[Red]-#,##0.00'

    excel_buffer.seek(0)

    st.download_button(
        "📥 Download Complete Business Report (Excel)",
        data=excel_buffer.getvalue(),
        file_name="Complete_Business_Intelligence_Report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        key="step8_complete_excel_download",
        use_container_width=True
    )

else:
    st.warning(
        "No records are available to export. "
        "Please adjust your filters or search text."
    )

# ------------------------------------------------------------
# 3. SUMMARY TEXT DOWNLOAD
# ------------------------------------------------------------

report_text = (
    "AI Business Intelligence Platform\n"
    f"Source file: {source_label}\n"
    f"Original rows: {len(source_df)}\n"
    f"Cleaned rows: {len(st.session_state.cleaned_df)}\n"
    f"Filtered rows: {len(filtered_df)}\n"
    f"Records displayed after search: {len(explorer_df)}\n"
    f"Remaining missing cells: "
    f"{int(st.session_state.cleaned_df.isna().sum().sum())}\n"
    f"Remaining duplicate rows: "
    f"{int(st.session_state.cleaned_df.duplicated().sum())}\n"
    f"Correction actions applied: "
    f"{len(st.session_state.correction_log)}\n"
)

st.download_button(
    "Download Summary Report (TXT)",
    data=report_text.encode("utf-8"),
    file_name="business_summary.txt",
    mime="text/plain",
    key="step8_text_download"
)

st.caption(
    "Always verify column meanings and business rules before "
    "using analysis results for financial decisions."
)

