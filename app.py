import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import hashlib
import re

st.set_page_config(
    page_title="Excel Business Cleaner & Sales Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Business Cleaner")
st.caption("Clean spreadsheets, check data quality, and analyse business sales.")
st.divider()


# ==========================================
# HELPER FUNCTIONS
# ==========================================

def clean_column_names(columns):
    """Standardise column names and make them unique."""
    cleaned = []
    used = set()

    for index, column in enumerate(columns, start=1):
        name = str(column).strip().lower()
        name = re.sub(r"\s+", "_", name)
        name = re.sub(r"[^a-z0-9_]", "", name)
        name = name.strip("_")

        if not name:
            name = f"unnamed_column_{index}"

        original = name
        counter = 2
        while name in used:
            name = f"{original}_{counter}"
            counter += 1

        used.add(name)
        cleaned.append(name)

    return cleaned


def clean_data(df, remove_duplicates=True, trim_text=True):
    df = df.copy()
    original_columns = list(df.columns)
    new_columns = clean_column_names(original_columns)

    column_mapping = pd.DataFrame({
        "Original Column": [str(c) for c in original_columns],
        "Cleaned Column": new_columns
    })

    df.columns = new_columns

    if trim_text:
        text_columns = df.select_dtypes(
            include=["object", "string"]
        ).columns
        for column in text_columns:
            df[column] = df[column].map(
                lambda value: value.strip() if isinstance(value, str) else value
            )

    duplicates = int(df.duplicated().sum())
    if remove_duplicates:
        df = df.drop_duplicates()

    return df, duplicates, column_mapping


def create_quality_report(original, cleaned, sheet_name):
    rows = []
    for index, column in enumerate(cleaned.columns):
        before = int(original.iloc[:, index].isna().sum())
        after = int(cleaned[column].isna().sum())

        rows.append({
            "Sheet": sheet_name,
            "Column": column,
            "Missing Before": before,
            "Missing After": after,
            "Data Type": str(cleaned[column].dtype)
        })

    return pd.DataFrame(rows)


def safe_sheet_name(name, used_names):
    name = re.sub(r"[\[\]:*?/\\]", "_", str(name))
    name = name[:31] or "Cleaned Data"
    original = name
    counter = 2

    while name.lower() in used_names:
        suffix = f"_{counter}"
        name = original[:31 - len(suffix)] + suffix
        counter += 1

    used_names.add(name.lower())
    return name


def read_csv_safely(file_bytes):
    try:
        return pd.read_csv(BytesIO(file_bytes), encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(BytesIO(file_bytes), encoding="cp1252")


def create_excel_export(results):
    """Create a multi-sheet Excel report in memory."""
    output = BytesIO()
    used_names = set()
    all_reports = []

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, data in results.items():
            cleaned_sheet = safe_sheet_name(
                f"Cleaned_{sheet_name}", used_names
            )
            mapping_sheet = safe_sheet_name(
                f"Columns_{sheet_name}", used_names
            )

            data["cleaned"].to_excel(
                writer, sheet_name=cleaned_sheet, index=False
            )
            data["mapping"].to_excel(
                writer, sheet_name=mapping_sheet, index=False
            )
            all_reports.append(data["report"])

        quality_report = pd.concat(all_reports, ignore_index=True)
        report_sheet = safe_sheet_name("Quality Report", used_names)
        quality_report.to_excel(
            writer, sheet_name=report_sheet, index=False
        )

    output.seek(0)
    return output.getvalue()


def guess_column(columns, keywords):
    """Suggest a likely column based on common header names."""
    for column in columns:
        if any(word in column.lower() for word in keywords):
            return column
    return columns[0] if len(columns) else None


def show_sales_dashboard(data, sheet_name):
    """Show optional sales analytics for a selected cleaned worksheet."""
    st.header("📈 Sales Performance Dashboard")
    st.caption(
        f"Analysis for worksheet: {sheet_name}. "
        "Choose the columns that match your spreadsheet."
    )

    df = data["cleaned"].copy()
    if df.empty:
        st.info("This worksheet has no rows to analyse.")
        return

    columns = list(df.columns)
    numeric_columns = [
        col for col in columns
        if pd.api.types.is_numeric_dtype(df[col])
    ]

    if not numeric_columns:
        st.info(
            "No numeric columns were detected. A sales dashboard needs "
            "a numeric amount/revenue column."
        )
        return

    revenue_default = guess_column(
        numeric_columns, ["revenue", "sales", "amount", "total", "price", "value"]
    )
    revenue_col = st.selectbox(
        "Which column contains sales / revenue?",
        numeric_columns,
        index=numeric_columns.index(revenue_default)
    )

    text_or_all_columns = columns
    date_default = guess_column(
        columns, ["date", "month", "day", "order_date", "invoice_date"]
    )
    product_default = guess_column(
        columns, ["product", "item", "service", "sku"]
    )
    customer_default = guess_column(
        columns, ["customer", "client", "buyer", "name"]
    )

    with st.expander("Choose optional analysis columns", expanded=True):
        date_options = ["Not available"] + text_or_all_columns
        product_options = ["Not available"] + text_or_all_columns
        customer_options = ["Not available"] + text_or_all_columns

        date_index = (
            date_options.index(date_default)
            if date_default in date_options else 0
        )
        product_index = (
            product_options.index(product_default)
            if product_default in product_options else 0
        )
        customer_index = (
            customer_options.index(customer_default)
            if customer_default in customer_options else 0
        )

        date_choice = st.selectbox(
            "Date column (optional)", date_options, index=date_index
        )
        product_choice = st.selectbox(
            "Product column (optional)", product_options, index=product_index
        )
        customer_choice = st.selectbox(
            "Customer column (optional)", customer_options,
            index=customer_index
        )

    revenue = pd.to_numeric(df[revenue_col], errors="coerce")
    valid_revenue = revenue.dropna()

    total_revenue = float(valid_revenue.sum()) if not valid_revenue.empty else 0
    average_sale = float(valid_revenue.mean()) if not valid_revenue.empty else 0
    transaction_count = int(valid_revenue.count())

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Total Revenue", f"{total_revenue:,.2f}")
    metric2.metric("Average Sale", f"{average_sale:,.2f}")
    metric3.metric("Transactions Analysed", f"{transaction_count:,}")

    if product_choice != "Not available":
        product_data = pd.DataFrame({
            "Product": df[product_choice].astype("string"),
            "Revenue": revenue
        }).dropna(subset=["Product", "Revenue"])

        if not product_data.empty:
            top_products = (
                product_data.groupby("Product", as_index=False)["Revenue"]
                .sum()
                .sort_values("Revenue", ascending=False)
                .head(10)
            )
            st.subheader("Top Products by Revenue")
            fig = px.bar(
                top_products.sort_values("Revenue"),
                x="Revenue", y="Product", orientation="h",
                title="Top 10 Products"
            )
            st.plotly_chart(fig, use_container_width=True)

    if customer_choice != "Not available":
        customer_data = pd.DataFrame({
            "Customer": df[customer_choice].astype("string"),
            "Revenue": revenue
        }).dropna(subset=["Customer", "Revenue"])

        if not customer_data.empty:
            top_customers = (
                customer_data.groupby("Customer", as_index=False)["Revenue"]
                .sum()
                .sort_values("Revenue", ascending=False)
                .head(10)
            )
            st.subheader("Top Customers by Revenue")
            fig = px.bar(
                top_customers.sort_values("Revenue"),
                x="Revenue", y="Customer", orientation="h",
                title="Top 10 Customers"
            )
            st.plotly_chart(fig, use_container_width=True)

    if date_choice != "Not available":
        date_data = pd.DataFrame({
            "Date": pd.to_datetime(df[date_choice], errors="coerce"),
            "Revenue": revenue
        }).dropna(subset=["Date", "Revenue"])

        if not date_data.empty:
            date_data["Month"] = date_data["Date"].dt.to_period("M").astype(str)
            monthly = (
                date_data.groupby("Month", as_index=False)["Revenue"]
                .sum()
                .sort_values("Month")
            )
            st.subheader("Revenue by Month")
            fig = px.line(
                monthly, x="Month", y="Revenue",
                markers=True, title="Monthly Revenue"
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(
                "A date column was selected, but no valid dates were found. "
                "Check the date format in your spreadsheet."
            )

    st.caption(
        "Dashboard figures are calculated from the selected cleaned worksheet. "
        "Revenue is the sum of the selected numeric column; verify that it "
        "represents sales amounts before using these figures for decisions."
    )


# ==========================================
# FILE UPLOAD
# ==========================================

uploaded_file = st.file_uploader(
    "Upload your Excel or CSV file",
    type=["xlsx", "csv"]
)

if uploaded_file is None:
    st.info("Upload a spreadsheet to start cleaning your data.")
    st.markdown("""
    ### Available features
    - Duplicate row detection and removal
    - Column name standardisation
    - Text whitespace cleaning
    - Missing-value analysis
    - Multiple Excel sheet processing
    - Downloadable Excel reports
    - Sales dashboard with interactive charts
    """)
    st.stop()


# ==========================================
# FILE VALIDATION
# ==========================================

file_bytes = uploaded_file.getvalue()

if len(file_bytes) == 0:
    st.error("The uploaded file is empty.")
    st.stop()

if len(file_bytes) > 25 * 1024 * 1024:
    st.error("Maximum supported file size is 25 MB.")
    st.stop()

file_hash = hashlib.sha256(file_bytes).hexdigest()

if st.session_state.get("file_hash") != file_hash:
    st.session_state["file_hash"] = file_hash
    st.session_state.pop("results", None)
    st.session_state.pop("run_key", None)


# ==========================================
# READ FILE
# ==========================================

try:
    if uploaded_file.name.lower().endswith(".csv"):
        original_data = read_csv_safely(file_bytes)
        if original_data.empty:
            st.warning("No data found in this CSV file.")
            st.stop()

        available_sheets = ["CSV Data"]
        sheet_data = {"CSV Data": original_data}
    else:
        excel_file = pd.ExcelFile(
            BytesIO(file_bytes), engine="openpyxl"
        )
        available_sheets = excel_file.sheet_names
        sheet_data = {
            name: excel_file.parse(name)
            for name in available_sheets
        }

        if not available_sheets:
            st.warning("No worksheets found.")
            st.stop()

except Exception as error:
    st.error("Unable to read this file.")
    st.exception(error)
    st.stop()


# ==========================================
# FILE INFORMATION
# ==========================================

st.subheader("File Information")
col1, col2 = st.columns(2)
col1.metric("File Name", uploaded_file.name)
col2.metric("Worksheets", len(available_sheets))


# ==========================================
# CLEANING SETTINGS
# ==========================================

st.subheader("Cleaning Settings")

process_choice = st.radio(
    "Choose processing mode",
    ["All Sheets", "Select One Sheet"],
    horizontal=True
)

selected_sheets = available_sheets

if process_choice == "Select One Sheet":
    selected_sheet = st.selectbox(
        "Choose a worksheet", available_sheets
    )
    selected_sheets = [selected_sheet]

remove_duplicates = st.checkbox("Remove duplicate rows", value=True)
trim_text = st.checkbox(
    "Remove unnecessary spaces from text", value=True
)

run_key = (
    f"{file_hash}:{selected_sheets}:"
    f"{remove_duplicates}:{trim_text}"
)

if st.session_state.get("run_key") != run_key:
    st.session_state.pop("results", None)


# ==========================================
# PROCESS DATA
# ==========================================

if st.button("Clean My Data", type="primary"):
    results = {}

    try:
        for sheet_name in selected_sheets:
            original_df = sheet_data[sheet_name]

            if original_df.empty:
                continue

            cleaned_df, duplicate_count, mapping = clean_data(
                original_df, remove_duplicates, trim_text
            )

            report = create_quality_report(
                original_df, cleaned_df, sheet_name
            )

            results[sheet_name] = {
                "cleaned": cleaned_df,
                "report": report,
                "mapping": mapping,
                "original_rows": len(original_df),
                "cleaned_rows": len(cleaned_df),
                "duplicates": duplicate_count,
                "missing_before": int(original_df.isna().sum().sum()),
                "missing_after": int(cleaned_df.isna().sum().sum())
            }

        if not results:
            st.warning("No non-empty worksheets to process.")
            st.stop()

        st.session_state["results"] = results
        st.session_state["run_key"] = run_key
        st.success("Data cleaning completed successfully!")

    except Exception as error:
        st.error("An error occurred during cleaning.")
        st.exception(error)


# ==========================================
# DISPLAY RESULTS
# ==========================================

if (
    st.session_state.get("run_key") == run_key
    and "results" in st.session_state
):
    results = st.session_state["results"]

    total_original = sum(x["original_rows"] for x in results.values())
    total_cleaned = sum(x["cleaned_rows"] for x in results.values())
    total_duplicates = sum(x["duplicates"] for x in results.values())
    total_missing_before = sum(
        x["missing_before"] for x in results.values()
    )
    total_missing_after = sum(
        x["missing_after"] for x in results.values()
    )

    st.divider()
    st.header("Cleaning Results")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Original Rows", total_original)
    col2.metric("Cleaned Rows", total_cleaned)
    col3.metric("Duplicates Found", total_duplicates)
    col4.metric("Missing Values", total_missing_after)

    st.subheader("Data Quality Comparison")
    comparison = pd.DataFrame({
        "Metric": ["Rows", "Missing Values"],
        "Before Cleaning": [total_original, total_missing_before],
        "After Cleaning": [total_cleaned, total_missing_after]
    })
    st.dataframe(comparison, use_container_width=True)

    st.subheader("Worksheet Results")
    for sheet_name, data in results.items():
        with st.expander(f"Worksheet: {sheet_name}", expanded=True):
            st.write(f"Original rows: {data['original_rows']}")
            st.write(f"Cleaned rows: {data['cleaned_rows']}")
            st.write(f"Duplicates found: {data['duplicates']}")
            st.write("Cleaned data preview")
            st.dataframe(data["cleaned"].head(20), use_container_width=True)
            st.write("Column mapping")
            st.dataframe(data["mapping"], use_container_width=True)

    # ======================================
    # SALES DASHBOARD
    # ======================================

    st.divider()
    st.subheader("Explore Your Sales Data")
    dashboard_sheet = st.selectbox(
        "Choose a worksheet for sales analysis",
        list(results.keys()),
        key="dashboard_sheet"
    )
    show_sales_dashboard(results[dashboard_sheet], dashboard_sheet)

    # ======================================
    # DOWNLOAD REPORT
    # ======================================

    st.divider()
    st.subheader("Download Your Results")

    try:
        excel_output = create_excel_export(results)
        output_name = (
            uploaded_file.name.rsplit(".", 1)[0] + "_cleaned.xlsx"
        )

        st.download_button(
            label="Download Complete Excel Report",
            data=excel_output,
            file_name=output_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary"
        )
    except Exception as error:
        st.error("Could not create the Excel export.")
        st.exception(error)

    st.caption(
        "Your original uploaded file is not modified. "
        "Missing values are reported, not automatically filled."
    )
