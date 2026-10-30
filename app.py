
import streamlit as st
import pandas as pd
from io import BytesIO
import re

# -------------------------------
# PAGE CONFIGURATION
# -------------------------------

st.set_page_config(
    page_title="Excel Business Cleaner",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Business Cleaner")
st.write("Clean, organise and analyse your business spreadsheets.")

# -------------------------------
# FILE UPLOAD
# -------------------------------

uploaded_file = st.file_uploader(
    "Upload your Excel or CSV file",
    type=["xlsx", "csv"]
)

# -------------------------------
# DATA CLEANING FUNCTION
# -------------------------------

def clean_column_name(column):
    column = str(column).strip().lower()
    column = re.sub(r"\s+", "_", column)
    column = re.sub(r"[^a-z0-9_]", "", column)
    return column


def clean_data(df, remove_duplicates):

    df = df.copy()

    # Clean column names
    df.columns = [
        clean_column_name(col)
        for col in df.columns
    ]

    # Remove unnecessary spaces from text
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].map(
            lambda value: value.strip()
            if isinstance(value, str) else value
        )

    # Count and remove duplicate rows
    duplicate_count = int(df.duplicated().sum())

    if remove_duplicates:
        df = df.drop_duplicates()

    return df, duplicate_count


# -------------------------------
# PROCESS UPLOADED FILE
# -------------------------------

if uploaded_file is not None:

    try:

        if uploaded_file.name.lower().endswith(".csv"):
            original_df = pd.read_csv(uploaded_file)
        else:
            original_df = pd.read_excel(
                uploaded_file,
                engine="openpyxl"
            )

        if original_df.empty:
            st.warning("Your uploaded file contains no data.")
            st.stop()

        if len(original_df.columns) == 0:
            st.warning("No columns found in the uploaded file.")
            st.stop()

        st.subheader("Original Data Preview")

        st.dataframe(
            original_df.head(20),
            use_container_width=True
        )

        # Original data statistics
        total_rows = len(original_df)
        total_columns = len(original_df.columns)
        missing_before = int(original_df.isna().sum().sum())

        st.subheader("Original File Statistics")

        col1, col2, col3 = st.columns(3)

        col1.metric("Total Rows", total_rows)
        col2.metric("Total Columns", total_columns)
        col3.metric("Missing Values", missing_before)

        # Cleaning settings
        st.subheader("Cleaning Settings")

        remove_duplicates = st.checkbox(
            "Remove duplicate rows",
            value=True
        )

        # Process data
        if st.button("Clean My Data", type="primary"):

            cleaned_df, duplicate_count = clean_data(
                original_df,
                remove_duplicates
            )

            st.session_state["cleaned_df"] = cleaned_df
            st.session_state["original_name"] = uploaded_file.name
            st.session_state["duplicate_count"] = duplicate_count
            st.session_state["remove_duplicates"] = remove_duplicates
            st.session_state["original_rows"] = total_rows
            st.session_state["missing_before"] = missing_before

        # Display results
        if "cleaned_df" in st.session_state:

            cleaned_df = st.session_state["cleaned_df"]

            st.success("Your data has been processed!")

            st.subheader("Cleaning Results")

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Original Rows",
                st.session_state["original_rows"]
            )

            col2.metric(
                "Cleaned Rows",
                len(cleaned_df)
            )

            col3.metric(
                "Duplicates Found",
                st.session_state["duplicate_count"]
            )

            # Missing value report
            missing_report = pd.DataFrame({
                "Column": cleaned_df.columns,
                "Missing Values": cleaned_df.isna().sum().values
            })

            st.subheader("Data Quality Report")

            st.dataframe(
                missing_report,
                use_container_width=True
            )

            st.subheader("Cleaned Data Preview")

            st.dataframe(
                cleaned_df.head(20),
                use_container_width=True
            )

            # Export Excel workbook
            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                cleaned_df.to_excel(
                    writer,
                    sheet_name="Cleaned Data",
                    index=False
                )

                missing_report.to_excel(
                    writer,
                    sheet_name="Quality Report",
                    index=False
                )

            output.seek(0)

            output_name = (
                st.session_state["original_name"]
                .rsplit(".", 1)[0] + "_cleaned.xlsx"
            )

            st.download_button(
                label="⬇ Download Cleaned Excel File",
                data=output.getvalue(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as error:

        st.error("Unable to process this file.")
        st.exception(error)

else:

    st.info("Upload an Excel or CSV file to get started.")
