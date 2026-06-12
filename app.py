from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Excel Watch", layout="centered")

REQUIRED_COLUMNS = {
    "customer": "Customer",
    "invoicedate": "Invoice Date",
    "total": "Total",
}


def load_theme():
    css_path = Path(__file__).with_name("style.css")
    st.markdown(f"<style>{css_path.read_text()}</style>", unsafe_allow_html=True)


def normalize_column_name(name):
    return "".join(char for char in str(name).lower() if char.isalnum())


def find_required_columns(df):
    found_columns = {}

    for column in df.columns:
        normalized = normalize_column_name(column)
        if normalized in REQUIRED_COLUMNS:
            found_columns[REQUIRED_COLUMNS[normalized]] = column

    missing_columns = [
        display_name
        for display_name in REQUIRED_COLUMNS.values()
        if display_name not in found_columns
    ]

    return found_columns, missing_columns


def prepare_data(df, found_columns):
    prepared = df.rename(
        columns={
            found_columns["Customer"]: "Customer",
            found_columns["Invoice Date"]: "Invoice Date",
            found_columns["Total"]: "Total",
        }
    ).copy()

    prepared["Customer"] = prepared["Customer"].astype(str).str.strip()
    prepared["Invoice Date"] = pd.to_datetime(prepared["Invoice Date"], errors="coerce")
    prepared["Total"] = pd.to_numeric(prepared["Total"], errors="coerce")

    prepared = prepared.dropna(subset=["Customer", "Invoice Date", "Total"]).copy()
    prepared = prepared[prepared["Customer"] != ""].copy()

    prepared["Year"] = prepared["Invoice Date"].dt.year
    prepared["Month Number"] = prepared["Invoice Date"].dt.month
    prepared["Month"] = prepared["Invoice Date"].dt.month_name()
    prepared["Year-Month"] = prepared["Invoice Date"].dt.to_period("M").astype(str)

    return prepared


def month_options(df):
    months = (
        df[["Month Number", "Month"]]
        .drop_duplicates()
        .sort_values("Month Number")
        .reset_index(drop=True)
    )
    return months["Month"].tolist()


def build_summary(df):
    rows = []

    for customer, customer_df in df.groupby("Customer", sort=True):
        rows.append(
            {
                "Row Labels": customer,
                "Sum of Total": customer_df["Total"].sum(),
            }
        )

        monthly_totals = (
            customer_df.groupby("Year-Month", as_index=False)["Total"]
            .sum()
            .sort_values("Year-Month")
        )

        for _, month_row in monthly_totals.iterrows():
            rows.append(
                {
                    "Row Labels": month_row["Year-Month"],
                    "Sum of Total": month_row["Total"],
                }
            )

    return pd.DataFrame(rows)


def summary_to_excel(summary):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")

    return buffer.getvalue()


def read_input_file(uploaded_file):
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    return pd.read_excel(uploaded_file, engine="openpyxl")


def select_filter(label, options):
    options = list(options)
    use_all = st.checkbox(f"All {label.lower()}", disabled=not options)

    if use_all:
        return options

    return st.multiselect(label, options, disabled=not options)


load_theme()

st.title("Excel Watch")

uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file is None:
    st.info("Upload an Excel or CSV file to start.")
    st.stop()

try:
    input_data = read_input_file(uploaded_file)
except Exception as error:
    st.error(f"Could not read this file: {error}")
    st.stop()

found_columns, missing_columns = find_required_columns(input_data)

if missing_columns:
    st.warning(
        "Missing required columns: " + ", ".join(missing_columns)
    )
    st.stop()

data = prepare_data(input_data, found_columns)

if data.empty:
    st.warning("No valid rows found after reading Customer, Invoice Date, and Total.")
    st.stop()

st.success(f"Loaded {len(data)} valid rows.")

st.subheader("Input Preview")
st.dataframe(input_data, use_container_width=True, hide_index=True)

customers = sorted(data["Customer"].unique())
selected_customers = select_filter("Customers", customers)

if selected_customers:
    customer_rows = data[data["Customer"].isin(selected_customers)]
else:
    customer_rows = data.iloc[0:0]

years = sorted(customer_rows["Year"].unique(), reverse=True)
selected_years = select_filter("Years", years)

if selected_years:
    year_rows = customer_rows[customer_rows["Year"].isin(selected_years)]
else:
    year_rows = data.iloc[0:0]

months = month_options(year_rows) if not year_rows.empty else []
selected_months = select_filter("Months", months)

if st.button("Calculate", type="primary"):
    if not selected_customers or not selected_years or not selected_months:
        st.warning("Select at least one customer, year, and month.")
        st.stop()

    selected_rows = year_rows[year_rows["Month"].isin(selected_months)].copy()

    if selected_rows.empty:
        st.warning("No rows match the selected filters.")
        st.stop()

    summary = build_summary(selected_rows)
    st.session_state["summary"] = summary

if "summary" in st.session_state:
    st.subheader("Output")
    st.dataframe(st.session_state["summary"], use_container_width=True, hide_index=True)
    st.download_button(
        label="Download Output",
        data=summary_to_excel(st.session_state["summary"]),
        file_name="excel_watch_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
