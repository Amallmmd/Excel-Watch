import re

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Excel Watch", layout="centered")


HEADER_ALIASES = {
    "srno": "Sr.No",
    "serialno": "Sr.No",
    "invoicenumber": "Invoice Number",
    "invoiceno": "Invoice Number",
    "invoicedate": "Invoice date",
    "typeofinvoice": "Type of Invoice",
    "subscription": "Subscription",
    "total": "Total",
}

REQUIRED_COLUMNS = [
    "Sr.No",
    "Invoice Number",
    "Invoice date",
    "Type of Invoice",
    "Subscription",
    "Total",
]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_header(value):
    return re.sub(r"[^a-z0-9]", "", clean_text(value).lower())


def find_labeled_value(row_values, pattern):
    for value in row_values:
        match = re.match(pattern, clean_text(value), flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def find_header(row_values):
    header = {}

    for index, value in enumerate(row_values):
        column_name = HEADER_ALIASES.get(normalize_header(value))
        if column_name:
            header[column_name] = index

    if all(column in header for column in REQUIRED_COLUMNS):
        return header

    return None


def value_at(row_values, index):
    if index >= len(row_values):
        return None
    return row_values[index]


def is_section_total_row(row_values, header):
    has_total_label = any(clean_text(value).lower() == "total" for value in row_values)
    invoice_number = clean_text(value_at(row_values, header["Invoice Number"]))
    invoice_date = clean_text(value_at(row_values, header["Invoice date"]))

    return has_total_label and not invoice_number and not invoice_date


def parse_invoice_sheet(raw_df):
    records = []
    current_customer = None
    current_sub_customer = None
    active_header = None

    for _, row in raw_df.iterrows():
        row_values = list(row)

        sub_customer = find_labeled_value(
            row_values, r"^\s*sub\s*[- ]?\s*customer\s*:?\s*(.+)$"
        )
        if sub_customer:
            current_sub_customer = sub_customer
            active_header = None
            continue

        customer = find_labeled_value(row_values, r"^\s*customer\s*:?\s*(.+)$")
        if customer:
            current_customer = customer
            current_sub_customer = None
            active_header = None
            continue

        header = find_header(row_values)
        if header:
            active_header = header
            continue

        if not active_header or not current_customer or not current_sub_customer:
            continue

        if not any(clean_text(value) for value in row_values):
            continue

        if is_section_total_row(row_values, active_header):
            active_header = None
            continue

        record = {
            "Customer": current_customer,
            "Sub-Customer": current_sub_customer,
        }

        for column in REQUIRED_COLUMNS:
            record[column] = value_at(row_values, active_header[column])

        records.append(record)

    invoices = pd.DataFrame(records)
    if invoices.empty:
        return invoices

    for column in ["Invoice Number", "Type of Invoice", "Subscription"]:
        invoices[column] = invoices[column].map(clean_text)

    invoices["Invoice date"] = pd.to_datetime(invoices["Invoice date"], errors="coerce")
    invoices["Total"] = pd.to_numeric(invoices["Total"], errors="coerce")
    invoices = invoices.dropna(subset=["Invoice date", "Total"]).copy()

    if invoices.empty:
        return invoices

    invoices["Year"] = invoices["Invoice date"].dt.year
    invoices["Month Number"] = invoices["Invoice date"].dt.month
    invoices["Month"] = invoices["Invoice date"].dt.month_name()
    invoices["Year-Month"] = invoices["Invoice date"].dt.to_period("M").astype(str)

    return invoices[
        [
            "Customer",
            "Sub-Customer",
            "Sr.No",
            "Invoice Number",
            "Invoice date",
            "Type of Invoice",
            "Subscription",
            "Total",
            "Year",
            "Month Number",
            "Month",
            "Year-Month",
        ]
    ]


def month_options(df):
    months = (
        df[["Month Number", "Month"]]
        .drop_duplicates()
        .sort_values("Month Number")
        .reset_index(drop=True)
    )
    return months["Month"].tolist()


st.title("Excel Watch")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is None:
    st.info("Upload an Excel file to start.")
    st.stop()

try:
    raw_data = pd.read_excel(uploaded_file, header=None, engine="openpyxl")
except Exception as error:
    st.error(f"Could not read this Excel file: {error}")
    st.stop()

invoices = parse_invoice_sheet(raw_data)

if invoices.empty:
    st.warning("No valid invoice table was found in this Excel file.")
    st.stop()

st.success(f"Loaded {len(invoices)} invoice rows.")

customer = st.selectbox("Customer", sorted(invoices["Customer"].unique()))

customer_rows = invoices[invoices["Customer"] == customer]
sub_customer = st.selectbox(
    "Sub-Customer", sorted(customer_rows["Sub-Customer"].unique())
)

sub_customer_rows = customer_rows[customer_rows["Sub-Customer"] == sub_customer]
year = st.selectbox("Year", sorted(sub_customer_rows["Year"].unique(), reverse=True))

year_rows = sub_customer_rows[sub_customer_rows["Year"] == year]
month = st.selectbox("Month", month_options(year_rows))

if st.button("Calculate", type="primary"):
    selected_rows = year_rows[year_rows["Month"] == month].copy()
    selected_total = selected_rows["Total"].sum()

    st.metric("Total", f"{selected_total:,.2f}")
    st.dataframe(
        selected_rows[
            [
                "Invoice Number",
                "Invoice date",
                "Type of Invoice",
                "Subscription",
                "Total",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Monthly Summary")
    summary = (
        invoices.groupby(["Customer", "Year-Month"], as_index=False)["Total"]
        .sum()
        .sort_values(["Customer", "Year-Month"])
    )
    st.dataframe(summary, use_container_width=True, hide_index=True)
