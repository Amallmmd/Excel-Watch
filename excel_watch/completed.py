import calendar

import pandas as pd
import streamlit as st

from excel_watch.common import (
    clean_text,
    clean_text_series,
    dataframe_to_csv,
    find_columns,
    read_input_file,
    select_filter,
    unique_options,
)


EXPECTED_COLUMNS = {
    "Customer": ["Customer"],
    "Sub Customer": ["SubCustomer", "Sub Customer"],
    "Invoice #": ["Invoice #"],
    "Invoice Date": ["Invoice Date"],
    "Type": ["Type"],
    "Subscription": ["Subscription"],
    "Total": ["Total"],
    "Created": ["Created"],
    "Sent": ["Sent"],
}

OUTPUT_COLUMNS = [
    "Customer",
    "Sub Customer",
    "Invoice #",
    "Invoice Date",
    "Type",
    "Subscription",
    "Total",
]

NAME_SEPARATORS = ("\u2022", "\u00e2\u20ac\u00a2")


def extract_name(value):
    text = clean_text(value)

    for separator in NAME_SEPARATORS:
        if separator in text:
            return text.rsplit(separator, 1)[-1].strip()

    return text


def prepare_data(df, found_columns):
    prepared = pd.DataFrame(index=df.index)

    for column in EXPECTED_COLUMNS:
        if column in found_columns:
            prepared[column] = df[found_columns[column]]
        else:
            prepared[column] = ""

    for column in ["Customer", "Sub Customer", "Invoice #", "Type", "Subscription"]:
        prepared[column] = clean_text_series(prepared[column])

    prepared["Created By"] = prepared["Created"].map(extract_name)
    prepared["Sent By"] = prepared["Sent"].map(extract_name)
    prepared["_Invoice Date Parsed"] = pd.to_datetime(
        prepared["Invoice Date"],
        errors="coerce",
    )

    parsed_dates = prepared["_Invoice Date Parsed"]
    prepared["Year"] = parsed_dates.dt.year.astype("Int64")
    prepared["Month Number"] = parsed_dates.dt.month.astype("Int64")
    prepared["Month"] = parsed_dates.dt.month_name()
    prepared["Date"] = parsed_dates.dt.date

    prepared["Invoice Date"] = parsed_dates.dt.strftime("%Y-%m-%d").fillna(
        clean_text_series(prepared["Invoice Date"])
    )

    return prepared


def filter_if_selected(df, column, selected_values, enabled=True):
    if not enabled:
        return df

    if not selected_values:
        return df.iloc[0:0]

    return df[df[column].isin(selected_values)]


def render_date_calendar(df, key):
    available_dates = sorted(date for date in df["Date"].dropna().unique())

    if not available_dates:
        st.caption("No available invoice dates for the current filters.")
        return []

    available_set = set(available_dates)
    selected_key = f"{key}_selected_dates"
    month_index_key = f"{key}_month_index"

    st.session_state[selected_key] = sorted(
        date for date in st.session_state.get(selected_key, []) if date in available_set
    )

    available_months = (
        df.dropna(subset=["Date", "Year", "Month Number"])[["Year", "Month Number"]]
        .drop_duplicates()
        .sort_values(["Year", "Month Number"])
        .apply(lambda row: (int(row["Year"]), int(row["Month Number"])), axis=1)
        .tolist()
    )

    if month_index_key not in st.session_state:
        st.session_state[month_index_key] = 0

    st.session_state[month_index_key] = min(
        max(st.session_state[month_index_key], 0),
        len(available_months) - 1,
    )

    previous_column, title_column, next_column = st.columns([1, 3, 1])

    if previous_column.button(
        "Previous",
        disabled=st.session_state[month_index_key] == 0,
        key=f"{key}_previous_month",
    ):
        st.session_state[month_index_key] -= 1

    if next_column.button(
        "Next",
        disabled=st.session_state[month_index_key] == len(available_months) - 1,
        key=f"{key}_next_month",
    ):
        st.session_state[month_index_key] += 1

    st.session_state[month_index_key] = min(
        max(st.session_state[month_index_key], 0),
        len(available_months) - 1,
    )

    year, month_number = available_months[st.session_state[month_index_key]]
    month_name = calendar.month_name[month_number]

    title_column.markdown(f"**{month_name} {year}**")

    header_columns = st.columns(7)
    for index, day_name in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
        header_columns[index].markdown(f"**{day_name}**")

    selected_dates = set(st.session_state[selected_key])

    for week in calendar.monthcalendar(year, month_number):
        columns = st.columns(7)
        for day_index, day in enumerate(week):
            if day == 0:
                columns[day_index].markdown("&nbsp;", unsafe_allow_html=True)
                continue

            current_date = pd.Timestamp(year=year, month=month_number, day=day).date()
            is_available = current_date in available_set
            is_selected = current_date in selected_dates
            button_key = f"{key}_{year}_{month_number}_{day}"
            clicked = columns[day_index].button(
                str(day),
                disabled=not is_available,
                key=button_key,
                use_container_width=True,
            )

            if clicked and is_available:
                if is_selected:
                    selected_dates.discard(current_date)
                else:
                    selected_dates.add(current_date)

                st.session_state[selected_key] = sorted(selected_dates)
                st.rerun()

    st.session_state[selected_key] = sorted(selected_dates)

    if st.session_state[selected_key]:
        formatted_dates = ", ".join(
            date.strftime("%d-%m-%Y") for date in st.session_state[selected_key]
        )
        st.caption(f"Selected dates: {formatted_dates}")
    else:
        st.caption("No dates selected.")

    return st.session_state[selected_key]


def build_output(df):
    output = df[OUTPUT_COLUMNS].copy()
    return output.reset_index(drop=True)


def render():
    uploaded_file = st.file_uploader(
        "Upload Excel or CSV file",
        type=["csv", "xlsx"],
        key="completed_upload",
    )

    if uploaded_file is None:
        st.info("Upload an Excel or CSV file to start.")
        st.stop()

    try:
        input_data = read_input_file(uploaded_file)
    except Exception as error:
        st.error(f"Could not read this file: {error}")
        st.stop()

    if input_data.empty:
        st.warning("The uploaded file does not contain any rows.")
        st.stop()

    found_columns, missing_columns = find_columns(input_data, EXPECTED_COLUMNS)

    if missing_columns:
        st.warning(
            "Missing expected columns: "
            + ", ".join(missing_columns)
            + ". Missing filters are disabled and missing output fields will be blank."
        )

    data = prepare_data(input_data, found_columns)

    st.success(f"Loaded {len(data)} rows.")

    st.subheader("Input Preview")
    st.dataframe(input_data, use_container_width=True, hide_index=True)

    customer_options = unique_options(data["Customer"])
    customer_filter_enabled = "Customer" in found_columns
    selected_customers = select_filter(
        "Customers",
        customer_options,
        key="completed_customers",
        disabled=not customer_filter_enabled,
    )
    customer_rows = filter_if_selected(
        data,
        "Customer",
        selected_customers,
        enabled=customer_filter_enabled,
    )

    sub_customer_options = unique_options(customer_rows["Sub Customer"])
    sub_customer_filter_enabled = "Sub Customer" in found_columns
    selected_sub_customers = select_filter(
        "Sub Customers",
        sub_customer_options,
        key="completed_sub_customers",
        disabled=not sub_customer_filter_enabled or customer_rows.empty,
    )
    sub_customer_rows = filter_if_selected(
        customer_rows,
        "Sub Customer",
        selected_sub_customers,
        enabled=sub_customer_filter_enabled,
    )

    has_invoice_dates = "Invoice Date" in found_columns and data["_Invoice Date Parsed"].notna().any()

    years = (
        sorted(sub_customer_rows["Year"].dropna().astype(int).unique(), reverse=True)
        if has_invoice_dates and not sub_customer_rows.empty
        else []
    )
    selected_years = select_filter(
        "Years",
        years,
        key="completed_years",
        disabled=not has_invoice_dates or sub_customer_rows.empty,
    )
    if has_invoice_dates:
        year_rows = (
            sub_customer_rows[sub_customer_rows["Year"].isin(selected_years)]
            if selected_years
            else data.iloc[0:0]
        )
    else:
        year_rows = sub_customer_rows

    st.markdown("Dates")
    selected_dates = render_date_calendar(
        year_rows,
        key="completed_dates",
    ) if has_invoice_dates and not year_rows.empty else []

    if has_invoice_dates:
        date_rows = (
            year_rows[year_rows["Date"].isin(selected_dates)]
            if selected_dates
            else data.iloc[0:0]
        )
    else:
        date_rows = year_rows

    created_by_options = unique_options(date_rows["Created By"])
    created_filter_enabled = "Created" in found_columns
    selected_created_by = select_filter(
        "Created By",
        created_by_options,
        key="completed_created_by",
        disabled=not created_filter_enabled or date_rows.empty,
    )
    created_rows = filter_if_selected(
        date_rows,
        "Created By",
        selected_created_by,
        enabled=created_filter_enabled,
    )

    subscription_options = unique_options(created_rows["Subscription"])
    subscription_filter_enabled = "Subscription" in found_columns
    selected_subscriptions = select_filter(
        "Subscription",
        subscription_options,
        key="completed_subscription",
        disabled=not subscription_filter_enabled or created_rows.empty,
    )
    subscription_rows = filter_if_selected(
        created_rows,
        "Subscription",
        selected_subscriptions,
        enabled=subscription_filter_enabled,
    )

    sent_by_options = unique_options(subscription_rows["Sent By"])
    sent_filter_enabled = "Sent" in found_columns
    selected_sent_by = select_filter(
        "Sent By",
        sent_by_options,
        key="completed_sent_by",
        disabled=not sent_filter_enabled or subscription_rows.empty,
    )
    final_rows = filter_if_selected(
        subscription_rows,
        "Sent By",
        selected_sent_by,
        enabled=sent_filter_enabled,
    )

    if final_rows.empty:
        st.warning("No rows match the selected filters.")
        st.session_state.pop("completed_output", None)
        return

    st.session_state["completed_output"] = build_output(final_rows)

    st.subheader("Output")
    st.dataframe(
        st.session_state["completed_output"],
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        label="Download Output",
        data=dataframe_to_csv(st.session_state["completed_output"]),
        file_name="completed_output.csv",
        mime="text/csv",
        key="completed_download",
    )
