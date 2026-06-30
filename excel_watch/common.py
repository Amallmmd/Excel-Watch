import pandas as pd
import streamlit as st


def normalize_column_name(name):
    return "".join(char for char in str(name).lower() if char.isalnum())


def find_columns(df, column_aliases):
    alias_lookup = {}

    for display_name, aliases in column_aliases.items():
        for alias in aliases:
            alias_lookup[normalize_column_name(alias)] = display_name

    found_columns = {}

    for column in df.columns:
        display_name = alias_lookup.get(normalize_column_name(column))
        if display_name and display_name not in found_columns:
            found_columns[display_name] = column

    missing_columns = [
        display_name
        for display_name in column_aliases
        if display_name not in found_columns
    ]

    return found_columns, missing_columns


def read_input_file(uploaded_file):
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    return pd.read_excel(uploaded_file)


def dataframe_to_csv(df):
    return df.to_csv(index=False).encode("utf-8")


def clean_text(value):
    if pd.isna(value):
        return ""

    return str(value).strip()


def clean_text_series(series):
    return series.map(clean_text)


def unique_options(series, reverse=False):
    values = [
        value
        for value in clean_text_series(series).drop_duplicates().tolist()
        if value != ""
    ]

    return sorted(values, reverse=reverse)


def select_filter(label, options, key, disabled=False):
    options = list(options)
    disable_controls = disabled or not options
    use_all = st.checkbox(
        f"All {label.lower()}",
        disabled=disable_controls,
        key=f"{key}_all",
    )

    if use_all and not disable_controls:
        return options

    return st.multiselect(
        label,
        options,
        disabled=disable_controls,
        key=f"{key}_selected",
    )
