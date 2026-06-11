import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excelwatch - Excel Editor", layout="centered")

st.title("📊 Excel File Editor")
st.markdown("Upload an Excel file, select a month, and export the updated file.")

# 1. Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    st.success(f"✅ File '{uploaded_file.name}' uploaded successfully!")
    st.markdown("### Preview of uploaded data")
    st.dataframe(df)

    # 2. Select a month
    st.markdown("### Select a Month")
    months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    selected_month = st.selectbox("Choose a month", months, index=None, placeholder="Select a month...")

    if selected_month:
        st.info(f"Selected month: **{selected_month}**")

        # 3. Export updated Excel button
        st.markdown("### Export Updated Excel")

        # Add the selected month as a new column in the dataframe for export
        df_export = df.copy()
        df_export["Selected_Month"] = selected_month

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Sheet1")

        st.download_button(
            label="📥 Download Updated Excel",
            data=buffer.getvalue(),
            file_name=f"updated_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("👆 Please upload an Excel file to get started.")