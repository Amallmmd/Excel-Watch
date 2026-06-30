from pathlib import Path

import streamlit as st

from excel_watch import completed, outstanding


st.set_page_config(page_title="Invoice Watch", layout="centered")


def load_theme():
    css_path = Path(__file__).with_name("style.css")
    st.markdown(f"<style>{css_path.read_text()}</style>", unsafe_allow_html=True)


load_theme()

st.title("Invoice Watch")

selected_option = st.selectbox(
    "Options",
    ["Outstanding", "Completed"],
    key="selected_option",
)

if selected_option == "Outstanding":
    outstanding.render()
else:
    completed.render()
