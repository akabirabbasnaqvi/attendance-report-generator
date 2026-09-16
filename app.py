from calendar import month_name

import pandas as pd
import streamlit as st

from attendance_report import build_report, load_punches, load_schedule, to_excel


st.set_page_config(page_title="Attendance Report", page_icon="📊", layout="wide")
st.title("Attendance Report Generator")
st.caption("Upload the device punches and the dated staff schedule to compare first arrival against each planned shift.")

attendance_file = st.file_uploader("Attendance file (.xls or .xlsx)", type=["xls", "xlsx"])
schedule_file = st.file_uploader("Schedule file (.xlsx)", type=["xlsx"])

if attendance_file and schedule_file:
    try:
        punches = load_punches(attendance_file)
        schedule = load_schedule(schedule_file)
        available_dates = sorted(schedule["Work Date"].unique())
        years = sorted({item.year for item in available_dates})
        selected_year = st.selectbox("Year", years)
        available_months = sorted({item.month for item in available_dates if item.year == selected_year})
        selected_month = st.selectbox("Month", available_months, format_func=lambda value: month_name[value])
        if st.button("Generate report", type="primary"):
            detail, summary = build_report(punches, schedule, selected_year, selected_month)
            st.session_state["detail"] = detail
            st.session_state["summary"] = summary
    except Exception as error:
        st.error(str(error))

if "summary" in st.session_state:
    st.subheader("Employee Summary")
    st.dataframe(st.session_state["summary"], use_container_width=True, hide_index=True)
    st.subheader("Daily Detail")
    st.dataframe(st.session_state["detail"], use_container_width=True, hide_index=True)
    st.download_button("Download Excel report", to_excel(st.session_state["detail"], st.session_state["summary"]), "attendance_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")