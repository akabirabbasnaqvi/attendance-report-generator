# Attendance Report Generator

Small Streamlit application for comparing device attendance punches with a dated staff schedule.

## Run

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

Upload the attendance export and schedule workbook, select the year and month, then generate the Excel report. The report contains an employee summary and daily detail. `OFF`, `WFH`, `LV`, and `Resigned` are kept as schedule states; scheduled shifts use the first punch of the day and a 15-minute grace period.