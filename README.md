# Attendance Report Generator

Small CustomTkinter desktop application for comparing device attendance punches with a dated staff schedule.

## Run

```powershell
python -m pip install -r requirements.txt
python app.py
```

Choose the attendance export and schedule workbook, select the month, then generate and export the Excel report. The report contains all employees found in either workbook, including ID-only attendance records and employees without a schedule match. `OFF`, `WFH`, `LV`, and `Resigned` are kept as schedule states; scheduled shifts use the first punch of the day and a 15-minute grace period.

## Build an executable

```powershell
pyinstaller --onefile --windowed --name AttendanceReport app.py
```

The executable will be created in `dist\AttendanceReport.exe`.