# Attendance Report Generator

A Windows desktop application built with CustomTkinter and pandas. It compares device attendance punches with employee schedules and generates a monthly Excel report.

## Project files

- `app.py`: Desktop user interface and report workflow.
- `attendance_report.py`: Attendance parsing, schedule matching, report generation, and Excel export API.
- `excel_files/`: Local input and generated Excel workbooks. This folder is ignored by Git because the files can contain employee data.
- `requirements.txt`: Python dependencies.

## Setup

Open PowerShell in `D:\schedule_software`:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

If PowerShell blocks activation, use this for the current terminal session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
```

## Run the application

```powershell
python app.py
```

In the application:

1. Choose the schedule workbook.
2. Choose the device attendance workbook.
3. Select the report month.
4. Click `Generate Report`.
5. Click `Export to Excel` and choose where to save the report.

The application supports legacy `.xls` and modern `.xlsx` workbooks. The schedule determines shifts, off-days, work-from-home days, leave states, and resigned records. Attendance uses the first daily punch and a 15-minute grace period; seconds are ignored when applying the grace-period comparison.

## Report worksheets

New exports can contain these worksheets:

- `Employee Summary`: Monthly totals for present, on-time, late, absent, WFH, and leave days.
- `Daily Detail`: Employee-by-day schedule, clock-in, clock-out, status, punch count, and remarks.
- `Badge ID Punches`: Punches whose device record contains only a numeric badge ID.
- `Name Comparison`: Names or IDs found only in the device data or only in the schedule.

Excel headings are bold with colored headers. Each worksheet has filters and a frozen first row.

## Build the Windows executable

Install PyInstaller in the active environment, if needed:

```powershell
python -m pip install pyinstaller
```

Build the executable:

```powershell
pyinstaller --onefile --windowed --name AttendanceReport app.py
```

The executable will be created at:

```text
dist\AttendanceReport.exe
```

Run it with:

```powershell
.\dist\AttendanceReport.exe
```

## Git notes

Employee workbooks, generated reports, PyInstaller output, and the virtual environment are ignored by Git. Track only source code and project documentation:

```powershell
git status
git add app.py attendance_report.py README.md requirements.txt .gitignore
git commit -m "Describe the change"
git push
```