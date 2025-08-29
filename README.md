# Windows Event Log Analyzer

**Terminal UI to read Windows Event Logs, aggregate by EventID, and write a nicely formatted Excel report.**

This utility provides a text-based UI to choose a Windows Event Log and a start/end datetime range, scans the chosen log using the legacy `win32evtlog` API (pywin32), aggregates events by normalized `EventID` (EventID & 0xFFFF), and writes a human-friendly Excel report (`log_analysis.xlsx`) with helpful formatting (freeze header, auto-filter, wrapped descriptions, computed column widths, etc).

---

## Features

* Interactive selection of event log (`Application`, `Security`, `Setup`, `System`, `ForwardedEvents`).
* Accepts `YYYY-MM-DD` or `YYYY-MM-DD HH:MM[:SS]` formats for start/end datetimes.
* Aggregates events by normalized EventID and computes frequency, sources, levels, timestamps and descriptions.
* Writes `log_analysis.xlsx` and backs up any existing file with a timestamp suffix.
* **Output will be exclusively stored in your Windows user's `Downloads` folder.**
* Applies post-processing with `openpyxl` for spreadsheet usability.
* Graceful, informative error messages (e.g. missing dependencies, permission hints).

---

## Requirements

* **Platform:** Windows
* **Python:** 3.7+
* **Python packages:**

  * `pywin32` (for `win32evtlog`, `win32evtlogutil`)
  * `pandas`
  * `openpyxl`

Install requirements with:

```bash
pip install pywin32 pandas openpyxl
```

> <span style="color: red; font-weight: bold;">**IMPORTANT:** If you want to access the <u>Security</u> log, you must run this script as Administrator (elevated prompt)!</span>

---

## Files

* `logger.py` — main script (interactive).
* Output: `log_analysis.xlsx` — generated Excel report. If a file exists, it will be backed up with a timestamp appended.
* **The Excel report will always be saved in your Windows user's `Downloads` folder.**

---

## Usage

Run the script from a Windows command prompt or PowerShell:

```powershell
python logger.py
```

Typical flow:

1. Choose Event Log to read.
2. Enter `Start datetime` and `End datetime`.
3. Confirm to proceed with aggregation.
4. Result: `log_analysis.xlsx` will be created in your `Downloads` folder.

Example:

```
=== Windows Event Log Analyzer ===
Enter choice: 1
Selected: Application

Start datetime: 2024-09-01
End datetime: 2024-09-30
Accepted range: 2024-09-01 00:00:00 -> 2024-09-30 23:59:59

Proceed? (y/n): y
Wrote report to 'C:\Users\<YourUser>\Downloads\log_analysis.xlsx'.
```

---

## Output format (columns)

* `SI no` — serial number
* `EventID` — normalized EventID
* `Source` — event source(s)
* `Level` — human-readable level
* `Task Category` — event category
* `Timestamp (logged)` — timestamps
* `Description` — concatenated descriptions
* `Frequency` — number of events found

---

## Notes & Troubleshooting

* Works only on Windows.
* <span style="color: red; font-weight: bold;">**To read the Security log, you must run the script as Administrator!**</span>
* If you give only a date (YYYY-MM-DD) as `end`, it is treated as `23:59:59`.
* Large logs may take time to scan.
* If formatting fails, a plain Excel file is still written.
* **All output Excel files are saved in your `Downloads` folder.**

---

**Thanks for using this tool!**
