# RGPV Analytical Scraper

Fetch a range of RGPV student results with Selenium and OCR, then write a ranked Excel report. The command-line interface asks for a roll-number range, semester, course, headless mode and worker count.

## Setup

The current entry point requires Windows because it imports `winsound` at startup. Install Python 3.10+, Google Chrome and the [Microsoft Visual C++ x64 runtime](https://aka.ms/vs/17/release/vc_redist.x64.exe) required by the OCR dependency.

```powershell
git clone https://github.com/Rhythm-Sanghi/RGPV_Result.git
cd RGPV_Result
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python scraper.py
```

`undetected-chromedriver` manages ChromeDriver. Browser and portal changes can require dependency or selector updates.

## Inputs

| Input | Example |
|---|---|
| Starting roll number | `0101CS211001` |
| Ending roll number | `0101CS211060` |
| Semester | `7` |
| Course | `B.Tech` |
| Headless mode | `Y` or `N` |
| Parallel workers | `2` |

## Output and resume

`Output/` is created next to `scraper.py` and is ignored by Git.

- `results.xlsx`: Results ranked by SGPA, Analytics and Backlog List worksheets. Failed grades/statuses are highlighted in red and passing statuses in green.
- `skipped_rolls.txt`: rolls that could not be fetched and their failure reasons.
- `debug.log`: worker diagnostics.
- `.session.json` and `LAST_UPDATE.txt`: session and progress state.

Running the same range again reads the existing workbook and skips recorded roll numbers. Keep the workbook to retain those records. Review skipped rolls separately.

## Behavior and limits

Workers use randomized delays and retry OCR with `ddddocr`; OCR can fail. Detected access-denied or rate-limit pages pause processing for operator intervention. These measures do not guarantee uninterrupted access. Portal markup changes, unavailable results and browser startup failures can leave gaps in a report.

## Files

| File | Purpose |
|---|---|
| `scraper.py` | Prompts, worker coordination, progress and resume |
| `engine.py` | Browser setup, portal interaction and OCR |
| `excel_report.py` | Workbook generation |
| `requirements.txt` | Python dependencies |
