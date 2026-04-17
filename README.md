# RGPV Analytical Scraper

A multi-threaded Python tool that batch-fetches student results from the RGPV result portal, solves CAPTCHAs automatically via OCR, and produces a ranked, formatted Excel report suitable for departmental review.

---

## Project Structure

```
RGPV_Result/
├── scraper.py          # Entry point — run this
├── engine.py           # Selenium scraping engine with OCR CAPTCHA bypass
├── excel_report.py     # Excel report builder using openpyxl
├── requirements.txt    # Python dependencies
└── Output/             # Created automatically on first run
    ├── results.xlsx        # Generated ranked report
    ├── skipped_rolls.txt   # Roll numbers that failed after all retries
    └── debug.log           # Per-thread debug trace
```

---

## Setup

### 1. Clone the repository
```bash
git clone https://github.com/Rhythm-Sanghi/RGPV_Result.git
cd RGPV_Result
```

### 2. Create and activate a virtual environment
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Make sure Google Chrome is installed
The scraper uses `undetected-chromedriver`, which automatically downloads the correct ChromeDriver version for your browser.

---

## Running

```bash
python scraper.py
```

You will be prompted for the following inputs:

| Input | Example |
|---|---|
| Starting Roll Number | `0101CS211001` |
| Ending Roll Number | `0101CS211060` |
| Semester | `7` |
| Course Type | `B.Tech` |
| Headless mode? | `Y` to run in background, `N` for visible browser |
| Parallel Threads | `2` |

---

## Output

The generated `Output/results.xlsx` contains three sheets:

| Sheet | Contents |
|---|---|
| Results | All students ranked by SGPA (highest to lowest), with colour-coded grades |
| Analytics | Batch pass percentage, average SGPA, subject-wise pass/fail breakdown |
| Backlog List | Students with any F grade or a Failed result status |

Formatting details:
- Red cells indicate an F grade or Fail status
- Green cells indicate a Pass status
- The header row is frozen, columns are auto-sized, and headers use bold dark-navy styling

---

## Anti-Ban Measures

- Randomised delay of 1 to 2 seconds between requests
- CAPTCHA retry loop with up to 5 OCR attempts per roll number using `ddddocr`
- IP-block detection that pauses all threads when a 503 or "Access Denied" response is detected, prompting you to reset your connection before continuing
- Headless mode via `undetected-chromedriver` to reduce bot-detection risk

---

## Resume After Interruption

If the scraper is stopped mid-run, just re-run it with the same roll number range. It reads the existing `Output/results.xlsx` and skips any roll numbers that were already scraped, picking up from where it left off.

---

## Skipped Rolls

Any roll number that fails after all retries is logged to `Output/skipped_rolls.txt` along with the reason for failure.

---

## Requirements

- Python 3.10 or higher
- Google Chrome (latest version)
- Windows, macOS, or Linux

> **Note:** The audible alert on IP block (`winsound.Beep`) is Windows-only and is silently ignored on other platforms.
