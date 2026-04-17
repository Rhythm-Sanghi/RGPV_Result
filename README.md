# 🎓 RGPV Analytical Scraper

A multi-threaded Python automation tool that batch-fetches student results from the RGPV result portal, solves CAPTCHAs via OCR, and generates a formatted, ranked Excel report.

---

## 📦 Project Structure

```
RGPV_Result/
├── scraper.py          ← Entry point — run this
├── engine.py           ← Selenium scraping engine + OCR CAPTCHA bypass
├── excel_report.py     ← openpyxl-based Excel report builder
├── requirements.txt    ← Python dependencies
└── Output/             ← Auto-created on first run
    ├── results.xlsx        ← Generated ranked report
    ├── skipped_rolls.txt   ← Log of failed/skipped roll numbers
    └── debug.log           ← Per-thread debug trace
```

---

## 🚀 Setup

### 1. Clone the repository
```bash
git clone https://github.com/your-username/RGPV_Result.git
cd RGPV_Result
```

### 2. Create & activate a virtual environment (recommended)
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

### 4. Ensure Google Chrome is installed
The scraper uses `undetected-chromedriver` which auto-downloads the matching ChromeDriver.

---

## ▶️ Running

```bash
python scraper.py
```

You will be prompted for:

| Input | Example |
|---|---|
| Starting Roll Number | `0101CS211001` |
| Ending Roll Number | `0101CS211060` |
| Semester | `7` |
| Course Type | `B.Tech` |
| Headless mode? | `Y` (background) or `N` (visible browser) |
| Parallel Threads | `2` |

---

## 📊 Output: Excel Report (`Output/results.xlsx`)

| Sheet | Contents |
|---|---|
| **Results** | All students ranked by SGPA (high → low), colour-coded grades |
| **Analytics** | Batch pass %, average SGPA, subject-wise pass/fail breakdown |
| **Backlog List** | Students with any F grade or Failed result status |

### Formatting
- 🔴 Red cells → F grade / Fail status
- 🟢 Green cells → Pass status
- 📌 Frozen header row, auto-column widths, bold dark-navy headers

---

## 🔒 Anti-Ban Features

- **Randomised delay** — 1–2 s between requests to avoid rate-limiting
- **CAPTCHA retry loop** — up to 5 OCR retries per roll number via `ddddocr`
- **IP-block detection** — detects 503 / "Access Denied" → beeps and pauses for manual reset
- **Headless UC driver** — bypasses bot detection via `undetected-chromedriver`

---

## 🔁 Resume / Crash Recovery

If the scraper is interrupted mid-batch:

1. Re-run `python scraper.py` with the **same roll range**.
2. It reads `Output/results.xlsx` and **automatically skips already-scraped rolls**.
3. Only the remaining pending rolls are processed.

---

## 📋 Skipped Rolls Log

`Output/skipped_rolls.txt` records every roll number that failed after all retries, along with the reason.

---

## ⚙️ Requirements

- Python 3.10+
- Google Chrome (latest)
- Windows / macOS / Linux

> **Note:** The audible IP-block alert (`winsound.Beep`) is Windows-only and is silently skipped on other platforms.
