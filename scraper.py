import os
import sys
import time
import random
import json
import winsound
import threading
import ddddocr
import queue
from pathlib import Path
from tqdm import tqdm
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

import engine
import excel_report

# Ensure stdout supports Unicode for the visual banner
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass

BANNER = r"""
╔══════════════════════════════════════════════════════════════════╗
║        RGPV ANALYTICAL SCRAPER  —  Rhythm Sanghi                 ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝
"""

# Portable output directory — always next to this script, works on any machine
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Output")
RESULTS_XLSX = os.path.join(OUTPUT_DIR, "results.xlsx")
SKIPPED_LOG = os.path.join(OUTPUT_DIR, "skipped_rolls.txt")
SESSION_FILE = os.path.join(OUTPUT_DIR, ".session.json")
DEBUG_LOG = os.path.join(OUTPUT_DIR, "debug.log")
HEARTBEAT_FILE = os.path.join(OUTPUT_DIR, "LAST_UPDATE.txt")

BLOCKED_KEYWORDS = ["service unavailable", "access denied", "403 forbidden", "too many requests"]

# Shared Threading Resources
THREAD_LOCK = threading.Lock()
REPORT_LOCK = threading.Lock()
TQDM_LOCK  = threading.Lock()   # dedicated lock for tqdm updates
STOP_EVENT = threading.Event()
BLOCK_EVENT = threading.Event()

def _print_banner():
    try:
        print(BANNER)
    except:
        print("\n  [ RGPV ANALYTICAL SCRAPER ]\n")

def _prompt(label, default=None):
    suffix = f" [{default}]" if default else ""
    val = input(f"  ➤  {label}{suffix}: ").strip()
    return val if val else default

def _prompt_bool(label, default=True):
    suffix = "(Y/n)" if default else "(y/N)"
    val = input(f"  ➤  {label} {suffix}: ").strip().lower()
    if not val: return default
    return val in ("y", "yes")

def _log_thread_debug(msg: str):
    with THREAD_LOCK:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        timestamp = time.strftime("%H:%M:%S")
        with open(DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] [Thread-{threading.get_ident()}] {msg}\n")

def _load_existing_rolls() -> set:
    if not os.path.exists(RESULTS_XLSX): return set()
    try:
        df = pd.read_excel(RESULTS_XLSX, sheet_name="Results", header=1)
        if "Roll No" in df.columns:
            return set(str(r).strip() for r in df["Roll No"].dropna().tolist())
    except: pass
    return set()

def _load_existing_records() -> list:
    """Read back previously scraped records from results.xlsx so they are
    preserved in the Excel file after subsequent runs."""
    if not os.path.exists(RESULTS_XLSX): return []
    try:
        df = pd.read_excel(RESULTS_XLSX, sheet_name="Results", header=1)
        if df.empty or "Roll No" not in df.columns: return []
        records = []
        base_cols = {"Rank", "Roll No", "Name", "Father's Name", "Result", "SGPA", "CGPA"}
        subject_cols = [c for c in df.columns if c not in base_cols]
        for _, row in df.iterrows():
            roll = str(row.get("Roll No", "")).strip()
            if not roll or roll == "nan": continue
            rec = {
                "roll_no":       roll,
                "name":          str(row.get("Name", "") or ""),
                "father_name":   str(row.get("Father's Name", "") or ""),
                "result_status": str(row.get("Result", "") or ""),
                "sgpa":          str(row.get("SGPA", "") or ""),
                "cgpa":          str(row.get("CGPA", "") or ""),
                "subjects":      {},
            }
            for sub in subject_cols:
                val = str(row.get(sub, "") or "").strip()
                if val and val.lower() != "nan":
                    rec["subjects"][sub] = val
            records.append(rec)
        return records
    except Exception as e:
        print(f"  ⚠  Could not load existing records: {e}")
        return []

def _log_skipped(roll_no: str, reason: str):
    with THREAD_LOCK:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(SKIPPED_LOG, "a", encoding="utf-8") as f:
            f.write(f"{roll_no}\t{reason}\n")

def _is_blocked(page_source: str) -> bool:
    source_lower = page_source.lower()
    return any(kw in source_lower for kw in BLOCKED_KEYWORDS)

def _generate_roll_sequence(start: str, end: str) -> list:
    import re
    start, end = start.strip(), end.strip()
    m_start = re.match(r'^(.*[a-zA-Z])(\d+)$', start)
    m_end   = re.match(r'^(.*[a-zA-Z])(\d+)$', end)
    if m_start and m_end:
        prefix = m_start.group(1)
        s_num, e_num = int(m_start.group(2)), int(m_end.group(2))
        width = len(m_start.group(2))
        return [f"{prefix}{str(n).zfill(width)}" for n in range(s_num, e_num + 1)]
    return [str(n) for n in range(int(start), int(end) + 1)]

def worker_task(rolls_batch, semester, course_type, headless, progress, data_queue):
    # ── Phase 1: initialise browser & OCR ─────────────────────────────────
    # This is the only place a thread can die silently.  We catch it here,
    # log it, and advance the progress bar for every roll we'll never reach,
    # so the bar doesn't stall at the end of the run.
    driver = None
    ocr    = None
    try:
        ocr    = ddddocr.DdddOcr(beta=True, show_ad=False)
        driver = engine.build_driver(headless=headless)
        _log_thread_debug("Browser & OCR initialized.")
    except Exception as exc:
        _log_thread_debug(f"FATAL: Thread init failed — {str(exc)[:120]}")
        for roll_no in rolls_batch:
            _log_skipped(roll_no, "Thread init failed — browser did not start")
            with TQDM_LOCK:
                progress.update(1)
        return  # Exit this worker cleanly; other threads continue

    # ── Phase 2: scrape each roll number ──────────────────────────────────
    try:
        for roll_no in rolls_batch:
            if STOP_EVENT.is_set(): break
            while BLOCK_EVENT.is_set(): time.sleep(1)
            _log_thread_debug(f"Processing roll: {roll_no}")
            try:
                page_source = driver.page_source if driver.current_url != "data:," else ""
                if _is_blocked(page_source):
                    BLOCK_EVENT.set()
                    continue
                result = engine.fetch_result(driver, roll_no, semester, course_type, ocr=ocr)
                if result is None:
                    _log_skipped(roll_no, "Failed after retries")
                elif result.get("status") == "NOT_FOUND":
                    _log_skipped(roll_no, "No result found on portal")
                    # Still push a stub so the roll appears in the Excel
                    data_queue.put({
                        "roll_no": roll_no, "name": "", "father_name": "",
                        "result_status": "NOT REGISTERED", "sgpa": "", "cgpa": "",
                        "subjects": {},
                    })
                else:
                    _log_thread_debug(f"Success: {roll_no}")
                    data_queue.put(result)  # Immediate handoff to merger
            except Exception as exc:
                _log_thread_debug(f"Error for {roll_no}: {str(exc)[:50]}")
            # Update progress bar safely with its own dedicated lock
            with TQDM_LOCK:
                progress.update(1)
            time.sleep(random.uniform(1, 2))
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
            try:
                driver.quit = lambda: None
            except Exception:
                pass

def merger_task(all_records, semester, course_type, data_queue):
    """Dedicated consumer thread: drains the data queue and updates Excel.
    
    Exits only when STOP_EVENT is set AND the queue is fully drained.
    """
    while True:
        # Exit condition: workers done AND queue empty
        if STOP_EVENT.is_set() and data_queue.empty():
            break
        try:
            new_record = data_queue.get(timeout=0.5)
            roll = new_record.get("roll_no")
            with THREAD_LOCK:
                if not any(r.get("roll_no") == roll for r in all_records):
                    all_records.append(new_record)
            with REPORT_LOCK:
                try:
                    excel_report.build_report(all_records, OUTPUT_DIR, semester, course_type)
                    print(f"  ➤ SUCCESS: {roll} captured and merged.")
                    sys.stdout.flush()
                except Exception as e:
                    _log_thread_debug(f"MERGE FAILED for {roll}: {e}")
            data_queue.task_done()
        except queue.Empty:
            continue  # Loop back and re-check STOP_EVENT

def main():
    _print_banner()
    print("  [ Session Configuration ]\n")
    s_roll = _prompt("Starting Roll Number")
    e_roll = _prompt("Ending Roll Number")
    sem = _prompt("Semester")
    c_type = _prompt("Course Type", default="B.Tech")
    headless = _prompt_bool("Run in Headless mode?", default=True)
    threads = int(_prompt("Number of Parallel Threads", default="2"))

    roll_list = _generate_roll_sequence(s_roll, e_roll)
    existing = _load_existing_rolls()
    pending = [r for r in roll_list if r not in existing]
    print(f"  ✔  Pending to scrape: {len(pending)}\n")
    if not pending: return

    # Seed all_records with already-scraped data so the Excel is always
    # the full cumulative dataset, not just the current session.
    all_records = _load_existing_records()
    if all_records:
        print(f"  ✔  Loaded {len(all_records)} existing records from previous runs.\n")
    data_queue = queue.Queue()

    # Round-robin distribution: guarantees every roll number is assigned to exactly
    # one batch with no skips, regardless of how evenly `len(pending)` divides by threads.
    batches = [pending[i::threads] for i in range(threads)]
    # Drop empty batches if threads > pending rolls
    batches = [b for b in batches if b]

    print(f"  [ Launching {len(batches)} Parallel Thread(s)... ]\n")
    progress = tqdm(total=len(pending), desc="  Scraping", unit="roll", ncols=80,
                    bar_format="  [{bar:30}] {n_fmt}/{total_fmt} ({percentage:3.0f}%) — ETA: {remaining}")

    merger = threading.Thread(
        target=merger_task,
        args=(all_records, sem, c_type, data_queue),
        daemon=True,
    )
    merger.start()

    try:
        with ThreadPoolExecutor(max_workers=len(batches)) as executor:
            futures = [
                executor.submit(worker_task, b, sem, c_type, headless, progress, data_queue)
                for b in batches
            ]
            while any(not f.done() for f in futures):
                if BLOCK_EVENT.is_set():
                    print("\n  ⚠️   IP BLOCK DETECTED! ALL THREADS PAUSED.")
                    input("  ➤  Reset IP/Restart Hotspot and press Enter to resume...")
                    BLOCK_EVENT.clear()
                time.sleep(1)
    except KeyboardInterrupt:
        STOP_EVENT.set()
    finally:
        # Signal the merger that all workers are done, then wait for it to
        # drain the remaining queue and write the final Excel report.
        STOP_EVENT.set()
        print("\n  [ All workers finished. Waiting for final report flush... ]")
        merger.join(timeout=60)  # Allow up to 60 s for final Excel write
        progress.close()
        print("  [ Final Report Generated in Output/ ]\n")

if __name__ == "__main__":
    main()
