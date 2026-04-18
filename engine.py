import os
import re
import ssl
import time
import base64
import random
import io
import threading
import subprocess

# Disable SSL certificate verification for ChromeDriver auto-download.
# Needed on institutional/corporate networks with custom certificate authorities.
os.environ.setdefault("WDM_SSL_VERIFY", "0")
ssl._create_default_https_context = ssl._create_unverified_context
import ddddocr
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoAlertPresentException, UnexpectedAlertPresentException
)
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from PIL import Image


def _get_chrome_major_version() -> int:
    import winreg
    reg_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Google\Chrome\BLBeacon"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Google\Chrome\BLBeacon"),
    ]
    for hive, key_path in reg_keys:
        try:
            with winreg.OpenKey(hive, key_path) as k:
                ver_str, _ = winreg.QueryValueEx(k, "version")
                major = int(str(ver_str).split(".")[0])
                return major
        except Exception:
            pass

    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for path in chrome_paths:
        if os.path.exists(path):
            try:
                out = subprocess.check_output(
                    ["powershell", "-Command",
                     f"(Get-Item '{path}').VersionInfo.FileVersion"],
                    stderr=subprocess.DEVNULL, timeout=5
                ).decode().strip()
                major = int(out.split(".")[0])
                return major
            except Exception:
                pass

    return None

RESULT_URL = "https://result.rgpv.ac.in/Result/ProgramSelect.aspx"
MAX_CAPTCHA_RETRIES = 5


def _log_thread_debug(msg: str):
    """Local fallback logger — scraper.py overrides this with a richer version."""
    try:
        log_path = os.path.join(
            r"c:\Users\Test\Documents\Projects\RGPV_Result\Output", "debug.log"
        )
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        timestamp = time.strftime("%H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] [engine] {msg}\n")
    except Exception:
        pass

PROGRAM_URL_MAP = {
    "B.Tech":   "https://result.rgpv.ac.in/Result/BErslt.aspx",
    "B.E.":     "https://result.rgpv.ac.in/Result/BErslt.aspx",
    "M.Tech":   "https://result.rgpv.ac.in/Result/MTResult.aspx",
    "MBA":      "https://result.rgpv.ac.in/Result/MBAResult.aspx",
    "MCA":      "https://result.rgpv.ac.in/Result/MCAResult.aspx",
    "B.Pharm":  "https://result.rgpv.ac.in/Result/BPResult.aspx",
    "Diploma":  "https://result.rgpv.ac.in/Result/DIResult.aspx",
}

PROGRAM_RADIO_LABEL = {
    "B.Tech":  "B.Tech.",
    "B.E.":    "B.E.",
    "M.Tech":  "M.Tech.",
    "MBA":     "MAM",
    "MCA":     "M.C.A.",
    "B.Pharm": "B.Pharmacy",
    "Diploma": "Diploma",
}



# Serialises uc.Chrome() calls: undetected_chromedriver patches the
# ChromeDriver binary in a shared temp location, so simultaneous
# calls from multiple threads collide and all but one crash.
_DRIVER_INIT_LOCK = threading.Lock()


def build_driver(headless: bool) -> uc.Chrome:
    ua = UserAgent()
    options = uc.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(f"--user-agent={ua.random}")
    options.add_argument("--window-size=1280,900")
    options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_ver = _get_chrome_major_version()
    kwargs = {"options": options, "use_subprocess": True}
    if chrome_ver:
        kwargs["version_main"] = chrome_ver
    # Hold the lock only for the uc.Chrome() call — released immediately
    # after the browser process is spawned so other threads can start theirs.
    with _DRIVER_INIT_LOCK:
        driver = uc.Chrome(**kwargs)
    return driver


def _refresh_captcha(driver, wait):
    try:
        refresh_btn = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//img[contains(@id,'CaptchaImage')]//following-sibling::*|//a[contains(@id,'captcha') or contains(@onclick,'refresh') or contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'refresh')]")
            )
        )
        refresh_btn.click()
    except Exception:
        pass
    time.sleep(1)


def _read_captcha(driver, wait, ocr) -> str:
    from PIL import ImageEnhance, ImageFilter

    if not ocr:
        return ""

    try:
        captcha_img_el = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[@id='ctl00_ContentPlaceHolder1_imgCaptcha']|//img[contains(@src,'CaptchaImage')]")
            )
        )
        
        # Use a unique filename per thread to prevent race conditions
        temp_filename = f"captcha_{threading.get_ident()}_{random.randint(1000, 9999)}.png"
        
        captcha_img_el.screenshot(temp_filename)
        img = Image.open(temp_filename).convert("L")

        img = img.resize((img.width * 3, img.height * 3), Image.LANCZOS)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.5)

        enhancer2 = ImageEnhance.Sharpness(img)
        img = enhancer2.enhance(2.0)

        img = img.filter(ImageFilter.MedianFilter(size=3))

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        img_bytes_final = buf.getvalue()

        result = ocr.classification(img_bytes_final)
        
        # Cleanup
        try:
            if os.path.exists(temp_filename):
                os.remove(temp_filename)
        except Exception:
            pass

        text = result.strip().upper()
        text = re.sub(r'[^A-Z0-9]', '', text)
        return text

    except Exception:
        return ""




def _navigate_to_result_form(driver, wait, course_type: str, semester: str):
    direct_url = PROGRAM_URL_MAP.get(course_type, PROGRAM_URL_MAP["B.Tech"])
    radio_label = PROGRAM_RADIO_LABEL.get(course_type, "B.Tech.")

    # [SESSION BRIDGE] Always start at Program Selection to lock DEC-2025 state
    driver.get(RESULT_URL)
    time.sleep(2)

    try:
        # Select the program radio button (e.g. B.Tech.)
        radio = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 f"//input[@type='radio' and following-sibling::text()[normalize-space()='{radio_label}']]"
                 f"|//label[normalize-space()='{radio_label}']/preceding-sibling::input[@type='radio']"
                 f"|//label[normalize-space()='{radio_label}']/input[@type='radio']"
                 f"|//td[normalize-space()='{radio_label}']//input[@type='radio']"
                 f"|//input[@type='radio' and @value='{radio_label}']"
                )
            )
        )
        driver.execute_script("arguments[0].click();", radio)
        time.sleep(1)
        
        # Click the 'Go' button to transition to the actual result form
        submit = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' or @type='button'][contains(@value,'Go') or contains(@value,'Submit')]"))
        )
        submit.click()
        
        # Wait for the actual result page (e.g. BErslt.aspx) to load
        wait.until(lambda d: "ProgramSelect" not in d.current_url)
        time.sleep(1.5)
    except Exception as e:
        # Fallback to direct URL if the first page is being temperamental, 
        # but the session bridge is the preferred reliable path.
        if "BErslt" not in driver.current_url:
            driver.get(direct_url)
            time.sleep(2)


def _select_semester(driver, wait, semester: str):
    try:
        sem_select = wait.until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_ContentPlaceHolder1_drpSemester")
            )
        )
        Select(sem_select).select_by_value(str(semester))
    except Exception:
        try:
            sem_select2 = driver.find_element(
                By.XPATH,
                "//select[contains(@id,'drpSemester') or contains(@id,'Sem') or contains(@name,'Sem')]"
            )
            Select(sem_select2).select_by_value(str(semester))
        except Exception:
            pass

        grading_radio = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_ContentPlaceHolder1_rbtnlstSType_0']|//input[@type='radio' and @value='G']"))
        )
        if not grading_radio.is_selected():
            driver.execute_script("arguments[0].click();", grading_radio)
            # CRITICAL: Wait for ASP.NET Postback to stabilize the DOM
            time.sleep(2) 
    except Exception:
        pass


def fetch_result(driver, roll_no: str, semester: str, course_type: str, ocr=None) -> dict | None:
    wait = WebDriverWait(driver, 20)
    _navigate_to_result_form(driver, wait, course_type, semester)
    _select_semester(driver, wait, semester) # clicks radio button here
    time.sleep(1) # Extra stability wait

    for attempt in range(1, MAX_CAPTCHA_RETRIES + 1):
        # [RECOVERY] If buttons are missing (cloaked by portal), Reset to restore
        if not driver.find_elements(By.ID, "ctl00_ContentPlaceHolder1_btnviewresult"):
            try:
                reset_btn = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnReset")
                driver.execute_script("arguments[0].click();", reset_btn)
                time.sleep(2)
                _select_semester(driver, wait, semester)
                time.sleep(1)
            except Exception: pass

        try:
            # We solve CAPTCHA only after the form is stable
            captcha_img = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@src,'CaptchaImage')]")))
            captcha_code = _read_captcha(driver, wait, ocr)
            
            roll_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtrollno")))
            roll_input.clear()
            roll_input.send_keys(roll_no)
            
            captcha_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1")
            captcha_input.clear()
            captcha_input.send_keys(captcha_code)
            
            submit_btn = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnviewresult")
            driver.execute_script("arguments[0].click();", submit_btn)
        except Exception as e:
            _log_thread_debug(f"Submission sync error: {str(e)[:50]}")
            continue

        # Wait for result or error
        found = False
        for _ in range(12):
            page_src = driver.page_source
            if "pnlGrading" in page_src or "lblNameGrading" in page_src:
                found = True
                break
            try:
                if driver.switch_to.alert: break
            except Exception: pass
            time.sleep(0.5)

        try:
            alert = driver.switch_to.alert
            txt = alert.text.lower()
            alert.accept()
            if "not found" in txt or "enrollment" in txt:
                return {"status": "NOT_FOUND", "roll_no": roll_no}
            time.sleep(1)
            continue # Retry for captcha/other errors
        except NoAlertPresentException: pass

        page_src = driver.page_source
        if found:
            data = parse_result(page_src, roll_no)
            if data and data.get("name"): return data

        if "No Result" in page_src or "not found" in page_src.lower():
            return {"status": "NOT_FOUND", "roll_no": roll_no}

    return None



def parse_result(page_source: str, roll_no: str) -> dict | None:
    soup = BeautifulSoup(page_source, "lxml")
    for scr in soup(["script", "style"]):
        scr.decompose()

    data = {"roll_no": roll_no}

    def _get_by_id(partial_id):
        tag = soup.find(id=re.compile(partial_id, re.I))
        return tag.get_text(strip=True) if tag else ""

    data["name"] = _get_by_id("lblNameGrading") or _get_by_id("lblName")
    data["father_name"] = _get_by_id("lblFnameGrading") or _get_by_id("lblFname")
    data["result_status"] = _get_by_id("lblResultNew") or _get_by_id("lblResult")
    data["sgpa"] = _get_by_id("lblSGPA")
    data["cgpa"] = _get_by_id("lblcgpa")

    if not data["roll_no"]:
        data["roll_no"] = _get_by_id("lblRollNo") or roll_no

    subjects = {}
    for tr in soup.find_all("tr"):
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
        if len(cells) >= 3:
            # RGPV format natively: Subject code is usually first or second, Grade is near the end
            sub_name = "Unknown Subject"
            
            # Check cell 0 and cell 1 for a valid subject code
            code_cell_idx = -1
            sub_code_raw = ""
            for idx in [0, 1]:
                if re.match(r"^[A-Za-z]{2,5}\s?\-?\s*\d{3,5}", cells[idx]):
                    code_cell_idx = idx
                    sub_code_raw = cells[idx].upper()
                    break

            if code_cell_idx != -1:
                # If Subject Name column exists, it typically follows the code
                if code_cell_idx + 1 < len(cells) and len(cells[code_cell_idx + 1]) > 5:
                    sub_name = cells[code_cell_idx + 1].title()

                # Search remaining cells for Grade
                grade_val = ""
                for cell in reversed(cells[code_cell_idx + 1:]):
                    c_up = cell.upper()
                    if c_up in ("O", "A+", "A", "B+", "B", "C+", "C", "D+", "D", "P", "F", "EX", "AB", "F(ABS)", "F (ABS)", "FAIL"):
                        grade_val = c_up
                        break
                    
                if grade_val:
                    subjects[sub_code_raw] = grade_val

    data["subjects"] = subjects

    if not data["name"] and not subjects:
        return None

    return data
