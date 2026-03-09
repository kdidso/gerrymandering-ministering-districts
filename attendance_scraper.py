from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ============================================================
# CONFIG
# ============================================================

LCR_BASE = "https://lcr.churchofjesuschrist.org"
ATTENDANCE_URL = f"{LCR_BASE}/report/class-and-quorum-attendance/overview?lang=eng"

USERNAME = os.getenv("LCR_USERNAME", "").strip()
PASSWORD = os.getenv("LCR_PASSWORD", "").strip()

START_DATE = os.getenv("START_DATE", "2025-12-28").strip()
END_DATE = os.getenv("END_DATE", "2026-03-08").strip()

OUTPUT_DIR = "data"
DEBUG_DIR = "debug"

HEADLESS = True

DEFAULT_WAIT = 30
LONG_WAIT = 60
POST_LOGIN_SLEEP = 2.0
POST_NAV_SLEEP = 1.5
SCROLL_SLEEP = 0.35

# The page shows 5 weeks at a time. Moving left one click shifts by one week.
DATE_BLOCK_CLICKS = 5

# Selectors from your inspection
DATE_HEADER_CLASS = "sc-arbpvo-0 sc-arbpvo-1 fRnemr fhFlsT sc-473b494-0 kpqFIx"
LEFT_ARROW_SVG_CLASS = "sc-2b11ed23-0 kPPSzB"

# Fallback table selector
ROW_XPATH = "//table//tbody//tr"

SAVE_DEBUG_ON_ERROR = True


# ============================================================
# BASIC HELPERS
# ============================================================

def log(msg: str) -> None:
    print(f"[INFO] {msg}")


def warn(msg: str) -> None:
    print(f"[WARN] {msg}")


def err(msg: str) -> None:
    print(f"[ERROR] {msg}", file=sys.stderr)


def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def midpoint_date(start_dt: date, end_dt: date) -> date:
    return start_dt + timedelta(days=(end_dt - start_dt).days // 2)


def dedupe_preserve_order(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def format_excel_header(dt: date) -> str:
    return f"{dt.strftime('%b')} {dt.day} {dt.year}"


def classes_to_css(class_string: str) -> str:
    return "." + ".".join(class_string.split())


def save_debug_page(driver: webdriver.Chrome, name: str) -> None:
    if not SAVE_DEBUG_ON_ERROR:
        return

    debug_dir = ensure_dir(DEBUG_DIR)

    try:
        html_path = debug_dir / f"{name}.html"
        html_path.write_text(driver.page_source, encoding="utf-8")
        log(f"Saved debug HTML to {html_path}")
    except Exception as ex:
        warn(f"Could not save debug HTML: {ex}")

    try:
        png_path = debug_dir / f"{name}.png"
        driver.save_screenshot(str(png_path))
        log(f"Saved debug screenshot to {png_path}")
    except Exception as ex:
        warn(f"Could not save debug screenshot: {ex}")


def wait_present(driver: webdriver.Chrome, locator: Tuple[str, str], timeout: int = DEFAULT_WAIT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator))


def safe_text(element: WebElement) -> str:
    try:
        return normalize_space(element.text)
    except Exception:
        return ""


def is_probable_date_label(text: str) -> bool:
    return bool(re.match(r"^\d{1,2}\s+[A-Za-z]{3}$", normalize_space(text)))


def infer_year_for_label(label: str, start_dt: date, end_dt: date) -> Optional[date]:
    label = normalize_space(label)
    if not is_probable_date_label(label):
        return None

    mid = midpoint_date(start_dt, end_dt)
    best_candidate = None
    best_distance = None

    for year in range(start_dt.year - 1, end_dt.year + 2):
        try:
            candidate = datetime.strptime(f"{label} {year}", "%d %b %Y").date()
        except ValueError:
            continue

        distance = abs((candidate - mid).days)
        if best_distance is None or distance < best_distance:
            best_candidate = candidate
            best_distance = distance

    return best_candidate


# ============================================================
# DRIVER / LOGIN
# ============================================================

def make_driver() -> webdriver.Chrome:
    opts = ChromeOptions()
    if HEADLESS or os.getenv("CI", "").lower() == "true":
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1600,2200")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--lang=en-US")
    return webdriver.Chrome(options=opts)


def login(driver: webdriver.Chrome) -> None:
    if not USERNAME or not PASSWORD:
        err("Missing env vars LCR_USERNAME and/or LCR_PASSWORD")
        sys.exit(1)

    log("Opening LCR base login page")
    driver.get(LCR_BASE)

    try:
        user_input = wait_present(driver, (By.ID, "username-input"), LONG_WAIT)
        user_input.clear()
        user_input.send_keys(USERNAME)
        user_input.send_keys(Keys.ENTER)

        pwd_input = wait_present(driver, (By.ID, "password-input"), LONG_WAIT)
        pwd_input.clear()
        pwd_input.send_keys(PASSWORD)
        pwd_input.send_keys(Keys.ENTER)

        WebDriverWait(driver, LONG_WAIT).until(EC.url_contains(LCR_BASE))
        time.sleep(POST_LOGIN_SLEEP)
        log("Login submitted successfully")
    except Exception:
        save_debug_page(driver, "login_failure")
        raise RuntimeError("Automated login failed with the known username/password field flow.")


def goto_attendance(driver: webdriver.Chrome) -> None:
    log("Opening attendance page")
    driver.get(ATTENDANCE_URL)
    wait_for_attendance_table(driver)


# ============================================================
# ATTENDANCE PAGE HELPERS
# ============================================================

def wait_for_attendance_table(driver: webdriver.Chrome) -> None:
    try:
        WebDriverWait(driver, DEFAULT_WAIT).until(
            lambda d: len(d.find_elements(By.XPATH, ROW_XPATH)) > 0
        )
        time.sleep(POST_NAV_SLEEP)
        return
    except TimeoutException:
        save_debug_page(driver, "attendance_table_missing")
        raise RuntimeError("Could not detect attendance table rows.")


def get_visible_date_labels(driver: webdriver.Chrome) -> List[str]:
    labels: List[str] = []

    try:
        elems = driver.find_elements(By.CSS_SELECTOR, classes_to_css(DATE_HEADER_CLASS))
        for elem in elems:
            txt = safe_text(elem)
            if is_probable_date_label(txt):
                labels.append(txt)
    except Exception:
        pass

    return dedupe_preserve_order(labels)


def get_visible_window_dates(driver: webdriver.Chrome, start_dt: date, end_dt: date) -> List[date]:
    out: List[date] = []
    for label in get_visible_date_labels(driver):
        actual = infer_year_for_label(label, start_dt, end_dt)
        if actual is not None:
            out.append(actual)
    return sorted(set(out))


def extract_attendees_from_page_source(driver: webdriver.Chrome) -> List[dict]:
    html = driver.page_source

    # Prefer the exact marker we saw in the debug html
    markers = [
        '"initialProps":{"attendees":',
        '"attendees":[',
    ]

    start = -1
    for marker in markers:
        start = html.find(marker)
        if start != -1:
            start = html.find("[", start)
            break

    if start == -1:
        save_debug_page(driver, "attendees_marker_missing")
        raise RuntimeError("Could not find attendees data in page source.")

    depth = 0
    end = None
    in_string = False
    escape = False

    for i in range(start, len(html)):
        ch = html[i]

        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue

        if ch == '"':
            in_string = True
        elif ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                end = i + 1
                break

    if end is None:
        save_debug_page(driver, "attendees_array_end_missing")
        raise RuntimeError("Could not find end of attendees array in page source.")

    attendees_json = html[start:end]

    try:
        attendees = json.loads(attendees_json)
    except Exception:
        debug_dir = ensure_dir(DEBUG_DIR)
        raw_path = debug_dir / "attendees_raw.json.txt"
        raw_path.write_text(attendees_json, encoding="utf-8")
        raise RuntimeError("Failed to parse attendees JSON from page source.")

    if not isinstance(attendees, list):
        raise RuntimeError("Attendees data was not a list.")

    return attendees


def scrape_current_loaded_window(
    driver: webdriver.Chrome,
    attendance_data: Dict[str, Dict[date, bool]],
    start_dt: date,
    end_dt: date,
) -> Tuple[int, List[date]]:
    wait_for_attendance_table(driver)

    attendees = extract_attendees_from_page_source(driver)
    found_dates: Set[date] = set()
    scraped_rows = 0

    for person in attendees:
        name = normalize_space(person.get("displayName", ""))
        if not name:
            continue

        scraped_rows += 1
        attendance_data.setdefault(name, {})

        for entry in person.get("entries", []):
            iso = entry.get("date", {}).get("isoYearMonthDay")
            if not iso:
                continue

            dt = datetime.strptime(iso, "%Y-%m-%d").date()
            if not (start_dt <= dt <= end_dt):
                continue

            found_dates.add(dt)
            attendance_data[name][dt] = bool(entry.get("isMarkedAttended", False))

    return scraped_rows, sorted(found_dates)


# ============================================================
# LEFT ARROW NAVIGATION
# ============================================================

def click_left_arrow(driver: webdriver.Chrome) -> None:
    # Get the header area into a stable view first
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.6)

    xpaths = [
        "//th[.//*[name()='svg' and contains(@class,'kPPSzB')]]",
        "//th[.//*[name()='svg' and contains(@class,'sc-2b11ed23-0')]]",
        "//span[.//*[name()='svg' and contains(@class,'kPPSzB')]]/ancestor::th[1]",
    ]

    for xp in xpaths:
        try:
            elem = driver.find_element(By.XPATH, xp)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            time.sleep(SCROLL_SLEEP)
            driver.execute_script("arguments[0].click();", elem)
            time.sleep(1.0)
            return
        except Exception:
            continue

    for xp in xpaths:
        try:
            elem = driver.find_element(By.XPATH, xp)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            time.sleep(SCROLL_SLEEP)
            ActionChains(driver).move_to_element(elem).click().perform()
            time.sleep(1.0)
            return
        except Exception:
            continue

    save_debug_page(driver, "left_arrow_click_failed")
    raise RuntimeError("Could not click the left attendance navigation arrow.")


def click_left_arrow_block(driver: webdriver.Chrome, clicks: int = DATE_BLOCK_CLICKS) -> None:
    for i in range(clicks):
        old_labels = get_visible_date_labels(driver)
        log(f"Before click {i + 1}: {old_labels}")

        click_left_arrow(driver)

        changed = False
        for _ in range(20):
            time.sleep(0.7)
            new_labels = get_visible_date_labels(driver)
            if new_labels and new_labels != old_labels:
                log(f"After click {i + 1}: {new_labels}")
                changed = True
                break

        if not changed:
            save_debug_page(driver, f"left_arrow_click_{i+1}_failed")
            raise RuntimeError(f"Left-arrow click {i+1} did not change the visible date headers.")


# ============================================================
# MAIN SCRAPE LOOP
# ============================================================

def scrape_attendance(
    driver: webdriver.Chrome,
    start_dt: date,
    end_dt: date,
) -> Tuple[Dict[str, Dict[date, bool]], List[date]]:
    attendance_data: Dict[str, Dict[date, bool]] = {}
    collected_dates: Set[date] = set()

    goto_attendance(driver)

    safety_counter = 0
    seen_windows: Set[Tuple[str, ...]] = set()

    while True:
        safety_counter += 1
        if safety_counter > 20:
            raise RuntimeError("Safety stop reached while paging through date windows.")

        visible_window_dates = get_visible_window_dates(driver, start_dt, end_dt)
        if not visible_window_dates:
            save_debug_page(driver, f"window_dates_missing_{safety_counter}")
            raise RuntimeError("Could not detect visible date headers on attendance page.")

        window_key = tuple(d.isoformat() for d in visible_window_dates)
        log(f"Visible week window: {list(window_key)}")

        if window_key in seen_windows:
            warn("Detected repeated week window; stopping to avoid infinite loop.")
            break
        seen_windows.add(window_key)

        rows_scraped, dates_found = scrape_current_loaded_window(driver, attendance_data, start_dt, end_dt)
        for dt in dates_found:
            collected_dates.add(dt)
        log(f"Rows scraped from source: {rows_scraped}; dates: {[d.isoformat() for d in dates_found]}")

        earliest_visible = min(visible_window_dates)
        if earliest_visible <= start_dt:
            log("Reached earliest required date window.")
            break

        log(f"Moving left to older {DATE_BLOCK_CLICKS}-week block")
        click_left_arrow_block(driver, clicks=DATE_BLOCK_CLICKS)

    final_dates = sorted(dt for dt in collected_dates if start_dt <= dt <= end_dt)
    if not final_dates:
        raise RuntimeError("No attendance dates were collected in the requested range.")

    return attendance_data, final_dates


# ============================================================
# EXCEL OUTPUT
# ============================================================

def write_excel(attendance_data: Dict[str, Dict[date, bool]], all_dates: List[date], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"

    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    percent_fill = PatternFill(fill_type="solid", fgColor="E2F0D9")
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    headers = ["Name", "% activity"] + [format_excel_header(dt) for dt in all_dates]
    ws.append(headers)

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = bold
        cell.fill = header_fill
        cell.alignment = left if col_idx == 1 else center

    for row_idx, name in enumerate(sorted(attendance_data.keys(), key=lambda s: s.casefold()), start=2):
        per_date = attendance_data[name]
        total = len(all_dates)
        present_count = sum(1 for d in all_dates if per_date.get(d, False))
        pct = (present_count / total) if total else 0.0

        ws.cell(row=row_idx, column=1, value=name)

        pct_cell = ws.cell(row=row_idx, column=2, value=pct)
        pct_cell.number_format = "0%"
        pct_cell.fill = percent_fill

        for col_idx, dt in enumerate(all_dates, start=3):
            ws.cell(row=row_idx, column=col_idx, value="☑" if per_date.get(dt, False) else "☐")

    ws.freeze_panes = "C2"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 12

    for col_idx in range(3, 3 + len(all_dates)):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = left if cell.column == 1 else center

    wb.save(out_path)


# ============================================================
# MAIN
# ============================================================

def main() -> int:
    if not USERNAME or not PASSWORD:
        err("Missing LCR_USERNAME and/or LCR_PASSWORD environment variables.")
        return 1

    start_dt = parse_iso_date(START_DATE)
    end_dt = parse_iso_date(END_DATE)

    if start_dt > end_dt:
        err("START_DATE must be on or before END_DATE.")
        return 1

    output_dir = ensure_dir(OUTPUT_DIR)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_dir / f"attendance_{start_dt.isoformat()}_to_{end_dt.isoformat()}_{timestamp}.xlsx"

    driver = make_driver()
    try:
        login(driver)
        attendance_data, all_dates = scrape_attendance(driver, start_dt, end_dt)
        write_excel(attendance_data, all_dates, out_path)
        log(f"Excel output written to {out_path}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    return 0


if __name__ == "__main__":
    sys.exit(main())
