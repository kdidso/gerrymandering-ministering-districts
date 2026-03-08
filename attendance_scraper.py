from __future__ import annotations

import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver import ChromeOptions
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

HEADLESS = True

DEFAULT_WAIT = 30
LONG_WAIT = 60
POST_CLICK_SLEEP = 1.1
POST_NAV_SLEEP = 1.5
POST_LOGIN_SLEEP = 2.0
SCROLL_SLEEP = 0.35

# 5 visible weeks at a time, per your description
DATE_BLOCK_CLICKS = 5

# Selectors from your inspection
NAME_CELL_CLASS = "sc-14fff288-0 llFqzd sc-24d75410-4 bMziSa sc-24d75410-0 lpsZiy"
PAGE_NUM_CLASS = "sc-66e0b3ee-0 hxtefx"
DATE_HEADER_CLASS = "sc-arbpvo-0 sc-arbpvo-1 fRnemr fhFlsT sc-473b494-0 kpqFIx"
LEFT_ARROW_CLASS = "sc-2b11ed23-0 kPPSzB"

# Generic fallbacks
HEADER_TH_XPATH = "//table//thead//th"
ROW_XPATH = "//table//tbody//tr"
TD_REL_XPATH = ".//td"

SAVE_DEBUG_ON_ERROR = True
DEBUG_DIR = "debug"


# ============================================================
# DATA TYPES
# ============================================================

@dataclass(frozen=True)
class ColumnInfo:
    column_index: int
    label_text: str
    actual_date: date


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
# ATTENDANCE TABLE DETECTION
# ============================================================

def wait_for_attendance_table(driver: webdriver.Chrome) -> None:
    try:
        WebDriverWait(driver, DEFAULT_WAIT).until(
            lambda d: len(d.find_elements(By.XPATH, ROW_XPATH)) > 0
        )
        time.sleep(POST_NAV_SLEEP)
        return
    except TimeoutException:
        pass

    try:
        WebDriverWait(driver, DEFAULT_WAIT).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, classes_to_css(NAME_CELL_CLASS))) > 0
        )
        time.sleep(POST_NAV_SLEEP)
        return
    except TimeoutException:
        pass

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

    if labels:
        return dedupe_preserve_order(labels)

    try:
        elems = driver.find_elements(By.XPATH, HEADER_TH_XPATH)
        for elem in elems:
            txt = safe_text(elem)
            if is_probable_date_label(txt):
                labels.append(txt)
    except Exception:
        pass

    return dedupe_preserve_order(labels)


def get_visible_columns(driver: webdriver.Chrome, start_dt: date, end_dt: date) -> List[ColumnInfo]:
    columns: List[ColumnInfo] = []

    try:
        headers = driver.find_elements(By.XPATH, HEADER_TH_XPATH)
        for idx, header in enumerate(headers):
            txt = safe_text(header)
            if not is_probable_date_label(txt):
                continue

            actual = infer_year_for_label(txt, start_dt, end_dt)
            if actual is None:
                continue

            columns.append(ColumnInfo(column_index=idx, label_text=txt, actual_date=actual))
    except Exception:
        pass

    if columns:
        return columns

    date_labels = get_visible_date_labels(driver)
    col_idx = 2  # Name, Gender first
    for lbl in date_labels:
        actual = infer_year_for_label(lbl, start_dt, end_dt)
        if actual is not None:
            columns.append(ColumnInfo(column_index=col_idx, label_text=lbl, actual_date=actual))
            col_idx += 1

    return columns


def get_visible_window_dates(driver: webdriver.Chrome, start_dt: date, end_dt: date) -> List[date]:
    out: List[date] = []
    for label in get_visible_date_labels(driver):
        actual = infer_year_for_label(label, start_dt, end_dt)
        if actual is not None:
            out.append(actual)
    return sorted(set(out))


# ============================================================
# MEMBER PAGE PAGINATION
# ============================================================

def get_bottom_page_numbers(driver: webdriver.Chrome) -> List[int]:
    found: List[int] = []

    try:
        elems = driver.find_elements(By.CSS_SELECTOR, classes_to_css(PAGE_NUM_CLASS))
        for elem in elems:
            txt = safe_text(elem)
            if txt.isdigit():
                found.append(int(txt))
    except Exception:
        pass

    if found:
        return sorted(set(found))

    try:
        elems = driver.find_elements(By.XPATH, "//*[normalize-space(text())!='']")
        for elem in elems:
            txt = safe_text(elem)
            if txt.isdigit():
                num = int(txt)
                if 1 <= num <= 200:
                    found.append(num)
    except Exception:
        pass

    found = sorted(set(found))
    return found if found else [1]


def click_member_page_number(driver: webdriver.Chrome, page_num: int) -> None:
    target = str(page_num)

    try:
        elems = driver.find_elements(By.CSS_SELECTOR, classes_to_css(PAGE_NUM_CLASS))
        for elem in elems:
            txt = safe_text(elem)
            if txt == target:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                time.sleep(SCROLL_SLEEP)
                driver.execute_script("arguments[0].click();", elem)
                time.sleep(POST_CLICK_SLEEP)
                wait_for_attendance_table(driver)
                return
    except Exception:
        pass

    try:
        elem = driver.find_element(By.XPATH, f"//*[normalize-space(text())='{target}']")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(SCROLL_SLEEP)
        driver.execute_script("arguments[0].click();", elem)
        time.sleep(POST_CLICK_SLEEP)
        wait_for_attendance_table(driver)
        return
    except Exception:
        pass

    raise RuntimeError(f"Could not click member page number {page_num}.")


# ============================================================
# WEEK NAVIGATION
# ============================================================

def click_left_arrow(driver: webdriver.Chrome) -> None:
    try:
        elems = driver.find_elements(By.CSS_SELECTOR, classes_to_css(LEFT_ARROW_CLASS))
        for elem in elems:
            if elem.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                time.sleep(SCROLL_SLEEP)
                driver.execute_script("arguments[0].click();", elem)
                time.sleep(POST_CLICK_SLEEP)
                wait_for_attendance_table(driver)
                return
    except Exception:
        pass

    fallback_xpaths = [
        "//button[.//*[name()='svg']]",
        "//*[@role='button'][.//*[name()='svg']]",
        "//*[name()='svg']/ancestor::*[@role='button' or self::button][1]",
    ]

    for xp in fallback_xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xp)
            for elem in elems:
                try:
                    if not elem.is_displayed():
                        continue
                    box = elem.rect
                    if box and box.get("x", 9999) < 900 and box.get("y", 9999) < 500:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                        time.sleep(SCROLL_SLEEP)
                        driver.execute_script("arguments[0].click();", elem)
                        time.sleep(POST_CLICK_SLEEP)
                        wait_for_attendance_table(driver)
                        return
                except StaleElementReferenceException:
                    continue
        except Exception:
            continue

    raise RuntimeError("Could not click the left attendance navigation arrow.")


def click_left_arrow_block(driver: webdriver.Chrome, clicks: int = DATE_BLOCK_CLICKS) -> None:
    old_labels = get_visible_date_labels(driver)
    log(f"Current labels before block move: {old_labels}")

    for i in range(clicks):
        log(f"Clicking left arrow ({i + 1}/{clicks})")
        click_left_arrow(driver)
        time.sleep(0.8)

    for _ in range(20):
        new_labels = get_visible_date_labels(driver)
        if new_labels and new_labels != old_labels:
            log(f"Date block changed to: {new_labels}")
            return
        time.sleep(0.7)

    save_debug_page(driver, "left_arrow_block_failed")
    raise RuntimeError("Left-arrow block click did not change the visible date headers.")


# ============================================================
# ROW / CELL EXTRACTION
# ============================================================

def extract_name_from_row(row: WebElement) -> str:
    try:
        elems = row.find_elements(By.CSS_SELECTOR, classes_to_css(NAME_CELL_CLASS))
        for elem in elems:
            txt = safe_text(elem)
            if txt:
                return txt
    except Exception:
        pass

    try:
        tds = row.find_elements(By.XPATH, TD_REL_XPATH)
        if tds:
            txt = safe_text(tds[0])
            if txt:
                return txt
    except Exception:
        pass

    return ""


def cell_is_present(cell: WebElement) -> bool:
    """
    Checked icons appear to include a filled square/checkmark structure.
    Unchecked icons appear outline-only.
    This is intentionally stricter than the old 'any path means present' logic.
    """
    try:
        html = (cell.get_attribute("innerHTML") or "").lower()
    except Exception:
        html = ""

    if "<svg" not in html:
        return False

    try:
        paths = cell.find_elements(By.XPATH, ".//*[name()='svg']//*[name()='path']")
    except Exception:
        paths = []

    try:
        circles = cell.find_elements(By.XPATH, ".//*[name()='svg']//*[name()='circle']")
    except Exception:
        circles = []

    try:
        rects = cell.find_elements(By.XPATH, ".//*[name()='svg']//*[name()='rect']")
    except Exception:
        rects = []

    # Strong indicator of checked icon complexity
    if len(paths) >= 2:
        return True

    # Filled square plus path often indicates checked state
    if len(rects) >= 1 and len(paths) >= 1:
        return True

    # Heuristic from your screenshot: checked icon markup included fill-rule/clip-rule/currentColor
    if "fill-rule" in html and "clip-rule" in html and "currentcolor" in html:
        return True

    # Simple outline-only icon should not count as present
    if len(paths) <= 1 and len(circles) <= 1 and len(rects) == 0:
        return False

    return False


def scrape_current_member_page(
    driver: webdriver.Chrome,
    attendance_data: Dict[str, Dict[date, bool]],
    start_dt: date,
    end_dt: date,
) -> Tuple[int, List[date]]:
    wait_for_attendance_table(driver)

    columns = get_visible_columns(driver, start_dt, end_dt)
    if not columns:
        save_debug_page(driver, "no_columns_detected")
        raise RuntimeError("Could not detect visible attendance date columns.")

    target_columns = [c for c in columns if start_dt <= c.actual_date <= end_dt]
    found_dates = sorted({c.actual_date for c in target_columns})

    rows = driver.find_elements(By.XPATH, ROW_XPATH)
    scraped_rows = 0

    for row in rows:
        try:
            name = extract_name_from_row(row)
            if not name:
                continue

            scraped_rows += 1
            attendance_data.setdefault(name, {})

            cells = row.find_elements(By.XPATH, TD_REL_XPATH)
            for col in target_columns:
                if col.column_index >= len(cells):
                    continue
                attendance_data[name][col.actual_date] = cell_is_present(cells[col.column_index])

        except StaleElementReferenceException:
            continue

    return scraped_rows, found_dates


def scrape_all_member_pages_for_current_dates(
    driver: webdriver.Chrome,
    attendance_data: Dict[str, Dict[date, bool]],
    start_dt: date,
    end_dt: date,
) -> List[date]:
    collected: Set[date] = set()

    member_pages = get_bottom_page_numbers(driver)
    log(f"Member page numbers found: {member_pages}")

    # Always scrape currently visible page first
    rows_scraped, dates_found = scrape_current_member_page(driver, attendance_data, start_dt, end_dt)
    for dt in dates_found:
        collected.add(dt)
    log(f"Scraped current visible page; rows: {rows_scraped}; dates: {[d.isoformat() for d in dates_found]}")

    # Then explicitly visit remaining pages, skipping 1 because current page is already scraped
    for member_page in member_pages:
        if member_page == 1:
            continue
        log(f"Scraping member page {member_page}")
        click_member_page_number(driver, member_page)
        rows_scraped, dates_found = scrape_current_member_page(driver, attendance_data, start_dt, end_dt)
        for dt in dates_found:
            collected.add(dt)
        log(f"Rows scraped: {rows_scraped}; dates: {[d.isoformat() for d in dates_found]}")

    return sorted(collected)


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
            raise RuntimeError("Safety stop reached while paging through date blocks.")

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

        dates_found = scrape_all_member_pages_for_current_dates(driver, attendance_data, start_dt, end_dt)
        for dt in dates_found:
            collected_dates.add(dt)

        earliest_visible = min(visible_window_dates)
        if earliest_visible <= start_dt:
            log("Reached earliest required date block.")
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
