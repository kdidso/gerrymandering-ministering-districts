from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import Browser, BrowserContext, Locator, Page, sync_playwright


# ============================================================
# CONFIG
# ============================================================

ATTENDANCE_URL = "https://lcr.churchofjesuschrist.org/report/class-and-quorum-attendance/overview?lang=eng"

# Override these with GitHub Actions env vars or local environment variables if desired.
START_DATE = os.getenv("START_DATE", "2024-12-28")
END_DATE = os.getenv("END_DATE", "2025-03-08")

OUTPUT_DIR = "data"
HEADLESS = True
SLOW_MO_MS = 0
DEFAULT_TIMEOUT_MS = 20000
POST_CLICK_WAIT_MS = 1200
PAGE_LOAD_WAIT_MS = 1800

# Login selectors
EMAIL_INPUT_SELECTORS = [
    "input[type='email']",
    "input[name='identifier']",
    "#identifierId",
]

PASSWORD_INPUT_SELECTORS = [
    "input[type='password']",
    "input[name='password']",
    "input[name='Passwd']",
]

# Attendance page selectors based on your inspection
NAME_CELL_SELECTOR = ".sc-14fff288-0.llFqzd.sc-24d75410-4.bMziSa.sc-24d75410-0.lpsZiy"
PAGE_NUMBER_SELECTOR = ".sc-66e0b3ee-0.hxtefx"
DATE_HEADER_SELECTOR = ".sc-arbpvo-0.sc-arbpvo-1.fRnemr.fhFlsT.sc-473b494-0.kpqFIx"
LEFT_ARROW_SELECTOR = ".sc-2b11ed23-0.kPPSzB"

# Fallback selectors
TABLE_SELECTOR = "table"
TABLE_ROW_SELECTOR = "table tbody tr"
TABLE_HEADER_CELL_SELECTOR = "table thead th"
TABLE_BODY_CELL_SELECTOR = "td"

# If true, store trace/screenshot-like debug HTML dumps on failure
SAVE_DEBUG_HTML = True


# ============================================================
# DATA CLASSES
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
    print(f"[ERROR] {msg}")


def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def safe_text(locator: Locator) -> str:
    try:
        return normalize_space(locator.text_content() or "")
    except Exception:
        return ""


def dedupe_preserve_order(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def first_visible_locator(page: Page, selectors: List[str], timeout_ms: int = 2500) -> Optional[Locator]:
    for selector in selectors:
        try:
            loc = page.locator(selector)
            if loc.count() > 0:
                loc.first.wait_for(state="visible", timeout=timeout_ms)
                return loc.first
        except Exception:
            continue
    return None


def try_click_by_text(page: Page, texts: List[str]) -> bool:
    for text in texts:
        patterns = [
            page.get_by_role("button", name=re.compile(fr"^{re.escape(text)}$", re.I)),
            page.get_by_role("link", name=re.compile(fr"^{re.escape(text)}$", re.I)),
            page.get_by_text(re.compile(fr"^{re.escape(text)}$", re.I)),
        ]
        for loc in patterns:
            try:
                if loc.count() > 0:
                    loc.first.click()
                    page.wait_for_timeout(800)
                    return True
            except Exception:
                continue
    return False


def is_probable_date_label(text: str) -> bool:
    return bool(re.match(r"^\d{1,2}\s+[A-Za-z]{3}$", normalize_space(text)))


def midpoint_date(start_dt: date, end_dt: date) -> date:
    return start_dt + timedelta(days=(end_dt - start_dt).days // 2)


def infer_year_for_label(label: str, start_dt: date, end_dt: date) -> Optional[date]:
    label = normalize_space(label)
    if not is_probable_date_label(label):
        return None

    mid = midpoint_date(start_dt, end_dt)
    best_candidate: Optional[date] = None
    best_distance: Optional[int] = None

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


def format_excel_header(dt: date) -> str:
    return f"{dt.strftime('%b')} {dt.day} {dt.year}"


def save_debug_html(page: Page, name: str) -> None:
    if not SAVE_DEBUG_HTML:
        return
    try:
        debug_dir = ensure_dir("debug")
        out_path = debug_dir / f"{name}.html"
        out_path.write_text(page.content(), encoding="utf-8")
        log(f"Saved debug HTML to {out_path}")
    except Exception as ex:
        warn(f"Could not save debug HTML: {ex}")


# ============================================================
# LOGIN
# ============================================================

def page_looks_logged_in(page: Page) -> bool:
    url = page.url.lower()
    if "class-and-quorum-attendance" in url:
        return True

    try:
        if page.locator(NAME_CELL_SELECTOR).count() > 0:
            return True
    except Exception:
        pass

    try:
        if page.get_by_text("Class and Quorum Attendance", exact=False).count() > 0:
            return True
    except Exception:
        pass

    return False


def automated_login(page: Page, username: str, password: str) -> bool:
    log("Opening attendance page")
    page.goto(ATTENDANCE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(PAGE_LOAD_WAIT_MS)

    if page_looks_logged_in(page):
        log("Already logged in")
        return True

    email_input = first_visible_locator(page, EMAIL_INPUT_SELECTORS, timeout_ms=6000)
    if email_input is not None:
        log("Filling username")
        email_input.fill(username)
        page.wait_for_timeout(300)
        try_click_by_text(page, ["Next", "Continue", "Sign In", "Continue to Church Account"])
        page.wait_for_timeout(1800)

    password_input = first_visible_locator(page, PASSWORD_INPUT_SELECTORS, timeout_ms=8000)
    if password_input is not None:
        log("Filling password")
        password_input.fill(password)
        page.wait_for_timeout(300)
        try_click_by_text(page, ["Sign In", "Log In", "Next", "Continue"])
        page.wait_for_timeout(2500)

    for _ in range(12):
        if page_looks_logged_in(page):
            return True
        page.wait_for_timeout(1000)

    return False


def ensure_logged_in(page: Page, username: str, password: str) -> None:
    if not automated_login(page, username, password):
        save_debug_html(page, "login_failure")
        raise RuntimeError(
            "Automated login failed in headless mode. "
            "This may mean the sign-in flow changed or requires an extra prompt."
        )

    page.goto(ATTENDANCE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(PAGE_LOAD_WAIT_MS)

    if not page_looks_logged_in(page):
        save_debug_html(page, "attendance_page_not_loaded")
        raise RuntimeError("Login seemed to complete, but attendance page did not load correctly.")


# ============================================================
# PAGE / TABLE EXTRACTION
# ============================================================

def wait_for_attendance_table(page: Page) -> None:
    try:
        page.locator(NAME_CELL_SELECTOR).first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
        page.wait_for_timeout(1000)
        return
    except Exception:
        pass

    try:
        page.locator(TABLE_SELECTOR).first.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
        page.wait_for_timeout(1000)
        return
    except Exception:
        pass

    save_debug_html(page, "attendance_table_missing")
    raise RuntimeError("Could not find attendance table or visible member rows.")


def get_visible_date_labels(page: Page) -> List[str]:
    labels: List[str] = []

    try:
        loc = page.locator(DATE_HEADER_SELECTOR)
        for i in range(loc.count()):
            txt = safe_text(loc.nth(i))
            if is_probable_date_label(txt):
                labels.append(txt)
    except Exception:
        pass

    if labels:
        return dedupe_preserve_order(labels)

    try:
        headers = page.locator(TABLE_HEADER_CELL_SELECTOR)
        for i in range(headers.count()):
            txt = safe_text(headers.nth(i))
            if is_probable_date_label(txt):
                labels.append(txt)
    except Exception:
        pass

    return dedupe_preserve_order(labels)


def get_visible_columns(page: Page, start_dt: date, end_dt: date) -> List[ColumnInfo]:
    columns: List[ColumnInfo] = []

    try:
        headers = page.locator(TABLE_HEADER_CELL_SELECTOR)
        for i in range(headers.count()):
            txt = safe_text(headers.nth(i))
            if not is_probable_date_label(txt):
                continue

            actual = infer_year_for_label(txt, start_dt, end_dt)
            if actual is None:
                continue

            columns.append(ColumnInfo(column_index=i, label_text=txt, actual_date=actual))
    except Exception:
        pass

    if columns:
        return columns

    # fallback if semantic th structure is missing
    labels = get_visible_date_labels(page)
    body_col_idx = 2  # Name and Gender assumed first two columns
    for lbl in labels:
        actual = infer_year_for_label(lbl, start_dt, end_dt)
        if actual is not None:
            columns.append(ColumnInfo(column_index=body_col_idx, label_text=lbl, actual_date=actual))
            body_col_idx += 1

    return columns


def get_visible_window_dates(page: Page, start_dt: date, end_dt: date) -> List[date]:
    dates: List[date] = []
    for label in get_visible_date_labels(page):
        actual = infer_year_for_label(label, start_dt, end_dt)
        if actual is not None:
            dates.append(actual)
    return sorted(set(dates))


def get_bottom_page_numbers(page: Page) -> List[int]:
    page_numbers: List[int] = []

    try:
        loc = page.locator(PAGE_NUMBER_SELECTOR)
        for i in range(loc.count()):
            txt = safe_text(loc.nth(i))
            if txt.isdigit():
                page_numbers.append(int(txt))
    except Exception:
        pass

    if page_numbers:
        return sorted(set(page_numbers))

    try:
        generic = page.locator("text=/^\\d+$/")
        for i in range(generic.count()):
            txt = safe_text(generic.nth(i))
            if txt.isdigit():
                page_numbers.append(int(txt))
    except Exception:
        pass

    page_numbers = sorted(set(page_numbers))
    return page_numbers if page_numbers else [1]


def click_bottom_page_number(page: Page, target_page_num: int) -> None:
    target = str(target_page_num)

    try:
        loc = page.locator(PAGE_NUMBER_SELECTOR)
        for i in range(loc.count()):
            item = loc.nth(i)
            if safe_text(item) == target:
                item.click()
                page.wait_for_timeout(POST_CLICK_WAIT_MS)
                return
    except Exception:
        pass

    try:
        page.get_by_text(re.compile(fr"^{re.escape(target)}$")).first.click()
        page.wait_for_timeout(POST_CLICK_WAIT_MS)
        return
    except Exception:
        pass

    raise RuntimeError(f"Could not click member page number {target_page_num}.")


def click_left_arrow(page: Page) -> None:
    try:
        loc = page.locator(LEFT_ARROW_SELECTOR)
        if loc.count() > 0:
            loc.first.click()
            page.wait_for_timeout(POST_CLICK_WAIT_MS)
            return
    except Exception:
        pass

    # conservative fallback: leftmost visible svg button near top
    candidates = [
        "button:has(svg)",
        "[role='button']:has(svg)",
        "svg",
    ]
    for selector in candidates:
        try:
            loc = page.locator(selector)
            for i in range(loc.count()):
                item = loc.nth(i)
                box = item.bounding_box()
                if box and box["x"] < 700 and box["y"] < 350:
                    item.click()
                    page.wait_for_timeout(POST_CLICK_WAIT_MS)
                    return
        except Exception:
            continue

    raise RuntimeError("Could not click the left attendance navigation arrow.")


def extract_name_from_row(row: Locator) -> str:
    try:
        loc = row.locator(NAME_CELL_SELECTOR)
        if loc.count() > 0:
            name = safe_text(loc.first)
            if name:
                return name
    except Exception:
        pass

    try:
        cells = row.locator(TABLE_BODY_CELL_SELECTOR)
        if cells.count() > 0:
            txt = safe_text(cells.nth(0))
            if txt and "," in txt:
                return txt
    except Exception:
        pass

    return ""


def cell_is_present(cell: Locator) -> bool:
    """
    Based on your inspection, filled attendance includes a check-mark path.
    Empty circles appear as outline only.
    """
    try:
        html = (cell.inner_html() or "").lower()
    except Exception:
        html = ""

    if "<path" in html:
        return True

    try:
        if cell.locator("svg path").count() > 0:
            return True
    except Exception:
        pass

    return False


def scrape_current_member_page(
    page: Page,
    attendance_data: Dict[str, Dict[date, bool]],
    start_dt: date,
    end_dt: date,
) -> Tuple[int, List[date]]:
    wait_for_attendance_table(page)

    columns = get_visible_columns(page, start_dt, end_dt)
    if not columns:
        save_debug_html(page, "no_columns_detected")
        raise RuntimeError("No attendance date columns were detected.")

    target_columns = [c for c in columns if start_dt <= c.actual_date <= end_dt]
    rows = page.locator(TABLE_ROW_SELECTOR)

    found_dates = sorted({c.actual_date for c in target_columns})
    scraped_rows = 0

    for row_idx in range(rows.count()):
        row = rows.nth(row_idx)
        name = extract_name_from_row(row)
        if not name:
            continue

        scraped_rows += 1
        attendance_data.setdefault(name, {})

        cells = row.locator(TABLE_BODY_CELL_SELECTOR)
        cell_count = cells.count()

        for col in target_columns:
            if col.column_index >= cell_count:
                continue
            present = cell_is_present(cells.nth(col.column_index))
            attendance_data[name][col.actual_date] = present

    return scraped_rows, found_dates


# ============================================================
# SCRAPE LOOP
# ============================================================

def scrape_attendance(page: Page, start_dt: date, end_dt: date) -> Tuple[Dict[str, Dict[date, bool]], List[date]]:
    attendance_data: Dict[str, Dict[date, bool]] = {}
    collected_dates: set[date] = set()

    page.goto(ATTENDANCE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(PAGE_LOAD_WAIT_MS)
    wait_for_attendance_table(page)

    safety_counter = 0

    while True:
        safety_counter += 1
        if safety_counter > 60:
            raise RuntimeError("Safety stop reached while paging through week windows.")

        visible_window_dates = get_visible_window_dates(page, start_dt, end_dt)
        if not visible_window_dates:
            save_debug_html(page, f"window_dates_missing_{safety_counter}")
            raise RuntimeError("Could not detect visible date headers.")

        log(f"Visible week window: {[d.isoformat() for d in visible_window_dates]}")

        member_pages = get_bottom_page_numbers(page)
        log(f"Member page numbers found: {member_pages}")

        for member_page in member_pages:
            log(f"Scraping member page {member_page}")
            click_bottom_page_number(page, member_page)
            rows_scraped, dates_found = scrape_current_member_page(page, attendance_data, start_dt, end_dt)
            for dt in dates_found:
                collected_dates.add(dt)
            log(f"Rows scraped: {rows_scraped}; in-range dates found: {[d.isoformat() for d in dates_found]}")

        earliest_visible = min(visible_window_dates)
        if earliest_visible <= start_dt:
            log("Reached earliest required date window.")
            break

        log("Clicking left arrow to move to older dates")
        click_left_arrow(page)

    final_dates = sorted(dt for dt in collected_dates if start_dt <= dt <= end_dt)
    if not final_dates:
        raise RuntimeError("No attendance dates were collected within the requested date range.")

    return attendance_data, final_dates


# ============================================================
# OUTPUT
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
    username = os.getenv("LCR_USERNAME", "").strip()
    password = os.getenv("LCR_PASSWORD", "").strip()

    if not username or not password:
        err("Missing LCR_USERNAME and/or LCR_PASSWORD environment variables.")
        return 1

    start_dt = parse_iso_date(START_DATE)
    end_dt = parse_iso_date(END_DATE)

    if start_dt > end_dt:
        err("START_DATE must be on or before END_DATE.")
        return 1

    output_dir = ensure_dir(OUTPUT_DIR)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = output_dir / f"attendance_{start_dt.isoformat()}_to_{end_dt.isoformat()}_{timestamp}.xlsx"

    with sync_playwright() as p:
        browser: Browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO_MS)
        context: BrowserContext = browser.new_context()
        context.set_default_timeout(DEFAULT_TIMEOUT_MS)

        try:
            page = context.new_page()
            ensure_logged_in(page, username, password)
            attendance_data, all_dates = scrape_attendance(page, start_dt, end_dt)
            write_excel(attendance_data, all_dates, excel_path)
            log(f"Excel output written to {excel_path}")
        finally:
            context.close()
            browser.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
