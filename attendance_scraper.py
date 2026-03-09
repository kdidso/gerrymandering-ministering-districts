from __future__ import annotations

import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Set, Tuple

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ============================================================
# CONFIG
# ============================================================

LCR_BASE = "https://lcr.churchofjesuschrist.org"
ATTENDANCE_PAGE_URL = f"{LCR_BASE}/report/class-and-quorum-attendance/overview?lang=eng"
MEMBER_LIST_PAGE_URL = f"{LCR_BASE}/records/member-list?lang=eng"

UNIT_NUMBER = os.getenv("UNIT_NUMBER", "253022").strip()

USERNAME = os.getenv("LCR_USERNAME", "").strip()
PASSWORD = os.getenv("LCR_PASSWORD", "").strip()

START_DATE = os.getenv("START_DATE", "2025-12-28").strip()
END_DATE = os.getenv("END_DATE", "2026-03-08").strip()

OUTPUT_DIR = "data"

HEADLESS = True
DEFAULT_WAIT = 30
LONG_WAIT = 60
WINDOW_DAYS = 28


# ============================================================
# HELPERS
# ============================================================

def log(msg: str) -> None:
    print(f"[INFO] {msg}")


def err(msg: str) -> None:
    print(f"[ERROR] {msg}", file=sys.stderr)


def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def to_iso(dt: date) -> str:
    return dt.strftime("%Y-%m-%d")


def format_excel_header(dt: date) -> str:
    return f"{dt.strftime('%b')} {dt.day} {dt.year}"


def sunday_on_or_before(dt: date) -> date:
    return dt - timedelta(days=(dt.weekday() + 1) % 7)


def build_date_windows(start_dt: date, end_dt: date) -> List[Tuple[date, date]]:
    """
    Build overlapping 28-day windows starting every Sunday.
    This covers the discovered API pattern without guessing exact block alignment.
    """
    windows: List[Tuple[date, date]] = []

    first_start = sunday_on_or_before(start_dt - timedelta(days=28))
    last_start = sunday_on_or_before(end_dt)

    current = first_start
    while current <= last_start:
        windows.append((current, current + timedelta(days=28)))
        current += timedelta(days=7)

    return windows


# ============================================================
# SELENIUM LOGIN
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
        user_input = WebDriverWait(driver, LONG_WAIT).until(
            EC.presence_of_element_located((By.ID, "username-input"))
        )
        user_input.clear()
        user_input.send_keys(USERNAME)
        user_input.send_keys(Keys.ENTER)

        pwd_input = WebDriverWait(driver, LONG_WAIT).until(
            EC.presence_of_element_located((By.ID, "password-input"))
        )
        pwd_input.clear()
        pwd_input.send_keys(PASSWORD)
        pwd_input.send_keys(Keys.ENTER)

        WebDriverWait(driver, LONG_WAIT).until(EC.url_contains(LCR_BASE))
        log("Login submitted successfully")
    except Exception as ex:
        raise RuntimeError("Automated login failed with the known username/password field flow.") from ex


def build_requests_session_from_driver(driver: webdriver.Chrome) -> requests.Session:
    session = requests.Session()

    for cookie in driver.get_cookies():
        session.cookies.set(
            cookie["name"],
            cookie["value"],
            domain=cookie.get("domain"),
            path=cookie.get("path", "/"),
        )

    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json, text/plain, */*",
            "Referer": ATTENDANCE_PAGE_URL,
        }
    )
    return session


# ============================================================
# API URLS
# ============================================================

def attendance_api_url(unit_number: str, start_dt: date, end_dt: date) -> str:
    return (
        f"{LCR_BASE}/api/umlu/v1/class-and-quorum/attendance/overview/"
        f"unitNumber/{unit_number}/start/{to_iso(start_dt)}/end/{to_iso(end_dt)}?lang=eng"
    )


def member_list_api_url(unit_number: str) -> str:
    return f"{LCR_BASE}/api/umlu/v1/member-list?lang=eng&unitNumber={unit_number}"


# ============================================================
# API FETCHES
# ============================================================

def fetch_json(session: requests.Session, url: str) -> dict | list:
    response = session.get(url, timeout=60)
    response.raise_for_status()
    return response.json()


def fetch_member_roster(session: requests.Session, unit_number: str) -> Dict[str, str]:
    url = member_list_api_url(unit_number)
    log(f"Fetching member roster: {url}")
    payload = fetch_json(session, url)

    if not isinstance(payload, list):
        raise RuntimeError("Member-list API did not return a list.")

    roster: Dict[str, str] = {}

    for person in payload:
        uuid = (person.get("uuid") or person.get("personUuid") or "").strip()
        if not uuid:
            continue

        name = (
            person.get("nameListPreferredLocal")
            or person.get("householdNameDirectoryLocal")
            or (person.get("nameFormats") or {}).get("listPreferredLocal")
            or person.get("houseHoldMemberNameForList")
            or ""
        )
        name = name.strip()

        if name:
            roster[uuid] = name

    log(f"Roster members mapped: {len(roster)}")
    if not roster:
        raise RuntimeError("No UUID→name mappings were found in member-list response.")

    return roster


def fetch_attendance_window(session: requests.Session, unit_number: str, start_dt: date, end_dt: date) -> dict:
    url = attendance_api_url(unit_number, start_dt, end_dt)
    log(f"Fetching attendance window: {to_iso(start_dt)} to {to_iso(end_dt)}")
    payload = fetch_json(session, url)

    if not isinstance(payload, dict):
        raise RuntimeError("Attendance API did not return an object.")

    return payload


# ============================================================
# MERGE ATTENDANCE
# ============================================================

def merge_attendance_window(
    payload: dict,
    roster: Dict[str, str],
    attendance_data: Dict[str, Dict[date, bool]],
    all_dates: Set[date],
    start_dt: date,
    end_dt: date,
) -> int:
    attendance_data_obj = payload.get("attendanceData") or {}
    attendees = attendance_data_obj.get("attendees") or []

    merged_people = 0

    for person in attendees:
        uuid = (person.get("uuid") or person.get("personUuid") or "").strip()
        if not uuid:
            continue

        name = roster.get(uuid)
        if not name:
            continue

        merged_people += 1
        attendance_data.setdefault(name, {})

        for entry in person.get("entries", []) or []:
            date_obj = entry.get("date") or {}
            iso = date_obj.get("isoYearMonthDay")
            if not iso:
                continue

            try:
                dt = datetime.strptime(iso, "%Y-%m-%d").date()
            except ValueError:
                continue

            if not (start_dt <= dt <= end_dt):
                continue

            all_dates.add(dt)
            attended = bool(
                entry.get("isMarkedAttended", entry.get("markedAttended", False))
            )
            attendance_data[name][dt] = attended

    return merged_people


def scrape_attendance_via_api(
    session: requests.Session,
    unit_number: str,
    start_dt: date,
    end_dt: date,
) -> Tuple[Dict[str, Dict[date, bool]], List[date]]:
    roster = fetch_member_roster(session, unit_number)

    attendance_data: Dict[str, Dict[date, bool]] = {}
    all_dates: Set[date] = set()

    windows = build_date_windows(start_dt, end_dt)
    log(f"Date windows to request: {[f'{to_iso(s)}..{to_iso(e)}' for s, e in windows]}")

    for win_start, win_end in windows:
        payload = fetch_attendance_window(session, unit_number, win_start, win_end)
        merged_people = merge_attendance_window(
            payload, roster, attendance_data, all_dates, start_dt, end_dt
        )
        log(f"Merged attendees from window: {merged_people}")

    final_dates = sorted(all_dates)
    if not final_dates:
        raise RuntimeError("No attendance dates were collected from the API.")

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

        # Touch both pages so auth/session context is definitely warm
        driver.get(ATTENDANCE_PAGE_URL)
        driver.get(MEMBER_LIST_PAGE_URL)

        session = build_requests_session_from_driver(driver)
        attendance_data, all_dates = scrape_attendance_via_api(session, UNIT_NUMBER, start_dt, end_dt)

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
