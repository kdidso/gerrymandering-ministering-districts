# Gerrymandering Ministering Districts

This repo currently includes a headless Playwright script that logs into LCR, visits the Class and Quorum Attendance page, scrapes attendance by date and member, and outputs an Excel spreadsheet into the `data` folder.

## Required GitHub Secrets

In the repo settings, create these secrets:

- `LCR_USERNAME`
- `LCR_PASSWORD`

## Run it manually from GitHub

Go to:

- **Actions**
- **Run Attendance Scraper**
- **Run workflow**

Enter:

- `start_date` in `YYYY-MM-DD`
- `end_date` in `YYYY-MM-DD`

The workflow will:

1. Run the scraper headlessly
2. Save the `.xlsx` file to `data/`
3. Upload it as an artifact
4. Commit it back into the repo

## Local run

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
export LCR_USERNAME="your_username"
export LCR_PASSWORD="your_password"
export START_DATE="2024-12-28"
export END_DATE="2025-03-08"
python attendance_scraper.py
