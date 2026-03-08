# Attendance Scraper

This repo contains a Selenium-based headless scraper for the LCR Class and Quorum Attendance page.

## Required GitHub Secrets

Add these in:

Settings > Secrets and variables > Actions

- `LCR_USERNAME`
- `LCR_PASSWORD`

## Run the workflow

Go to:

Actions > Run Attendance Scraper > Run workflow

Enter:

- `start_date` in `YYYY-MM-DD`
- `end_date` in `YYYY-MM-DD`

The workflow will:

1. Log in to LCR
2. Scrape attendance across all visible member pages
3. Move left through attendance weeks until the start date is covered
4. Generate an Excel file in `data/`
5. Upload the file as an artifact
6. Attempt to commit the spreadsheet back into the repo

## Output

The spreadsheet includes:

- Name
- % activity
- one column per attendance date
- ☑ for present
- ☐ for not present
