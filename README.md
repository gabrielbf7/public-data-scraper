# ACNC Charity Scraper

Python scraper that collects public data from Australian charities listed on the [ACNC register](https://www.acnc.gov.au) and exports it to a formatted Excel file.

**Collected fields:** Name, ABN, Address for Service Email, Charity Size.

> This repository contains a **sample run (~30 charities)** to demonstrate the approach. The same code scales to the full register (~55,000 charities), see "Scaling" below.

## How it works

1. Queries the ACNC internal JSON search API (paginated) and keeps only charities with status "Registered".
2. Fetches each charity's detail record by UUID to get the Address for Service Email.
3. Skips records without an ABN (e.g. internal group records).
4. Exports the results to `acnc_charities_sample.xlsx` with a styled header, auto-sized columns and a frozen header row.
5. Charity size is shown as "Not reported" when the ACNC has not classified the charity.

Requests are sent with a browser User-Agent and a configurable pause (`DELAY_SECONDS`) between calls, to keep the load on the server low.

## Tech Stack

- Python 3
- requests (HTTP/JSON API calls)
- openpyxl (Excel export)

## Setup

    python -m venv venv
    venv\Scripts\activate
    pip install requests openpyxl

## Usage

    python scraper.py

Change `SEARCH_TERM`, `SAMPLE_SIZE` and `DELAY_SECONDS` at the top of `scraper.py` to adjust the run.

## Scaling to the full dataset

For the full register (~55,000 charities) the plan is:

- Paginate the search with 100 items per page and no search term.
- Save progress in batches and resume from where it stopped after an interruption.
- Adaptive delay: slow down automatically if the server returns 429/403 errors.
- Log skipped and failed records to a separate file.

## Notes

- Only publicly available data is collected, via the same endpoints the website itself uses.
- Please respect the website's terms and robots.txt, and keep a polite delay between requests.

## Author

Gabriel Fantin - [GitHub](https://github.com/gabrielbf7)