# ACNC Charity Scraper

Python scraper that collects public data from Australian charities listed on the [ACNC register](https://www.acnc.gov.au) and exports it to a formatted Excel file.

**Collected fields:** Name, ABN, Address for Service Email, Charity Size.

> This repository contains a **sample run (~100 charities)** to demonstrate the approach. The same code scales to the full register (~55,000 charities), see "Scaling" below.

## Features

- **Resume after failure:** progress is saved after every record in `progress.json`, so an interrupted run continues where it stopped.
- **Adaptive delay:** the pause between requests doubles (up to 30s) when the server answers 429/403, retries up to 3 times, and returns to normal when requests succeed.
- **Batch processing:** records are processed in batches (`BATCH_SIZE`) and the Excel file is saved after each batch, so partial results are always available.
- **Clean output:** styled header, auto-sized columns, frozen header row.

## How it works

1. Queries the ACNC internal JSON search API (paginated) and keeps only charities with status "Registered".
2. Fetches each charity's detail record by UUID to get the Address for Service Email.
3. Skips records without an ABN (e.g. internal group records).
4. Exports the results to `acnc_charities_sample.xlsx`.
5. Charity size is shown as "Not reported" when the ACNC has not classified the charity.

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

Change `SEARCH_TERM`, `SAMPLE_SIZE`, `BATCH_SIZE` and `DELAY_SECONDS` at the top of `scraper.py` to adjust the run. To start a fresh collection, delete `progress.json`.

## Scaling to the full dataset

Already implemented: resume, adaptive delay and batch saving. Planned for the full register (~55,000 charities):

- Paginate the search with 100 items per page and no search term.
- Log skipped and failed records to a separate file.

## Notes

- Only publicly available data is collected, via the same endpoints the website itself uses.
- Please respect the website's terms and robots.txt, and keep a polite delay between requests.

## Author

Gabriel Fantin - [GitHub](https://github.com/gabrielbf7)