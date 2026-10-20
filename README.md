# ACNC Charity Scraper

Python scraper that collects public data from Australian charities listed on the [ACNC register](https://www.acnc.gov.au) and exports it to a formatted Excel file.

**Collected fields:** Name, ABN, Address for Service Email, Charity Size.

> This repository contains **sample outputs** (not the full dataset) to demonstrate the approach. The code handles the full register (66,907 registered charities), see "Performance" below.

## Features

- **Full register listing:** one paginated search lists all 95,638 charities (957 pages of 100) with ABN, status and size.
- **Resume after failure:** every step saves its progress (page by page, record by record), so an interrupted run continues where it stopped.
- **Adaptive delay:** the pause between requests doubles (up to 30s) when the server answers 429/403, retries up to 3 times, and returns to normal when requests succeed.
- **Cross-checked data:** the ABNs were compared against the official data.gov.au dataset (0 missing).
- **Clean output:** styled header, auto-sized columns, frozen header row.

## How it works

1. `collect_list.py` queries the ACNC internal JSON search API with a blank search term and writes `charity_list.jsonl` (UUID, name, ABN, status, size).
2. `collect_emails.py` reads that list, keeps only charities with status "Registered", and fetches each detail record by UUID to get the Address for Service Email (plus the general contact email as an extra). Results are appended to `emails.jsonl`.
3. The Excel file `acnc_registered_charities.xlsx` is rebuilt every 1,000 records and at the end of the run.
4. `scraper.py` is the original single-file version (search term + sample size), kept as a simple demo.

## Tech Stack

- Python 3
- requests (HTTP/JSON API calls)
- openpyxl (Excel export)

## Setup

    python -m venv venv
    venv\Scripts\activate
    pip install requests openpyxl

## Usage

    python collect_list.py
    python collect_emails.py 1000

The number after `collect_emails.py` limits how many charities are processed in the run. Omit it to process all of them. To resume after an interruption, run the same command again. To start over, delete `emails.jsonl`.

## Sample output

`acnc_sample_masked.xlsx` has 50 rows from a real run. Email local parts are masked (`***@domain`), because some are personal addresses. The complete files are not published.

## Performance

Measured on a 1,000-charity run:

- About 2.5 s per record (request plus a 1 s polite delay): ~42 minutes for 1,000 records.
- Estimated full run for 66,907 registered charities: ~46 hours sequential.
- 982 of 1,000 records (98.2%) had an Address for Service Email.
- Parallel requests could shorten the full run, but increase the risk of being blocked, so they are not enabled.

## Notes

- Group records (names ending in "_ACNC Group") have no ABN of their own, so ABN and email may be blank for them.
- Only publicly available data is collected, via the same endpoints the website itself uses.
- Please respect the website's terms and robots.txt, and keep a polite delay between requests.

## Author

Gabriel Fantin - [GitHub](https://github.com/gabrielbf7)