import json
import os
import time

import requests

URL = "https://www.acnc.gov.au/api/dynamics/search/charity"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}
OUT_FILE = "charity_list.jsonl"
DELAY = 1


def get_page(page):
    """Fetch one page of 100 results; an empty result counts as a failure."""
    wait = DELAY
    for _ in range(6):
        resp = requests.get(
            URL,
            params={"search": " ", "items_per_page": 100, "page": page},
            headers=HEADERS,
            timeout=30,
        )
        if resp.status_code == 200:
            payload = resp.json()
            if payload.get("results"):
                return payload
            print(f"  Empty results on page {page}; retrying in {wait}s")
        else:
            print(f"  HTTP {resp.status_code} on page {page}; retrying in {wait}s")
        time.sleep(wait)
        wait = min(wait * 2, 60)
    raise RuntimeError(f"Page {page} failed after retries")


def saved_pages():
    """Return the set of page numbers already written to OUT_FILE."""
    pages = set()
    if os.path.exists(OUT_FILE):
        with open(OUT_FILE, encoding="utf-8") as f:
            for line in f:
                pages.add(json.loads(line)["page"])
    return pages


def main():
    first = get_page(0)
    total_pages = first["pager"]["total_pages"]
    done = saved_pages()
    missing = [p for p in range(total_pages) if p not in done]
    print(f"{len(missing)} pages to fetch (of {total_pages})")
    for n, page in enumerate(missing, start=1):
        payload = get_page(page)
        with open(OUT_FILE, "a", encoding="utf-8") as f:
            for item in payload["results"]:
                row = {"page": page, "uuid": item["uuid"], **item["data"]}
                f.write(json.dumps(row, ensure_ascii=False) + "\n")
        if n % 10 == 0:
            print(f"{n}/{len(missing)} pages fetched")
        time.sleep(DELAY)
    print("Done.")


if __name__ == "__main__":
    main()