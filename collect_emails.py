"""Collect Address for Service Email for all registered ACNC charities.

Reads charity_list.jsonl, fetches each detail record and appends the result
to emails.jsonl (one line per charity), so an interrupted run can resume.
Reuses the HTTP, retry and adaptive-delay logic from scraper.py.
Usage: python collect_emails.py [limit]
"""

import json
import os
import sys

import requests

import scraper

LIST_FILE = "charity_list.jsonl"
EMAILS_FILE = "emails.jsonl"
OUTPUT_FILE = "acnc_registered_charities.xlsx"
EXCEL_EVERY = 1000


def read_jsonl(path):
    """Return a list of dicts from a JSON Lines file (empty if missing)."""
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return [json.loads(line) for line in f if line.strip()]


def save_excel(charities, done):
    """Export every charity already collected; warn if the file is open."""
    rows = [
        {
            "name": c["Name"],
            "abn": c["Abn"],
            "email": done[c["uuid"]]["email"],
            "size": c.get("CharitySize") or "Not reported",
        }
        for c in charities
        if c["uuid"] in done
    ]
    try:
        scraper.export_to_excel(rows, OUTPUT_FILE)
        print(f"  Saved {len(rows)} rows to {OUTPUT_FILE}")
    except PermissionError:
        print(f"  Could not save {OUTPUT_FILE}. Close it in Excel.")


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else None

    charities = [c for c in read_jsonl(LIST_FILE) if c.get("Status") == "Registered"]
    done = {r["uuid"]: r for r in read_jsonl(EMAILS_FILE)}
    todo = [c for c in charities if c["uuid"] not in done]
    if limit:
        todo = todo[:limit]
    print(f"Registered: {len(charities)} | done: {len(done)} | this run: {len(todo)}")

    with open(EMAILS_FILE, "a", encoding="utf-8") as out:
        for i, charity in enumerate(todo, start=1):
            label = f"[{i}/{len(todo)}]"
            try:
                detail = scraper.fetch_detail(charity["uuid"])
            except requests.RequestException as exc:
                print(f"{label} ERROR - {charity['Name']}: {exc}")
                continue  # not saved, so the next run retries it

            record = {
                "uuid": charity["uuid"],
                "email": detail.get("AFSEmail"),
                "contact_email": detail.get("Email"),
            }
            out.write(json.dumps(record, ensure_ascii=False) + "\n")
            out.flush()
            done[charity["uuid"]] = record
            print(f"{label} OK - {charity['Name']}")

            if i % EXCEL_EVERY == 0:
                save_excel(charities, done)
            scraper.throttle.wait()

    save_excel(charities, done)
    print("Done.")


if __name__ == "__main__":
    main()