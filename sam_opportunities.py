"""
Fetch opportunities from the SAM.gov Opportunities API and save them to an
Excel spreadsheet.

All parameters are read from a .env file located next to this script:

    SAM_API_KEY   your SAM.gov API key
    INITIAL_DATE  (optional) postedFrom date, format MM/DD/YYYY
    FINAL_DATE    (optional) postedTo date,   format MM/DD/YYYY
    KEYWORD       title keyword to search for
    OUTPUT_FILE   (optional) output filename, defaults to GELSIGHT_OPPORTUNITIES.xlsx

If INITIAL_DATE / FINAL_DATE are not set, today's date (from datetime) is used
for both, so a daily scheduled run always searches the current day.

Usage:
    pip install requests python-dotenv pandas openpyxl
    python sam_opportunities.py
"""

import os
import sys
from datetime import date

import requests
import pandas as pd
from dotenv import load_dotenv

API_URL = "https://api.sam.gov/opportunities/v2/search"
PAGE_SIZE = 100  # max records per request allowed by the API


def get_config():
    """Load and validate configuration from the .env file."""
    load_dotenv()

    # Default the date range to today (MM/DD/YYYY) when not provided, so a
    # daily scheduled run always searches the current day.
    today = date.today().strftime("%m/%d/%Y")

    config = {
        "api_key": os.getenv("SAM_API_KEY"),
        "posted_from": os.getenv("INITIAL_DATE") or today,
        "posted_to": os.getenv("FINAL_DATE") or today,
        "keyword": os.getenv("KEYWORD", ""),
        "output_file": os.getenv("OUTPUT_FILE", "GELSIGHT_OPPORTUNITIES.xlsx"),
    }

    if not config["api_key"]:
        sys.exit("Missing required value in .env: SAM_API_KEY")

    return config


def fetch_opportunities(config):
    """Fetch all matching opportunity records, handling pagination."""
    records = []
    offset = 0

    while True:
        params = {
            "api_key": config["api_key"],
            "postedFrom": config["posted_from"],
            "postedTo": config["posted_to"],
            "title": config["keyword"],
            "limit": PAGE_SIZE,
            "offset": offset,
        }

        response = requests.get(API_URL, params=params, timeout=60)
        response.raise_for_status()
        data = response.json()

        page = data.get("opportunitiesData", []) or []
        records.extend(page)

        total = data.get("totalRecords", len(records))
        print(f"Fetched {len(records)} of {total} record(s)...")

        offset += PAGE_SIZE
        if offset >= total or not page:
            break

    return records


def flatten(records):
    """Flatten nested fields so they fit cleanly into spreadsheet columns."""
    rows = []
    for r in records:
        office = r.get("officeAddress") or {}
        poc_list = r.get("pointOfContact") or []
        poc = poc_list[0] if poc_list else {}

        rows.append({
            "Notice ID": r.get("noticeId"),
            "Title": r.get("title"),
            "Solicitation Number": r.get("solicitationNumber"),
            "Department/Agency": r.get("fullParentPathName"),
            "Type": r.get("type"),
            "Base Type": r.get("baseType"),
            "Posted Date": r.get("postedDate"),
            "Response Deadline": r.get("responseDeadLine"),
            "NAICS Code": r.get("naicsCode"),
            "Classification Code": r.get("classificationCode"),
            "Set Aside": r.get("typeOfSetAside"),
            "Active": r.get("active"),
            "Office City": office.get("city"),
            "Office State": office.get("state"),
            "Office ZIP": office.get("zipcode"),
            "Contact Name": poc.get("fullName"),
            "Contact Email": poc.get("email"),
            "Contact Phone": poc.get("phone"),
            "Description Link": r.get("description"),
            "UI Link": r.get("uiLink"),
        })
    return rows


def main():
    config = get_config()
    print(
        f"Searching SAM.gov for title '{config['keyword']}' "
        f"from {config['posted_from']} to {config['posted_to']}..."
    )

    records = fetch_opportunities(config)

    if not records:
        print("No opportunities found for the given parameters.")
        new_df = pd.DataFrame()
    else:
        new_df = pd.DataFrame(flatten(records))

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               config["output_file"])

    # Append to the existing spreadsheet instead of overwriting it.
    if os.path.exists(output_path):
        existing_df = pd.read_excel(output_path, sheet_name="Opportunities")
        before = len(existing_df)
        combined = pd.concat([existing_df, new_df], ignore_index=True)
        # Drop duplicates so re-running does not add the same opportunity twice.
        if "Notice ID" in combined.columns:
            combined = combined.drop_duplicates(subset=["Notice ID"], keep="first")
        else:
            combined = combined.drop_duplicates()
        added = len(combined) - before
        print(f"Existing file had {before} record(s); appended {added} new record(s).")
    else:
        combined = new_df
        print(f"Creating new file with {len(combined)} record(s).")

    if combined.empty:
        combined = pd.DataFrame(columns=["Notice ID", "Title"])

    combined.to_excel(output_path, index=False, sheet_name="Opportunities")
    print(f"Saved {len(combined)} total record(s) to {output_path}")


if __name__ == "__main__":
    main()
