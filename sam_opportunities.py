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

The spreadsheet is overwritten on every run. If any opportunities are found it
is emailed as an attachment via the Postmark HTTP API (no SMTP ports needed).
Email settings (all read from .env / environment):

    POSTMARK_API_TOKEN  Postmark "Server API Token" (from the server's API Tokens tab)
    EMAIL_FROM          from address — MUST be a verified Postmark Sender Signature
                        (or on a verified domain in your Postmark account)
    EMAIL_TO            recipient address(es), comma-separated

Usage:
    pip install requests python-dotenv pandas openpyxl
    python sam_opportunities.py
"""

import os
import sys
import base64
from datetime import date

import requests
import pandas as pd
from dotenv import load_dotenv

API_URL = "https://api.sam.gov/opportunities/v2/search"
POSTMARK_URL = "https://api.postmarkapp.com/email"
PAGE_SIZE = 100  # max records per request allowed by the API


def get_config():
    """Load and validate configuration from the .env file."""
    load_dotenv()

    # Default the date range to today (MM/DD/YYYY) when not provided, so a
    # daily scheduled run always searches the current day.
    today = date.today().strftime("%m/%d/%Y")

    # KEYWORD may be a single term or a comma-separated list (each is searched).
    keywords = [k.strip() for k in os.getenv("KEYWORD", "").split(",") if k.strip()]

    config = {
        "api_key": os.getenv("SAM_API_KEY"),
        "posted_from": os.getenv("INITIAL_DATE") or today,
        "posted_to": os.getenv("FINAL_DATE") or today,
        "keywords": keywords,
        "keyword": ", ".join(keywords),  # display string for emails/logs
        "output_file": os.getenv("OUTPUT_FILE", "GELSIGHT_OPPORTUNITIES.xlsx"),
        # Email settings (Postmark)
        "postmark_token": os.getenv("POSTMARK_API_TOKEN"),
        "email_from": os.getenv("EMAIL_FROM"),
        "email_to": os.getenv("EMAIL_TO"),
    }

    if not config["api_key"]:
        sys.exit("Missing required value in .env: SAM_API_KEY")

    return config


def send_email(config, attachment_path, num_records):
    """Email the spreadsheet as an attachment via the Postmark HTTP API."""
    required = {
        "POSTMARK_API_TOKEN": config["postmark_token"],
        "EMAIL_FROM": config["email_from"],
        "EMAIL_TO": config["email_to"],
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        print(f"Skipping email; missing setting(s): {', '.join(missing)}")
        return

    with open(attachment_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("ascii")

    payload = {
        "From": config["email_from"],
        "To": config["email_to"],  # Postmark accepts a comma-separated string
        "Subject": (
            f"SAM.gov opportunities for '{config['keyword']}' "
            f"({config['posted_from']}) - {num_records} found"
        ),
        "TextBody": (
            f"{num_records} opportunity(ies) found on SAM.gov for title "
            f"'{config['keyword']}' on {config['posted_from']}.\n\n"
            f"The full list is attached as {os.path.basename(attachment_path)}."
        ),
        "MessageStream": "outbound",
        "Attachments": [
            {
                "Name": os.path.basename(attachment_path),
                "Content": encoded,
                "ContentType": (
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
            }
        ],
    }

    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Postmark-Server-Token": config["postmark_token"],
    }

    response = requests.post(POSTMARK_URL, json=payload, headers=headers, timeout=60)
    body = response.json() if response.content else {}

    if response.status_code == 200 and body.get("ErrorCode") == 0:
        print(f"Emailed {attachment_path} to {config['email_to']}")
    else:
        sys.exit(
            f"Email failed (HTTP {response.status_code}, "
            f"Postmark ErrorCode {body.get('ErrorCode')}): {body.get('Message')}"
        )


def fetch_opportunities(config, keyword):
    """Fetch all records matching one keyword, handling pagination."""
    records = []
    offset = 0

    while True:
        params = {
            "api_key": config["api_key"],
            "postedFrom": config["posted_from"],
            "postedTo": config["posted_to"],
            "title": keyword,
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


def flatten(records, keyword):
    """Flatten nested fields so they fit cleanly into spreadsheet columns."""
    rows = []
    for r in records:
        office = r.get("officeAddress") or {}
        poc_list = r.get("pointOfContact") or []
        poc = poc_list[0] if poc_list else {}

        rows.append({
            "Notice ID": r.get("noticeId"),
            "Matched Keyword": keyword,
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

    if not config["keywords"]:
        sys.exit("No KEYWORD set in .env — nothing to search for.")

    # Search each keyword and merge results, deduping by Notice ID. When the
    # same notice matches more than one keyword, the keywords are combined.
    merged = {}
    for keyword in config["keywords"]:
        print(
            f"Searching SAM.gov for title '{keyword}' "
            f"from {config['posted_from']} to {config['posted_to']}..."
        )
        records = fetch_opportunities(config, keyword)
        for row in flatten(records, keyword):
            nid = row["Notice ID"]
            if nid in merged:
                seen = set(merged[nid]["Matched Keyword"].split(", "))
                seen.add(keyword)
                merged[nid]["Matched Keyword"] = ", ".join(sorted(seen))
            else:
                merged[nid] = row

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               config["output_file"])

    if not merged:
        print("No opportunities found for the given parameters.")
        df = pd.DataFrame(columns=["Notice ID", "Matched Keyword", "Title"])
    else:
        df = pd.DataFrame(list(merged.values()))

    # Overwrite the spreadsheet on every run.
    df.to_excel(output_path, index=False, sheet_name="Opportunities")
    print(f"Saved {len(merged)} unique record(s) to {output_path}")

    # Email the spreadsheet only when opportunities were found.
    if merged:
        send_email(config, output_path, len(merged))
    else:
        print("No records found; skipping email.")


if __name__ == "__main__":
    main()
