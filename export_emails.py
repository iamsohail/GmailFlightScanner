"""Export all matched flight emails to a readable text file for verification."""

import base64
import csv
import os
import re
from datetime import datetime

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def authenticate():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise SystemExit("No valid token.json found. Run scanner.py first.")
    return build("gmail", "v1", credentials=creds)


def get_body(payload):
    plain, html = "", ""
    mime = payload.get("mimeType", "")
    if mime == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            plain = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    elif mime == "text/html":
        data = payload.get("body", {}).get("data", "")
        if data:
            html = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    elif "parts" in payload:
        for part in payload["parts"]:
            p, h = _extract_parts(part)
            plain += p
            html += h
    if plain:
        return plain
    if html:
        return BeautifulSoup(html, "html.parser").get_text(separator=" ", strip=True)
    return "(no body)"


def _extract_parts(part):
    plain, html = "", ""
    mime = part.get("mimeType", "")
    if mime == "text/plain":
        data = part.get("body", {}).get("data", "")
        if data:
            plain = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    elif mime == "text/html":
        data = part.get("body", {}).get("data", "")
        if data:
            html = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    elif "parts" in part:
        for sub in part["parts"]:
            p, h = _extract_parts(sub)
            plain += p
            html += h
    return plain, html


def search_email(service, row):
    """Search for the email matching this flight row."""
    pnr = row["PNR/Booking Ref"].strip()
    subject = row["Email Subject"].strip()

    # Try PNR search first (most unique)
    if pnr:
        result = service.users().messages().list(userId="me", q=pnr, includeSpamTrash=True).execute()
        msgs = result.get("messages", [])
        if msgs:
            return msgs[0]["id"]

    # Fallback: search by subject
    if subject:
        # Use first 60 chars of subject, escape quotes
        q = subject[:60].replace('"', '\\"')
        result = service.users().messages().list(userId="me", q=f'subject:"{q}"', includeSpamTrash=True).execute()
        msgs = result.get("messages", [])
        if msgs:
            return msgs[0]["id"]

    return None


def main():
    print("Exporting flight emails for verification...")
    service = authenticate()

    rows = []
    with open("flights.csv", "r") as f:
        rows = list(csv.DictReader(f))

    print(f"Found {len(rows)} flights in CSV\n")

    output_lines = []
    output_lines.append("=" * 80)
    output_lines.append("FLIGHT EMAILS - VERIFICATION EXPORT")
    output_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    output_lines.append(f"Total flights: {len(rows)}")
    output_lines.append("=" * 80)

    for i, row in enumerate(rows, 1):
        print(f"  Fetching {i}/{len(rows)}: {row['Date']} {row['Airline']} {row['PNR/Booking Ref']}...")

        output_lines.append("")
        output_lines.append(f"{'─' * 80}")
        output_lines.append(f"FLIGHT #{i}")
        output_lines.append(f"{'─' * 80}")
        output_lines.append(f"  Date:          {row['Date']}")
        output_lines.append(f"  Airline:       {row['Airline']}")
        output_lines.append(f"  Flight Number: {row['Flight Number']}")
        output_lines.append(f"  From:          {row['From']}")
        output_lines.append(f"  To:            {row['To']}")
        output_lines.append(f"  PNR:           {row['PNR/Booking Ref']}")
        output_lines.append(f"  Email Subject: {row['Email Subject'].strip()}")
        output_lines.append(f"  Email Date:    {row['Email Date']}")
        output_lines.append("")

        msg_id = search_email(service, row)
        if msg_id:
            msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
            payload = msg.get("payload", {})
            headers = {h["name"]: h["value"] for h in payload.get("headers", [])}

            output_lines.append(f"  --- EMAIL HEADERS ---")
            output_lines.append(f"  From:    {headers.get('From', 'N/A')}")
            output_lines.append(f"  To:      {headers.get('To', 'N/A')}")
            output_lines.append(f"  Date:    {headers.get('Date', 'N/A')}")
            output_lines.append(f"  Subject: {headers.get('Subject', 'N/A')}")
            output_lines.append("")
            output_lines.append(f"  --- EMAIL BODY ---")

            body = get_body(payload)
            # Trim very long bodies to 2000 chars
            if len(body) > 2000:
                body = body[:2000] + "\n  ... [TRUNCATED] ..."
            for line in body.split("\n"):
                output_lines.append(f"  {line}")
        else:
            output_lines.append("  [EMAIL NOT FOUND]")

    output_file = "flight_emails_verification.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))

    print(f"\nDone! Saved to {output_file}")
    print(f"Open it with: open {output_file}")


if __name__ == "__main__":
    main()
