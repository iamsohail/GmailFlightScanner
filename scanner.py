"""Gmail Flight Scanner - Scans Gmail for flight emails and exports details to CSV."""

import base64
import csv
import os
import re
from datetime import datetime

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

SEARCH_QUERIES = [
    '"boarding pass"',
    '"flight confirmation"',
    '"flight itinerary"',
    '"e-ticket"',
    '"booking confirmation" flight',
    '"trip confirmation" flight',
    '"PNR"',
    '"flight booking"',
]

# Known airlines for name extraction
KNOWN_AIRLINES = {
    "airindia": "Air India",
    "indigo": "IndiGo",
    "goindigo": "IndiGo",
    "spicejet": "SpiceJet",
    "vistara": "Vistara",
    "airvistara": "Vistara",
    "akasaair": "Akasa Air",
    "airasia": "AirAsia",
    "emirates": "Emirates",
    "etihad": "Etihad",
    "qatar": "Qatar Airways",
    "qatarairways": "Qatar Airways",
    "singapore": "Singapore Airlines",
    "singaporeair": "Singapore Airlines",
    "lufthansa": "Lufthansa",
    "british": "British Airways",
    "britishairways": "British Airways",
    "klm": "KLM",
    "airfrance": "Air France",
    "united": "United Airlines",
    "delta": "Delta Airlines",
    "american": "American Airlines",
    "southwest": "Southwest Airlines",
    "thai": "Thai Airways",
    "cathay": "Cathay Pacific",
    "cathaypacific": "Cathay Pacific",
    "jet": "Jet Airways",
    "jetairways": "Jet Airways",
    "goair": "Go First",
    "gofirst": "Go First",
    "allianceair": "Alliance Air",
    "starair": "Star Air",
    "flydubai": "FlyDubai",
    "omanair": "Oman Air",
    "saudia": "Saudia",
    "turkish": "Turkish Airlines",
    "turkishairlines": "Turkish Airlines",
}

# IATA 2-letter airline codes
AIRLINE_CODES = {
    "AI": "Air India",
    "6E": "IndiGo",
    "SG": "SpiceJet",
    "UK": "Vistara",
    "QP": "Akasa Air",
    "I5": "AirAsia India",
    "EK": "Emirates",
    "EY": "Etihad",
    "QR": "Qatar Airways",
    "SQ": "Singapore Airlines",
    "LH": "Lufthansa",
    "BA": "British Airways",
    "KL": "KLM",
    "AF": "Air France",
    "UA": "United Airlines",
    "DL": "Delta Airlines",
    "AA": "American Airlines",
    "WN": "Southwest Airlines",
    "TG": "Thai Airways",
    "CX": "Cathay Pacific",
    "9W": "Jet Airways",
    "G8": "Go First",
    "9I": "Alliance Air",
    "S5": "Star Air",
    "FZ": "FlyDubai",
    "WY": "Oman Air",
    "SV": "Saudia",
    "TK": "Turkish Airlines",
}


def authenticate():
    """Authenticate with Gmail API via OAuth2."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists("credentials.json"):
                print("ERROR: credentials.json not found.")
                print("Download your OAuth client secret from Google Cloud Console")
                print("and place it in this directory as credentials.json")
                raise SystemExit(1)
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def search_flights(service):
    """Search Gmail for flight-related emails across multiple queries."""
    message_ids = set()
    messages = []

    for query in SEARCH_QUERIES:
        page_token = None
        while True:
            result = (
                service.users()
                .messages()
                .list(userId="me", q=query, pageToken=page_token)
                .execute()
            )
            batch = result.get("messages", [])
            for msg in batch:
                if msg["id"] not in message_ids:
                    message_ids.add(msg["id"])
                    messages.append(msg)
            page_token = result.get("nextPageToken")
            if not page_token:
                break
        print(f"  Query {query}: found messages (total unique so far: {len(message_ids)})")

    print(f"\nTotal unique flight emails found: {len(messages)}")
    return messages


def get_message_body(payload):
    """Extract plain text or HTML body from a Gmail message payload."""
    plain_text = ""
    html_text = ""

    if payload.get("mimeType") == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            plain_text = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    elif payload.get("mimeType") == "text/html":
        data = payload.get("body", {}).get("data", "")
        if data:
            html_text = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
    elif "parts" in payload:
        for part in payload["parts"]:
            p, h = _extract_parts(part)
            if p:
                plain_text += p
            if h:
                html_text += h

    if plain_text:
        return plain_text
    if html_text:
        soup = BeautifulSoup(html_text, "html.parser")
        return soup.get_text(separator=" ", strip=True)
    return ""


def _extract_parts(part):
    """Recursively extract text from message parts."""
    plain = ""
    html = ""
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


def extract_flight_number(text):
    """Extract flight number in IATA format (e.g., AI302, 6E2341)."""
    pattern = r"\b([A-Z0-9]{2})\s?(\d{1,4})\b"
    matches = re.findall(pattern, text)
    for code, num in matches:
        if code in AIRLINE_CODES or (code[0].isalpha() and code[1].isalpha()):
            return f"{code}{num}"
    # Fallback: look for common patterns with airline context
    pattern2 = r"(?:flight|flt|flt\.)\s*(?:no\.?\s*)?([A-Z0-9]{2}\s?\d{1,4})"
    match = re.search(pattern2, text, re.IGNORECASE)
    if match:
        return match.group(1).replace(" ", "")
    return ""


def extract_airport_codes(text):
    """Extract origin and destination airport codes."""
    from_code = ""
    to_code = ""

    # Pattern: "from XXX to YYY" or "XXX → YYY" or "XXX - YYY"
    patterns = [
        r"(?:from|departure|depart|origin)\s*:?\s*.*?\b([A-Z]{3})\b",
        r"(?:to|arrival|arrive|destination)\s*:?\s*.*?\b([A-Z]{3})\b",
    ]

    from_match = re.search(patterns[0], text, re.IGNORECASE)
    to_match = re.search(patterns[1], text, re.IGNORECASE)

    if from_match:
        candidate = from_match.group(1)
        if candidate.isupper() and candidate not in ("THE", "AND", "FOR", "YOU", "ARE", "HAS", "WAS", "HIS", "HER", "OUR", "NOT", "BUT", "ALL", "CAN", "HAD", "HER", "ONE", "OUR", "OUT", "DAY", "GET", "HIM", "HOW", "ITS", "MAY", "NEW", "NOW", "OLD", "SEE", "WAY", "WHO", "DID", "GOT", "LET", "SAY", "SHE", "TOO", "USE", "PNR"):
            from_code = candidate

    if to_match:
        candidate = to_match.group(1)
        if candidate.isupper() and candidate not in ("THE", "AND", "FOR", "YOU", "ARE", "HAS", "WAS", "HIS", "HER", "OUR", "NOT", "BUT", "ALL", "CAN", "HAD", "HER", "ONE", "OUR", "OUT", "DAY", "GET", "HIM", "HOW", "ITS", "MAY", "NEW", "NOW", "OLD", "SEE", "WAY", "WHO", "DID", "GOT", "LET", "SAY", "SHE", "TOO", "USE", "PNR"):
            to_code = candidate

    # Fallback: look for "XXX → YYY" or "XXX - YYY" or "XXX to YYY" direct patterns
    if not from_code or not to_code:
        route_pattern = r"\b([A-Z]{3})\s*(?:→|->|–|—|-|to)\s*([A-Z]{3})\b"
        route_match = re.search(route_pattern, text)
        if route_match:
            if not from_code:
                from_code = route_match.group(1)
            if not to_code:
                to_code = route_match.group(2)

    return from_code, to_code


def extract_flight_date(text):
    """Extract flight date from email body."""
    date_patterns = [
        # 15 Jan 2025, 15-Jan-2025
        (r"\b(\d{1,2})\s*[-/]?\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*[-/,]?\s*(\d{4})\b", "%d %b %Y"),
        # Jan 15, 2025
        (r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2})\s*[,]?\s*(\d{4})\b", "%b %d %Y"),
        # 2025-01-15
        (r"\b(\d{4})-(\d{2})-(\d{2})\b", "%Y-%m-%d"),
        # 15/01/2025 or 15-01-2025
        (r"\b(\d{2})[/-](\d{2})[/-](\d{4})\b", "%d/%m/%Y"),
    ]

    # Search near flight-related keywords first
    flight_context = re.findall(
        r"(?:date|departure|depart|travel|journey|flight).{0,80}",
        text,
        re.IGNORECASE,
    )
    search_texts = flight_context + [text]

    for search_text in search_texts:
        for pattern, fmt in date_patterns:
            match = re.search(pattern, search_text, re.IGNORECASE)
            if match:
                try:
                    groups = match.groups()
                    if fmt == "%d %b %Y":
                        date_str = f"{groups[0]} {groups[1][:3]} {groups[2]}"
                        return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                    elif fmt == "%b %d %Y":
                        date_str = f"{groups[0][:3]} {groups[1]} {groups[2]}"
                        return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                    elif fmt == "%Y-%m-%d":
                        date_str = f"{groups[0]}-{groups[1]}-{groups[2]}"
                        return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                    elif fmt == "%d/%m/%Y":
                        date_str = f"{groups[0]}/{groups[1]}/{groups[2]}"
                        return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                except ValueError:
                    continue

    return ""


def extract_airline(text, sender):
    """Extract airline name from sender or email body."""
    # Check sender domain
    domain_match = re.search(r"@([\w.-]+)", sender)
    if domain_match:
        domain = domain_match.group(1).lower().replace(".", "")
        for key, airline in KNOWN_AIRLINES.items():
            if key in domain:
                return airline

    # Check flight number for airline code
    flight_num = extract_flight_number(text)
    if flight_num and len(flight_num) >= 2:
        code = flight_num[:2]
        if code in AIRLINE_CODES:
            return AIRLINE_CODES[code]

    # Check body for airline names
    text_lower = text.lower()
    for key, airline in KNOWN_AIRLINES.items():
        if key in text_lower or airline.lower() in text_lower:
            return airline

    return ""


def extract_pnr(text):
    """Extract PNR or booking reference."""
    patterns = [
        r"(?:PNR|pnr|Pnr)\s*(?:no\.?|number|#|:)?\s*:?\s*([A-Z0-9]{5,8})",
        r"(?:booking\s*(?:ref|reference|id|code|no)|confirmation\s*(?:no|number|code|#))\s*:?\s*([A-Z0-9]{5,8})",
        r"(?:reference|ref\.?)\s*(?:no\.?|number|#|:)?\s*:?\s*([A-Z0-9]{6})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def parse_email(service, msg_id):
    """Fetch and parse a single email for flight details."""
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    payload = msg.get("payload", {})
    headers = {h["name"]: h["value"] for h in payload.get("headers", [])}

    subject = headers.get("Subject", "")
    sender = headers.get("From", "")
    email_date_raw = headers.get("Date", "")

    # Parse email received date
    email_date = ""
    if email_date_raw:
        # Clean timezone info for parsing
        clean_date = re.sub(r"\s*\(.*\)", "", email_date_raw)
        for fmt in (
            "%a, %d %b %Y %H:%M:%S %z",
            "%d %b %Y %H:%M:%S %z",
            "%a, %d %b %Y %H:%M:%S",
            "%d %b %Y %H:%M:%S",
        ):
            try:
                email_date = datetime.strptime(clean_date.strip(), fmt).strftime("%Y-%m-%d")
                break
            except ValueError:
                continue

    body = get_message_body(payload)
    full_text = f"{subject} {body}"

    flight_number = extract_flight_number(full_text)
    from_code, to_code = extract_airport_codes(full_text)
    flight_date = extract_flight_date(full_text)
    airline = extract_airline(full_text, sender)
    pnr = extract_pnr(full_text)

    return {
        "Date": flight_date or email_date,
        "Airline": airline,
        "Flight Number": flight_number,
        "From": from_code,
        "To": to_code,
        "PNR/Booking Ref": pnr,
        "Email Subject": subject,
        "Email Date": email_date,
    }


def main():
    print("Gmail Flight Scanner")
    print("=" * 40)

    print("\n1. Authenticating with Gmail...")
    service = authenticate()
    print("   Authenticated successfully.")

    print("\n2. Searching for flight emails...")
    messages = search_flights(service)

    if not messages:
        print("\nNo flight-related emails found.")
        return

    print(f"\n3. Parsing {len(messages)} emails for flight details...")
    flights = []
    for i, msg in enumerate(messages, 1):
        if i % 10 == 0 or i == len(messages):
            print(f"   Processed {i}/{len(messages)} emails...")
        try:
            flight = parse_email(service, msg["id"])
            flights.append(flight)
        except Exception as e:
            print(f"   Warning: Failed to parse message {msg['id']}: {e}")

    # Sort by date (oldest first), empty dates go last
    flights.sort(key=lambda f: f["Date"] if f["Date"] else "9999-99-99")

    output_file = "flights.csv"
    fieldnames = [
        "Date",
        "Airline",
        "Flight Number",
        "From",
        "To",
        "PNR/Booking Ref",
        "Email Subject",
        "Email Date",
    ]

    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(flights)

    print(f"\n4. Results saved to {output_file}")
    print(f"   Total flight emails: {len(flights)}")

    # Summary stats
    with_flight_num = sum(1 for f in flights if f["Flight Number"])
    with_pnr = sum(1 for f in flights if f["PNR/Booking Ref"])
    with_route = sum(1 for f in flights if f["From"] and f["To"])
    print(f"   With flight number: {with_flight_num}")
    print(f"   With PNR/booking ref: {with_pnr}")
    print(f"   With route (from/to): {with_route}")


if __name__ == "__main__":
    main()
