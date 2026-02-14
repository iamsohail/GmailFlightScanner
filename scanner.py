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

# Passenger name variants to filter bookings (case-insensitive)
PASSENGER_NAMES = ["mohammad sohail ahmad", "sohail ahmad"]

SEARCH_QUERIES = [
    # Generic flight terms
    '"boarding pass"',
    '"flight confirmation"',
    '"flight itinerary"',
    '"e-ticket"',
    '"booking confirmation" flight',
    '"trip confirmation" flight',
    '"PNR"',
    '"flight booking"',
    # Active Indian airlines
    'from:indigo subject:itinerary',
    'from:goindigo subject:itinerary',
    'from:airindia',
    'from:airindiaexpress',
    'from:spicejet',
    'from:akasaair',
    'from:allianceair',
    'from:starair',
    'from:flybig',
    # Defunct / renamed Indian airlines
    'from:jetairways',
    'from:goair',
    'from:gofirst',
    'from:airasiago',
    'from:airasia subject:booking',
    'from:airdeccan',
    'from:airsahara',
    'from:kingfisherairlines',
    'from:flygokingfisher',
    'from:airvistara',
    'from:vfrpl',
    'from:aircosta',
    'from:airpegasus',
    'from:trujet',
    'from:paramountairways',
    'from:mdlrairlines',
    'from:zoomair',
    # Booking platforms (flight-specific)
    'from:makemytrip flight',
    'from:ixigo flight',
    'from:cleartrip flight',
    'from:yatra flight',
    'from:easemytrip flight',
    'from:goibibo flight',
    'from:happyeasygo flight',
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

# Common 3-letter English words to exclude from airport code detection
AIRPORT_STOPWORDS = {
    "THE", "AND", "FOR", "YOU", "ARE", "HAS", "WAS", "HIS", "HER", "OUR",
    "NOT", "BUT", "ALL", "CAN", "HAD", "ONE", "OUT", "DAY", "GET", "HIM",
    "HOW", "ITS", "MAY", "NEW", "NOW", "OLD", "SEE", "WAY", "WHO", "DID",
    "GOT", "LET", "SAY", "SHE", "TOO", "USE", "PNR", "ADD", "COM", "NON",
    "END", "OFF", "RUN", "SET", "TRY", "PUT", "BIG", "FEW", "FAR", "OWN",
    "SAT", "SIT", "TOP", "RED", "HOT", "CUT", "AGO", "YES", "YET", "RAN",
    "BED", "BOX", "BOY", "CAR", "DOG", "EAR", "EAT", "EYE", "FLY", "GAS",
    "GUN", "HIT", "JOB", "KEY", "LAY", "LEG", "LET", "LIE", "MAP", "MRS",
    "OIL", "PAY", "PER", "RUN", "SIX", "SUN", "TEN", "WAR", "WET", "WIN",
    "WON", "YES", "AIR", "ACT", "AGE", "AID", "AIM", "ART", "ASK", "BAD",
    "BAR", "BIT", "BUY", "COP", "CRY", "DIE", "DIG", "DRY", "DUE", "ERA",
    "FAN", "FAT", "FEE", "FIT", "FUN", "GAP", "HAT", "HIT", "ICE", "ILL",
    "JAM", "JET", "LAW", "LAP", "LOG", "LOT", "LOW", "MAN", "MEN", "MET",
    "MIX", "MOB", "MUD", "NET", "NOR", "NUT", "ODD", "PAN", "PEN", "PET",
    "PIN", "PIT", "POT", "RAW", "RIB", "RID", "ROB", "ROD", "ROW", "RUB",
    "SAD", "SIP", "SKI", "TAP", "TAX", "TIE", "TIN", "TIP", "TOE", "TON",
    "TOW", "TOY", "TUB", "VAN", "VIA", "VOW", "WEB", "WIG", "WIT", "WOE",
    "YEN", "ZOO", "FWD", "REF", "INR", "USD", "EUR", "SMS", "OTP", "URL",
    "PDF", "APP", "API", "RSS", "FAQ", "TBA", "TBD", "ETA", "ETD", "GMT",
    "IST", "EST", "PST", "CST", "UTC", "BAG", "DEP", "ARR", "VIA", "FLT",
    "ONS", "UAE", "USA", "DGR", "VRM", "STD", "STA", "AVL", "CNF", "RAC",
    "GEN", "TAT", "OBC", "INF", "ADT", "CHD", "PAX", "SEQ", "ROW", "QTY",
    "AMT", "TAX", "SUB", "TTL", "NET", "MAX", "MIN", "AVG", "REQ", "RES",
    "MOB", "TEL", "WEB", "ORG", "GOV", "EDU", "MIL", "INT", "EXT", "SRC",
    "DST", "MSG", "ERR", "LOG", "CMD", "SYS", "BUS", "CAB",
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
                .list(userId="me", q=query, pageToken=page_token, includeSpamTrash=True)
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

    # Pattern: keyword followed by a 3-letter uppercase airport code (not IGNORECASE for the code)
    from_patterns = [
        r"(?i)(?:from|departure|depart|origin)\s*:?\s*.{0,30}?\b([A-Z]{3})\b",
    ]
    to_patterns = [
        r"(?i)(?:to|arrival|arrive|destination)\s*:?\s*.{0,30}?\b([A-Z]{3})\b",
    ]

    def _valid_airport(code):
        return code.isupper() and code not in AIRPORT_STOPWORDS

    for pattern in from_patterns:
        for match in re.finditer(pattern, text):
            candidate = match.group(1)
            if _valid_airport(candidate):
                from_code = candidate
                break

    for pattern in to_patterns:
        for match in re.finditer(pattern, text):
            candidate = match.group(1)
            if _valid_airport(candidate):
                to_code = candidate
                break

    # Fallback: look for "XXX → YYY" or "XXX - YYY" route patterns (strict uppercase only)
    if not from_code or not to_code:
        route_pattern = r"\b([A-Z]{3})\s*(?:→|->|–|—|-)\s*([A-Z]{3})\b"
        route_match = re.search(route_pattern, text)
        if route_match:
            c1, c2 = route_match.group(1), route_match.group(2)
            if not from_code and _valid_airport(c1):
                from_code = c1
            if not to_code and _valid_airport(c2):
                to_code = c2

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
                        parsed = datetime.strptime(date_str, fmt)
                    elif fmt == "%b %d %Y":
                        date_str = f"{groups[0][:3]} {groups[1]} {groups[2]}"
                        parsed = datetime.strptime(date_str, fmt)
                    elif fmt == "%Y-%m-%d":
                        date_str = f"{groups[0]}-{groups[1]}-{groups[2]}"
                        parsed = datetime.strptime(date_str, fmt)
                    elif fmt == "%d/%m/%Y":
                        date_str = f"{groups[0]}/{groups[1]}/{groups[2]}"
                        parsed = datetime.strptime(date_str, fmt)
                    else:
                        continue
                    # Reject dates outside reasonable range
                    if 1990 <= parsed.year <= 2030:
                        return parsed.strftime("%Y-%m-%d")
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
        r"(?:PNR|pnr|Pnr)\s*(?:no\.?|number|#|:)?\s*:?\s*\b([A-Z0-9]{5,8})\b",
        r"(?:booking\s*(?:ref|reference|id|code|no)|confirmation\s*(?:no|number|code|#))\s*:?\s*\b([A-Z0-9]{5,8})\b",
        r"(?:reference|ref\.?)\s*(?:no\.?|number|#|:)?\s*:?\s*\b([A-Z0-9]{6})\b",
    ]
    # Common words/substrings that get falsely captured as PNR codes
    pnr_stopwords = {
        "NUMBER", "REFERENCE", "BOOKING", "CONFIRM", "DETAIL", "DETAILS",
        "FLIGHT", "STATUS", "CANCEL", "CHANGE", "UPDATE", "ERENCE",
        "RENCE", "UMBER", "ATION", "UMBER", "TICKET", "TRAVEL",
        "PLEASE", "REFUND", "AMOUNT", "TOTAL", "PRICE", "CHARGE",
        "EMAIL", "ISSUE", "BOARD", "CHECK", "PRINT", "VALID",
        "NUMERIC", "STRING", "FORMAT", "RETURN",
    }
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            candidate = match.group(1).upper()
            if candidate not in pnr_stopwords:
                return candidate
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

    # Check if passenger name appears in the email
    text_lower = full_text.lower()
    has_passenger_name = any(name in text_lower for name in PASSENGER_NAMES)

    return {
        "Date": flight_date or email_date,
        "Airline": airline,
        "Flight Number": flight_number,
        "From": from_code,
        "To": to_code,
        "PNR/Booking Ref": pnr,
        "Email Subject": subject,
        "Email Date": email_date,
        "_has_name": has_passenger_name,
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

    # Exclude non-flight emails by subject keywords
    EXCLUDE_SUBJECTS = [
        "hotel booking", "bus booking", "bus ticket", "credit card",
        "account summary", "savings of rs", "missed out on saving",
        "message from our ceo", "message from the ceo",
        "discounts in dubai", "big discounts", "credit note",
        "tax invoice", "gst invoice", "vrl travels",
        "credit card communication", "voucher worth",
        "challenge #", "intermiles credited",
    ]
    def _is_excluded(subject):
        s = subject.lower()
        return any(kw in s for kw in EXCLUDE_SUBJECTS)

    non_excluded = [f for f in flights if not _is_excluded(f["Email Subject"])]
    print(f"   Excluded {len(flights) - len(non_excluded)} non-flight emails by subject")
    flights = non_excluded

    # Filter: keep only actual flight bookings (must have flight number, PNR, or route)
    actual_flights = [
        f for f in flights
        if f["Flight Number"] or f["PNR/Booking Ref"] or (f["From"] and f["To"])
    ]
    print(f"   Filtered to {len(actual_flights)} actual flight records (removed {len(flights) - len(actual_flights)} non-flight emails)")

    # Filter: keep only bookings with passenger name
    named_flights = [f for f in actual_flights if f["_has_name"]]
    print(f"   Filtered to {len(named_flights)} records matching passenger name (removed {len(actual_flights) - len(named_flights)} others)")
    flights = named_flights

    # Remove internal field before writing CSV
    for f in flights:
        del f["_has_name"]

    # Deduplicate by PNR: keep the record with the most extracted data
    def _richness(f):
        """Score how much data a record has (more = better)."""
        return sum([
            bool(f["Flight Number"]),
            bool(f["From"]),
            bool(f["To"]),
            bool(f["Airline"]),
            bool(f["Date"]),
            bool(f["PNR/Booking Ref"]),
        ])

    seen_pnrs = {}
    deduped = []
    for f in flights:
        pnr = f["PNR/Booking Ref"]
        if not pnr:
            # No PNR — keep as-is (deduplicate by flight number + date instead)
            key = (f["Flight Number"], f["Date"])
            if key == ("", ""):
                deduped.append(f)
            elif key not in seen_pnrs:
                seen_pnrs[key] = f
                deduped.append(f)
            elif _richness(f) > _richness(seen_pnrs[key]):
                deduped.remove(seen_pnrs[key])
                seen_pnrs[key] = f
                deduped.append(f)
        elif pnr not in seen_pnrs:
            seen_pnrs[pnr] = f
            deduped.append(f)
        elif _richness(f) > _richness(seen_pnrs[pnr]):
            deduped.remove(seen_pnrs[pnr])
            seen_pnrs[pnr] = f
            deduped.append(f)

    print(f"   Deduplicated to {len(deduped)} unique flights (removed {len(flights) - len(deduped)} duplicates)")
    flights = deduped

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
