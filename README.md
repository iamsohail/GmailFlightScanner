# Gmail Flight Scanner

CLI tool that scans your Gmail for flight-related emails and exports structured flight details to a CSV file.

## Features

- OAuth2 authentication with Gmail (read-only access)
- Searches across multiple flight-related queries (boarding pass, e-ticket, PNR, etc.)
- Extracts: flight number, airline, route, date, PNR/booking reference
- Deduplicates results across queries
- Outputs sorted CSV file

## Setup

### 1. Google Cloud Console

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (or select an existing one)
3. Enable the **Gmail API**
4. Go to **Credentials** → **Create Credentials** → **OAuth client ID**
5. Application type: **Desktop app**
6. Download the JSON file and save it as `credentials.json` in this directory

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run

```bash
python scanner.py
```

On first run, a browser window opens for OAuth consent. After authorization, a `token.json` is saved for future runs.

## Output

The script generates `flights.csv` with columns:

| Column | Description |
|---|---|
| Date | Flight date (extracted from email body) |
| Airline | Airline name |
| Flight Number | IATA flight number (e.g., AI302) |
| From | Departure airport code |
| To | Arrival airport code |
| PNR/Booking Ref | PNR or booking reference number |
| Email Subject | Original email subject line |
| Email Date | Date the email was received |

## Files

- `scanner.py` - Main script
- `credentials.json` - Your OAuth client secret (not tracked in git)
- `token.json` - OAuth token (auto-generated, not tracked in git)
- `flights.csv` - Output file (not tracked in git)
