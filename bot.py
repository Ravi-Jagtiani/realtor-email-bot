import imaplib
import smtplib
import email
import os
import time
import logging
from datetime import datetime
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── CONFIG ────────────────────────────────────────────────────────────────────
IMAP_SERVER    = os.environ.get("IMAP_SERVER",    "imap.mail.yahoo.com")
IMAP_PORT      = int(os.environ.get("IMAP_PORT",  "993"))
SMTP_SERVER    = os.environ.get("SMTP_SERVER",    "smtp.mail.yahoo.com")
SMTP_PORT      = int(os.environ.get("SMTP_PORT",  "465"))
EMAIL_ADDRESS  = os.environ.get("EMAIL_ADDRESS",  "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
AGENT_NAME     = os.environ.get("AGENT_NAME",     "Ravi Jagtiani")
AGENT_EMAIL    = os.environ.get("AGENT_EMAIL",    "")
BROKERAGE_NAME = os.environ.get("BROKERAGE_NAME", "")
# ─────────────────────────────────────────────────────────────────────────────

# ── INQUIRY KEYWORDS ──────────────────────────────────────────────────────────
INQUIRY_KEYWORDS = [
    "interested in", "inquiry", "enquiry", "open house",
    "listing", "property", "schedule a tour", "more info",
    "available", "asking price", "for sale", "for lease",
    "can you send", "would like to know", "tell me more",
    "viewing", "visit", "showing"
]
# ─────────────────────────────────────────────────────────────────────────────

# ── EMAIL TEMPLATE ────────────────────────────────────────────────────────────
# Replace with Ravi's actual template when received.
# Placeholders: {sender_name}, {property_type}, {listing_url},
#               {contact_name}, {contact_email}, {brokerage_name}
EMAIL_TEMPLATE = """Hi {sender_name},

Thank you for your enquiry regarding our {property_type} listings.

<a href="{listing_url}">Click here</a> for all my {property_type} listings across CA.
Below each listing you will find their respect NDA links to sign and access “Due Diligence” including financials. 

To setup a tour or to know more about any of the listings please call me on 669.226.7416.

Best regards,
{contact_name}
{brokerage_name}
"""
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
def get_listings() -> list[dict]:
    """
    Load property types from Google Sheet.
    Columns: property_type, keywords, listing_url, contact_name, contact_email
    """
    try:
        import json
        import gspread
        from google.oauth2.service_account import Credentials

        creds_json = os.environ.get("GOOGLE_CREDENTIALS", "")
        sheet_name = os.environ.get("SHEET_NAME", "Realtor Listings")

        creds = Credentials.from_service_account_info(
            json.loads(creds_json),
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.readonly"
            ]
        )
        client  = gspread.authorize(creds)
        sheet   = client.open(sheet_name).sheet1
        records = sheet.get_all_records()
        log.info(f"Loaded {len(records)} property type(s) from Google Sheet.")
        return records

    except Exception as e:
        log.error(f"Google Sheets error: {e}")
        return []


# ── EMAIL UTILITIES ───────────────────────────────────────────────────────────
def decode_str(s) -> str:
    if not s:
        return ""
    decoded, enc = decode_header(s)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(enc or "utf-8", errors="replace")
    return decoded


def get_body(msg) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode("utf-8", errors="replace")
    return msg.get_payload(decode=True).decode("utf-8", errors="replace")


def parse_sender(raw: str):
    if "<" in raw:
        parts = raw.split("<")
        return parts[0].strip().strip('"'), parts[1].strip().rstrip(">")
    return "", raw.strip()


# ── INQUIRY DETECTION ─────────────────────────────────────────────────────────
def is_inquiry(subject: str, body: str) -> bool:
    text = (subject + " " + body).lower()
    return any(kw in text for kw in INQUIRY_KEYWORDS)


# ── PROPERTY TYPE MATCHING ────────────────────────────────────────────────────
def match_property_type(subject: str, body: str, listings: list[dict]):
    """
    Match email to a property type by checking keywords column.
    Falls back to first row (general) if nothing matches.
    """
    text = (subject + " " + body).lower()

    for listing in listings:
        keywords = [k.strip().lower() for k in listing.get("keywords", "").split(",")]
        if any(kw in text for kw in keywords if kw):
            return listing

    # No match — return None, bot will send fallback reply
    return None


# ── SEND REPLY ────────────────────────────────────────────────────────────────
def send_reply(to_email: str, sender_name: str, subject: str, listing: dict) -> bool:
    body = EMAIL_TEMPLATE.format(
        sender_name   = sender_name or "there",
        property_type = listing["property_type"],
        listing_url   = listing["listing_url"],
        contact_name  = listing["contact_name"],
        contact_email = listing["contact_email"],
        brokerage_name= BROKERAGE_NAME
    )
    msg = MIMEMultipart()
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = to_email
    msg["Subject"] = f"Re: {subject}" if not subject.lower().startswith("re:") else subject
    msg.attach(MIMEText(body, "html"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        save_to_sent(msg["Subject"], to_email, body)
        log.info(f"✓ Reply sent to {to_email} — {listing['property_type']}")
        return True
    except Exception as e:
        log.error(f"Send failed: {e}")
        return False


def send_fallback_reply(to_email: str, sender_name: str, subject: str) -> bool:
    """
    Send a generic reply when no property type is matched.
    Directs enquirer to Ravi's general Crexi profile.
    """
    body = f"""Hi {sender_name or 'there'},

Thank you for your enquiry. We'd love to help you find the right property.

You can browse all of our available listings here:
https://www.crexi.com/profile/ravi-jagtiani-ravijag

{AGENT_NAME} will be in touch with you shortly. You can also reach them directly at {AGENT_EMAIL}.

Best regards,
{AGENT_NAME}
{BROKERAGE_NAME}
"""
    msg = MIMEMultipart()
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = to_email
    msg["Subject"] = f"Re: {subject}" if not subject.lower().startswith("re:") else subject
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        save_to_sent(msg["Subject"], to_email, body)
        log.info(f"✓ Fallback reply sent to {to_email}")
        return True
    except Exception as e:
        log.error(f"Fallback send failed: {e}")
        return False

def save_to_sent(subject: str, to_email: str, body: str):
    """Save a copy of the sent email to the Sent folder via IMAP."""
    try:
        msg = MIMEMultipart()
        msg["From"]    = EMAIL_ADDRESS
        msg["To"]      = to_email
        msg["Subject"] = subject
        msg["Date"]    = datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0000")
        msg.attach(MIMEText(body, "html"))

        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        # Try common Sent folder names
        for folder in ["Sent", "Sent Items", "INBOX.Sent", "[Gmail]/Sent Mail"]:
            try:
                mail.append(
                    folder, "\\Seen",
                    imaplib.Time2Internaldate(time.time()),
                    msg.as_bytes()
                )
                log.info(f"Saved to {folder} folder.")
                break
            except Exception:
                continue

        mail.logout()
    except Exception as e:
        log.error(f"Could not save to Sent folder: {e}")


# ── MAIN (runs once and exits) ────────────────────────────────────────────────
def main():
    log.info("=" * 50)
    log.info(f"Bot run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("=" * 50)

    # 1. Load property types from Google Sheet
    listings = get_listings()
    if not listings:
        log.warning("No listings loaded — cannot match enquiries. Exiting.")
        return

    # 2. Connect to inbox
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        mail.select("inbox")
    except Exception as e:
        log.error(f"Inbox connection failed: {e}")
        return

    _, data = mail.search(None, "UNSEEN")
    ids     = data[0].split()
    log.info(f"{len(ids)} unread email(s) found.")

    replied  = 0
    skipped  = 0

    for eid in ids:
        _, msg_data = mail.fetch(eid, "(RFC822)")
        msg         = email.message_from_bytes(msg_data[0][1])

        subject    = decode_str(msg.get("Subject", ""))
        raw_from   = decode_str(msg.get("From",    ""))
        body       = get_body(msg)
        name, addr = parse_sender(raw_from)

        log.info(f"→ '{subject}' from {addr}")

        if not is_inquiry(subject, body):
            log.info("  Not an enquiry, skipping.")
            mail.store(eid, '+FLAGS', '\\Seen')
            skipped += 1
            continue

        listing = match_property_type(subject, body, listings)

        if listing:
            log.info(f"  Matched type: {listing['property_type']}")
            if send_reply(addr, name, subject, listing):
                mail.store(eid, '+FLAGS', '\\Seen')
                replied += 1
        else:
            log.info("  No type matched — sending fallback reply.")
            if send_fallback_reply(addr, name, subject):
                mail.store(eid, '+FLAGS', '\\Seen')
                replied += 1

    mail.logout()

    log.info("-" * 50)
    log.info(f"Done. Replied: {replied} | Skipped: {skipped}")
    log.info("=" * 50)


if __name__ == "__main__":
    main()
