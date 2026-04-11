"""
Realtor Email Auto-Reply Bot
────────────────────────────
Runs once and exits (GitHub Actions cron).
- Loads property types from Google Sheet
- Reads unread enquiry emails
- Extracts the actual enquirer's email from the form email body
- Sends HTML reply with signature to enquirer, CC's the listing contact
- Marks each email as read after processing so it never replies twice
- Saves a copy to Sent folder
"""

import imaplib
import smtplib
import email
import re
import os
import json
import time
import logging
from datetime import datetime
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

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

# ── TEMPLATES ─────────────────────────────────────────────────────────────────
EMAIL_TEMPLATE = """\
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.6; max-width: 700px; margin: 0; padding: 20px;">

  <p>Hi {sender_name},</p>

  <p>Thank you for your enquiry regarding our <strong>{property_type}</strong> listings.</p>

  <p>You can <a href="{listing_url}">click here</a> to view all available {property_type} properties.</p>

  <p>{contact_name} will be in touch with you shortly to discuss your requirements.<br>
  You can also reach them directly at <a href="mailto:{contact_email}">{contact_email}</a>.</p>

  <br>
  <hr style="border: none; border-top: 1px solid #cccccc; margin: 24px 0;">
  {signature}

</body>
</html>"""

FALLBACK_TEMPLATE = """\
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.6; max-width: 700px; margin: 0; padding: 20px;">

  <p>Hi {sender_name},</p>

  <p>Thank you for your enquiry. We would love to help you find the right property.</p>

  <p>You can <a href="https://www.crexi.com/profile/ravi-jagtiani-ravijag">click here</a>
  to browse all of our available listings.</p>

  <p>{contact_name} will be in touch with you shortly.<br>
  You can also reach them directly at <a href="mailto:{contact_email}">{contact_email}</a>.</p>

  <br>
  <hr style="border: none; border-top: 1px solid #cccccc; margin: 24px 0;">
  {signature}

</body>
</html>"""

SIGNATURE_HTML = """\
<table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; font-size: 13px; color: #222; line-height: 1.7;">
  <tr><td style="padding-bottom:14px;">
    <<img src="https://raw.githubusercontent.com/ShreyaJ3147/realtor-email-bot/main/signature.png" width="600" alt="Ravi Jagtiani - Jagtiani Group" style="display:block; max-width:100%;"/>"
         style="display:block; max-width:100%;"/>
  </td></tr>
  <tr><td>
    <strong>Ravi R Jagtiani | Jagtiani Group | Managing Director</strong><br>
    <strong>Cal DRE# 02044082 - Realtor&reg; | President's Circle |
    America's Top 1% Real Estate Professional |
    Voted the Best Commercial Realtor in San Mateo County</strong>
  </td></tr>
  <tr><td style="padding-top:10px;">
    mobile: <a href="tel:6692267416" style="color:#222;">669.226.7416</a><br>
    email: <a href="mailto:ravi@jagtianigroup.com" style="color:#1a73e8;">ravi@jagtianigroup.com</a><br>
    web: <a href="https://www.JagtianiGroup.com/commercial" style="color:#1a73e8;">www.JagtianiGroup.com/commercial</a>
  </td></tr>
  <tr><td style="padding-top:10px; color:#555; font-style:italic;">
    In the business of wealth creation
  </td></tr>
  <tr><td style="padding-top:10px;">
    <a href="https://www.linkedin.com/in/ravijagtiani" style="color:#1a73e8;">Linkedin</a> &nbsp;|&nbsp;
    <a href="https://www.facebook.com/JagtianiGroup" style="color:#1a73e8;">Facebook Business Page</a> &nbsp;|&nbsp;
    <a href="https://www.youtube.com/@jagtianigroup" style="color:#1a73e8;">Youtube</a> &nbsp;|&nbsp;
    <a href="https://www.zillow.com/profile/ravijag" style="color:#1a73e8;">My Reviews</a> &nbsp;|&nbsp;
    <a href="https://blog.jagtianigroup.com" style="color:#1a73e8;">My blog</a>
  </td></tr>
  <tr><td style="padding-top:12px; font-size:12px; color:#777;">
    Intero has been voted the &lsquo;Best Real Estate Company&rsquo; in the East Bay
    and Silicon Valley by the Bay Area News Group for 2016, 2017, &amp; 2018!
  </td></tr>
</table>"""
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
def get_listings() -> list[dict]:
    try:
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
    payload = msg.get_payload(decode=True)
    return payload.decode("utf-8", errors="replace") if payload else ""


def parse_sender(raw: str):
    if "<" in raw:
        parts = raw.split("<")
        return parts[0].strip().strip('"'), parts[1].strip().rstrip(">")
    return "", raw.strip()


def extract_enquirer_email(body: str) -> str:
    """
    Extract the actual enquirer email from a form submission body.
    Tries labelled patterns first (Email: x@x.com), then any email in body.
    Skips Ravi's own email address.
    """
    labelled = re.search(
        r'(?:email|e-mail|reply.?to|contact)[^\n:]*[:\s]+([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
        body, re.IGNORECASE
    )
    if labelled:
        return labelled.group(1).strip()

    all_emails = re.findall(
        r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',
        body
    )
    for addr in all_emails:
        if addr.lower() != EMAIL_ADDRESS.lower():
            return addr
    return ""


def extract_enquirer_name(body: str) -> str:
    """Extract enquirer name from common form field patterns."""
    match = re.search(
        r'(?:full name|name|first name)[^\n:]*[:\s]+([A-Za-z][A-Za-z\s\-]{1,40})',
        body, re.IGNORECASE
    )
    if match:
        name = match.group(1).strip()
        name = re.split(r'\n|last name|email|phone|message', name, flags=re.IGNORECASE)[0].strip()
        return name
    return ""


# ── INQUIRY DETECTION ─────────────────────────────────────────────────────────
def is_inquiry(subject: str, body: str) -> bool:
    text = (subject + " " + body).lower()
    return any(kw in text for kw in INQUIRY_KEYWORDS)


# ── PROPERTY TYPE MATCHING ────────────────────────────────────────────────────
def match_property_type(subject: str, body: str, listings: list[dict]):
    text = (subject + " " + body).lower()
    for listing in listings:
        keywords = [k.strip().lower() for k in listing.get("keywords", "").split(",")]
        if any(kw in text for kw in keywords if kw):
            return listing
    return None


# ── EMAIL BUILDER ─────────────────────────────────────────────────────────────
def build_email_msg(to_email: str, cc_email: str, subject: str, html_body: str) -> MIMEMultipart:
    msg = MIMEMultipart("related")
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = to_email
    if cc_email:
        msg["Cc"]  = cc_email
    msg["Subject"] = subject

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt)

    try:
        with open("signature.png", "rb") as f:
            img = MIMEImage(f.read(), _subtype="png")
            img.add_header("Content-ID", "<signature>")
            img.add_header("Content-Disposition", "inline", filename="signature.png")
            msg.attach(img)
    except Exception as e:
        log.warning(f"Signature image not found: {e}")

    return msg


# ── SEND + SAVE ───────────────────────────────────────────────────────────────
def send_message(msg: MIMEMultipart, to_email: str, cc_email: str) -> bool:
    recipients = [to_email] + ([cc_email] if cc_email else [])
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, recipients, msg.as_bytes())
        return True
    except Exception as e:
        log.error(f"SMTP send failed: {e}")
        return False


def save_to_sent(msg: MIMEMultipart):
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        for folder in ["Sent", "Sent Items", "INBOX.Sent"]:
            try:
                mail.append(
                    folder, "\\Seen",
                    imaplib.Time2Internaldate(time.time()),
                    msg.as_bytes()
                )
                log.info(f"Saved to '{folder}' folder.")
                break
            except Exception:
                continue
        mail.logout()
    except Exception as e:
        log.warning(f"Could not save to Sent folder: {e}")


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    log.info("=" * 50)
    log.info(f"Bot run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("=" * 50)

    listings = get_listings()
    if not listings:
        log.warning("No listings loaded. Exiting.")
        return

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        mail.select("inbox")
    except Exception as e:
        log.error(f"Inbox connection failed: {e}")
        return

    _, data = mail.search(None, "UNSEEN")
    ids = data[0].split()
    log.info(f"{len(ids)} unread email(s) found.")

    replied = 0
    skipped = 0

    for eid in ids:
        _, msg_data = mail.fetch(eid, "(RFC822)")
        msg         = email.message_from_bytes(msg_data[0][1])

        subject  = decode_str(msg.get("Subject", ""))
        raw_from = decode_str(msg.get("From", ""))
        body     = get_body(msg)
        name , form_sender = parse_sender(raw_from)

        log.info(f"→ '{subject}' from {form_sender}")

        # Mark as read immediately — prevents any double processing
        mail.store(eid, '+FLAGS', '\\Seen')

        if not is_inquiry(subject, body):
            log.info("  Not an enquiry, skipping.")
            skipped += 1
            continue

        # Extract actual enquirer contact from form body
        # Falls back to the direct sender if no email found in body
        enquirer_email = extract_enquirer_email(body) or form_sender
        enquirer_name  = extract_enquirer_name(body) or name or "there"

        if not enquirer_email:
            log.warning("  No enquirer email found — skipping.")
            skipped += 1
            continue

        log.info(f"  Enquirer: {enquirer_name} <{enquirer_email}>")

        listing    = match_property_type(subject, body, listings)
        cc_email   = listing["contact_email"] if listing else AGENT_EMAIL
        reply_subj = f"Re: {subject}" if not subject.lower().startswith("re:") else subject

        if listing:
            log.info(f"  Matched type: {listing['property_type']}")
            html_body = EMAIL_TEMPLATE.format(
                sender_name   = enquirer_name,
                property_type = listing["property_type"],
                listing_url   = listing["listing_url"],
                contact_name  = listing["contact_name"],
                contact_email = listing["contact_email"],
                brokerage_name= BROKERAGE_NAME,
                signature     = SIGNATURE_HTML
            )
        else:
            log.info("  No type matched — sending fallback.")
            html_body = FALLBACK_TEMPLATE.format(
                sender_name   = enquirer_name,
                contact_name  = AGENT_NAME,
                contact_email = AGENT_EMAIL,
                signature     = SIGNATURE_HTML
            )

        out_msg = build_email_msg(enquirer_email, cc_email, reply_subj, html_body)

        if send_message(out_msg, enquirer_email, cc_email):
            save_to_sent(out_msg)
            log.info(f"  ✓ Sent to {enquirer_email} | CC: {cc_email}")
            replied += 1
        else:
            skipped += 1

    mail.logout()
    log.info("-" * 50)
    log.info(f"Done. Replied: {replied} | Skipped: {skipped}")
    log.info("=" * 50)


if __name__ == "__main__":
    main()
