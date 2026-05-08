"""
Realtor Email Auto-Reply Bot
────────────────────────────
Runs once and exits (GitHub Actions cron).
- Loads property types + team members from Google Sheet
- Handles both direct emails and form submission emails
- Extracts enquirer contact from Crexi, BizBuySell, LoopNet and direct emails
- Sends HTML reply with signature image (hosted on GitHub)
- CC's the relevant team member + Ravi always
- Marks emails as read to prevent duplicate replies
- Saves copy to Sent folder
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

# ── CONFIG ────────────────────────────────────────────────────────────────────
IMAP_SERVER    = os.environ.get("IMAP_SERVER",    "imap.mail.yahoo.com")
IMAP_PORT      = int(os.environ.get("IMAP_PORT",  "993"))
SMTP_SERVER    = os.environ.get("SMTP_SERVER",    "smtp.mail.yahoo.com")
SMTP_PORT      = int(os.environ.get("SMTP_PORT",  "465"))
EMAIL_ADDRESS  = os.environ.get("EMAIL_ADDRESS",  "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
AGENT_NAME     = os.environ.get("AGENT_NAME",     "Ravi Jagtiani")
AGENT_EMAIL    = os.environ.get("AGENT_EMAIL",    "")
AGENT_PHONE    = os.environ.get("AGENT_PHONE",    "669.226.7416")
BROKERAGE_NAME = os.environ.get("BROKERAGE_NAME", "Jagtiani Group")
# Hosted signature image — update GITHUB_USER to your GitHub username
GITHUB_USER    = os.environ.get("GITHUB_USER",    "")
SIGNATURE_URL  = f"https://raw.githubusercontent.com/ShreyaJ3147/realtor-email-bot/main/signature.png"
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
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222222;
             line-height: 1.7; max-width: 680px; margin: 0; padding: 20px;">

  <p>Hi {sender_name},</p>

  <p>Here is a link to all the
  <a href="{listing_url}">{property_type} listings</a>.</p>

  <p>The link above includes the {region} location. Below each listing you will
  find the respective NDA links to sign and access due diligence.</p>

  <p>You can call me to know more about the business:
  <a href="tel:{agent_phone_digits}">{agent_phone}</a></p>

  <p>To set up a tour, please call my team member
  <strong>{team_name}</strong>:
  <a href="tel:{team_phone_digits}">{team_phone}</a></p>

  <p>Thanks,<br>
  {agent_name}</p>

  <br>
  <hr style="border: none; border-top: 1px solid #cccccc; margin: 24px 0;">

  <table cellpadding="0" cellspacing="0" border="0"
         style="font-family: Arial, sans-serif; font-size: 13px;
                color: #222222; line-height: 1.7;">
    <tr><td style="padding-bottom: 14px;">
      <img src="{signature_url}" width="600"
           alt="Ravi Jagtiani - Jagtiani Group"
           style="display: block; max-width: 100%;"/>
    </td></tr>
    <tr><td>
      <strong>Ravi R Jagtiani | Jagtiani Group | Managing Director</strong><br>
      <strong>Cal DRE# 02044082 - Realtor&reg; | President's Circle |
      America's Top 1% Real Estate Professional |
      Voted the Best Commercial Realtor in San Mateo County</strong>
    </td></tr>
    <tr><td style="padding-top: 10px;">
      mobile: <a href="tel:6692267416" style="color: #222222;">669.226.7416</a><br>
      email: <a href="mailto:ravi@jagtianigroup.com"
                style="color: #1a73e8;">ravi@jagtianigroup.com</a><br>
      web: <a href="https://www.JagtianiGroup.com/commercial"
              style="color: #1a73e8;">www.JagtianiGroup.com/commercial</a>
    </td></tr>
    <tr><td style="padding-top: 10px; color: #555555; font-style: italic;">
      In the business of wealth creation
    </td></tr>
    <tr><td style="padding-top: 10px;">
      <a href="https://www.linkedin.com/in/ravijagtiani"
         style="color: #1a73e8;">Linkedin</a> &nbsp;|&nbsp;
      <a href="https://www.facebook.com/JagtianiGroup"
         style="color: #1a73e8;">Facebook Business Page</a> &nbsp;|&nbsp;
      <a href="https://www.youtube.com/@jagtianigroup"
         style="color: #1a73e8;">Youtube</a> &nbsp;|&nbsp;
      <a href="https://www.zillow.com/profile/ravijag"
         style="color: #1a73e8;">My Reviews</a> &nbsp;|&nbsp;
      <a href="https://blog.jagtianigroup.com"
         style="color: #1a73e8;">My blog</a>
    </td></tr>
    <tr><td style="padding-top: 12px; font-size: 12px; color: #777777;">
      Intero has been voted the &lsquo;Best Real Estate Company&rsquo; in the
      East Bay and Silicon Valley by the Bay Area News Group for
      2016, 2017, &amp; 2018!
    </td></tr>
  </table>

</body>
</html>"""


FALLBACK_TEMPLATE = """\
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #222222;
             line-height: 1.7; max-width: 680px; margin: 0; padding: 20px;">

  <p>Hi {sender_name},</p>

  <p>Thank you for your enquiry. Please find all our available listings here:<br>
  <a href="https://www.crexi.com/profile/ravi-jagtiani-ravijag">
  View all listings</a></p>

  <p>You can call me to know more:
  <a href="tel:{agent_phone_digits}">{agent_phone}</a></p>

  <p>Thanks,<br>
  {agent_name}</p>

  <br>
  <hr style="border: none; border-top: 1px solid #cccccc; margin: 24px 0;">

  <table cellpadding="0" cellspacing="0" border="0"
         style="font-family: Arial, sans-serif; font-size: 13px;
                color: #222222; line-height: 1.7;">
    <tr><td style="padding-bottom: 14px;">
      <img src="{signature_url}" width="600"
           alt="Ravi Jagtiani - Jagtiani Group"
           style="display: block; max-width: 100%;"/>
    </td></tr>
    <tr><td>
      <strong>Ravi R Jagtiani | Jagtiani Group | Managing Director</strong><br>
      <strong>Cal DRE# 02044082 - Realtor&reg; | President's Circle |
      America's Top 1% Real Estate Professional |
      Voted the Best Commercial Realtor in San Mateo County</strong>
    </td></tr>
    <tr><td style="padding-top: 10px;">
      mobile: <a href="tel:6692267416" style="color: #222222;">669.226.7416</a><br>
      email: <a href="mailto:ravi@jagtianigroup.com"
                style="color: #1a73e8;">ravi@jagtianigroup.com</a><br>
      web: <a href="https://www.JagtianiGroup.com/commercial"
              style="color: #1a73e8;">www.JagtianiGroup.com/commercial</a>
    </td></tr>
    <tr><td style="padding-top: 10px; color: #555555; font-style: italic;">
      In the business of wealth creation
    </td></tr>
    <tr><td style="padding-top: 10px;">
      <a href="https://www.linkedin.com/in/ravijagtiani"
         style="color: #1a73e8;">Linkedin</a> &nbsp;|&nbsp;
      <a href="https://www.facebook.com/JagtianiGroup"
         style="color: #1a73e8;">Facebook Business Page</a> &nbsp;|&nbsp;
      <a href="https://www.youtube.com/@jagtianigroup"
         style="color: #1a73e8;">Youtube</a> &nbsp;|&nbsp;
      <a href="https://www.zillow.com/profile/ravijag"
         style="color: #1a73e8;">My Reviews</a> &nbsp;|&nbsp;
      <a href="https://blog.jagtianigroup.com"
         style="color: #1a73e8;">My blog</a>
    </td></tr>
    <tr><td style="padding-top: 12px; font-size: 12px; color: #777777;">
      Intero has been voted the &lsquo;Best Real Estate Company&rsquo; in the
      East Bay and Silicon Valley by the Bay Area News Group for
      2016, 2017, &amp; 2018!
    </td></tr>
  </table>

</body>
</html>"""
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
    Required columns:
      property_type, keywords, listing_url, region,
      team_name, team_phone, team_email
    """
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
        log.info(f"Loaded {len(records)} listing row(s) from Google Sheet.")
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


def digits_only(phone: str) -> str:
    """Strip non-digit chars for use in tel: links."""
    return re.sub(r"\D", "", phone)


def extract_enquirer_email(body: str) -> str:
    """
    Handles all known notification formats + direct emails:
    1. LoopNet   — 'From: Name | phone | email | Listing ID'
    2. BizBuySell — 'Contact Email:' label then email on next line
    3. Crexi CA  — standalone email on its own line
    4. Same-line label — 'Email: address@email.com'
    5. Any email in body as last resort fallback
    """
    lines = [l.strip() for l in body.splitlines()]

    # Format 1 — LoopNet: "From: Name | phone | email | ..."
    loopnet = re.search(
        r'From:\s+[^|]+\|[^|]+\|\s*([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
        body, re.IGNORECASE
    )
    if loopnet:
        return loopnet.group(1).strip()

    # Format 2 — BizBuySell: "Contact Email:" on one line, email on the next
    for i, line in enumerate(lines):
        if re.search(r'contact\s*email', line, re.IGNORECASE):
            for next_line in lines[i+1:]:
                if next_line:
                    match = re.match(
                        r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$',
                        next_line
                    )
                    if match:
                        return next_line
                    break

    # Format 3 — Crexi CA: standalone email on its own line
    for line in lines:
        match = re.match(
            r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$',
            line
        )
        if match and line.lower() != EMAIL_ADDRESS.lower():
            return line

    # Format 4 — same-line label fallback
    labelled = re.search(
        r'(?:email|e-mail|reply.?to|contact)[^\n:]*[:\s]+'
        r'([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
        body, re.IGNORECASE
    )
    if labelled:
        return labelled.group(1).strip()

    # Format 5 — any email in body, skip Ravi's own
    all_emails = re.findall(
        r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',
        body
    )
    for addr in all_emails:
        if addr.lower() != EMAIL_ADDRESS.lower():
            return addr

    return ""


def extract_enquirer_name(body: str) -> str:
    """
    Handles all known notification formats:
    1. LoopNet   — 'From: Javier Feliciano | ...'
    2. BizBuySell — 'Contact Name:' then name on next line
    3. Crexi CA  — 'Harsh Garg has executed/requested...'
    4. Generic label fallback
    """
    lines = [l.strip() for l in body.splitlines()]

    # Format 1 — LoopNet: "From: Name | ..."
    loopnet = re.search(
        r'From:\s+([A-Z][a-zA-Z\s\-]{2,40}?)\s*\|',
        body, re.IGNORECASE
    )
    if loopnet:
        return loopnet.group(1).strip()

    # Format 2 — BizBuySell: "Contact Name:" on one line, name on the next
    for i, line in enumerate(lines):
        if re.search(r'contact\s*name', line, re.IGNORECASE):
            for next_line in lines[i+1:]:
                if next_line:
                    return next_line.strip()
            break

    # Format 3 — Crexi CA: "Name has executed/requested..."
    match = re.search(
        r'^([A-Z][a-zA-Z\s\-]{2,40}?)\s+has\s+(?:executed|requested|submitted)',
        body, re.MULTILINE
    )
    if match:
        return match.group(1).strip()

    # Format 4 — generic label fallback
    match = re.search(
        r'(?:full name|name|first name)[^\n:]*[:\s]+'
        r'([A-Za-z][A-Za-z\s\-]{1,40})',
        body, re.IGNORECASE
    )
    if match:
        name = match.group(1).strip()
        name = re.split(
            r'\n|last name|email|phone|message',
            name, flags=re.IGNORECASE
        )[0].strip()
        return name

    return ""


# ── INQUIRY DETECTION ─────────────────────────────────────────────────────────
def is_inquiry(subject: str, body: str) -> bool:
    text = (subject + " " + body).lower()
    return any(kw in text for kw in INQUIRY_KEYWORDS)


# ── PROPERTY TYPE + REGION MATCHING ──────────────────────────────────────────
def match_listing(subject: str, body: str, listings: list[dict]):
    """
    Two-stage matching:
    1. Filter rows where property type keywords match the email
    2. Among those, find the row whose region cities appear in the email
    3. If no city match, fall back to the last matching row (Ravi fallback)

    Returns (listing_row, matched_city) where matched_city is the exact
    city name extracted from the email to use in the reply, or the full
    region string if no specific city was detected.
    """
    text = (subject + " " + body).lower()

    # Stage 1 — collect all rows that match the property type keywords
    type_matches = []
    for listing in listings:
        keywords = [k.strip().lower() for k in listing.get("keywords", "").split(",")]
        if any(kw in text for kw in keywords if kw):
            type_matches.append(listing)

    if not type_matches:
        return None, ""

    # Stage 2 — scan region column of each matched row for city names
    for listing in type_matches:
        region_str = listing.get("region", "")
        cities = [c.strip() for c in region_str.split(",") if c.strip()]
        for city in cities:
            if city.lower() in text:
                log.info(f"  Region match: {city} → {listing.get('team_name')}")
                return listing, city  # return exact city for use in reply

    # Stage 3 — no city found, use last type match (should be Ravi fallback row)
    fallback = type_matches[-1]
    return fallback, fallback.get("region", "your area")


# ── EMAIL BUILDER ─────────────────────────────────────────────────────────────
def build_email_msg(
    to_email: str,
    cc_email: str,
    subject: str,
    html_body: str
) -> MIMEMultipart:
    # Always CC Ravi + the relevant team member
    # De-duplicate in case team member IS Ravi
    cc_addresses = list({a for a in [AGENT_EMAIL, cc_email] if a})
    cc_str = ", ".join(cc_addresses)

    msg = MIMEMultipart("alternative")
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = to_email
    msg["Cc"]      = cc_str
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    return msg


# ── SEND + SAVE ───────────────────────────────────────────────────────────────
def send_message(msg: MIMEMultipart, to_email: str, cc_email: str) -> bool:
    # Recipients = To + all CC addresses (Ravi always included)
    cc_list = [a.strip() for a in msg.get("Cc", "").split(",") if a.strip()]
    recipients = list({to_email} | set(cc_list))
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

        subject    = decode_str(msg.get("Subject", ""))
        raw_from   = decode_str(msg.get("From", ""))
        body       = get_body(msg)
        name, form_sender = parse_sender(raw_from)

        log.info(f"→ '{subject}' from {form_sender}")

        # Mark as read immediately — prevents double processing
        mail.store(eid, '+FLAGS', '\\Seen')

        if not is_inquiry(subject, body):
            log.info("  Not an enquiry, skipping.")
            skipped += 1
            continue

        # Known platform notification senders — never reply to these
        BLOCKED_SENDERS = [
            "support@crexi.com",
            "noreply@crexi.com",
            "notifications@crexi.com",
            "support@bizbuysell.com",
            "noreply@bizbuysell.com",
            "leads@bizbuysell.com",
            "support@loopnet.com",
            "noreply@loopnet.com",
            "leads@loopnet.com",
            "donotreply@loopnet.com",
        ]

        # Extract real enquirer from body, do NOT fall back to blocked senders
        enquirer_email = extract_enquirer_email(body)
        if not enquirer_email:
            if form_sender.lower() in [b.lower() for b in BLOCKED_SENDERS]:
                log.warning(f"  Could not extract enquirer email and sender is a platform notification ({form_sender}) — skipping to avoid account ban.")
                skipped += 1
                continue
            # Safe to use sender directly (genuine direct email)
            enquirer_email = form_sender

        enquirer_name = extract_enquirer_name(body) or name or "there"

        if not enquirer_email:
            log.warning("  No enquirer email found — skipping.")
            skipped += 1
            continue

        log.info(f"  Enquirer: {enquirer_name} <{enquirer_email}>")

        listing, matched_city = match_listing(subject, body, listings)
        reply_subj = (
            f"Re: {subject}"
            if not subject.lower().startswith("re:") else subject
        )

        if listing:
            log.info(f"  Matched: {listing['property_type']} — {matched_city} → {listing.get('team_name')}")
            cc_email  = listing.get("team_email", AGENT_EMAIL)
            html_body = EMAIL_TEMPLATE.format(
                sender_name       = enquirer_name,
                property_type     = listing["property_type"],
                listing_url       = listing["listing_url"],
                region            = matched_city,
                team_name         = listing.get("team_name", AGENT_NAME),
                team_phone        = listing.get("team_phone", AGENT_PHONE),
                team_phone_digits = digits_only(listing.get("team_phone", AGENT_PHONE)),
                agent_name        = AGENT_NAME,
                agent_phone       = AGENT_PHONE,
                agent_phone_digits= digits_only(AGENT_PHONE),
                signature_url     = SIGNATURE_URL
            )
        else:
            log.info("  No type matched — sending fallback.")
            cc_email  = AGENT_EMAIL
            html_body = FALLBACK_TEMPLATE.format(
                sender_name       = enquirer_name,
                agent_name        = AGENT_NAME,
                agent_phone       = AGENT_PHONE,
                agent_phone_digits= digits_only(AGENT_PHONE),
                signature_url     = SIGNATURE_URL
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
