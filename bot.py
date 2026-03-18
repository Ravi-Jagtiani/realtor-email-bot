"""
Realtor Email Auto-Reply Bot
────────────────────────────
Designed to run once and exit (for GitHub Actions / cron scheduling).
- Scrapes the realtor's Crexi profile for live listings
- Checks inbox for unread property inquiry emails
- Matches inquiry to listing, fills template, sends reply
- Logs everything to replies.log
"""

import imaplib
import smtplib
import email
import re
import os
import logging
from datetime import datetime
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── CONFIG — all values come from environment variables (GitHub Secrets) ──────
IMAP_SERVER       = os.environ.get("IMAP_SERVER",   "imap.mail.yahoo.com")
IMAP_PORT         = int(os.environ.get("IMAP_PORT", "993"))
SMTP_SERVER       = os.environ.get("SMTP_SERVER",   "smtp.mail.yahoo.com")
SMTP_PORT         = int(os.environ.get("SMTP_PORT", "465"))
EMAIL_ADDRESS     = os.environ.get("EMAIL_ADDRESS",  "")
EMAIL_PASSWORD    = os.environ.get("EMAIL_PASSWORD", "")
CREXI_PROFILE_URL = os.environ.get("CREXI_PROFILE_URL", "https://www.crexi.com/profile/ravi-jagtiani-ravijag")
AGENT_NAME        = os.environ.get("AGENT_NAME",    "Ravi Jagtiani")
AGENT_EMAIL       = os.environ.get("AGENT_EMAIL",   "")
BROKERAGE_NAME    = os.environ.get("BROKERAGE_NAME","")
# ─────────────────────────────────────────────────────────────────────────────

# ── KEYWORDS that flag an email as a property inquiry ────────────────────────
INQUIRY_KEYWORDS = [
    "interested in", "inquiry", "enquiry", "open house",
    "listing", "property", "schedule a tour", "more info",
    "available", "asking price", "for sale", "for lease",
    "can you send", "would like to know", "tell me more",
    "viewing", "visit", "showing"
]
# ─────────────────────────────────────────────────────────────────────────────

# ── EMAIL REPLY TEMPLATE ──────────────────────────────────────────────────────
# Replace this entire block with the real template when available.
# Placeholders: {sender_name}, {property_address}, {listing_url},
#               {agent_name}, {agent_email}, {brokerage_name}
EMAIL_TEMPLATE = """Hi {sender_name},

Thank you for reaching out! We received your inquiry about {property_address}.

You can view the full listing details here:
{listing_url}

{agent_name} is handling this property and will be in touch with you shortly.
In the meantime, feel free to reach them directly at {agent_email}.

Best regards,
{agent_name}
{brokerage_name}
"""
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


# ── CREXI SCRAPER ─────────────────────────────────────────────────────────────
def scrape_crexi_listings() -> list[dict]:
    """Scrape the realtor's active listings from their Crexi profile page."""
    try:
        from playwright.sync_api import sync_playwright

        listings = []
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ]
            )
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/122.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 800},
                locale="en-US",
                timezone_id="America/New_York",
            )
            # Remove headless detection flags
            context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3]});
                Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
            """)
            page = context.new_page()
            log.info(f"Scraping Crexi profile: {CREXI_PROFILE_URL}")
            page.goto(CREXI_PROFILE_URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(5000)

            # Debug — log all hrefs found on page so we can find the right selector
            all_links = page.query_selector_all("a[href]")
            hrefs = [a.get_attribute("href") for a in all_links if a.get_attribute("href")]
            log.info(f"All hrefs found on page ({len(hrefs)} total):")
            for h in hrefs[:50]:  # first 50 to avoid log flood
                log.info(f"  {h}")

            # Grab all links pointing to /properties/ pages
            cards = page.query_selector_all("a[href*='/properties/']")
            seen  = set()

            for card in cards:
                try:
                    href = card.get_attribute("href") or ""
                    if not href or href in seen:
                        continue
                    seen.add(href)

                    full_url = (
                        f"https://www.crexi.com{href}"
                        if href.startswith("/") else href
                    )
                    text  = card.inner_text().strip()
                    lines = [l.strip() for l in text.split("\n") if l.strip()]
                    address = (
                        lines[0] if lines
                        else href.split("/")[-1].replace("-", " ").title()
                    )

                    listings.append({
                        "property_address": address,
                        "agent_name":       AGENT_NAME,
                        "agent_email":      AGENT_EMAIL or EMAIL_ADDRESS,
                        "listing_url":      full_url
                    })
                except Exception:
                    continue

            browser.close()

        log.info(f"Found {len(listings)} listing(s) on Crexi.")
        return listings

    except Exception as e:
        log.error(f"Crexi scrape error: {e}")
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


# ── LISTING MATCH ─────────────────────────────────────────────────────────────
def match_listing(subject: str, body: str, listings: list[dict]):
    text       = (subject + " " + body).lower()
    best, best_score = None, 0

    for listing in listings:
        words = [w for w in re.split(r"[\s,]+", listing["property_address"].lower()) if len(w) > 2]
        score = sum(1 for w in words if w in text)
        if score > best_score:
            best_score = score
            best       = listing

    return best if best_score >= 2 else None


# ── SEND REPLY ────────────────────────────────────────────────────────────────
def send_reply(to_email: str, sender_name: str, subject: str, listing: dict) -> bool:
    body = EMAIL_TEMPLATE.format(
        sender_name      = sender_name or "there",
        property_address = listing["property_address"],
        listing_url      = listing["listing_url"],
        agent_name       = listing["agent_name"],
        agent_email      = listing["agent_email"],
        brokerage_name   = BROKERAGE_NAME
    )
    msg = MIMEMultipart()
    msg["From"]    = EMAIL_ADDRESS
    msg["To"]      = to_email
    msg["Subject"] = f"Re: {subject}" if not subject.lower().startswith("re:") else subject
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        log.info(f"✓ Reply sent to {to_email} — {listing['property_address']}")
        return True
    except Exception as e:
        log.error(f"Send failed: {e}")
        return False


# ── MAIN (runs once and exits) ────────────────────────────────────────────────
def main():
    log.info("=" * 50)
    log.info(f"Bot run started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("=" * 50)

    # 1. Get listings from Crexi
    listings = scrape_crexi_listings()
    if not listings:
        log.warning("No listings found — cannot match inquiries. Exiting.")
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

        subject     = decode_str(msg.get("Subject", ""))
        raw_from    = decode_str(msg.get("From",    ""))
        body        = get_body(msg)
        name, addr  = parse_sender(raw_from)

        log.info(f"→ '{subject}' from {addr}")

        if not is_inquiry(subject, body):
            log.info("  Not an inquiry, skipping.")
            skipped += 1
            continue

        listing = match_listing(subject, body, listings)
        if not listing:
            log.info("  No listing match found, skipping.")
            skipped += 1
            continue

        log.info(f"  Matched: {listing['property_address']}")
        if send_reply(addr, name, subject, listing):
            replied += 1

    mail.logout()

    log.info("-" * 50)
    log.info(f"Done. Replied: {replied} | Skipped: {skipped}")
    log.info("=" * 50)


if __name__ == "__main__":
    main()
