#!/usr/bin/env python3
"""Serveur local pour l'app Todo & Mail — sauvegarde dans data.json."""

import base64
import csv
import email as email_lib
import email.policy
import hashlib
import http.server
import imaplib
import json
import logging
import mailbox
import os
import poplib
import re
import secrets
import smtplib
import socket
import ssl
import subprocess
import time
import urllib.request
import urllib.parse
import urllib.error
from urllib.parse import parse_qs, urlparse
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from email import encoders
from email import policy as email_policy
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import getaddresses, parsedate_to_datetime

from app_config import (
    CONTACTS_CSV,
    DATA,
    DIR,
    DOWNLOADS,
    GOOGLE_CALENDAR_SCOPE,
    GOOGLE_MAIL_SCOPE,
    INBOX_INDEX_FILE,
    LOG_FILE,
    MAILS_DIR,
    OBSIDIAN_ATT_DIR,
    OBSIDIAN_MD_DIR,
    OBSIDIAN_VAULT,
    PORT,
    PROJECT_ROOT,
    RENDERER_INDEX,
    SEEN_UIDS_FILE,
)
from account_store import (
    find_account_by_email,
    find_account_index_by_email,
    load_accounts,
    normalize_auth_fields,
    save_accounts,
)
from calendar_routes import (
    handle_calendar_accounts_get,
    handle_calendar_calendars_get,
    handle_calendar_event_create_post,
    handle_calendar_event_delete_post,
    handle_calendar_event_update_post,
    handle_calendar_events_get,
    handle_oauth_callback,
)
from google_calendar_service import (
    build_calendar_http_error_response,
    build_oauth_callback_page,
    create_google_calendar_event,
    delete_google_calendar_event,
    exchange_google_auth_code,
    generate_pkce_pair,
    get_valid_gmail_access_token,
    get_google_oauth_accounts,
    list_google_calendars,
    list_google_calendar_events,
    parse_google_error_payload,
    pick_google_oauth_account,
    update_google_calendar_event,
)
from mail_service import (
    build_xoauth2_string as _build_xoauth2_string_impl,
    delete_mail_on_server as _delete_mail_on_server_impl,
    fetch_imap as _fetch_imap_impl,
    fetch_pop3 as _fetch_pop3_impl,
    send_email_smtp as _send_email_smtp_impl,
    smtp_auth_xoauth2 as _smtp_auth_xoauth2_impl,
)
from json_store import atomic_write_json, read_json_with_backup

try:
    import html2text
    HAS_HTML2TEXT = True
except ImportError:
    HAS_HTML2TEXT = False

# In-memory OAuth state store (state -> metadata) for current server process.
GOOGLE_OAUTH_PENDING = {}

logging.basicConfig(
    filename=LOG_FILE,
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.ERROR,
)
logger = logging.getLogger("todoapp")


def loadAppState():
    return read_json_with_backup(DATA, {"sections": [], "settings": {}})


def saveAppState(data):
    atomic_write_json(DATA, data)


def loadContactsData():
    contacts = []
    try:
        with open(CONTACTS_CSV, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row.get("Display Name", "").strip()
                if not name:
                    first = row.get("First Name", "").strip()
                    last = row.get("Last Name", "").strip()
                    name = f"{first} {last}".strip()
                email_addr = row.get("Primary Email", "").strip()
                if email_addr:
                    contacts.append({"name": name, "email": email_addr})
    except Exception:
        pass
    return contacts


def ai_call(token, prompt):
    """Generic AI call via Google Gemini API — single request."""
    body = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.3},
    }).encode()

    url = ("https://generativelanguage.googleapis.com/v1beta/"
           f"models/gemma-3-27b-it:generateContent?key={token}")

    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            result = json.loads(r.read())
        return result["candidates"][0]["content"]["parts"][0]["text"]
    except urllib.error.HTTPError as e:
        error_body = e.read().decode()
        try:
            msg = json.loads(error_body).get("error", {}).get("message", error_body)
        except Exception:
            msg = error_body
        logger.error("Gemini API %d: %s", e.code, msg)
        raise RuntimeError(f"Gemini {e.code}: {msg}")
    except Exception as e:
        logger.error("Gemini API error: %s", e)
        raise


def ai_reformulate(payload):
    token = payload.get("token", "")
    text = payload.get("text", "")
    prompt = (
        "Corriges la syntaxe, la grammaire et l'orthographe du texte suivant. "
        "Réponds UNIQUEMENT avec le texte corrigé, sans commentaire ni explication :\n\n"
        + text
    )
    return ai_call(token, prompt)


def ai_generate_reminder(payload):
    token = payload.get("token", "")
    original_subject = payload.get("subject", "")
    original_to = payload.get("to", "")
    original_body = payload.get("body", "")
    prompt = (
        "Tu es un assistant professionnel. Il y a 3 jours j'ai envoyé un mail et je n'ai pas reçu de réponse. "
        "Génère un mail de relance poli et professionnel en français. "
        "Réponds UNIQUEMENT en JSON valide (sans balises markdown) avec cette structure :\n"
        '{"subject":"...","body":"..."}\n\n'
        f"Mail original :\n"
        f"À : {original_to}\n"
        f"Sujet : {original_subject}\n"
        f"Corps :\n{original_body}"
    )
    content = ai_call(token, prompt)
    if "```" in content:
        content = content.split("```json")[-1] if "```json" in content else content.split("```")[1]
        content = content.split("```")[0]
    return json.loads(content.strip())


def ai_generate_reply(payload):
    token = payload.get("token", "")
    user_prompt = payload.get("prompt", "")
    subject = payload.get("subject", "")
    sender = payload.get("from", "")
    original_text = payload.get("original_text", "")
    draft = payload.get("draft", "")

    prompt = (
        "Tu es un assistant de redaction email professionnel en francais. "
        "Genere UNIQUEMENT le texte de reponse (sans objet, sans salutation imposee, sans commentaire). "
        "Respecte strictement les instructions utilisateur ci-dessous. "
        "N'inclus pas le message original dans la sortie.\n\n"
        f"Sujet du fil : {subject}\n"
        f"Expediteur original : {sender}\n\n"
        "Instructions utilisateur :\n"
        f"{user_prompt}\n\n"
        "Brouillon actuel (a ameliorer si present) :\n"
        f"{draft}\n\n"
        "Message original recu (contexte, NE PAS recopier integralement) :\n"
        f"{original_text}"
    )
    return ai_call(token, prompt)


def build_eml(from_addr, to_addr, subject, body_text, html_body=None):
    if html_body:
        msg = MIMEMultipart("alternative")
        msg.attach(MIMEText(body_text, "plain", "utf-8"))
        msg.attach(MIMEText(html_body, "html", "utf-8"))
    else:
        msg = MIMEText(body_text, "plain", "utf-8")
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0100")
    return msg.as_string()


def save_eml_to_downloads(from_addr, to_addr, subject, body_text, html_body=None):
    eml_content = build_eml(from_addr, to_addr, subject, body_text, html_body=html_body)
    safe_subject = "".join(c for c in subject if c.isalnum() or c in " _-").strip()[:80] or "mail"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{safe_subject}_{ts}.eml"
    filepath = os.path.join(DOWNLOADS, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(eml_content)
    return filepath


def build_xoauth2_string(username, access_token):
    return _build_xoauth2_string_impl(username, access_token)


def smtp_auth_xoauth2(server, username, access_token):
    return _smtp_auth_xoauth2_impl(server, username, access_token)


# ═══════════════════════════════════════════════════════
#  Seen UIDs — deduplication
# ═══════════════════════════════════════════════════════
def load_seen_uids():
    return read_json_with_backup(SEEN_UIDS_FILE, {})


def save_seen_uids(seen):
    atomic_write_json(SEEN_UIDS_FILE, seen)


# ═══════════════════════════════════════════════════════
#  Inbox Index — local mail metadata
# ═══════════════════════════════════════════════════════
def load_inbox_index():
    index = read_json_with_backup(INBOX_INDEX_FILE, [])
    if not isinstance(index, list):
        return []

    # Drop entries pointing to missing .eml files to avoid stale inbox rows.
    filtered = []
    changed = False
    for m in index:
        eml_file = m.get("eml_file", "")
        if eml_file:
            eml_path = os.path.join(MAILS_DIR, eml_file)
            if not os.path.isfile(eml_path):
                changed = True
                continue
        filtered.append(m)

    if changed:
        save_inbox_index(filtered)
    return filtered


def save_inbox_index(index):
    atomic_write_json(INBOX_INDEX_FILE, index)


def compute_mail_id(raw_bytes):
    """Compute a stable hash for deduplication."""
    return hashlib.sha256(raw_bytes).hexdigest()[:24]


def clean_string_for_file(name):
    if not name:
        return ""
    name = str(name).replace('\n', ' ').replace('\r', '')
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


def unique_eml_filename_from_subject(subject, prefix=""):
    """Build a unique .eml filename from subject with _1, _2... suffixes."""
    safe_subject = clean_string_for_file(subject)[:120] or "mail"
    if prefix:
        safe_subject = f"{prefix}{safe_subject}"

    candidate = f"{safe_subject}.eml"
    index = 1
    while os.path.exists(os.path.join(MAILS_DIR, candidate)):
        candidate = f"{safe_subject}_{index}.eml"
        index += 1
    return candidate


def extract_bodies(msg):
    """Extract plain-text and HTML bodies from a parsed email message."""
    body_text = ""
    body_html = ""
    h = None
    if HAS_HTML2TEXT:
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.body_width = 0

    for part in msg.walk():
        content_type = part.get_content_type()
        content_disposition = str(part.get("Content-Disposition", ""))

        if "attachment" in content_disposition:
            continue

        if content_type == "text/plain":
            if not body_text:
                try:
                    charset = part.get_content_charset('utf-8') or 'utf-8'
                    body_text = part.get_payload(decode=True).decode(charset, errors='replace')
                except Exception:
                    pass
        elif content_type == "text/html":
            try:
                charset = part.get_content_charset('utf-8') or 'utf-8'
                body_html = part.get_payload(decode=True).decode(charset, errors='replace')
            except Exception:
                pass

    if not body_text and body_html and h:
        try:
            body_text = h.handle(body_html)
        except Exception:
            body_text = ""

    return body_text, body_html


def get_attachment_payload(msg, index=None, filename=None):
    """Return (bytes, filename, content_type) for an attachment by index or filename."""
    found_idx = 0
    for part in msg.walk():
        content_disposition = str(part.get("Content-Disposition", ""))
        part_filename = part.get_filename()
        if not (("attachment" in content_disposition or part_filename) and part_filename):
            continue

        if index is not None:
            if found_idx != index:
                found_idx += 1
                continue
        elif filename is not None and part_filename != filename:
            found_idx += 1
            continue

        payload = part.get_payload(decode=True)
        if payload is None:
            return None, None, None
        return payload, part_filename, part.get_content_type() or "application/octet-stream"

    return None, None, None


def enrich_mail_from_eml(mail):
    """Populate body/body_html/attachments from local .eml when available."""
    eml_file = mail.get("eml_file", "")
    if not eml_file:
        return mail
    eml_path = os.path.join(MAILS_DIR, eml_file)
    if not os.path.isfile(eml_path):
        return mail

    try:
        with open(eml_path, "rb") as f:
            raw_bytes = f.read()
        parsed = parse_email_metadata(raw_bytes, mail.get("account", ""))
        mail["body"] = parsed.get("body", mail.get("body", ""))
        mail["body_html"] = parsed.get("body_html", "")
        mail["attachments"] = parsed.get("attachments", mail.get("attachments", []))
    except Exception:
        pass

    return mail


def extract_attachments_info(msg):
    """Return list of attachment filenames from a message."""
    attachments = []
    for part in msg.walk():
        content_disposition = str(part.get("Content-Disposition", ""))
        filename = part.get_filename()
        if ("attachment" in content_disposition or filename) and filename:
            attachments.append(filename)
    return attachments


def parse_email_metadata(raw_bytes, account_email=""):
    """Parse raw email bytes into metadata dict."""
    msg = email_lib.message_from_bytes(raw_bytes, policy=email_policy.default)

    subject = msg.get('Subject', 'Sans sujet') or 'Sans sujet'
    from_hdr = msg.get('From', '')
    to_hdr = msg.get('To', '')
    cc_hdr = msg.get('Cc', '')
    date_str = msg.get('Date', '')
    message_id = msg.get('Message-ID', '') or ''

    # Parse sender
    from_addrs = getaddresses([from_hdr])
    sender_name = ''
    sender_email = ''
    if from_addrs:
        sender_name = from_addrs[0][0] or ''
        sender_email = from_addrs[0][1] or ''

    # Parse date
    date_ts = 0
    date_display = date_str
    try:
        dt = parsedate_to_datetime(date_str)
        date_ts = int(dt.timestamp() * 1000)
        date_display = dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        date_ts = int(time.time() * 1000)
        date_display = datetime.now().strftime("%Y-%m-%d %H:%M")

    body_text, body_html = extract_bodies(msg)
    attachments = extract_attachments_info(msg)

    return {
        "subject": subject,
        "from_name": sender_name,
        "from_email": sender_email,
        "to": to_hdr,
        "cc": cc_hdr,
        "date": date_display,
        "date_ts": date_ts,
        "message_id": message_id,
        "body": body_text,
        "body_html": body_html,
        "attachments": attachments,
        "account": account_email,
    }


# ═══════════════════════════════════════════════════════
#  POP3 Fetch
# ═══════════════════════════════════════════════════════
def fetch_pop3(account):
    return _fetch_pop3_impl(
        account,
        load_seen_uids=load_seen_uids,
        save_seen_uids=save_seen_uids,
        load_inbox_index=load_inbox_index,
        save_inbox_index=save_inbox_index,
        compute_mail_id=compute_mail_id,
        parse_email_metadata=parse_email_metadata,
        unique_eml_filename_from_subject=unique_eml_filename_from_subject,
        mails_dir=MAILS_DIR,
    )


# ═══════════════════════════════════════════════════════
#  IMAP Fetch
# ═══════════════════════════════════════════════════════
def fetch_imap(account):
    return _fetch_imap_impl(
        account,
        normalize_auth_fields=normalize_auth_fields,
        get_valid_gmail_access_token=get_valid_gmail_access_token,
        load_seen_uids=load_seen_uids,
        save_seen_uids=save_seen_uids,
        load_inbox_index=load_inbox_index,
        save_inbox_index=save_inbox_index,
        compute_mail_id=compute_mail_id,
        parse_email_metadata=parse_email_metadata,
        unique_eml_filename_from_subject=unique_eml_filename_from_subject,
        mails_dir=MAILS_DIR,
    )


def fetch_all_accounts():
    """Fetch from all configured accounts (POP3 or IMAP)."""
    accounts = load_accounts()
    total_new = 0
    all_errors = []

    for acc in accounts:
        if not acc.get("enabled", True):
            continue
        try:
            protocol = acc.get("protocol", "pop3").lower()
            if protocol == "imap":
                n, errs = _fetch_imap_impl(
                    acc,
                    normalize_auth_fields=normalize_auth_fields,
                    get_valid_gmail_access_token=get_valid_gmail_access_token,
                    load_seen_uids=load_seen_uids,
                    save_seen_uids=save_seen_uids,
                    load_inbox_index=load_inbox_index,
                    save_inbox_index=save_inbox_index,
                    compute_mail_id=compute_mail_id,
                    parse_email_metadata=parse_email_metadata,
                    unique_eml_filename_from_subject=unique_eml_filename_from_subject,
                    mails_dir=MAILS_DIR,
                )
            else:
                n, errs = _fetch_pop3_impl(
                    acc,
                    load_seen_uids=load_seen_uids,
                    save_seen_uids=save_seen_uids,
                    load_inbox_index=load_inbox_index,
                    save_inbox_index=save_inbox_index,
                    compute_mail_id=compute_mail_id,
                    parse_email_metadata=parse_email_metadata,
                    unique_eml_filename_from_subject=unique_eml_filename_from_subject,
                    mails_dir=MAILS_DIR,
                )
            total_new += n
            all_errors.extend(errs)
        except Exception as e:
            all_errors.append(f"{acc.get('email', '?')}: {e}")

    return total_new, all_errors


# ═══════════════════════════════════════════════════════
#  Email Autoconfig (Mozilla Thunderbird database)
# ═══════════════════════════════════════════════════════
def autoconfig_email(email_addr):
    """Auto-detect IMAP/SMTP settings from Mozilla's autoconfig database."""
    domain = email_addr.strip().split("@")[-1].lower()

    config = None
    # Try Mozilla autoconfig
    url = f"https://autoconfig.thunderbird.net/v1.1/{domain}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "ISENAPP/1.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            xml_data = resp.read()
        config = _parse_autoconfig_xml(xml_data, email_addr)
    except Exception:
        config = None

    if config:
        return config

    # Fallback: probe common hostnames
    return _autoconfig_fallback(domain, email_addr)


def _parse_autoconfig_xml(xml_data, email_addr):
    """Parse Mozilla autoconfig XML and return structured config dict."""
    root = ET.fromstring(xml_data)
    ns = ''
    # Handle potential namespace
    if root.tag.startswith('{'):
        ns = root.tag.split('}')[0] + '}'

    result = {"imap": None, "smtp": None, "source": "mozilla"}

    for provider in root.iter(f"{ns}emailProvider"):
        # Find IMAP
        for inc in provider.iter(f"{ns}incomingServer"):
            if inc.get("type") == "imap":
                hostname = (inc.findtext(f"{ns}hostname") or "").strip()
                port = int(inc.findtext(f"{ns}port") or "993")
                socket_type = (inc.findtext(f"{ns}socketType") or "SSL").strip()
                username_tpl = (inc.findtext(f"{ns}username") or "%EMAILADDRESS%").strip()
                username = username_tpl.replace("%EMAILADDRESS%", email_addr).replace("%EMAILLOCALPART%", email_addr.split("@")[0])
                result["imap"] = {
                    "server": hostname, "port": port,
                    "ssl": socket_type in ("SSL", "STARTTLS"),
                    "socket_type": socket_type, "username": username
                }
                break

        # Find SMTP
        for out in provider.iter(f"{ns}outgoingServer"):
            if out.get("type") == "smtp":
                hostname = (out.findtext(f"{ns}hostname") or "").strip()
                port = int(out.findtext(f"{ns}port") or "587")
                socket_type = (out.findtext(f"{ns}socketType") or "STARTTLS").strip()
                username_tpl = (out.findtext(f"{ns}username") or "%EMAILADDRESS%").strip()
                username = username_tpl.replace("%EMAILADDRESS%", email_addr).replace("%EMAILLOCALPART%", email_addr.split("@")[0])
                result["smtp"] = {
                    "server": hostname, "port": port,
                    "ssl": socket_type == "SSL",
                    "starttls": socket_type == "STARTTLS",
                    "socket_type": socket_type, "username": username
                }
                break

    if result["imap"] or result["smtp"]:
        return result
    return None


def _autoconfig_fallback(domain, email_addr):
    """Fallback: test common IMAP/SMTP hostnames and ports."""
    result = {"imap": None, "smtp": None, "source": "fallback"}

    # Try IMAP
    for host in [f"imap.{domain}", f"mail.{domain}"]:
        for port, use_ssl in [(993, True), (143, False)]:
            try:
                if use_ssl:
                    conn = imaplib.IMAP4_SSL(host, port, timeout=5)
                else:
                    conn = imaplib.IMAP4(host, port)
                    conn.socket().settimeout(5)
                conn.logout()
                result["imap"] = {
                    "server": host, "port": port, "ssl": use_ssl,
                    "socket_type": "SSL" if use_ssl else "plain",
                    "username": email_addr
                }
                break
            except Exception:
                continue
        if result["imap"]:
            break

    # Try SMTP
    for host in [f"smtp.{domain}", f"mail.{domain}"]:
        for port, use_ssl, use_starttls in [(465, True, False), (587, False, True), (25, False, False)]:
            try:
                if use_ssl:
                    srv = smtplib.SMTP_SSL(host, port, timeout=5)
                else:
                    srv = smtplib.SMTP(host, port, timeout=5)
                    if use_starttls:
                        srv.starttls()
                srv.quit()
                result["smtp"] = {
                    "server": host, "port": port, "ssl": use_ssl,
                    "starttls": use_starttls,
                    "socket_type": "SSL" if use_ssl else ("STARTTLS" if use_starttls else "plain"),
                    "username": email_addr
                }
                break
            except Exception:
                continue
        if result["smtp"]:
            break

    if result["imap"] or result["smtp"]:
        return result
    return None


# ═══════════════════════════════════════════════════════
#  SMTP Send
# ═══════════════════════════════════════════════════════
def send_email_smtp(account, to_addr, subject, body_text, cc="", attachments=None, html_body=None):
    return _send_email_smtp_impl(
        account,
        to_addr,
        subject,
        body_text,
        cc=cc,
        attachments=attachments,
        html_body=html_body,
        normalize_auth_fields=normalize_auth_fields,
        get_valid_gmail_access_token=get_valid_gmail_access_token,
        compute_mail_id=compute_mail_id,
        unique_eml_filename_from_subject=unique_eml_filename_from_subject,
        parse_email_metadata=parse_email_metadata,
        load_inbox_index=load_inbox_index,
        save_inbox_index=save_inbox_index,
        mails_dir=MAILS_DIR,
    )


# ═══════════════════════════════════════════════════════
#  Delete mail from POP3 server
# ═══════════════════════════════════════════════════════
def delete_mail_on_server(account, uid_to_delete):
    return _delete_mail_on_server_impl(
        account,
        uid_to_delete,
        normalize_auth_fields=normalize_auth_fields,
        get_valid_gmail_access_token=get_valid_gmail_access_token,
    )


# ═══════════════════════════════════════════════════════
#  Obsidian Export (v3.py logic)
# ═══════════════════════════════════════════════════════
MOTS_CLES = ['projet', 'stage', 'facture', 'urgent', 'réunion', 'candidature', 'rapport', 'admin', 'examen']

WIKILINK_RE = re.compile(r'\[\[([^\]|]+?)(?:\|[^\]]*)?\]\]')


def scan_vault_graph():
    """Scan Obsidian vault, extract nodes (md files + attachments) and edges (wikilinks)."""
    vault = OBSIDIAN_VAULT
    nodes = {}  # name -> {id, label, path, type, tags, group}
    edges = []  # [{source, target}]

    # Collect all files
    for root, dirs, files in os.walk(vault):
        # Skip .obsidian config
        dirs[:] = [d for d in dirs if d != '.obsidian']
        for fname in files:
            fpath = os.path.join(root, fname)
            relpath = os.path.relpath(fpath, vault)
            name_no_ext = os.path.splitext(fname)[0]

            if fname.lower().endswith('.md'):
                # Parse frontmatter for tags
                tags = []
                try:
                    with open(fpath, 'r', encoding='utf-8', errors='replace') as f:
                        content = f.read(4096)  # Read just enough for frontmatter
                    if content.startswith('---'):
                        end = content.find('---', 3)
                        if end != -1:
                            fm = content[3:end]
                            for line in fm.split('\n'):
                                line = line.strip()
                                if line.startswith('- '):
                                    tags.append(line[2:].strip())
                except Exception:
                    pass

                # Determine group from path
                group = 'mail' if '/mails/' in relpath or relpath.startswith('mails/') else 'note'

                nodes[name_no_ext] = {
                    'id': name_no_ext,
                    'label': name_no_ext,
                    'path': relpath,
                    'type': 'md',
                    'tags': tags,
                    'group': group,
                }
            else:
                # Attachment (pdf, jpg, etc.)
                ext = os.path.splitext(fname)[1].lower()
                if ext in ('.png', '.jpg', '.jpeg', '.gif', '.svg', '.pdf',
                           '.docx', '.xlsx', '.pptx', '.odt', '.csv', '.zip'):
                    nodes[fname] = {
                        'id': fname,
                        'label': fname,
                        'path': relpath,
                        'type': 'attachment',
                        'tags': [],
                        'group': 'attachment',
                    }

    # Extract edges from wikilinks in md files
    for name, node in list(nodes.items()):
        if node['type'] != 'md':
            continue
        fpath = os.path.join(vault, node['path'])
        try:
            with open(fpath, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            links = WIKILINK_RE.findall(content)
            for link in links:
                link = link.strip()
                if link in nodes:
                    edges.append({'source': name, 'target': link})
                # Also try with known extensions for attachments
                elif link + '.md' in nodes:
                    pass  # wikilinks usually reference without .md
                else:
                    # Target might not exist yet — create an "orphan" node
                    if link not in nodes:
                        nodes[link] = {
                            'id': link,
                            'label': link,
                            'path': '',
                            'type': 'orphan',
                            'tags': [],
                            'group': 'orphan',
                        }
                    edges.append({'source': name, 'target': link})
        except Exception:
            pass

    return {'nodes': list(nodes.values()), 'edges': edges}


def read_vault_file(relpath):
    """Read a file from the Obsidian vault by relative path."""
    # Sanitize: prevent directory traversal
    safe = os.path.normpath(relpath)
    if safe.startswith('..') or os.path.isabs(safe):
        raise ValueError('Invalid path')
    fpath = os.path.join(OBSIDIAN_VAULT, safe)
    if not fpath.startswith(OBSIDIAN_VAULT):
        raise ValueError('Path outside vault')
    if not os.path.isfile(fpath):
        raise FileNotFoundError('File not found')
    with open(fpath, 'r', encoding='utf-8', errors='replace') as f:
        return f.read()


def export_email_to_obsidian(mail_meta):
    """Export a single email to Obsidian markdown, replicating v3.py logic."""
    os.makedirs(OBSIDIAN_MD_DIR, exist_ok=True)
    os.makedirs(OBSIDIAN_ATT_DIR, exist_ok=True)

    eml_path = os.path.join(MAILS_DIR, mail_meta.get("eml_file", ""))
    if not os.path.isfile(eml_path):
        raise FileNotFoundError(f"Fichier .eml introuvable: {eml_path}")

    with open(eml_path, "rb") as f:
        raw_bytes = f.read()

    msg = email_lib.message_from_bytes(raw_bytes, policy=email_policy.default)

    subject = msg.get('Subject', 'Sans_Sujet') or 'Sans_Sujet'
    from_hdr = msg.get('From', '')
    to_hdr = msg.get('To', '')
    cc_hdr = msg.get('Cc', '')
    date_str = msg.get('Date', '')

    # Clean subject — move RE/FW prefixes to end
    subject_clean = subject
    prefixes = []
    prefix_pattern = r'^(\s*(re|fw|fwd)\s*[:：\-]+)'
    while True:
        m = re.match(prefix_pattern, subject_clean, re.IGNORECASE)
        if m:
            prefixes.append(m.group(1).strip())
            subject_clean = subject_clean[m.end():].lstrip()
        else:
            break
    subject_final = f"{subject_clean} ({' '.join(prefixes)})" if prefixes else subject_clean
    safe_subject = clean_string_for_file(subject_final) or "Sans_Sujet"

    # Parse addresses
    def parse_addresses_list(header_value):
        if not header_value:
            return []
        addresses = getaddresses([header_value])
        results = []
        for name, addr in addresses:
            results.append(clean_string_for_file(name) if name else clean_string_for_file(addr))
        return results

    sender_list = parse_addresses_list(from_hdr)
    sender_name = sender_list[0] if sender_list else 'Inconnu'
    raw_sender = getaddresses([from_hdr])
    sender_domain = ""
    if raw_sender and raw_sender[0][1] and '@' in raw_sender[0][1]:
        sender_domain = raw_sender[0][1].split('@')[-1].lower()

    to_list = parse_addresses_list(to_hdr)
    cc_list = parse_addresses_list(cc_hdr)

    # Date parsing
    daily_note_link = ""
    year_month_tag = ""
    mois_fr = ["janvier", "février", "mars", "avril", "mai", "juin",
               "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
    try:
        dt = parsedate_to_datetime(date_str)
        daily_note_link = dt.strftime("%Y-%m-%d")
        mois = mois_fr[dt.month - 1]
        year_month_tag = f"{mois}-{dt.year}"
        file_time = dt.strftime("%Y-%m-%d_%H%M%S")
    except Exception:
        file_time = datetime.now().strftime("%Y-%m-%d_%H%M%S")

    # Filename
    base_md_filename = safe_subject[:100]
    md_filename = f"{base_md_filename}.md"
    md_filepath = os.path.join(OBSIDIAN_MD_DIR, md_filename)
    r_idx = 1
    while os.path.exists(md_filepath):
        md_filename = f"{base_md_filename}_r{r_idx}.md"
        md_filepath = os.path.join(OBSIDIAN_MD_DIR, md_filename)
        r_idx += 1

    # Tags
    tags = ["email"]
    if sender_domain:
        tags.append(f"domaine/{sender_domain.replace('.', '_')}")
    if year_month_tag:
        tags.append(f"periode/{year_month_tag}")
    subject_lower = subject.lower()
    for kw in MOTS_CLES:
        if kw in subject_lower:
            tags.append(f"sujet/{kw}")

    # Body & attachments
    h = None
    if HAS_HTML2TEXT:
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.body_width = 0

    body_content = ""
    attachments_links = []

    for part in msg.walk():
        content_type = part.get_content_type()
        content_disposition = str(part.get("Content-Disposition", ""))

        if "attachment" in content_disposition or part.get_filename():
            filename = part.get_filename()
            if filename:
                if content_type.startswith('image/'):
                    continue
                safe_filename = clean_string_for_file(filename)
                att_filename = f"{file_time}_{safe_filename}"
                att_filepath = os.path.join(OBSIDIAN_ATT_DIR, att_filename)
                payload = part.get_payload(decode=True)
                if payload:
                    with open(att_filepath, 'wb') as att_file:
                        att_file.write(payload)
                    attachments_links.append(f"[[{att_filename}]]")

        elif content_type == "text/plain" and "attachment" not in content_disposition:
            if not body_content:
                try:
                    charset = part.get_content_charset('utf-8') or 'utf-8'
                    body_content = part.get_payload(decode=True).decode(charset, errors='replace')
                except Exception:
                    pass
        elif content_type == "text/html" and "attachment" not in content_disposition and h:
            try:
                charset = part.get_content_charset('utf-8') or 'utf-8'
                html_content = part.get_payload(decode=True).decode(charset, errors='replace')
                body_content = h.handle(html_content)
            except Exception:
                pass

    body_lower = body_content[:500].lower()
    for kw in MOTS_CLES:
        if kw in body_lower and f"sujet/{kw}" not in tags:
            tags.append(f"sujet/{kw}")

    # Write markdown
    with open(md_filepath, 'w', encoding='utf-8') as md_file:
        md_file.write("---\n")
        md_file.write("type: email\n")
        md_file.write("tags:\n")
        for tag in tags:
            md_file.write(f"  - {tag}\n")
        md_file.write("---\n\n")
        md_file.write(f"# {subject_final}\n\n")
        if daily_note_link:
            md_file.write(f"**🗓️ Date :** {daily_note_link} ({date_str})\n")
        else:
            md_file.write(f"**🗓️ Date :** {date_str}\n")
        md_file.write(f"**👤 De :** [[{sender_name}]]\n")
        if to_list:
            to_links = ", ".join([f"[[{dest}]]" for dest in to_list])
            md_file.write(f"**👥 À :** {to_links}\n")
        if cc_list:
            cc_links = ", ".join([f"[[{cc}]]" for cc in cc_list])
            md_file.write(f"**👀 Cc :** {cc_links}\n")
        md_file.write("\n---\n\n")
        md_file.write(body_content)
        md_file.write("\n\n")
        if attachments_links:
            md_file.write("---\n### 📎 Pièces Jointes\n")
            for link in attachments_links:
                md_file.write(f"- {link}\n")

    return md_filepath


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=PROJECT_ROOT, **kw)

    def end_headers(self):
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        super().end_headers()

    def do_GET(self):
        if self.path == "/" or self.path.startswith("/index.html"):
            try:
                with open(RENDERER_INDEX, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
                return
            except FileNotFoundError:
                self.send_error(404)
                return

        if self.path.startswith("/api/oauth/google/callback"):
            return handle_oauth_callback(
                self,
                pending_store=GOOGLE_OAUTH_PENDING,
                load_accounts=load_accounts,
                find_account_index_by_email=find_account_index_by_email,
                normalize_auth_fields=normalize_auth_fields,
                save_accounts=save_accounts,
                exchange_google_auth_code=exchange_google_auth_code,
                build_oauth_callback_page=build_oauth_callback_page,
                now_ts=lambda: int(time.time()),
            )

        if self.path == "/api/state":
            return self._json(loadAppState())
        if self.path == "/api/contacts":
            return self._json(loadContactsData())
        if self.path == "/api/accounts":
            return self._json(load_accounts())
        if self.path == "/api/calendar/accounts":
            return handle_calendar_accounts_get(self, get_google_oauth_accounts=get_google_oauth_accounts)
        if self.path.startswith("/api/calendar/calendars"):
            return handle_calendar_calendars_get(
                self,
                pick_google_oauth_account=pick_google_oauth_account,
                list_google_calendars=list_google_calendars,
                build_calendar_http_error_response=build_calendar_http_error_response,
            )
        if self.path.startswith("/api/calendar/events"):
            return handle_calendar_events_get(
                self,
                pick_google_oauth_account=pick_google_oauth_account,
                list_google_calendar_events=list_google_calendar_events,
                parse_google_error_payload=parse_google_error_payload,
            )
        if self.path == "/api/inbox":
            inbox = load_inbox_index()
            # Filter out deleted and sent, sort by date desc
            visible = [m for m in inbox if not m.get("deleted") and m.get("folder") != "sent"]
            visible.sort(key=lambda m: m.get("date_ts", 0), reverse=True)
            return self._json(visible)
        if self.path == "/api/inbox/sent":
            inbox = load_inbox_index()
            sent = [m for m in inbox if m.get("folder") == "sent" and not m.get("deleted")]
            sent.sort(key=lambda m: m.get("date_ts", 0), reverse=True)
            return self._json(sent)
        if self.path.startswith("/api/mail/attachment?"):
            try:
                qs = parse_qs(urlparse(self.path).query)
                mail_id = qs.get("id", [""])[0]
                idx_raw = qs.get("idx", [None])[0]
                filename = qs.get("name", [None])[0]
                idx = int(idx_raw) if idx_raw is not None else None

                inbox = load_inbox_index()
                mail = next((m for m in inbox if m.get("id") == mail_id), None)
                if not mail:
                    self.send_error(404)
                    return

                eml_path = os.path.join(MAILS_DIR, mail.get("eml_file", ""))
                if not os.path.isfile(eml_path):
                    self.send_error(404)
                    return

                with open(eml_path, "rb") as f:
                    raw_bytes = f.read()
                msg = email_lib.message_from_bytes(raw_bytes, policy=email_policy.default)
                payload, resolved_name, content_type = get_attachment_payload(msg, index=idx, filename=filename)
                if payload is None:
                    self.send_error(404)
                    return
                content_type = content_type or "application/octet-stream"

                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Disposition", f'inline; filename="{resolved_name}"')
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            except Exception:
                self.send_error(500)
                return

        if self.path.startswith("/api/mail/"):
            mail_id = self.path.split("/api/mail/")[1]
            inbox = load_inbox_index()
            mail = next((m for m in inbox if m.get("id") == mail_id), None)
            if mail:
                mail = enrich_mail_from_eml(mail)
                return self._json(mail)
            self.send_error(404)
            return
        if self.path == "/api/vault/graph":
            try:
                return self._json(scan_vault_graph())
            except Exception as e:
                return self._json({"error": str(e)}, 500)
        if self.path.startswith("/api/vault/read?"):
            try:
                qs = parse_qs(urlparse(self.path).query)
                relpath = qs.get('path', [''])[0]
                content = read_vault_file(relpath)
                return self._json({"ok": True, "content": content, "path": relpath})
            except Exception as e:
                return self._json({"error": str(e)}, 500)
        super().do_GET()

    def do_POST(self):
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", 0)))
            data = json.loads(raw) if raw else {}
        except (json.JSONDecodeError, ValueError):
            return self._json({"error": "JSON invalide"}, 400)

        if self.path == "/api/state":
            saveAppState(data)
            return self._json({"ok": True})

        if self.path == "/api/run-v3":
            v3_path = os.path.join(DIR, "v3.py")
            try:
                result = subprocess.run(
                    ["python3", v3_path],
                    capture_output=True, text=True, timeout=300
                )
                output = result.stdout + result.stderr
                return self._json({"ok": result.returncode == 0, "output": output})
            except Exception as e:
                return self._json({"ok": False, "error": str(e)}, 500)

        if self.path == "/api/reformulate":
            try:
                corrected = ai_reformulate(data)
                return self._json({"ok": True, "text": corrected})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        if self.path == "/api/save-eml":
            try:
                filepath = save_eml_to_downloads(
                    data.get("from", ""),
                    data.get("to", ""),
                    data.get("subject", ""),
                    data.get("body", ""),
                    html_body=data.get("html_body", None)
                )
                return self._json({"ok": True, "path": filepath})
            except Exception as e:
                return self._json({"error": str(e)}, 500)



        if self.path == "/api/generate-reminder":
            try:
                result = ai_generate_reminder(data)
                return self._json({"ok": True, "reminder": result})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        if self.path == "/api/generate-reply":
            try:
                text = ai_generate_reply(data)
                return self._json({"ok": True, "text": text})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Account management ──
        if self.path == "/api/accounts/save":
            try:
                accounts = data.get("accounts", [])
                for acc in accounts:
                    normalize_auth_fields(acc)
                    if (acc.get("provider", "") or "").lower() == "gmail_oauth":
                        acc["protocol"] = "imap"
                        acc["auth_type"] = "oauth2"
                        acc["username"] = acc.get("email", acc.get("username", ""))
                        acc["imap_server"] = "imap.gmail.com"
                        acc["imap_port"] = 993
                        acc["imap_ssl"] = True
                        acc["smtp_server"] = "smtp.gmail.com"
                        acc["smtp_port"] = 587
                        acc["smtp_ssl"] = False
                        acc["smtp_starttls"] = True
                        if not acc.get("oauth_redirect_uri"):
                            acc["oauth_redirect_uri"] = "http://127.0.0.1:8080/api/oauth/google/callback"
                        if not acc.get("oauth_scope"):
                            acc["oauth_scope"] = "https://mail.google.com/"
                save_accounts(accounts)
                return self._json({"ok": True})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Start Google OAuth flow for Gmail account ──
        if self.path == "/api/oauth/google/start":
            try:
                account_email = (data.get("email", "") or "").strip()
                requested_scope = (data.get("scope", "") or "").strip()
                if not account_email or "@" not in account_email:
                    return self._json({"error": "Adresse email invalide"}, 400)

                accounts = load_accounts()
                idx = find_account_index_by_email(accounts, account_email)
                if idx < 0:
                    return self._json({"error": "Compte introuvable"}, 404)

                account = normalize_auth_fields(accounts[idx])
                client_id = (account.get("oauth_client_id", "") or "").strip()
                client_secret = (account.get("oauth_client_secret", "") or "").strip()
                redirect_uri = (account.get("oauth_redirect_uri", "") or "").strip() or "http://127.0.0.1:8080/api/oauth/google/callback"
                scope = requested_scope or (account.get("oauth_scope", "") or "").strip() or GOOGLE_MAIL_SCOPE

                if not client_id:
                    return self._json({"error": "Client ID OAuth requis pour Gmail."}, 400)

                verifier, challenge = generate_pkce_pair()
                state = secrets.token_urlsafe(24)
                GOOGLE_OAUTH_PENDING[state] = {
                    "account_email": account_email,
                    "code_verifier": verifier,
                    "created_at": int(time.time()),
                }

                query = {
                    "client_id": client_id,
                    "redirect_uri": redirect_uri,
                    "response_type": "code",
                    "scope": scope,
                    "access_type": "offline",
                    "prompt": "consent",
                    "include_granted_scopes": "true",
                    "state": state,
                    "code_challenge": challenge,
                    "code_challenge_method": "S256",
                }
                auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(query)

                # Persist normalized values before opening auth URL.
                account["provider"] = "gmail_oauth"
                account["auth_type"] = "oauth2"
                account["protocol"] = "imap"
                account["username"] = account_email
                account["email"] = account_email
                account["oauth_redirect_uri"] = redirect_uri
                account["oauth_scope"] = scope
                if client_secret:
                    account["oauth_client_secret"] = client_secret
                accounts[idx] = account
                save_accounts(accounts)

                return self._json({"ok": True, "auth_url": auth_url})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Google Calendar events CRUD ──
        if self.path == "/api/calendar/events":
            return handle_calendar_event_create_post(
                self,
                data,
                pick_google_oauth_account=pick_google_oauth_account,
                create_google_calendar_event=create_google_calendar_event,
                build_calendar_http_error_response=build_calendar_http_error_response,
            )

        if self.path == "/api/calendar/events/update":
            return handle_calendar_event_update_post(
                self,
                data,
                pick_google_oauth_account=pick_google_oauth_account,
                update_google_calendar_event=update_google_calendar_event,
                build_calendar_http_error_response=build_calendar_http_error_response,
            )

        if self.path == "/api/calendar/events/delete":
            return handle_calendar_event_delete_post(
                self,
                data,
                pick_google_oauth_account=pick_google_oauth_account,
                delete_google_calendar_event=delete_google_calendar_event,
                parse_google_error_payload=parse_google_error_payload,
            )

        # ── Email autoconfig (Mozilla Thunderbird DB) ──
        if self.path == "/api/autoconfig":
            try:
                email_addr = data.get("email", "").strip()
                if not email_addr or "@" not in email_addr:
                    return self._json({"error": "Adresse email invalide"}, 400)
                result = autoconfig_email(email_addr)
                if result:
                    return self._json({"ok": True, "config": result})
                else:
                    domain = email_addr.split("@")[-1]
                    return self._json({"error": f"Aucune configuration trouvée pour {domain}"}, 404)
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Fetch emails (POP3/IMAP) ──
        if self.path == "/api/fetch-emails":
            try:
                new_count, errors = fetch_all_accounts()
                return self._json({
                    "ok": True,
                    "new_count": new_count,
                    "errors": errors
                })
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Send email (SMTP) ──
        if self.path == "/api/send-email":
            try:
                from_addr = data.get("from", "")
                to_addr = data.get("to", "")
                subject = data.get("subject", "")
                body = data.get("body", "")
                cc = data.get("cc", "")
                attachments = data.get("attachments", None)
                html_body = data.get("html_body", None)
                account = find_account_by_email(from_addr)
                if not account:
                    return self._json({"error": f"Aucun compte configuré pour {from_addr}"}, 400)
                send_email_smtp(account, to_addr, subject, body, cc, attachments=attachments, html_body=html_body)
                return self._json({"ok": True})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Mark email read/unread/starred ──
        if self.path == "/api/mail/mark-read":
            try:
                mail_id = data.get("id", "")
                inbox = load_inbox_index()
                for m in inbox:
                    if m.get("id") == mail_id:
                        if "read" in data:
                            m["read"] = data["read"]
                        if "starred" in data:
                            m["starred"] = data["starred"]
                        break
                save_inbox_index(inbox)
                return self._json({"ok": True})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Delete email ──
        if self.path == "/api/mail/delete":
            try:
                mail_id = data.get("id", "")
                delete_on_server = data.get("delete_on_server", False)
                inbox = load_inbox_index()
                mail = next((m for m in inbox if m.get("id") == mail_id), None)
                if not mail:
                    return self._json({"error": "Mail introuvable"}, 404)

                # Delete on POP3 server if requested
                if delete_on_server and mail.get("uid") and mail.get("account"):
                    account = find_account_by_email(mail["account"])
                    if account:
                        try:
                            delete_mail_on_server(account, mail["uid"])
                        except Exception as del_err:
                            pass  # Continue even if server delete fails

                # Remove local .eml file
                eml_path = os.path.join(MAILS_DIR, mail.get("eml_file", ""))
                if os.path.isfile(eml_path):
                    os.remove(eml_path)

                # Remove from seen UIDs so we don't have stale entries
                if mail.get("uid") and mail.get("account"):
                    seen = load_seen_uids()
                    for key, uids in seen.items():
                        if mail["uid"] in uids:
                            uids.remove(mail["uid"])
                    save_seen_uids(seen)

                # Mark as deleted in index
                mail["deleted"] = True
                save_inbox_index(inbox)

                return self._json({"ok": True})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Export email to Obsidian markdown ──
        if self.path == "/api/mail/export-obsidian":
            try:
                mail_id = data.get("id", "")
                inbox = load_inbox_index()
                mail = next((m for m in inbox if m.get("id") == mail_id), None)
                if not mail:
                    return self._json({"error": "Mail introuvable"}, 404)
                md_path = export_email_to_obsidian(mail)
                return self._json({"ok": True, "path": md_path})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Bulk export to Obsidian ──
        if self.path == "/api/mail/export-obsidian-all":
            try:
                inbox = load_inbox_index()
                visible = [m for m in inbox if not m.get("deleted")]
                exported = 0
                errors = []
                for mail in visible:
                    try:
                        export_email_to_obsidian(mail)
                        exported += 1
                    except Exception as e:
                        errors.append(f"{mail.get('subject', '?')}: {e}")
                return self._json({"ok": True, "exported": exported, "errors": errors})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        # ── Import contacts CSV ──
        if self.path == "/api/contacts/import":
            try:
                csv_content = data.get("csv", "")
                if not csv_content:
                    return self._json({"error": "Aucun contenu CSV"}, 400)
                with open(CONTACTS_CSV, "w", encoding="utf-8") as f:
                    f.write(csv_content)
                new_contacts = loadContactsData()
                return self._json({"ok": True, "count": len(new_contacts)})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        self.send_error(404)

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass  # silencieux


if __name__ == "__main__":
    print(f"🚀 Todo → http://localhost:{PORT}")
    http.server.HTTPServer(("", PORT), Handler).serve_forever()
