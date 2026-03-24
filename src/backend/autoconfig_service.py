"""Auto-détection de configuration email (IMAP / SMTP).

Interroge la base Mozilla Thunderbird Autoconfig et, en cas d'échec,
sonde les hostnames IMAP/SMTP courants pour trouver les paramètres
de connexion d'un domaine email.

Dépendances internes :
    (aucune)

Dépendances externes :
    - Base Mozilla Autoconfig (autoconfig.thunderbird.net)
"""

import imaplib
import smtplib
import urllib.request
import xml.etree.ElementTree as ET


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
