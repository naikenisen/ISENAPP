import base64
import hashlib
import json
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta

from account_store import (
    find_account_index_by_email,
    load_accounts,
    normalize_auth_fields,
    save_accounts,
)


def build_oauth_callback_page(ok, message):
    color = "#34d399" if ok else "#ef4444"
    icon = "✅" if ok else "❌"
    title = "Connexion Gmail réussie" if ok else "Connexion Gmail échouée"
    return f"""<!doctype html>
<html lang=\"fr\"><head><meta charset=\"utf-8\"><title>{title}</title>
<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"></head>
<body style=\"font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:#0f172a;color:#e2e8f0;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;padding:1rem\">
  <div style=\"max-width:560px;width:100%;background:#111827;border:1px solid #374151;border-radius:12px;padding:1rem 1.2rem;box-shadow:0 8px 30px rgba(0,0,0,.35)\">
    <h1 style=\"margin:.1rem 0 .6rem 0;font-size:1.2rem;color:{color}\">{icon} {title}</h1>
    <p style=\"line-height:1.5;margin:0 0 .7rem 0\">{message}</p>
    <p style=\"line-height:1.5;margin:0;color:#94a3b8\">Tu peux maintenant fermer cet onglet et revenir dans ISENAPP.</p>
  </div>
</body></html>"""


def _b64url(data_bytes):
    return base64.urlsafe_b64encode(data_bytes).decode().rstrip("=")


def generate_pkce_pair():
    verifier = _b64url(secrets.token_bytes(64))
    challenge = _b64url(hashlib.sha256(verifier.encode("utf-8")).digest())
    return verifier, challenge


def exchange_google_auth_code(client_id, client_secret, redirect_uri, code, code_verifier):
    payload = {
        "grant_type": "authorization_code",
        "client_id": client_id,
        "code": code,
        "redirect_uri": redirect_uri,
        "code_verifier": code_verifier,
    }
    if client_secret:
        payload["client_secret"] = client_secret

    body = urllib.parse.urlencode(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token",
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read())


def refresh_google_token(account):
    refresh_token = (account.get("oauth_refresh_token", "") or "").strip()
    client_id = (account.get("oauth_client_id", "") or "").strip()
    client_secret = (account.get("oauth_client_secret", "") or "").strip()
    if not refresh_token:
        raise RuntimeError("Refresh token Gmail manquant. Reconnecte le compte OAuth.")
    if not client_id:
        raise RuntimeError("Client ID OAuth manquant pour ce compte Gmail.")

    payload = {
        "grant_type": "refresh_token",
        "client_id": client_id,
        "refresh_token": refresh_token,
    }
    if client_secret:
        payload["client_secret"] = client_secret

    body = urllib.parse.urlencode(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token",
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        token_data = json.loads(resp.read())

    access_token = token_data.get("access_token", "")
    expires_in = int(token_data.get("expires_in", 3600))
    if not access_token:
        raise RuntimeError("Google OAuth: access_token absent après refresh.")

    account["oauth_access_token"] = access_token
    account["oauth_token_expiry"] = int(time.time()) + max(30, expires_in - 30)
    if token_data.get("refresh_token"):
        account["oauth_refresh_token"] = token_data["refresh_token"]

    return account


def get_valid_gmail_access_token(account_email):
    """Load account from storage, refresh token when needed, and return valid token."""
    accounts = load_accounts()
    idx = find_account_index_by_email(accounts, account_email)
    if idx < 0:
        raise RuntimeError(f"Compte introuvable: {account_email}")

    account = normalize_auth_fields(accounts[idx])
    if account.get("auth_type") != "oauth2":
        raise RuntimeError("Ce compte n'est pas configuré en OAuth 2.0.")

    now = int(time.time())
    access_token = (account.get("oauth_access_token", "") or "").strip()
    expiry = int(account.get("oauth_token_expiry", 0) or 0)

    if access_token and expiry > now + 60:
        return access_token

    try:
        account = refresh_google_token(account)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google OAuth refresh HTTP {e.code}: {body}")
    except Exception as e:
        raise RuntimeError(f"Refresh token Gmail impossible: {e}")

    accounts[idx] = account
    save_accounts(accounts)
    return account.get("oauth_access_token", "")


def get_google_oauth_accounts():
    """Return enabled OAuth2 accounts suitable for Google APIs."""
    oauth_accounts = []
    for acc in load_accounts():
        acc = normalize_auth_fields(acc)
        if acc.get("enabled", True) is False:
            continue
        if acc.get("auth_type") != "oauth2":
            continue
        email_addr = (acc.get("email", "") or "").strip()
        if not email_addr:
            continue
        oauth_accounts.append(acc)
    return oauth_accounts


def pick_google_oauth_account(preferred_email=""):
    """Pick an OAuth account; prefer the requested email when available."""
    accounts = get_google_oauth_accounts()
    if preferred_email:
        target = preferred_email.strip().lower()
        for acc in accounts:
            if (acc.get("email", "") or "").strip().lower() == target:
                return acc
    return accounts[0] if accounts else None


def map_google_calendar_event(ev, calendar_meta=None, calendar_id="primary"):
    """Normalize Google Calendar event payload for frontend use."""
    start_data = ev.get("start", {}) or {}
    end_data = ev.get("end", {}) or {}
    start_value = start_data.get("dateTime") or start_data.get("date") or ""
    end_value = end_data.get("dateTime") or end_data.get("date") or ""
    calendar_meta = calendar_meta or {}
    return {
        "id": ev.get("id", ""),
        "summary": ev.get("summary", "(Sans titre)"),
        "description": ev.get("description", "") or "",
        "location": ev.get("location", "") or "",
        "start": start_value,
        "end": end_value,
        "allDay": bool(start_data.get("date") and not start_data.get("dateTime")),
        "htmlLink": ev.get("htmlLink", "") or "",
        "status": ev.get("status", "") or "",
        "calendarId": calendar_id,
        "calendarName": calendar_meta.get("summary", calendar_id),
        "calendarColor": calendar_meta.get("backgroundColor", "#6c8aff"),
        "calendarTextColor": calendar_meta.get("foregroundColor", "#ffffff"),
        "canEdit": bool(calendar_meta.get("canEdit", True)),
    }


def normalize_google_calendar_datetime(dt_raw):
    """Normalize local/naive datetime string to RFC3339 with timezone."""
    value = (dt_raw or "").strip()
    if not value:
        raise RuntimeError("Date/heure manquante.")

    candidate = value[:-1] + "+00:00" if value.endswith("Z") else value
    try:
        dt = datetime.fromisoformat(candidate)
    except ValueError:
        raise RuntimeError("Format date/heure invalide (ISO attendu).")

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.now().astimezone().tzinfo)

    return dt.isoformat(timespec="seconds")


def list_google_calendars(account_email):
    """Fetch the account calendar list with colors and edit rights."""
    access_token = get_valid_gmail_access_token(account_email)
    params = {
        "minAccessRole": "reader",
        "maxResults": 2500,
    }
    url = (
        "https://www.googleapis.com/calendar/v3/users/me/calendarList?"
        + urllib.parse.urlencode(params)
    )
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
    )
    with urllib.request.urlopen(req, timeout=25) as resp:
        payload = json.loads(resp.read())

    items = payload.get("items", []) if isinstance(payload, dict) else []
    calendars = []
    for cal in items:
        cal_id = (cal.get("id", "") or "").strip()
        if not cal_id:
            continue
        access_role = (cal.get("accessRole", "") or "").strip().lower()
        calendars.append({
            "id": cal_id,
            "summary": cal.get("summary", cal_id),
            "backgroundColor": cal.get("backgroundColor", "#6c8aff"),
            "foregroundColor": cal.get("foregroundColor", "#ffffff"),
            "primary": bool(cal.get("primary", False)),
            "selected": bool(cal.get("selected", True)),
            "accessRole": access_role,
            "canEdit": access_role in {"owner", "writer"},
        })
    return calendars


def list_google_calendar_events(account_email, time_min_iso, time_max_iso, calendar_ids=None):
    """Fetch Google Calendar events for an arbitrary time range and calendars."""
    access_token = get_valid_gmail_access_token(account_email)
    calendars = list_google_calendars(account_email)
    calendars_by_id = {c["id"]: c for c in calendars}

    if calendar_ids:
        target_ids = [c for c in calendar_ids if c in calendars_by_id]
    else:
        target_ids = [c["id"] for c in calendars]

    if not target_ids:
        target_ids = ["primary"]

    events = []
    for cal_id in target_ids:
        params = {
            "timeMin": time_min_iso,
            "timeMax": time_max_iso,
            "singleEvents": "true",
            "orderBy": "startTime",
            "maxResults": 2500,
        }
        url = (
            "https://www.googleapis.com/calendar/v3/calendars/"
            + urllib.parse.quote(cal_id, safe="")
            + "/events?"
            + urllib.parse.urlencode(params)
        )
        req = urllib.request.Request(
            url,
            headers={"Authorization": f"Bearer {access_token}"},
        )
        with urllib.request.urlopen(req, timeout=25) as resp:
            payload = json.loads(resp.read())

        items = payload.get("items", []) if isinstance(payload, dict) else []
        calendar_meta = calendars_by_id.get(cal_id, {
            "summary": cal_id,
            "backgroundColor": "#6c8aff",
            "foregroundColor": "#ffffff",
            "canEdit": True,
        })
        for ev in items:
            events.append(map_google_calendar_event(ev, calendar_meta=calendar_meta, calendar_id=cal_id))

    return {
        "events": events,
        "calendars": calendars,
    }


def create_google_calendar_event(account_email, payload):
    """Create an event in the primary Google Calendar."""
    access_token = get_valid_gmail_access_token(account_email)
    calendar_id = (payload.get("calendarId", "") or "").strip() or "primary"
    summary = (payload.get("summary", "") or "").strip()
    if not summary:
        raise RuntimeError("Le titre (summary) est requis.")

    all_day = bool(payload.get("allDay"))
    event_data = {
        "summary": summary,
        "description": (payload.get("description", "") or "").strip(),
        "location": (payload.get("location", "") or "").strip(),
    }

    if all_day:
        start_date = (payload.get("startDate", "") or "").strip()
        end_date = (payload.get("endDate", "") or "").strip()
        if not start_date:
            raise RuntimeError("Date de début requise pour un événement journée entière.")

        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d")
            if end_date:
                end_dt = datetime.strptime(end_date, "%Y-%m-%d")
            else:
                end_dt = start_dt + timedelta(days=1)
            if end_dt <= start_dt:
                end_dt = start_dt + timedelta(days=1)
        except ValueError:
            raise RuntimeError("Format de date invalide (AAAA-MM-JJ attendu).")

        event_data["start"] = {"date": start_dt.strftime("%Y-%m-%d")}
        event_data["end"] = {"date": end_dt.strftime("%Y-%m-%d")}
    else:
        start_dt_iso = (payload.get("startDateTime", "") or "").strip()
        end_dt_iso = (payload.get("endDateTime", "") or "").strip()
        if not start_dt_iso or not end_dt_iso:
            raise RuntimeError("Dates/horaires de début et fin requis.")

        event_data["start"] = {"dateTime": normalize_google_calendar_datetime(start_dt_iso)}
        event_data["end"] = {"dateTime": normalize_google_calendar_datetime(end_dt_iso)}

    body = json.dumps(event_data).encode("utf-8")
    req = urllib.request.Request(
        "https://www.googleapis.com/calendar/v3/calendars/"
        + urllib.parse.quote(calendar_id, safe="")
        + "/events",
        data=body,
        method="POST",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=25) as resp:
        created = json.loads(resp.read())
    calendars = list_google_calendars(account_email)
    meta = next((c for c in calendars if c["id"] == calendar_id), None) or {}
    return map_google_calendar_event(created, calendar_meta=meta, calendar_id=calendar_id)


def update_google_calendar_event(account_email, payload):
    """Patch an existing event on a specific calendar."""
    access_token = get_valid_gmail_access_token(account_email)
    calendar_id = (payload.get("calendarId", "") or "").strip() or "primary"
    event_id = (payload.get("eventId", "") or "").strip()
    if not event_id:
        raise RuntimeError("eventId requis.")

    body_payload = {}
    if "summary" in payload:
        body_payload["summary"] = (payload.get("summary", "") or "").strip()
    if "description" in payload:
        body_payload["description"] = payload.get("description", "") or ""
    if "location" in payload:
        body_payload["location"] = payload.get("location", "") or ""

    if payload.get("allDay") is True:
        start_date = (payload.get("startDate", "") or "").strip()
        end_date = (payload.get("endDate", "") or "").strip()
        if not start_date:
            raise RuntimeError("startDate requis pour un événement journée entière.")
        if not end_date:
            end_dt = datetime.strptime(start_date, "%Y-%m-%d") + timedelta(days=1)
            end_date = end_dt.strftime("%Y-%m-%d")
        body_payload["start"] = {"date": start_date}
        body_payload["end"] = {"date": end_date}
    elif payload.get("allDay") is False:
        start_dt_iso = (payload.get("startDateTime", "") or "").strip()
        end_dt_iso = (payload.get("endDateTime", "") or "").strip()
        if not start_dt_iso or not end_dt_iso:
            raise RuntimeError("startDateTime/endDateTime requis pour un événement horaire.")
        body_payload["start"] = {"dateTime": normalize_google_calendar_datetime(start_dt_iso)}
        body_payload["end"] = {"dateTime": normalize_google_calendar_datetime(end_dt_iso)}

    if not body_payload:
        raise RuntimeError("Aucune propriété à mettre à jour.")

    body = json.dumps(body_payload).encode("utf-8")
    url = (
        "https://www.googleapis.com/calendar/v3/calendars/"
        + urllib.parse.quote(calendar_id, safe="")
        + "/events/"
        + urllib.parse.quote(event_id, safe="")
    )
    req = urllib.request.Request(
        url,
        data=body,
        method="PATCH",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=25) as resp:
        updated = json.loads(resp.read())

    calendars = list_google_calendars(account_email)
    meta = next((c for c in calendars if c["id"] == calendar_id), None) or {}
    return map_google_calendar_event(updated, calendar_meta=meta, calendar_id=calendar_id)


def delete_google_calendar_event(account_email, event_id):
    """Delete an event from the primary Google Calendar."""
    access_token = get_valid_gmail_access_token(account_email)
    calendar_id = "primary"
    if isinstance(event_id, dict):
        calendar_id = (event_id.get("calendarId", "") or "").strip() or "primary"
        event_id = event_id.get("eventId", "")

    event_id = (event_id or "").strip()
    if not event_id:
        raise RuntimeError("eventId requis.")

    url = (
        "https://www.googleapis.com/calendar/v3/calendars/"
        + urllib.parse.quote(calendar_id, safe="")
        + "/events/"
        + urllib.parse.quote(event_id, safe="")
    )
    req = urllib.request.Request(
        url,
        method="DELETE",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    with urllib.request.urlopen(req, timeout=25):
        return True


def parse_google_error_payload(body_text):
    """Parse Google API error JSON and expose actionable metadata."""
    info = {
        "message": body_text,
        "reason": "",
        "status": "",
        "activation_url": "",
    }
    try:
        payload = json.loads(body_text or "{}")
        err = payload.get("error", {}) if isinstance(payload, dict) else {}
        if isinstance(err, dict):
            info["message"] = err.get("message") or info["message"]
            info["status"] = err.get("status") or ""

            errors = err.get("errors") if isinstance(err.get("errors"), list) else []
            if errors and isinstance(errors[0], dict):
                info["reason"] = errors[0].get("reason", "") or info["reason"]

            details = err.get("details") or []
            if not isinstance(details, list):
                details = []
            for detail in details:
                if not isinstance(detail, dict):
                    continue
                if detail.get("reason"):
                    info["reason"] = detail.get("reason")
                links = detail.get("links") or []
                if not isinstance(links, list):
                    links = []
                for link in links:
                    if isinstance(link, dict) and link.get("url"):
                        info["activation_url"] = link.get("url")
                        break
                if info["activation_url"]:
                    break

        if not info["activation_url"] and "console.developers.google.com/apis/api/calendar-json.googleapis.com/overview" in body_text:
            marker = "https://console.developers.google.com/apis/api/calendar-json.googleapis.com/overview"
            start = body_text.find(marker)
            if start >= 0:
                end = body_text.find('"', start)
                info["activation_url"] = body_text[start:end] if end > start else marker
    except Exception:
        pass
    return info


def build_calendar_http_error_response(http_err):
    """Build a normalized JSON payload for Google Calendar HTTP errors."""
    body = http_err.read().decode("utf-8", errors="replace")
    info = parse_google_error_payload(body)
    reason = (info.get("reason", "") or "").lower()
    status = (info.get("status", "") or "").upper()

    if reason in {"insufficientpermissions", "access_token_scope_insufficient"}:
        return {
            "ok": False,
            "error_code": "CALENDAR_SCOPE_INSUFFICIENT",
            "error": "Le token OAuth n'a pas le scope Google Calendar requis.",
            "details": info.get("message", body),
        }, 502

    if reason in {"forbiddenfornonorganizer", "forbidden"} or (http_err.code == 403 and status == "PERMISSION_DENIED"):
        return {
            "ok": False,
            "error_code": "CALENDAR_EVENT_FORBIDDEN",
            "error": "Cet événement ou agenda ne peut pas être modifié avec ce compte.",
            "details": info.get("message", body),
        }, 502

    if reason in {"accessnotconfigured", "service_disabled"}:
        return {
            "ok": False,
            "error_code": "CALENDAR_API_DISABLED",
            "error": "Google Calendar API n'est pas activée pour ce projet Google Cloud.",
            "details": info.get("message", body),
            "activation_url": info.get("activation_url", ""),
        }, 502

    return {
        "ok": False,
        "error_code": "CALENDAR_HTTP_ERROR",
        "error": f"Google Calendar HTTP {http_err.code}",
        "details": info.get("message", body),
    }, 502
