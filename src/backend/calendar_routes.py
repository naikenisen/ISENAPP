from datetime import datetime
from urllib.parse import parse_qs, urlparse
import urllib.error


def handle_oauth_callback(
    handler,
    *,
    pending_store,
    load_accounts,
    find_account_index_by_email,
    normalize_auth_fields,
    save_accounts,
    exchange_google_auth_code,
    build_oauth_callback_page,
    now_ts,
):
    try:
        qs = parse_qs(urlparse(handler.path).query)
        err = (qs.get("error", [""])[0] or "").strip()
        state = (qs.get("state", [""])[0] or "").strip()
        code = (qs.get("code", [""])[0] or "").strip()

        if err:
            html = build_oauth_callback_page(False, f"Google a renvoyé une erreur: {err}")
            handler.send_response(400)
            handler.send_header("Content-Type", "text/html; charset=utf-8")
            handler.end_headers()
            handler.wfile.write(html.encode("utf-8"))
            return True

        pending = pending_store.pop(state, None)
        if not pending:
            html = build_oauth_callback_page(False, "État OAuth invalide ou expiré.")
            handler.send_response(400)
            handler.send_header("Content-Type", "text/html; charset=utf-8")
            handler.end_headers()
            handler.wfile.write(html.encode("utf-8"))
            return True

        if not code:
            html = build_oauth_callback_page(False, "Code OAuth absent.")
            handler.send_response(400)
            handler.send_header("Content-Type", "text/html; charset=utf-8")
            handler.end_headers()
            handler.wfile.write(html.encode("utf-8"))
            return True

        account_email = pending["account_email"]
        accounts = load_accounts()
        idx = find_account_index_by_email(accounts, account_email)
        if idx < 0:
            html = build_oauth_callback_page(False, "Compte cible introuvable. Réessaie depuis l'application.")
            handler.send_response(404)
            handler.send_header("Content-Type", "text/html; charset=utf-8")
            handler.end_headers()
            handler.wfile.write(html.encode("utf-8"))
            return True

        account = normalize_auth_fields(accounts[idx])
        client_id = (account.get("oauth_client_id", "") or "").strip()
        client_secret = (account.get("oauth_client_secret", "") or "").strip()
        redirect_uri = (account.get("oauth_redirect_uri", "") or "").strip() or "http://127.0.0.1:8080/api/oauth/google/callback"

        token_data = exchange_google_auth_code(
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            code=code,
            code_verifier=pending["code_verifier"],
        )

        access_token = token_data.get("access_token", "")
        refresh_token = token_data.get("refresh_token", "")
        expires_in = int(token_data.get("expires_in", 3600))
        if not access_token:
            raise RuntimeError("Google OAuth: access_token absent")

        account["provider"] = "gmail_oauth"
        account["auth_type"] = "oauth2"
        account["protocol"] = "imap"
        account["email"] = account_email
        account["username"] = account_email
        account["imap_server"] = "imap.gmail.com"
        account["imap_port"] = 993
        account["imap_ssl"] = True
        account["imap_post_action"] = account.get("imap_post_action", "mark_read")
        account["smtp_server"] = "smtp.gmail.com"
        account["smtp_port"] = 587
        account["smtp_ssl"] = False
        account["smtp_starttls"] = True
        account["oauth_access_token"] = access_token
        account["oauth_token_expiry"] = now_ts() + max(30, expires_in - 30)
        if refresh_token:
            account["oauth_refresh_token"] = refresh_token
        granted_scope = (token_data.get("scope", "") or "").strip()
        if granted_scope:
            account["oauth_scope"] = granted_scope

        accounts[idx] = account
        save_accounts(accounts)

        html = build_oauth_callback_page(True, f"Le compte {account_email} est désormais connecté via Google OAuth 2.0.")
        handler.send_response(200)
        handler.send_header("Content-Type", "text/html; charset=utf-8")
        handler.end_headers()
        handler.wfile.write(html.encode("utf-8"))
        return True
    except Exception as e:
        html = build_oauth_callback_page(False, f"Impossible de finaliser OAuth: {e}")
        handler.send_response(500)
        handler.send_header("Content-Type", "text/html; charset=utf-8")
        handler.end_headers()
        handler.wfile.write(html.encode("utf-8"))
        return True


def handle_calendar_accounts_get(handler, *, get_google_oauth_accounts):
    accounts = get_google_oauth_accounts()
    return handler._json([
        {
            "email": (acc.get("email", "") or "").strip(),
            "provider": acc.get("provider", ""),
            "connected": bool((acc.get("oauth_refresh_token", "") or "").strip()),
        }
        for acc in accounts
    ])


def handle_calendar_calendars_get(
    handler,
    *,
    pick_google_oauth_account,
    list_google_calendars,
    build_calendar_http_error_response,
):
    try:
        qs = parse_qs(urlparse(handler.path).query)
        account_email = (qs.get("account", [""])[0] or "").strip()
        account = pick_google_oauth_account(account_email)
        if not account:
            return handler._json({"error": "Aucun compte Google OAuth disponible."}, 404)

        calendars = list_google_calendars(account.get("email", ""))
        return handler._json({
            "ok": True,
            "account": account.get("email", ""),
            "calendars": calendars,
        })
    except urllib.error.HTTPError as e:
        payload, code = build_calendar_http_error_response(e)
        return handler._json(payload, code)
    except Exception as e:
        return handler._json({"error": str(e)}, 500)


def handle_calendar_events_get(
    handler,
    *,
    pick_google_oauth_account,
    list_google_calendar_events,
    parse_google_error_payload,
):
    try:
        qs = parse_qs(urlparse(handler.path).query)
        year_raw = (qs.get("year", [""])[0] or "").strip()
        month_raw = (qs.get("month", [""])[0] or "").strip()
        start_raw = (qs.get("start", [""])[0] or "").strip()
        end_raw = (qs.get("end", [""])[0] or "").strip()
        account_email = (qs.get("account", [""])[0] or "").strip()
        calendars_raw = (qs.get("calendars", [""])[0] or "").strip()
        calendar_ids = [c.strip() for c in calendars_raw.split(",") if c.strip()]

        now = datetime.now()
        year = int(year_raw) if year_raw.isdigit() else now.year
        month = int(month_raw) if month_raw.isdigit() else now.month

        if start_raw and end_raw:
            try:
                start_dt = datetime.strptime(start_raw, "%Y-%m-%d")
                end_dt = datetime.strptime(end_raw, "%Y-%m-%d")
            except ValueError:
                return handler._json({"error": "Format start/end invalide (AAAA-MM-JJ attendu)."}, 400)

            if end_dt <= start_dt:
                return handler._json({"error": "La date de fin doit être après la date de début."}, 400)

            time_min_iso = start_dt.strftime("%Y-%m-%dT00:00:00Z")
            time_max_iso = end_dt.strftime("%Y-%m-%dT00:00:00Z")
        else:
            if month < 1 or month > 12:
                return handler._json({"error": "Mois invalide (1-12)."}, 400)

            start_dt = datetime(year, month, 1)
            if month == 12:
                end_dt = datetime(year + 1, 1, 1)
            else:
                end_dt = datetime(year, month + 1, 1)
            time_min_iso = start_dt.strftime("%Y-%m-%dT00:00:00Z")
            time_max_iso = end_dt.strftime("%Y-%m-%dT00:00:00Z")

        account = pick_google_oauth_account(account_email)
        if not account:
            return handler._json({"error": "Aucun compte Google OAuth disponible."}, 404)

        result_data = list_google_calendar_events(
            account.get("email", ""),
            time_min_iso,
            time_max_iso,
            calendar_ids=calendar_ids,
        )
        return handler._json({
            "ok": True,
            "account": account.get("email", ""),
            "year": year,
            "month": month,
            "start": start_dt.strftime("%Y-%m-%d"),
            "end": end_dt.strftime("%Y-%m-%d"),
            "events": result_data.get("events", []),
            "calendars": result_data.get("calendars", []),
        })
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        info = parse_google_error_payload(body)
        reason = (info.get("reason", "") or "").lower()
        status = (info.get("status", "") or "").upper()

        if reason in {"accessnotconfigured", "service_disabled"} or status == "PERMISSION_DENIED":
            return handler._json({
                "ok": False,
                "error_code": "CALENDAR_API_DISABLED",
                "error": "Google Calendar API n'est pas activée pour ce projet Google Cloud.",
                "details": info.get("message", ""),
                "activation_url": info.get("activation_url", ""),
            }, 502)

        if reason in {"insufficientpermissions", "access_token_scope_insufficient"}:
            return handler._json({
                "ok": False,
                "error_code": "CALENDAR_SCOPE_INSUFFICIENT",
                "error": "Le token OAuth n'a pas le scope Google Calendar requis.",
                "details": info.get("message", ""),
            }, 502)

        return handler._json({
            "ok": False,
            "error_code": "CALENDAR_HTTP_ERROR",
            "error": f"Google Calendar HTTP {e.code}",
            "details": info.get("message", body),
        }, 502)
    except Exception as e:
        return handler._json({"error": str(e)}, 500)


def handle_calendar_event_create_post(
    handler,
    data,
    *,
    pick_google_oauth_account,
    create_google_calendar_event,
    build_calendar_http_error_response,
):
    try:
        account_email = (data.get("account", "") or "").strip()
        account = pick_google_oauth_account(account_email)
        if not account:
            return handler._json({"error": "Aucun compte Google OAuth disponible."}, 404)

        created = create_google_calendar_event(account.get("email", ""), data)
        return handler._json({"ok": True, "event": created})
    except urllib.error.HTTPError as e:
        payload, code = build_calendar_http_error_response(e)
        return handler._json(payload, code)
    except Exception as e:
        return handler._json({"error": str(e)}, 500)


def handle_calendar_event_update_post(
    handler,
    data,
    *,
    pick_google_oauth_account,
    update_google_calendar_event,
    build_calendar_http_error_response,
):
    try:
        account_email = (data.get("account", "") or "").strip()
        account = pick_google_oauth_account(account_email)
        if not account:
            return handler._json({"error": "Aucun compte Google OAuth disponible."}, 404)

        updated = update_google_calendar_event(account.get("email", ""), data)
        return handler._json({"ok": True, "event": updated})
    except urllib.error.HTTPError as e:
        payload, code = build_calendar_http_error_response(e)
        return handler._json(payload, code)
    except Exception as e:
        return handler._json({"error": str(e)}, 500)


def handle_calendar_event_delete_post(
    handler,
    data,
    *,
    pick_google_oauth_account,
    delete_google_calendar_event,
    parse_google_error_payload,
):
    try:
        account_email = (data.get("account", "") or "").strip()
        event_id = (data.get("eventId", "") or "").strip()
        calendar_id = (data.get("calendarId", "") or "").strip() or "primary"
        account = pick_google_oauth_account(account_email)
        if not account:
            return handler._json({"error": "Aucun compte Google OAuth disponible."}, 404)
        if not event_id:
            return handler._json({"error": "eventId requis."}, 400)

        delete_google_calendar_event(account.get("email", ""), {
            "eventId": event_id,
            "calendarId": calendar_id,
        })
        return handler._json({"ok": True})
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        info = parse_google_error_payload(body)
        return handler._json({
            "ok": False,
            "error_code": "CALENDAR_HTTP_ERROR",
            "error": f"Google Calendar HTTP {e.code}",
            "details": info.get("message", body),
        }, 502)
    except Exception as e:
        return handler._json({"error": str(e)}, 500)
