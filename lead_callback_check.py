#!/usr/bin/env python3
"""
=============================================================================
 Lead Callback Check
=============================================================================

WHAT THIS DOES
--------------
Runs once at the end of each business day and answers one question for every
lead that came in: "did a loan officer call this person, and how fast?"

The flow, top to bottom:

    1. HubSpot list  ->  who submitted a form, and WHEN they were added to the
                         list. That add time is the "lead time".
    2. Google Sheet  ->  any leads from previous days that are still marked
                         NOT CALLED get pulled back in and re-checked.
    3. HubSpot       ->  for each lead, pull the calls logged on the contact
                         that happened AFTER the lead time.
    4. Math          ->  business hours between lead time and the first call,
                         bucketed into a tier (OK / Flag 1 / Flag 2 / Flag 3).
                         No call at all = NOT CALLED.
    5. Google Sheet  ->  one row per lead. Existing rows (same contact + same
                         lead time) are updated in place, new ones appended.
    6. Email         ->  styled HTML report in two parts:
                           - the report day's leads (grouped by owner: not called,
                             then called late, then a one-liner for on-time)
                           - the backlog: anyone from earlier days still uncalled

REPORT DAY
----------
The "report day" is the business day being scored. Normally that's today, but
if GitHub runs the job late (after midnight, before the office opens) it's
yesterday. Weekend and after-hours leads roll into the next business day,
because the rep couldn't have called before the office opened anyway. So a
Monday report covers Friday 5pm through Monday close.

HOW TO READ THIS FILE
---------------------
    Config            all the knobs, read from environment variables
    Small helpers     date parsing/formatting and the business-hours math
    HubSpot           every call to the HubSpot API lives here
    Evaluation        turns one contact + its calls into one result row
    Google Sheets     reading the Master tab and writing results back
    Email             builds the HTML and sends it
    Main              glues the sections together in order

ENVIRONMENT VARIABLES
---------------------
    HUBSPOT_TOKEN                private app token               (secret)
    GOOGLE_SERVICE_ACCOUNT_JSON  service account key file, as text (secret)
    GOOGLE_SHEET_ID              the ID in the sheet URL         (secret)
    GMAIL_USER                   sending Gmail address           (secret)
    GMAIL_APP_PASSWORD           16-char Google App Password     (secret)

    HUBSPOT_LIST_ID    default 19329         HUBSPOT_PORTAL_ID  default 6624066
    EMAIL_TO           default daniel's      SHEET_TAB          default Master
    TIMEZONE           default America/Phoenix
    BUSINESS_START     default 8             BUSINESS_END       default 17
    RECHECK_DAYS       default 7             DRY_RUN            default false
    EXCLUDE_FORMS      default "Credit Auth" (comma-separated, case-insensitive
                       substring match on the Form / Event column)
    REPORT_DAY_ONLY    default true (only score list members whose lead time
                       falls in the report day window; older ones come back
                       through the sheet carry-forward instead)
"""

# --------------------------------------------------------------------------- #
# Imports
# --------------------------------------------------------------------------- #

import json
import logging
import os
import smtplib
import sys
import time
from datetime import date, datetime, timedelta, timezone
from datetime import time as dtime          # renamed so it doesn't clash with the `time` module
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import requests


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #

# --- HubSpot ---
HUBSPOT_TOKEN = os.environ.get("HUBSPOT_TOKEN", "")
LIST_ID = os.environ.get("HUBSPOT_LIST_ID", "19329")        # "Form Submissions Today"
PORTAL_ID = os.environ.get("HUBSPOT_PORTAL_ID", "6624066")  # only used to build record links

# --- Google Sheets ---
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
SHEET_TAB = os.environ.get("SHEET_TAB", "Master")
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/edit" if GOOGLE_SHEET_ID else ""

# --- Email ---
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
EMAIL_TO = [e.strip() for e in os.environ.get("EMAIL_TO", "daniel.haycraft@capitalfund1.com").split(",") if e.strip()]

# --- Business-hours clock ---
TZ = ZoneInfo(os.environ.get("TIMEZONE", "America/Phoenix"))
BUSINESS_START = int(os.environ.get("BUSINESS_START", "8"))   # 8  = 8:00am
BUSINESS_END = int(os.environ.get("BUSINESS_END", "17"))      # 17 = 5:00pm
WORKDAYS = {0, 1, 2, 3, 4}                                    # Python weekday(): Mon=0 ... Sun=6
DAY_HOURS = BUSINESS_END - BUSINESS_START                     # 9 hours = "one full business day"

# --- Behaviour ---
RECHECK_DAYS = int(os.environ.get("RECHECK_DAYS", "7"))
REPORT_DAY_ONLY = os.environ.get("REPORT_DAY_ONLY", "true").strip().lower() in ("1", "true", "yes")
EXCLUDE_FORMS = [s.strip().lower() for s in os.environ.get("EXCLUDE_FORMS", "Credit Auth").split(",") if s.strip()]
DRY_RUN = os.environ.get("DRY_RUN", "false").strip().lower() in ("1", "true", "yes")

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("lead_callback_check")

HS_BASE = "https://api.hubapi.com"
SESSION = requests.Session()
SESSION.headers.update({"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"})

CONTACT_PROPERTIES = [
    "firstname", "lastname", "email", "phone", "mobilephone",
    "hubspot_owner_id", "hs_lead_status", "sales_lead_source_2_0", "contact_type_2_0",
    "createdate", "recent_conversion_date", "recent_conversion_event_name",
    "heyflow_date", "heyflow_form", "hs_object_source_detail_1",
    "hs_time_to_first_engagement",
]

CALL_PROPERTIES = [
    "hs_timestamp", "hs_call_direction", "hs_call_status", "hs_call_disposition",
    "hs_call_title", "hs_call_duration", "hubspot_owner_id", "hs_createdate",
]

# Column order in the Google Sheet. evaluate()'s result dict uses these exact keys
# (plus a few private "_" keys for the email that never get written to the sheet).
SHEET_HEADERS = [
    "Run Date", "Lead Added At", "Contact ID", "Name", "Email", "Phone",
    "Contact Owner", "Lead Status", "Lead Source", "Form / Event", "Recent Conversion Date",
    "First Call At", "Call Owner", "Call Direction", "Call Outcome", "Calls After Lead",
    "Biz Hours To First Call", "Biz Hours Elapsed", "Tier", "Status",
    "HS Lead Response Time (hrs)", "HubSpot Link", "Last Checked",
]

DT_FMT = "%Y-%m-%d %H:%M:%S"

# Tier labels and the colours the email uses for them.
TIER_OK = "OK (<1h)"
TIER_1 = "Flag 1 (1-3h)"
TIER_2 = "Flag 2 (3h+)"
TIER_3 = "Flag 3 (1+ biz day)"
TIER_AFTER_HOURS = "After hours (clock not started)"

TIER_STYLE = {   # background, text colour, short label for the pill
    TIER_OK:          ("#2E7D32", "#FFFFFF", "On time"),
    TIER_1:           ("#F9A825", "#1B1B1B", "Flag 1"),
    TIER_2:           ("#EF6C00", "#FFFFFF", "Flag 2"),
    TIER_3:           ("#C62828", "#FFFFFF", "Flag 3"),
    TIER_AFTER_HOURS: ("#90A4AE", "#FFFFFF", "After hrs"),
}

# Email palette. Swap these for CF1 brand colours if you want.
C_HEADER_BG = "#1F3A5F"
C_HEADER_TEXT = "#FFFFFF"
C_PAGE_BG = "#F2F4F7"
C_CARD_BG = "#FFFFFF"
C_TEXT = "#1B1B1B"
C_MUTED = "#6B7280"
C_LINE = "#E5E7EB"
C_OWNER_BG = "#EEF2F7"
C_LINK = "#1F3A5F"


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #

def now_tz() -> datetime:
    """Current time as a timezone-aware datetime in the configured TZ (Arizona)."""
    return datetime.now(timezone.utc).astimezone(TZ)


def parse_hs_datetime(value):
    """
    HubSpot date value -> aware datetime, or None.
    Handles ISO strings ('2026-11-02T15:00:14.907Z') and epoch-ms strings ('1793491200000').
    """
    if value in (None, ""):
        return None
    try:
        if isinstance(value, (int, float)) or str(value).isdigit():
            return datetime.fromtimestamp(int(value) / 1000, tz=timezone.utc)
        s = str(value)
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        return datetime.fromisoformat(s)
    except Exception:
        return None


def fmt_dt(dt) -> str:
    """Aware datetime -> 'YYYY-MM-DD HH:MM:SS' in TZ. Empty string for None. Used for the sheet."""
    if not dt:
        return ""
    return dt.astimezone(TZ).strftime(DT_FMT)


def fmt_short(dt, with_date=False) -> str:
    """Human time for the email: 'Mon 9:39am', or 'Tue 11/3 2:28pm' with_date=True."""
    if not dt:
        return ""
    dt = dt.astimezone(TZ)
    t = dt.strftime("%I:%M%p").lstrip("0").lower()
    if with_date:
        return f"{dt.strftime('%a')} {dt.month}/{dt.day} {t}"
    return f"{dt.strftime('%a')} {t}"


def parse_sheet_dt(s: str):
    """Reverse of fmt_dt(): sheet cell -> aware datetime, or None."""
    try:
        return datetime.strptime(s.strip(), DT_FMT).replace(tzinfo=TZ)
    except Exception:
        return None


def business_hours_between(start: datetime, end: datetime) -> float:
    """
    Business hours (float) between two datetimes, counting only
    BUSINESS_START-BUSINESS_END on WORKDAYS in TZ. Walks day by day and adds
    the overlap of [start, end] with each workday's [open, close].
    """
    if not start or not end or end <= start:
        return 0.0
    start = start.astimezone(TZ)
    end = end.astimezone(TZ)
    total = 0.0
    day = start.date()
    while day <= end.date():
        if day.weekday() in WORKDAYS:
            opens = datetime.combine(day, dtime(BUSINESS_START), tzinfo=TZ)
            closes = datetime.combine(day, dtime(BUSINESS_END), tzinfo=TZ)
            s = max(start, opens)
            e = min(end, closes)
            if e > s:
                total += (e - s).total_seconds() / 3600.0
        day += timedelta(days=1)
    return round(total, 2)


def tier_for(hours: float) -> str:
    """Bucket a business-hour count into a tier label. Change the cutoffs here."""
    if hours < 1:
        return TIER_OK
    if hours < 3:
        return TIER_1
    if hours < DAY_HOURS:
        return TIER_2
    return TIER_3


def report_day_for(run_at: datetime) -> date:
    """
    Which business day this run is reporting on.
    - Before the office opens -> yesterday (GitHub ran us late).
    - Weekend -> the previous Friday.
    """
    d = run_at.astimezone(TZ).date()
    if run_at.astimezone(TZ).hour < BUSINESS_START:
        d -= timedelta(days=1)
    while d.weekday() not in WORKDAYS:
        d -= timedelta(days=1)
    return d


def previous_business_close(day: date) -> datetime:
    """Closing time of the last workday before `day`. Leads after this belong to `day`."""
    p = day - timedelta(days=1)
    while p.weekday() not in WORKDAYS:
        p -= timedelta(days=1)
    return datetime.combine(p, dtime(BUSINESS_END), tzinfo=TZ)


def is_excluded(source: str) -> bool:
    """True if the Form / Event text matches anything in EXCLUDE_FORMS."""
    s = (source or "").lower()
    return any(x in s for x in EXCLUDE_FORMS)


def hs_link(contact_id) -> str:
    return f"https://app.hubspot.com/contacts/{PORTAL_ID}/record/0-1/{contact_id}"


# --------------------------------------------------------------------------- #
# HubSpot
# --------------------------------------------------------------------------- #

def hs_request(method: str, path: str, **kwargs):
    """One HubSpot API call with retries on 429 and 5xx. Raises on other errors with the body in the message."""
    url = f"{HS_BASE}{path}"
    for attempt in range(6):
        resp = SESSION.request(method, url, timeout=30, **kwargs)
        if resp.status_code == 429:
            wait = float(resp.headers.get("Retry-After", "2"))
            log.warning("HubSpot 429, sleeping %.1fs", wait)
            time.sleep(wait)
            continue
        if resp.status_code >= 500:
            log.warning("HubSpot %s on %s, retry %d", resp.status_code, path, attempt + 1)
            time.sleep(2 * (attempt + 1))
            continue
        if not resp.ok:
            raise RuntimeError(f"HubSpot {method} {path} -> {resp.status_code}: {resp.text[:500]}")
        return resp.json() if resp.text else {}
    raise RuntimeError(f"HubSpot {method} {path} failed after retries")


def get_list_memberships(list_id: str) -> dict:
    """
    Everyone in the list with the moment they were added (the lead time).
    GET /crm/v3/lists/{listId}/memberships/join-order, paginated.
    Returns {contact_id (str): added_at (aware datetime)}.
    """
    members = {}
    after = None
    while True:
        params = {"limit": 250}
        if after:
            params["after"] = after
        data = hs_request("GET", f"/crm/v3/lists/{list_id}/memberships/join-order", params=params)
        for m in data.get("results", []):
            added = parse_hs_datetime(m.get("membershipTimestamp"))
            if added:
                members[str(m["recordId"])] = added
        after = (data.get("paging") or {}).get("next", {}).get("after")
        if not after:
            break
    return members


def get_contacts(contact_ids) -> dict:
    """Batch read CONTACT_PROPERTIES, 100 IDs per call. Returns {contact_id: properties}."""
    out = {}
    ids = list(contact_ids)
    for i in range(0, len(ids), 100):
        chunk = ids[i:i + 100]
        body = {"properties": CONTACT_PROPERTIES, "inputs": [{"id": cid} for cid in chunk]}
        data = hs_request("POST", "/crm/v3/objects/contacts/batch/read", data=json.dumps(body))
        for r in data.get("results", []):
            out[str(r["id"])] = r.get("properties", {}) or {}
    return out


def get_owners() -> dict:
    """Owner ID -> 'First Last', active and archived. GET /crm/v3/owners/."""
    owners = {}
    for archived in ("false", "true"):
        after = None
        while True:
            params = {"limit": 500, "archived": archived}
            if after:
                params["after"] = after
            data = hs_request("GET", "/crm/v3/owners/", params=params)
            for o in data.get("results", []):
                name = f"{o.get('firstName', '') or ''} {o.get('lastName', '') or ''}".strip() or o.get("email", "") or str(o["id"])
                owners[str(o["id"])] = name
            after = (data.get("paging") or {}).get("next", {}).get("after")
            if not after:
                break
    return owners


def get_call_dispositions() -> dict:
    """Call outcome GUID -> label. Falls back to {} if the endpoint isn't available to this token."""
    try:
        data = hs_request("GET", "/calling/v1/dispositions")
        return {d["id"]: d.get("label", d["id"]) for d in data}
    except Exception as exc:
        log.warning("Could not load call dispositions: %s", exc)
        return {}


def get_calls_after(contact_id: str, since: datetime) -> list:
    """
    Calls associated with the contact with hs_timestamp >= since, oldest first.
    POST /crm/v3/objects/calls/search. This is the "what counts as a call" rule:
    any call, any direction, any outcome. Tighten it by adding filters here.
    """
    since_ms = int(since.timestamp() * 1000)
    body = {
        "filterGroups": [{
            "filters": [
                {"propertyName": "associations.contact", "operator": "EQ", "value": str(contact_id)},
                {"propertyName": "hs_timestamp", "operator": "GTE", "value": str(since_ms)},
            ]
        }],
        "sorts": [{"propertyName": "hs_timestamp", "direction": "ASCENDING"}],
        "properties": CALL_PROPERTIES,
        "limit": 50,
    }
    data = hs_request("POST", "/crm/v3/objects/calls/search", data=json.dumps(body))
    calls = []
    for r in data.get("results", []):
        p = r.get("properties", {}) or {}
        ts = parse_hs_datetime(p.get("hs_timestamp"))
        if ts and ts >= since:
            calls.append({"id": r["id"], "ts": ts, **p})
    calls.sort(key=lambda c: c["ts"])
    return calls


# --------------------------------------------------------------------------- #
# Evaluation
# --------------------------------------------------------------------------- #

def describe_source(props: dict, added_at: datetime) -> str:
    """
    Which form/event put this contact in the list. Informational, but also
    what EXCLUDE_FORMS matches against.

    HubSpot form names come as "Page title: Form name"; we keep the part after
    the last ': ' so "Contact Me Website" shows instead of the whole page title.
    """
    added_day = added_at.astimezone(TZ).date()

    # Collect every same-day signal, then prefer one that is NOT on the exclude
    # list. So a contact who did a Heyflow app AND a credit auth the same day is
    # scored as a Heyflow lead; only a contact whose sole same-day signal is an
    # excluded form (credit auth and nothing else) gets excluded.
    signals = []
    rc = parse_hs_datetime(props.get("recent_conversion_date"))
    if rc and rc.astimezone(TZ).date() == added_day and props.get("recent_conversion_event_name"):
        signals.append(props["recent_conversion_event_name"].rsplit(": ", 1)[-1])
    hf = parse_hs_datetime(props.get("heyflow_date"))
    if hf and hf.date() == added_day:
        signals.append(f"Heyflow: {props.get('heyflow_form') or 'unknown form'}")
    for s in signals:
        if not is_excluded(s):
            return s
    if signals:
        return signals[0]          # everything same-day is excluded; return it so evaluate() drops the lead

    # No same-day signal (older contact re-added, Heyflow date not updated, etc.): best-effort label.
    if props.get("hs_object_source_detail_1"):
        return props["hs_object_source_detail_1"]
    if props.get("recent_conversion_event_name"):
        return props["recent_conversion_event_name"].rsplit(": ", 1)[-1]
    return "unknown"


def evaluate(contact_id: str, added_at: datetime, props: dict, owners: dict, dispositions: dict, run_at: datetime):
    """
    Score one lead. Returns a result dict keyed by SHEET_HEADERS (plus private
    "_" keys the email uses), or None if the source is excluded.
    """
    source = describe_source(props, added_at)
    if is_excluded(source):
        return None

    calls = get_calls_after(contact_id, added_at)
    first = calls[0] if calls else None

    name = f"{props.get('firstname', '') or ''} {props.get('lastname', '') or ''}".strip() or props.get("email", "") or contact_id
    owner_name = owners.get(str(props.get("hubspot_owner_id") or ""), "(unassigned)")
    elapsed = business_hours_between(added_at, run_at)

    if first:
        hrs = business_hours_between(added_at, first["ts"])
        tier = tier_for(hrs)
        status = "CALLED"
        first_call_at = fmt_dt(first["ts"])
        call_owner = owners.get(str(first.get("hubspot_owner_id") or ""), "")
        direction = first.get("hs_call_direction") or ""
        outcome = dispositions.get(first.get("hs_call_disposition") or "", first.get("hs_call_disposition") or "") or (first.get("hs_call_status") or "")
    else:
        hrs = ""
        tier = tier_for(elapsed) if elapsed > 0 else TIER_AFTER_HOURS
        status = "NOT CALLED"
        first_call_at = call_owner = direction = outcome = ""

    hs_rt = props.get("hs_time_to_first_engagement")
    hs_rt_hrs = round(int(hs_rt) / 3_600_000, 1) if hs_rt and str(hs_rt).lstrip("-").isdigit() else ""

    return {
        "Run Date": run_at.strftime("%Y-%m-%d"),
        "Lead Added At": fmt_dt(added_at),
        "Contact ID": contact_id,
        "Name": name,
        "Email": props.get("email") or "",
        "Phone": props.get("phone") or props.get("mobilephone") or "",
        "Contact Owner": owner_name,
        "Lead Status": props.get("hs_lead_status") or "",
        "Lead Source": props.get("sales_lead_source_2_0") or "",
        "Form / Event": source,
        "Recent Conversion Date": fmt_dt(parse_hs_datetime(props.get("recent_conversion_date"))),
        "First Call At": first_call_at,
        "Call Owner": call_owner,
        "Call Direction": direction,
        "Call Outcome": outcome,
        "Calls After Lead": len(calls),
        "Biz Hours To First Call": hrs,
        "Biz Hours Elapsed": elapsed,
        "Tier": tier,
        "Status": status,
        "HS Lead Response Time (hrs)": hs_rt_hrs,
        "HubSpot Link": hs_link(contact_id),
        "Last Checked": fmt_dt(run_at),
        # private keys for the email (never written to the sheet)
        "_added": added_at,
        "_first_call": first["ts"] if first else None,
    }


# --------------------------------------------------------------------------- #
# Google Sheets
# --------------------------------------------------------------------------- #
# The sheet is both the output and the memory: NOT CALLED rows from earlier
# days are read back in and re-checked. Row identity = (Contact ID, Lead Added At).

def open_sheet():
    """Authenticate and return the worksheet, creating the tab with headers if needed. None if not configured."""
    if not GOOGLE_SHEET_ID or not GOOGLE_SERVICE_ACCOUNT_JSON:
        log.warning("Google Sheet not configured (GOOGLE_SHEET_ID / GOOGLE_SERVICE_ACCOUNT_JSON missing). Skipping sheet.")
        return None
    import gspread
    from google.oauth2.service_account import Credentials

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet(SHEET_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_TAB, rows=2000, cols=len(SHEET_HEADERS) + 2)
        ws.append_row(SHEET_HEADERS, value_input_option="USER_ENTERED")
        ws.freeze(rows=1)
    return ws


def load_sheet_rows(ws):
    """Read the whole tab once. Returns (idx {header: col}, rows [(row_number, {header: value})])."""
    values = ws.get_all_values()
    if not values:
        ws.append_row(SHEET_HEADERS, value_input_option="USER_ENTERED")
        ws.freeze(rows=1)
        return {h: i for i, h in enumerate(SHEET_HEADERS)}, []
    header = values[0]
    idx = {h: i for i, h in enumerate(header)}
    rows = []
    for n, raw in enumerate(values[1:], start=2):
        d = {h: (raw[i] if i < len(raw) else "") for h, i in idx.items()}
        rows.append((n, d))
    return idx, rows


def open_leads_from_sheet(rows, run_at: datetime):
    """
    From existing rows, pick the NOT CALLED ones within RECHECK_DAYS to re-check.
    Rows whose Form / Event now matches EXCLUDE_FORMS are not re-checked; their
    row numbers come back separately so main() can mark them EXCLUDED.

    Returns ({contact_id: added_at}, [row_numbers_to_exclude]).
    """
    cutoff = run_at - timedelta(days=RECHECK_DAYS)
    carried, to_exclude = {}, []
    for n, d in rows:
        if d.get("Status") != "NOT CALLED":
            continue
        added = parse_sheet_dt(d.get("Lead Added At", ""))
        cid = d.get("Contact ID", "").strip()
        if not (added and cid and added >= cutoff):
            continue
        if is_excluded(d.get("Form / Event", "")):
            to_exclude.append(n)
            continue
        carried[cid] = added
    return carried, to_exclude


def mark_rows_status(ws, idx, row_numbers, status: str, run_at: datetime):
    """Set Status (and Last Checked) on specific rows. Used to flip old rows to EXCLUDED."""
    if not row_numbers:
        return
    status_col = _col_letter(idx["Status"] + 1)
    checked_col = _col_letter(idx["Last Checked"] + 1) if "Last Checked" in idx else None
    updates = []
    for n in row_numbers:
        updates.append({"range": f"{status_col}{n}", "values": [[status]]})
        if checked_col:
            updates.append({"range": f"{checked_col}{n}", "values": [[fmt_dt(run_at)]]})
    ws.batch_update(updates, value_input_option="USER_ENTERED")
    log.info("Sheet: %d rows marked %s", len(row_numbers), status)


def write_results(ws, idx, existing_rows, results: list):
    """Update rows that already exist (same Contact ID + Lead Added At) in place, append the rest."""
    key_to_rownum = {}
    for n, d in existing_rows:
        key_to_rownum[(d.get("Contact ID", "").strip(), d.get("Lead Added At", "").strip())] = n

    header = [h for h, _ in sorted(idx.items(), key=lambda kv: kv[1])]
    last_col = _col_letter(len(header))

    updates, appends = [], []
    for r in results:
        values = [r.get(h, "") for h in header]
        key = (str(r["Contact ID"]), r["Lead Added At"])
        if key in key_to_rownum:
            n = key_to_rownum[key]
            updates.append({"range": f"A{n}:{last_col}{n}", "values": [values]})
        else:
            appends.append(values)

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
    if appends:
        ws.append_rows(appends, value_input_option="USER_ENTERED")
    log.info("Sheet: %d rows updated, %d rows appended", len(updates), len(appends))


def _col_letter(n: int) -> str:
    """1 -> 'A', 26 -> 'Z', 27 -> 'AA'."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# --------------------------------------------------------------------------- #
# Email
# --------------------------------------------------------------------------- #
# Built from nested tables with inline styles only, because that's what
# Outlook and Gmail actually render. No <style> block, no flexbox, no CSS grid.

FONT = "font-family:Arial,Helvetica,sans-serif;"


def _esc(s) -> str:
    return (str(s) if s is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _pill(tier: str) -> str:
    bg, fg, label = TIER_STYLE.get(tier, ("#9E9E9E", "#FFFFFF", tier))
    return (f"<span style='display:inline-block;padding:3px 9px;border-radius:10px;background:{bg};"
            f"color:{fg};font-size:11px;font-weight:bold;white-space:nowrap;{FONT}'>{_esc(label)}</span>")


def _hrs(v) -> str:
    try:
        return f"{float(v):.1f}"
    except (TypeError, ValueError):
        return ""


def _cell(content, align="left", extra="") -> str:
    return (f"<td align='{align}' style='padding:9px 10px;border-bottom:1px solid {C_LINE};"
            f"font-size:13px;color:{C_TEXT};vertical-align:top;{FONT}{extra}'>{content}</td>")


def _th(label, align="left") -> str:
    return (f"<th align='{align}' style='padding:6px 10px;font-size:11px;font-weight:bold;color:{C_MUTED};"
            f"text-transform:uppercase;letter-spacing:.4px;border-bottom:2px solid {C_LINE};{FONT}'>{_esc(label)}</th>")


NOWRAP = "white-space:nowrap;"


def _name_link(r) -> str:
    return f"<a href='{r['HubSpot Link']}' style='color:{C_LINK};font-weight:bold;text-decoration:none'>{_esc(r['Name'])}</a>"


def _table(header_cells: str, body_rows: str) -> str:
    return (f"<table width='100%' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse'>"
            f"<tr>{header_cells}</tr>{body_rows}</table>")


def _section_title(text: str, count=None, color=C_TEXT) -> str:
    badge = ""
    if count is not None:
        badge = (f" <span style='display:inline-block;margin-left:6px;padding:2px 8px;border-radius:10px;"
                 f"background:{C_OWNER_BG};color:{C_TEXT};font-size:12px;font-weight:bold'>{count}</span>")
    return (f"<tr><td style='padding:22px 0 8px;font-size:16px;font-weight:bold;color:{color};{FONT}'>"
            f"{_esc(text)}{badge}</td></tr>")


def _owner_header(owner: str, count: int) -> str:
    return (f"<tr><td colspan='6' style='padding:10px 10px 6px;background:{C_OWNER_BG};font-size:13px;"
            f"font-weight:bold;color:{C_TEXT};{FONT}'>{_esc(owner)} "
            f"<span style='color:{C_MUTED};font-weight:normal'>({count})</span></td></tr>")


def _stat_tile(number, label, color=C_TEXT) -> str:
    return (f"<td width='20%' align='center' style='padding:12px 4px;background:{C_CARD_BG};border:1px solid {C_LINE};"
            f"border-radius:6px'>"
            f"<div style='font-size:24px;font-weight:bold;color:{color};{FONT}'>{number}</div>"
            f"<div style='font-size:11px;color:{C_MUTED};text-transform:uppercase;letter-spacing:.4px;{FONT}'>{_esc(label)}</div>"
            f"</td>")


def build_email(results: list, run_at: datetime, report_day: date) -> tuple:
    """
    Returns (subject, html). Two halves:
      1. Report day: leads whose lead time is after the previous business close.
         Not called (grouped by owner) -> called late (flat) -> on time (one line).
      2. Backlog: NOT CALLED leads from before that, flat, oldest first.
    """
    cutoff = previous_business_close(report_day)
    today_rows = [r for r in results if r["_added"] > cutoff]
    backlog = [r for r in results if r["_added"] <= cutoff and r["Status"] == "NOT CALLED"]

    t_uncalled = [r for r in today_rows if r["Status"] == "NOT CALLED"]
    t_late = [r for r in today_rows if r["Status"] == "CALLED" and r["Tier"] != TIER_OK]
    t_ontime = [r for r in today_rows if r["Status"] == "CALLED" and r["Tier"] == TIER_OK]

    day_label = report_day.strftime("%a %b %d")          # "Mon Nov 09"
    day_short = f"{report_day.month}/{report_day.day}"    # "11/9"
    subject = f"Lead callback {day_short}: {len(t_uncalled)} not called, {len(t_late)} late, {len(backlog)} in backlog"

    parts = []

    # ---- outer wrapper ----
    parts.append(f"<body style='margin:0;padding:0;background:{C_PAGE_BG}'>"
                 f"<table width='100%' cellpadding='0' cellspacing='0' border='0' style='background:{C_PAGE_BG}'>"
                 f"<tr><td align='center' style='padding:20px 10px'>"
                 f"<table width='720' cellpadding='0' cellspacing='0' border='0' style='max-width:720px;width:100%'>")

    # ---- header bar ----
    parts.append(f"<tr><td style='background:{C_HEADER_BG};padding:18px 22px;border-radius:8px 8px 0 0'>"
                 f"<div style='font-size:20px;font-weight:bold;color:{C_HEADER_TEXT};{FONT}'>Lead Callback Check</div>"
                 f"<div style='font-size:13px;color:#C9D4E3;margin-top:3px;{FONT}'>Report for {_esc(day_label)}"
                 f" &nbsp;&middot;&nbsp; sent {_esc(fmt_short(run_at, True))}</div></td></tr>")

    # ---- card ----
    parts.append(f"<tr><td style='background:{C_CARD_BG};padding:18px 22px 24px;border-radius:0 0 8px 8px'>"
                 f"<table width='100%' cellpadding='0' cellspacing='0' border='0'>")

    # stat tiles
    parts.append("<tr><td><table width='100%' cellpadding='0' cellspacing='6' border='0'><tr>"
                 + _stat_tile(len(today_rows), "Leads in")
                 + _stat_tile(len(t_ontime), "On time", "#2E7D32")
                 + _stat_tile(len(t_late), "Called late", "#EF6C00")
                 + _stat_tile(len(t_uncalled), "Not called", "#C62828")
                 + _stat_tile(len(backlog), "Backlog", "#6B7280")
                 + "</tr></table></td></tr>")

    # ---- section 1a: not called, grouped by owner ----
    parts.append(_section_title(f"{day_label}: not called yet", len(t_uncalled), "#C62828"))
    if t_uncalled:
        by_owner = {}
        for r in t_uncalled:
            by_owner.setdefault(r["Contact Owner"], []).append(r)
        body = ""
        for owner in sorted(by_owner):
            body += _owner_header(owner, len(by_owner[owner]))
            for r in sorted(by_owner[owner], key=lambda r: r["_added"]):
                body += ("<tr>" + _cell(_name_link(r), extra=NOWRAP) + _cell(_esc(r["Phone"]), extra=NOWRAP)
                         + _cell(_esc(fmt_short(r["_added"])), extra=NOWRAP) + _cell(_esc(r["Form / Event"]), extra=f"color:{C_MUTED};")
                         + _cell(_hrs(r["Biz Hours Elapsed"]), "right") + _cell(_pill(r["Tier"]), "center") + "</tr>")
        parts.append("<tr><td>" + _table(_th("Name") + _th("Phone") + _th("Came in") + _th("Source")
                                        + _th("Hrs waiting", "right") + _th("Tier", "center"), body) + "</td></tr>")
    else:
        parts.append(f"<tr><td style='padding:6px 10px;font-size:13px;color:{C_MUTED};{FONT}'>Everyone got a call. Nice.</td></tr>")

    # ---- section 1b: called late ----
    parts.append(_section_title(f"{day_label}: called, but late", len(t_late), "#EF6C00"))
    if t_late:
        body = ""
        for r in sorted(t_late, key=lambda r: -float(r["Biz Hours To First Call"] or 0)):
            body += ("<tr>" + _cell(_name_link(r), extra=NOWRAP) + _cell(_esc(r["Contact Owner"]), extra=NOWRAP)
                     + _cell(_esc(fmt_short(r["_added"])), extra=NOWRAP) + _cell(_esc(fmt_short(r["_first_call"])), extra=NOWRAP)
                     + _cell(_hrs(r["Biz Hours To First Call"]), "right") + _cell(_pill(r["Tier"]), "center") + "</tr>")
        parts.append("<tr><td>" + _table(_th("Name") + _th("Owner") + _th("Came in") + _th("First call")
                                        + _th("Hrs to call", "right") + _th("Tier", "center"), body) + "</td></tr>")
    else:
        parts.append(f"<tr><td style='padding:6px 10px;font-size:13px;color:{C_MUTED};{FONT}'>None.</td></tr>")

    # ---- section 1c: on time, one line ----
    if t_ontime:
        by_owner = {}
        for r in t_ontime:
            by_owner[r["Contact Owner"]] = by_owner.get(r["Contact Owner"], 0) + 1
        line = ", ".join(f"{_esc(o)} ({n})" for o, n in sorted(by_owner.items()))
        parts.append(f"<tr><td style='padding:14px 10px 0;font-size:13px;color:{C_TEXT};{FONT}'>"
                     f"<span style='color:#2E7D32;font-weight:bold'>Called within the hour:</span> {line}</td></tr>")

    # ---- section 2: backlog ----
    parts.append(f"<tr><td style='padding-top:22px'><hr style='border:0;border-top:1px solid {C_LINE};margin:0'></td></tr>")
    parts.append(_section_title(f"Still not called from before {day_short}", len(backlog), C_TEXT))
    if backlog:
        body = ""
        for r in sorted(backlog, key=lambda r: r["_added"]):
            body += ("<tr>" + _cell(_name_link(r), extra=NOWRAP) + _cell(_esc(r["Contact Owner"]), extra=NOWRAP) + _cell(_esc(r["Phone"]), extra=NOWRAP)
                     + _cell(_esc(fmt_short(r["_added"], True)), extra=NOWRAP) + _cell(_esc(r["Form / Event"]), extra=f"color:{C_MUTED};")
                     + _cell(_hrs(r["Biz Hours Elapsed"]), "right") + "</tr>")
        parts.append("<tr><td>" + _table(_th("Name") + _th("Owner") + _th("Phone") + _th("Came in") + _th("Source")
                                        + _th("Hrs waiting", "right"), body) + "</td></tr>")
    else:
        parts.append(f"<tr><td style='padding:6px 10px;font-size:13px;color:{C_MUTED};{FONT}'>Backlog is clear.</td></tr>")

    # ---- footer ----
    sheet_link = f" &nbsp;&middot;&nbsp; <a href='{SHEET_URL}' style='color:{C_LINK}'>Open the sheet</a>" if SHEET_URL else ""
    excl = f" Excluded: {_esc(', '.join(EXCLUDE_FORMS))}." if EXCLUDE_FORMS else ""
    parts.append(f"<tr><td style='padding-top:22px;font-size:11px;color:{C_MUTED};line-height:1.5;{FONT}'>"
                 f"Lead time is the moment the contact joined the HubSpot list. A call is any call logged on the contact after that. "
                 f"Hours count Mon-Fri {BUSINESS_START}:00-{BUSINESS_END}:00 {_esc(TZ.key)} only; weekend and after-hours leads "
                 f"start their clock at the next open. Tiers: On time &lt;1h, Flag 1 1-3h, Flag 2 3h+, Flag 3 1+ business day.{excl}"
                 f"{sheet_link}</td></tr>")

    parts.append("</table></td></tr></table></td></tr></table></body>")
    return subject, "".join(parts)


def send_email(subject: str, html: str):
    """Send via Gmail SMTP (smtp.gmail.com:587, STARTTLS, App Password)."""
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        log.warning("Gmail not configured (GMAIL_USER / GMAIL_APP_PASSWORD missing). Skipping email.")
        return
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(EMAIL_TO)
    msg.attach(MIMEText("This report is HTML. Open in an HTML-capable mail client.", "plain"))
    msg.attach(MIMEText(html, "html"))
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
        s.ehlo()
        s.starttls()
        s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        s.sendmail(GMAIL_USER, EMAIL_TO, msg.as_string())
    log.info("Email sent to %s", EMAIL_TO)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main():
    if not HUBSPOT_TOKEN:
        log.error("HUBSPOT_TOKEN is not set")
        sys.exit(1)

    run_at = now_tz()
    report_day = report_day_for(run_at)
    window_start = previous_business_close(report_day)
    log.info("Run at %s, report day %s, window opens %s (DRY_RUN=%s)",
             fmt_dt(run_at), report_day, fmt_dt(window_start), DRY_RUN)

    # 1. List members with the time they joined
    members = get_list_memberships(LIST_ID)
    log.info("List %s has %d members", LIST_ID, len(members))
    if REPORT_DAY_ONLY:
        # The list's "less than 1 day ago" filter really means yesterday + today, so trim it
        # to the report-day window. Older uncalled leads come back via the sheet carry-forward.
        members = {cid: added for cid, added in members.items() if added > window_start}
        log.info("%d of those fall in the report-day window", len(members))

    # 2. Leads from earlier runs still waiting on a call
    ws = None if DRY_RUN else open_sheet()
    idx, existing_rows = (None, [])
    carried, to_exclude = {}, []
    if ws:
        idx, existing_rows = load_sheet_rows(ws)
        carried, to_exclude = open_leads_from_sheet(existing_rows, run_at)
        for cid in list(carried):
            if cid in members:          # re-submitted today; today's entry wins
                carried.pop(cid)
        log.info("Carrying forward %d uncalled leads, %d old rows to mark EXCLUDED", len(carried), len(to_exclude))
        mark_rows_status(ws, idx, to_exclude, "EXCLUDED", run_at)

    targets = {**carried, **members}
    if not targets:
        log.info("Nothing to check.")
        if not DRY_RUN:
            send_email(f"Lead callback {report_day.month}/{report_day.day}: no leads",
                       "<p>The list was empty at run time and the backlog is clear.</p>")
        return

    # 3. Context lookups
    owners = get_owners()
    dispositions = get_call_dispositions()
    contacts = get_contacts(targets.keys())

    # 4. Evaluate
    results, excluded = [], 0
    for cid, added_at in sorted(targets.items(), key=lambda kv: kv[1]):
        props = contacts.get(cid)
        if props is None:
            log.warning("Contact %s not found (merged or deleted?), skipping", cid)
            continue
        r = evaluate(cid, added_at, props, owners, dispositions, run_at)
        if r is None:
            excluded += 1
            continue
        results.append(r)
        time.sleep(0.25)
    log.info("%d leads evaluated, %d excluded by EXCLUDE_FORMS, %d not called",
             len(results), excluded, sum(1 for r in results if r["Status"] == "NOT CALLED"))

    # 5. Output
    if DRY_RUN:
        cols = ["Name", "Contact Owner", "Lead Added At", "Form / Event", "First Call At",
                "Biz Hours To First Call", "Biz Hours Elapsed", "Tier", "Status"]
        print("\t".join(cols))
        for r in results:
            print("\t".join(str(r.get(c, "")) for c in cols))
        subject, html = build_email(results, run_at, report_day)
        print(f"\n[DRY RUN] Would email: {subject}")
        out = os.environ.get("DRY_RUN_HTML")      # optional: save the email to a file to eyeball it
        if out:
            with open(out, "w", encoding="utf-8") as f:
                f.write(html)
            print(f"[DRY RUN] Email HTML written to {out}")
        return

    if ws:
        write_results(ws, idx, existing_rows, results)

    subject, html = build_email(results, run_at, report_day)
    send_email(subject, html)


if __name__ == "__main__":
    main()
