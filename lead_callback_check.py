#!/usr/bin/env python3
"""
=============================================================================
 Lead Callback Check
=============================================================================

WHAT THIS DOES
--------------
Runs once at the end of each business day and answers one question for every
lead that came in today: "did a loan officer call this person, and how fast?"

The flow, top to bottom:

    1. HubSpot list  ->  who submitted a form today, and WHEN they were added
                         to the list. That add time is the "lead time".
    2. Google Sheet  ->  any leads from previous days that are still marked
                         NOT CALLED get pulled back in and re-checked.
    3. HubSpot       ->  for each lead, pull the calls logged on the contact
                         that happened AFTER the lead time.
    4. Math          ->  business hours between lead time and the first call,
                         bucketed into a tier (OK / Flag 1 / Flag 2 / Flag 3).
                         No call at all = NOT CALLED.
    5. Google Sheet  ->  one row per lead. Existing rows (same contact + same
                         lead time) are updated in place, new ones appended.
    6. Email         ->  the NOT CALLED list grouped by contact owner, plus a
                         short "called but late" section, sent via Gmail SMTP.

HOW TO READ THIS FILE
---------------------
The file is split into sections with big banner comments:

    Config            all the knobs, read from environment variables
    Small helpers     date parsing/formatting and the business-hours math
    HubSpot           every call to the HubSpot API lives here
    Evaluation        turns one contact + its calls into one result row
    Google Sheets     reading the Master tab and writing results back
    Email             builds the HTML and sends it
    Main              glues the sections together in order

If you want to change a rule (what counts as a call, the tier cutoffs, the
business hours), it's almost always in Config, tier_for(), or get_calls_after().

ENVIRONMENT VARIABLES
---------------------
Set as GitHub Secrets / Variables (see README.md), or exported in your shell
for a local run. Only HUBSPOT_TOKEN is strictly required. If the Google or
Gmail variables are missing the script logs a warning and skips that step
instead of crashing, so you can test HubSpot on its own.

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
"""

# --------------------------------------------------------------------------- #
# Imports
# --------------------------------------------------------------------------- #
# Standard library only, except `requests` (HTTP) and, lazily, `gspread` +
# `google-auth` for Sheets. The Google imports live inside open_sheet() so a
# dry run that never touches Sheets doesn't need them installed.

import json                 # building API request bodies, reading the SA key
import logging              # progress output you'll see in the Actions log
import os                   # environment variables
import smtplib              # sending the email over SMTP
import sys                  # sys.exit on fatal config problems
import time                 # small sleeps between API calls
from datetime import date, datetime, timedelta, timezone
from datetime import time as dtime          # renamed so it doesn't clash with the `time` module
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo               # timezone handling (Python 3.9+)

import requests


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
# os.environ.get("NAME", "fallback") reads an env var and uses the fallback if
# it isn't set. Nothing is hard-coded below that you can't override from the
# GitHub repo's Secrets/Variables pages.

# --- HubSpot ---
HUBSPOT_TOKEN = os.environ.get("HUBSPOT_TOKEN", "")
LIST_ID = "19329"        # "Form Submissions Today"
PORTAL_ID = os.environ.get("HUBSPOT_PORTAL_ID", "6624066")  # only used to build record links

# --- Google Sheets ---
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
SHEET_TAB = os.environ.get("SHEET_TAB", "Master")

# --- Email ---
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
# EMAIL_TO can be "a@x.com,b@x.com". Split on commas, strip spaces, drop blanks.
EMAIL_TO = [e.strip() for e in os.environ.get("EMAIL_TO", "daniel.haycraft@capitalfund1.com").split(",") if e.strip()]

# --- Business-hours clock ---
TZ = ZoneInfo(os.environ.get("TIMEZONE", "America/Phoenix"))
BUSINESS_START = int(os.environ.get("BUSINESS_START", "8"))   # 8  = 8:00am
BUSINESS_END = int(os.environ.get("BUSINESS_END", "17"))      # 17 = 5:00pm
WORKDAYS = {0, 1, 2, 3, 4}                                    # Python weekday(): Mon=0 ... Sun=6
DAY_HOURS = BUSINESS_END - BUSINESS_START                     # 9 hours = "one full business day"

# --- Behaviour ---
RECHECK_DAYS = int(os.environ.get("RECHECK_DAYS", "7"))       # keep re-checking NOT CALLED leads this many days
DRY_RUN = os.environ.get("DRY_RUN", "false").strip().lower() in ("1", "true", "yes")

# Logging goes to stdout, which is what you see in the GitHub Actions log.
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("lead_callback_check")

# One requests.Session so the auth header is set once and TCP connections are reused.
HS_BASE = "https://api.hubapi.com"
SESSION = requests.Session()
SESSION.headers.update({"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"})

# Contact properties we ask HubSpot for. Internal names, not labels.
CONTACT_PROPERTIES = [
    "firstname", "lastname", "email", "phone", "mobilephone",
    "hubspot_owner_id", "hs_lead_status", "sales_lead_source_2_0", "contact_type_2_0",
    "createdate", "recent_conversion_date", "recent_conversion_event_name",
    "heyflow_date", "heyflow_form", "hs_object_source_detail_1",
    "hs_time_to_first_engagement",   # HubSpot's own "Lead response time", kept as a sanity-check column
]

# Call engagement properties. hs_timestamp is when the call happened (not when it was logged).
CALL_PROPERTIES = [
    "hs_timestamp", "hs_call_direction", "hs_call_status", "hs_call_disposition",
    "hs_call_title", "hs_call_duration", "hubspot_owner_id", "hs_createdate",
]

# Column order in the Google Sheet. The keys in evaluate()'s result dict must
# match these exactly, because write_results() looks values up by header name.
SHEET_HEADERS = [
    "Run Date", "Lead Added At", "Contact ID", "Name", "Email", "Phone",
    "Contact Owner", "Lead Status", "Lead Source", "Form / Event", "Recent Conversion Date",
    "First Call At", "Call Owner", "Call Direction", "Call Outcome", "Calls After Lead",
    "Biz Hours To First Call", "Biz Hours Elapsed", "Tier", "Status",
    "HS Lead Response Time (hrs)", "HubSpot Link", "Last Checked",
]

# How datetimes are written to the sheet. Also used to read them back, so don't
# change one without the other (they're the same constant on purpose).
DT_FMT = "%Y-%m-%d %H:%M:%S"


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #

def now_tz() -> datetime:
    """Current time as a timezone-aware datetime in the configured TZ (Arizona)."""
    return datetime.now(timezone.utc).astimezone(TZ)


def parse_hs_datetime(value):
    """
    Turn a HubSpot date value into an aware datetime, or None.

    HubSpot is inconsistent: the CRM API returns ISO strings like
    '2026-11-02T15:00:14.907Z', while some properties (and the lists API in
    older versions) return epoch milliseconds as a string like '1793491200000'.
    This handles both. The 'Z' suffix is swapped for '+00:00' because
    datetime.fromisoformat() didn't accept 'Z' before Python 3.11.
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
    """Aware datetime -> 'YYYY-MM-DD HH:MM:SS' in Arizona time. Empty string for None."""
    if not dt:
        return ""
    return dt.astimezone(TZ).strftime(DT_FMT)


def parse_sheet_dt(s: str):
    """The reverse of fmt_dt(): read a sheet cell back into an aware datetime. None if it won't parse."""
    try:
        return datetime.strptime(s.strip(), DT_FMT).replace(tzinfo=TZ)
    except Exception:
        return None


def business_hours_between(start: datetime, end: datetime) -> float:
    """
    Business hours between two datetimes, as a float (e.g. 2.5).

    Only counts time inside BUSINESS_START-BUSINESS_END on WORKDAYS, in TZ.
    Walks day by day from start to end; for each workday it clips the
    [start, end] window to that day's [open, close] window and adds whatever
    overlap is left.

    Examples with 8-5 Mon-Fri:
        Mon 10:00 -> Mon 14:00        = 4.0
        Fri 19:00 -> Mon 09:00        = 1.0   (Friday evening and weekend don't count)
        Mon 16:00 -> Tue 09:00        = 2.0   (1h Monday + 1h Tuesday)
        Mon 08:00 -> Tue 08:00        = 9.0   (exactly one business day)
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
            s = max(start, opens)   # later of "lead came in" and "office opened"
            e = min(end, closes)    # earlier of "call happened" and "office closed"
            if e > s:
                total += (e - s).total_seconds() / 3600.0
        day += timedelta(days=1)
    return round(total, 2)


def tier_for(hours: float) -> str:
    """
    Bucket a business-hour count into a tier label.

    Used two ways:
      - for CALLED leads, `hours` is lead time -> first call
      - for NOT CALLED leads, `hours` is lead time -> now (how overdue they are)

    Change the cutoffs here if the rules change.
    """
    if hours < 1:
        return "OK (<1h)"
    if hours < 3:
        return "Flag 1 (1-3h)"
    if hours < DAY_HOURS:
        return "Flag 2 (3h+)"
    return "Flag 3 (1+ biz day)"


def hs_link(contact_id) -> str:
    """Clickable link to the contact record in HubSpot."""
    return f"https://app.hubspot.com/contacts/{PORTAL_ID}/record/0-1/{contact_id}"


# --------------------------------------------------------------------------- #
# HubSpot
# --------------------------------------------------------------------------- #
# All HubSpot traffic goes through hs_request() so retries and error handling
# live in one place. The functions below each wrap one endpoint.

def hs_request(method: str, path: str, **kwargs):
    """
    Make one HubSpot API call and return the parsed JSON.

    Handles the two things that bite you in a scheduled job:
      - 429 Too Many Requests: sleep for the Retry-After header and try again
      - 5xx server errors: back off a little and retry (up to 6 attempts)
    Anything else that isn't 2xx raises with the response body in the message
    so the Actions log tells you exactly what HubSpot complained about.

    `kwargs` are passed straight to requests (params=..., data=...).
    """
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
    Everyone currently in the list, with the moment they were added.

    Endpoint: GET /crm/v3/lists/{listId}/memberships/join-order
    Each result looks like {"recordId": "259736052305", "membershipTimestamp": "2026-11-02T15:00:14.907Z"}.
    That timestamp is our lead time. It's the same source for HubSpot forms and
    Heyflow, which is why we use it instead of per-form date properties.

    Paginated: HubSpot returns up to 250 per page and a paging.next.after
    cursor when there's more. Loop until there isn't one.

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
    """
    Pull CONTACT_PROPERTIES for a set of contact IDs in one request per 100.

    Endpoint: POST /crm/v3/objects/contacts/batch/read (max 100 IDs per call).

    Returns {contact_id (str): {property_name: value, ...}}.
    A contact that was merged or deleted since the list was read just won't be
    in the result; main() logs and skips those.
    """
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
    """
    Map owner IDs to names so the sheet and email say "Lawrence Harrison"
    instead of "135255920".

    Endpoint: GET /crm/v3/owners/ (paginated).
    Fetched twice, once for active and once for archived owners, so a contact
    still owned by someone who left the company resolves to a name instead of
    "(unassigned)".

    Returns {owner_id (str): "First Last"}.
    """
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
    """
    Map call outcome IDs to labels.

    HubSpot stores a call's outcome (hs_call_disposition) as a GUID like
    '9d9162e7-6cf3-4944-bf63-4dff82258764', not the text "Connected". This
    endpoint (GET /calling/v1/dispositions) returns the lookup table.

    If the token doesn't have access to it, we log a warning and return {} so
    the sheet shows the raw GUID rather than the whole run failing.
    """
    try:
        data = hs_request("GET", "/calling/v1/dispositions")
        return {d["id"]: d.get("label", d["id"]) for d in data}
    except Exception as exc:
        log.warning("Could not load call dispositions: %s", exc)
        return {}


def get_calls_after(contact_id: str, since: datetime) -> list:
    """
    All calls associated with one contact that happened at or after `since`,
    oldest first.

    Endpoint: POST /crm/v3/objects/calls/search
    Two filters ANDed together inside one filterGroup:
      - associations.contact EQ <contact_id>   (calls linked to this contact)
      - hs_timestamp GTE <since in epoch ms>   (on/after the lead time)

    This is the "what counts as a call" rule. Right now: any call, any
    direction, any outcome. To tighten it (e.g. outbound only), add another
    filter here such as {"propertyName": "hs_call_direction", "operator": "EQ", "value": "OUTBOUND"}.

    Limit is 50 because we only need the first one; the count is a bonus column.

    Returns a list of dicts, each with "id", "ts" (aware datetime) and the raw
    call properties merged in.
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
        if ts and ts >= since:     # belt and braces; the API filter already did this
            calls.append({"id": r["id"], "ts": ts, **p})
    calls.sort(key=lambda c: c["ts"])
    return calls


# --------------------------------------------------------------------------- #
# Evaluation
# --------------------------------------------------------------------------- #

def describe_source(props: dict, added_at: datetime) -> str:
    """
    Best guess at WHICH form or event put this contact in the list today.
    Purely informational, it doesn't affect the scoring.

    Order of guesses:
      1. A HubSpot form submission dated the same day as the list add
         -> use its name (e.g. "Capital Fund 1 Credit Auth: Credit Auth Form")
      2. A Heyflow date on the same day -> "Heyflow: LP" / "Heyflow: Main App"
      3. Whatever created the record (hs_object_source_detail_1)
      4. Last known form name, or "unknown"
    """
    added_day = added_at.astimezone(TZ).date()
    rc = parse_hs_datetime(props.get("recent_conversion_date"))
    if rc and rc.astimezone(TZ).date() == added_day and props.get("recent_conversion_event_name"):
        return props["recent_conversion_event_name"]
    hf = parse_hs_datetime(props.get("heyflow_date"))
    if hf and hf.date() == added_day:
        return f"Heyflow: {props.get('heyflow_form') or 'unknown form'}"
    if props.get("hs_object_source_detail_1"):
        return props["hs_object_source_detail_1"]
    return props.get("recent_conversion_event_name") or "unknown"


def evaluate(contact_id: str, added_at: datetime, props: dict, owners: dict, dispositions: dict, run_at: datetime) -> dict:
    """
    Score one lead. This is the heart of the job.

    Inputs:
        contact_id    HubSpot contact ID
        added_at      when they joined the list (the lead time)
        props         the contact's properties from get_contacts()
        owners        owner ID -> name map
        dispositions  call outcome GUID -> label map
        run_at        "now", passed in so every row in a run shares one timestamp

    Logic:
        calls   = calls on this contact after added_at
        if there is at least one:
            hours = business hours from added_at to the FIRST call
            tier  = tier_for(hours)              status = CALLED
        else:
            hours = business hours from added_at to now (how overdue)
            tier  = tier_for(hours), or "After hours" if the clock hasn't started
            status = NOT CALLED

    Returns one dict whose keys are exactly SHEET_HEADERS. That dict becomes a
    row in the sheet and a line in the email.
    """
    calls = get_calls_after(contact_id, added_at)
    first = calls[0] if calls else None

    # "First Last", falling back to email, falling back to the ID.
    name = f"{props.get('firstname', '') or ''} {props.get('lastname', '') or ''}".strip() or props.get("email", "") or contact_id
    owner_name = owners.get(str(props.get("hubspot_owner_id") or ""), "(unassigned)")

    # Business hours from lead time until right now. For CALLED leads this is
    # just context; for NOT CALLED leads it's what the tier is based on.
    elapsed = business_hours_between(added_at, run_at)

    if first:
        hrs = business_hours_between(added_at, first["ts"])
        tier = tier_for(hrs)
        status = "CALLED"
        first_call_at = fmt_dt(first["ts"])
        call_owner = owners.get(str(first.get("hubspot_owner_id") or ""), "")
        direction = first.get("hs_call_direction") or ""
        # Outcome: translate the GUID if we can, else show the GUID, else fall back to call status.
        outcome = dispositions.get(first.get("hs_call_disposition") or "", first.get("hs_call_disposition") or "") or (first.get("hs_call_status") or "")
    else:
        hrs = ""
        # elapsed == 0 means the lead came in after 5pm (or on a weekend) and
        # the business-hours clock hasn't started yet, so no flag.
        tier = tier_for(elapsed) if elapsed > 0 else "After hours (clock not started)"
        status = "NOT CALLED"
        first_call_at = call_owner = direction = outcome = ""

    # HubSpot's native "Lead response time" is stored in milliseconds. Convert to hours.
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
        "Form / Event": describe_source(props, added_at),
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
    }


# --------------------------------------------------------------------------- #
# Google Sheets
# --------------------------------------------------------------------------- #
# The sheet is both the output AND the memory of the job: NOT CALLED rows from
# earlier days are read back in and re-checked, so a lead that gets called on
# Wednesday after coming in Monday has its Monday row flipped to CALLED.
#
# Row identity = (Contact ID, Lead Added At). Same person submitting again on
# a different day gets a second row, which is what you want for reporting.

def open_sheet():
    """
    Authenticate with the service account and return the worksheet object for
    SHEET_TAB, creating the tab (with headers, frozen) if it doesn't exist.
    Returns None (and logs a warning) if the Google env vars aren't set.
    """
    if not GOOGLE_SHEET_ID or not GOOGLE_SERVICE_ACCOUNT_JSON:
        log.warning("Google Sheet not configured (GOOGLE_SHEET_ID / GOOGLE_SERVICE_ACCOUNT_JSON missing). Skipping sheet.")
        return None
    import gspread
    from google.oauth2.service_account import Credentials

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
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
    """
    Read the whole tab once.

    Returns (idx, rows):
        idx   {header_name: column_index}  e.g. {"Run Date": 0, "Lead Added At": 1, ...}
        rows  [(sheet_row_number, {header_name: cell_value}), ...]
              row numbers start at 2 because row 1 is the header

    Reading everything in one call is far cheaper than cell-by-cell reads, and
    the sheet is small (a few hundred rows a month).
    """
    values = ws.get_all_values()
    if not values:
        # Brand-new empty tab: write the header so the rest of the code has something to key off.
        ws.append_row(SHEET_HEADERS, value_input_option="USER_ENTERED")
        ws.freeze(rows=1)
        return {h: i for i, h in enumerate(SHEET_HEADERS)}, []
    header = values[0]
    idx = {h: i for i, h in enumerate(header)}
    rows = []
    for n, raw in enumerate(values[1:], start=2):
        # Pad short rows: gspread trims trailing empty cells, so raw may be shorter than header.
        d = {h: (raw[i] if i < len(raw) else "") for h, i in idx.items()}
        rows.append((n, d))
    return idx, rows


def open_leads_from_sheet(rows, run_at: datetime) -> dict:
    """
    From the rows already in the sheet, pick the ones still waiting on a call
    and young enough to keep chasing (Lead Added At within RECHECK_DAYS).

    Returns {contact_id: added_at} in the same shape as get_list_memberships()
    so main() can merge the two.
    """
    cutoff = run_at - timedelta(days=RECHECK_DAYS)
    out = {}
    for _, d in rows:
        if d.get("Status") != "NOT CALLED":
            continue
        added = parse_sheet_dt(d.get("Lead Added At", ""))
        cid = d.get("Contact ID", "").strip()
        if added and cid and added >= cutoff:
            out[cid] = added
    return out


def write_results(ws, idx, existing_rows, results: list):
    """
    Write result rows to the sheet.

    For each result:
      - if a row with the same (Contact ID, Lead Added At) already exists,
        overwrite that row in place (batch_update)
      - otherwise append it (append_rows)

    Both are done as single bulk API calls, not one call per row, to stay
    inside Google's quota.
    """
    key_to_rownum = {}
    for n, d in existing_rows:
        key_to_rownum[(d.get("Contact ID", "").strip(), d.get("Lead Added At", "").strip())] = n

    # Rebuild the header in column order from idx, in case someone reordered columns in the sheet.
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
    """1 -> 'A', 26 -> 'Z', 27 -> 'AA'. Needed to build A1-style ranges like 'A5:W5'."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# --------------------------------------------------------------------------- #
# Email
# --------------------------------------------------------------------------- #

def _esc(s) -> str:
    """Escape &, <, > so a contact named 'Bob <test>' can't break the HTML."""
    return (str(s) if s is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_email(results: list, run_at: datetime) -> tuple:
    """
    Build the subject line and HTML body.

    Layout:
        header line with counts
        "Not called yet"      grouped by Contact Owner, oldest lead first, name links to HubSpot
        "Called, but late"    anyone CALLED whose tier starts with "Flag", slowest first
        footer                reminds the reader what the clock rules are

    Returns (subject, html).
    """
    uncalled = [r for r in results if r["Status"] == "NOT CALLED"]
    late = [r for r in results if r["Status"] == "CALLED" and r["Tier"].startswith("Flag")]
    today_str = run_at.strftime("%a %b %d, %Y")

    subject = f"Lead callback check {run_at.strftime('%m/%d')}: {len(uncalled)} uncalled of {len(results)}"

    def table(rows, cols):
        """Render a list of result dicts as a simple HTML table showing only `cols`."""
        h = "".join(f"<th style='text-align:left;padding:4px 8px;border-bottom:1px solid #ccc'>{_esc(c)}</th>" for c in cols)
        body = ""
        for r in rows:
            cells = ""
            for c in cols:
                v = r.get(c, "")
                if c == "Name":
                    v = f"<a href='{r['HubSpot Link']}'>{_esc(v)}</a>"
                else:
                    v = _esc(v)
                cells += f"<td style='padding:4px 8px;border-bottom:1px solid #eee'>{v}</td>"
            body += f"<tr>{cells}</tr>"
        return f"<table style='border-collapse:collapse;font-size:13px'><tr>{h}</tr>{body}</table>"

    html = [f"<div style='font-family:Arial,sans-serif;font-size:14px'>"]
    html.append(f"<h2 style='margin-bottom:4px'>Lead callback check, {today_str}</h2>")
    html.append(f"<p>{len(results)} leads checked, <b>{len(uncalled)} not called</b>, {len(late)} called late.</p>")

    if uncalled:
        html.append("<h3>Not called yet</h3>")
        by_owner = {}
        for r in uncalled:
            by_owner.setdefault(r["Contact Owner"], []).append(r)
        for owner in sorted(by_owner):
            rows = sorted(by_owner[owner], key=lambda r: r["Lead Added At"])
            html.append(f"<p style='margin:12px 0 4px'><b>{_esc(owner)}</b> ({len(rows)})</p>")
            html.append(table(rows, ["Name", "Phone", "Lead Added At", "Form / Event", "Biz Hours Elapsed", "Tier"]))
    else:
        html.append("<p>Everyone got a call. Nice.</p>")

    if late:
        html.append("<h3 style='margin-top:24px'>Called, but late</h3>")
        html.append(table(sorted(late, key=lambda r: -float(r["Biz Hours To First Call"] or 0)),
                          ["Name", "Contact Owner", "Lead Added At", "First Call At", "Biz Hours To First Call", "Tier"]))

    html.append("<p style='color:#777;font-size:12px;margin-top:24px'>Lead time = time added to the HubSpot list. "
                "Business hours Mon-Fri "
                f"{BUSINESS_START}:00-{BUSINESS_END}:00 {TZ.key}. Full log in the Google Sheet.</p></div>")
    return subject, "".join(html)


def send_email(subject: str, html: str):
    """
    Send via Gmail SMTP (smtp.gmail.com:587 with STARTTLS).

    Needs a Google App Password, not the account password. The message is
    multipart/alternative with a plain-text stub plus the HTML, which is the
    polite way to send HTML mail. Skips with a warning if Gmail isn't configured.
    """
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
    """
    Orchestrates the whole run. Read this top to bottom and it mirrors the
    numbered flow in the file header.

    DRY_RUN=true short-circuits before any writes: no sheet (and therefore no
    carry-forward, since the sheet is the memory), no email. It prints a
    tab-separated table to stdout instead, which is what you'll see in the
    Actions log.
    """
    if not HUBSPOT_TOKEN:
        log.error("HUBSPOT_TOKEN is not set")
        sys.exit(1)

    run_at = now_tz()
    log.info("Run at %s (DRY_RUN=%s)", fmt_dt(run_at), DRY_RUN)

    # 1. Today's list members with the time they joined the list
    members = get_list_memberships(LIST_ID)
    log.info("List %s has %d members", LIST_ID, len(members))

    # 2. Leads from earlier runs still waiting on a call (needs the sheet)
    ws = None if DRY_RUN else open_sheet()
    idx, existing_rows = (None, [])
    carried = {}
    if ws:
        idx, existing_rows = load_sheet_rows(ws)
        carried = open_leads_from_sheet(existing_rows, run_at)
        for cid in list(carried):
            if cid in members:          # they re-submitted today; today's list entry wins
                carried.pop(cid)
        log.info("Carrying forward %d uncalled leads from earlier runs", len(carried))

    # Merge. If a contact is in both, the later (today's) value from `members` wins.
    targets = {**carried, **members}
    if not targets:
        log.info("Nothing to check today.")
        if not DRY_RUN:
            send_email(f"Lead callback check {run_at.strftime('%m/%d')}: no leads today",
                       "<p>The list was empty at run time. Nothing to check.</p>")
        return

    # 3. Context lookups, done once per run rather than once per lead
    owners = get_owners()
    dispositions = get_call_dispositions()
    contacts = get_contacts(targets.keys())

    # 4. Evaluate each lead, oldest lead time first
    results = []
    for cid, added_at in sorted(targets.items(), key=lambda kv: kv[1]):
        props = contacts.get(cid)
        if props is None:
            log.warning("Contact %s not found (merged or deleted?), skipping", cid)
            continue
        results.append(evaluate(cid, added_at, props, owners, dispositions, run_at))
        time.sleep(0.25)   # one search call per lead; this keeps us well under HubSpot's 4/sec search limit

    uncalled = sum(1 for r in results if r["Status"] == "NOT CALLED")
    log.info("%d leads evaluated, %d not called", len(results), uncalled)

    # 5. Output
    if DRY_RUN:
        cols = ["Name", "Contact Owner", "Lead Added At", "Form / Event", "First Call At",
                "Biz Hours To First Call", "Biz Hours Elapsed", "Tier", "Status"]
        print("\t".join(cols))
        for r in results:
            print("\t".join(str(r.get(c, "")) for c in cols))
        subject, _ = build_email(results, run_at)
        print(f"\n[DRY RUN] Would email: {subject}")
        return

    if ws:
        write_results(ws, idx, existing_rows, results)

    subject, html = build_email(results, run_at)
    send_email(subject, html)


# Standard Python idiom: only run main() when the file is executed directly
# (python lead_callback_check.py), not when it's imported by a test.
if __name__ == "__main__":
    main()
