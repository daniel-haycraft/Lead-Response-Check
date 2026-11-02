# Lead Callback Check

End-of-day job that checks whether every contact who hit the HubSpot
"Form Submissions Today" list got a call after they were added, scores the
response time in business hours, logs it to a Google Sheet, and emails the
uncalled list grouped by contact owner.

Runs Mon-Fri at 6:00pm Arizona via GitHub Actions. Can also be triggered by
hand (Actions tab -> Lead Callback Check -> Run workflow), with a dry-run option.

## How it decides

| Thing | Rule |
|---|---|
| Lead time | The moment the contact was added to list 19329 (Lists API membership timestamp). Same rule for HubSpot forms and Heyflow. |
| Called | Any HubSpot call engagement associated with the **contact** with a timestamp after the lead time. Any direction, any outcome. |
| Clock | Business hours only, Mon-Fri 8:00-17:00 America/Phoenix. A lead that comes in at 7pm starts its clock at 8am the next business day. |
| Tiers | `OK (<1h)` / `Flag 1 (1-3h)` / `Flag 2 (3h+)` / `Flag 3 (1+ biz day)` / `NOT CALLED` |
| Carry forward | Leads still `NOT CALLED` are re-checked on each later run for up to `RECHECK_DAYS` (7) and their sheet row is updated in place. |

## Repo layout

```
lead_callback_check.py
requirements.txt
.github/workflows/lead_callback_check.yml
README.md
```

## Secrets (Settings -> Secrets and variables -> Actions -> Secrets)

| Name | What it is |
|---|---|
| `HUBSPOT_TOKEN` | Private app access token. Scopes: `crm.objects.contacts.read`, `crm.lists.read`, `crm.objects.owners.read`. Calls are read through the contacts scope; if the scope picker shows a separate calls read scope, tick it too. Phase 2 (notes/@mentions) will need `crm.objects.contacts.write`. |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | The full JSON key file for the service account, pasted as one secret. Enable the Google Sheets API on the project, then share the target sheet with the service account's `client_email` (Editor). |
| `GOOGLE_SHEET_ID` | The long ID in the sheet URL: `docs.google.com/spreadsheets/d/<THIS PART>/edit`. |
| `GMAIL_USER` | The Google account that sends the email (daniel.haycraft@capitalfund1.com). |
| `GMAIL_APP_PASSWORD` | A 16-character App Password from Google Account -> Security -> 2-Step Verification -> App passwords. The normal password will not work. If App passwords isn't offered, the Workspace admin has it disabled. |

## Variables (Settings -> Secrets and variables -> Actions -> Variables), all optional

| Name | Default |
|---|---|
| `HUBSPOT_LIST_ID` | `19329` |
| `HUBSPOT_PORTAL_ID` | `6624066` |
| `EMAIL_TO` | `daniel.haycraft@capitalfund1.com` (comma-separate for more) |
| `SHEET_TAB` | `Master` (created automatically with headers if missing) |
| `TIMEZONE` | `America/Phoenix` |
| `BUSINESS_START` / `BUSINESS_END` | `8` / `17` |
| `RECHECK_DAYS` | `7` |

## First run

1. Create the repo, push these files.
2. Add the five secrets.
3. Actions -> Lead Callback Check -> Run workflow -> tick **dry run**. Check the log output. This hits HubSpot only, no sheet or email.
4. Run again without dry run. The Master tab gets created and the email lands.

## Running locally

```bash
pip install -r requirements.txt
export HUBSPOT_TOKEN=pat-na1-...
export DRY_RUN=true
python lead_callback_check.py
```

Add the Google and Gmail variables and drop `DRY_RUN` to do a full run from your machine.

## Sheet columns

Run Date, Lead Added At, Contact ID, Name, Email, Phone, Contact Owner, Lead Status,
Lead Source, Form / Event, Recent Conversion Date, First Call At, Call Owner,
Call Direction, Call Outcome, Calls After Lead, Biz Hours To First Call,
Biz Hours Elapsed, Tier, Status, HS Lead Response Time (hrs), HubSpot Link, Last Checked

`HS Lead Response Time (hrs)` is HubSpot's own `hs_time_to_first_engagement`, there as a sanity check. It counts emails, tasks and meetings too and resets on owner change, so it will not always match.

## Phase 2 (not built yet)

Notes on the contact with an @mention of the owner for Flag 2 / Flag 3 leads, and optionally a task. Needs the contacts write scope.
