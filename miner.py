import os

# Must be before playwright import
if os.getenv("RAILWAY_ENVIRONMENT"):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "/app/pw-browsers"

from playwright.sync_api import sync_playwright
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import json
import time
import random
import re
import sys
import io
import glob
import csv
import pandas as pd
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

print("Imports done.", flush=True)

dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path)

USERNAME = os.getenv("MINER_USER")
PASSWORD = os.getenv("MINER_PASSWORD")
ADMIN_URL = "https://evolvemedspa.zenoti.com/Admin/Admin.aspx"
COOKIES_FILE = os.path.join(os.path.dirname(__file__), "cookies.json")
TOKEN_FILE = os.path.join(os.path.dirname(__file__), "token.json")
LOGIN_MAX_ATTEMPTS = 3
LOGIN_RETRY_BACKOFF = 5
REDOWNLOAD_MAX_ATTEMPTS = 2
REDOWNLOAD_RETRY_BACKOFF = 10

# Current Stock exports as CSV. Flip to True to download it as a workbook
# (#export_excel_v2) instead; the Excel path and validate_excel() stay in place.
CURRENT_STOCK_AS_EXCEL = False

if not USERNAME or not PASSWORD:
    raise ValueError("MINER_USER and MINER_PASSWORD must be set in the .env file.")

# yesterday = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
# START_DATE = yesterday
# END_DATE = yesterday

today = date.today()
date_to = today - timedelta(days=1)
# Anchor the window on date_to, not today: on the 1st of a month date_to has
# already rolled back into the previous month, so anchoring on today would
# shift the whole range forward by one month (Aug 1 -> Jul 1..Jul 31 instead
# of Jun 1..Jul 31).
first_day_of_period = date_to.replace(day=1)
date_from = first_day_of_period - relativedelta(months=1)

START_DATE = date_from.strftime("%Y-%m-%d")
END_DATE = date_to.strftime("%Y-%m-%d")
BKP_START_DATE = date_to.replace(day=1).strftime("%Y-%m-%d")

print("Date From:", date_from.strftime("%m/%d/%Y"))
print("Date To  :", date_to.strftime("%m/%d/%Y"))

IS_LOCAL = os.getenv("RAILWAY_ENVIRONMENT") is None

DRIVE_FOLDER_ID = "1wKLZcbe8p9Qpgl9g9KZ4G6bGk__JowY5"
DONE_FOLDER_ID = "1icNO-KvNyolmdOAL7d4HSacKVMR72nmz"
REPORT_FOLDERS = {
    "Attendance": "1YKoroJ8l_YSlQGCEBvp9vJIm8sFZOzX0",
    "Cost of Goods": "1M6xHpZAKtBlu6ageNr2KhZYxOTX9iExg",
    "Sales-Cash": "1FXYnXXQiwQxVAu8IBQOm5GROddoBOXwp",
    "Appointments": "12jqbWWMgpgioR_23KJKLXSvDignwrcP2",
    "Sales-Accrual": "1TBdw_u-ADwb3m6GH-HY4WOVYblIPBxd-",
    "Employee Sales": "123hD_j54WtSPIAjdACWFEGaO877lztRs",
    "Business KPI": "1GjkgXcKrGFqa8l9iM-rW8u2MeRVCaB_M",
    "Memberships": "172HJzXYy_9_qtlgTSlZUgUJmZmT-7qwH",
    "Current Stock": "174ZiUaKjIjEKJNKe75mZXKAwya0F4GNK",
    "Stock Ledger": "1JwZGmMBu-3ZHb67edqOZ8vsj5u9eMRd9",
    "FBAds": "1rs8hu18v64Xml3ZQ4F1Mr5uytZ6V5ppC",
    "GoogleAds": "15Cxii7nKW4XXhJNPjAUd2Y819GxneYNa",
}
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

GSHEET_ID = "1ebRZa2y25O5wuPTdbIOeAqokc3FEawToSBVx5yUthfQ"
GSHEET_TABS = {
    "FBAds": {"gid": 1784599697, "folder_id": "1rs8hu18v64Xml3ZQ4F1Mr5uytZ6V5ppC"},
    "GoogleAds": {"gid": 1126535667, "folder_id": "15Cxii7nKW4XXhJNPjAUd2Y819GxneYNa"},
}


_drive_service = None
_drive_creds = None


def get_drive_service():
    from google.auth.transport.requests import Request

    global _drive_service, _drive_creds

    # Load creds once from env; keep them in-memory so a refreshed access token
    # persists across calls (can't save back to env var on Railway).
    if _drive_creds is None:
        token_json = os.getenv("GOOGLE_TOKEN_JSON")
        if not token_json:
            print("GOOGLE_TOKEN_JSON not set. Skipping upload.")
            return None
        creds_info = json.loads(token_json)
        _drive_creds = Credentials.from_authorized_user_info(creds_info, SCOPES)

    # Refresh only when the live cached creds are actually expired (~once/hour),
    # not on every report. Rebuild the service after a refresh.
    if _drive_creds.expired and _drive_creds.refresh_token:
        print("Refreshing Google OAuth token...")
        _drive_creds.refresh(Request())
        _drive_service = None
        print("Token refreshed.")

    if _drive_service is None:
        _drive_service = build("drive", "v3", credentials=_drive_creds)

    return _drive_service

def upload_to_drive(filepath, folder_id=DRIVE_FOLDER_ID):
    service = get_drive_service()
    if not service:
        return None

    filename = os.path.basename(filepath)

    # Phase 1 (moving existing files to Done) runs upfront in
    # move_existing_reports_to_done() before any download, so no pre-move here.
    print(f"Uploading new file: {filename}")
    file_metadata = {"name": filename, "parents": [folder_id]}
    # Pick the mimetype from the extension: Current Stock uploads a workbook, so a
    # hardcoded text/csv would make Drive mislabel (and fail to preview) it.
    mimetypes_by_ext = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls": "application/vnd.ms-excel",
        ".csv": "text/csv",
    }
    ext = os.path.splitext(filename)[1].lower()
    mimetype = mimetypes_by_ext.get(ext, "application/octet-stream")
    media = MediaFileUpload(filepath, mimetype=mimetype, resumable=True)
    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,webViewLink",
    ).execute()
    time.sleep(3)

    print(f"Uploaded to Drive: {filename} ({uploaded.get('webViewLink')})")

    if not uploaded or not uploaded.get('id'):
        raise Exception(f"Upload failed: {filename}")

    # Post-upload sweep: the Google API client silently retries files().create()
    # on a lost/timed-out response, which can create a SECOND server-side copy
    # even though this call returned once with no error. Keep the file we just
    # got back; move any other same-name copy to Done (can't delete under the
    # drive.file scope). This is the guard that catches single-run duplicates.
    if folder_id != DONE_FOLDER_ID:
        kept_id = uploaded.get("id")
        copies = service.files().list(
            q=f"name='{filename}' and '{folder_id}' in parents and trashed=false",
            fields="files(id, name)",
        ).execute().get("files", [])
        for c in copies:
            if c["id"] == kept_id:
                continue
            print(f"Duplicate copy detected: moving extra {filename} to Done")
            try:
                service.files().update(
                    fileId=c["id"],
                    addParents=DONE_FOLDER_ID,
                    removeParents=folder_id,
                    fields="id",
                ).execute()
                time.sleep(1)
            except Exception as sweep_err:
                print(f"  Could not move duplicate {filename}: {sweep_err}")

    return uploaded


def list_all_files(service, folder_id):
    """Every file in a folder, following pagination.

    files().list() returns at most pageSize results (Drive v3 defaults to 100),
    so a single call silently misses anything past the first page. Ask for the
    max and follow nextPageToken so callers really do see all files.

    Caveat: under the drive.file scope this only ever returns files created by
    THIS OAuth client. A file uploaded by hand, or by a run using a different
    client_id, is invisible here and therefore unmovable."""
    files = []
    page_token = None
    while True:
        response = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id, name)",
            pageSize=1000,
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            return files


def move_existing_reports_to_done():
    for folder_name, folder_id in REPORT_FOLDERS.items():
        service = get_drive_service()
        if not service:
            return

        existing = list_all_files(service, folder_id)
        if not existing:
            continue

        print(f"{folder_name}: {len(existing)} existing file(s) to move")
        moved = 0
        for old_file in existing:
            print(f"Moving {folder_name}/{old_file['name']} to Done folder")
            # Isolate per file: one un-movable file (permissions, a concurrent
            # delete) shouldn't abandon the rest of this folder or the folders
            # after it. This sweep runs pre-login, so raising here would kill
            # the whole run before a single report is downloaded.
            try:
                service.files().update(
                    fileId=old_file["id"],
                    addParents=DONE_FOLDER_ID,
                    removeParents=folder_id,
                    fields="id, parents",
                    supportsAllDrives=True,
                ).execute()
                moved += 1
                time.sleep(2)
            except Exception as e:
                print(f"  Could not move {folder_name}/{old_file['name']}: {e}")

        if moved != len(existing):
            print(f"  {folder_name}: moved {moved}/{len(existing)}, {len(existing) - moved} left behind")

        # Verify rather than trust the update() response: re-list the folder and
        # report anything still visible. A file that survives here was either
        # re-parented back, or update() reported success without taking effect.
        try:
            remaining = list_all_files(service, folder_id)
        except Exception as e:
            print(f"  {folder_name}: could not verify sweep: {e}")
            continue

        if remaining:
            print(
                f"  WARNING {folder_name}: {len(remaining)} file(s) STILL in the folder "
                f"after the sweep: {[f['name'] for f in remaining]}"
            )
        else:
            print(f"  {folder_name}: verified empty.")


def dedupe_report_folders():
    """Guardrail: if a report folder holds more than one file with the same
    name (e.g. a container restart re-ran the script and re-uploaded), keep one
    and move the extras to Done. Move — not delete — because the drive.file
    scope cannot delete files it did not create."""
    service = get_drive_service()
    if not service:
        return

    for folder_name, folder_id in REPORT_FOLDERS.items():
        files = list_all_files(service, folder_id)

        by_name = {}
        for f in files:
            by_name.setdefault(f["name"], []).append(f)

        for name, dupes in by_name.items():
            if len(dupes) < 2:
                continue
            # keep the first, move the rest to Done
            for dup in dupes[1:]:
                print(f"Duplicate in {folder_name}: moving extra {name} to Done folder")
                try:
                    service.files().update(
                        fileId=dup["id"],
                        addParents=DONE_FOLDER_ID,
                        removeParents=folder_id,
                        fields="id",
                    ).execute()
                    time.sleep(2)
                except Exception as e:
                    print(f"  Could not move duplicate {name}: {e}")


def validate_csv(filepath):
    filename = os.path.basename(filepath)

    if not os.path.exists(filepath):
        raise Exception(f"File not found: {filename}")

    filesize = os.path.getsize(filepath)
    if filesize == 0:
        raise Exception(f"File is empty: {filename}")

    filesize_mb = filesize / (1024 * 1024)
    print(f"  File size: {filesize_mb:.2f} MB")

    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        head = f.read(1024)

    if "<html" in head.lower() or "<!doctype" in head.lower():
        raise Exception(f"File is HTML, not CSV (possible error page): {filename}")

    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        try:
            headers = next(reader)
        except StopIteration:
            raise Exception(f"CSV has no headers: {filename}")

        if len(headers) < 2:
            raise Exception(f"CSV has only {len(headers)} column(s), likely corrupt: {filename}")

        row_count = 0
        for row in reader:
            row_count += 1
            if row_count >= 5:
                break

        if row_count == 0:
            raise Exception(f"CSV has headers but no data rows: {filename}")

    print(f"  CSV valid: {len(headers)} columns, {row_count}+ data rows")
    return True


def validate_excel(filepath):
    """Sanity-check a downloaded workbook.

    Zenoti sometimes serves an HTML error page with a spreadsheet filename, so
    check the magic bytes: .xlsx is a zip ("PK"), legacy .xls is OLE2. Row/column
    counts are only checked when openpyxl is installed.
    """
    filename = os.path.basename(filepath)

    if not os.path.exists(filepath):
        raise Exception(f"File not found: {filename}")

    filesize = os.path.getsize(filepath)
    if filesize == 0:
        raise Exception(f"File is empty: {filename}")

    print(f"  File size: {filesize / (1024 * 1024):.2f} MB")

    with open(filepath, "rb") as f:
        magic = f.read(8)

    if magic[:2] == b"PK":
        kind = "xlsx"
    elif magic == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        kind = "xls"
    else:
        head = magic.decode("utf-8", errors="replace").lower()
        if "<html" in head or "<!doc" in head or magic.lstrip()[:1] == b"<":
            raise Exception(f"File is HTML, not a workbook (possible error page): {filename}")
        raise Exception(f"File is not a valid Excel workbook: {filename}")

    if kind == "xls":
        # Legacy format; openpyxl cannot read it. Magic bytes are as far as we go.
        print(f"  Workbook valid: legacy .xls, {filesize} bytes")
        return True

    try:
        from openpyxl import load_workbook
    except ImportError:
        print("  Workbook valid: xlsx container OK (install openpyxl for row checks)")
        return True

    wb = load_workbook(filepath, read_only=True)
    try:
        ws = wb.active
        rows = ws.max_row or 0
        cols = ws.max_column or 0
        if cols < 2:
            raise Exception(f"Workbook has only {cols} column(s), likely corrupt: {filename}")
        if rows < 2:
            raise Exception(f"Workbook has headers but no data rows: {filename}")
    finally:
        wb.close()

    print(f"  Workbook valid: {cols} columns, {rows} rows")
    return True


def cleanup_old_csvs():
    script_dir = os.path.dirname(__file__) or "."
    # Current Stock now downloads as a workbook, so sweep those too — otherwise
    # a stale current_stock_*.xlsx sits in the repo root forever.
    for pattern in ("*.csv", "*.xlsx", "*.xls"):
        for f in glob.glob(os.path.join(script_dir, pattern)):
            if date_to.strftime("%Y-%m-%d") not in os.path.basename(f):
                os.remove(f)
                print(f"Cleaned up old report: {f}")


def create_browser_and_context(pw):
    launch_args = {
        "headless": True,
        "args": [
            "--start-maximized",
            "--disable-blink-features=AutomationControlled",
        ],
    }

    if IS_LOCAL:
        launch_args["channel"] = "chrome"
    else:
        launch_args["args"] += [
            "--no-sandbox",
            "--disable-dev-shm-usage",
        ]

    browser = pw.chromium.launch(**launch_args)

    context_args = {"no_viewport": True, "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"}
    if os.path.exists(COOKIES_FILE):
        print(f"Loading saved cookies from {COOKIES_FILE}")
        context_args["storage_state"] = COOKIES_FILE

    context = browser.new_context(**context_args)
    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        window.chrome = {runtime: {}};
    """)
    return browser, context


def save_cookies(context):
    context.storage_state(path=COOKIES_FILE)
    print(f"Cookies saved to {COOKIES_FILE}")


def needs_login(page):
    # print(f"Checking if login is needed. Current URL: {page.url}")
    page.goto(ADMIN_URL, wait_until="domcontentloaded")
    try:
        page.wait_for_url("**/Admin/**", timeout=10000)
        return False
    except:
        return True


def _login_attempt(page):
    print(f"Current URL before login: {page.url}")

    username_sel = "input#Username, input[name='Username'], input[name='username'], input[type='email']"
    if not page.locator(username_sel).first.is_visible():
        print("Login form not visible, navigating to admin...")
        page.goto(ADMIN_URL, wait_until="networkidle")

    try:
        page.wait_for_selector(username_sel, state="visible", timeout=30000)
    except Exception as e:
        print(f"Login form not found. URL: {page.url}")
        print(f"Page title: {page.title()}")
        print(f"Page content preview: {page.content()[:1000]}")
        raise e
    print(f"Login page loaded. URL: {page.url}")

    page.locator(username_sel).first.click()
    page.locator(username_sel).first.fill("")
    page.locator(username_sel).first.press_sequentially(USERNAME, delay=50)
    time.sleep(random.uniform(0.5, 1.5))
    print("Username entered.")

    page.locator('#Password').click()
    page.locator('#Password').fill("")
    page.locator('#Password').press_sequentially(PASSWORD, delay=50)
    time.sleep(random.uniform(0.5, 1.5))
    print("Password entered.")
    time.sleep(2)

    login_button = page.locator('#btnLogin')
    print("Waiting for login button...")
    try:
        login_button.click(timeout=10000)
    except:
        print("Button disabled (captcha pending). Forcing submit via JS...")
        page.evaluate("document.getElementById('btnLogin').removeAttribute('disabled')")
        page.evaluate("document.getElementById('btnLogin').click()")
    print("Login button clicked.")
    time.sleep(random.uniform(2.0, 3.0))

    page.wait_for_url("**/Admin/**", timeout=30000)
    print("Login successful!")


def do_login(page):
    """Log in, retrying up to LOGIN_MAX_ATTEMPTS times.

    Login is a single point of failure for the whole run: all three call sites
    abandon their remaining work if it raises. Transient causes (slow login
    form, captcha interstitial, a dropped nav) usually clear on a second try."""
    last_error = None

    for attempt in range(1, LOGIN_MAX_ATTEMPTS + 1):
        print(f"Login attempt {attempt}/{LOGIN_MAX_ATTEMPTS}...")
        try:
            _login_attempt(page)
            return
        except Exception as e:
            last_error = e
            print(f"Login attempt {attempt}/{LOGIN_MAX_ATTEMPTS} failed: {e}")

        if attempt == LOGIN_MAX_ATTEMPTS:
            break

        backoff = LOGIN_RETRY_BACKOFF * attempt
        print(f"Retrying login in {backoff}s...")
        time.sleep(backoff)

        # Reload the login page so the next attempt starts from a clean form
        # rather than whatever half-submitted state the failure left behind.
        try:
            page.goto(ADMIN_URL, wait_until="networkidle", timeout=60000)
        except Exception as nav_err:
            print(f"  Could not reload login page: {nav_err}")
            continue

        # The submit may have actually gone through and only the post-login
        # wait timed out — same check needs_login() uses.
        if "/Admin/" in page.url:
            print("Already authenticated after reload. Login complete.")
            return

    raise Exception(
        f"Login failed after {LOGIN_MAX_ATTEMPTS} attempts: {last_error}"
    ) from last_error


def wait_for_dashboard(page):
    try:
        page.locator('#menuLinkreports').wait_for(state='visible', timeout=60000)
        print("Dashboard loaded.")
    except Exception as e:
        print(f"Error: Dashboard menu not found. URL: {page.url}")
        raise e


def apply_appointments_filters(report_page):
    print("  Applying Appointments filters...")
    report_page.evaluate("""
        (function() {
            // All multi-selects → selectAll (Centers, Appointment Status, Appointment Source)
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });
            // Date Type (single select) → Appointment Date
            $('select:not([multiple])').each(function() {
                var opts = Array.from(this.options);
                var match = opts.find(function(o) { return o.text.trim().indexOf('Appointment Date') !== -1; });
                if (match) {
                    $(this).multiselect('select', match.value);
                }
            });
        })();
    """)
    time.sleep(2)
    print("  Appointments filters applied.")


def apply_attendance_filters(report_page):
    print("  Applying Attendance filters...")
    report_page.evaluate("""
        (function() {
            // Zenoti dropdowns: Centers + Employee Jobs → All
            ['elm_centers', 'elm_employee_jobs'].forEach(function(id) {
                var cb = document.getElementById(id + '-zenoti-dropdown-options-all');
                if (cb && !cb.checked) cb.click();
            });

            // All multi-selects → selectAll
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });

            // Override: Schedule Status → Working only
            $('select').each(function() {
                var opts = Array.from(this.options);
                if (opts.some(function(o) { return o.value === '6a9d2c87-d452-471f-ba33-90af26ae4edb'; })) {
                    $(this).multiselect('deselectAll', false);
                    $(this).multiselect('select', ['6a9d2c87-d452-471f-ba33-90af26ae4edb']);
                }
            });

            // Single selects: View By → Date, Check-in/Checkout Status → All
            $('select:not([multiple])').each(function() {
                var opts = Array.from(this.options);
                var texts = opts.map(function(o) { return o.text.trim(); });
                if (texts.indexOf('Date') !== -1 && texts.indexOf('Check-in') !== -1) {
                    $(this).multiselect('select', '1');
                } else if (texts.some(function(t) { return t.indexOf('Missed check-ins') !== -1; })) {
                    $(this).multiselect('select', '0');
                }
            });
        })();
    """)
    time.sleep(2)
    print("  Attendance filters applied.")


def apply_cost_of_goods_filters(report_page):
    print("  Applying Cost of Goods filters...")
    report_page.evaluate("""
        (function() {
            // All multi-selects → selectAll (Centers, Product Type, Consumption Type, Brand, Category, Sub Category, Business Unit)
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });
            // Stock Costing Method (single select) → Perpetual Average Cost
            $('select:not([multiple])').each(function() {
                var opts = Array.from(this.options);
                if (opts.some(function(o) { return o.text.indexOf('Perpetual') !== -1; })) {
                    $(this).multiselect('select', '1');
                }
            });
        })();
    """)
    time.sleep(2)
    print("  Cost of Goods filters applied.")


def apply_sales_accrual_filters(report_page):
    print("  Applying Sales-Accrual filters...")
    report_page.evaluate("""
        (function() {
            // Zenoti dropdown: Centers → All
            var cb = document.getElementById('elm_centers-zenoti-dropdown-options-all');
            if (cb && !cb.checked) cb.click();

            // All multi-selects → selectAll (Category, Sub Category, Business Unit, Payment Type, Sale Type, Invoice Status)
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });

            // Override: Item Type → Service + Product only
            var $itemType = $('#elm_item_type');
            if ($itemType.length) {
                $itemType.multiselect('deselectAll', false);
                $itemType.multiselect('select', ['Service', 'Product']);
            }
        })();
    """)
    time.sleep(1)
    selected = report_page.evaluate(
        "Array.from(document.querySelectorAll('#elm_item_type option')).filter(function(o){return o.selected}).map(function(o){return o.value})"
    )
    print(f"  Item Type selected: {selected}")
    time.sleep(2)
    print("  Sales-Accrual filters applied.")


def apply_employee_sales_filters(report_page):
    print("  Applying Employee Sales filters...")
    # Targeted per-id selection, NOT a blanket $('select[multiple]') sweep:
    # the Employee filter is also a multiselect and must stay empty.
    report_page.evaluate("""
        (function() {
            ['elm_allowed_centers', 'elm_payment_type', 'elm_sale_type', 'elm_invoice_status'].forEach(function(id) {
                var $sel = $('#' + id);
                if ($sel.length) {
                    $sel.multiselect('selectAll', false);
                    $sel.multiselect('updateButtonText');
                }
            });

            // Item Type → Service + Product only
            var $itemType = $('#elm_item_type');
            if ($itemType.length) {
                $itemType.multiselect('deselectAll', false);
                $itemType.multiselect('select', ['Service', 'Product']);
                $itemType.multiselect('updateButtonText');
            }
        })();
    """)
    time.sleep(2)

    selected = report_page.evaluate("""
        (function() {
            var out = {};
            ['elm_allowed_centers', 'elm_item_type', 'elm_payment_type', 'elm_sale_type', 'elm_invoice_status'].forEach(function(id) {
                var el = document.getElementById(id);
                out[id] = el ? Array.from(el.options).filter(function(o){return o.selected}).length : null;
            });
            return out;
        })();
    """)
    print(f"  Selected counts: {selected}")
    print("  Employee filter left empty.")
    print("  Employee Sales filters applied.")


def apply_sales_cash_filters(report_page):
    print("  Applying Sales-Cash filters...")
    report_page.evaluate("""
        (function() {
            // Zenoti dropdown: Centers → All
            var cb = document.getElementById('elm_centers-zenoti-dropdown-options-all');
            if (cb && !cb.checked) cb.click();

            // Level of Detail (single select) → Item
            var $lod = $('#elm_level_of_detail');
            if ($lod.length) {
                $lod.multiselect('select', '1');
            }

            // All multi-selects → selectAll (Item Type, Category, Sub Category, Business Unit, Sale Type, Invoice Status)
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });

            // Override: Payment Type → Cash, Card, Check, Custom-Financial, CustomNon-Financial only
            $('select[multiple]').each(function() {
                var values = Array.from(this.options).map(function(o) { return o.value; });
                if (values.indexOf('16') !== -1 && values.indexOf('10') !== -1) {
                    $(this).multiselect('deselectAll', false);
                    $(this).multiselect('select', ['0', '1', '2', '3', '4']);
                }
            });
        })();
    """)
    time.sleep(2)
    print("  Sales-Cash filters applied.")


def apply_business_kpi_filters(report_page):
    print("  Applying Business KPI filters...")
    report_page.evaluate("""
        (function() {
            // Centers → All
            var cb = document.getElementById('elm_centers-zenoti-dropdown-options-all');
            if (cb && !cb.checked) cb.click();

            // Invoice Status → select All
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });

            // Uncheck 'Show Sales Including Tax'
            var taxCb = document.getElementById('elm_include_tax');
            if (taxCb && taxCb.checked) taxCb.click();
        })();
    """)
    time.sleep(2)
    print("  Business KPI filters applied.")


def apply_memberships_filters(report_page):
    print("  Applying Memberships filters...")
    report_page.evaluate("""
        (function() {
            // All multi-selects → selectAll (Sale Centers, Membership Type, Membership Stats)
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });

            // Liability Type → By Sale (value "1")
            $('select:not([multiple])').each(function() {
                var opts = Array.from(this.options);
                if (opts.some(function(o) { return o.text.indexOf('By Sale') !== -1 && o.text.indexOf('By Value') === -1; })) {
                    $(this).multiselect('select', '1');
                }
            });

            // Date Type → Sale Date (value "2")
            $('select:not([multiple])').each(function() {
                var opts = Array.from(this.options);
                if (opts.some(function(o) { return o.text.indexOf('Sale Date') !== -1; }) &&
                    opts.some(function(o) { return o.text.indexOf('Balance As On Date') !== -1; })) {
                    $(this).multiselect('select', '2');
                }
            });

            // Status Type → Membership
            var $statusType = $('#elm_status_type');
            if ($statusType.length) {
                $statusType.multiselect('select', 'Membership');
            }
        })();
    """)
    time.sleep(2)
    print("  Memberships filters applied.")


def apply_current_stock_filters(report_page):
    print("  Applying Current Stock filters...")
    report_page.evaluate("""
        (function() {
            // All multi-selects → selectAll (Centers, Category, Sub Category,
            // Vendor, Brand, Business Unit)
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });

            // Single selects (radios)
            $('select:not([multiple])').each(function() {
                var opts = Array.from(this.options);
                var texts = opts.map(function(o) { return o.text.trim(); });

                // Product Type → All (value "3")
                if (texts.indexOf('Retail') !== -1 && texts.indexOf('Consumable') !== -1) {
                    $(this).multiselect('select', '3');
                }
                // On-Hand Qty → All (value "0")
                else if (texts.indexOf('Greater than 0') !== -1 && texts.indexOf('Less than 0') !== -1) {
                    $(this).multiselect('select', '0');
                }
            });
        })();
    """)
    time.sleep(2)
    print("  Current Stock filters applied.")


def apply_stock_ledger_filters(report_page):
    print("  Applying Stock Ledger filters...")
    report_page.evaluate("""
        (function() {
            // Zenoti dropdown: Centers → All
            var cb = document.getElementById('elm_centers-zenoti-dropdown-options-all');
            if (cb && !cb.checked) cb.click();

            // All multi-selects → selectAll (Category, Sub Category, Product Type,
            // Vendor, Brand, Transaction Type, Business Unit)
            $('select[multiple]').each(function() {
                $(this).multiselect('selectAll', false);
            });
            // Stock Costing Method (single select) → Perpetual Average Cost
            $('select:not([multiple])').each(function() {
                var opts = Array.from(this.options);
                if (opts.some(function(o) { return o.text.indexOf('Perpetual') !== -1; })) {
                    $(this).multiselect('select', '1');
                }
            });
        })();
    """)
    time.sleep(2)
    print("  Stock Ledger filters applied.")


def download_gsheet_reports():
    script_dir = os.path.dirname(__file__) or "."
    succeeded = []
    failed = []

    for tab_name, info in GSHEET_TABS.items():
        try:
            url = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/export?format=csv&gid={info['gid']}"
            print(f"\nFetching Google Sheet: {tab_name} (gid={info['gid']})")
            df = pd.read_csv(url)
            print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")

            filename = os.path.join(script_dir, f"{tab_name}_{END_DATE}.csv")
            df.to_csv(filename, index=False)

            validate_csv(filename)
            upload_to_drive(filename, info["folder_id"])
            os.remove(filename)
            succeeded.append(tab_name)
        except Exception as e:
            print(f"FAILED Google Sheet: {tab_name} — {e}")
            failed.append((tab_name, str(e)))

    return succeeded, failed


def validate_report_folders(succeeded_reports):
    service = get_drive_service()
    if not service:
        return []

    missing = []
    for report in succeeded_reports:
        folder_id = REPORT_FOLDERS.get(report)
        if not folder_id:
            continue

        if report == "Business KPI":
            expected = [f"business_kpi_{BKP_START_DATE}_to_{END_DATE}.csv"]
        elif report == "Current Stock":
            if CURRENT_STOCK_AS_EXCEL:
                # Extension follows whatever the server served, so accept both.
                expected = [
                    f"current_stock_{END_DATE}.xlsx",
                    f"current_stock_{END_DATE}.xls",
                ]
            else:
                expected = [f"current_stock_{END_DATE}.csv"]
        else:
            safe_name = report.replace(" ", "_").lower()
            expected = [f"{safe_name}_{START_DATE}_to_{END_DATE}.csv"]

        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name)",
            pageSize=1000,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = results.get("files", [])
        names = [f["name"] for f in files]

        found = next((e for e in expected if e in names), None)
        if found:
            print(f"  OK {report}: found '{found}'")
        else:
            print(f"  MISSING {report}: expected '{' or '.join(expected)}' — found: {names}")
            missing.append(report)

    if missing:
        print(f"Missing reports in Drive: {missing}")
    else:
        print("All report folders validated successfully.")
    return missing


REPORT_FILTERS = {
    "Appointments": apply_appointments_filters,
    "Attendance": apply_attendance_filters,
    "Cost of Goods": apply_cost_of_goods_filters,
    "Sales-Accrual": apply_sales_accrual_filters,
    "Employee Sales": apply_employee_sales_filters,
    "Sales-Cash": apply_sales_cash_filters,
    "Business KPI": apply_business_kpi_filters,
    "Memberships": apply_memberships_filters,
    "Current Stock": apply_current_stock_filters,
    "Stock Ledger": apply_stock_ledger_filters,
}


def download_report(context, page, report_name, start_date, end_date):
    page.goto("https://evolvemedspa.zenoti.com/Admin/Reports/ReportsDashboard.aspx")
    page.wait_for_load_state("networkidle", timeout=120000)
    time.sleep(2)
    print(f"Opening report: {report_name}")

    if report_name == "Business KPI":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(2)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'business_kpi')")
    elif report_name == "Memberships":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(2)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'memberships')")
    elif report_name == "Stock Ledger":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(2)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'stock_ledger')")
    elif report_name == "Current Stock":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(2)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'current_stock')")
    elif report_name == "Employee Sales":
        # Unlike the other bookmarked reports there is no
        # ReportsGrid_Row_Click(event,'<id>') handle for this one, so find the
        # row by its visible name. "View All" renders the grid inside the
        # #dialog-reports modal and the same span may also exist in the grid
        # *behind* it, so scope the lookup to the modal — and dispatch the click
        # from JS, because a Playwright click is eaten by the modal overlay
        # intercepting pointer events.
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(3)

        find_row_js = """
            (function() {
                var modal = document.querySelector('#dialog-reports.show') ||
                            document.querySelector('#dialog-reports');
                var scope = modal || document;
                return Array.from(scope.querySelectorAll('span.report-name'))
                    .find(function(s) { return s.textContent.trim() === 'Employee Sales'; }) || null;
            })()
        """.strip()

        # Fail fast on a missing row instead of burning the 120s expect_page timeout.
        if not page.evaluate(f"!!{find_row_js}"):
            raise Exception("'Employee Sales' not found in the View All report grid")

        click_row_js = f"""
            (function() {{
                var span = {find_row_js};
                var row = span.closest('tr') || span.closest('td') || span;
                var handler = row.getAttribute('onclick') ||
                              (row.parentElement && row.parentElement.getAttribute('onclick')) || '';
                row.click();
                return handler || '(no onclick attribute; dispatched click)';
            }})();
        """

        with context.expect_page(timeout=120000) as new_page_info:
            print(f"  Row handler: {page.evaluate(click_row_js)}")
    else:
        with context.expect_page(timeout=120000) as new_page_info:
            page.locator('#gridReports span.report-name').get_by_text(report_name, exact=True).click(timeout=60000)

    time.sleep(2)
    report_page = new_page_info.value
    report_page.wait_for_load_state("load", timeout=120000)
    report_page.wait_for_load_state("networkidle", timeout=120000)
    time.sleep(2)
    print(f"{report_name} report page loaded.")

    if report_name == "Sales-Accrual":
        start_dt = f"{start_date} 00:00"
        end_dt = f"{end_date} 23:59"
        dt_format = "YYYY-MM-DD HH:mm"
    elif report_name == "Current Stock":
        # Single "Stock as on" datetime = T-1 at 07:00 AM
        start_dt = f"{start_date} 07:00"
        end_dt = f"{start_date} 07:00"
        dt_format = "YYYY-MM-DD HH:mm"
    else:
        start_dt = start_date
        end_dt = end_date
        dt_format = "YYYY-MM-DD"

    report_page.evaluate(f"""
        (function() {{
            var picker = $('#elm_dates').data('daterangepicker');
            if (picker) {{
                var startDate = moment('{start_dt}', '{dt_format}');
                var endDate = moment('{end_dt}', '{dt_format}');
                picker.setStartDate(startDate);
                picker.setEndDate(endDate);
                picker.element.trigger('apply.daterangepicker', picker);
            }}
        }})();
    """)
    time.sleep(2)
    print("Date range set.")

    filter_fn = REPORT_FILTERS.get(report_name)
    if filter_fn:
        filter_fn(report_page)
    time.sleep(2)

    print("Refreshing report...")
    if report_name == "Current Stock":
        report_page.evaluate("""
            (function() {
                if (typeof btnCurrentStock_onClick === 'function') {
                    btnCurrentStock_onClick();
                } else {
                    var b = document.querySelector('#btnRefresh');
                    if (b) b.click();
                }
            })();
        """)
    elif report_name == "Employee Sales":
        # #btnRefresh renders disabled until the page decides the filter set is
        # complete; drop the attribute so the click registers.
        report_page.evaluate("""
            (function() {
                var b = document.querySelector('#btnRefresh');
                if (b) {
                    b.removeAttribute('disabled');
                    b.click();
                }
            })();
        """)
    else:
        report_page.evaluate("document.querySelector('#btnRefresh').click()")
    time.sleep(2)
    report_page.wait_for_load_state("networkidle", timeout=300000)
    time.sleep(2)

    # The dropdown offers CSV (#export_csv), Excel (#export_excel_v2) and a hidden
    # "Excel with subtotals"; #export_excel_v2 runs Export_Click(true, "excel_v2").
    # Current Stock's Excel export is gated behind CURRENT_STOCK_AS_EXCEL.
    is_excel = report_name == "Current Stock" and CURRENT_STOCK_AS_EXCEL
    export_sel = "#export_excel_v2" if is_excel else "#export_csv"

    print(f"Exporting report to {'Excel' if is_excel else 'CSV'}...")
    report_page.locator('#dropdownMenuLink').click()
    time.sleep(2)
    report_page.wait_for_selector(export_sel, state='attached', timeout=30000)

    download_timeout = 900000 if report_name == "Stock Ledger" else 600000 if report_name == "Employee Sales" else 300000
    with report_page.expect_download(timeout=download_timeout) as download_info:
        report_page.evaluate(f"document.querySelector('{export_sel}').click()")

    time.sleep(10)
    download = download_info.value
    script_dir = os.path.dirname(__file__) or "."
    if report_name == "Business KPI":
        filename = os.path.join(script_dir, f"business_kpi_{start_date}_to_{end_date}.csv")
    elif report_name == "Current Stock":
        if is_excel:
            # Trust the server's extension when it gives one (.xlsx vs .xls);
            # validate_report_folders() accepts either.
            suggested = download.suggested_filename or ""
            ext = os.path.splitext(suggested)[1].lower()
            if ext not in (".xlsx", ".xls"):
                ext = ".xlsx"
        else:
            ext = ".csv"
        filename = os.path.join(script_dir, f"current_stock_{end_date}{ext}")
    else:
        safe_name = report_name.replace(" ", "_").lower()
        filename = os.path.join(script_dir, f"{safe_name}_{start_date}_to_{end_date}.csv")
    download.save_as(filename)
    time.sleep(2)

    print(f"Validating downloaded file: {filename}")
    if is_excel:
        validate_excel(filename)
    else:
        validate_csv(filename)
    print(f"Downloaded: {filename}")

    report_page.close()
    time.sleep(2)
    page.bring_to_front()
    time.sleep(2)
    return filename


print("Script starting...")
sys.stdout.flush()

LOG_FILENAME = os.path.join(os.path.dirname(__file__) or ".", f"download_report_logs_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.txt")
log_file = open(LOG_FILENAME, "w", encoding="utf-8")


class Tee:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for s in self.streams:
            s.write(data)
            s.flush()

    def flush(self):
        for s in self.streams:
            s.flush()


sys.stdout = Tee(sys.__stdout__, log_file)

cleanup_old_csvs()

# Phase 1: scan every report folder upfront and move any existing file
# (1 or more) to Done before downloading anything. Runs before the browser
# launches so leftovers from a previous run are cleared even if login fails,
# and so a bad GOOGLE_TOKEN_JSON surfaces before a full login.
print("Moving existing report files to Done...")
move_existing_reports_to_done()

with sync_playwright() as p:
    print("Playwright started.")
    browser, context = create_browser_and_context(p)
    print("Browser launched.")
    page = context.new_page()

    try:
        if needs_login(page):
            print("No valid session. Logging in...")
            do_login(page)
            save_cookies(context)
        else:
            print("Session valid from saved cookies. Skipping login.")

        wait_for_dashboard(page)
        save_cookies(context)

        reports = ["Stock Ledger", "Appointments", "Sales-Cash", "Cost of Goods", "Attendance", "Business KPI", "Memberships", "Current Stock", "Employee Sales"]
        # reports = ["Employee Sales"]
        failed_reports = []
        succeeded_reports = []

        for report in reports:
            try:
                if report == "Business KPI":
                    report_start = BKP_START_DATE
                    report_end = END_DATE
                    print(f"Business KPI date filter: {report_start} to {report_end}")
                elif report == "Current Stock":
                    report_start = END_DATE
                    report_end = END_DATE
                else:
                    report_start = START_DATE
                    report_end = END_DATE
                filename = download_report(context, page, report, report_start, report_end)
                folder_id = REPORT_FOLDERS.get(report, DRIVE_FOLDER_ID)
                upload_to_drive(filename, folder_id)
                os.remove(filename)
                save_cookies(context)
                succeeded_reports.append(report)
                time.sleep(5)
            except Exception as e:
                print(f"FAILED: {report} — {e}")
                failed_reports.append((report, str(e)))
                for p in context.pages:
                    if p != page:
                        try:
                            p.close()
                        except Exception:
                            pass
                page.bring_to_front()
                time.sleep(2)

        if failed_reports:
            print(f"\n--- Retrying {len(failed_reports)} failed report(s) ---")
            relogin_ok = True
            try:
                if needs_login(page):
                    print("Re-logging in before retry...")
                    do_login(page)
                    save_cookies(context)
                    wait_for_dashboard(page)
            except Exception as e:
                print(f"Re-login failed, skipping retries: {e}")
                relogin_ok = False

            retry_still_failed = []
            if not relogin_ok:
                retry_still_failed = list(failed_reports)
            for report, prev_error in (failed_reports if relogin_ok else []):
                try:
                    print(f"Retrying: {report}")
                    if report == "Business KPI":
                        report_start = BKP_START_DATE
                        report_end = END_DATE
                    elif report == "Current Stock":
                        report_start = END_DATE
                        report_end = END_DATE
                    else:
                        report_start = START_DATE
                        report_end = END_DATE
                    filename = download_report(context, page, report, report_start, report_end)
                    folder_id = REPORT_FOLDERS.get(report, DRIVE_FOLDER_ID)
                    upload_to_drive(filename, folder_id)
                    os.remove(filename)
                    save_cookies(context)
                    succeeded_reports.append(report)
                    time.sleep(5)
                except Exception as e:
                    print(f"RETRY FAILED: {report} — {e}")
                    retry_still_failed.append((report, str(e)))
                    for p in context.pages:
                        if p != page:
                            try:
                                p.close()
                            except Exception:
                                pass
                    page.bring_to_front()
                    time.sleep(2)

            failed_reports = retry_still_failed

        print(f"\n--- Report Summary ---")
        print(f"Succeeded: {succeeded_reports}")
        if failed_reports:
            print(f"Failed: {[r for r, _ in failed_reports]}")

        print(f"\n--- Google Sheet Extraction ---")
        gsheet_succeeded, gsheet_failed = download_gsheet_reports()
        print(f"Google Sheets succeeded: {gsheet_succeeded}")
        if gsheet_failed:
            print(f"Google Sheets failed: {[r for r, _ in gsheet_failed]}")

        print("Checking report folders for duplicate filenames...")
        dedupe_report_folders()

        print("Validating report folders contain recent downloads...")
        missing_reports = validate_report_folders(succeeded_reports)

        if missing_reports:
            print(f"\n--- Re-checking {len(missing_reports)} missing report(s) before re-download ---")
            time.sleep(10)
            missing_reports = validate_report_folders(missing_reports)

        # Re-download anything Drive says is absent, then dedupe + re-validate and
        # try again. Bounded at REDOWNLOAD_MAX_ATTEMPTS: a report that fails twice
        # is failing for a reason a third pass won't fix, and each pass costs a
        # full download.
        redownload_failed = []
        attempted_redownload = bool(missing_reports)
        for attempt in range(1, REDOWNLOAD_MAX_ATTEMPTS + 1):
            if not missing_reports:
                break

            print(
                f"\n--- Re-downloading {len(missing_reports)} missing report(s) "
                f"(attempt {attempt}/{REDOWNLOAD_MAX_ATTEMPTS}) ---"
            )
            try:
                if needs_login(page):
                    print("Re-logging in before re-download...")
                    do_login(page)
                    save_cookies(context)
                    wait_for_dashboard(page)
            except Exception as e:
                print(f"Re-login failed, abandoning re-download: {e}")
                break

            redownload_failed = []
            for report in list(missing_reports):
                try:
                    print(f"Re-downloading: {report}")
                    if report == "Business KPI":
                        report_start = BKP_START_DATE
                        report_end = END_DATE
                    elif report == "Current Stock":
                        report_start = END_DATE
                        report_end = END_DATE
                    else:
                        report_start = START_DATE
                        report_end = END_DATE
                    filename = download_report(context, page, report, report_start, report_end)
                    folder_id = REPORT_FOLDERS.get(report, DRIVE_FOLDER_ID)
                    upload_to_drive(filename, folder_id)
                    os.remove(filename)
                    save_cookies(context)
                    time.sleep(2)
                except Exception as e:
                    print(f"RE-DOWNLOAD FAILED: {report} — {e}")
                    redownload_failed.append(report)
                    for p in context.pages:
                        if p != page:
                            try:
                                p.close()
                            except Exception:
                                pass
                    page.bring_to_front()
                    time.sleep(2)

            if redownload_failed:
                print(f"Re-download failures: {redownload_failed}")

            # Validate against the reports we just re-downloaded, not the full
            # succeeded list, so the next pass retries only what is still absent.
            print(f"Validating after re-download attempt {attempt}...")
            dedupe_report_folders()
            missing_reports = validate_report_folders(missing_reports)

            if missing_reports and attempt < REDOWNLOAD_MAX_ATTEMPTS:
                print(f"Still missing {missing_reports}; retrying in {REDOWNLOAD_RETRY_BACKOFF}s...")
                time.sleep(REDOWNLOAD_RETRY_BACKOFF)

        # Final sweep across every report that downloaded successfully, so one that
        # vanished from Drive after its own pass is still caught. Skipped when the
        # first validation was already clean — line above would just repeat it.
        if attempted_redownload:
            print("Final validation...")
            dedupe_report_folders()
            still_missing = validate_report_folders(succeeded_reports)
        else:
            still_missing = missing_reports

        if still_missing:
            raise Exception(
                f"Reports still missing after {REDOWNLOAD_MAX_ATTEMPTS} re-download attempt(s): {still_missing}"
            )

        print("Logging out...")
        page.goto("https://evolvemedspa.zenoti.com/Admin/Reports/ReportsDashboard.aspx")
        page.wait_for_load_state("networkidle", timeout=60000)
        time.sleep(1)
        page.locator('#usernameBtn').click
        time.sleep(1)
        page.locator('.userLogoutCls').click
        time.sleep(2)
        print("Logged out.")

        if failed_reports:
            raise Exception(f"Reports failed after retry: {[r for r, _ in failed_reports]}")

    except Exception as e:
        print(f"Error: {e}")
        raise
    finally:
        context.close()
        browser.close()

    sys.stdout = sys.__stdout__
    log_file.close()
    upload_to_drive(LOG_FILENAME, DRIVE_FOLDER_ID)
    os.remove(LOG_FILENAME)