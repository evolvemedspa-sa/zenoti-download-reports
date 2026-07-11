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

if not USERNAME or not PASSWORD:
    raise ValueError("MINER_USER and MINER_PASSWORD must be set in the .env file.")

# yesterday = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
# START_DATE = yesterday
# END_DATE = yesterday

today = date.today()
date_to = today - timedelta(days=1)
first_day_this_month = today.replace(day=1)
date_from = first_day_this_month - relativedelta(months=1)

START_DATE = date_from.strftime("%Y-%m-%d")
END_DATE = date_to.strftime("%Y-%m-%d")

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
    "Business KPI": "1GjkgXcKrGFqa8l9iM-rW8u2MeRVCaB_M",
    "Memberships": "172HJzXYy_9_qtlgTSlZUgUJmZmT-7qwH",
    "Inventory Aging": "174ZiUaKjIjEKJNKe75mZXKAwya0F4GNK",
    "Stock Ledger": "1JwZGmMBu-3ZHb67edqOZ8vsj5u9eMRd9",
}
SCOPES = ["https://www.googleapis.com/auth/drive.file"]


def get_drive_service():
    from google.auth.transport.requests import Request

    token_json = os.getenv("GOOGLE_TOKEN_JSON")
    if not token_json:
        print("GOOGLE_TOKEN_JSON not set. Skipping upload.")
        return None

    creds_info = json.loads(token_json)
    creds = Credentials.from_authorized_user_info(creds_info, SCOPES)

    # Refresh if expired
    if creds.expired and creds.refresh_token:
        print("Refreshing Google OAuth token...")
        creds.refresh(Request())
        # Note: can't save back to env var on Railway, but refresh_token stays valid
        print("Token refreshed.")

    return build("drive", "v3", credentials=creds)

def upload_to_drive(filepath, folder_id=DRIVE_FOLDER_ID):
    service = get_drive_service()
    if not service:
        return None

    filename = os.path.basename(filepath)

    # Phase 1 (moving existing files to Done) runs upfront in
    # move_existing_reports_to_done() before any download, so no pre-move here.
    print(f"Uploading new file: {filename}")
    file_metadata = {"name": filename, "parents": [folder_id]}
    media = MediaFileUpload(filepath, mimetype="text/csv", resumable=True)
    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,webViewLink",
    ).execute()
    time.sleep(5)

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


def move_existing_reports_to_done():
    service = get_drive_service()
    if not service:
        return

    for folder_name, folder_id in REPORT_FOLDERS.items():
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name)",
        ).execute()
        existing = results.get("files", [])
        for old_file in existing:
            print(f"Moving {folder_name}/{old_file['name']} to Done folder")
            service.files().update(
                fileId=old_file["id"],
                addParents=DONE_FOLDER_ID,
                removeParents=folder_id,
                fields="id",
            ).execute()
            time.sleep(2)


def dedupe_report_folders():
    """Guardrail: if a report folder holds more than one file with the same
    name (e.g. a container restart re-ran the script and re-uploaded), keep one
    and move the extras to Done. Move — not delete — because the drive.file
    scope cannot delete files it did not create."""
    service = get_drive_service()
    if not service:
        return

    for folder_name, folder_id in REPORT_FOLDERS.items():
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name)",
        ).execute()
        files = results.get("files", [])

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


def cleanup_old_csvs():
    script_dir = os.path.dirname(__file__) or "."
    for f in glob.glob(os.path.join(script_dir, "*.csv")):
        if date_to.strftime("%Y-%m-%d") not in os.path.basename(f):
            os.remove(f)
            print(f"Cleaned up old CSV: {f}")


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
    print(f"Checking if login is needed. Current URL: {page.url}")
    page.goto(ADMIN_URL, wait_until="domcontentloaded")
    try:
        page.wait_for_url("**/Admin/**", timeout=10000)
        return False
    except:
        return True


def do_login(page):
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
    page.locator(username_sel).first.press_sequentially(USERNAME, delay=50)
    time.sleep(random.uniform(0.5, 1.5))
    print("Username entered.")

    page.locator('#Password').click()
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


def apply_inventory_aging_filters(report_page):
    print("  Applying Inventory Aging filters...")
    report_page.evaluate("""
        (function() {
            // All multi-selects → selectAll (Centers, Category, Sub Category, Vendor, Brand, Business Unit)
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
                // Stock Costing Method → Perpetual Average Cost (value "1")
                else if (texts.some(function(t) { return t.indexOf('Perpetual') !== -1; })) {
                    $(this).multiselect('select', '1');
                }
            });
        })();
    """)
    time.sleep(2)
    print("  Inventory Aging filters applied.")


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


REPORT_FILTERS = {
    "Appointments": apply_appointments_filters,
    "Attendance": apply_attendance_filters,
    "Cost of Goods": apply_cost_of_goods_filters,
    "Sales-Accrual": apply_sales_accrual_filters,
    "Sales-Cash": apply_sales_cash_filters,
    "Business KPI": apply_business_kpi_filters,
    "Memberships": apply_memberships_filters,
    "Inventory Aging": apply_inventory_aging_filters,
    "Stock Ledger": apply_stock_ledger_filters,
}


def download_report(context, page, report_name, start_date, end_date):
    page.goto("https://evolvemedspa.zenoti.com/Admin/Reports/ReportsDashboard.aspx")
    page.wait_for_load_state("networkidle", timeout=120000)
    time.sleep(5)
    print(f"Opening report: {report_name}")

    if report_name == "Business KPI":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(5)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'business_kpi')")
    elif report_name == "Memberships":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(5)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'memberships')")
    elif report_name == "Inventory Aging":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(5)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'inventory_aging')")
    elif report_name == "Stock Ledger":
        page.evaluate('loadBookmarksViewAllGrid("Bookmarked")')
        time.sleep(5)
        with context.expect_page(timeout=120000) as new_page_info:
            page.evaluate("ReportsGrid_Row_Click(event,'stock_ledger')")
    else:
        with context.expect_page(timeout=120000) as new_page_info:
            page.locator('#gridReports span.report-name').get_by_text(report_name, exact=True).click(timeout=60000)

    time.sleep(5)
    report_page = new_page_info.value
    report_page.wait_for_load_state("load", timeout=120000)
    report_page.wait_for_load_state("networkidle", timeout=120000)
    time.sleep(5)
    print(f"{report_name} report page loaded.")

    if report_name == "Sales-Accrual":
        start_dt = f"{start_date} 00:00"
        end_dt = f"{end_date} 23:59"
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
    time.sleep(3)
    print("Date range set.")

    filter_fn = REPORT_FILTERS.get(report_name)
    if filter_fn:
        filter_fn(report_page)
    time.sleep(2)

    print("Refreshing report...")
    report_page.evaluate("document.querySelector('#btnRefresh').click()")
    time.sleep(5)
    report_page.wait_for_load_state("networkidle", timeout=300000)
    time.sleep(5)

    print("Exporting report to CSV...")
    report_page.locator('#dropdownMenuLink').click()
    time.sleep(2)
    report_page.wait_for_selector('#export_csv', state='attached', timeout=30000)

    download_timeout = 900000 if report_name == "Stock Ledger" else 300000
    with report_page.expect_download(timeout=download_timeout) as download_info:
        report_page.evaluate("document.querySelector('#export_csv').click()")

    time.sleep(10)
    download = download_info.value
    script_dir = os.path.dirname(__file__) or "."
    if report_name == "Business KPI":
        filename = os.path.join(script_dir, f"Business_Kpi_{end_date}.csv")
    else:
        safe_name = report_name.replace(" ", "_").lower()
        filename = os.path.join(script_dir, f"{safe_name}_{start_date}_to_{end_date}.csv")
    download.save_as(filename)
    time.sleep(5)

    print(f"Validating downloaded file: {filename}")
    validate_csv(filename)
    print(f"Downloaded: {filename}")

    report_page.close()
    time.sleep(2)
    page.bring_to_front()
    time.sleep(2)
    return filename


print("Script starting...")
sys.stdout.flush()

LOG_FILENAME = os.path.join(os.path.dirname(__file__) or ".", f"logs_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.txt")
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

        # Phase 1: scan every report folder upfront and move any existing file
        # (1 or more) to Done before downloading anything.
        print("Moving existing report files to Done...")
        move_existing_reports_to_done()

        reports = ["Stock Ledger", "Appointments", "Sales-Cash", "Cost of Goods", "Attendance", "Business KPI", "Memberships"]
        # reports = ["Stock Ledger", "Appointments"]
        failed_reports = []
        succeeded_reports = []

        for report in reports:
            try:
                if report == "Business KPI":
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
                time.sleep(10)
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
                time.sleep(5)

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
                    time.sleep(10)
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
                    time.sleep(5)

            failed_reports = retry_still_failed

        print(f"\n--- Report Summary ---")
        print(f"Succeeded: {succeeded_reports}")
        if failed_reports:
            print(f"Failed: {[r for r, _ in failed_reports]}")

        print("Checking report folders for duplicate filenames...")
        dedupe_report_folders()

        print("Logging out...")
        page.goto("https://evolvemedspa.zenoti.com/Admin/Reports/ReportsDashboard.aspx")
        page.wait_for_load_state("networkidle", timeout=60000)
        time.sleep(1)
        page.locator('#usernameBtn').click
        time.sleep(1)
        page.locator('.userLogoutCls').click
        time.sleep(5)
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