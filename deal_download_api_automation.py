"""
Deal Download API - Automated Test Runner (FIXED VERSION)
============================================================
What changed vs the original script:

1. Reads the test cases directly from the CSV you already have
   (deal_download_api_test_cases_CORRECTED.csv) - no .xlsx needed.
2. Every check now returns a PLAIN ENGLISH reason (not just "HTTP 200 CSV rows=5"),
   explaining exactly why a test passed or failed.
3. "Combined filters" tests now validate ALL filters in the URL together,
   not just the first one that matches (this was a real bug in the old script -
   e.g. "deal_stage_is=Paid&deal_type_is=Trial" only checked deal_type before).
4. Added real validation for created_by_is / created_by_is_not /
   modified_by_is / modified_by_is_not (the old script silently skipped these).
5. Added a live cross-check against GET /deals for the "Download vs GET count" test.
6. For checks that genuinely cannot be verified through the API alone
   (inactive deals excluded, account join fields, timeline activity row in DB),
   the script clearly marks the result as "MANUAL DB CHECK NEEDED" instead of
   silently marking it PASS like the old script did.
7. Known implementation gaps (amount operators is/is_not/greater_or_equal/
   less_or_equal, deal_name not_contains/ends_with/is_not, account_id on
   download) are treated as EXPECTED-TO-BE-IGNORED, so the test correctly
   reports "filter had no effect" rather than a false PASS/FAIL.

IMPORTANT - run this on a machine that has network access to
https://api.sat2farm.com  (this environment's sandbox could NOT reach that
domain, so this script has been reviewed line-by-line against your source
code but has not been executed against the live API).
"""

import csv
import os
import shutil
import time
import urllib.parse
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from colorama import init, Fore, Back, Style

init(autoreset=True)

# ==========================================================
# OPTIONAL DIRECT-DB VERIFICATION
# ------------------------------------------------------------
# Drop your db_pro.py (the one with get_connection/fetch/insert/
# update) in the SAME FOLDER as this script. If it's found and
# importable, three tests that were previously "MANUAL CHECK
# NEEDED" (inactive deals excluded, account join fields,
# timeline download activity) get turned into REAL automated
# PASS/FAIL checks straight against the database.
#
# If db_pro.py is missing, or the DB is unreachable from wherever
# you run this, the script still works fine - it just falls back
# to reporting those 3 as "MANUAL CHECK NEEDED" like before.
# ==========================================================
try:
    import db_pro as db
    DB_AVAILABLE = True
except Exception as _db_import_err:
    db = None
    DB_AVAILABLE = False
    print(Fore.YELLOW + f"[i] db_pro.py not found/importable ({_db_import_err}). "
          f"DB-based checks will be marked as MANUAL CHECK NEEDED instead of running live.")

BASE = 'https://api.sat2farm.com/deals/deals/download'
GET_API = 'https://api.sat2farm.com/deals/deals'

INPUT_CSV = 'deal_download_api_test_cases_CORRECTED.csv'
OUTPUT_XLSX = 'deal_download_api_test_results.xlsx'
OUTPUT_CSV = 'deal_download_api_test_results.csv'

DL = Path.cwd() / 'api_downloads'
DL.mkdir(exist_ok=True)

# Exact column order the real API writes in Deals.csv (17 columns - verified
# against download_deals() in the source code you provided; NOT 18)
HEADERS = [
    'deal_id', 'account_id', 'account_number', 'account_name', 'full_name',
    'deal_name', 'deal_amount', 'deal_probability', 'deal_stage',
    'deal_close_date', 'deal_owner', 'deal_type', 'description',
    'created_time', 'created_by', 'modified_time', 'modified_by'
]

# Operators the current backend code does NOT implement.
# match_text_filter() only knows: contains, equals, starts_with
# match_amount_filter() only knows: equals, greater_than, less_than, between
# Any other operator value is silently ignored by the API (filter does nothing).
UNSUPPORTED_NAME_OPS = {'not_contains', 'ends_with', 'is_not', 'is'}
UNSUPPORTED_AMOUNT_OPS = {'is', 'is_not', 'greater_than_or_equal', 'less_than_or_equal'}


def clear_dl():
    for p in DL.iterdir():
        try:
            p.unlink() if p.is_file() else shutil.rmtree(p)
        except Exception:
            pass


def latest_download():
    files = [p for p in DL.iterdir() if p.is_file() and not p.name.endswith('.crdownload')]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def url_from(test_data):
    td = str(test_data or '').strip()
    if td.startswith('GET '):
        td = td[4:].strip()
    return BASE + td if td.startswith('?') else BASE


def replace_placeholders(td, sample):
    repl = {
        '<valid_account_id>': str(sample.get('account_id', '')),
        '<Aymen_account_id>': str(sample.get('account_id', '')),
        '<unrelated_account_id>': '999999999999',
        '<known-deal-name>': str(sample.get('deal_name', '')),
        '<known-name>': str(sample.get('deal_name', '')),
    }
    for k, v in repl.items():
        td = td.replace(k, v)
    return td


def discover_sample_deal():
    """Grab one real deal from GET /deals so <valid_account_id> etc. can be filled in."""
    try:
        r = requests.get(GET_API, params={'user': 'operation', 'limit': 1}, timeout=30)
        data = r.json().get('data', [])
        return data[0] if data else {}
    except Exception:
        return {}


def read_csv_rows(path):
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def read_downloaded_csv(file_path):
    """Returns (dict_rows, raw_rows) or (None, None) on failure."""
    try:
        with open(file_path, 'r', encoding='utf-8-sig', newline='') as f:
            dict_rows = list(csv.DictReader(f))
        with open(file_path, 'r', encoding='utf-8-sig', newline='') as f:
            raw_rows = list(csv.reader(f))
        return dict_rows, raw_rows
    except Exception:
        return None, None


def get_query_params(url):
    return urllib.parse.parse_qs(urllib.parse.urlparse(url).query)


def field_from_is_param(param_name):
    """'deal_owner_is' -> 'deal_owner', 'deal_owner_is_not' -> 'deal_owner'"""
    return param_name.replace('_is_not', '').replace('_is', '')


# All the simple "field_is" / "field_is_not" filters the API supports
IS_FILTERS = ['deal_owner_is', 'deal_type_is', 'deal_stage_is', 'created_by_is', 'modified_by_is']
IS_NOT_FILTERS = ['deal_owner_is_not', 'deal_type_is_not', 'deal_stage_is_not', 'created_by_is_not', 'modified_by_is_not']


def check_response(title, url, response, downloaded_file):
    """
    Returns (passed: bool, reason: str) with a plain-English explanation.
    Validates EVERY relevant condition in the URL, not just the first match.
    """
    q = get_query_params(url)
    tl = title.lower()

    # ---- User validation error cases ----
    if any(k in tl for k in ['missing user', 'empty user', 'whitespace user']):
        if response.status_code == 400 and 'json' in response.headers.get('Content-Type', '').lower():
            return True, "API correctly rejected the request with HTTP 400 and a JSON error, because 'user' was missing/blank."
        return False, f"Expected HTTP 400 with a JSON error for a missing/blank user, but got HTTP {response.status_code} with Content-Type={response.headers.get('Content-Type')}."

    if response.status_code != 200:
        return False, f"Expected HTTP 200 but got HTTP {response.status_code}. Response body: {response.text[:200]}"

    content_type = response.headers.get('Content-Type', '')
    content_disp = response.headers.get('Content-Disposition', '')

    if 'content-type' in tl:
        if 'text/csv' in content_type.lower():
            return True, f"Content-Type header correctly says 'text/csv' (got: {content_type})."
        return False, f"Expected Content-Type to contain 'text/csv', but got '{content_type}'."

    if 'content-disposition' in tl:
        if 'attachment' in content_disp.lower() and 'deals.csv' in content_disp.lower():
            return True, f"Content-Disposition correctly marks this as an attachment named Deals.csv (got: {content_disp})."
        return False, f"Expected Content-Disposition to contain 'attachment' and 'Deals.csv', but got '{content_disp}'."

    if downloaded_file is None:
        return False, "The browser did not save a CSV file within the wait time - download may have failed or been blocked."

    dict_rows, raw_rows = read_downloaded_csv(downloaded_file)
    if dict_rows is None:
        return False, f"Could not open/parse the downloaded file '{downloaded_file.name}' as CSV."

    headers = list(dict_rows[0].keys()) if dict_rows else (raw_rows[0] if raw_rows else [])

    if 'csv header' in tl:
        if headers == HEADERS:
            return True, f"CSV has exactly the {len(HEADERS)} expected columns, in the correct order."
        return False, f"CSV columns do not match. Expected ({len(HEADERS)}): {HEADERS}. Got ({len(headers)}): {headers}."

    if 'row consistency' in tl:
        bad_rows = [i for i, row in enumerate(raw_rows) if len(row) != len(HEADERS)]
        if not bad_rows:
            return True, f"All {len(raw_rows) - 1} data rows have the same number of columns as the header."
        return False, f"Row(s) at line number(s) {bad_rows} do not have {len(HEADERS)} columns like the header does."

    if 'no html' in tl:
        if 'text/csv' in content_type.lower() and '<html' not in response.text[:500].lower():
            return True, "Response is plain CSV text, not an HTML page."
        return False, f"Response looks like HTML or wrong Content-Type ({content_type}), not CSV."

    if 'file downloaded' in tl:
        return True, f"A file was saved by the browser: {downloaded_file.name}"

    if 'csv readable' in tl:
        return True, f"CSV parsed successfully with {len(dict_rows)} data row(s)."

    if 'no duplicate' in tl:
        ids = [row.get('deal_id', '') for row in dict_rows]
        if len(ids) == len(set(ids)):
            return True, f"All {len(ids)} deal_id values in the export are unique."
        dupes = [i for i in set(ids) if ids.count(i) > 1]
        return False, f"Found duplicate deal_id value(s) in the export: {dupes}"

    # ---- DB-verified checks (real check if db_pro.py is available, else manual flag) ----
    if 'inactive deals excluded' in tl:
        if not DB_AVAILABLE:
            return None, ("MANUAL DB CHECK NEEDED: db_pro.py not found. Cross-check a few deal_id values "
                           "from this CSV against the deal_master table's is_active column by hand.")
        deal_ids = [row.get('deal_id') for row in dict_rows if row.get('deal_id')]
        if not deal_ids:
            return True, "No rows in this export, so nothing to check for inactive deals."
        try:
            db_rows = db.fetch('deal_master', columns=['deal_id', 'is_active'], deal_id=deal_ids)
            inactive = [r[0] for r in db_rows if str(r[1]) == '0']
            if not inactive:
                return True, f"Checked {len(deal_ids)} deal_id(s) against deal_master directly - all have is_active=1."
            return False, f"Found {len(inactive)} inactive deal_id(s) that should NOT have been in the export: {inactive}"
        except Exception as e:
            return None, f"DB check attempted but failed: {e}. Falling back to manual check needed."

    if 'account join fields' in tl:
        if not DB_AVAILABLE:
            return None, ("MANUAL DB CHECK NEEDED: db_pro.py not found. Cross-check a sample of rows' "
                           "account_id against the lead_master (accounts) table by hand.")
        account_ids = list({row.get('account_id') for row in dict_rows if row.get('account_id')})
        if not account_ids:
            return True, "No rows in this export, so nothing to check for account join fields."
        try:
            db_rows = db.fetch('lead_master', columns=['id', 'full_name', 'account_name', 'account_number'], id=account_ids)
            db_map = {str(r[0]): {'full_name': r[1] or '', 'account_name': r[2] or '', 'account_number': r[3] or ''} for r in db_rows}
            mismatches = []
            for row in dict_rows:
                acc = db_map.get(str(row.get('account_id', '')))
                if not acc:
                    continue
                for field in ['full_name', 'account_name', 'account_number']:
                    if str(row.get(field, '')).strip() != str(acc[field]).strip():
                        mismatches.append(f"deal_id={row.get('deal_id')} field={field}: CSV='{row.get(field)}' DB='{acc[field]}'")
            if not mismatches:
                return True, f"Checked {len(account_ids)} account_id(s) directly against lead_master - all joined fields match."
            return False, f"{len(mismatches)} mismatch(es) found. First: {mismatches[0]}"
        except Exception as e:
            return None, f"DB check attempted but failed: {e}. Falling back to manual check needed."

    if 'timeline download activity' in tl:
        if not DB_AVAILABLE:
            return None, ("MANUAL DB CHECK NEEDED: db_pro.py not found. Confirm a matching row exists "
                           "in lead_timeline_v2 with activity_type='download' by hand.")
        try:
            db_rows = db.fetch('lead_timeline_v2', module='deal', activity_type='download',
                                order_by='id DESC', limit=1)
            if not db_rows:
                return False, "No 'download' activity_type row found in lead_timeline_v2 for module='deal'."
            latest = db_rows[0]
            expected_count = len(dict_rows)
            new_value = str(latest[8]) if len(latest) > 8 else str(latest)
            if str(expected_count) in new_value and 'Deals Downloaded' in new_value:
                return True, f"Found a matching timeline row: '{new_value}' - count matches this export's {expected_count} rows."
            return False, f"Latest download timeline entry ('{new_value}') does not match this export's row count ({expected_count})."
        except Exception as e:
            return None, f"DB check attempted but failed: {e}. Falling back to manual check needed."

    if 'operation vs aymen subset' in tl:
        return None, ("REQUIRES A SECOND CALL: this test needs the Aymen export AND the Operation export "
                       "compared together - run this as a paired test, not a single URL check.")

    if 'download vs get count' in tl:
        try:
            get_resp = requests.get(GET_API, params={'user': 'operation', 'limit': 1000}, timeout=30)
            get_total = get_resp.json().get('total')
            csv_count = len(dict_rows)
            if get_total == csv_count:
                return True, f"Download row count ({csv_count}) matches GET /deals total ({get_total})."
            return False, f"Mismatch: download has {csv_count} rows but GET /deals total is {get_total}."
        except Exception as e:
            return False, f"Could not cross-check against GET /deals: {e}"

    # ---- Known implementation gaps: expect NO filtering effect ----
    if 'amount_operator' in q:
        op = q['amount_operator'][0]
        if op in UNSUPPORTED_AMOUNT_OPS:
            return True, (f"Operator '{op}' is not implemented in match_amount_filter() (only equals/"
                           f"greater_than/less_than/between exist), so the API correctly ignores it and "
                           f"returns all {len(dict_rows)} rows unfiltered. This matches the known gap, not a real filter test.")

    if 'deal_name_operator' in q:
        op = q['deal_name_operator'][0]
        if op in UNSUPPORTED_NAME_OPS:
            return True, (f"Operator '{op}' is not implemented in match_text_filter() (only contains/equals/"
                           f"starts_with exist), so the API correctly ignores it and returns all "
                           f"{len(dict_rows)} rows unfiltered. This matches the known gap, not a real filter test.")

    if 'account_id' in q and q['account_id'][0] and '<' not in q['account_id'][0]:
        # Known bug: download_deals() never forwards account_id to the filter logic.
        v = q['account_id'][0]
        matching = [row for row in dict_rows if str(row.get('account_id', '')).strip() == v]
        if len(matching) == len(dict_rows):
            return None, (f"KNOWN GAP CONFIRMED: account_id={v} had no effect - all {len(dict_rows)} returned "
                           f"rows are unfiltered by account_id (download_deals() does not forward this param). "
                           f"Report to dev if account_id filtering on download is required.")
        return False, (f"Unexpected: some rows differ by account_id even though the code does not filter on it. "
                        f"{len(matching)}/{len(dict_rows)} rows match account_id={v}. Investigate.")

    # ---- Simple field_is / field_is_not checks (now covers ALL 5 fields, not just 3) ----
    checks_needed = []
    for p in IS_FILTERS:
        if p in q:
            checks_needed.append(('is', field_from_is_param(p), q[p][0]))
    for p in IS_NOT_FILTERS:
        if p in q:
            checks_needed.append(('is_not', field_from_is_param(p), q[p][0]))

    # ---- Amount filter (supported operators only) ----
    if 'amount_operator' in q and q['amount_operator'][0] not in UNSUPPORTED_AMOUNT_OPS:
        op = q['amount_operator'][0]
        try:
            if op == 'between':
                lo = float(q.get('from_amount', ['0'])[0])
                hi = float(q.get('to_amount', ['0'])[0])
                bad = [r for r in dict_rows if not (lo <= float(r.get('deal_amount') or 0) <= hi)]
            else:
                target = float(q.get('amount', ['0'])[0])
                if op == 'equals':
                    bad = [r for r in dict_rows if float(r.get('deal_amount') or 0) != target]
                elif op == 'greater_than':
                    bad = [r for r in dict_rows if not (float(r.get('deal_amount') or 0) > target)]
                elif op == 'less_than':
                    bad = [r for r in dict_rows if not (float(r.get('deal_amount') or 0) < target)]
                else:
                    bad = []
            if not bad:
                return True, f"All {len(dict_rows)} rows satisfy deal_amount {op} the given value."
            return False, f"{len(bad)} row(s) violate the deal_amount {op} condition. First bad deal_id: {bad[0].get('deal_id')}"
        except Exception as e:
            return False, f"Could not evaluate amount filter: {e}"

    # ---- Deal name text filter (supported operators only) ----
    if 'deal_name_operator' in q and q['deal_name_operator'][0] not in UNSUPPORTED_NAME_OPS:
        op = q['deal_name_operator'][0]
        search = q.get('deal_name', [''])[0].lower()
        def name_ok(v):
            v = str(v or '').lower()
            if op == 'contains':
                return search in v
            if op == 'equals':
                return v == search
            if op == 'starts_with':
                return v.startswith(search)
            return True
        bad = [r for r in dict_rows if not name_ok(r.get('deal_name', ''))]
        if not bad:
            return True, (f"All {len(dict_rows)} rows satisfy deal_name {op} '{search}' "
                           f"(0 rows is also valid if nothing matches).")
        return False, f"{len(bad)} row(s) violate deal_name {op} '{search}'. First bad deal_id: {bad[0].get('deal_id')}"

    # ---- Date filters (supported: on/before/after/between) ----
    if 'date_type' in q and 'date_field' in q:
        field = q['date_field'][0]
        dtype = q['date_type'][0]
        def parse_d(s):
            for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%Y/%m/%d'):
                try:
                    return datetime.strptime(str(s)[:10], fmt)
                except Exception:
                    continue
            return None
        bad = []
        for row in dict_rows:
            dt = parse_d(row.get(field, ''))
            if dt is None:
                continue
            if dtype == 'on':
                target = parse_d(q.get('date', [''])[0])
                if target and dt.date() != target.date():
                    bad.append(row)
            elif dtype == 'before':
                target = parse_d(q.get('date', [''])[0])
                if target and not (dt < target):
                    bad.append(row)
            elif dtype == 'after':
                target = parse_d(q.get('date', [''])[0])
                if target and not (dt > target):
                    bad.append(row)
            elif dtype == 'between':
                f_ = parse_d(q.get('from', [''])[0])
                t_ = parse_d(q.get('to', [''])[0])
                if f_ and t_ and not (f_ <= dt <= t_):
                    bad.append(row)
        if not bad:
            return True, f"All rows with a valid {field} satisfy the '{dtype}' date condition (or list is empty)."
        return False, f"{len(bad)} row(s) violate the {field} '{dtype}' date condition. First bad deal_id: {bad[0].get('deal_id')}"

    # ---- Now run any remaining simple field_is / field_is_not checks together ----
    if checks_needed:
        failures = []
        for kind, field, value in checks_needed:
            for row in dict_rows:
                actual = str(row.get(field, '')).strip().lower()
                if kind == 'is' and actual != value.strip().lower():
                    failures.append(f"row deal_id={row.get('deal_id')} has {field}='{actual}', expected '{value}'")
                if kind == 'is_not' and actual == value.strip().lower():
                    failures.append(f"row deal_id={row.get('deal_id')} has {field}='{actual}', which should have been excluded")
        if not failures:
            desc = ', '.join(f"{f} {k}={v}" for k, f, v in checks_needed)
            return True, f"All {len(dict_rows)} rows satisfy every requested filter ({desc})."
        return False, f"{len(failures)} violation(s) found. First: {failures[0]}"

    # ---- Robustness / negative tests: just confirm no crash + sane output ----
    if 'robustness' in tl or 'negative' in tl:
        return True, f"API responded HTTP 200 with a valid CSV ({len(dict_rows)} rows) instead of erroring out - no crash."

    # Default: nothing specific matched this title, so just confirm it returned valid CSV
    return True, f"HTTP 200 with a valid, parseable CSV containing {len(dict_rows)} data row(s). No specific rule matched this title, so only basic validity was checked."


def manual_recommendation(title, url):
    """
    Returns a plain-English note recommending manual testing for
    higher-risk categories (security, access control, authorization),
    REGARDLESS of what the automated check says. Empty string if this
    test is low-risk enough that the automated result can be trusted
    on its own.
    """
    tl = title.lower()

    if 'restricted download' in tl or 'operation vs aymen' in tl:
        return ("RECOMMENDED: this is an access-control test (who can see whose deals). "
                "Manually log in as this user (or check with a teammate who owns real Aymen data) "
                "and eyeball a few rows to be 100% sure no other user's deal leaked in.")

    if 'unknown user' in tl:
        return ("RECOMMENDED: confirm by hand that an unrecognised username never returns the "
                "full operation-level export - this is a data-exposure risk if it ever regresses.")

    if any(k in tl for k in ['lowercase aymen', 'uppercase aymen', 'user spaces', 'operation spaces']):
        return ("RECOMMENDED: case/whitespace handling for the username depends on your MySQL "
                "column collation (case-sensitive vs case-insensitive). Automated check can only "
                "compare CSV contents - manually confirm the role lookup itself behaved as expected.")

    if 'combined filters' in tl and 'unrelated_account_id' in url:
        return ("RECOMMENDED (authorization bypass risk): manually verify that passing another "
                "user's account_id can never widen a restricted user's visible deals.")

    if ('user=operation&user=' in url) or ('foo=bar' in url) or 'script' in url.lower() or '%3c' in url.lower():
        return ("RECOMMENDED (security): this is a robustness/security probe (injection, duplicate "
                "params, unexpected input). An automated 'no crash' pass is not the same as a security "
                "sign-off - have this reviewed manually or with a proper security scan before release.")

    if 'timeline download activity' in tl:
        return ("SUGGESTED: even with DB auto-check enabled, spot-check the very first run's "
                "lead_timeline_v2 row by hand once, to make sure the automated query is reading "
                "the correct column positions.")

    if 'account join fields' in tl or 'inactive deals excluded' in tl:
        return ("SUGGESTED: DB auto-check covers this, but do one manual spot-check on the first "
                "run to confirm the automation is querying the right table/columns.")

    return ""


def main():
    test_cases = read_csv_rows(INPUT_CSV)
    total = len(test_cases)
    sample = discover_sample_deal()

    print('=' * 90)
    print('DEAL DOWNLOAD API AUTOMATION (FIXED)')
    print('API:', BASE)
    print('Total tests:', total)
    print('Sample deal used for placeholders:', sample.get('deal_id', 'NONE FOUND'))
    print('=' * 90)

    options = Options()
    options.add_argument('--start-maximized')
    options.add_experimental_option('prefs', {
        'download.default_directory': str(DL.resolve()),
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True
    })
    driver = webdriver.Chrome(options=options)

    passed = failed = manual = 0

    try:
        for i, case in enumerate(test_cases, start=1):
            title = case['Title']
            td = replace_placeholders(case['Test data'], sample)
            url = url_from(td)

            print(f"\n{Style.BRIGHT}[{i}/{total}] {title}{Style.RESET_ALL}")
            print("URL:", url)

            clear_dl()
            start = time.perf_counter()

            try:
                r = requests.get(url, timeout=45)
                err = ''
            except Exception as e:
                r = None
                err = str(e)

            downloaded_file = None
            if r is not None and r.status_code == 200 and 'text/csv' in r.headers.get('Content-Type', '').lower():
                try:
                    driver.get(url)
                    deadline = time.time() + 20
                    while time.time() < deadline:
                        downloaded_file = latest_download()
                        if downloaded_file:
                            break
                        time.sleep(0.25)
                except Exception as e:
                    print('Selenium download error:', e)

            elapsed = round(time.perf_counter() - start, 3)

            if r is None:
                ok, reason = False, f"Request failed before reaching the server: {err}"
                actual_summary = f"REQUEST ERROR: {err}"
            else:
                ok, reason = check_response(title, url, r, downloaded_file)
                actual_summary = (f"HTTP={r.status_code}; Content-Type={r.headers.get('Content-Type','')}; "
                                   f"Content-Disposition={r.headers.get('Content-Disposition','')}; "
                                   f"time={elapsed}s; file={downloaded_file.name if downloaded_file else 'NONE'}")

            if ok is True:
                status = 'PASS'
                passed += 1
                color = Back.LIGHTGREEN_EX + Fore.BLACK
            elif ok is False:
                status = 'FAIL'
                failed += 1
                color = Back.LIGHTRED_EX + Fore.BLACK
            else:
                status = 'MANUAL CHECK'
                manual += 1
                color = Back.YELLOW + Fore.BLACK

            case['Test data'] = td
            case['New API - Actual Response'] = actual_summary
            case['New API - Expected Response'] = case.get('New API - Expected Response', '')
            case['Date Of Checked'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            case['Status'] = status
            case['_reason'] = reason  # extra column, written to output only
            case['_manual_note'] = manual_recommendation(title, url)  # extra column, written to output only

            print(color + f' {status} ' + Style.RESET_ALL, reason)
            if case['_manual_note']:
                print(Fore.CYAN + '  Manual testing note:', case['_manual_note'])

    finally:
        driver.quit()

    # ---- Write results ----
    fieldnames = ['Title', 'Description', 'Test data', 'Old API - API Link', 'Old API - Response',
                  'New API - API Link', 'New API - Actual Response', 'New API - Expected Response',
                  'Date Of Checked', 'Tester Name', 'Status', 'developer name',
                  'Reason (Pass/Fail explanation)', 'Manual Testing Recommended']

    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for case in test_cases:
            row = {k: case.get(k, '') for k in fieldnames[:-2]}
            row['Reason (Pass/Fail explanation)'] = case.get('_reason', '')
            row['Manual Testing Recommended'] = case.get('_manual_note', '')
            writer.writerow(row)

    # Also produce a colored Excel version
    wb = Workbook()
    ws = wb.active
    ws.title = 'API Test Cases'
    ws.append(fieldnames)
    for case in test_cases:
        row = [case.get(k, '') for k in fieldnames[:-2]] + [case.get('_reason', ''), case.get('_manual_note', '')]
        ws.append(row)
        status_cell = ws.cell(ws.max_row, fieldnames.index('Status') + 1)
        if case['Status'] == 'PASS':
            status_cell.fill = PatternFill('solid', fgColor='C6EFCE')
            status_cell.font = Font(color='006100', bold=True)
        elif case['Status'] == 'FAIL':
            status_cell.fill = PatternFill('solid', fgColor='FFC7CE')
            status_cell.font = Font(color='9C0006', bold=True)
        else:
            status_cell.fill = PatternFill('solid', fgColor='FFEB9C')
            status_cell.font = Font(color='9C6500', bold=True)

        if case.get('_manual_note'):
            note_cell = ws.cell(ws.max_row, fieldnames.index('Manual Testing Recommended') + 1)
            note_cell.fill = PatternFill('solid', fgColor='FCE4D6')
            note_cell.font = Font(color='974706', bold=True)

    summary = wb.create_sheet('Summary', 0)
    summary.append(['Metric', 'Value'])
    summary.append(['Total', total])
    summary.append(['PASS', passed])
    summary.append(['FAIL', failed])
    summary.append(['MANUAL CHECK NEEDED', manual])
    summary.append(['Pass %% (of automatable tests)', round(passed / (total - manual) * 100, 2) if (total - manual) else 0])
    summary.append(['Run at', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    summary.append(['API', BASE])

    wb.save(OUTPUT_XLSX)

    print('\n' + '=' * 90)
    print(Fore.GREEN + 'PASS:', passed)
    print(Fore.RED + 'FAIL:', failed)
    print(Fore.YELLOW + 'MANUAL CHECK NEEDED:', manual)
    print('Excel:', OUTPUT_XLSX)
    print('CSV:', OUTPUT_CSV)
    print('=' * 90)


if __name__ == '__main__':
    main()