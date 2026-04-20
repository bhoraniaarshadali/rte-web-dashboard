# --
"""
RTE Application Status Bulk Checker — v3.0
============================================
✅ Parallel processing (5x faster)
✅ Fetches: Status, Child Name, Mobile Number, Village (Gam)
✅ WhatsApp integration with real mobile number
✅ Resume support (skip already APPROVED)
✅ Live dashboard via data.js
✅ Sync server for dashboard Force Sync button
"""

import time, re, os, json, sys, threading
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse

# ── CONFIG ─────────────────────────────────────────────────────
GSHEET_URL   = "https://docs.google.com/spreadsheets/d/1baLAUi9REHf_1dMj-RjtNy7Bc88L8vqbbrxzemK6cpA/export?format=xlsx"
OUTPUT_FILE  = "RTE_Status_Results.xlsx"
DATA_JS_FILE = "data.js"
SYNC_PORT    = 5001
MAX_WORKERS  = 5      # Parallel requests — safe for server
DELAY_SEC    = 0.5    # Small delay between batches
MAX_RETRIES  = 3
BASE_URL     = "https://rte.orpgujarat.com/ApplicationFormStatus"
USER_AGENT   = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
# ───────────────────────────────────────────────────────────────

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Thread-safe lock for file writes and prints
WRITE_LOCK = threading.Lock()
PRINT_LOCK = threading.Lock()
GLOBAL_DF  = None


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%H:%M:%S")
    symbols = {
        "INFO":  "ℹ️ ", "OK":    "✅", "SKIP":  "⏩",
        "WARN":  "⚠️ ", "ERROR": "❌", "FETCH": "🌐",
        "SAVE":  "💾", "START": "🚀", "SPEED": "⚡",
    }
    sym = symbols.get(level, "• ")
    with PRINT_LOCK:
        print(f"[{ts}] {sym} {msg}", flush=True)


def make_session():
    """Create a new session per thread"""
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT})
    return s


def get_csrf_token(session):
    resp = session.get(BASE_URL, timeout=15)
    resp.raise_for_status()
    match = re.search(
        r'name="__RequestVerificationToken".*?value="([^"]+)"',
        resp.text, re.DOTALL
    )
    if not match:
        raise ValueError("CSRF token not found")
    return match.group(1)


def fetch_status(app_id: str, dob: str) -> dict:
    """
    Fetch from government portal:
    - status message
    - child name
    - parent mobile number
    - village (gam) — stored but not displayed
    Each thread has its own session.
    """
    result = {
        "status":     "Status not found",
        "child_name": "N/A",
        "mobile":     "N/A",
        "gam":        "N/A",
        "lig":        "N/A",  # Gender
        "area":       "N/A",
        "pincode":    "N/A",
    }

    session = make_session()

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            if attempt > 1:
                log(f"  Retry {attempt}/{MAX_RETRIES} → {app_id}", "WARN")
                time.sleep(3)

            token = get_csrf_token(session)
            resp  = session.post(
                BASE_URL,
                data={
                    "__RequestVerificationToken": token,
                    "ApplicationNumber": app_id.strip(),
                    "DateOfBirth":       dob.strip(),
                },
                headers={
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Origin":       "https://rte.orpgujarat.com",
                    "Referer":      BASE_URL,
                },
                timeout=20,
                allow_redirects=True,
            )

            soup = BeautifulSoup(resp.text, "html.parser")

            # ── 1. Status ───────────────────────────────────────
            for fs in soup.select("fieldset"):
                legend = fs.find("legend")
                if legend and "અરજીની સ્થિતિ" in legend.get_text():
                    p = fs.find("p")
                    result["status"] = (
                        p.get_text(strip=True) if p
                        else fs.get_text().replace(legend.get_text(), "").strip()
                    )
                    break
            else:
                err_div = soup.find("div", class_="alert")
                if err_div:
                    result["status"] = err_div.get_text(strip=True)

            # ── 2. Child Name (બાળકની માહિતી) ──────────────────
            for fs in soup.select("fieldset"):
                legend = fs.find("legend")
                if legend and "બાળકની માહિતી" in legend.get_text():
                    first = middle = surname = ""
                    for dt in fs.select("dt"):
                        label = dt.get_text(strip=True)
                        dd    = dt.find_next_sibling("dd")
                        val   = dd.get_text(strip=True) if dd else ""
                        if "બાળકનું નામ" in label:
                            first = val
                        elif "પિતા" in label or "વાલી" in label:
                            middle = val
                        elif "અટક" in label:
                            surname = val
                        elif "લિગ" in label or "લિગ" in label: # Gender
                            result["lig"] = val.strip().upper()
                    name = " ".join(p for p in [first, middle, surname] if p)
                    if name:
                        result["child_name"] = name
                    break

            # ── 3. Mobile Number (સંપર્કની માહિતી) ─────────────
            for fs in soup.select("fieldset"):
                legend = fs.find("legend")
                if legend and "સંપર્કની માહિતી" in legend.get_text():
                    for dt in fs.select("dt"):
                        label = dt.get_text(strip=True)
                        dd    = dt.find_next_sibling("dd")
                        val   = dd.get_text(strip=True) if dd else ""
                        # "મોબાઇલ" = Mobile
                        if "મોબાઇલ" in label and val:
                            # Take first mobile found (primary mobile)
                            result["mobile"] = val.strip()
                            break
                    break

            # ── 4. Village / Gam (Address fieldset) ─────────────
            for fs in soup.select("fieldset"):
                legend = fs.find("legend")
                if legend and ("Address" in legend.get_text() or
                               "સરનામ" in legend.get_text()):
                    for dt in fs.select("dt"):
                        label = dt.get_text(strip=True)
                        dd    = dt.find_next_sibling("dd")
                        val   = dd.get_text(strip=True) if dd else ""
                        if "ગામ" in label and val:
                            result["gam"] = val.strip()
                        elif "વિસ્તાર" in label and val:
                            result["area"] = val.strip()
                        elif "પીનકોડ" in label and val:
                            result["pincode"] = val.strip()
                    break

            return result

        except requests.exceptions.ConnectionError as e:
            log(f"  Connection error [{app_id}]: {e}", "ERROR")
            if attempt == MAX_RETRIES:
                result["status"] = "ERROR: Connection failed"
                return result
        except Exception as e:
            log(f"  Error [{app_id}]: {e}", "ERROR")
            result["status"] = f"ERROR: {str(e)}"
            return result

    return result


def classify_status(status: str) -> str:
    s = status.lower()
    # SUBMITTED check first — contains 'manjur' but is NOT yet approved
    if any(x in s for x in ["બાકી છે", "સમીક્ષા (મંજૂ", "સમીક્ષા (મંજુ",
                              "bakī che", "samīkṣā"]):
        return "SUBMITTED"
    if any(x in s for x in ["મંજૂર", "મંજુર", "ફાળવ", "approved", "approve"]):
        return "APPROVED"
    if any(x in s for x in ["નામંજૂ", "નામંજુ", "cancel", "કેન્સલ",
                              "reject", "refused", "રદ"]):
        return "ERROR"
    if "error:" in s or "not found" in s:
        return "ERROR"
    return "PENDING"


def format_dob(val) -> str:
    if isinstance(val, datetime):
        return val.strftime("%d-%m-%Y")
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return str(val).split(" ")[0]


def export_data_js(df):
    """Write data.js for live dashboard"""
    summary = {
        "total":        int(len(df)),
        "approved":     int((df["Result"] == "APPROVED").sum()),
        "submitted":    int((df["Result"] == "SUBMITTED").sum()),
        "pending":      int((df["Result"] == "PENDING").sum()),
        "error":        int((df["Result"] == "ERROR").sum()),
        "girls":        int((df["Gender"] == "GIRL").sum()),
        "boys":         int((df["Gender"] == "BOY").sum()),
        "last_updated": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
    }
    records = []
    for _, row in df.iterrows():
        records.append({
            "Token No":          str(row.get("Token No", "")),
            "Application Id":    str(row.get("Application Id", "")),
            "Child Name":        str(row.get("Child Name", "N/A")),
            "DOB":               str(row.get("DOB", "")),
            "Mobile":            str(row.get("Mobile", "N/A")),
            "Gender":            str(row.get("Gender", "N/A")),
            "Area":              str(row.get("Area", "N/A")),
            "Pincode":           str(row.get("Pincode", "N/A")),
            "Gam":               str(row.get("Gam", "N/A")),
            "Filled By":         str(row.get("કોને ફોર્મ ભર્યું છે?", "")),
            "Status (Gujarati)": str(row.get("Status (Gujarati)", "")),
            "Result":            str(row.get("Result", "PENDING")),
        })
    js = (
        f"const RTE_SUMMARY = {json.dumps(summary, ensure_ascii=False, indent=2)};\n"
        f"const RTE_DATA = {json.dumps(records, ensure_ascii=False, indent=2)};\n"
    )
    with open(DATA_JS_FILE, "w", encoding="utf-8") as f:
        f.write(js)


def save_excel(df):
    df.to_excel(OUTPUT_FILE, index=False)
    wb = load_workbook(OUTPUT_FILE)
    ws = wb.active
    fills = {
        "APPROVED":  PatternFill("solid", start_color="C6EFCE"),
        "SUBMITTED": PatternFill("solid", start_color="BDD7EE"),
        "PENDING":   PatternFill("solid", start_color="FFEB9C"),
        "ERROR":     PatternFill("solid", start_color="FFC7CE"),
    }
    header_fill = PatternFill("solid", start_color="1F4E79")
    thin   = Side(border_style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = Font(bold=True, color="FFFFFF", size=11, name="Arial")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border
    ws.row_dimensions[1].height = 28
    result_col = ws.max_column
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cat  = row[result_col - 1].value or ""
        fill = fills.get(cat, PatternFill())
        for cell in row:
            cell.border    = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font      = Font(name="Arial", size=10)
        row[result_col - 1].fill = fill
        row[result_col - 1].font = Font(name="Arial", size=10, bold=True)
        row[result_col - 2].fill = fill
    col_widths = {1:22, 2:10, 3:28, 4:18, 5:20, 6:10, 7:20, 8:12, 9:20, 10:22, 11:55, 12:16}
    for col, w in col_widths.items():
        if col <= ws.max_column:
            ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    wb.save(OUTPUT_FILE)


# ── Sync Server ────────────────────────────────────────────────

class SyncHandler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        self.end_headers()

    def do_GET(self):
        global GLOBAL_DF
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        if parsed.path == '/sync':
            app_id = params.get('app_id', [None])[0]
            if app_id and GLOBAL_DF is not None:
                success = sync_single(app_id)
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps(
                    {"status": "success" if success else "failed"}
                ).encode())
                return
        self.send_response(404)
        self.end_headers()

    def log_message(self, format, *args):
        return  # silence server logs


def sync_single(app_id):
    global GLOBAL_DF
    idx_list = GLOBAL_DF.index[
        GLOBAL_DF['Application Id'].astype(str).str.strip() == app_id.strip()
    ].tolist()
    if not idx_list:
        return False
    idx = idx_list[0]
    dob = str(GLOBAL_DF.at[idx, "DOB"]).strip()
    data = fetch_status(app_id, dob)
    with WRITE_LOCK:
        GLOBAL_DF.at[idx, "Status (Gujarati)"] = data["status"]
        GLOBAL_DF.at[idx, "Child Name"]        = data["child_name"]
        GLOBAL_DF.at[idx, "Mobile"]            = data["mobile"]
        GLOBAL_DF.at[idx, "Gender"]            = data["lig"]
        GLOBAL_DF.at[idx, "Area"]              = data["area"]
        GLOBAL_DF.at[idx, "Pincode"]           = data["pincode"]
        GLOBAL_DF.at[idx, "Gam"]               = data["gam"]
        GLOBAL_DF.at[idx, "Result"]            = classify_status(data["status"])
        try:
            save_excel(GLOBAL_DF)
            export_data_js(GLOBAL_DF)
        except:
            pass
    log(f"Remote sync: {app_id} → {data['child_name']} | {data['mobile']}", "OK")
    return True


def start_server():
    server = HTTPServer(('localhost', SYNC_PORT), SyncHandler)
    log(f"Sync server on http://localhost:{SYNC_PORT}", "START")
    server.serve_forever()


# ── Parallel Worker ────────────────────────────────────────────

def process_record(args):
    """Worker function for ThreadPoolExecutor"""
    idx, app_id, dob, num, total = args
    data   = fetch_status(app_id, dob)
    cat    = classify_status(data["status"])
    result = {
        "idx":    idx,
        "status": data["status"],
        "child":  data["child_name"],
        "mobile": data["mobile"],
        "elig":   data["lig"],
        "earea":  data["area"],
        "epin":   data["pincode"],
        "gam":    data["gam"],
        "cat":    cat,
        "app_id": app_id,
        "num":    num,
        "total":  total,
    }
    return result


# ── Main ───────────────────────────────────────────────────────

def main():
    global GLOBAL_DF
    print("=" * 65)
    print("   RTE Application Status Bulk Checker  v3.0")
    print(f"   ⚡ Parallel Mode: {MAX_WORKERS} simultaneous checks")
    print("=" * 65)

    # Step 1: Load Google Sheets
    log("Connecting to Google Sheets...", "FETCH")
    try:
        df_source = pd.read_excel(GSHEET_URL)
        log(f"Downloaded {len(df_source)} records", "OK")
    except Exception as e:
        log(f"Failed: {e}", "ERROR")
        return

    # Step 2: Load cache
    cache = {}
    if os.path.exists(OUTPUT_FILE):
        try:
            df_ex = pd.read_excel(OUTPUT_FILE)
            for _, r in df_ex.iterrows():
                aid = str(r.get("Application Id", "")).strip()
                if aid:
                    cache[aid] = {
                        "Result": str(r.get("Result", "PENDING")),
                        "Status": str(r.get("Status (Gujarati)", "")),
                        "Child":  str(r.get("Child Name", "N/A")),
                        "Mobile": str(r.get("Mobile", "N/A")),
                        "Gender": str(r.get("Gender", "N/A")),
                        "Area":   str(r.get("Area", "N/A")),
                        "Pincode":str(r.get("Pincode", "N/A")),
                        "Gam":    str(r.get("Gam", "N/A")),
                    }
            log(f"Cache loaded: {len(cache)} previous results", "OK")
        except Exception as e:
            log(f"Cache load failed: {e}", "WARN")
    else:
        log("Starting fresh — no previous results", "INFO")

    # Step 3: Prepare DataFrame
    GLOBAL_DF = df_source.copy()
    GLOBAL_DF["DOB"]               = GLOBAL_DF["બાળક ની જન્મ તારીખ "].apply(format_dob)
    GLOBAL_DF["Child Name"]        = "N/A"
    GLOBAL_DF["Mobile"]            = "N/A"
    GLOBAL_DF["Gender"]            = "N/A"
    GLOBAL_DF["Area"]              = "N/A"
    GLOBAL_DF["Pincode"]           = "N/A"
    GLOBAL_DF["Gam"]               = "N/A"
    GLOBAL_DF["Status (Gujarati)"] = ""
    GLOBAL_DF["Result"]            = "PENDING"

    for i, row in GLOBAL_DF.iterrows():
        aid = str(row["Application Id"]).strip()
        if aid in cache:
            GLOBAL_DF.at[i, "Result"]            = cache[aid]["Result"]
            GLOBAL_DF.at[i, "Status (Gujarati)"] = cache[aid]["Status"]
            GLOBAL_DF.at[i, "Child Name"]        = cache[aid]["Child"]
            GLOBAL_DF.at[i, "Mobile"]            = cache[aid]["Mobile"]
            GLOBAL_DF.at[i, "Gender"]            = cache[aid]["Gender"]
            GLOBAL_DF.at[i, "Area"]              = cache[aid]["Area"]
            GLOBAL_DF.at[i, "Pincode"]           = cache[aid]["Pincode"]
            GLOBAL_DF.at[i, "Gam"]               = cache[aid]["Gam"]

    total    = len(GLOBAL_DF)
    approved = int((GLOBAL_DF["Result"] == "APPROVED").sum())
    to_check = total - approved
    log(f"Total: {total} | Approved (skip): {approved} | To check: {to_check}", "INFO")

    # Estimate time
    est_min = (to_check / MAX_WORKERS) * 3 / 60
    log(f"Estimated time: ~{est_min:.0f} min (was ~{to_check * 3 / 60:.0f} min sequential)", "SPEED")
    print("-" * 65)

    # Step 4: Start sync server
    threading.Thread(target=start_server, daemon=True).start()

    # Step 5: Build work queue
    work_queue = []
    for i, row in GLOBAL_DF.iterrows():
        if GLOBAL_DF.at[i, "Result"] == "APPROVED":
            continue
        app_id = str(row["Application Id"]).strip()
        dob    = str(row["DOB"]).strip()
        work_queue.append((i, app_id, dob, i + 1, total))

    # Step 6: Parallel processing
    checked = skipped = errors = 0
    skipped = approved  # already counted

    completed_count = 0
    batch_size = MAX_WORKERS * 4  # save every N completions

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_record, args): args for args in work_queue}

        for future in as_completed(futures):
            try:
                res = future.result()
                idx    = res["idx"]
                app_id = res["app_id"]
                cat    = res["cat"]
                num    = res["num"]

                # Update DataFrame (thread-safe)
                with WRITE_LOCK:
                    GLOBAL_DF.at[idx, "Status (Gujarati)"] = res["status"]
                    GLOBAL_DF.at[idx, "Child Name"]        = res["child"]
                    GLOBAL_DF.at[idx, "Mobile"]            = res["mobile"]
                    GLOBAL_DF.at[idx, "Gender"]            = res["elig"]
                    GLOBAL_DF.at[idx, "Area"]              = res["earea"]
                    GLOBAL_DF.at[idx, "Pincode"]           = res["epin"]
                    GLOBAL_DF.at[idx, "Gam"]               = res["gam"]
                    GLOBAL_DF.at[idx, "Result"]            = cat

                checked += 1
                completed_count += 1
                if cat == "ERROR":
                    errors += 1

                # Log result
                mobile_short = res["mobile"][:10] if res["mobile"] != "N/A" else "N/A"
                if cat == "APPROVED":
                    log(f"[{num:3d}/{total}] ✅ {app_id} | {res['child']} | 📱{mobile_short}", "OK")
                elif cat == "ERROR":
                    log(f"[{num:3d}/{total}] ❌ {app_id} | {res['status'][:50]}", "ERROR")
                elif cat == "SUBMITTED":
                    log(f"[{num:3d}/{total}] 🔵 {app_id} | {res['child']} | 📱{mobile_short}", "INFO")
                else:
                    log(f"[{num:3d}/{total}] ⏳ {app_id} | {res['child']} | 📱{mobile_short}", "INFO")

                # Save every batch_size completions
                if completed_count % batch_size == 0 or completed_count == len(work_queue):
                    with WRITE_LOCK:
                        try:
                            save_excel(GLOBAL_DF)
                            export_data_js(GLOBAL_DF)
                            done_a = int((GLOBAL_DF["Result"] == "APPROVED").sum())
                            done_s = int((GLOBAL_DF["Result"] == "SUBMITTED").sum())
                            done_p = int((GLOBAL_DF["Result"] == "PENDING").sum())
                            print(f"\n{'─'*65}")
                            log(f"SAVED — ✅{done_a} approved | 🔵{done_s} submitted | ⏳{done_p} pending | ❌{errors} errors")
                            print(f"{'─'*65}\n")
                        except Exception as e:
                            log(f"Save failed: {e}", "WARN")

            except Exception as e:
                log(f"Future error: {e}", "ERROR")

    # Step 7: Final save
    with WRITE_LOCK:
        save_excel(GLOBAL_DF)
        export_data_js(GLOBAL_DF)

    # Summary
    fa = int((GLOBAL_DF["Result"] == "APPROVED").sum())
    fs = int((GLOBAL_DF["Result"] == "SUBMITTED").sum())
    fp = int((GLOBAL_DF["Result"] == "PENDING").sum())
    fe = int((GLOBAL_DF["Result"] == "ERROR").sum())

    print("\n" + "=" * 65)
    print("  ✅ COMPLETE!")
    print("=" * 65)
    print(f"\n  ✅  APPROVED   : {fa}")
    print(f"  🔵  SUBMITTED  : {fs}")
    print(f"  ⏳  PENDING    : {fp}")
    print(f"  ❌  ERROR      : {fe}")
    print(f"  {'─'*17}")
    print(f"  📊  TOTAL      : {total}")
    print(f"  ⏩  SKIPPED    : {skipped}")
    print(f"  🌐  CHECKED    : {checked}")
    print(f"\n  💾  Excel  → {OUTPUT_FILE}")
    print(f"  🌍  Dashboard → Open dashboard.html in browser")
    print(f"\n  Sync server still running on port {SYNC_PORT}.")
    print("  Press Ctrl+C to stop.\n")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()