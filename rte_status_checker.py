import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import os
import json
from datetime import datetime
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse

# --- CONFIGURATION ---
GSHEET_URL = 'https://docs.google.com/spreadsheets/d/1baLAUi9REHf_1dMj-RjtNy7Bc88L8vqbbrxzemK6cpA/export?format=xlsx'
OUTPUT_FILE = 'RTE_Status_Results.xlsx'
DATA_JS = 'data.js'
# Set to True for continuous loop, False for single run
AUTO_LOOP = True
LOOP_DELAY = 900  # 15 minutes
PORT = 5001      # Port for local sync server

# GLOBAL DATAFRAME FOR LIVE SYNC
df_main = None

class RteSyncHandler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'X-Requested-With')
        self.end_headers()

    def do_GET(self):
        global df_main
        query = urllib.parse.urlparse(self.path).query
        params = urllib.parse.parse_qs(query)
        
        # Root path info
        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b"<h1>RTE Sync Server is Running</h1><p>The dashboard uses this to sync records.</p>")
            return

        if self.path.startswith('/sync'):
            app_id = params.get('app_id', [None])[0]
            if app_id:
                print(f"\n[REMOTE SYNC] Triggered for: {app_id}")
                success = sync_single_record(app_id)
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "success" if success else "failed"}).encode())
                return

        self.send_response(404)
        self.end_headers()

def fetch_status(app_id, dob):
    """Fetches status from government portal"""
    url = "https://rte.orpgujarat.com/ApplicationStatus"
    data = {
        "ApplicationNumber": app_id,
        "BirthDate": dob
    }
    
    try:
        # Simulate browser headers
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.post(url, data=data, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Look for status message in common alert boxes or fieldsets
        status_box = soup.find('div', class_='alert') or soup.find('fieldset')
        if status_box:
            status_text = status_box.get_text(separator=" ").strip()
            # Clean up extra spaces/newlines
            status_text = " ".join(status_text.split())
            return status_text
        return "Status information not found on page."
    except Exception as e:
        return f"Error: {str(e)}"

def classify_status(msg):
    """Categorizes status for the dashboard"""
    msg = msg.upper()
    if any(k in msg for k in ['MANJUR', 'APPROVE', 'અપ્રૂવ', 'મંજૂર']):
        return 'APPROVED'
    if any(k in msg for k in ['REJECT', 'ERROR', 'CANCEL', 'કૅન્સલ', 'રદ']):
        return 'ERROR'
    if any(k in msg for k in ['SUBMIT', 'PENDING AT DISTRICT', 'સબમિટ']):
        return 'SUBMITTED'
    return 'PENDING'

def export_to_web(df):
    """Exports current data to data.js for dashboard"""
    if df is None: return
    
    summary = {
        "total": len(df),
        "approved": len(df[df['Result'] == 'APPROVED']),
        "submitted": len(df[df['Result'] == 'SUBMITTED']),
        "pending": len(df[df['Result'] == 'PENDING']),
        "error": len(df[df['Result'] == 'ERROR']),
        "last_updated": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    }
    
    records = df.to_dict(orient='records')
    js_content = f"const RTE_SUMMARY = {json.dumps(summary, indent=4)};\n"
    js_content += f"const RTE_DATA = {json.dumps(records, indent=4)};"
    
    with open(DATA_JS, 'w', encoding='utf-8') as f:
        f.write(js_content)
    print(f"[*] Dashboard data updated: {summary['last_updated']}")

def sync_single_record(app_id):
    """Force sync a single record by App ID"""
    global df_main
    if df_main is None: return False
    
    idx_list = df_main.index[df_main['Application Id'] == app_id].tolist()
    if not idx_list: return False
    
    idx = idx_list[0]
    
    # Robustly find DOB column
    dob_col = next((c for c in df_main.columns if 'જન્મ તારીખ' in c or 'Birth Date' in c), None)
    if not dob_col: return False
    
    dob = str(df_main.at[idx, dob_col])
    if ' ' in dob: dob = dob.split(' ')[0] 
    
    print(f"Connecting for {app_id}...")
    status_msg = fetch_status(app_id, dob)
    df_main.at[idx, 'Status (Gujarati)'] = status_msg
    df_main.at[idx, 'Result'] = classify_status(status_msg)
    
    # Save results
    df_main.to_excel(OUTPUT_FILE, index=False)
    export_to_web(df_main)
    return True

def run_background_sync():
    """Main auto-pilot loop"""
    global df_main
    while True:
        print(f"\n[AUTO-PILOT] Starting sync loop: {datetime.now().strftime('%H:%M:%S')}")
        
        # Load master file
        # Load master file from Google Sheets
        try:
            df_new = pd.read_excel(GSHEET_URL)
        except Exception as e:
            print(f"[!] Error fetching Google Sheet: {e}")
            time.sleep(60)
            continue
        
        # Robustly find DOB column if not already renamed
        dob_col = next((c for c in df_new.columns if 'જન્મ તારીખ' in c or 'Birth Date' in c), 'Birth Date')
        if dob_col in df_new.columns and dob_col != 'Birth Date':
            df_new = df_new.rename(columns={dob_col: 'Birth Date'})
        
        col_map = {
            'Application Id': 'Application Id',
            'Token No': 'Token No',
            'Student First Name': 'First Name',
            'Father/Guardian Name': 'Father Name',
            'Surname': 'Surname'
        }
        # Rename other columns
        df_new.columns = [col_map.get(c.strip(), c.strip()) for c in df_new.columns]

        # Check for Child Name parts
        if 'Child Name' not in df_new.columns:
            parts = ['First Name', 'Father Name', 'Surname']
            available_parts = [p for p in parts if p in df_new.columns]
            if available_parts:
                df_new['Child Name'] = df_new[available_parts].fillna('').agg(' '.join, axis=1)
            else:
                df_new['Child Name'] = '-'
        
        # Merge with existing df_main to preserve Mobile/Status
        if df_main is None:
            df_main = df_new
            if 'Result' not in df_main.columns:
                df_main['Result'] = 'PENDING'
                df_main['Status (Gujarati)'] = ''
        else:
            # Sync only missing records from df_new into df_main
            for _, row in df_new.iterrows():
                app_id = str(row['Application Id'])
                if app_id not in df_main['Application Id'].astype(str).values:
                    # New record found
                    new_row = row.copy()
                    new_row['Result'] = 'PENDING'
                    new_row['Status (Gujarati)'] = ''
                    df_main = pd.concat([df_main, pd.DataFrame([new_row])], ignore_index=True)

        # Process Pending/Submitted Items
        for index, row in df_main.iterrows():
            if row['Result'] in ['APPROVED', 'ERROR']: continue
            
            app_id = str(row['Application Id']).strip()
            dob_val = row.get('Birth Date')
            dob = str(dob_val).split(' ')[0] if pd.notna(dob_val) else ''
            
            if not dob or dob == 'nan' or dob == 'None':
                print(f"Skipping {app_id}: Invalid DOB '{dob_val}'")
                continue
            
            print(f"Checking {app_id}...")
            status = fetch_status(app_id, dob)
            df_main.at[index, 'Status (Gujarati)'] = status
            df_main.at[index, 'Result'] = classify_status(status)
            time.sleep(1)
        
        export_to_web(df_main)
        df_main.to_excel(OUTPUT_FILE, index=False)
        print(f"[AUTO-PILOT] Loop complete. Waiting {LOOP_DELAY}s...")
        time.sleep(LOOP_DELAY)

def start_server():
    server = HTTPServer(('localhost', PORT), RteSyncHandler)
    print(f"[*] Sync Server running on http://localhost:{PORT}")
    server.serve_forever()

if __name__ == "__main__":
    # Load initial data from local results if exists, otherwise from GSheet
    def apply_mapping(df):
        # Robustly find DOB column
        dob_col = next((c for c in df.columns if 'જન્મ તારીખ' in c or 'Birth Date' in c), None)
        if dob_col:
            df = df.rename(columns={dob_col: 'Birth Date'})
        
        col_map = {
            'Application Id': 'Application Id',
            'Token No': 'Token No'
        }
        df.columns = [col_map.get(c.strip(), c.strip()) for c in df.columns]
        return df

    if os.path.exists(OUTPUT_FILE):
        print(f"[*] Loading existing results from {OUTPUT_FILE}")
        df_main = apply_mapping(pd.read_excel(OUTPUT_FILE))
    else:
        print(f"[*] Fetching initial data from Google Sheets...")
        try:
            df_main = apply_mapping(pd.read_excel(GSHEET_URL))
            df_main['Result'] = 'PENDING'
            df_main['Status (Gujarati)'] = ''
        except Exception as e:
            print(f"[!] Critical Error: Could not fetch initial data. {e}")
            exit(1)
    
    # Start Sync Server in thread
    thread_srv = threading.Thread(target=start_server, daemon=True)
    thread_srv.start()
    
    # Start Background Loop
    run_background_sync()