"""
PDO Riyah QHSE Dashboard - Flask Server
----------------------------------------
Reads MY_WORK.xlsx from OneDrive synced folder
Serves live data to office network users
"""

from flask import Flask, jsonify, render_template, send_from_directory
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

# ── CONFIG ──────────────────────────────────────────────
# Apna OneDrive path yahan daalo (Windows pe yahi hota hai)
# Example: C:/Users/YourName/OneDrive - CompanyName/ProjectFolder/MY_WORK.xlsx
EXCEL_PATH = r"C:\Users\Shaghaf Ahmed\OneDrive\Desktop\Work\AUTOMATE EXCEL\MY_WORK.xlsx"
# ────────────────────────────────────────────────────────


def read_sheet(sheet_name, header_row=1):
    """Read a sheet from Excel, skip empty rows"""
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    # Remove header-repeat rows and empty rows
    df = df.dropna(how='all')
    return df


def safe_str(val):
    if pd.isna(val):
        return "—"
    if isinstance(val, datetime):
        return val.strftime("%d %b %Y")
    return str(val).strip()


def parse_r_sheet(sheet_name):
    """Parse R1 or R2 document register sheet"""
    df = read_sheet(sheet_name)
    # Set proper column names
    df.columns = ['Doc_Code','Doc_Name','Rev','Deadline','Review_Status',
                  'Issued_For','Sender','Recipient','Doc_Type','Prev_Status','Date']
    # Drop the first row if it's still a header
    df = df[df['Doc_Code'] != 'Doc-code'].reset_index(drop=True)
    df = df[df['Doc_Code'].notna()].reset_index(drop=True)

    docs = []
    for _, row in df.iterrows():
        docs.append({
            'code':        safe_str(row['Doc_Code']),
            'name':        safe_str(row['Doc_Name']),
            'rev':         safe_str(row['Rev']),
            'deadline':    safe_str(row['Deadline']),
            'status':      safe_str(row['Review_Status']),
            'doc_type':    safe_str(row['Doc_Type']),
            'prev_status': safe_str(row['Prev_Status']),
            'date':        safe_str(row['Date']),
        })

    # Counts
    statuses = df['Review_Status'].str.strip().value_counts().to_dict()
    doc_types = {}
    for t in df['Doc_Type'].dropna():
        key = str(t).split('-')[0].strip()
        doc_types[key] = doc_types.get(key, 0) + 1

    prev = {}
    for p in df['Prev_Status'].dropna():
        ps = str(p).strip()
        if 'Approved with' in ps:   prev['AWC'] = prev.get('AWC', 0) + 1
        elif 'Review with' in ps:   prev['RWC'] = prev.get('RWC', 0) + 1
        elif 'Rejected' in ps:      prev['REJ'] = prev.get('REJ', 0) + 1
        elif ps == '-/-':           prev['First'] = prev.get('First', 0) + 1

    return {
        'total':    len(docs),
        'statuses': statuses,
        'doc_types': doc_types,
        'prev_status': prev,
        'docs':     docs
    }


def parse_ncr_sheet(sheet_name, is_client=False):
    """Parse Internal NCR or Client NCR sheet"""
    df = read_sheet(sheet_name)

    if is_client:
        df.columns = ['SNo','NCR_No','Date_Issued','Category','Location',
                      'Contractor','Originator','Object','Completion_Date','Status']
        df = df[df['SNo'] != 'S No'].dropna(subset=['NCR_No']).reset_index(drop=True)
    else:
        df.columns = ['SNo','NCR_No','Date_Issued','Location','Area',
                      'Contractor','Issuer','Verified_By','Completion_Date','Status']
        df = df[df['SNo'] != 'S.NO'].dropna(subset=['Contractor']).reset_index(drop=True)

    df['Status'] = df['Status'].str.strip()
    df['Contractor'] = df['Contractor'].str.strip()

    ncrs = []
    for _, row in df.iterrows():
        rec = {
            'sno':             safe_str(row['SNo']),
            'ncr_no':          safe_str(row['NCR_No']),
            'date_issued':     safe_str(row['Date_Issued']),
            'status':          safe_str(row['Status']),
            'location':        safe_str(row['Location']),
            'contractor':      safe_str(row['Contractor']),
            'completion_date': safe_str(row['Completion_Date']),
        }
        if is_client:
            rec['originator'] = safe_str(row['Originator'])
            rec['object']     = safe_str(row['Object'])
            rec['category']   = safe_str(row['Category'])
        else:
            rec['area']       = safe_str(row['Area'])
            rec['issuer']     = safe_str(row['Issuer'])
        ncrs.append(rec)

    # Stats
    by_status     = df['Status'].value_counts().to_dict()
    by_contractor = df.groupby(['Contractor','Status']).size().unstack(fill_value=0).to_dict('index')
    by_location   = df['Location'].str.strip().value_counts().to_dict()

    # Monthly trend
    df['Date_Issued'] = pd.to_datetime(df['Date_Issued'], errors='coerce')
    monthly = df.dropna(subset=['Date_Issued'])
    monthly = monthly.groupby(monthly['Date_Issued'].dt.to_period('M')).size()
    monthly_data = {str(k): int(v) for k, v in monthly.items()}

    return {
        'total':          len(ncrs),
        'by_status':      {k: int(v) for k,v in by_status.items()},
        'by_contractor':  {k: {s: int(v) for s,v in sv.items()} for k,sv in by_contractor.items()},
        'by_location':    {k: int(v) for k,v in by_location.items()},
        'monthly':        monthly_data,
        'records':        ncrs,
        'open_records':   [r for r in ncrs if r['status'] == 'Open'],
    }


# ── ROUTES ──────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('.', 'PDO_Riyah_QHSE_Dashboard.html')


@app.route('/api/data')
def get_data():
    """Main API - returns ALL data fresh from Excel"""
    try:
        r1   = parse_r_sheet('R1')
        r2   = parse_r_sheet('R2')
        ncr  = parse_ncr_sheet('NCRs', is_client=False)
        cncr = parse_ncr_sheet('Client_NCRs', is_client=True)

        return jsonify({
            'success':    True,
            'updated_at': datetime.now().strftime("%d %b %Y, %H:%M"),
            'r1':         r1,
            'r2':         r2,
            'ncr':        ncr,
            'client_ncr': cncr,
        })
    except FileNotFoundError:
        return jsonify({'success': False, 'error': f'Excel file not found at: {EXCEL_PATH}'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/status')
def status():
    """Health check - check if Excel is readable"""
    try:
        mtime = os.path.getmtime(EXCEL_PATH)
        last_modified = datetime.fromtimestamp(mtime).strftime("%d %b %Y, %H:%M")
        return jsonify({'ok': True, 'excel_last_modified': last_modified})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory('.', filename)


# ── START ────────────────────────────────────────────────
if __name__ == '__main__':
    print("=" * 55)
    print("  PDO Riyah QHSE Dashboard Server")
    print("=" * 55)
    print(f"  Excel Path : {EXCEL_PATH}")
    print(f"  Dashboard  : http://localhost:5000")
    print(f"  Network    : http://<YOUR-IP>:5000")
    print("  (Share the network link with your team)")
    print("=" * 55)

    # 0.0.0.0 means accessible from office network
    app.run(host='0.0.0.0', port=5000, debug=False)
