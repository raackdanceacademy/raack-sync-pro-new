import os
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import pandas as pd
import zipfile
import io
import tempfile
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import numpy as np
import json
import time
from collections import defaultdict
import shutil
# 📊 Branch Data Management System

# ## Project Information
# - **Project Name**: Branch Data Management & Google Sheets Automation System
# - **Version**: 2.0
# - **Developed By**: Vishali Nandhakumar
# - **Developer Role**: Data Analyst
# - **Development Period**: [Month Year] - [Month Year]
# - **Current Organization**: Raack Academy of Dance
# - **Contact Email**: vishalinandha2002@gmail.com
# - **GitHub**: https://github.com/VishaliNandhakumar

# ## Project Overview
# This application automates the process of handling branch-wise bill reports by integrating Excel file uploads with Google Sheets. It eliminates manual data entry, prevents duplicate records, and provides organized data export capabilities across 45+ branches.

# ## Technical Stack
# | Technology | Purpose |
# |------------|---------|
# | Python 3.x | Core programming language |
# | Flask Framework | Web application backend |
# | Pandas & NumPy | Data processing and manipulation |
# | Google Sheets API (gspread) | Google Sheets integration |
# | PythonAnywhere | Cloud hosting platform |
# | HTML5/CSS3/JavaScript | Frontend interface |

# ## Key Features Implemented by Developer
# 1. **Excel File Processing** - Upload and validate Excel files with specific column requirements
# 2. **Intelligent Duplicate Detection** - Prevents duplicate entries based on (Bill No + Status) combinations
# 3. **Google Sheets Sync** - Real-time synchronization with 7 main sheets and 7 summary sheets
# 4. **Branch-wise Organization** - Automatic worksheet creation for 45+ branches
# 5. **ZIP Export System** - Generates organized ZIP files with status and branch folders
# 6. **Data Validation** - Cleans and standardizes bill numbers (removes spaces, .0 suffixes)
# 7. **Rate Limiting** - Implements 15-second delays to prevent API throttling



app = Flask(__name__)
CORS(app)

# ===============================
# ADD THIS FOR PYTHONANYWHERE
# ===============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ===============================
# UPLOAD CONFIG - MODIFIED FOR PYTHONANYWHERE
# ===============================
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'temp_uploads')
ZIP_FOLDER = os.path.join(BASE_DIR, 'temp_zips')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ZIP_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ZIP_FOLDER'] = ZIP_FOLDER

# ===============================
# MASTER DATA
# ===============================
BRANCHES = [
    "KILPAUK", "MYLAPORE", "VELACHERY", "CUDDALORE", "TAMBARAM", "MOGAPPAIR",
    "THORAIPAKKAM", "AVADI", "KEELKATTALAI", "MUGALIVAKKAM", "SHOLINGANALLUR",
    "NEELANKARAI", "KOLATHUR", "PALLIKARANAI", "OLD PERUNGALATHUR",
    "GUDUVANCHERI", "PUDUCHERRY", "RAMAPURAM", "SAIDAPET", "OLD PALLAVARAM",
    "MANNIVAKKAM", "CHIDAMBARAM", "HASTHINAPURAM", "THIRUVERKADU", "SURAPET",
    "MARAIMALAI NAGAR", "PADUR", "MEDAVAKKAM", "PADAPPAI", "AMBATTUR",
    "ARUMBAKKAM", "AYAPAKKAM", "SITHALAPAKKAM", "PERUMBAKKAM", "BASAVANAGUDI",
    "PUDUPAKKAM", "URAPAKKAM", "THANJAVUR", "PAMMAL", "KUMBAKONAM",
    "MADURAVOYAL", "KANDIGAI", "KUNDRATHUR","MADAMBAKKAM","NAVALUR","IYYAPANTHANGAL","KELAMBAKKAM","MAPPAEDU","VYASARPADI"
]

STATUSES = [
    "Success", "Failure", "Initiated", "Awaited",
    "Timeout", "Unsuccessful", "Aborted"
]

# ===============================
# GOOGLE SHEET IDS - MAIN SHEETS
# ===============================
SHEET_IDS = {
    "Aborted": "1x8cyu1-n7YykmCAcZQ1VcMMtYWEvK4R_J50nqhGKTVg",
    "Awaited": "1Xy_pOmG9rr2u0R8OQVUP8eg1JBss8QmotrVMNFIHb3E",
    "Failure": "1UwI2C9WwlAa4rvZajwZiDuZrYZn6rpXeLS_xoI7OLuY",
    "Initiated": "1XhqOC2hM7T-glTiJp97B9DdxhLcJhfjxy058Ydg9ngs",
    "Success": "1v8IKnleCqpixOFG6vwHrykQO612ImhzKI5J1M14KXL0",
    "Timeout": "1Kd43afefe7rmGcw65MaTIYuPGEvpaq-o3SMrmOJM1vY",
    "Unsuccessful": "1KVPGEY6KcdssAeHYGlJejkMYoBYN2tVjEj7ruD_9zSM"
}

# ===============================
# GOOGLE SHEET IDS - SUMMARY SHEETS (11 columns)
# ===============================
SUMMARY_SHEET_IDS = {
    "Aborted": "1WK7UEDPHhXfTsc7ONmIX4DGfov_qa_I85PW-rZjHbxs",
    "Awaited": "1KaV4nVuCrA2YGcOVB39tkg9XDgrhgKpnIrHH_2PU1FU",
    "Failure": "1ZbIfJ_69ktl6x6ItCgEDmdngIxE3SaHsyt7dX5xZqxk",
    "Initiated": "1eqvwDeJ4VruZZAxn208D8qAqadKiwrKY5LIsQZTTN1c",
    "Success": "12sdd2mBL85mMO7WHic_WACTS5NsXs8hn-ay3IdPHUoM",
    "Timeout": "1vylrQprlx9kYjK18go9PhUFw1gZZyFty66HBHEqDqCA",
    "Unsuccessful": "1qtPV9fNhvIjBHvtuAzoSLXMLRbeOQls37mJDoBPp5UY"
}

# ===============================
# SERVICE ACCOUNT
# ===============================
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "credentials/service_account.json")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ===============================
# 🔥 DATA CLEANING
# ===============================
def clean_dataframe_for_json(df):
    df = df.copy()
    df.replace([float("inf"), float("-inf")], "", inplace=True)
    df.fillna("", inplace=True)
 
    for col in df.columns:
        if df[col].dtype == 'int64':
            df[col] = df[col].astype('Int64')
        elif df[col].dtype == 'float64':
            df[col] = df[col].astype('float')
 
    return df

# ===============================
# JSON SERIALIZER FOR NUMPY TYPES
# ===============================
class NumpyJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.integer, np.int64, np.int32)):
            return int(obj)
        elif isinstance(obj, (np.floating, np.float64, np.float32)):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, pd.Timestamp):
            return obj.strftime('%Y-%m-%d %H:%M:%S')
        elif pd.isna(obj):
            return None
        return super().default(obj)

# ===============================
# GOOGLE AUTH
# ===============================
def get_google_sheets_client():
    try:
        print("🔍 Checking JSON file:", SERVICE_ACCOUNT_FILE)

        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError("Service account JSON file not found")

        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=SCOPES
        )

        gc = gspread.authorize(creds)
        print("✅ Google Sheets authenticated")
        return gc

    except Exception as e:
        print("❌ Google Auth Error:", e)
        return None

# ===============================
# HELPER FUNCTIONS FOR GOOGLE SHEETS
# ===============================
def clean_bill_number(bill_no):
    """Clean and standardize bill number format - REMOVES ALL SPACES and STANDARDIZES"""
    if pd.isna(bill_no):
        return ""
 
    # Convert to string and strip whitespace
    bill_no = str(bill_no).strip()
 
    # Remove .0 if present
    if bill_no.endswith('.0'):
        bill_no = bill_no[:-2]
 
    # Remove ALL whitespace (including spaces in the middle)
    bill_no = ''.join(bill_no.split())
 
    # Convert to uppercase for consistency
    bill_no = bill_no.upper()
 
    return bill_no

def find_empty_row_for_append(worksheet):
    """Find the first empty row to append data"""
    try:
        all_values = worksheet.get_all_values()
 
        last_row = 0
        for i, row in enumerate(all_values, start=1):
            if any(cell and str(cell).strip() for cell in row):
                last_row = i
 
        return last_row + 1
    except Exception as e:
        print(f"Error finding empty row: {e}")
        return 2

def get_existing_records(worksheet):
    """Get all existing (bill_no, status) combinations from the worksheet"""
    records = set()
    try:
        all_values = worksheet.get_all_values()
 
        for row in all_values:
            if len(row) > 17:  # Need at least up to column R (order status)
                bill_no = row[2]  # Column C (index 2)
                status = row[17]  # Column R (index 17) - order status
 
                if bill_no and str(bill_no).strip() and status and str(status).strip():
                    # Apply cleaning to bill number
                    bill_no_clean = clean_bill_number(bill_no)
                    status_clean = str(status).strip()
 
                    # Skip headers
                    if (bill_no_clean and not bill_no_clean.startswith("BILLNO") and
                        not bill_no_clean.startswith("SNO") and
                        bill_no_clean != "" and bill_no_clean != "TOTAL" and
                        status_clean and status_clean != ""):
                        records.add((bill_no_clean, status_clean))
 
        return records
    except Exception as e:
        print(f"Error getting existing records: {e}")
        return set()

def get_existing_bills_from_summary(worksheet):
    """Get all existing bill numbers from summary worksheet"""
    records = set()
    try:
        all_values = worksheet.get_all_values()
 
        for row in all_values:
            if len(row) > 1:
                bill_no = row[1]  # Column B (index 1) - Bill No
 
                if bill_no and str(bill_no).strip():
                    bill_no_clean = clean_bill_number(bill_no)
 
                    # Skip headers
                    if (bill_no_clean and not bill_no_clean.startswith("BILLNO") and
                        not bill_no_clean.startswith("SNO") and
                        bill_no_clean != "" and bill_no_clean != "TOTAL"):
                        records.add(bill_no_clean)
 
        return records
    except Exception as e:
        print(f"Error getting existing bills from summary: {e}")
        return set()

def validate_bill_no_uniqueness(df):
    """Validate that Bill No + Status combination is unique in the uploaded data itself"""
    try:
        # Clean Bill No column
        df = df.copy()
        df["Bill No"] = df["Bill No"].apply(clean_bill_number)
        df["order status"] = df["order status"].apply(lambda x: str(x).strip())
 
        # Remove empty bill numbers
        df = df[df["Bill No"] != ""]
        df = df[df["Bill No"] != "nan"]
        df = df[df["order status"] != ""]
 
        # Check for duplicates based on Bill No + Status combination
        duplicates_in_upload = df[df.duplicated(subset=['Bill No', 'order status'], keep=False)]
 
        if not duplicates_in_upload.empty:
            print(f"⚠️  Found {len(duplicates_in_upload)} duplicate (Bill No + Status) combinations in uploaded data")
 
            # Group by combination to show counts
            combo_counts = duplicates_in_upload.groupby(['Bill No', 'order status']).size()
            print(f"🔍 Duplicate combinations: {combo_counts.to_dict()}")
 
            # Remove duplicates (keep first occurrence of each Bill No+Status combo)
            df_unique = df.drop_duplicates(subset=['Bill No', 'order status'], keep='first')
            removed_count = len(df) - len(df_unique)
            print(f"🗑️  Removed {removed_count} duplicate rows from uploaded data")
 
            return df_unique, removed_count
        else:
            print("✅ No duplicates found within uploaded data")
            return df, 0
 
    except Exception as e:
        print(f"Error validating bill number uniqueness: {e}")
        return df, 0

def convert_numpy_to_python(obj):
    """Convert numpy/pandas types to Python native types"""
    if isinstance(obj, dict):
        return {k: convert_numpy_to_python(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_numpy_to_python(item) for item in obj]
    elif isinstance(obj, (np.integer, np.int64, np.int32)):
        return int(obj)
    elif isinstance(obj, (np.floating, np.float64, np.float32)):
        return float(obj)
    elif isinstance(obj, pd.Timestamp):
        return obj.strftime('%Y-%m-%d %H:%M:%S')
    elif pd.isna(obj):
        return None
    else:
        return obj

def normalize_sheet_name(name):
    """Normalize sheet name by removing extra spaces and special characters"""
    name = str(name).strip()
    for char in ['\\', '/', '*', '?', ':', '[', ']']:
        name = name.replace(char, '_')
    return name[:31]

# ===============================
# ZIP FILE GENERATION FUNCTIONS
# ===============================
def create_excel_with_summary(df, sheet_name, folder_path):
    """Create Excel file with summary for a specific status"""
    try:
        file_path = os.path.join(folder_path, f"{sheet_name}.xlsx")
 
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            summary_data = []
 
            if not df.empty:
                grouped = df.groupby('Branch Name').agg({
                    'Total Bill Amount': 'sum',
                    'Total Discount Amount': 'sum',
                    'Total Tax Amount': 'sum',
                    'Net Amount': 'sum',
                    'Bill No': 'count'
                }).round(2)
 
                grouped = grouped.reset_index()
                grouped.columns = ['Branch Name', 'Total Bill Amount', 'Total Discount Amount',
                                 'Total Tax Amount', 'Net Amount', 'Record Count']
 
                grand_totals = pd.DataFrame({
                    'Branch Name': ['GRAND TOTAL'],
                    'Total Bill Amount': [grouped['Total Bill Amount'].sum()],
                    'Total Discount Amount': [grouped['Total Discount Amount'].sum()],
                    'Total Tax Amount': [grouped['Total Tax Amount'].sum()],
                    'Net Amount': [grouped['Net Amount'].sum()],
                    'Record Count': [grouped['Record Count'].sum()]
                })
 
                summary_df = pd.concat([grouped, grand_totals], ignore_index=True)
                summary_data = summary_df
            else:
                summary_data = pd.DataFrame({
                    'Branch Name': ['No Data Available'],
                    'Total Bill Amount': [0],
                    'Total Discount Amount': [0],
                    'Total Tax Amount': [0],
                    'Net Amount': [0],
                    'Record Count': [0]
                })
 
            summary_data.to_excel(writer, sheet_name='Summary', index=False)
 
            if not df.empty:
                df.to_excel(writer, sheet_name='Detailed Data', index=False)
 
            workbook = writer.book
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 30)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
 
        return file_path
    except Exception as e:
        print(f"Error creating Excel file for {sheet_name}: {e}")
        return None

def generate_zip_files(df):
    """Generate ZIP files organized by status and branch"""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"branch_data_{timestamp}.zip"
        zip_path = os.path.join(app.config['ZIP_FOLDER'], zip_filename)
 
        temp_dir = tempfile.mkdtemp()
        status_files = {}
 
        for status in STATUSES:
            status_df = df[df["order status"] == status]
 
            if status_df.empty:
                print(f"⏭️  No data for status: {status}")
                continue
 
            status_folder = os.path.join(temp_dir, status)
            os.makedirs(status_folder, exist_ok=True)
 
            status_file = create_excel_with_summary(status_df, status, status_folder)
            if status_file:
                status_files[status] = status_file
 
            for branch, branch_df in status_df.groupby("Branch Name"):
                if branch_df.empty:
                    continue
 
                branch_file = create_excel_with_summary(
                    branch_df,
                    f"{branch}_{status}",
                    status_folder
                )
 
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file.endswith('.xlsx'):
                        file_path = os.path.join(root, file)
                        rel_path = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, rel_path)
 
        shutil.rmtree(temp_dir)
 
        return zip_filename, zip_path, len(status_files)
 
    except Exception as e:
        print(f"Error generating ZIP files: {e}")
        return None, None, 0

# ===============================
# ROUTES
# ===============================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({'error': 'No file uploaded'}), 400

        df = pd.read_excel(file)

        required_cols = [
            "Branch Name", "order status", "Bill No",
            "Total Bill Amount", "Total Discount Amount",
            "Total Tax Amount", "Net Amount"
        ]

        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return jsonify({'error': f'Missing columns: {missing}'}), 400

        df = clean_dataframe_for_json(df)
 
        # Clean Bill No column with improved function
        df["Bill No"] = df["Bill No"].apply(clean_bill_number)
        df["order status"] = df["order status"].apply(lambda x: str(x).strip())
        df = df[df["Bill No"] != ""]
        df = df[df["Bill No"] != "nan"]
        df = df[df["order status"] != ""]
 
        # Check for duplicates WITHIN the uploaded file (based on Bill No + Status)
        df, internal_duplicates_removed = validate_bill_no_uniqueness(df)
 
        # Now check for duplicates AGAINST Google Sheets (based on Bill No + Status)
        gc = get_google_sheets_client()
        duplicates_in_google_sheets = {}
        total_duplicates_with_gs = 0
 
        if gc:
            # Group data by status
            for status in STATUSES:
                if status not in SHEET_IDS:
                    continue
 
                status_df = df[df["order status"] == status]
                if status_df.empty:
                    continue
 
                try:
                    spreadsheet = gc.open_by_key(SHEET_IDS[status])
                    all_worksheets = spreadsheet.worksheets()
 
                    for branch, branch_df in status_df.groupby("Branch Name"):
                        ws_name = normalize_sheet_name(branch)
 
                        # Find worksheet
                        ws = None
                        for worksheet in all_worksheets:
                            if worksheet.title.lower() == ws_name.lower():
                                ws = worksheet
                                break
 
                        if ws:
                            # Get existing records (bill_no, status) from Google Sheets
                            existing_records = get_existing_records(ws)
 
                            # Find duplicates
                            duplicate_records_in_branch = []
                            for _, row in branch_df.iterrows():
                                bill_no_clean = clean_bill_number(row["Bill No"])
                                status_clean = str(row["order status"]).strip()
                                if (bill_no_clean, status_clean) in existing_records:
                                    duplicate_records_in_branch.append(bill_no_clean)
 
                            if duplicate_records_in_branch:
                                if status not in duplicates_in_google_sheets:
                                    duplicates_in_google_sheets[status] = {}
                                if branch not in duplicates_in_google_sheets[status]:
                                    duplicates_in_google_sheets[status][branch] = []
                                duplicates_in_google_sheets[status][branch].extend(duplicate_records_in_branch)
                                total_duplicates_with_gs += len(duplicate_records_in_branch)
 
                except Exception as e:
                    print(f"Error checking Google Sheets for {status}: {e}")
                    continue
 
        # Remove duplicates with Google Sheets from the dataframe
        if total_duplicates_with_gs > 0:
            # Create a mask for rows that are NOT duplicates with Google Sheets
            mask = pd.Series([True] * len(df), index=df.index)
 
            for status, branches in duplicates_in_google_sheets.items():
                for branch, bill_nos in branches.items():
                    # Create condition for each duplicate row
                    condition = (
                        (df["order status"] == status) &
                        (df["Branch Name"] == branch) &
                        (df["Bill No"].isin(bill_nos))
                    )
                    mask = mask & ~condition
 
            df_filtered = df[mask].copy()
            gs_duplicates_removed = len(df) - len(df_filtered)
            df = df_filtered
        else:
            gs_duplicates_removed = 0
 
        # Save cleaned data
        df.to_csv(os.path.join(UPLOAD_FOLDER, 'temp_data.csv'), index=False)
 
        df_json = df.head(10).to_dict(orient='records')
        df_json = convert_numpy_to_python(df_json)
 
        duplicate_message = ""
        if internal_duplicates_removed > 0 or gs_duplicates_removed > 0:
            duplicate_message = f"Removed {internal_duplicates_removed} duplicate rows within file and {gs_duplicates_removed} rows already in Google Sheets"
 
        return jsonify({
            'success': True,
            'rows': len(df),
            'preview': df_json,
            'columns': list(df.columns),
            'duplicate_info': {
                'internal_duplicates_removed': int(internal_duplicates_removed),
                'google_sheets_duplicates_removed': int(gs_duplicates_removed),
                'total_duplicates_removed': int(internal_duplicates_removed + gs_duplicates_removed),
                'unique_rows_remaining': int(len(df)),
                'message': duplicate_message,
                'duplicates_by_status': convert_numpy_to_python(duplicates_in_google_sheets)
            }
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/process-data', methods=['POST'])
def process_data():
    """Process data and provide option to download ZIP files"""
    try:
        option = request.json.get('option')
        if not option:
            return jsonify({'error': 'No option specified'}), 400
 
        path = os.path.join(UPLOAD_FOLDER, 'temp_data.csv')
        if not os.path.exists(path):
            return jsonify({'error': 'Upload file first'}), 400

        df = pd.read_csv(path)
        df = clean_dataframe_for_json(df)
 
        if option == 'google_sheets':
            return update_google_sheets()
 
        elif option == 'summary_sheets':
            return update_summary_sheets()
 
        elif option == 'download_zip':
            zip_filename, zip_path, status_count = generate_zip_files(df)
 
            if not zip_filename:
                return jsonify({'error': 'Failed to generate ZIP files'}), 500
 
            status_summary = {}
            for status in STATUSES:
                status_df = df[df["order status"] == status]
                if not status_df.empty:
                    status_summary[status] = len(status_df)
 
            return jsonify({
                'success': True,
                'message': f'ZIP file generated with {status_count} status folders',
                'zip_filename': zip_filename,
                'status_count': status_count,
                'status_summary': status_summary,
                'total_records': len(df)
            })
 
        else:
            return jsonify({'error': 'Invalid option'}), 400

    except Exception as e:
        print(f"Error in process-data: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/update-google-sheets', methods=['POST'])
def update_google_sheets():
    try:
        path = os.path.join(UPLOAD_FOLDER, 'temp_data.csv')
        if not os.path.exists(path):
            return jsonify({'error': 'Upload file first'}), 400

        df = pd.read_csv(path)
        df = clean_dataframe_for_json(df)

        gc = get_google_sheets_client()
        if not gc:
            return jsonify({'error': 'Google authentication failed'}), 500

        today = datetime.now().strftime("%d-%m-%Y")
        current_time = datetime.now().strftime("%H:%M:%S")
        total_rows_updated = 0
        summary = {}
 
        batch_size = 5
        processed_count = 0
 
        # Track processed (bill_no, status) combinations in this run
        processed_combos_this_run = set()
 
        # Group data by status first
        for status in STATUSES:
            if status not in SHEET_IDS:
                continue

            spreadsheet = gc.open_by_key(SHEET_IDS[status])
            status_df = df[df["order status"] == status]
 
            if status_df.empty:
                print(f"⏭️  No data for status: {status}")
                continue
 
            print(f"📊 Processing status: {status}")
            all_worksheets = spreadsheet.worksheets()
            existing_worksheets = {}
            for ws in all_worksheets:
                existing_worksheets[ws.title.lower()] = ws
 
            branches_data = list(status_df.groupby("Branch Name"))
 
            for branch, branch_df in branches_data:
                processed_count += 1
                if processed_count % batch_size == 0:
                    print(f"⏳ Rate limiting: Waiting 15 seconds...")
                    time.sleep(15)
 
                ws_name = normalize_sheet_name(branch)
 
                # Get or create worksheet
                if ws_name.lower() in existing_worksheets:
                    ws = existing_worksheets[ws_name.lower()]
                    print(f"✅ Found worksheet: {ws.title}")
                else:
                    print(f"📄 Creating new worksheet: {ws_name}")
                    try:
                        ws = spreadsheet.add_worksheet(title=ws_name, rows="1000", cols="20")
                        print(f"✅ Created new worksheet: {ws_name}")
 
                        date_header = [f"Data Saved On: {today} {current_time}"] + [""] * 19
                        headers_list = ["S No", "Id", "Bill No", "Branch Name", "FinancialYearName",
                                      "Bill Date", "Total Bill Amount", "Total Discount Amount",
                                      "Total Tax Amount", "Net Amount", "Paid AT", "Bill Status",
                                      "Created By", "Created On", "order id", "tracking id",
                                      "bank ref no", "order status", "payment mode", "card name"]
 
                        ws.batch_update([
                            {'range': 'A1:T1', 'values': [date_header]},
                            {'range': 'A2:T2', 'values': [headers_list]}
                        ])
 
                        existing_worksheets[ws_name.lower()] = ws
                        print(f"✅ Initialized new worksheet with headers")
 
                    except Exception as e:
                        print(f"❌ Error creating worksheet: {e}")
                        continue
 
                # ============ GET ALL EXISTING RECORDS ============
                # Get ALL values from the worksheet for fresh read every time
                all_values = ws.get_all_values()
 
                # Extract all (bill_no, status) combinations
                existing_records = set()
 
                for row in all_values:
                    if len(row) > 17:  # Need at least up to column R
                        bill_no = row[2]  # Column C
                        status_val = row[17] if len(row) > 17 else ""  # Column R
 
                        if bill_no and str(bill_no).strip() and status_val and str(status_val).strip():
                            bill_no_clean = clean_bill_number(bill_no)
                            status_clean = str(status_val).strip()
 
                            # Skip headers
                            if (bill_no_clean and not bill_no_clean.startswith("BILLNO") and
                                not bill_no_clean.startswith("SNO") and
                                bill_no_clean != "" and bill_no_clean != "TOTAL" and
                                status_clean and status_clean != ""):
                                existing_records.add((bill_no_clean, status_clean))
 
                print(f"📊 Worksheet '{ws.title}' has {len(existing_records)} existing unique (bill_no, status) combinations")
                if len(existing_records) > 0:
                    print(f"📊 Sample: {list(existing_records)[:5]}")
 
                # ============ CLEAN UPLOADED DATA ============
                branch_df = branch_df.copy()
 
                # Clean Bill No column with improved function
                branch_df["Bill No"] = branch_df["Bill No"].apply(clean_bill_number)
                branch_df["order status"] = branch_df["order status"].apply(lambda x: str(x).strip())
 
                # Remove invalid records
                branch_df = branch_df[branch_df["Bill No"] != ""]
                branch_df = branch_df[branch_df["Bill No"] != "nan"]
                branch_df = branch_df[branch_df["order status"] != ""]
 
                print(f"🔍 Branch: {branch}, Total records in upload: {len(branch_df)}")
 
                # ============ REMOVE DUPLICATES WITHIN UPLOAD ============
                dup_in_upload = branch_df[branch_df.duplicated(subset=['Bill No', 'order status'], keep=False)]
                if not dup_in_upload.empty:
                    print(f"⚠️  Found {len(dup_in_upload)} duplicate (bill_no, status) combinations WITHIN uploaded data for {branch}")
                    branch_df = branch_df.drop_duplicates(subset=['Bill No', 'order status'], keep='first')
                    print(f"✅ After removing internal duplicates: {len(branch_df)} records")
 
                # ============ FILTER OUT ALREADY EXISTING RECORDS ============
                # Create a list of (bill_no, status) tuples for filtering
                branch_df['combo'] = list(zip(branch_df["Bill No"], branch_df["order status"]))
 
                # First, filter out records that already exist in the Google Sheet
                new_records_mask = ~branch_df['combo'].isin(existing_records)
                new_data = branch_df[new_records_mask].copy()
 
                # Also filter out records we've already processed in this same run
                if processed_combos_this_run:
                    not_processed_this_run_mask = ~new_data['combo'].isin(processed_combos_this_run)
                    new_data = new_data[not_processed_this_run_mask].copy()
 
                # Calculate stats
                total_in_upload = len(branch_df)
                duplicates_with_gs = len(branch_df[~new_records_mask])
                duplicates_this_run = len(set(branch_df['combo']) & processed_combos_this_run) if processed_combos_this_run else 0
                new_rows = len(new_data)
 
                print(f"🔍 Duplicate breakdown for {branch}:")
                print(f"   - Total in upload: {total_in_upload}")
                print(f"   - Already in Google Sheet: {duplicates_with_gs}")
                print(f"   - Already processed this run: {duplicates_this_run}")
                print(f"   - New rows to add: {new_rows}")
 
                if new_rows == 0:
                    print(f"⏭️  No new data for {branch}")
                    # Drop the combo column before continuing
                    branch_df = branch_df.drop('combo', axis=1, errors='ignore')
                    new_data = new_data.drop('combo', axis=1, errors='ignore')
                    continue
 
                # Drop the combo column before further processing
                branch_df = branch_df.drop('combo', axis=1, errors='ignore')
                new_data = new_data.drop('combo', axis=1, errors='ignore')
 
                # ============ FIND CORRECT SERIAL NUMBER ============
                # Find the highest serial number in the sheet
                max_serial = 0
                for row in all_values:
                    if row and len(row) > 0 and row[0] and str(row[0]).strip():
                        serial_str = str(row[0]).strip()
                        if serial_str.isdigit():
                            try:
                                serial = int(serial_str)
                                if serial > max_serial:
                                    max_serial = serial
                            except:
                                pass
 
                start_serial = max_serial + 1
                print(f"🔢 Starting serial number: {start_serial}")
 
                # ============ FIND NEXT EMPTY ROW ============
                # Find the first empty row at the bottom
                last_row_with_data = 0
                for i, row in enumerate(all_values, start=1):
                    if any(cell and str(cell).strip() for cell in row):
                        last_row_with_data = i
 
                append_row = last_row_with_data + 1
                print(f"📝 Appending at row: {append_row}")
 
                # ============ PREPARE DATA ============
                # Prepare headers list
                headers_list = ["S No", "Id", "Bill No", "Branch Name", "FinancialYearName",
                              "Bill Date", "Total Bill Amount", "Total Discount Amount",
                              "Total Tax Amount", "Net Amount", "Paid AT", "Bill Status",
                              "Created By", "Created On", "order id", "tracking id",
                              "bank ref no", "order status", "payment mode", "card name"]
 
                # Prepare data for insertion
                data_to_append = []
 
                # Add date header
                date_row = [f"Data Saved On: {today} {current_time}"] + [""] * 19
                data_to_append.append(date_row)
 
                # Add column headers
                data_to_append.append(headers_list)
 
                # Add data rows with correct serial numbers
                for idx, (_, row) in enumerate(new_data.iterrows()):
                    row_data = []
                    row_data.append(start_serial + idx)  # Serial No
                    row_data.append(row.get("Id", ""))
                    row_data.append(clean_bill_number(row.get("Bill No", "")))
                    row_data.append(row.get("Branch Name", ""))
                    row_data.append(row.get("FinancialYearName", ""))
 
                    # Bill Date
                    bill_date = row.get("Bill Date", "")
                    if pd.isna(bill_date):
                        bill_date = ""
                    elif isinstance(bill_date, (pd.Timestamp, datetime)):
                        bill_date = bill_date.strftime('%Y-%m-%d')
                    row_data.append(str(bill_date))
 
                    # Amounts
                    for col in ["Total Bill Amount", "Total Discount Amount", "Total Tax Amount", "Net Amount"]:
                        val = row.get(col, 0)
                        row_data.append(float(val) if not pd.isna(val) else 0)
 
                    # Remaining columns
                    for col in ["Paid AT", "Bill Status", "Created By", "Created On",
                              "order id", "tracking id", "bank ref no", "order status",
                              "payment mode", "card name"]:
                        val = row.get(col, "")
                        row_data.append(str(val) if not pd.isna(val) else "")
 
                    data_to_append.append(row_data)
 
                    # Add this combo to processed set
                    combo = (clean_bill_number(row.get("Bill No", "")), str(row.get("order status", "")).strip())
                    processed_combos_this_run.add(combo)
 
                # Add empty row
                data_to_append.append([""] * 20)
 
                # Add totals row
                if not new_data.empty:
                    totals = ["", "", "TOTAL", "", "", "",
                             float(new_data["Total Bill Amount"].sum()),
                             float(new_data["Total Discount Amount"].sum()),
                             float(new_data["Total Tax Amount"].sum()),
                             float(new_data["Net Amount"].sum())]
                    totals.extend([""] * 10)
                    data_to_append.append(totals)
 
                # Add 3 empty rows
                for _ in range(3):
                    data_to_append.append([""] * 20)
 
                # ============ UPDATE GOOGLE SHEET ============
                try:
                    # Calculate range
                    end_row = append_row + len(data_to_append) - 1
                    full_range = f"A{append_row}:T{end_row}"
 
                    print(f"📝 Writing {len(new_data)} rows to {full_range}")
 
                    # Update the sheet
                    ws.update(full_range, data_to_append, value_input_option='USER_ENTERED')
 
                    print(f"✅ Successfully added {len(new_data)} rows to {ws.title}")
 
                    # Update summary
                    rows_added = len(new_data)
                    total_rows_updated += rows_added
 
                    if status not in summary:
                        summary[status] = {}
                    if branch not in summary[status]:
                        summary[status][branch] = 0
                    summary[status][branch] += rows_added
 
                except Exception as e:
                    print(f"❌ Error updating sheet: {e}")
                    time.sleep(5)
                    try:
                        ws.update(full_range, data_to_append, value_input_option='USER_ENTERED')
                        print(f"✅ Retry successful for {ws.title}")
 
                        rows_added = len(new_data)
                        total_rows_updated += rows_added
 
                        if status not in summary:
                            summary[status] = {}
                        if branch not in summary[status]:
                            summary[status][branch] = 0
                        summary[status][branch] += rows_added
 
                    except Exception as retry_e:
                        print(f"❌ Retry failed: {retry_e}")
 
                time.sleep(2)  # Small delay between sheets
 
        # ============ PREPARE RESPONSE ============
        summary = convert_numpy_to_python(summary)
 
        response_message = f"Google Sheets updated successfully!\n"
        response_message += f"Total rows added: {total_rows_updated}\n"
        response_message += f"Date: {today} {current_time}\n\n"
 
        for status, branches in summary.items():
            if branches:
                response_message += f"{status}:\n"
                for branch, count in branches.items():
                    response_message += f"  {branch}: {count} rows\n"
 
        return jsonify({
            'success': True,
            'message': response_message,
            'rows_updated': int(total_rows_updated),
            'summary': summary,
            'date': today,
            'time': current_time
        })

    except Exception as e:
        print(f"Error in update-google-sheets: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/update-summary-sheets', methods=['POST'])
def update_summary_sheets():
    """Update only summary Google Sheets with specific columns (11 columns)"""
    try:
        path = os.path.join(UPLOAD_FOLDER, 'temp_data.csv')
        if not os.path.exists(path):
            return jsonify({'error': 'Upload file first'}), 400

        df = pd.read_csv(path)
        df = clean_dataframe_for_json(df)

        gc = get_google_sheets_client()
        if not gc:
            return jsonify({'error': 'Google authentication failed'}), 500

        today = datetime.now().strftime("%d-%m-%Y")
        current_time = datetime.now().strftime("%H:%M:%S")
        total_rows_updated = 0
        summary = {}
        
        batch_size = 5
        processed_count = 0
        
        # Track processed bill numbers in this run
        processed_bills_this_run = set()
        
        # Group data by status first
        for status in STATUSES:
            if status not in SUMMARY_SHEET_IDS:
                continue

            spreadsheet = gc.open_by_key(SUMMARY_SHEET_IDS[status])
            status_df = df[df["order status"] == status]
            
            if status_df.empty:
                print(f"⏭️  No data for status: {status}")
                continue
            
            print(f"\n📊 Processing status: {status} (SUMMARY SHEET ONLY)")
            
            all_worksheets = spreadsheet.worksheets()
            existing_worksheets = {ws.title.lower(): ws for ws in all_worksheets}
            
            branches_data = list(status_df.groupby("Branch Name"))
            
            for branch, branch_df in branches_data:
                processed_count += 1
                if processed_count % batch_size == 0:
                    print(f"⏳ Rate limiting: Waiting 15 seconds...")
                    time.sleep(15)
                
                # Clean branch data
                branch_df = branch_df.copy()
                branch_df["Bill No"] = branch_df["Bill No"].apply(clean_bill_number)
                branch_df["order status"] = branch_df["order status"].apply(lambda x: str(x).strip())
                branch_df = branch_df[branch_df["Bill No"] != ""]
                branch_df = branch_df[branch_df["Bill No"] != "nan"]
                branch_df = branch_df[branch_df["order status"] != ""]
                
                # Remove duplicates within branch data
                dup_in_upload = branch_df[branch_df.duplicated(subset=['Bill No'], keep=False)]
                if not dup_in_upload.empty:
                    branch_df = branch_df.drop_duplicates(subset=['Bill No'], keep='first')
                
                if branch_df.empty:
                    continue
                
                print(f"🔍 Processing branch: {branch}")
                
                # Worksheet name for summary
                ws_name = f"{normalize_sheet_name(branch)}_summary"
                
                # Get or create worksheet
                if ws_name.lower() in existing_worksheets:
                    ws = existing_worksheets[ws_name.lower()]
                    print(f"   ✅ Found SUMMARY worksheet: {ws.title}")
                else:
                    print(f"   📄 Creating new SUMMARY worksheet: {ws_name}")
                    try:
                        ws = spreadsheet.add_worksheet(title=ws_name, rows="1000", cols="11")
                        date_header_summary = [f"Data Saved On: {today} {current_time}"] + [""] * 10
                        headers_list_summary = [
                            "S No", "Bill No", "Branch Name", "Bill Date",
                            "Total Bill Amount", "Total Discount Amount", "Total Tax Amount", "Net Amount",
                            "Paid AT", "Created By", "order id"
                        ]
                        ws.batch_update([
                            {'range': 'A1:K1', 'values': [date_header_summary]},
                            {'range': 'A2:K2', 'values': [headers_list_summary]}
                        ])
                        existing_worksheets[ws_name.lower()] = ws
                        print(f"   ✅ Created and initialized SUMMARY worksheet with 11 columns")
                    except Exception as e:
                        print(f"   ❌ Error creating worksheet: {e}")
                        continue
                
                # Get existing bills from SUMMARY sheet
                existing_bills = get_existing_bills_from_summary(ws)
                print(f"   📊 SUMMARY sheet has {len(existing_bills)} existing bill numbers")
                
                # Filter out already existing bills
                new_records_mask = ~branch_df["Bill No"].isin(existing_bills)
                
                # Also filter out records already processed in this run
                if processed_bills_this_run:
                    not_processed_mask = ~branch_df["Bill No"].isin(processed_bills_this_run)
                    new_records_mask = new_records_mask & not_processed_mask
                
                new_data = branch_df[new_records_mask].copy()
                
                print(f"   📈 New records to add: {len(new_data)} rows")
                
                if new_data.empty:
                    print(f"   ⏭️  No new data for {branch}")
                    continue
                
                # Find serial number
                all_values = ws.get_all_values()
                max_serial = 0
                for row in all_values:
                    if row and len(row) > 0 and row[0] and str(row[0]).strip():
                        serial_str = str(row[0]).strip()
                        if serial_str.isdigit():
                            max_serial = max(max_serial, int(serial_str))
                
                start_serial = max_serial + 1
                
                # Find next empty row
                last_row_with_data = 0
                for i, row in enumerate(all_values, start=1):
                    if any(cell and str(cell).strip() for cell in row):
                        last_row_with_data = i
                
                append_row = last_row_with_data + 1
                
                # Prepare data for SUMMARY sheet
                data_to_append = []
                data_to_append.append([f"Data Saved On: {today} {current_time}"] + [""] * 10)
                data_to_append.append([
                    "S No", "Bill No", "Branch Name", "Bill Date",
                    "Total Bill Amount", "Total Discount Amount", "Total Tax Amount", "Net Amount",
                    "Paid AT", "Created By", "order id"
                ])
                
                for idx, (_, row) in enumerate(new_data.iterrows()):
                    row_data = []
                    row_data.append(start_serial + idx)
                    row_data.append(clean_bill_number(row.get("Bill No", "")))
                    row_data.append(row.get("Branch Name", ""))
                    
                    # Bill Date
                    bill_date = row.get("Bill Date", "")
                    if pd.isna(bill_date):
                        bill_date = ""
                    elif isinstance(bill_date, (pd.Timestamp, datetime)):
                        bill_date = bill_date.strftime('%Y-%m-%d')
                    row_data.append(str(bill_date))
                    
                    # Amounts
                    total_bill = row.get("Total Bill Amount", 0)
                    row_data.append(float(total_bill) if not pd.isna(total_bill) else 0)
                    
                    discount = row.get("Total Discount Amount", 0)
                    row_data.append(float(discount) if not pd.isna(discount) else 0)
                    
                    tax = row.get("Total Tax Amount", 0)
                    row_data.append(float(tax) if not pd.isna(tax) else 0)
                    
                    net_amount = row.get("Net Amount", 0)
                    row_data.append(float(net_amount) if not pd.isna(net_amount) else 0)
                    
                    # Other fields
                    paid_at = row.get("Paid AT", "")
                    row_data.append(str(paid_at) if not pd.isna(paid_at) else "")
                    
                    created_by = row.get("Created By", "")
                    row_data.append(str(created_by) if not pd.isna(created_by) else "")
                    
                    order_id = row.get("order id", "")
                    row_data.append(str(order_id) if not pd.isna(order_id) else "")
                    
                    data_to_append.append(row_data)
                    
                    # Add bill to processed set
                    processed_bills_this_run.add(clean_bill_number(row.get("Bill No", "")))
                
                data_to_append.append([""] * 11)
                
                if not new_data.empty:
                    totals = [
                        "",
                        "TOTAL",
                        "",
                        "",
                        float(new_data["Total Bill Amount"].sum()),
                        float(new_data["Total Discount Amount"].sum()),
                        float(new_data["Total Tax Amount"].sum()),
                        float(new_data["Net Amount"].sum()),
                        "",
                        "",
                        ""
                    ]
                    data_to_append.append(totals)
                
                for _ in range(3):
                    data_to_append.append([""] * 11)
                
                # Update SUMMARY sheet
                try:
                    end_row = append_row + len(data_to_append) - 1
                    ws.update(f"A{append_row}:K{end_row}", data_to_append, value_input_option='USER_ENTERED')
                    print(f"   ✅ SUMMARY sheet updated with {len(new_data)} rows")
                    
                    total_rows_updated += len(new_data)
                    
                    if status not in summary:
                        summary[status] = {}
                    if branch not in summary[status]:
                        summary[status][branch] = 0
                    summary[status][branch] += len(new_data)
                    
                except Exception as e:
                    print(f"   ❌ Error updating SUMMARY sheet: {e}")
                    time.sleep(5)
                    try:
                        ws.update(f"A{append_row}:K{end_row}", data_to_append, value_input_option='USER_ENTERED')
                        print(f"   ✅ Retry successful for {ws.title}")
                        
                        total_rows_updated += len(new_data)
                        
                        if status not in summary:
                            summary[status] = {}
                        if branch not in summary[status]:
                            summary[status][branch] = 0
                        summary[status][branch] += len(new_data)
                        
                    except Exception as retry_e:
                        print(f"   ❌ Retry failed: {retry_e}")
                
                time.sleep(2)  # Small delay between branches
        
        # Prepare response
        summary = convert_numpy_to_python(summary)
        
        response_message = f"✅ SUMMARY SHEETS UPDATE COMPLETED!\n\n"
        response_message += f"📋 SUMMARY SHEETS (11 columns):\n"
        response_message += f"   Columns: S No, Bill No, Branch Name, Bill Date, Total Bill Amount, Total Discount Amount, Total Tax Amount, Net Amount, Paid AT, Created By, order id\n"
        response_message += f"   Total rows added: {total_rows_updated}\n\n"
        response_message += f"📅 Date: {today} {current_time}\n\n"
        
        response_message += "📈 SUMMARY SHEETS BREAKDOWN:\n"
        for status, branches in summary.items():
            if branches:
                response_message += f"   {status}:\n"
                for branch, count in branches.items():
                    response_message += f"      {branch}: {count} rows\n"
        
        return jsonify({
            'success': True,
            'message': response_message,
            'rows_updated': int(total_rows_updated),
            'summary': summary,
            'date': today,
            'time': current_time
        })
        
    except Exception as e:
        print(f"Error in update-summary-sheets: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download-zip/<filename>')
def download_zip(filename):
    """Download generated ZIP file"""
    try:
        return send_from_directory(
            app.config['ZIP_FOLDER'],
            filename,
            as_attachment=True,
            mimetype='application/zip'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 404

@app.route('/check-google-sheets')
def check_google_sheets():
    gc = get_google_sheets_client()
    if not gc:
        return jsonify({'accessible': False})

    sheet = gc.open_by_key(SHEET_IDS["Success"])
    return jsonify({
        'accessible': True,
        'worksheets': [ws.title for ws in sheet.worksheets()]
    })

# ===============================
# CHECK DUPLICATES ROUTE
# ===============================
@app.route('/check-duplicates', methods=['POST'])
def check_duplicates():
    """Check for duplicate Bill Nos in uploaded data"""
    try:
        path = os.path.join(UPLOAD_FOLDER, 'temp_data.csv')
        if not os.path.exists(path):
            return jsonify({'error': 'Upload file first'}), 400
 
        df = pd.read_csv(path)
        df = clean_dataframe_for_json(df)
 
        # Clean Bill No column with improved function
        df["Bill No"] = df["Bill No"].apply(clean_bill_number)
        df["order status"] = df["order status"].apply(lambda x: str(x).strip())
        df = df[df["Bill No"] != ""]
        df = df[df["Bill No"] != "nan"]
        df = df[df["order status"] != ""]
 
        # Check for duplicates based on Bill No + Status
        duplicate_df = df[df.duplicated(subset=['Bill No', 'order status'], keep=False)]
 
        if duplicate_df.empty:
            return jsonify({
                'has_duplicates': False,
                'message': 'No duplicate (Bill No + Status) combinations found in uploaded data'
            })
 
        # Group duplicates
        duplicates_summary = []
        for (bill_no, status), group in duplicate_df.groupby(['Bill No', 'order status']):
            duplicates_summary.append({
                'bill_no': str(bill_no),
                'status': str(status),
                'count': int(len(group)),
                'branches': group['Branch Name'].unique().tolist()
            })
 
        # Sort by count descending
        duplicates_summary.sort(key=lambda x: x['count'], reverse=True)
 
        return jsonify({
            'has_duplicates': True,
            'total_duplicate_combinations': len(duplicate_df.groupby(['Bill No', 'order status']).size()),
            'total_duplicate_rows': len(duplicate_df),
            'duplicates': duplicates_summary[:20],
            'message': f'Found {len(duplicate_df)} duplicate rows across {len(duplicate_df.groupby(["Bill No", "order status"]))} unique combinations'
        })
 
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ===============================
# DEBUG ROUTE FOR DUPLICATES
# ===============================
@app.route('/debug-worksheet/<status>/<branch>')
def debug_worksheet(status, branch):
    """Debug endpoint to check worksheet data for duplicates"""
    try:
        gc = get_google_sheets_client()
        if not gc:
            return jsonify({'error': 'Google auth failed'}), 500
 
        if status not in SHEET_IDS:
            return jsonify({'error': 'Invalid status'}), 400
 
        spreadsheet = gc.open_by_key(SHEET_IDS[status])
        ws_name = normalize_sheet_name(branch)
 
        ws = None
        for worksheet in spreadsheet.worksheets():
            if worksheet.title.lower() == ws_name.lower():
                ws = worksheet
                break
 
        if not ws:
            return jsonify({'error': f'Worksheet {branch} not found'}), 404
 
        all_values = ws.get_all_values()
 
        records = []
        for row in all_values:
            if len(row) > 17:
                bill_no = clean_bill_number(row[2]) if len(row) > 2 else ""
                status_val = row[17] if len(row) > 17 else ""
                if bill_no and status_val and not bill_no.startswith("BILLNO") and not bill_no.startswith("SNO") and bill_no != "TOTAL":
                    records.append((bill_no, status_val))
 
        # Find duplicates
        seen = set()
        duplicates = []
        for record in records:
            if record in seen:
                duplicates.append(record)
            else:
                seen.add(record)
 
        return jsonify({
            'worksheet': ws.title,
            'status': status,
            'total_rows': len(all_values),
            'total_records': len(records),
            'unique_records': len(seen),
            'duplicates_found': len(duplicates),
            'duplicate_list': duplicates[:20],
            'sample_records': list(seen)[:10]
        })
 
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ===============================
# COMPARE BILLS ROUTE
# ===============================
@app.route('/compare-bills', methods=['POST'])
def compare_bills():
    """Compare bill numbers between uploaded data and Google Sheets"""
    try:
        data = request.json
        status = data.get('status')
        branch = data.get('branch')
 
        if not status or not branch:
            return jsonify({'error': 'Status and branch required'}), 400
 
        path = os.path.join(UPLOAD_FOLDER, 'temp_data.csv')
        if not os.path.exists(path):
            return jsonify({'error': 'Upload file first'}), 400
 
        df = pd.read_csv(path)
        df = clean_dataframe_for_json(df)
 
        gc = get_google_sheets_client()
        if not gc:
            return jsonify({'error': 'Google auth failed'}), 500
 
        if status not in SHEET_IDS:
            return jsonify({'error': 'Invalid status'}), 400
 
        spreadsheet = gc.open_by_key(SHEET_IDS[status])
        ws_name = normalize_sheet_name(branch)
 
        ws = None
        for worksheet in spreadsheet.worksheets():
            if worksheet.title.lower() == ws_name.lower():
                ws = worksheet
                break
 
        if not ws:
            return jsonify({'error': f'Worksheet {branch} not found'}), 404
 
        # Get existing records from Google Sheets
        existing_records = get_existing_records(ws)
 
        branch_df = df[(df["order status"] == status) & (df["Branch Name"] == branch)]
        branch_df = branch_df.copy()
        branch_df["Bill No"] = branch_df["Bill No"].apply(clean_bill_number)
        branch_df["order status"] = branch_df["order status"].apply(lambda x: str(x).strip())
        branch_df = branch_df[branch_df["Bill No"] != ""]
        branch_df = branch_df[branch_df["Bill No"] != "nan"]
        branch_df = branch_df[branch_df["order status"] != ""]
 
        new_records = []
        duplicate_records = []
 
        for _, row in branch_df.iterrows():
            bill_no = clean_bill_number(row["Bill No"])
            status_val = str(row["order status"]).strip()
            if (bill_no, status_val) in existing_records:
                duplicate_records.append((bill_no, status_val))
            else:
                new_records.append((bill_no, status_val))
 
        return jsonify({
            'status': status,
            'branch': branch,
            'total_in_upload': len(branch_df),
            'total_in_sheet': len(existing_records),
            'new_records_count': len(new_records),
            'duplicate_records_count': len(duplicate_records),
            'new_records_sample': new_records[:10],
            'duplicate_records_sample': duplicate_records[:10]
        })
 
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ===============================
# CLEANUP ROUTINE
# ===============================
@app.route('/cleanup', methods=['POST'])
def cleanup():
    """Clean up temporary files"""
    try:
        for file in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
 
        for file in os.listdir(ZIP_FOLDER):
            file_path = os.path.join(ZIP_FOLDER, file)
            if os.path.isfile(file_path):
                file_age = time.time() - os.path.getmtime(file_path)
                if file_age > 3600:
                    os.remove(file_path)
 
        return jsonify({'success': True, 'message': 'Cleanup completed'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
    