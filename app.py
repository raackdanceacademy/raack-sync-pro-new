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
    "MADURAVOYAL", "KANDIGAI", "KUNDRATHUR"
]

STATUSES = [
    "Success", "Failure", "Initiated", "Awaited",
    "Timeout", "Unsuccessful", "Aborted"
]

# ===============================
# GOOGLE SHEET IDS
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
def find_empty_row_for_append(worksheet):
    """Find the first empty row to append data - OPTIMIZED VERSION"""
    try:
        all_values = worksheet.get_all_values()
        
        last_row = 0
        for i, row in enumerate(all_values, start=1):
            if any(cell.strip() for cell in row):
                last_row = i
        
        return last_row + 1
    except Exception as e:
        print(f"Error finding empty row: {e}")
        return 2

def get_existing_bill_nos(worksheet):
    """Get all existing bill numbers from the worksheet - OPTIMIZED"""
    bill_nos = set()
    try:
        all_values = worksheet.get_all_values()
        
        for row in all_values:
            if len(row) > 2:
                bill_no = row[2]
                # Clean and check the bill number
                if bill_no:
                    bill_no = str(bill_no).strip()
                    # Skip headers and empty values
                    if bill_no and not bill_no.startswith("Bill No"):
                        bill_nos.add(bill_no)
        return bill_nos
    except Exception as e:
        print(f"Error getting existing bill numbers: {e}")
        return set()

def validate_bill_no_uniqueness(df):
    """Validate that Bill No is unique in the uploaded data itself"""
    try:
        # Clean Bill No column
        df = df.copy()
        df["Bill No"] = df["Bill No"].astype(str).str.strip()
        
        # Remove empty bill numbers
        df = df[df["Bill No"] != ""]
        
        # Check for duplicates in uploaded data
        duplicates_in_upload = df[df.duplicated(subset=['Bill No'], keep=False)]
        
        if not duplicates_in_upload.empty:
            print(f"⚠️  Found {len(duplicates_in_upload)} duplicate Bill Nos in uploaded data")
            duplicate_counts = duplicates_in_upload['Bill No'].value_counts()
            print(f"🔍 Most common duplicates: {duplicate_counts.head(5).to_dict()}")
            
            # Remove duplicates within uploaded data (keep first occurrence)
            df_unique = df.drop_duplicates(subset=['Bill No'], keep='first')
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

def prepare_data_for_sheet(new_data, today, current_time, start_serial=1):
    """Prepare data in the correct format for Google Sheets"""
    data_to_append = []
    
    # Add date separator
    date_row = [f"Data Saved On: {today} {current_time}"]
    date_row.extend([""] * 19)
    data_to_append.append(date_row)
    
    # Add column headers
    headers_list = ["S No", "Id", "Bill No", "Branch Name", "FinancialYearName", 
                   "Bill Date", "Total Bill Amount", "Total Discount Amount", 
                   "Total Tax Amount", "Net Amount", "Paid AT", "Bill Status", 
                   "Created By", "Created On", "order id", "tracking id", 
                   "bank ref no", "order status", "payment mode", "card name"]
    data_to_append.append(headers_list)
    
    # Add the actual data rows with proper serial numbers
    for idx, row in new_data.iterrows():
        row_data = []
        serial_no = start_serial + idx
        
        # Add serial number
        row_data.append(serial_no)
        
        # Add ID (if exists in data)
        row_data.append(row.get("Id", ""))
        
        # Add Bill No
        row_data.append(str(row.get("Bill No", "")))
        
        # Add Branch Name
        row_data.append(row.get("Branch Name", ""))
        
        # Add FinancialYearName (if exists)
        row_data.append(row.get("FinancialYearName", ""))
        
        # Add Bill Date
        bill_date = row.get("Bill Date", "")
        if pd.isna(bill_date):
            bill_date = ""
        elif isinstance(bill_date, (pd.Timestamp, datetime)):
            bill_date = bill_date.strftime('%Y-%m-%d')
        row_data.append(str(bill_date))
        
        # Add Total Bill Amount
        total_bill = row.get("Total Bill Amount", 0)
        row_data.append(float(total_bill) if not pd.isna(total_bill) else 0)
        
        # Add Total Discount Amount
        discount = row.get("Total Discount Amount", 0)
        row_data.append(float(discount) if not pd.isna(discount) else 0)
        
        # Add Total Tax Amount
        tax = row.get("Total Tax Amount", 0)
        row_data.append(float(tax) if not pd.isna(tax) else 0)
        
        # Add Net Amount
        net_amount = row.get("Net Amount", 0)
        row_data.append(float(net_amount) if not pd.isna(net_amount) else 0)
        
        # Add remaining columns
        for col in ["Paid AT", "Bill Status", "Created By", "Created On", 
                   "order id", "tracking id", "bank ref no", "order status", 
                   "payment mode", "card name"]:
            value = row.get(col, "")
            if pd.isna(value):
                value = ""
            row_data.append(str(value))
        
        data_to_append.append(row_data)
    
    # Add empty row
    data_to_append.append([""] * 20)
    
    # Add totals row
    totals = ["", "", "TOTAL", "", "", "",
             float(new_data["Total Bill Amount"].sum()) if not new_data.empty else 0,
             float(new_data["Total Discount Amount"].sum()) if not new_data.empty else 0,
             float(new_data["Total Tax Amount"].sum()) if not new_data.empty else 0,
             float(new_data["Net Amount"].sum()) if not new_data.empty else 0]
    totals.extend([""] * 10)
    data_to_append.append(totals)
    
    # Add 3 more empty rows for separation
    for _ in range(3):
        data_to_append.append([""] * 20)
    
    return data_to_append

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
        
        # Clean Bill No column
        df["Bill No"] = df["Bill No"].astype(str).str.strip()
        df = df[df["Bill No"] != ""]
        
        # Check for duplicates WITHIN the uploaded file
        df, internal_duplicates_removed = validate_bill_no_uniqueness(df)
        
        # Now check for duplicates AGAINST Google Sheets
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
                            # Get existing bill numbers from Google Sheets
                            existing_bill_nos = get_existing_bill_nos(ws)
                            existing_bill_nos_clean = {str(bill).strip() for bill in existing_bill_nos if str(bill).strip()}
                            
                            # Find duplicates
                            duplicate_bills_in_branch = []
                            for bill_no in branch_df["Bill No"]:
                                if str(bill_no).strip() in existing_bill_nos_clean:
                                    duplicate_bills_in_branch.append(str(bill_no))
                            
                            if duplicate_bills_in_branch:
                                if status not in duplicates_in_google_sheets:
                                    duplicates_in_google_sheets[status] = {}
                                if branch not in duplicates_in_google_sheets[status]:
                                    duplicates_in_google_sheets[status][branch] = []
                                duplicates_in_google_sheets[status][branch].extend(duplicate_bills_in_branch)
                                total_duplicates_with_gs += len(duplicate_bills_in_branch)
                
                except Exception as e:
                    print(f"Error checking Google Sheets for {status}: {e}")
                    continue
        
        # Remove duplicates with Google Sheets from the dataframe
        if total_duplicates_with_gs > 0:
            # Create a mask for rows that are NOT duplicates with Google Sheets
            mask = pd.Series([True] * len(df), index=df.index)
            
            for status, branches in duplicates_in_google_sheets.items():
                for branch, bill_nos in branches.items():
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
            # This will now only process the already-filtered data
            return update_google_sheets()
            
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
        
        worksheet_cache = {}
        
        # Group data by status first
        for status in STATUSES:
            if status not in SHEET_IDS:
                continue

            spreadsheet = gc.open_by_key(SHEET_IDS[status])
            status_df = df[df["order status"] == status]
            
            if status_df.empty:
                print(f"⏭️  No data for status: {status}")
                continue
            
            print(f"📊 Getting worksheets for {status}...")
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
                
                if ws_name.lower() in existing_worksheets:
                    ws = existing_worksheets[ws_name.lower()]
                    print(f"✅ Found existing worksheet: {ws.title}")
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
                        
                        ws.batch_update([{
                            'range': 'A1:T1',
                            'values': [date_header]
                        }, {
                            'range': 'A2:T2',
                            'values': [headers_list]
                        }])
                        
                        existing_worksheets[ws_name.lower()] = ws
                        worksheet_cache[ws_name.lower()] = {
                            'data': [],
                            'last_updated': datetime.now()
                        }
                        print(f"💾 Initialized empty cache for new worksheet: {ws_name}")
                        
                    except Exception as e:
                        error_msg = str(e)
                        if "already exists" in error_msg.lower():
                            print(f"⚠️  Worksheet '{ws_name}' exists. Trying to find it...")
                            all_worksheets = spreadsheet.worksheets()
                            existing_worksheets = {}
                            for ws_obj in all_worksheets:
                                existing_worksheets[ws_obj.title.lower()] = ws_obj
                            
                            if ws_name.lower() in existing_worksheets:
                                ws = existing_worksheets[ws_name.lower()]
                                print(f"✅ Now found worksheet: {ws.title}")
                            else:
                                print(f"❌ Worksheet '{ws_name}' not found after refresh")
                                continue
                        else:
                            print(f"❌ Error creating worksheet: {e}")
                            continue
                
                cache_key = ws_name.lower()
                
                if cache_key in worksheet_cache:
                    cache_age = (datetime.now() - worksheet_cache[cache_key]['last_updated']).total_seconds()
                    if cache_age < 300:
                        existing_bill_nos = set(worksheet_cache[cache_key]['data'])
                        print(f"📦 Using cached bill numbers for {ws.title} ({len(existing_bill_nos)} bills)")
                    else:
                        existing_bill_nos = get_existing_bill_nos(ws)
                        worksheet_cache[cache_key] = {
                            'data': list(existing_bill_nos),
                            'last_updated': datetime.now()
                        }
                        print(f"🔄 Refreshed cache for {ws.title} ({len(existing_bill_nos)} bills)")
                else:
                    existing_bill_nos = get_existing_bill_nos(ws)
                    worksheet_cache[cache_key] = {
                        'data': list(existing_bill_nos),
                        'last_updated': datetime.now()
                    }
                    print(f"📊 Fetched fresh bill numbers for {ws.title} ({len(existing_bill_nos)} bills)")
                
                print(f"🔍 Worksheet '{ws.title}' has {len(existing_bill_nos)} existing bill numbers")
                
                # Clean bill numbers properly
                existing_bill_nos_clean = {str(bill).strip() for bill in existing_bill_nos if str(bill).strip()}
                
                # Clean branch_df bill numbers
                branch_df = branch_df.copy()
                branch_df["Bill No"] = branch_df["Bill No"].astype(str).str.strip()
                
                # Remove any empty bill numbers
                branch_df = branch_df[branch_df["Bill No"] != ""]
                
                sample_new_bills = branch_df["Bill No"].head(5).tolist()
                sample_existing_bills = list(existing_bill_nos_clean)[:5] if existing_bill_nos_clean else []
                print(f"🔍 Sample new bills: {sample_new_bills}")
                print(f"🔍 Sample existing bills: {sample_existing_bills}")
                
                # Filter for new bills only
                new_data = branch_df[~branch_df["Bill No"].isin(existing_bill_nos_clean)]
                
                total_rows = len(branch_df)
                new_rows = len(new_data)
                duplicate_rows = total_rows - new_rows
                
                # Detailed logging for duplicate check
                print(f"🔍 Duplicate check details:")
                print(f"   - Total rows in branch data: {total_rows}")
                print(f"   - Existing bill numbers in sheet: {len(existing_bill_nos_clean)}")
                print(f"   - New rows to add: {new_rows}")
                print(f"   - Duplicates filtered out: {duplicate_rows}")
                
                # Log specific duplicates if any
                if duplicate_rows > 0:
                    duplicate_bills = set(branch_df["Bill No"]) - set(new_data["Bill No"])
                    if duplicate_bills:
                        print(f"   - Example duplicate bill numbers: {list(duplicate_bills)[:5]}")
                
                if len(new_data) == 0:
                    print(f"⏭️  No new data for {branch} ({status}) - all {duplicate_rows} rows were duplicates")
                    continue
                
                append_row = find_empty_row_for_append(ws)
                
                start_serial = 1
                if append_row > 2:
                    try:
                        existing_values = ws.get_all_values()
                        
                        last_serial = 0
                        for row in existing_values:
                            if row and row[0] and row[0].isdigit():
                                try:
                                    serial = int(row[0])
                                    if serial > last_serial:
                                        last_serial = serial
                                except:
                                    continue
                        
                        start_serial = last_serial + 1
                        print(f"🔢 Starting serial number: {start_serial}")
                    except:
                        start_serial = 1
                
                data_to_append = prepare_data_for_sheet(new_data, today, current_time, start_serial)
                
                if data_to_append:
                    try:
                        start_range = f"A{append_row}"
                        end_row = append_row + len(data_to_append) - 1
                        end_range = f"T{end_row}"
                        full_range = f"{start_range}:{end_range}"
                        
                        print(f"📝 Updating range {full_range} with {len(new_data)} new rows")
                        
                        ws.update(full_range, data_to_append)
                        
                        new_bill_nos = new_data["Bill No"].astype(str).str.strip().tolist()
                        
                        if cache_key in worksheet_cache:
                            current_cache = worksheet_cache[cache_key]['data']
                            updated_cache_data = list(set(current_cache + new_bill_nos))
                            worksheet_cache[cache_key] = {
                                'data': updated_cache_data,
                                'last_updated': datetime.now()
                            }
                            print(f"💾 Updated cache with {len(new_bill_nos)} new bill numbers")
                            print(f"💾 Cache now has {len(updated_cache_data)} total bill numbers")
                        else:
                            worksheet_cache[cache_key] = {
                                'data': new_bill_nos,
                                'last_updated': datetime.now()
                            }
                        
                        print(f"✅ Successfully added {len(new_data)} rows to {ws.title} ({status})")
                        
                    except Exception as e:
                        print(f"❌ Error updating sheet {ws.title}: {e}")
                        time.sleep(30)
                        try:
                            ws.update(full_range, data_to_append)
                            print(f"✅ Retry successful for {ws.title}")
                            
                            new_bill_nos = new_data["Bill No"].astype(str).str.strip().tolist()
                            if cache_key in worksheet_cache:
                                current_cache = worksheet_cache[cache_key]['data']
                                updated_cache_data = list(set(current_cache + new_bill_nos))
                                worksheet_cache[cache_key] = {
                                    'data': updated_cache_data,
                                    'last_updated': datetime.now()
                                }
                        except:
                            print(f"❌ Retry failed for {ws.title}")
                            continue
                
                rows_added = len(new_data)
                total_rows_updated += rows_added
                
                if status not in summary:
                    summary[status] = {}
                if branch not in summary[status]:
                    summary[status][branch] = 0
                summary[status][branch] += rows_added
                
                time.sleep(2)
        
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
        
        # Clean Bill No column
        df["Bill No"] = df["Bill No"].astype(str).str.strip()
        df = df[df["Bill No"] != ""]
        
        # Check for duplicates
        duplicate_df = df[df.duplicated(subset=['Bill No'], keep=False)]
        
        if duplicate_df.empty:
            return jsonify({
                'has_duplicates': False,
                'message': 'No duplicate Bill Nos found in uploaded data'
            })
        
        # Group duplicates
        duplicates_summary = []
        for bill_no, group in duplicate_df.groupby('Bill No'):
            duplicates_summary.append({
                'bill_no': str(bill_no),
                'count': int(len(group)),
                'branches': group['Branch Name'].unique().tolist(),
                'statuses': group['order status'].unique().tolist()
            })
        
        # Sort by count descending
        duplicates_summary.sort(key=lambda x: x['count'], reverse=True)
        
        return jsonify({
            'has_duplicates': True,
            'total_duplicate_bills': len(duplicate_df['Bill No'].unique()),
            'total_duplicate_rows': len(duplicate_df),
            'duplicates': duplicates_summary[:20],  # Top 20 duplicates
            'message': f'Found {len(duplicate_df)} duplicate rows across {len(duplicate_df["Bill No"].unique())} bill numbers'
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
        
        bill_numbers = []
        for row in all_values:
            if len(row) > 2:
                bill_no = row[2]
                if bill_no and bill_no.strip() and not bill_no.startswith("Bill No"):
                    bill_numbers.append(bill_no.strip())
        
        duplicates = []
        seen = set()
        for bill in bill_numbers:
            if bill in seen:
                duplicates.append(bill)
            else:
                seen.add(bill)
        
        return jsonify({
            'worksheet': ws.title,
            'status': status,
            'total_rows': len(all_values),
            'total_bill_numbers': len(bill_numbers),
            'unique_bill_numbers': len(seen),
            'duplicates_found': len(duplicates),
            'duplicate_list': duplicates[:20],
            'sample_bill_numbers': list(seen)[:10]
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
        
        existing_bill_nos = get_existing_bill_nos(ws)
        
        branch_df = df[(df["order status"] == status) & (df["Branch Name"] == branch)]
        branch_df = branch_df.copy()
        branch_df["Bill No"] = branch_df["Bill No"].astype(str).str.strip()
        branch_df = branch_df[branch_df["Bill No"] != ""]
        
        existing_bill_nos_clean = {str(bill).strip() for bill in existing_bill_nos if str(bill).strip()}
        
        new_bills = []
        duplicate_bills = []
        
        for idx, row in branch_df.iterrows():
            bill_no = str(row["Bill No"]).strip()
            if bill_no in existing_bill_nos_clean:
                duplicate_bills.append(bill_no)
            else:
                new_bills.append(bill_no)
        
        return jsonify({
            'status': status,
            'branch': branch,
            'total_in_upload': len(branch_df),
            'total_in_sheet': len(existing_bill_nos_clean),
            'new_bills_count': len(new_bills),
            'duplicate_bills_count': len(duplicate_bills),
            'new_bills_sample': new_bills[:10],
            'duplicate_bills_sample': duplicate_bills[:10]
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