import os
from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env file

class Config:
    SECRET_KEY = os.getenv('SECRET_KEY', 'your-secret-key-change-this-in-production')
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    
    # Google Sheets Configuration
    GOOGLE_SHEETS_CREDENTIALS = os.getenv('GOOGLE_SHEETS_CREDENTIALS', 'credentials.json')
    
    # Email for Google Sheets sharing (optional)
    GOOGLE_SHEETS_SHARE_EMAIL = os.getenv('GOOGLE_SHEETS_SHARE_EMAIL', '')
    
    # Branches and Statuses
    BRANCHES = [
        "KILPAUK", "MYLAPORE", "VELACHERY", "CUDDALORE", "TAMBARAM", "MOGAPPAIR",
        "THORAIPAKKAM", "AVADI", "KEELKATTALAI", "MUGALIVAKKAM", "SHOLINGANALLUR",
        "NEELANKARAI", "KOLATHUR", "PALLIKARANAI", "OLD PERUNGALATHUR",
        "GUDUVANCHERI", "PUDUCHERRY", "RAMAPURAM", "SAIDAPET", "OLD PALLAVARAM",
        "MANNIVAKKAM", "CHIDAMBARAM", "HASTHINAPURAM", "THIRUVERKADU", "SURAPET",
        "MARAIMALAI NAGAR", "PADUR", "MEDAVAKKAM", "PADAPPAI", "AMBATTUR",
        "ARUMBAKKAM", "AYAPAKKAM", "SITHALAPAKKAM", "PERUMBAKKAM", "BASAVANAGUDI",
        "PUDUPAKKAM", "URAPAKKAM", "THANJAVUR", "PAMMAL", "KUMBAKONAM",
        "MADURAVOYAL", "KANDIGAI"
    ]
    
    STATUSES = [
        "Success", "Failure", "Initiated", "Awaited", 
        "Timeout", "Unsuccessful", "Aborted"
    ]
    
    AMOUNT_COLUMNS = [
        "Total Bill Amount", "Total Discount Amount", 
        "Total Tax Amount", "Net Amount"
    ]
    
    BILL_COLUMN = "Bill No"