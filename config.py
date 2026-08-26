import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # Google Sheets API
    SPREADSHEET_ID = '1h12aEo5pwGl_5dl6VZjemuqL_3hA_yauaFMDpVItdB0'
    
    # Google Drive Folder ID
    DRIVE_FOLDER_ID = '1dGN_0wVCIb30gzF7_kn6ciG0Y3U2VMAs'
    
    # Sheets names
    SHEETS = {
        'camion_t1': 'Camion T1',
        'camion_t2': 'Camion T2',
        'autoelevadores': 'Autoelevadores',
        'choferes': 'Choferes y Ayudantes'
    }
    
    # Google API credentials file
    CREDENTIALS_FILE = 'credentials.json'
    
    # Secret key for Flask
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-key-12345')