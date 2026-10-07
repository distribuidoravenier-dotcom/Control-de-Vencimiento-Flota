import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # ID del Google Sheet
    SPREADSHEET_ID = os.getenv(
        'SPREADSHEET_ID',
        '1Rdq1lsHQgSSTZHPUj_CLIeky-lTTzjEquyKB5IZMp50'
    )

    # Nombre exacto de la pestaña
    SHEET_NAME = 'Maestro de Flota - Camion T2'

    # Primary key
    PRIMARY_KEY = 'PATENTE'

    # Credenciales
    CREDENTIALS_FILE = 'credentials.json'
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-key-change-me')

    # Definición de columnas (orden, tipo, opciones)
    # tipo: 'text' | 'number' | 'date' | 'select'
    COLUMNS = [
        {'name': 'FLOTA',                    'type': 'text',   'options': []},
        {'name': 'PATENTE',                  'type': 'text',   'options': [], 'pk': True},
        {'name': 'CODIGO DE CAMIÓN',         'type': 'number', 'options': []},
        {'name': 'PROPIETARIO',              'type': 'text',   'options': []},
        {'name': 'CIA SEGURO',               'type': 'text',   'options': []},
        {'name': 'MARCA',                    'type': 'text',   'options': []},
        {'name': 'MODELO',                   'type': 'text',   'options': []},
        {'name': 'AÑO',                      'type': 'number', 'options': []},
        {'name': 'ANTIGÜEDAD',               'type': 'number', 'options': []},
        {'name': 'ODOMETRO',                 'type': 'number', 'options': []},
        {'name': 'Cumple Polìtica?',         'type': 'select', 'options': ['SI', 'NO']},
        {'name': 'TIPO',                     'type': 'text',   'options': []},
        {'name': 'Operación',                'type': 'text',   'options': []},
        {'name': 'Tara',                     'type': 'number', 'options': []},
        {'name': 'Capacidad Maxima (kg)',    'type': 'number', 'options': []},
        {'name': 'Capacidad Maxima (pal)',   'type': 'number', 'options': []},
        {'name': 'Tiene telemetria',         'type': 'select', 'options': ['SI', 'NO']},
        {'name': 'STATUS',                   'type': 'select', 'options': ['SALE', 'NO SALE']},
        {'name': 'VENC VTV',                 'type': 'date',   'options': []},
        {'name': 'VENC SEGURO',              'type': 'date',   'options': []},
        {'name': 'UTA',                      'type': 'date',   'options': []},
        {'name': 'EXTINTOR',                 'type': 'date',   'options': []},
        {'name': 'BOTIQUIN',                 'type': 'date',   'options': []},
    ]