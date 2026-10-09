# config.py
import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # ID del Google Sheet
    SPREADSHEET_ID = os.getenv(
        'SPREADSHEET_ID',
        '1Rdq1lsHQgSSTZHPUj_CLIeky-lTTzjEquyKB5IZMp50'
    )

    # Credenciales
    CREDENTIALS_FILE = 'credentials.json'
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-key-change-me')

    # Definición de solapas (sheets) dentro del mismo spreadsheet
    SHEETS = {
        'camiones_t2': {
            'name': 'Maestro de Flota - Camion T2',
            'label': 'Camiones T2',
            'primary_key': 'PATENTE',
            'columns': [
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
                {'name': 'Cumple Polìtica?',         'type': 'select', 'options': ['SI', 'NO'], 'readonly': True},
                {'name': 'TIPO',                     'type': 'text',   'options': []},
                {'name': 'Operación',                'type': 'text',   'options': []},
                {'name': 'Tara',                     'type': 'number', 'options': []},
                {'name': 'Capacidad Maxima (kg)',    'type': 'number', 'options': []},
                {'name': 'Capacidad Maxima (pal)',   'type': 'number', 'options': []},
                {'name': 'Tiene telemetria',         'type': 'select', 'options': ['SI', 'NO']},
                {'name': 'STATUS',                   'type': 'select', 'options': ['SALE', 'NO SALE'], 'readonly': True},
                {'name': 'VENC VTV',                 'type': 'date',   'options': [], 'readonly': True},
                {'name': 'VENC SEGURO',              'type': 'date',   'options': [], 'readonly': True},
                {'name': 'UTA',                      'type': 'date',   'options': [], 'readonly': True},
                {'name': 'EXTINTOR',                 'type': 'date',   'options': [], 'readonly': True},
                {'name': 'BOTIQUIN',                 'type': 'date',   'options': [], 'readonly': True},
            ]
        },
        'maestro_ae': {
            'name': 'Maestro de Flota - AE',
            'label': 'Maestro AE',
            'primary_key': 'PATENTE',
            'columns': [
                {'name': 'FLOTA',                  'type': 'text',   'options': []},
                {'name': 'PATENTE',                'type': 'text',   'options': [], 'pk': True},
                {'name': 'CODIGO DE AE',           'type': 'text',   'options': []},
                {'name': 'PROPIETARIO',            'type': 'text',   'options': []},
                {'name': 'CIA SEGURO',             'type': 'text',   'options': []},
                {'name': 'MARCA',                  'type': 'text',   'options': []},
                {'name': 'MODELO',                 'type': 'text',   'options': []},
                {'name': 'AÑO',                    'type': 'number', 'options': []},
                {'name': 'Motor',                  'type': 'text',   'options': []},
                {'name': 'HORAS',                  'type': 'number', 'options': []},
                {'name': 'ANTIGÜEDAD',             'type': 'number', 'options': []},
                {'name': 'Cumple Polìtica?',       'type': 'select', 'options': ['SI', 'NO'], 'readonly': True},
                {'name': 'TIPO',                   'type': 'text',   'options': []},
                {'name': 'Operación',              'type': 'text',   'options': []},
                {'name': 'Capacidad de Carga',     'type': 'text',   'options': []},
                {'name': 'VENC SEGURO',            'type': 'date',   'options': [], 'readonly': True},
                {'name': 'EXTINTOR',               'type': 'date',   'options': [], 'readonly': True},
            ]
        }
    }

    # Definición de solapas para el módulo "Control Documentario"
    CONTROL_DOCUMENTARIO = {
        'camiones_t2': {
            'sheet_key': 'camiones_t2',
            'label': 'Camiones T2',
            'columns': ['PATENTE', 'VENC VTV', 'VENC SEGURO', 'UTA', 'EXTINTOR', 'BOTIQUIN']
        },
        'maestro_ae': {
            'sheet_key': 'maestro_ae',
            'label': 'Maestro AE',
            'columns': ['PATENTE', 'VENC SEGURO', 'EXTINTOR']
        }
    }

    # ============================================================
    # MÓDULO HISTORIAL DE VENCIMIENTOS
    # ============================================================
    # Solapa destino donde se acumula el historial
    HISTORIAL_SHEET_NAME = 'Historial de Vencimiento'

    # Mapeo: nombre de columna en Maestro -> valor de "TIPO" en Historial
    HISTORIAL_TIPO_MAP = {
        'VENC VTV':    'VTV',
        'VENC SEGURO': 'SEGURO',
        'UTA':         'UTA',
        'EXTINTOR':    'MATAFUEGO',
        'BOTIQUIN':    'BOTIQUIN',
    }

    # Días de anticipación para cargar automáticamente en el historial
    HISTORIAL_DIAS_ALERTA = 30

    # Estados permitidos en la columna ESTADO
    HISTORIAL_ESTADOS = ['EN PROCESO', 'VENCIDO', 'COMPLETO']

    # Columnas de la solapa "Historial de Vencimiento"
    HISTORIAL_COLUMNS = [
        {'name': 'FECHA',                 'type': 'date',   'options': []},
        {'name': 'TIPO',                  'type': 'text',   'options': []},
        {'name': 'DESCRIPCION',           'type': 'text',   'options': []},
        {'name': 'FECHA VENCIMIENTO',     'type': 'date',   'options': []},
        {'name': 'ESTADO',                'type': 'select', 'options': ['EN PROCESO', 'VENCIDO', 'COMPLETO']},
        {'name': 'SEMANA DE VENCIMIENTO', 'type': 'number', 'options': []},
    ]
