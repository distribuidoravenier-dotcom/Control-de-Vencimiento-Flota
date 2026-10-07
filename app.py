import os
import json
from flask import Flask, render_template, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import Config

app = Flask(__name__)
app.config['SECRET_KEY'] = Config.SECRET_KEY

SPREADSHEET_ID = Config.SPREADSHEET_ID
SHEET_NAME = Config.SHEET_NAME
PRIMARY_KEY = Config.PRIMARY_KEY
COLUMNS = Config.COLUMNS
HEADERS = [c['name'] for c in COLUMNS]

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

_service_cache = None


# ============================================================
# GOOGLE SHEETS
# ============================================================

def get_sheets_service():
    global _service_cache
    if _service_cache is not None:
        return _service_cache
    try:
        creds_env = os.getenv('GOOGLE_CREDENTIALS_JSON')
        if creds_env:
            info = json.loads(creds_env)
            creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        else:
            creds = service_account.Credentials.from_service_account_file(
                Config.CREDENTIALS_FILE, scopes=SCOPES
            )
        _service_cache = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        return _service_cache
    except Exception as e:
        print(f"❌ Error credenciales: {e}")
        raise


def col_letter(n):
    """1 -> A, 26 -> Z, 27 -> AA ..."""
    result = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


LAST_COL = col_letter(len(HEADERS))


def ensure_sheet_headers():
    """Verifica que la fila 1 del Sheet tenga exactamente nuestros headers."""
    try:
        service = get_sheets_service()
        current = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!1:1"
        ).execute().get('values', [[]])
        current_headers = current[0] if current else []

        if current_headers != HEADERS:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{SHEET_NAME}'!A1",
                valueInputOption='RAW',
                body={'values': [HEADERS]}
            ).execute()
            print("✅ Headers sincronizados en el Sheet.")
        return True
    except HttpError as err:
        print(f"❌ Error asegurando headers: {err}")
        return False


def get_all_data():
    """Devuelve {'headers': [...], 'rows': [{..., '_row_number': N}]}"""
    try:
        service = get_sheets_service()
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A:{LAST_COL}"
        ).execute()

        values = result.get('values', [])
        if not values:
            return {'headers': HEADERS, 'rows': []}

        headers = values[0]
        rows = []
        for i, raw in enumerate(values[1:], start=2):
            if not any(cell.strip() for cell in raw):
                continue  # fila vacía
            row_data = {}
            for j, h in enumerate(headers):
                row_data[h] = raw[j] if j < len(raw) else ''
            row_data['_row_number'] = i
            rows.append(row_data)
        return {'headers': headers, 'rows': rows}
    except HttpError as err:
        print(f"❌ Error leyendo datos: {err}")
        return {'headers': HEADERS, 'rows': []}


# ============================================================
# RUTAS
# ============================================================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/config', methods=['GET'])
def api_config():
    """Devuelve la definición de columnas para que el frontend sepa tipos y opciones."""
    return jsonify({
        'success': True,
        'columns': COLUMNS,
        'primaryKey': PRIMARY_KEY,
        'sheetName': SHEET_NAME
    })


@app.route('/api/rows', methods=['GET'])
def api_get_rows():
    ensure_sheet_headers()
    data = get_all_data()
    return jsonify({
        'success': True,
        'headers': data['headers'],
        'rows': data['rows']
    })


@app.route('/api/rows', methods=['POST'])
def api_add_row():
    """
    Body: { "values": { "PATENTE": "AB123CD", ... } }
    Valida que PATENTE no esté vacía ni duplicada.
    """
    try:
        body = request.get_json() or {}
        values = body.get('values', {})

        pk_value = (values.get(PRIMARY_KEY) or '').strip()
        if not pk_value:
            return jsonify({'success': False, 'error': f'El campo {PRIMARY_KEY} es obligatorio'}), 400

        # Chequear duplicado
        data = get_all_data()
        for r in data['rows']:
            if (r.get(PRIMARY_KEY) or '').strip().upper() == pk_value.upper():
                return jsonify({'success': False, 'error': f'Ya existe un registro con {PRIMARY_KEY} = {pk_value}'}), 400

        row_values = [str(values.get(h, '') or '') for h in HEADERS]

        service = get_sheets_service()
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A:{LAST_COL}",
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': [row_values]}
        ).execute()

        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ add_row: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/rows/<int:row_number>', methods=['PUT'])
def api_update_row(row_number):
    """
    Body: { "values": { "PATENTE": "AB123CD", ... } }
    Valida que la nueva PATENTE no duplique otra fila.
    """
    try:
        body = request.get_json() or {}
        values = body.get('values', {})

        pk_value = (values.get(PRIMARY_KEY) or '').strip()
        if not pk_value:
            return jsonify({'success': False, 'error': f'El campo {PRIMARY_KEY} es obligatorio'}), 400

        data = get_all_data()
        for r in data['rows']:
            if r['_row_number'] == row_number:
                continue
            if (r.get(PRIMARY_KEY) or '').strip().upper() == pk_value.upper():
                return jsonify({'success': False, 'error': f'Otra fila ya tiene {PRIMARY_KEY} = {pk_value}'}), 400

        row_values = [str(values.get(h, '') or '') for h in HEADERS]

        service = get_sheets_service()
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A{row_number}:{LAST_COL}{row_number}",
            valueInputOption='RAW',
            body={'values': [row_values]}
        ).execute()

        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ update_row: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/rows/<int:row_number>', methods=['DELETE'])
def api_delete_row(row_number):
    """Elimina una fila por número de fila real del Sheet."""
    try:
        sheet_id = get_sheet_id()
        if sheet_id is None:
            return jsonify({'success': False, 'error': 'Pestaña no encontrada'}), 500

        service = get_sheets_service()
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'requests': [{
                'deleteDimension': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'ROWS',
                        'startIndex': row_number - 1,
                        'endIndex': row_number
                    }
                }
            }]}
        ).execute()

        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ delete_row: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


def get_sheet_id():
    service = get_sheets_service()
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        fields='sheets.properties(sheetId,title)'
    ).execute()
    for s in spreadsheet.get('sheets', []):
        if s['properties']['title'] == SHEET_NAME:
            return s['properties']['sheetId']
    return None


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)