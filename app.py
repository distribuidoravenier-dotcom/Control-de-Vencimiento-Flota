# app.py
import os
import json
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from config import Config

app = Flask(__name__)
app.config['SECRET_KEY'] = Config.SECRET_KEY

SPREADSHEET_ID = Config.SPREADSHEET_ID
SHEETS = Config.SHEETS
CONTROL_DOCUMENTARIO = Config.CONTROL_DOCUMENTARIO

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

_service_cache = None

# Columnas de vencimiento por solapa (las que definen STATUS)
FECHAS_VENCIMIENTO = {
    'camiones_t2': ['VENC VTV', 'VENC SEGURO', 'UTA', 'EXTINTOR', 'BOTIQUIN'],
    'maestro_ae':  ['VENC SEGURO', 'EXTINTOR'],
}


# ============================================================
# CÁLCULOS AUTOMÁTICOS
# ============================================================

def parse_date_any(str_val):
    """Intenta parsear dd/mm/yyyy, dd/mm/yy, yyyy-mm-dd. Devuelve date o None."""
    if not str_val:
        return None
    s = str(str_val).strip()
    if not s:
        return None
    formatos = ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d-%m-%Y', '%d-%m-%y')
    for fmt in formatos:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def calc_cumple_politica(valor_antiguedad):
    """ANTIGÜEDAD > 10 => 'NO', si no => 'SI'."""
    try:
        ant = float(str(valor_antiguedad).replace(',', '.').strip())
    except (ValueError, TypeError):
        return ''
    return 'NO' if ant > 10 else 'SI'


def calc_status(sheet_key, row_dict):
    """
    Si CUALQUIERA de las fechas de vencimiento es menor a hoy => 'NO SALE'.
    Si no hay fechas cargadas o todas son futuras => 'SALE'.
    """
    fechas = FECHAS_VENCIMIENTO.get(sheet_key, [])
    hoy = datetime.now().date()
    for col in fechas:
        d = parse_date_any(row_dict.get(col, ''))
        if d and d < hoy:
            return 'NO SALE'
    return 'SALE'


def recalcular_campos_auto(sheet_key, row_dict):
    """Aplica los cálculos automáticos sobre un dict de fila (in-place)."""
    if 'Cumple Polìtica?' in row_dict:
        row_dict['Cumple Polìtica?'] = calc_cumple_politica(row_dict.get('ANTIGÜEDAD', ''))
    if 'STATUS' in row_dict:
        row_dict['STATUS'] = calc_status(sheet_key, row_dict)
    return row_dict


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


def get_sheet_config(sheet_key):
    """Devuelve la configuración de la solapa solicitada."""
    if sheet_key not in SHEETS:
        return None
    cfg = SHEETS[sheet_key]
    headers = [c['name'] for c in cfg['columns']]
    return {
        'key': sheet_key,
        'name': cfg['name'],
        'label': cfg['label'],
        'primary_key': cfg['primary_key'],
        'columns': cfg['columns'],
        'headers': headers,
        'last_col': col_letter(len(headers))
    }


def ensure_sheet_headers(sheet_cfg):
    """Verifica que la fila 1 del Sheet tenga exactamente nuestros headers."""
    try:
        service = get_sheets_service()
        current = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!1:1"
        ).execute().get('values', [[]])
        current_headers = current[0] if current else []

        if current_headers != sheet_cfg['headers']:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{sheet_cfg['name']}'!A1",
                valueInputOption='RAW',
                body={'values': [sheet_cfg['headers']]}
            ).execute()
            print(f"✅ Headers sincronizados en '{sheet_cfg['name']}'.")
        return True
    except HttpError as err:
        print(f"❌ Error asegurando headers: {err}")
        return False


def get_all_data(sheet_cfg):
    """Devuelve {'headers': [...], 'rows': [{..., '_row_number': N}]}"""
    try:
        service = get_sheets_service()
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!A:{sheet_cfg['last_col']}"
        ).execute()

        values = result.get('values', [])
        if not values:
            return {'headers': sheet_cfg['headers'], 'rows': []}

        headers = values[0]
        rows = []
        for i, raw in enumerate(values[1:], start=2):
            if not any(cell.strip() for cell in raw):
                continue  # fila vacía
            row_data = {}
            for j, h in enumerate(headers):
                row_data[h] = raw[j] if j < len(raw) else ''
            row_data['_row_number'] = i
            recalcular_campos_auto(sheet_cfg['key'], row_data)
            rows.append(row_data)
        return {'headers': headers, 'rows': rows}
    except HttpError as err:
        print(f"❌ Error leyendo datos: {err}")
        return {'headers': sheet_cfg['headers'], 'rows': []}


def get_sheet_id(sheet_cfg):
    service = get_sheets_service()
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        fields='sheets.properties(sheetId,title)'
    ).execute()
    for s in spreadsheet.get('sheets', []):
        if s['properties']['title'] == sheet_cfg['name']:
            return s['properties']['sheetId']
    return None


# ============================================================
# HISTORIAL DE VENCIMIENTOS
# ============================================================

def get_historial_config():
    """Configuración de la solapa Historial de Vencimiento."""
    cfg = {
        'name': Config.HISTORIAL_SHEET_NAME,
        'label': 'Historial de Vencimientos',
        'primary_key': None,
        'columns': Config.HISTORIAL_COLUMNS,
    }
    headers = [c['name'] for c in cfg['columns']]
    return {
        'key': 'historial_vencimientos',
        'name': cfg['name'],
        'label': cfg['label'],
        'primary_key': None,
        'columns': cfg['columns'],
        'headers': headers,
        'last_col': col_letter(len(headers))
    }


def semana_del_anio(fecha):
    """Devuelve el número de semana ISO de una fecha date."""
    if not fecha:
        return ''
    return fecha.isocalendar()[1]


def get_historial_data():
    """Lee todas las filas de la solapa Historial."""
    sheet_cfg = get_historial_config()
    ensure_sheet_headers(sheet_cfg)
    return get_all_data(sheet_cfg)


def ensure_historial_headers():
    sheet_cfg = get_historial_config()
    return ensure_sheet_headers(sheet_cfg)


def cargar_historial_automatico():
    """
    Recorre todas las solapas del Maestro, detecta vencimientos a <= 30 días
    y agrega filas nuevas en 'Historial de Vencimiento' que no existan ya
    (mismo TIPO + DESCRIPCION + FECHA VENCIMIENTO).
    Devuelve la cantidad de filas agregadas.
    """
    historial_cfg = get_historial_config()
    ensure_sheet_headers(historial_cfg)

    # Leer historial existente para evitar duplicados
    hist_data = get_all_data(historial_cfg)
    existentes = set()
    for r in hist_data['rows']:
        key = (
            (r.get('TIPO') or '').strip().upper(),
            (r.get('DESCRIPCION') or '').strip().upper(),
            (r.get('FECHA VENCIMIENTO') or '').strip(),
        )
        existentes.add(key)

    hoy = date.today()
    limite = hoy + timedelta(days=Config.HISTORIAL_DIAS_ALERTA)

    nuevas_filas = []

    for sheet_key, sheet_cfg_orig in SHEETS.items():
        sheet_cfg = get_sheet_config(sheet_key)
        if not sheet_cfg:
            continue
        data = get_all_data(sheet_cfg)
        for row in data['rows']:
            patente = (row.get('PATENTE') or '').strip()
            if not patente:
                continue
            for col_name in FECHAS_VENCIMIENTO.get(sheet_key, []):
                fecha_str = row.get(col_name, '')
                fecha_v = parse_date_any(fecha_str)
                if not fecha_v:
                    continue
                # Solo vencimientos dentro de los próximos 30 días (incluye ya vencidos)
                if fecha_v > limite:
                    continue

                tipo = Config.HISTORIAL_TIPO_MAP.get(col_name, col_name)
                descripcion = patente
                fecha_venc_str = fecha_v.strftime('%d/%m/%Y')

                key = (tipo.strip().upper(),
                       descripcion.strip().upper(),
                       fecha_venc_str.strip())
                if key in existentes:
                    continue

                fila = {
                    'FECHA':                 hoy.strftime('%d/%m/%Y'),
                    'TIPO':                  tipo,
                    'DESCRIPCION':           descripcion,
                    'FECHA VENCIMIENTO':     fecha_venc_str,
                    'ESTADO':                'EN PROCESO',
                    'SEMANA DE VENCIMIENTO': str(semana_del_anio(fecha_v)),
                }
                nuevas_filas.append(fila)
                existentes.add(key)

    if nuevas_filas:
        headers = historial_cfg['headers']
        values = [[fila.get(h, '') for h in headers] for fila in nuevas_filas]
        service = get_sheets_service()
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{historial_cfg['name']}'!A:{historial_cfg['last_col']}",
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': values}
        ).execute()

    return len(nuevas_filas)


# ============================================================
# RUTAS
# ============================================================

@app.route('/')
def index():
    """Página principal unificada con módulos y solapas."""
    return render_template('index.html', sheets=SHEETS, control=CONTROL_DOCUMENTARIO)


# ============================================================
# API
# ============================================================

@app.route('/api/config/<sheet_key>', methods=['GET'])
def api_config(sheet_key):
    """Devuelve la definición de columnas para que el frontend sepa tipos y opciones."""
    sheet_cfg = get_sheet_config(sheet_key)
    if not sheet_cfg:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    return jsonify({
        'success': True,
        'columns': sheet_cfg['columns'],
        'primaryKey': sheet_cfg['primary_key'],
        'sheetName': sheet_cfg['name'],
        'sheetLabel': sheet_cfg['label']
    })


@app.route('/api/rows/<sheet_key>', methods=['GET'])
def api_get_rows(sheet_key):
    sheet_cfg = get_sheet_config(sheet_key)
    if not sheet_cfg:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    ensure_sheet_headers(sheet_cfg)
    data = get_all_data(sheet_cfg)
    return jsonify({
        'success': True,
        'headers': data['headers'],
        'rows': data['rows']
    })


@app.route('/api/control-documentario', methods=['GET'])
def api_control_documentario():
    """Devuelve los datos filtrados (solo columnas de documentos) para ambas solapas."""
    result = {}
    for key, cfg in CONTROL_DOCUMENTARIO.items():
        sheet_cfg = get_sheet_config(cfg['sheet_key'])
        if not sheet_cfg:
            continue
        ensure_sheet_headers(sheet_cfg)
        data = get_all_data(sheet_cfg)
        rows = []
        for r in data['rows']:
            row = {}
            for col_name in cfg['columns']:
                row[col_name] = r.get(col_name, '')
            row['_row_number'] = r['_row_number']
            rows.append(row)
        result[key] = {
            'label': cfg['label'],
            'columns': cfg['columns'],
            'rows': rows
        }
    return jsonify({'success': True, 'data': result})


@app.route('/api/rows/<sheet_key>', methods=['POST'])
def api_add_row(sheet_key):
    """
    Body: { "values": { "PATENTE": "AB123CD", ... } }
    Valida que PATENTE no esté vacía ni duplicada.
    Las columnas readonly se calculan automáticamente.
    """
    sheet_cfg = get_sheet_config(sheet_key)
    if not sheet_cfg:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    try:
        body = request.get_json() or {}
        values = body.get('values', {})
        pk = sheet_cfg['primary_key']
        headers = sheet_cfg['headers']
        last_col = sheet_cfg['last_col']

        pk_value = (values.get(pk) or '').strip()
        if not pk_value:
            return jsonify({'success': False, 'error': f'El campo {pk} es obligatorio'}), 400

        data = get_all_data(sheet_cfg)
        for r in data['rows']:
            if (r.get(pk) or '').strip().upper() == pk_value.upper():
                return jsonify({'success': False, 'error': f'Ya existe un registro con {pk} = {pk_value}'}), 400

        row_dict = {}
        for h in headers:
            row_dict[h] = str(values.get(h, '') or '')

        recalcular_campos_auto(sheet_key, row_dict)
        row_values = [row_dict.get(h, '') for h in headers]

        service = get_sheets_service()
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!A:{last_col}",
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': [row_values]}
        ).execute()

        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ add_row: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/rows/<sheet_key>/<int:row_number>', methods=['PUT'])
def api_update_row(sheet_key, row_number):
    """
    Body: { "values": { "PATENTE": "AB123CD", ... } }
    Valida que la nueva PATENTE no duplique otra fila.
    """
    sheet_cfg = get_sheet_config(sheet_key)
    if not sheet_cfg:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    try:
        body = request.get_json() or {}
        values = body.get('values', {})
        pk = sheet_cfg['primary_key']
        headers = sheet_cfg['headers']
        last_col = sheet_cfg['last_col']

        pk_value = (values.get(pk) or '').strip()
        if not pk_value:
            return jsonify({'success': False, 'error': f'El campo {pk} es obligatorio'}), 400

        data = get_all_data(sheet_cfg)
        current_row = None
        for r in data['rows']:
            if r['_row_number'] == row_number:
                current_row = r
                continue
            if (r.get(pk) or '').strip().upper() == pk_value.upper():
                return jsonify({'success': False, 'error': f'Otra fila ya tiene {pk} = {pk_value}'}), 400

        readonly_names = {c['name'] for c in sheet_cfg['columns'] if c.get('readonly')}
        fechas = set(FECHAS_VENCIMIENTO.get(sheet_key, []))

        new_row = {}
        for h in headers:
            if h in fechas:
                new_row[h] = str(current_row.get(h, '') if current_row else '')
            elif h in readonly_names:
                new_row[h] = str(values.get(h, '') or '')
            else:
                new_row[h] = str(values.get(h, '') or '')

        recalcular_campos_auto(sheet_key, new_row)
        row_values = [new_row.get(h, '') for h in headers]

        service = get_sheets_service()
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!A{row_number}:{last_col}{row_number}",
            valueInputOption='RAW',
            body={'values': [row_values]}
        ).execute()

        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ update_row: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/control-documentario/<sheet_key>/<int:row_number>', methods=['PUT'])
def api_update_control_documentario(sheet_key, row_number):
    """
    Edita SOLO las columnas del módulo Control Documentario de una fila.
    Recalcula Cumple Polìtica? y STATUS automáticamente.
    """
    if sheet_key not in CONTROL_DOCUMENTARIO:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    sheet_cfg = get_sheet_config(sheet_key)
    if not sheet_cfg:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    try:
        body = request.get_json() or {}
        values = body.get('values', {})

        cd_cols = CONTROL_DOCUMENTARIO[sheet_key]['columns']
        editable_cols = [c for c in cd_cols if c != sheet_cfg['primary_key']]

        data = get_all_data(sheet_cfg)
        current_row = None
        for r in data['rows']:
            if r['_row_number'] == row_number:
                current_row = r
                break
        if current_row is None:
            return jsonify({'success': False, 'error': 'Fila no encontrada'}), 404

        headers = sheet_cfg['headers']
        last_col = sheet_cfg['last_col']

        new_row = {}
        for h in headers:
            if h in editable_cols:
                new_row[h] = str(values.get(h, '') or '')
            else:
                new_row[h] = str(current_row.get(h, '') or '')

        recalcular_campos_auto(sheet_key, new_row)
        new_row_values = [new_row.get(h, '') for h in headers]

        service = get_sheets_service()
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!A{row_number}:{last_col}{row_number}",
            valueInputOption='RAW',
            body={'values': [new_row_values]}
        ).execute()

        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ update_control_documentario: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/rows/<sheet_key>/<int:row_number>', methods=['DELETE'])
def api_delete_row(sheet_key, row_number):
    """Elimina una fila por número de fila real del Sheet."""
    sheet_cfg = get_sheet_config(sheet_key)
    if not sheet_cfg:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    try:
        sheet_id = get_sheet_id(sheet_cfg)
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


# ============================================================
# API HISTORIAL DE VENCIMIENTOS
# ============================================================

@app.route('/api/historial-vencimientos', methods=['GET'])
def api_historial_vencimientos():
    """
    Devuelve las filas del historial.
    Antes de leer, corre la carga automática de vencimientos próximos (<= 30 días).
    """
    try:
        # Carga automática (solo agrega lo que falte)
        agregadas = cargar_historial_automatico()
    except Exception as e:
        print(f"⚠️ Error en carga automática: {e}")
        agregadas = 0

    sheet_cfg = get_historial_config()
    ensure_sheet_headers(sheet_cfg)
    data = get_all_data(sheet_cfg)
    return jsonify({
        'success': True,
        'headers': data['headers'],
        'rows': data['rows'],
        'agregadas': agregadas,
        'estados': Config.HISTORIAL_ESTADOS
    })


@app.route('/api/historial-vencimientos/<int:row_number>', methods=['PUT'])
def api_update_historial(row_number):
    """
    Actualiza el ESTADO de una fila del historial.
    Body: { "estado": "EN PROCESO" | "VENCIDO" | "COMPLETO" }
    """
    if row_number < 2:
        return jsonify({'success': False, 'error': 'Fila inválida'}), 400
    try:
        body = request.get_json() or {}
        nuevo_estado = (body.get('estado') or '').strip().upper()
        if nuevo_estado not in [e.upper() for e in Config.HISTORIAL_ESTADOS]:
            return jsonify({'success': False, 'error': 'Estado inválido'}), 400

        sheet_cfg = get_historial_config()
        headers = sheet_cfg['headers']
        if 'ESTADO' not in headers:
            return jsonify({'success': False, 'error': 'Columna ESTADO no encontrada'}), 500

        col_idx = headers.index('ESTADO')  # 0-based
        col_letra = col_letter(col_idx + 1)

        service = get_sheets_service()
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!{col_letra}{row_number}",
            valueInputOption='RAW',
            body={'values': [[nuevo_estado]]}
        ).execute()

        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ update_historial: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
