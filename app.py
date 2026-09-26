import os
import io
import json
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

app = Flask(__name__)

# Configuración
SPREADSHEET_ID = '1h12aEo5pwGl_5dl6VZjemuqL_3hA_yauaFMDpVItdB0'
DRIVE_FOLDER_ID = '1dGN_0wVCIb30gzF7_kn6ciG0Y3U2VMAs'

# Configuración Mantenimientos
SPREADSHEET_ID_MP = '1ZCgxrjf84jAlc6F8KsiuvFXkZ2qwbUEHfDAHwnfDhzg'
SHEET_HISTORIAL_MP = 'Historial Mantenimientos'
SHEET_CONFIG_MP = 'Configuracion MP'
SHEET_CONFIG_DROPDOWNS = 'Configuracion Desplegables'
SHEET_PROGRAMACION_MP = 'Programacion Mantenimiento Preventivo'

# Headers de la hoja de Programación
# (SIN GASTO en PESOS / GASTO en USD: el gasto solo se registra en Historial Mantenimientos,
# cuando la tarea ya se realizó)
PROG_HEADERS = [
    'Marca Temporal',
    'PATENTE',
    'FECHA ULTIMO MANTENIMIENTO',
    'PROXIMO MANTENIMIENTO FECHA',
    'KM ULTIMO MANTENIMIENTO',
    'PROXIMO MANTENIMIENTO KM',
    'TIPO MANTENIMIENTO',
    'TIPO REPARACIÓN',
    'DETALLE REPARACIÓN',
    'Estado'
]

# Headers del Historial de Mantenimientos (para asegurar que existan los campos de gasto)
HIST_HEADERS = [
    'FECHA',
    'PATENTE',
    'TIPO MANTENIMIENTO',
    'TIPO REPARACIÓN',
    'CONCEPTO GENERAL',
    'DETALLE REPARACIÓN',
    'KM REALIZADO',
    'KM PROXIMO CAMBIO',
    'PROVEEDOR',
    'GASTO en PESOS',
    'GASTO en USD',
    'OBSERVACIONES'
]

SHEETS = {
    'Camion T1': 'Camion T1',
    'Camion T2': 'Camion T2',
    'Autoelevadores': 'Autoelevadores',
    'Choferes y Ayudantes': 'Choferes y Ayudantes',
    'Historial de Vencimiento de Documentación': 'Historial de Vencimiento de Documentación'
}

HISTORY_SHEET = 'Historial de Vencimiento de Documentación'

DATE_COLUMNS_CONFIG = {
    'Camion T1': ['VENC VTV', 'VENC SEGURO', 'SENASA', 'LICENCIA DE CONDUCIR', 'PAGO MONOTRIBUTO',
                  'POLIZA DE SEGUROS', 'SEGURO DE ACCIDENTES PERSONALES', 'CLAUSULA DE NO REPETICION',
                  'INDUCCION DE CMQ', 'CAPACITACION DE MANEJO DEFENSIVO'],
    'Camion T2': ['VENC VTV', 'VENC SEGURO', 'UTA', 'EXTINTOR', 'BOTIQUIN'],
    'Autoelevadores': ['VENC SEGURO', 'EXTINTOR'],
    'Choferes y Ayudantes': ['VENCIMIENTO REGISTRO', 'LIBRETA SANITARIA']
}

ID_FIELD_CONFIG = {
    'Camion T1': 'PATENTE',
    'Camion T2': 'PATENTE',
    'Autoelevadores': 'CODIGO DE AE',
    'Choferes y Ayudantes': 'APELLIDO Y NOMBRE'
}

app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-key-12345')
CORS(app)

_creds_cache = None
_sheets_service_cache = None
_drive_service_cache = None

def get_google_creds():
    global _creds_cache
    if _creds_cache is not None:
        return _creds_cache
    try:
        if 'GOOGLE_APPLICATION_CREDENTIALS_JSON' in os.environ:
            creds_json = json.loads(os.environ['GOOGLE_APPLICATION_CREDENTIALS_JSON'])
            creds = service_account.Credentials.from_service_account_info(
                creds_json,
                scopes=[
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive'
                ]
            )
        else:
            creds = service_account.Credentials.from_service_account_file(
                'credentials.json',
                scopes=[
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive'
                ]
            )
        _creds_cache = creds
        return creds
    except Exception as e:
        print(f"Error al obtener credenciales: {e}")
        raise

def get_sheets_service():
    """Devuelve un cliente de Sheets reutilizable en vez de crear uno nuevo en cada llamada."""
    global _sheets_service_cache
    if _sheets_service_cache is None:
        _sheets_service_cache = build('sheets', 'v4', credentials=get_google_creds())
    return _sheets_service_cache

def get_drive_service():
    """Devuelve un cliente de Drive reutilizable en vez de crear uno nuevo en cada llamada."""
    global _drive_service_cache
    if _drive_service_cache is None:
        _drive_service_cache = build('drive', 'v3', credentials=get_google_creds())
    return _drive_service_cache

def get_all_data(sheet_name, spreadsheet_id=None):
    try:
        service = get_sheets_service()
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        result = sheet.values().get(
            spreadsheetId=sid,
            range=f"'{sheet_name}'!A:Z"
        ).execute()

        values = result.get('values', [])
        if not values:
            return {'headers': [], 'rows': []}

        headers = values[0] if values else []
        rows = []

        for i, row in enumerate(values[1:], start=2):
            row_data = {}
            for j, header in enumerate(headers):
                if j < len(row):
                    row_data[header] = row[j]
                else:
                    row_data[header] = ''
            row_data['_row_number'] = i
            rows.append(row_data)

        return {'headers': headers, 'rows': rows}

    except HttpError as err:
        print(f"Error getting sheet data [{sheet_name}]: {err}")
        return {'headers': [], 'rows': []}

def update_row(sheet_name, row_number, values, spreadsheet_id=None):
    try:
        service = get_sheets_service()
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        data = get_all_data(sheet_name, sid)
        num_columns = len(data.get('headers', []))

        while len(values) < num_columns:
            values.append('')

        last_col = chr(64 + num_columns) if num_columns <= 26 else 'Z'

        body = {'values': [values[:num_columns]]}

        sheet.values().update(
            spreadsheetId=sid,
            range=f"'{sheet_name}'!A{row_number}:{last_col}{row_number}",
            valueInputOption='RAW',
            body=body
        ).execute()

        return True

    except HttpError as err:
        print(f"Error updating row: {err}")
        return False

def update_row_partial(sheet_name, row_number, column_updates):
    try:
        service = get_sheets_service()
        sheet = service.spreadsheets()

        data = get_all_data(sheet_name)
        headers = data.get('headers', [])

        requests = []
        for col_name, value in column_updates.items():
            if col_name in headers:
                col_idx = headers.index(col_name)
                col_letter = chr(65 + col_idx) if col_idx < 26 else 'Z'
                requests.append({
                    'range': f"'{sheet_name}'!{col_letter}{row_number}",
                    'values': [[value]]
                })

        if not requests:
            return False

        body = {'valueInputOption': 'RAW', 'data': requests}

        sheet.values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()

        return True

    except HttpError as err:
        print(f"Error updating row partial: {err}")
        return False

def update_prog_estado(row_number, estado):
    try:
        service = get_sheets_service()
        sheet = service.spreadsheets()

        data = get_all_data(SHEET_PROGRAMACION_MP, SPREADSHEET_ID_MP)
        headers = data.get('headers', [])

        if 'Estado' not in headers:
            print(f"⚠️ 'Estado' no está en los headers de la hoja: {headers}")
            return False

        col_idx = headers.index('Estado')
        col_letter = chr(65 + col_idx) if col_idx < 26 else 'Z'

        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID_MP,
            range=f"'{SHEET_PROGRAMACION_MP}'!{col_letter}{row_number}",
            valueInputOption='RAW',
            body={'values': [[estado]]}
        ).execute()

        return True
    except HttpError as err:
        print(f"Error en update_prog_estado: {err}")
        return False

def delete_row(sheet_name, row_number, spreadsheet_id=None):
    try:
        service = get_sheets_service()
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        spreadsheet = service.spreadsheets().get(spreadsheetId=sid).execute()

        sheet_id = None
        for s in spreadsheet.get('sheets', []):
            if s['properties']['title'] == sheet_name:
                sheet_id = s['properties']['sheetId']
                break

        if sheet_id is None:
            return False

        requests = [{
            'deleteDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': row_number - 1,
                    'endIndex': row_number
                }
            }
        }]

        body = {'requests': requests}

        sheet.batchUpdate(spreadsheetId=sid, body=body).execute()

        return True

    except HttpError as err:
        print(f"Error deleting row: {err}")
        return False

def add_row_to_sheet(sheet_name, values, spreadsheet_id=None):
    try:
        service = get_sheets_service()
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        data = get_all_data(sheet_name, sid)
        num_columns = len(data.get('headers', []))

        while len(values) < num_columns:
            values.append('')

        body = {'values': [values[:num_columns]]}

        sheet.values().append(
            spreadsheetId=sid,
            range=f"'{sheet_name}'!A:Z",
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()

        return True

    except HttpError as err:
        print(f"Error adding row: {err}")
        return False

def ensure_sheet_exists(sheet_name, spreadsheet_id=None, headers=None):
    """
    Crea la hoja si no existe Y escribe/actualiza los headers.
    Si la hoja existe pero los headers no coinciden, los reescribe.
    Esta versión es TOLERANTE a fallos: si no puede actualizar headers,
    NO rompe el flujo (devuelve False en vez de lanzar excepción).
    """
    try:
        service = get_sheets_service()
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        spreadsheet = service.spreadsheets().get(spreadsheetId=sid).execute()
        exists = False
        for s in spreadsheet.get('sheets', []):
            if s['properties']['title'] == sheet_name:
                exists = True
                break

        if not exists:
            print(f"📄 Creando hoja '{sheet_name}'...")
            requests = [{
                'addSheet': {
                    'properties': {'title': sheet_name}
                }
            }]
            try:
                sheet.batchUpdate(spreadsheetId=sid, body={'requests': requests}).execute()
            except HttpError as err:
                print(f"⚠️ No se pudo crear la hoja '{sheet_name}': {err}")
                return False

            if headers:
                try:
                    sheet.values().update(
                        spreadsheetId=sid,
                        range=f"'{sheet_name}'!A1",
                        valueInputOption='RAW',
                        body={'values': [headers]}
                    ).execute()
                except HttpError as err:
                    print(f"⚠️ No se pudieron escribir headers en '{sheet_name}': {err}")
                    return False
            print(f"✅ Hoja '{sheet_name}' creada con headers: {headers}")
        else:
            # Verificar/actualizar headers
            if headers:
                try:
                    result = sheet.values().get(
                        spreadsheetId=sid,
                        range=f"'{sheet_name}'!A1:Z1"
                    ).execute()
                    current_headers = result.get('values', [[]])[0] if result.get('values') else []

                    # Si los headers están vacíos o difieren, los reescribimos
                    if not current_headers or current_headers != headers:
                        print(f"⚠️ Headers de '{sheet_name}' no coinciden. Actualizando...")
                        print(f"   Actuales: {current_headers}")
                        print(f"   Esperados: {headers}")

                        # Leer TODOS los datos actuales para preservarlos
                        try:
                            full_result = sheet.values().get(
                                spreadsheetId=sid,
                                range=f"'{sheet_name}'!A:Z"
                            ).execute()
                            full_values = full_result.get('values', [])
                        except HttpError as err:
                            print(f"⚠️ No se pudieron leer datos de '{sheet_name}': {err}")
                            full_values = []

                        # Mapear filas actuales a los nuevos headers
                        new_rows = []
                        if full_values and len(full_values) > 1:
                            old_headers = full_values[0]
                            for old_row in full_values[1:]:
                                new_row = []
                                for h in headers:
                                    if h in old_headers:
                                        idx = old_headers.index(h)
                                        new_row.append(old_row[idx] if idx < len(old_row) else '')
                                    else:
                                        new_row.append('')
                                new_rows.append(new_row)

                        # Limpiar la hoja
                        try:
                            sheet.values().clear(
                                spreadsheetId=sid,
                                range=f"'{sheet_name}'!A:Z"
                            ).execute()
                        except HttpError as err:
                            print(f"⚠️ No se pudo limpiar '{sheet_name}': {err}")
                            return False

                        # Escribir headers + datos mapeados
                        body_values = [headers] + new_rows
                        try:
                            sheet.values().update(
                                spreadsheetId=sid,
                                range=f"'{sheet_name}'!A1",
                                valueInputOption='RAW',
                                body={'values': body_values}
                            ).execute()
                            print(f"✅ Headers actualizados en '{sheet_name}' (preservando {len(new_rows)} filas)")
                        except HttpError as err:
                            print(f"⚠️ No se pudieron escribir headers/datos en '{sheet_name}': {err}")
                            return False
                except HttpError as err:
                    print(f"⚠️ No se pudo verificar headers de '{sheet_name}': {err}")
                    return False

        return True
    except HttpError as err:
        print(f"❌ Error ensuring sheet '{sheet_name}': {err}")
        return False
    except Exception as e:
        print(f"❌ Error inesperado ensuring sheet '{sheet_name}': {e}")
        return False

def upload_file_to_drive(file_content, filename, folder_id):
    try:
        service = get_drive_service()

        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }

        media = MediaIoBaseUpload(
            io.BytesIO(file_content),
            mimetype='image/jpeg',
            resumable=True
        )

        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        return file.get('id')

    except HttpError as err:
        print(f"Error uploading file: {err}")
        return None

def parse_date(date_str):
    if not date_str:
        return None
    date_str = str(date_str).strip()
    if date_str == '':
        return None

    if '/' in date_str:
        parts = date_str.split('/')
        if len(parts) == 3:
            try:
                day = int(parts[0])
                month = int(parts[1])
                year = int(parts[2])
                if year < 100:
                    year += 2000
                if 1 <= day <= 31 and 1 <= month <= 12:
                    return datetime(year, month, day)
            except ValueError:
                pass

    if '-' in date_str:
        parts = date_str.split('-')
        if len(parts) == 3:
            try:
                year = int(parts[0])
                month = int(parts[1])
                day = int(parts[2])
                if 1 <= day <= 31 and 1 <= month <= 12:
                    return datetime(year, month, day)
            except ValueError:
                pass

    return None

def get_politica_por_tipo_reparacion(tipo_reparacion):
    try:
        config_data = get_all_data(SHEET_CONFIG_MP, SPREADSHEET_ID_MP)
        for row in config_data.get('rows', []):
            tr = (row.get('TIPO REPARACIÓN') or '').strip()
            if tr and tr == tipo_reparacion.strip():
                pol = (row.get('POLITICA') or '').strip()
                try:
                    return float(pol)
                except ValueError:
                    return 0
        return 0
    except Exception as e:
        print(f"Error en get_politica_por_tipo_reparacion: {e}")
        return 0

def check_and_register_expirations():
    try:
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        limit_date = today + timedelta(days=30)

        history_data = get_all_data(HISTORY_SHEET)
        history_rows = history_data.get('rows', [])
        history_headers = history_data.get('headers', [])

        if not history_headers:
            print("La hoja de historial no existe o está vacía. Omitiendo detección automática.")
            return False

        existing_keys = set()
        for row in history_rows:
            tipo = row.get('TIPO', '').strip()
            desc = row.get('DESCRIPCION', '').strip()
            fecha_venc = row.get('FECHA VENCIMIENTO', '').strip()
            if tipo and desc and fecha_venc:
                existing_keys.add((tipo, desc, fecha_venc))

        new_rows = []
        fecha_deteccion = today.strftime('%d/%m/%Y')

        for sheet_name, date_columns in DATE_COLUMNS_CONFIG.items():
            if sheet_name == HISTORY_SHEET:
                continue

            data = get_all_data(sheet_name)
            rows = data.get('rows', [])
            id_field = ID_FIELD_CONFIG.get(sheet_name, '')

            for row in rows:
                identificador = row.get(id_field, '').strip() if id_field else ''
                if not identificador:
                    continue

                for date_col in date_columns:
                    date_value = row.get(date_col, '').strip()
                    if not date_value:
                        continue

                    exp_date = parse_date(date_value)
                    if not exp_date:
                        continue

                    if exp_date <= limit_date:
                        tipo = date_col
                        descripcion = identificador
                        fecha_venc_str = exp_date.strftime('%d/%m/%Y')

                        key = (tipo, descripcion, fecha_venc_str)
                        if key not in existing_keys:
                            fecha_deteccion = today.strftime('%d/%m/%Y')

                            new_row_values = []
                            for h in history_headers:
                                if h == 'FECHA':
                                    new_row_values.append(fecha_deteccion)
                                elif h == 'TIPO':
                                    new_row_values.append(tipo)
                                elif h == 'DESCRIPCION':
                                    new_row_values.append(descripcion)
                                elif h == 'FECHA VENCIMIENTO':
                                    new_row_values.append(fecha_venc_str)
                                elif h == 'ESTADO':
                                    new_row_values.append('En Proceso')
                                else:
                                    new_row_values.append('')

                            new_rows.append(new_row_values)
                            existing_keys.add(key)

        if new_rows:
            try:
                service = get_sheets_service()
                sheet = service.spreadsheets()

                body = {'values': new_rows}

                sheet.values().append(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"'{HISTORY_SHEET}'!A:Z",
                    valueInputOption='RAW',
                    insertDataOption='INSERT_ROWS',
                    body=body
                ).execute()

                print(f"Se agregaron {len(new_rows)} registros al historial.")
            except HttpError as err:
                print(f"Error al insertar batch en historial: {err}")

        return True

    except Exception as e:
        print(f"Error en check_and_register_expirations: {e}")
        return False

def _fecha_mas_cercana(row, date_columns):
    """Devuelve (fecha_datetime, nombre_columna) del vencimiento MAS PROXIMO de una fila."""
    nearest = None
    nearest_col = None
    for col in date_columns:
        val = (row.get(col) or '').strip()
        if not val:
            continue
        d = parse_date(val)
        if not d:
            continue
        if nearest is None or d < nearest:
            nearest = d
            nearest_col = col
    return nearest, nearest_col

def build_panel_flota():
    """
    Correlaciona, por PATENTE / CODIGO DE AE, los 3 módulos que hoy viven separados:
    - Documentación (Camion T1/T2, Autoelevadores, Choferes) -> próximo vencimiento
    - Mantenimiento (Historial Mantenimientos MP)             -> último mantenimiento
    - Programación de Mantenimiento Preventivo (MP)           -> próximo mantenimiento
    Devuelve una lista de "fichas" unificadas, una por vehículo/persona.
    """
    hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    # --- Último mantenimiento realizado, por PATENTE ---
    hist_data = get_all_data(SHEET_HISTORIAL_MP, SPREADSHEET_ID_MP)
    ultimo_mant_por_patente = {}
    for row in hist_data.get('rows', []):
        patente = (row.get('PATENTE') or '').strip()
        if not patente:
            continue
        fecha = parse_date(row.get('FECHA', ''))
        if not fecha:
            continue
        actual = ultimo_mant_por_patente.get(patente)
        if actual is None or fecha > actual['fecha']:
            ultimo_mant_por_patente[patente] = {
                'fecha': fecha,
                'tipo': row.get('TIPO REPARACIÓN', '') or row.get('TIPO MANTENIMIENTO', '')
            }

    # --- Próximo mantenimiento programado (no completo), por PATENTE ---
    prog_data = get_all_data(SHEET_PROGRAMACION_MP, SPREADSHEET_ID_MP)
    proximo_mant_por_patente = {}
    for row in prog_data.get('rows', []):
        if (row.get('Estado') or '').strip() == 'Completo':
            continue
        patente = (row.get('PATENTE') or '').strip()
        if not patente:
            continue
        fecha = parse_date(row.get('PROXIMO MANTENIMIENTO FECHA', ''))
        actual = proximo_mant_por_patente.get(patente)
        if fecha and (actual is None or fecha < actual['fecha']):
            proximo_mant_por_patente[patente] = {
                'fecha': fecha,
                'km': row.get('PROXIMO MANTENIMIENTO KM', ''),
                'tipo_reparacion': row.get('TIPO REPARACIÓN', '')
            }

    # --- Documentación: recorremos cada sheet de flota/personal ---
    panel = []
    for sheet_name, date_columns in DATE_COLUMNS_CONFIG.items():
        if sheet_name == HISTORY_SHEET:
            continue
        id_field = ID_FIELD_CONFIG.get(sheet_name, '')
        data = get_all_data(sheet_name)

        for row in data.get('rows', []):
            identificador = (row.get(id_field) or '').strip() if id_field else ''
            if not identificador:
                continue

            nearest, nearest_col = _fecha_mas_cercana(row, date_columns)
            if nearest:
                dias = (nearest - hoy).days
                if dias < 0:
                    estado_doc = 'VENCIDO'
                elif dias <= 30:
                    estado_doc = 'PROXIMO_A_VENCER'
                else:
                    estado_doc = 'OK'
            else:
                dias = None
                estado_doc = 'SIN_DATOS'

            ultimo = ultimo_mant_por_patente.get(identificador)
            proximo = proximo_mant_por_patente.get(identificador)

            panel.append({
                'categoria': sheet_name,
                'identificador': identificador,
                'estado_documentacion': estado_doc,
                'proxima_doc_tipo': nearest_col or '',
                'proxima_doc_fecha': nearest.strftime('%d/%m/%Y') if nearest else '',
                'proxima_doc_dias': dias,
                'ultimo_mantenimiento_fecha': ultimo['fecha'].strftime('%d/%m/%Y') if ultimo else '',
                'ultimo_mantenimiento_tipo': ultimo['tipo'] if ultimo else '',
                'proximo_mantenimiento_fecha': proximo['fecha'].strftime('%d/%m/%Y') if proximo else '',
                'proximo_mantenimiento_km': proximo['km'] if proximo else '',
            })

    orden_estado = {'VENCIDO': 0, 'PROXIMO_A_VENCER': 1, 'OK': 2, 'SIN_DATOS': 3}
    panel.sort(key=lambda p: (orden_estado.get(p['estado_documentacion'], 4), p['categoria'], p['identificador']))
    return panel

@app.route('/api/panel_flota')
def panel_flota():
    """
    Endpoint único que junta Documentación + Mantenimiento + Programación
    por PATENTE/CODIGO, para tener TODO correlacionado en una sola vista.
    """
    try:
        panel = build_panel_flota()
        return jsonify({'success': True, 'panel': panel})
    except Exception as e:
        print(f"Error en panel_flota: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/')
def index():
    return render_template('index.html', sheets=SHEETS)

@app.route('/api/sheet/<sheet_name>')
def get_sheet(sheet_name):
    data = get_all_data(sheet_name)
    return jsonify(data)

@app.route('/api/history')
def get_history():
    try:
        check_and_register_expirations()
        data = get_all_data(HISTORY_SHEET)
        return jsonify(data)
    except Exception as e:
        print(f"Error en get_history: {e}")
        return jsonify({'headers': [], 'rows': [], 'error': str(e)}), 500

@app.route('/api/history/update_status/<int:row_number>', methods=['POST'])
def update_history_status(row_number):
    try:
        data = request.json
        new_status = data.get('estado', '')

        if new_status not in ['En Proceso', 'Completo', 'Vencido', 'Notificado']:
            return jsonify({'success': False, 'error': 'Estado no válido'}), 400

        success = update_row_partial(HISTORY_SHEET, row_number, {'ESTADO': new_status})

        if success:
            return jsonify({'success': True, 'message': 'Estado actualizado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al actualizar el estado'}), 500

    except Exception as e:
        print(f"Error en update_history_status: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/history/delete/<int:row_number>', methods=['DELETE'])
def delete_history_row(row_number):
    try:
        success = delete_row(HISTORY_SHEET, row_number)

        if success:
            return jsonify({'success': True, 'message': 'Registro eliminado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al eliminar el registro'}), 500

    except Exception as e:
        print(f"Error en delete_history_row: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/add', methods=['POST'])
def add_document():
    try:
        sheet_name = request.form.get('sheet_name')

        form_data = {}
        for key in request.form:
            form_data[key] = request.form[key]

        data = get_all_data(sheet_name)
        headers = data.get('headers', [])

        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            elif header in form_data:
                row_values.append(form_data[header])
            else:
                row_values.append('')

        for key in request.files:
            if key.startswith('foto_'):
                foto = request.files[key]
                if foto.filename != '':
                    doc_name = key.replace('foto_', '')

                    if sheet_name in ['Camion T1', 'Camion T2']:
                        identificador = form_data.get('PATENTE', 'SIN_PATENTE')
                    elif sheet_name == 'Autoelevadores':
                        identificador = form_data.get('CODIGO DE AE', 'SIN_CODIGO')
                    elif sheet_name == 'Choferes y Ayudantes':
                        identificador = form_data.get('APELLIDO Y NOMBRE', 'SIN_NOMBRE')
                    else:
                        identificador = 'DOCUMENTO'

                    file_extension = os.path.splitext(foto.filename)[1]
                    filename = f"{identificador} - {doc_name}{file_extension}"
                    file_content = foto.read()

                    file_id = upload_file_to_drive(file_content, filename, DRIVE_FOLDER_ID)

                    if file_id:
                        drive_url = f"https://drive.google.com/file/d/{file_id}/view"
                        foto_header = f"{doc_name}_FOTO"
                        for i, header in enumerate(headers):
                            if header == foto_header:
                                row_values[i] = drive_url
                                break

        success = add_row_to_sheet(sheet_name, row_values)

        if success:
            return jsonify({'success': True, 'message': 'Documento agregado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al guardar en Google Sheets'}), 500

    except Exception as e:
        print(f"Error en add_document: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/update/<sheet_name>/<int:row_number>', methods=['POST'])
def update_document(sheet_name, row_number):
    try:
        data = request.json

        sheet_data = get_all_data(sheet_name)
        headers = sheet_data.get('headers', [])

        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            elif header in data:
                row_values.append(data[header])
            else:
                row_values.append('')

        success = update_row(sheet_name, row_number, row_values)

        if success:
            return jsonify({'success': True, 'message': 'Documento actualizado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al actualizar en Google Sheets'}), 500

    except Exception as e:
        print(f"Error en update_document: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/update_with_photo/<sheet_name>/<int:row_number>', methods=['POST'])
def update_document_with_photo(sheet_name, row_number):
    try:
        form_data = {}
        for key in request.form:
            form_data[key] = request.form[key]

        for key in request.files:
            if key.startswith('foto_'):
                foto = request.files[key]
                if foto.filename != '':
                    doc_name = key.replace('foto_', '')

                    if sheet_name in ['Camion T1', 'Camion T2']:
                        identificador = form_data.get('PATENTE', 'SIN_PATENTE')
                    elif sheet_name == 'Autoelevadores':
                        identificador = form_data.get('CODIGO DE AE', 'SIN_CODIGO')
                    elif sheet_name == 'Choferes y Ayudantes':
                        identificador = form_data.get('APELLIDO Y NOMBRE', 'SIN_NOMBRE')
                    else:
                        identificador = 'DOCUMENTO'

                    file_extension = os.path.splitext(foto.filename)[1]
                    filename = f"{identificador} - {doc_name}{file_extension}"
                    file_content = foto.read()

                    file_id = upload_file_to_drive(file_content, filename, DRIVE_FOLDER_ID)

                    if file_id:
                        drive_url = f"https://drive.google.com/file/d/{file_id}/view"
                        foto_header = f"{doc_name}_FOTO"
                        form_data[foto_header] = drive_url

        sheet_data = get_all_data(sheet_name)
        headers = sheet_data.get('headers', [])

        current_row = None
        for row in sheet_data.get('rows', []):
            if row.get('_row_number') == row_number:
                current_row = row
                break

        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            elif header in form_data and form_data[header] is not None:
                row_values.append(form_data[header])
            elif current_row and header in current_row:
                row_values.append(current_row[header])
            else:
                row_values.append('')

        success = update_row(sheet_name, row_number, row_values)

        if success:
            return jsonify({'success': True, 'message': 'Documento actualizado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al actualizar en Google Sheets'}), 500

    except Exception as e:
        print(f"Error en update_document_with_photo: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/delete/<sheet_name>/<int:row_number>', methods=['DELETE'])
def delete_document(sheet_name, row_number):
    try:
        success = delete_row(sheet_name, row_number)

        if success:
            return jsonify({'success': True, 'message': 'Documento eliminado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al eliminar el documento'}), 500

    except Exception as e:
        print(f"Error en delete_document: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

# ==================== MANTENIMIENTOS ====================

@app.route('/api/mantenimientos/config', methods=['GET'])
def get_mantenimiento_config():
    try:
        print("=" * 60)
        print("🔍 GET /api/mantenimientos/config")
        print("=" * 60)

        # Asegurar y actualizar headers de las hojas (tolerante a fallos)
        ensure_sheet_exists(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)
        ensure_sheet_exists(SHEET_CONFIG_MP, SPREADSHEET_ID_MP)
        ensure_sheet_exists(
            SHEET_PROGRAMACION_MP,
            SPREADSHEET_ID_MP,
            headers=PROG_HEADERS
        )
        # Asegurar que el historial tenga los headers nuevos (GASTO en PESOS / GASTO en USD)
        ensure_sheet_exists(
            SHEET_HISTORIAL_MP,
            SPREADSHEET_ID_MP,
            headers=HIST_HEADERS
        )

        dropdown_data = get_all_data(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)
        config_mp_data = get_all_data(SHEET_CONFIG_MP, SPREADSHEET_ID_MP)
        hist_data = get_all_data(SHEET_HISTORIAL_MP, SPREADSHEET_ID_MP)

        print(f"📋 Headers dropdowns: {dropdown_data.get('headers', [])}")
        print(f"📋 Headers config_mp: {config_mp_data.get('headers', [])}")
        print(f"📋 Headers historial: {hist_data.get('headers', [])}")

        # PATENTES + ODOMETRO desde Camion T2
        camion_t2_data = get_all_data('Camion T2', SPREADSHEET_ID)
        patentes_camion_t2 = []
        odometro_por_patente = {}
        for row in camion_t2_data.get('rows', []):
            patente = (row.get('PATENTE') or '').strip()
            if patente:
                if patente not in patentes_camion_t2:
                    patentes_camion_t2.append(patente)
                odometro = (row.get('ODOMETRO') or '').strip()
                if odometro:
                    odometro_por_patente[patente] = odometro
        patentes_camion_t2.sort()
        print(f"🚛 Patentes Camion T2: {patentes_camion_t2}")
        print(f"🛣️ Odómetros: {odometro_por_patente}")

        # Programación
        programacion_data = get_all_data(SHEET_PROGRAMACION_MP, SPREADSHEET_ID_MP)
        print(f"📋 Headers programación: {programacion_data.get('headers', [])}")
        print(f"📋 Filas programación: {len(programacion_data.get('rows', []))}")

        return jsonify({
            'success': True,
            'dropdowns': dropdown_data,
            'config_mp': config_mp_data,
            'historial_headers': hist_data.get('headers', []),
            'historial_rows': hist_data.get('rows', []),
            'patentes_camion_t2': patentes_camion_t2,
            'odometro_por_patente': odometro_por_patente,
            'programacion_headers': programacion_data.get('headers', []),
            'programacion_rows': programacion_data.get('rows', [])
        })
    except Exception as e:
        print(f"❌ Error en get_mantenimiento_config: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/mantenimientos/dropdowns', methods=['POST'])
def save_mantenimiento_dropdowns():
    try:
        data = request.json
        dropdowns = data.get('dropdowns', {})

        ensure_sheet_exists(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)

        headers = list(dropdowns.keys())

        if not headers:
            body = {'values': []}
        else:
            max_len = max(len(v) for v in dropdowns.values()) if dropdowns else 0
            rows = [headers]
            for i in range(max_len):
                row = []
                for h in headers:
                    opts = dropdowns[h]
                    row.append(opts[i] if i < len(opts) else '')
                rows.append(row)
            body = {'values': rows}

        service = get_sheets_service()
        sheet = service.spreadsheets()

        sheet.values().clear(
            spreadsheetId=SPREADSHEET_ID_MP,
            range=f"'{SHEET_CONFIG_DROPDOWNS}'!A:Z"
        ).execute()

        if headers:
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID_MP,
                range=f"'{SHEET_CONFIG_DROPDOWNS}'!A1",
                valueInputOption='RAW',
                body=body
            ).execute()

        return jsonify({'success': True, 'message': 'Desplegables guardados correctamente'})
    except Exception as e:
        print(f"Error en save_mantenimiento_dropdowns: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/mantenimientos/config_mp', methods=['POST'])
def save_config_mp():
    try:
        data = request.json
        rows = data.get('rows', [])

        ensure_sheet_exists(SHEET_CONFIG_MP, SPREADSHEET_ID_MP)

        headers = ['TIPO MANTENIMIENTO', 'TIPO REPARACIÓN', 'POLITICA']

        body_values = [headers]
        for r in rows:
            body_values.append([
                r.get('TIPO MANTENIMIENTO', ''),
                r.get('TIPO REPARACIÓN', ''),
                str(r.get('POLITICA', '')) if r.get('POLITICA', '') != '' else ''
            ])

        service = get_sheets_service()
        sheet = service.spreadsheets()

        sheet.values().clear(
            spreadsheetId=SPREADSHEET_ID_MP,
            range=f"'{SHEET_CONFIG_MP}'!A:Z"
        ).execute()

        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID_MP,
            range=f"'{SHEET_CONFIG_MP}'!A1",
            valueInputOption='RAW',
            body={'values': body_values}
        ).execute()

        return jsonify({'success': True, 'message': 'Configuración MP guardada correctamente'})
    except Exception as e:
        print(f"Error en save_config_mp: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/mantenimientos/add', methods=['POST'])
def add_mantenimiento():
    try:
        data = request.json
        sheet_name = SHEET_HISTORIAL_MP

        sheet_data = get_all_data(sheet_name, SPREADSHEET_ID_MP)
        headers = sheet_data.get('headers', [])

        if not headers:
            return jsonify({'success': False, 'error': 'La hoja Historial Mantenimientos no tiene encabezados'}), 500

        row_values = []
        for header in headers:
            row_values.append(str(data.get(header, '')))

        success = add_row_to_sheet(sheet_name, row_values, SPREADSHEET_ID_MP)

        if success:
            return jsonify({'success': True, 'message': 'Mantenimiento agregado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al guardar en Google Sheets'}), 500
    except Exception as e:
        print(f"Error en add_mantenimiento: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/mantenimientos/update/<int:row_number>', methods=['POST'])
def update_mantenimiento(row_number):
    try:
        data = request.json
        sheet_name = SHEET_HISTORIAL_MP

        sheet_data = get_all_data(sheet_name, SPREADSHEET_ID_MP)
        headers = sheet_data.get('headers', [])

        row_values = []
        for header in headers:
            row_values.append(str(data.get(header, '')))

        success = update_row(sheet_name, row_number, row_values, SPREADSHEET_ID_MP)

        if success:
            return jsonify({'success': True, 'message': 'Mantenimiento actualizado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al actualizar'}), 500
    except Exception as e:
        print(f"Error en update_mantenimiento: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/mantenimientos/delete/<int:row_number>', methods=['DELETE'])
def delete_mantenimiento(row_number):
    try:
        success = delete_row(SHEET_HISTORIAL_MP, row_number, SPREADSHEET_ID_MP)

        if success:
            return jsonify({'success': True, 'message': 'Mantenimiento eliminado correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al eliminar'}), 500
    except Exception as e:
        print(f"Error en delete_mantenimiento: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

# ==================== PROGRAMACIÓN MANTENIMIENTO PREVENTIVO ====================

@app.route('/api/programacion/add', methods=['POST'])
def add_programacion():
    try:
        data = request.json
        sheet_name = SHEET_PROGRAMACION_MP

        ensure_sheet_exists(sheet_name, SPREADSHEET_ID_MP, headers=PROG_HEADERS)

        sheet_data = get_all_data(sheet_name, SPREADSHEET_ID_MP)
        headers = sheet_data.get('headers', [])

        if not headers:
            return jsonify({'success': False, 'error': 'La hoja Programación no tiene encabezados'}), 500

        marca_temporal = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        fecha_ultimo = data.get('FECHA ULTIMO MANTENIMIENTO', '')
        proximo_fecha = ''
        parsed = parse_date(fecha_ultimo)
        if parsed:
            try:
                proximo = parsed.replace(year=parsed.year + 1)
            except ValueError:
                proximo = parsed.replace(year=parsed.year + 1, day=28)
            proximo_fecha = proximo.strftime('%d/%m/%Y')

        tipo_reparacion = (data.get('TIPO REPARACIÓN') or '').strip()
        km_ultimo_str = (data.get('KM ULTIMO MANTENIMIENTO') or '').strip()
        politica = get_politica_por_tipo_reparacion(tipo_reparacion) if tipo_reparacion else 0
        proximo_km = ''
        if km_ultimo_str:
            try:
                km_ultimo = float(km_ultimo_str)
                proximo_km = str(int(km_ultimo + politica)) if politica else str(int(km_ultimo))
            except ValueError:
                proximo_km = ''

        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(marca_temporal)
            elif header == 'PROXIMO MANTENIMIENTO FECHA':
                row_values.append(proximo_fecha)
            elif header == 'PROXIMO MANTENIMIENTO KM':
                row_values.append(proximo_km)
            elif header == 'TIPO MANTENIMIENTO':
                row_values.append('PREVENTIVO')
            elif header == 'Estado':
                row_values.append('En Proceso')
            else:
                row_values.append(str(data.get(header, '')))

        success = add_row_to_sheet(sheet_name, row_values, SPREADSHEET_ID_MP)

        if success:
            return jsonify({'success': True, 'message': 'Programación agregada correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al guardar en Google Sheets'}), 500
    except Exception as e:
        print(f"Error en add_programacion: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/programacion/update/<int:row_number>', methods=['POST'])
def update_programacion(row_number):
    try:
        data = request.json
        sheet_name = SHEET_PROGRAMACION_MP

        sheet_data = get_all_data(sheet_name, SPREADSHEET_ID_MP)
        headers = sheet_data.get('headers', [])

        fecha_ultimo = data.get('FECHA ULTIMO MANTENIMIENTO', '')
        proximo_fecha = ''
        parsed = parse_date(fecha_ultimo)
        if parsed:
            try:
                proximo = parsed.replace(year=parsed.year + 1)
            except ValueError:
                proximo = parsed.replace(year=parsed.year + 1, day=28)
            proximo_fecha = proximo.strftime('%d/%m/%Y')

        tipo_reparacion = (data.get('TIPO REPARACIÓN') or '').strip()
        km_ultimo_str = (data.get('KM ULTIMO MANTENIMIENTO') or '').strip()
        politica = get_politica_por_tipo_reparacion(tipo_reparacion) if tipo_reparacion else 0
        proximo_km = ''
        if km_ultimo_str:
            try:
                km_ultimo = float(km_ultimo_str)
                proximo_km = str(int(km_ultimo + politica)) if politica else str(int(km_ultimo))
            except ValueError:
                proximo_km = ''

        current_row = None
        for row in sheet_data.get('rows', []):
            if row.get('_row_number') == row_number:
                current_row = row
                break
        marca_temporal = current_row.get('Marca Temporal', '') if current_row else ''
        estado_actual = current_row.get('Estado', 'En Proceso') if current_row else 'En Proceso'

        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(marca_temporal)
            elif header == 'PROXIMO MANTENIMIENTO FECHA':
                row_values.append(proximo_fecha)
            elif header == 'PROXIMO MANTENIMIENTO KM':
                row_values.append(proximo_km)
            elif header == 'TIPO MANTENIMIENTO':
                row_values.append('PREVENTIVO')
            elif header == 'Estado':
                row_values.append(str(data.get('Estado', estado_actual)))
            else:
                row_values.append(str(data.get(header, '')))

        success = update_row(sheet_name, row_number, row_values, SPREADSHEET_ID_MP)

        if success:
            return jsonify({'success': True, 'message': 'Programación actualizada correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al actualizar'}), 500
    except Exception as e:
        print(f"Error en update_programacion: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/programacion/delete/<int:row_number>', methods=['DELETE'])
def delete_programacion(row_number):
    try:
        success = delete_row(SHEET_PROGRAMACION_MP, row_number, SPREADSHEET_ID_MP)

        if success:
            return jsonify({'success': True, 'message': 'Programación eliminada correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al eliminar'}), 500
    except Exception as e:
        print(f"Error en delete_programacion: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/programacion/completar/<int:row_number>', methods=['POST'])
def completar_programacion(row_number):
    """
    Registra el mantenimiento en el Historial y resuelve la fila de Programación:
    - Si se vuelve a programar: SOBRESCRIBE esa misma fila con la nueva fecha/km (no crea una fila nueva).
    - Si NO se vuelve a programar: ELIMINA la fila de Programación.
    "Programacion Mantenimiento Preventivo" debe contener siempre solo tareas en proceso.
    Body esperado (JSON):
    {
        "gasto_pesos": "1234.56",
        "gasto_usd": "12.34",
        "volver_a_programar": true/false
    }
    """
    try:
        data = request.json or {}
        gasto_pesos = str(data.get('gasto_pesos', '') or '')
        gasto_usd = str(data.get('gasto_usd', '') or '')
        volver_a_programar = bool(data.get('volver_a_programar', False))

        # 1) Traer la fila de programación
        sheet_data = get_all_data(SHEET_PROGRAMACION_MP, SPREADSHEET_ID_MP)
        prog_row = None
        for row in sheet_data.get('rows', []):
            if row.get('_row_number') == row_number:
                prog_row = row
                break

        if not prog_row:
            return jsonify({'success': False, 'error': 'No se encontró la programación'}), 404

        # 2) Insertar en Historial Mantenimientos (el gasto se carga acá, no en Programación)
        hist_data = get_all_data(SHEET_HISTORIAL_MP, SPREADSHEET_ID_MP)
        hist_headers = hist_data.get('headers', [])

        if not hist_headers:
            hist_headers = HIST_HEADERS

        hoy = datetime.now()
        hoy_str = hoy.strftime('%d/%m/%Y')

        nuevo_hist = {
            'FECHA': hoy_str,
            'PATENTE': prog_row.get('PATENTE', ''),
            'TIPO MANTENIMIENTO': prog_row.get('TIPO MANTENIMIENTO', ''),
            'TIPO REPARACIÓN': prog_row.get('TIPO REPARACIÓN', ''),
            'CONCEPTO GENERAL': prog_row.get('TIPO REPARACIÓN', ''),
            'DETALLE REPARACIÓN': prog_row.get('DETALLE REPARACIÓN', ''),
            'KM REALIZADO': prog_row.get('PROXIMO MANTENIMIENTO KM', '') or prog_row.get('KM ULTIMO MANTENIMIENTO', ''),
            'KM PROXIMO CAMBIO': '',
            'PROVEEDOR': '',
            'GASTO en PESOS': gasto_pesos,
            'GASTO en USD': gasto_usd,
            'OBSERVACIONES': 'Generado automáticamente desde Programación (Tarea Completa)'
        }

        tipo_rep = (nuevo_hist['TIPO REPARACIÓN'] or '').strip()
        politica = get_politica_por_tipo_reparacion(tipo_rep) if tipo_rep else 0
        km_realizado_str = str(nuevo_hist['KM REALIZADO'] or '').strip()
        if km_realizado_str:
            try:
                km_realizado = float(km_realizado_str)
                nuevo_hist['KM PROXIMO CAMBIO'] = str(int(km_realizado + politica)) if politica else str(int(km_realizado))
            except ValueError:
                pass

        hist_row_values = []
        for header in hist_headers:
            hist_row_values.append(str(nuevo_hist.get(header, '')))

        success_hist = add_row_to_sheet(SHEET_HISTORIAL_MP, hist_row_values, SPREADSHEET_ID_MP)

        # 3) Resolver la fila de Programación
        nueva_prog_ok = True
        if volver_a_programar:
            # Se SOBRESCRIBE la misma fila con la nueva fecha/km (no se agrega una fila nueva)
            patente = prog_row.get('PATENTE', '')
            camion_t2_data = get_all_data('Camion T2', SPREADSHEET_ID)
            odometro_actual = ''
            for row in camion_t2_data.get('rows', []):
                if (row.get('PATENTE') or '').strip() == patente:
                    odometro_actual = (row.get('ODOMETRO') or '').strip()
                    break

            km_para_nueva = odometro_actual or nuevo_hist['KM REALIZADO']

            nueva_prog = {
                'PATENTE': patente,
                'FECHA ULTIMO MANTENIMIENTO': hoy_str,
                'KM ULTIMO MANTENIMIENTO': km_para_nueva,
                'TIPO MANTENIMIENTO': 'PREVENTIVO',
                'TIPO REPARACIÓN': prog_row.get('TIPO REPARACIÓN', ''),
                'DETALLE REPARACIÓN': prog_row.get('DETALLE REPARACIÓN', '')
            }

            try:
                prog_headers = sheet_data.get('headers', [])
                nueva_row_values = []
                for header in prog_headers:
                    if header == 'Marca Temporal':
                        nueva_row_values.append(hoy.strftime('%Y-%m-%d %H:%M:%S'))
                    elif header == 'PROXIMO MANTENIMIENTO FECHA':
                        try:
                            prox = hoy.replace(year=hoy.year + 1)
                        except ValueError:
                            prox = hoy.replace(year=hoy.year + 1, day=28)
                        nueva_row_values.append(prox.strftime('%d/%m/%Y'))
                    elif header == 'PROXIMO MANTENIMIENTO KM':
                        km_ult = str(km_para_nueva or '').strip()
                        if km_ult:
                            try:
                                km_ult_f = float(km_ult)
                                nueva_row_values.append(str(int(km_ult_f + politica)) if politica else str(int(km_ult_f)))
                            except ValueError:
                                nueva_row_values.append('')
                        else:
                            nueva_row_values.append('')
                    elif header == 'TIPO MANTENIMIENTO':
                        nueva_row_values.append('PREVENTIVO')
                    elif header == 'Estado':
                        nueva_row_values.append('En Proceso')
                    else:
                        nueva_row_values.append(str(nueva_prog.get(header, '')))

                nueva_prog_ok = update_row(SHEET_PROGRAMACION_MP, row_number, nueva_row_values, SPREADSHEET_ID_MP)
            except Exception as e:
                print(f"Error al reprogramar (sobrescribir fila): {e}")
                nueva_prog_ok = False
        else:
            # La tarea no se vuelve a repetir: se elimina la fila de Programación
            nueva_prog_ok = delete_row(SHEET_PROGRAMACION_MP, row_number, SPREADSHEET_ID_MP)

        return jsonify({
            'success': True,
            'message': 'Tarea completada',
            'historial_ok': success_hist,
            'nueva_programacion_ok': nueva_prog_ok
        })

    except Exception as e:
        print(f"Error en completar_programacion: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
