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

# Headers de la hoja de Programación (se crean automáticamente)
PROG_HEADERS = [
    'Marca Temporal',
    'PATENTE',
    'FECHA ULTIMO MANTENIMIENTO',
    'PROXIMO MANTENIMIENTO FECHA',
    'KM ULTIMO MANTENIMIENTO',
    'TIPO MANTENIMIENTO',
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

# Configuración de columnas de fecha por hoja (para detección de vencimientos)
DATE_COLUMNS_CONFIG = {
    'Camion T1': ['VENC VTV', 'VENC SEGURO', 'SENASA', 'LICENCIA DE CONDUCIR', 'PAGO MONOTRIBUTO',
                  'POLIZA DE SEGUROS', 'SEGURO DE ACCIDENTES PERSONALES', 'CLAUSULA DE NO REPETICION',
                  'INDUCCION DE CMQ', 'CAPACITACION DE MANEJO DEFENSIVO'],
    'Camion T2': ['VENC VTV', 'VENC SEGURO', 'UTA', 'EXTINTOR', 'BOTIQUIN'],
    'Autoelevadores': ['VENC SEGURO', 'EXTINTOR'],
    'Choferes y Ayudantes': ['VENCIMIENTO REGISTRO', 'LIBRETA SANITARIA']
}

# Campo identificador por hoja
ID_FIELD_CONFIG = {
    'Camion T1': 'PATENTE',
    'Camion T2': 'PATENTE',
    'Autoelevadores': 'CODIGO DE AE',
    'Choferes y Ayudantes': 'APELLIDO Y NOMBRE'
}

app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-key-12345')
CORS(app)

def get_google_creds():
    """Obtiene credenciales de Service Account para Google APIs"""
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
        return creds
    except Exception as e:
        print(f"Error al obtener credenciales: {e}")
        raise

def get_all_data(sheet_name, spreadsheet_id=None):
    """Obtiene todos los datos de una hoja incluyendo headers y filas"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
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
        print(f"Error getting sheet data: {err}")
        return {'headers': [], 'rows': []}

def update_row(sheet_name, row_number, values, spreadsheet_id=None):
    """Actualiza una fila completa en Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        data = get_all_data(sheet_name, sid)
        num_columns = len(data.get('headers', []))

        while len(values) < num_columns:
            values.append('')

        last_col = chr(64 + num_columns) if num_columns <= 26 else 'Z'

        body = {
            'values': [values[:num_columns]]
        }

        result = sheet.values().update(
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
    """Actualiza celdas específicas de una fila en Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
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

        body = {
            'valueInputOption': 'RAW',
            'data': requests
        }

        sheet.values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()

        return True

    except HttpError as err:
        print(f"Error updating row partial: {err}")
        return False

def delete_row(sheet_name, row_number, spreadsheet_id=None):
    """Elimina una fila de Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        spreadsheet = service.spreadsheets().get(
            spreadsheetId=sid
        ).execute()

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

        result = sheet.batchUpdate(
            spreadsheetId=sid,
            body=body
        ).execute()

        return True

    except HttpError as err:
        print(f"Error deleting row: {err}")
        return False

def add_row_to_sheet(sheet_name, values, spreadsheet_id=None):
    """Agrega una nueva fila a Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        data = get_all_data(sheet_name, sid)
        num_columns = len(data.get('headers', []))

        while len(values) < num_columns:
            values.append('')

        body = {
            'values': [values[:num_columns]]
        }

        result = sheet.values().append(
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
    """Crea la hoja si no existe y opcionalmente escribe headers"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        sid = spreadsheet_id or SPREADSHEET_ID

        spreadsheet = service.spreadsheets().get(spreadsheetId=sid).execute()
        exists = False
        for s in spreadsheet.get('sheets', []):
            if s['properties']['title'] == sheet_name:
                exists = True
                break

        if not exists:
            requests = [{
                'addSheet': {
                    'properties': {'title': sheet_name}
                }
            }]
            sheet.batchUpdate(spreadsheetId=sid, body={'requests': requests}).execute()
            if headers:
                sheet.values().update(
                    spreadsheetId=sid,
                    range=f"'{sheet_name}'!A1",
                    valueInputOption='RAW',
                    body={'values': [headers]}
                ).execute()
        return True
    except HttpError as err:
        print(f"Error ensuring sheet: {err}")
        return False

def upload_file_to_drive(file_content, filename, folder_id):
    """Sube un archivo a Google Drive"""
    try:
        creds = get_google_creds()
        service = build('drive', 'v3', credentials=creds)

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
    """Parsea una fecha en formatos comunes (dd/mm/yyyy, yyyy-mm-dd)"""
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

def check_and_register_expirations():
    """Recorre las hojas de documentos, detecta vencimientos dentro de 30 días
    y los registra en la hoja de historial si no existen ya.
    Optimizado: usa batch append para evitar timeout."""
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
                creds = get_google_creds()
                service = build('sheets', 'v4', credentials=creds)
                sheet = service.spreadsheets()

                body = {
                    'values': new_rows
                }

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

@app.route('/')
def index():
    """Página principal"""
    return render_template('index.html', sheets=SHEETS)

@app.route('/api/sheet/<sheet_name>')
def get_sheet(sheet_name):
    """API para obtener datos de una hoja"""
    data = get_all_data(sheet_name)
    return jsonify(data)

@app.route('/api/history')
def get_history():
    """API para obtener el historial de vencimientos (ejecuta detección automática)"""
    try:
        check_and_register_expirations()
        data = get_all_data(HISTORY_SHEET)
        return jsonify(data)
    except Exception as e:
        print(f"Error en get_history: {e}")
        return jsonify({'headers': [], 'rows': [], 'error': str(e)}), 500

@app.route('/api/history/update_status/<int:row_number>', methods=['POST'])
def update_history_status(row_number):
    """Actualiza el estado de un registro del historial"""
    try:
        data = request.json
        new_status = data.get('estado', '')

        if new_status not in ['En Proceso', 'Completo', 'Vencido', 'Notificado']:
            return jsonify({
                'success': False,
                'error': 'Estado no válido'
            }), 400

        success = update_row_partial(HISTORY_SHEET, row_number, {'ESTADO': new_status})

        if success:
            return jsonify({
                'success': True,
                'message': 'Estado actualizado correctamente'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Error al actualizar el estado'
            }), 500

    except Exception as e:
        print(f"Error en update_history_status: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/history/delete/<int:row_number>', methods=['DELETE'])
def delete_history_row(row_number):
    """Elimina un registro del historial"""
    try:
        success = delete_row(HISTORY_SHEET, row_number)

        if success:
            return jsonify({
                'success': True,
                'message': 'Registro eliminado correctamente'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Error al eliminar el registro'
            }), 500

    except Exception as e:
        print(f"Error en delete_history_row: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/add', methods=['POST'])
def add_document():
    """API para agregar un nuevo documento"""
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

                    file_id = upload_file_to_drive(
                        file_content,
                        filename,
                        DRIVE_FOLDER_ID
                    )

                    if file_id:
                        drive_url = f"https://drive.google.com/file/d/{file_id}/view"
                        foto_header = f"{doc_name}_FOTO"
                        for i, header in enumerate(headers):
                            if header == foto_header:
                                row_values[i] = drive_url
                                break

        success = add_row_to_sheet(sheet_name, row_values)

        if success:
            return jsonify({
                'success': True,
                'message': 'Documento agregado correctamente'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Error al guardar en Google Sheets'
            }), 500

    except Exception as e:
        print(f"Error en add_document: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/update/<sheet_name>/<int:row_number>', methods=['POST'])
def update_document(sheet_name, row_number):
    """Actualiza un documento existente (solo datos)"""
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
            return jsonify({
                'success': True,
                'message': 'Documento actualizado correctamente'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Error al actualizar en Google Sheets'
            }), 500

    except Exception as e:
        print(f"Error en update_document: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/update_with_photo/<sheet_name>/<int:row_number>', methods=['POST'])
def update_document_with_photo(sheet_name, row_number):
    """Actualiza un documento existente con múltiples fotos"""
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

                    file_id = upload_file_to_drive(
                        file_content,
                        filename,
                        DRIVE_FOLDER_ID
                    )

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
            return jsonify({
                'success': True,
                'message': 'Documento actualizado correctamente'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Error al actualizar en Google Sheets'
            }), 500

    except Exception as e:
        print(f"Error en update_document_with_photo: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/delete/<sheet_name>/<int:row_number>', methods=['DELETE'])
def delete_document(sheet_name, row_number):
    """Elimina un documento"""
    try:
        success = delete_row(sheet_name, row_number)

        if success:
            return jsonify({
                'success': True,
                'message': 'Documento eliminado correctamente'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Error al eliminar el documento'
            }), 500

    except Exception as e:
        print(f"Error en delete_document: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

# ==================== MANTENIMIENTOS ====================

@app.route('/api/mantenimientos/config', methods=['GET'])
def get_mantenimiento_config():
    """Obtiene la configuración completa de mantenimientos: dropdowns y políticas MP"""
    try:
        ensure_sheet_exists(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)

        dropdown_data = get_all_data(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)
        config_mp_data = get_all_data(SHEET_CONFIG_MP, SPREADSHEET_ID_MP)
        hist_data = get_all_data(SHEET_HISTORIAL_MP, SPREADSHEET_ID_MP)

        # PATENTES disponibles desde Camion T2 del spreadsheet original
        # También obtenemos el ODOMETRO por patente (para KM ULTIMO MANTENIMIENTO)
        camion_t2_data = get_all_data('Camion T2', SPREADSHEET_ID)
        patentes_camion_t2 = []
        odometro_por_patente = {}
        for row in camion_t2_data.get('rows', []):
            patente = (row.get('PATENTE') or '').strip()
            if patente:
                if patente not in patentes_camion_t2:
                    patentes_camion_t2.append(patente)
                # Guardar odómetro (si hay múltiples filas con misma patente, guarda la última encontrada)
                odometro = (row.get('ODOMETRO') or '').strip()
                if odometro:
                    odometro_por_patente[patente] = odometro
        patentes_camion_t2.sort()

        # Programación de mantenimiento preventivo
        ensure_sheet_exists(
            SHEET_PROGRAMACION_MP,
            SPREADSHEET_ID_MP,
            headers=PROG_HEADERS
        )
        programacion_data = get_all_data(SHEET_PROGRAMACION_MP, SPREADSHEET_ID_MP)

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
        print(f"Error en get_mantenimiento_config: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/mantenimientos/dropdowns', methods=['POST'])
def save_mantenimiento_dropdowns():
    """Guarda la configuración de desplegables"""
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

        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
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
    """Guarda la configuración MP (TIPO MANTENIMIENTO, TIPO REPARACIÓN, POLITICA)"""
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

        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
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
    """Agrega un nuevo registro de mantenimiento"""
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
    """Actualiza un registro de mantenimiento"""
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
    """Elimina un registro de mantenimiento"""
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
    """Agrega un nuevo registro de programación de mantenimiento preventivo"""
    try:
        data = request.json
        sheet_name = SHEET_PROGRAMACION_MP

        ensure_sheet_exists(
            sheet_name,
            SPREADSHEET_ID_MP,
            headers=PROG_HEADERS
        )

        sheet_data = get_all_data(sheet_name, SPREADSHEET_ID_MP)
        headers = sheet_data.get('headers', [])

        if not headers:
            return jsonify({'success': False, 'error': 'La hoja Programación no tiene encabezados'}), 500

        # Marca Temporal automática
        marca_temporal = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Calcular PROXIMO MANTENIMIENTO FECHA = FECHA ULTIMO MANTENIMIENTO + 1 año
        fecha_ultimo = data.get('FECHA ULTIMO MANTENIMIENTO', '')
        proximo_fecha = ''
        parsed = parse_date(fecha_ultimo)
        if parsed:
            try:
                proximo = parsed.replace(year=parsed.year + 1)
            except ValueError:
                # 29 de febrero en año no bisiesto -> 28 de febrero
                proximo = parsed.replace(year=parsed.year + 1, day=28)
            proximo_fecha = proximo.strftime('%d/%m/%Y')

        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(marca_temporal)
            elif header == 'PROXIMO MANTENIMIENTO FECHA':
                row_values.append(proximo_fecha)
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
    """Actualiza un registro de programación"""
    try:
        data = request.json
        sheet_name = SHEET_PROGRAMACION_MP

        sheet_data = get_all_data(sheet_name, SPREADSHEET_ID_MP)
        headers = sheet_data.get('headers', [])

        # Recalcular PROXIMO MANTENIMIENTO FECHA = FECHA ULTIMO MANTENIMIENTO + 1 año
        fecha_ultimo = data.get('FECHA ULTIMO MANTENIMIENTO', '')
        proximo_fecha = ''
        parsed = parse_date(fecha_ultimo)
        if parsed:
            try:
                proximo = parsed.replace(year=parsed.year + 1)
            except ValueError:
                proximo = parsed.replace(year=parsed.year + 1, day=28)
            proximo_fecha = proximo.strftime('%d/%m/%Y')

        # Preservar Marca Temporal original (no la modificamos al editar)
        current_row = None
        for row in sheet_data.get('rows', []):
            if row.get('_row_number') == row_number:
                current_row = row
                break
        marca_temporal = current_row.get('Marca Temporal', '') if current_row else ''

        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(marca_temporal)
            elif header == 'PROXIMO MANTENIMIENTO FECHA':
                row_values.append(proximo_fecha)
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
    """Elimina un registro de programación"""
    try:
        success = delete_row(SHEET_PROGRAMACION_MP, row_number, SPREADSHEET_ID_MP)

        if success:
            return jsonify({'success': True, 'message': 'Programación eliminada correctamente'})
        else:
            return jsonify({'success': False, 'error': 'Error al eliminar'}), 500
    except Exception as e:
        print(f"Error en delete_programacion: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
