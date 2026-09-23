import os
import io
import json
from datetime import datetime
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

# Mantenimientos spreadsheet
SPREADSHEET_ID_MP = '1ZCgxrjf84jAlc6F8KsiuvFXkZ2qwbUEHfDAHwnfDhzg'
SHEET_HISTORIAL_MP = 'Historial Mantenimientos'
SHEET_CONFIG_MP = 'Configuracion MP'
SHEET_CONFIG_DROPDOWNS = 'Configuracion Desplegables'

SHEETS = {
    'Camion T1': 'Camion T1',
    'Camion T2': 'Camion T2',
    'Autoelevadores': 'Autoelevadores',
    'Choferes y Ayudantes': 'Choferes y Ayudantes'
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

def ensure_sheet_exists(sheet_name, spreadsheet_id=None):
    """Crea la hoja si no existe"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        sid = spreadsheet_id or SPREADSHEET_ID
        
        spreadsheet = service.spreadsheets().get(spreadsheetId=sid).execute()
        for s in spreadsheet.get('sheets', []):
            if s['properties']['title'] == sheet_name:
                return True
        
        requests = [{
            'addSheet': {
                'properties': {'title': sheet_name}
            }
        }]
        sheet.batchUpdate(spreadsheetId=sid, body={'requests': requests}).execute()
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

@app.route('/')
def index():
    """Página principal"""
    return render_template('index.html', sheets=SHEETS)

@app.route('/api/sheet/<sheet_name>')
def get_sheet(sheet_name):
    """API para obtener datos de una hoja"""
    data = get_all_data(sheet_name)
    return jsonify(data)

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
        
        # Procesar múltiples fotos
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

@app.route('/api/mantenimientos/sheet/<sheet_name>')
def get_mantenimiento_sheet(sheet_name):
    """API para obtener datos de una hoja del spreadsheet de mantenimientos"""
    data = get_all_data(sheet_name, SPREADSHEET_ID_MP)
    return jsonify(data)

@app.route('/api/mantenimientos/config', methods=['GET'])
def get_mantenimiento_config():
    """Obtiene la configuración completa de mantenimientos: dropdowns y políticas MP"""
    try:
        ensure_sheet_exists(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)
        
        # Leer hoja de configuración de dropdowns
        dropdown_data = get_all_data(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)
        
        # Leer hoja Configuracion MP (TIPO MANTENIMIENTO, TIPO REPARACIÓN, POLITICA)
        config_mp_data = get_all_data(SHEET_CONFIG_MP, SPREADSHEET_ID_MP)
        
        # Obtener headers reales de Historial Mantenimientos
        hist_data = get_all_data(SHEET_HISTORIAL_MP, SPREADSHEET_ID_MP)
        
        return jsonify({
            'success': True,
            'dropdowns': dropdown_data,
            'config_mp': config_mp_data,
            'historial_headers': hist_data.get('headers', []),
            'historial_rows': hist_data.get('rows', [])
        })
    except Exception as e:
        print(f"Error en get_mantenimiento_config: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/mantenimientos/dropdowns', methods=['POST'])
def save_mantenimiento_dropdowns():
    """Guarda la configuración de desplegables"""
    try:
        data = request.json
        # data = { 'dropdowns': { 'PATENTE': ['AAA111', 'BBB222'], 'TIPO MANTENIMIENTO': [...], ... } }
        
        dropdowns = data.get('dropdowns', {})
        
        ensure_sheet_exists(SHEET_CONFIG_DROPDOWNS, SPREADSHEET_ID_MP)
        
        # Reconstruir la hoja completa: una columna por cada campo con desplegables
        # Formato: primera fila = nombre del campo, siguientes filas = opciones
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
        
        # Limpiar toda la hoja primero
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
        rows = data.get('rows', [])  # [{'TIPO MANTENIMIENTO': 'Preventivo', 'TIPO REPARACIÓN': '...', 'POLITICA': 10000}, ...]
        
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

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
