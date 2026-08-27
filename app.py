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

def get_all_data(sheet_name):
    """Obtiene todos los datos de una hoja incluyendo headers y filas"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
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

def update_row(sheet_name, row_number, values):
    """Actualiza una fila completa en Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        data = get_all_data(sheet_name)
        num_columns = len(data.get('headers', []))
        
        while len(values) < num_columns:
            values.append('')
        
        last_col = chr(64 + num_columns) if num_columns <= 26 else 'Z'
        
        body = {
            'values': [values[:num_columns]]
        }
        
        result = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_name}'!A{row_number}:{last_col}{row_number}",
            valueInputOption='RAW',
            body=body
        ).execute()
        
        return True
        
    except HttpError as err:
        print(f"Error updating row: {err}")
        return False

def delete_row(sheet_name, row_number):
    """Elimina una fila de Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID
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
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()
        
        return True
        
    except HttpError as err:
        print(f"Error deleting row: {err}")
        return False

def add_row_to_sheet(sheet_name, values):
    """Agrega una nueva fila a Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        data = get_all_data(sheet_name)
        num_columns = len(data.get('headers', []))
        
        while len(values) < num_columns:
            values.append('')
        
        body = {
            'values': [values[:num_columns]]
        }
        
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_name}'!A:Z",
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        
        return True
        
    except HttpError as err:
        print(f"Error adding row: {err}")
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

def is_date_header(header):
    """Determina si una columna corresponde a una fecha/vencimiento."""
    if not header:
        return False
    header_upper = header.upper()
    return (
        'FECHA' in header_upper or
        'VENCIMIENTO' in header_upper or
        'VENC' in header_upper
    )

def get_date_photo_requirements(headers):
    """Devuelve pares (columna_fecha, columna_foto) para fechas que requieren respaldo."""
    return [(header, f'{header}_FOTO') for header in headers if is_date_header(header)]


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
        module = request.form.get('module', 'master')

        # En Control de Documentación, toda fecha cargada debe tener foto de respaldo.
        if module == 'docs':
            for date_header, photo_header in get_date_photo_requirements(headers):
                date_value = form_data.get(date_header, '').strip()
                if date_value:
                    if photo_header not in headers:
                        return jsonify({
                            'success': False,
                            'error': f'No existe la columna "{photo_header}" en Google Sheets para guardar la foto de "{date_header}".'
                        }), 400

                    photo_key = f'foto_{date_header}'
                    photo = request.files.get(photo_key)
                    if not photo or not photo.filename:
                        return jsonify({
                            'success': False,
                            'error': f'La fecha "{date_header}" requiere obligatoriamente una foto de respaldo.'
                        }), 400

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
                    # Extraer el nombre del documento de la clave (ej: foto_VENC_VTV)
                    doc_name = key.replace('foto_', '')
                    
                    # Generar nombre según la pestaña
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
                        # Buscar la columna correspondiente (agregar _FOTO al nombre)
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
        data = request.json or {}
        module = data.get('__module', 'master')
        
        sheet_data = get_all_data(sheet_name)
        headers = sheet_data.get('headers', [])

        # En Control de Documentación no se permite modificar/cargar una fecha sin su foto.
        if module == 'docs':
            current_row = next((r for r in sheet_data.get('rows', []) if r.get('_row_number') == row_number), None)
            if current_row is None:
                return jsonify({'success': False, 'error': 'No se encontró el documento.'}), 404

            for date_header, photo_header in get_date_photo_requirements(headers):
                new_date = str(data.get(date_header, '') or '').strip()
                old_date = str(current_row.get(date_header, '') or '').strip()
                old_photo = str(current_row.get(photo_header, '') or '').strip()

                if new_date and new_date != old_date:
                    return jsonify({
                        'success': False,
                        'error': f'La fecha "{date_header}" requiere obligatoriamente una foto de respaldo nueva.'
                    }), 400
                if new_date and not old_photo:
                    return jsonify({
                        'success': False,
                        'error': f'El documento tiene cargada la fecha "{date_header}" pero no tiene foto de respaldo.'
                    }), 400
        
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
        # Obtener datos del formulario
        form_data = {}
        for key in request.form:
            form_data[key] = request.form[key]
        module = form_data.get('module', 'master')

        # Obtener datos actuales antes de procesar fotos, para validar cambios de fechas.
        sheet_data = get_all_data(sheet_name)
        headers = sheet_data.get('headers', [])
        current_row = None
        for row in sheet_data.get('rows', []):
            if row.get('_row_number') == row_number:
                current_row = row
                break

        if current_row is None:
            return jsonify({'success': False, 'error': 'No se encontró el documento.'}), 404

        # En Control de Documentación, toda fecha nueva/modificada debe tener foto.
        if module == 'docs':
            for date_header, photo_header in get_date_photo_requirements(headers):
                new_date = str(form_data.get(date_header, '') or '').strip()
                old_date = str(current_row.get(date_header, '') or '').strip()
                old_photo = str(current_row.get(photo_header, '') or '').strip()
                date_changed = new_date != old_date

                if new_date and date_changed:
                    photo_key = f'foto_{date_header}'
                    photo = request.files.get(photo_key)
                    if not photo or not photo.filename:
                        return jsonify({
                            'success': False,
                            'error': f'La fecha "{date_header}" requiere obligatoriamente una foto de respaldo nueva.'
                        }), 400

                if new_date and not old_photo and not request.files.get(f'foto_{date_header}'):
                    return jsonify({
                        'success': False,
                        'error': f'El documento tiene la fecha "{date_header}" pero no tiene foto de respaldo.'
                    }), 400

        # Procesar múltiples fotos
        for key in request.files:
            if key.startswith('foto_'):
                foto = request.files[key]
                if foto.filename != '':
                    # Extraer el nombre del documento de la clave (ej: foto_VENC_VTV)
                    doc_name = key.replace('foto_', '')
                    
                    # Generar nombre según la pestaña
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
                        # Buscar la columna correspondiente (agregar _FOTO al nombre)
                        foto_header = f"{doc_name}_FOTO"
                        form_data[foto_header] = drive_url
        
        # Construir valores de fila
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

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
