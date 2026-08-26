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
        
        if 'foto' in request.files and request.files['foto'].filename != '':
            foto = request.files['foto']
            
            # Generar nombre según la pestaña
            if sheet_name in ['Camion T1', 'Camion T2']:
                identificador = form_data.get('PATENTE', 'SIN_PATENTE')
                documento = form_data.get('VENC VTV', 'DOCUMENTO')
            elif sheet_name == 'Autoelevadores':
                identificador = form_data.get('CODIGO DE AE', 'SIN_CODIGO')
                documento = form_data.get('VENC SEGURO', 'DOCUMENTO')
            elif sheet_name == 'Choferes y Ayudantes':
                identificador = form_data.get('APELLIDO Y NOMBRE', 'SIN_NOMBRE')
                documento = form_data.get('VENCIMIENTO REGISTRO', 'DOCUMENTO')
            else:
                identificador = 'DOCUMENTO'
                documento = 'FOTO'
            
            file_extension = os.path.splitext(foto.filename)[1]
            filename = f"{identificador} - {documento}{file_extension}"
            file_content = foto.read()
            
            file_id = upload_file_to_drive(
                file_content, 
                filename, 
                DRIVE_FOLDER_ID
            )
            
            if file_id:
                drive_url = f"https://drive.google.com/file/d/{file_id}/view"
                # Buscar columna de foto o link
                foto_col = None
                for i, header in enumerate(headers):
                    if header.lower() in ['link', 'url', 'foto', 'imagen']:
                        foto_col = i
                        break
                if foto_col is not None:
                    row_values[foto_col] = drive_url
        
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
    """Actualiza un documento existente"""
    try:
        data = request.json if request.json else {}
        
        # Obtener datos del formulario si es multipart
        if request.files:
            form_data = {}
            for key in request.form:
                form_data[key] = request.form[key]
            
            # Procesar foto si se subió
            if 'foto' in request.files and request.files['foto'].filename != '':
                foto = request.files['foto']
                
                # Generar nombre según la pestaña
                if sheet_name in ['Camion T1', 'Camion T2']:
                    identificador = form_data.get('PATENTE', 'SIN_PATENTE')
                    documento = form_data.get('VENC VTV', 'DOCUMENTO')
                elif sheet_name == 'Autoelevadores':
                    identificador = form_data.get('CODIGO DE AE', 'SIN_CODIGO')
                    documento = form_data.get('VENC SEGURO', 'DOCUMENTO')
                elif sheet_name == 'Choferes y Ayudantes':
                    identificador = form_data.get('APELLIDO Y NOMBRE', 'SIN_NOMBRE')
                    documento = form_data.get('VENCIMIENTO REGISTRO', 'DOCUMENTO')
                else:
                    identificador = 'DOCUMENTO'
                    documento = 'FOTO'
                
                file_extension = os.path.splitext(foto.filename)[1]
                filename = f"{identificador} - {documento}{file_extension}"
                file_content = foto.read()
                
                file_id = upload_file_to_drive(
                    file_content, 
                    filename, 
                    DRIVE_FOLDER_ID
                )
                
                if file_id:
                    drive_url = f"https://drive.google.com/file/d/{file_id}/view"
                    form_data['FOTO_URL'] = drive_url
                    form_data['LINK'] = drive_url
                    form_data['URL'] = drive_url
            
            # Fusionar datos del formulario con los datos JSON
            data.update(form_data)
        
        sheet_data = get_all_data(sheet_name)
        headers = sheet_data.get('headers', [])
        
        # Obtener fila actual para preservar datos no enviados
        current_row = None
        for row in sheet_data.get('rows', []):
            if row.get('_row_number') == row_number:
                current_row = row
                break
        
        row_values = []
        for header in headers:
            if header == 'Marca Temporal':
                row_values.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            elif header in data and data[header] is not None:
                row_values.append(data[header])
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
        print(f"Error en update_document: {e}")
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
