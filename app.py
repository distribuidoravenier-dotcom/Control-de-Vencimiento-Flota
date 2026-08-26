import os
import io
import json
import base64
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, url_for, flash
from flask_cors import CORS
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

app = Flask(__name__)
app.config.from_object('config.Config')
CORS(app)

# Scopes necesarios
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

def get_google_creds():
    """Obtiene credenciales de Google para API"""
    creds = None
    
    # El archivo token.json almacena las credenciales del usuario
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # Si no hay credenciales válidas, pedir login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                app.config['CREDENTIALS_FILE'], SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Guardar credenciales para próximo uso
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def get_sheet_data(sheet_name):
    """Obtiene datos de una hoja específica"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        result = sheet.values().get(
            spreadsheetId=app.config['SPREADSHEET_ID'],
            range=f"'{sheet_name}'!A:F"
        ).execute()
        
        values = result.get('values', [])
        if not values:
            return []
        
        # Obtener headers
        headers = values[0] if values else []
        data = []
        
        # Procesar datos
        for row in values[1:]:
            row_data = {}
            for i, header in enumerate(headers):
                if i < len(row):
                    row_data[header] = row[i]
                else:
                    row_data[header] = ''
            data.append(row_data)
        
        return data
        
    except HttpError as err:
        print(f"Error getting sheet data: {err}")
        return []

def update_sheet_cell(sheet_name, row, col, value):
    """Actualiza una celda específica en Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        # Convertir columna a letra (1=A, 2=B, etc)
        col_letter = chr(64 + col)
        
        body = {
            'values': [[value]]
        }
        
        result = sheet.values().update(
            spreadsheetId=app.config['SPREADSHEET_ID'],
            range=f"'{sheet_name}'!{col_letter}{row}",
            valueInputOption='RAW',
            body=body
        ).execute()
        
        return True
        
    except HttpError as err:
        print(f"Error updating cell: {err}")
        return False

def add_row_to_sheet(sheet_name, values):
    """Agrega una nueva fila a Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        body = {
            'values': [values]
        }
        
        result = sheet.values().append(
            spreadsheetId=app.config['SPREADSHEET_ID'],
            range=f"'{sheet_name}'!A:F",
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
        
        # Crear el archivo en Drive
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
    return render_template('index.html', sheets=app.config['SHEETS'])

@app.route('/api/sheet/<sheet_name>')
def get_sheet(sheet_name):
    """API para obtener datos de una hoja"""
    data = get_sheet_data(sheet_name)
    return jsonify(data)

@app.route('/api/add', methods=['POST'])
def add_document():
    """API para agregar un nuevo documento"""
    try:
        # Obtener datos del formulario
        sheet_name = request.form.get('sheet_name')
        patente = request.form.get('patente')
        documento = request.form.get('documento')
        fecha_vencimiento = request.form.get('fecha_vencimiento')
        
        # Validar que los campos obligatorios estén completos
        if not all([sheet_name, patente, documento, fecha_vencimiento]):
            return jsonify({
                'success': False, 
                'error': 'Todos los campos son obligatorios'
            }), 400
        
        # Validar fecha
        try:
            datetime.strptime(fecha_vencimiento, '%Y-%m-%d')
        except ValueError:
            return jsonify({
                'success': False, 
                'error': 'Formato de fecha inválido'
            }), 400
        
        # Procesar foto
        if 'foto' not in request.files:
            return jsonify({
                'success': False, 
                'error': 'La foto es obligatoria'
            }), 400
        
        foto = request.files['foto']
        if foto.filename == '':
            return jsonify({
                'success': False, 
                'error': 'Debe seleccionar una foto'
            }), 400
        
        # Generar nombre de archivo: "patente - documento.ext"
        file_extension = os.path.splitext(foto.filename)[1]
        filename = f"{patente} - {documento}{file_extension}"
        
        # Leer contenido del archivo
        file_content = foto.read()
        
        # Subir a Google Drive
        file_id = upload_file_to_drive(
            file_content, 
            filename, 
            app.config['DRIVE_FOLDER_ID']
        )
        
        if not file_id:
            return jsonify({
                'success': False, 
                'error': 'Error al subir la foto a Google Drive'
            }), 500
        
        # Crear URL de Google Drive
        drive_url = f"https://drive.google.com/file/d/{file_id}/view"
        
        # Agregar fila a Google Sheets
        row_values = [
            patente,
            documento,
            fecha_vencimiento,
            drive_url,
            '',  # Campo para observaciones
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Fecha de carga
        ]
        
        success = add_row_to_sheet(sheet_name, row_values)
        
        if success:
            return jsonify({
                'success': True,
                'message': 'Documento agregado correctamente',
                'file_id': file_id,
                'drive_url': drive_url
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

@app.route('/api/update/<sheet_name>/<row>', methods=['POST'])
def update_document(sheet_name, row):
    """Actualiza un documento existente"""
    try:
        col = int(request.form.get('col', 0))
        value = request.form.get('value', '')
        
        success = update_sheet_cell(sheet_name, int(row), col, value)
        
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
        return jsonify({
            'success': False, 
            'error': str(e)
        }), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5000)))