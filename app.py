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
            range=f"'{sheet_name}'!A:G"
        ).execute()
        
        values = result.get('values', [])
        if not values:
            return {'headers': [], 'rows': []}
        
        headers = values[0] if values else []
        rows = []
        
        for i, row in enumerate(values[1:], start=2):  # start=2 porque la fila 1 es header
            row_data = {}
            for j, header in enumerate(headers):
                if j < len(row):
                    row_data[header] = row[j]
                else:
                    row_data[header] = ''
            row_data['_row_number'] = i  # Guardamos el número de fila para actualizaciones
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
        
        # Asegurar que values tenga al menos 7 columnas
        while len(values) < 7:
            values.append('')
        
        body = {
            'values': [values[:7]]  # Solo las primeras 7 columnas
        }
        
        result = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_name}'!A{row_number}:G{row_number}",
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
        
        # Crear solicitud para eliminar la fila
        requests = [{
            'deleteDimension': {
                'range': {
                    'sheetId': get_sheet_id(sheet_name),
                    'dimension': 'ROWS',
                    'startIndex': row_number - 1,  # 0-indexed
                    'endIndex': row_number
                }
            }
        }]
        
        body = {
            'requests': requests
        }
        
        result = sheet.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()
        
        return True
        
    except HttpError as err:
        print(f"Error deleting row: {err}")
        return False

def get_sheet_id(sheet_name):
    """Obtiene el ID de una hoja por su nombre"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID
        ).execute()
        
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        return None
        
    except HttpError as err:
        print(f"Error getting sheet ID: {err}")
        return None

def add_row_to_sheet(sheet_name, values):
    """Agrega una nueva fila a Google Sheets"""
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        # Asegurar que values tenga al menos 7 columnas
        while len(values) < 7:
            values.append('')
        
        body = {
            'values': [values[:7]]
        }
        
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_name}'!A:G",
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
        patente = request.form.get('patente')
        documento = request.form.get('documento')
        fecha_vencimiento = request.form.get('fecha_vencimiento')
        observaciones = request.form.get('observaciones', '')
        
        if not all([sheet_name, patente, documento, fecha_vencimiento]):
            return jsonify({
                'success': False, 
                'error': 'Todos los campos son obligatorios'
            }), 400
        
        try:
            datetime.strptime(fecha_vencimiento, '%Y-%m-%d')
        except ValueError:
            return jsonify({
                'success': False, 
                'error': 'Formato de fecha inválido'
            }), 400
        
        file_id = None
        drive_url = ''
        
        if 'foto' in request.files and request.files['foto'].filename != '':
            foto = request.files['foto']
            file_extension = os.path.splitext(foto.filename)[1]
            filename = f"{patente} - {documento}{file_extension}"
            file_content = foto.read()
            
            file_id = upload_file_to_drive(
                file_content, 
                filename, 
                DRIVE_FOLDER_ID
            )
            
            if file_id:
                drive_url = f"https://drive.google.com/file/d/{file_id}/view"
        
        row_values = [
            patente,
            documento,
            fecha_vencimiento,
            drive_url,
            observaciones,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            ''
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

@app.route('/api/update/<sheet_name>/<int:row_number>', methods=['POST'])
def update_document(sheet_name, row_number):
    """Actualiza un documento existente"""
    try:
        data = request.json
        patente = data.get('patente', '')
        documento = data.get('documento', '')
        fecha_vencimiento = data.get('fecha_vencimiento', '')
        observaciones = data.get('observaciones', '')
        drive_url = data.get('drive_url', '')
        
        # Validar fecha si está presente
        if fecha_vencimiento:
            try:
                datetime.strptime(fecha_vencimiento, '%Y-%m-%d')
            except ValueError:
                return jsonify({
                    'success': False, 
                    'error': 'Formato de fecha inválido'
                }), 400
        
        row_values = [
            patente,
            documento,
            fecha_vencimiento,
            drive_url,
            observaciones,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            ''
        ]
        
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