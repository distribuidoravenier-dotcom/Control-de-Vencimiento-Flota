# app.py
import os
import json
import io
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

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
_drive_service_cache = None

# Columnas de vencimiento por solapa (las que definen STATUS)
FECHAS_VENCIMIENTO = {
    'camiones_t2':        ['VENC VTV', 'VENC SEGURO', 'UTA', 'EXTINTOR', 'BOTIQUIN'],
    'maestro_ae':         ['VENC SEGURO', 'EXTINTOR'],
    'choferes_ayudantes': ['VENCIMIENTO REGISTRO', 'LIBRETA SANITARIA'],
}

# Columnas que se usan como DESCRIPCION en el historial por solapa
DESCRIPCION_COL = {
    'camiones_t2':        'PATENTE',
    'maestro_ae':         'PATENTE',
    'choferes_ayudantes': 'APELLIDO Y NOMBRE',
}

# Columna extra (solo en app.py) para marcar si ya se generó el PDF
HISTORIAL_PDF_COL = 'PDF GENERADO'

# ============================================================
# EVIDENCIAS DE DOCUMENTACIÓN (Google Drive)
# ============================================================
# Carpeta raíz en Drive donde se crearán todas las evidencias.
# Se puede sobrescribir con la variable de entorno DRIVE_ROOT_FOLDER_ID.
DRIVE_ROOT_FOLDER_ID = os.getenv('DRIVE_ROOT_FOLDER_ID', '').strip()


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

def _load_credentials():
    creds_env = os.getenv('GOOGLE_CREDENTIALS_JSON')
    if creds_env:
        info = json.loads(creds_env)
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return service_account.Credentials.from_service_account_file(
        Config.CREDENTIALS_FILE, scopes=SCOPES
    )


def get_sheets_service():
    global _service_cache
    if _service_cache is not None:
        return _service_cache
    try:
        creds = _load_credentials()
        _service_cache = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        return _service_cache
    except Exception as e:
        print(f"❌ Error credenciales (Sheets): {e}")
        raise


def get_drive_service():
    global _drive_service_cache
    if _drive_service_cache is not None:
        return _drive_service_cache
    try:
        creds = _load_credentials()
        _drive_service_cache = build('drive', 'v3', credentials=creds, cache_discovery=False)
        return _drive_service_cache
    except Exception as e:
        print(f"❌ Error credenciales (Drive): {e}")
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


def ensure_sheet_exists(sheet_cfg):
    """Crea la solapa si no existe en el spreadsheet."""
    try:
        service = get_sheets_service()
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=SPREADSHEET_ID,
            fields='sheets.properties(title)'
        ).execute()
        titles = [s['properties']['title'] for s in spreadsheet.get('sheets', [])]
        if sheet_cfg['name'] in titles:
            return True
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={'requests': [{'addSheet': {'properties': {'title': sheet_cfg['name']}}}]}
        ).execute()
        print(f"✅ Solapa '{sheet_cfg['name']}' creada.")
        return True
    except HttpError as err:
        print(f"❌ Error creando solapa '{sheet_cfg['name']}': {err}")
        return False


def ensure_sheet_headers(sheet_cfg):
    """
    Verifica que la fila 1 del Sheet tenga exactamente nuestros headers.
    Si faltan columnas al final, las agrega SIN borrar las existentes.
    """
    if not ensure_sheet_exists(sheet_cfg):
        return False
    try:
        service = get_sheets_service()
        current = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!1:1"
        ).execute().get('values', [[]])
        current_headers = current[0] if current else []

        expected = sheet_cfg['headers']

        if current_headers[:len(expected)] == expected and len(current_headers) >= len(expected):
            return True

        extras = []
        if len(current_headers) > len(expected):
            extras = current_headers[len(expected):]
        new_headers = expected + extras

        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_cfg['name']}'!A1",
            valueInputOption='RAW',
            body={'values': [new_headers]}
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
                continue
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
    columns = list(Config.HISTORIAL_COLUMNS) + [
        {'name': HISTORIAL_PDF_COL, 'type': 'select', 'options': ['SI', '']}
    ]
    cfg = {
        'name': Config.HISTORIAL_SHEET_NAME,
        'label': 'Historial de Vencimientos',
        'primary_key': None,
        'columns': columns,
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
    if not fecha:
        return ''
    return fecha.isocalendar()[1]


def get_historial_data():
    sheet_cfg = get_historial_config()
    ensure_sheet_headers(sheet_cfg)
    return get_all_data(sheet_cfg)


def ensure_historial_headers():
    sheet_cfg = get_historial_config()
    return ensure_sheet_headers(sheet_cfg)


def cargar_historial_automatico():
    historial_cfg = get_historial_config()
    ensure_sheet_headers(historial_cfg)

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
        descripcion_col = DESCRIPCION_COL.get(sheet_key)
        if not descripcion_col:
            continue
        data = get_all_data(sheet_cfg)
        for row in data['rows']:
            descripcion_val = (row.get(descripcion_col) or '').strip()
            if not descripcion_val:
                continue
            for col_name in FECHAS_VENCIMIENTO.get(sheet_key, []):
                fecha_str = row.get(col_name, '')
                fecha_v = parse_date_any(fecha_str)
                if not fecha_v:
                    continue
                if fecha_v > limite:
                    continue

                tipo = Config.HISTORIAL_TIPO_MAP.get(col_name, col_name)
                descripcion = descripcion_val
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
                    HISTORIAL_PDF_COL:       '',
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


def marcar_pdf_generado(row_number):
    sheet_cfg = get_historial_config()
    ensure_sheet_headers(sheet_cfg)
    headers = sheet_cfg['headers']
    if HISTORIAL_PDF_COL not in headers:
        return False
    col_idx = headers.index(HISTORIAL_PDF_COL)
    col_letra = col_letter(col_idx + 1)
    service = get_sheets_service()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_cfg['name']}'!{col_letra}{row_number}",
        valueInputOption='RAW',
        body={'values': [['SI']]}
    ).execute()
    return True


# ============================================================
# CONTADOR DE NOTIFICACIONES (N° de Orden correlativo)
# ============================================================

def get_contador_config():
    cfg = {
        'name': Config.NOTIFICACION_CONTADOR_SHEET,
        'label': 'Notificaciones Contador',
        'primary_key': 'AÑO',
        'columns': Config.NOTIFICACION_CONTADOR_COLUMNS,
    }
    headers = [c['name'] for c in cfg['columns']]
    return {
        'key': 'notificaciones_contador',
        'name': cfg['name'],
        'label': cfg['label'],
        'primary_key': 'AÑO',
        'columns': cfg['columns'],
        'headers': headers,
        'last_col': col_letter(len(headers))
    }


def siguiente_numero_orden():
    contador_cfg = get_contador_config()
    if not ensure_sheet_headers(contador_cfg):
        raise RuntimeError(f"No se pudo preparar la solapa '{contador_cfg['name']}'")

    anio_actual = str(date.today().year)
    service = get_sheets_service()

    data = get_all_data(contador_cfg)
    rows = data['rows']

    fila_existente = None
    fila_numero = None
    ultimo = 0

    for idx, r in enumerate(rows):
        if (r.get('AÑO') or '').strip() == anio_actual:
            fila_existente = r
            fila_numero = r['_row_number']
            try:
                ultimo = int(str(r.get('ULTIMO_NUMERO', '0')).strip() or '0')
            except (ValueError, TypeError):
                ultimo = 0
            break

    nuevo = ultimo + 1

    if fila_existente:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{contador_cfg['name']}'!B{fila_numero}",
            valueInputOption='RAW',
            body={'values': [[str(nuevo)]]}
        ).execute()
    else:
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{contador_cfg['name']}'!A:{contador_cfg['last_col']}",
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': [[anio_actual, str(nuevo)]]}
        ).execute()

    numero_formateado = str(nuevo).zfill(4)
    return f"{numero_formateado}-{anio_actual}"


# ============================================================
# EVIDENCIAS DE DOCUMENTACIÓN (Google Drive)
# ============================================================

def drive_resolver_root():
    """Devuelve el ID de la carpeta raíz a usar.
    Si DRIVE_ROOT_FOLDER_ID está definido, se usa; si no, se usa 'root' (Mi unidad).
    """
    return DRIVE_ROOT_FOLDER_ID or 'root'


def drive_listar(folder_id):
    """Lista subcarpetas e imágenes de una carpeta de Drive."""
    service = get_drive_service()
    q = f"'{folder_id}' in parents and trashed = false"
    # OJO: 'fields' debe ser 'files(id,name,...)', NO 'files(files(...))'
    fields = "files(id,name,mimeType,size,modifiedTime,webViewLink,iconLink,thumbnailLink)"
    resp = service.files().list(
        q=q,
        fields=fields,
        orderBy="folder,name",
        pageSize=1000,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = resp.get('files', [])

    carpetas = []
    imagenes = []
    for f in files:
        mime = f.get('mimeType', '')
        if mime == 'application/vnd.google-apps.folder':
            carpetas.append({
                'id': f['id'],
                'name': f['name'],
                'modifiedTime': f.get('modifiedTime', '')
            })
        elif mime.startswith('image/'):
            imagenes.append({
                'id': f['id'],
                'name': f['name'],
                'mimeType': mime,
                'size': f.get('size', ''),
                'modifiedTime': f.get('modifiedTime', ''),
                'webViewLink': f.get('webViewLink', ''),
                'thumbnailLink': f.get('thumbnailLink', ''),
                'iconLink': f.get('iconLink', '')
            })
        else:
            # otros archivos: los mostramos igual como "archivo"
            imagenes.append({
                'id': f['id'],
                'name': f['name'],
                'mimeType': mime,
                'size': f.get('size', ''),
                'modifiedTime': f.get('modifiedTime', ''),
                'webViewLink': f.get('webViewLink', ''),
                'thumbnailLink': f.get('thumbnailLink', ''),
                'iconLink': f.get('iconLink', '')
            })

    return carpetas, imagenes


def drive_obtener_info(folder_id):
    service = get_drive_service()
    return service.files().get(
        fileId=folder_id,
        fields="id,name,mimeType,parents",
        supportsAllDrives=True
    ).execute()


def drive_crear_carpeta(nombre, parent_id):
    service = get_drive_service()
    metadata = {
        'name': nombre,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = service.files().create(
        body=metadata,
        fields='id,name',
        supportsAllDrives=True
    ).execute()
    return folder


def drive_subir_imagen(nombre, parent_id, file_stream, mimetype):
    service = get_drive_service()
    metadata = {
        'name': nombre,
        'parents': [parent_id]
    }
    media = MediaIoBaseUpload(file_stream, mimetype=mimetype, resumable=False)
    file = service.files().create(
        body=metadata,
        media_body=media,
        fields='id,name,mimeType,webViewLink,thumbnailLink,size',
        supportsAllDrives=True
    ).execute()
    return file


def drive_eliminar(file_id):
    service = get_drive_service()
    service.files().delete(fileId=file_id, supportsAllDrives=True).execute()


def drive_renombrar(file_id, nuevo_nombre):
    service = get_drive_service()
    return service.files().update(
        fileId=file_id,
        body={'name': nuevo_nombre},
        fields='id,name',
        supportsAllDrives=True
    ).execute()


def drive_buscar_carpeta_por_nombre(nombre, parent_id):
    service = get_drive_service()
    q = (
        f"'{parent_id}' in parents and trashed = false "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and name = '{nombre.replace(chr(39), chr(92)+chr(39))}'"
    )
    resp = service.files().list(
        q=q,
        fields='files(id,name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = resp.get('files', [])
    return files[0] if files else None


def drive_obtener_o_crear_path(path_str, root_id):
    """path_str con '/' crea la jerarquía de carpetas y devuelve el ID final."""
    partes = [p.strip() for p in str(path_str).split('/') if p.strip()]
    current = root_id
    for parte in partes:
        existente = drive_buscar_carpeta_por_nombre(parte, current)
        if existente:
            current = existente['id']
        else:
            nueva = drive_crear_carpeta(parte, current)
            current = nueva['id']
    return current


# ============================================================
# RUTAS
# ============================================================

@app.route('/')
def index():
    return render_template('index.html', sheets=SHEETS, control=CONTROL_DOCUMENTARIO)


# ============================================================
# API
# ============================================================

@app.route('/api/config/<sheet_key>', methods=['GET'])
def api_config(sheet_key):
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
    if sheet_key not in CONTROL_DOCUMENTARIO:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    sheet_cfg = get_sheet_config(sheet_key)
    if not sheet_cfg:
        return jsonify({'success': False, 'error': 'Solapa no encontrada'}), 404
    try:
        body = request.get_json() or {}
        values = body.get('values', {})

        cd_cols = CONTROL_DOCUMENTARIO[sheet_key]['columns']
        # Columnas que NO se pueden editar desde Control Documentario
        non_editable = {sheet_cfg['primary_key']}
        if 'APELLIDO Y NOMBRE' in sheet_cfg['headers']:
            non_editable.add('APELLIDO Y NOMBRE')
        if 'PATENTE' in sheet_cfg['headers']:
            non_editable.add('PATENTE')
        non_editable.add('STATUS')  # STATUS siempre se recalcula

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

        # Construir la nueva fila: solo se actualizan las columnas que vienen en values
        # y que NO están en non_editable. El resto se preserva del current_row.
        new_row = {}
        for h in headers:
            if h in non_editable:
                # Preservar el valor existente
                new_row[h] = str(current_row.get(h, '') or '')
            elif h in values:
                # Actualizar con el valor enviado
                new_row[h] = str(values.get(h, '') or '')
            else:
                # Columna no enviada: preservar el valor existente
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
    try:
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
        'estados': Config.HISTORIAL_ESTADOS,
        'tipos_notificables': Config.NOTIFICACION_TIPOS
    })


@app.route('/api/historial-vencimientos/<int:row_number>', methods=['PUT'])
def api_update_historial(row_number):
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

        col_idx = headers.index('ESTADO')
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


@app.route('/api/notificacion/generar', methods=['POST'])
def api_generar_notificacion():
    try:
        body = request.get_json() or {}
        tipo = (body.get('tipo') or '').strip().upper()
        if tipo not in [t.upper() for t in Config.NOTIFICACION_TIPOS]:
            return jsonify({'success': False, 'error': 'Tipo no permitido'}), 400

        row_number = body.get('row_number')
        if not row_number or int(row_number) < 2:
            return jsonify({'success': False, 'error': 'row_number inválido'}), 400

        numero_orden = siguiente_numero_orden()

        try:
            marcar_pdf_generado(int(row_number))
        except Exception as e:
            print(f"⚠️ No se pudo marcar PDF generado en fila {row_number}: {e}")

        return jsonify({'success': True, 'numero_orden': numero_orden})
    except Exception as e:
        print(f"❌ generar_notificacion: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


# ============================================================
# API EVIDENCIAS DE DOCUMENTACIÓN (Drive)
# ============================================================

@app.route('/api/evidencias/root', methods=['GET'])
def api_evidencias_root():
    """Devuelve el ID y nombre de la carpeta raíz configurada."""
    try:
        root_id = drive_resolver_root()
        if root_id == 'root':
            return jsonify({'success': True, 'root': {'id': 'root', 'name': 'Mi unidad'}})
        info = drive_obtener_info(root_id)
        return jsonify({'success': True, 'root': {'id': info['id'], 'name': info['name']}})
    except Exception as e:
        print(f"❌ evidencias_root: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/evidencias/listar', methods=['GET'])
def api_evidencias_listar():
    """Lista carpetas e imágenes dentro de una carpeta (folder_id opcional)."""
    try:
        folder_id = (request.args.get('folder_id') or '').strip() or drive_resolver_root()
        carpetas, imagenes = drive_listar(folder_id)
        return jsonify({
            'success': True,
            'folder_id': folder_id,
            'carpetas': carpetas,
            'imagenes': imagenes
        })
    except Exception as e:
        print(f"❌ evidencias_listar: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/evidencias/carpeta', methods=['POST'])
def api_evidencias_crear_carpeta():
    """Crea una carpeta dentro de parent_id (o en la raíz)."""
    try:
        body = request.get_json() or {}
        nombre = (body.get('nombre') or '').strip()
        parent_id = (body.get('parent_id') or '').strip() or drive_resolver_root()
        if not nombre:
            return jsonify({'success': False, 'error': 'El nombre es obligatorio'}), 400
        folder = drive_crear_carpeta(nombre, parent_id)
        return jsonify({'success': True, 'carpeta': folder})
    except Exception as e:
        print(f"❌ evidencias_crear_carpeta: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/evidencias/subir', methods=['POST'])
def api_evidencias_subir():
    """Sube una o más imágenes a la carpeta indicada (form-data)."""
    try:
        parent_id = (request.form.get('parent_id') or '').strip() or drive_resolver_root()
        if 'archivos' not in request.files:
            return jsonify({'success': False, 'error': 'No se enviaron archivos'}), 400

        files = request.files.getlist('archivos')
        if not files:
            return jsonify({'success': False, 'error': 'No se enviaron archivos'}), 400

        subidos = []
        for f in files:
            if not f or not f.filename:
                continue
            mimetype = f.mimetype or 'application/octet-stream'
            stream = io.BytesIO(f.read())
            info = drive_subir_imagen(f.filename, parent_id, stream, mimetype)
            subidos.append({
                'id': info.get('id'),
                'name': info.get('name'),
                'webViewLink': info.get('webViewLink'),
                'thumbnailLink': info.get('thumbnailLink'),
                'size': info.get('size', '')
            })

        return jsonify({'success': True, 'subidos': subidos})
    except Exception as e:
        print(f"❌ evidencias_subir: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/evidencias/eliminar', methods=['POST'])
def api_evidencias_eliminar():
    """Elimina un archivo o carpeta de Drive."""
    try:
        body = request.get_json() or {}
        file_id = (body.get('file_id') or '').strip()
        if not file_id:
            return jsonify({'success': False, 'error': 'file_id es obligatorio'}), 400
        drive_eliminar(file_id)
        return jsonify({'success': True})
    except Exception as e:
        print(f"❌ evidencias_eliminar: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/evidencias/renombrar', methods=['POST'])
def api_evidencias_renombrar():
    """Renombra un archivo o carpeta de Drive."""
    try:
        body = request.get_json() or {}
        file_id = (body.get('file_id') or '').strip()
        nuevo_nombre = (body.get('nuevo_nombre') or '').strip()
        if not file_id or not nuevo_nombre:
            return jsonify({'success': False, 'error': 'file_id y nuevo_nombre son obligatorios'}), 400
        info = drive_renombrar(file_id, nuevo_nombre)
        return jsonify({'success': True, 'file': info})
    except Exception as e:
        print(f"❌ evidencias_renombrar: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/evidencias/thumbnail/<file_id>', methods=['GET'])
def api_evidencias_thumbnail(file_id):
    """Devuelve la imagen desde Drive usando la cuenta de servicio (proxy)."""
    try:
        service = get_drive_service()
        # Descargar el contenido del archivo
        request_drive = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        from googleapiclient.http import MediaIoBaseDownload
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_drive)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        from flask import Response
        # Detectar content-type
        meta = service.files().get(
            fileId=file_id, fields='mimeType', supportsAllDrives=True
        ).execute()
        mimetype = meta.get('mimeType', 'image/jpeg')
        return Response(fh.read(), mimetype=mimetype)
    except Exception as e:
        print(f"❌ evidencias_thumbnail: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
