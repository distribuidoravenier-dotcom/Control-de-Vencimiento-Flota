# app.py
import os
import json
import io
from datetime import datetime, date, timedelta
from flask import Flask, render_template, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import cloudinary
import cloudinary.uploader
import cloudinary.api

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

cloudinary.config(
    cloud_name=os.getenv('CLOUDINARY_CLOUD_NAME'),
    api_key=os.getenv('CLOUDINARY_API_KEY'),
    api_secret=os.getenv('CLOUDINARY_API_SECRET'),
    secure=True
)

# CARPETA RAÍZ FIJA EN CLOUDINARY
CARPETA_RAIZ = "evidencias_flota"

FECHAS_VENCIMIENTO = {
    'camiones_t2':        ['VENC VTV', 'VENC SEGURO', 'UTA', 'EXTINTOR', 'BOTIQUIN'],
    'maestro_ae':         ['VENC SEGURO', 'EXTINTOR'],
    'choferes_ayudantes': ['VENCIMIENTO REGISTRO', 'LIBRETA SANITARIA'],
}

DESCRIPCION_COL = {
    'camiones_t2':        'PATENTE',
    'maestro_ae':         'PATENTE',
    'choferes_ayudantes': 'APELLIDO Y NOMBRE',
}

HISTORIAL_PDF_COL = 'PDF GENERADO'

# ============================================================
# CÁLCULOS AUTOMÁTICOS
# ============================================================

def parse_date_any(str_val):
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
    try:
        ant = float(str(valor_antiguedad).replace(',', '.').strip())
    except (ValueError, TypeError):
        return ''
    return 'NO' if ant > 10 else 'SI'

def calc_status(sheet_key, row_dict):
    fechas = FECHAS_VENCIMIENTO.get(sheet_key, [])
    hoy = datetime.now().date()
    for col in fechas:
        d = parse_date_any(row_dict.get(col, ''))
        if d and d < hoy:
            return 'NO SALE'
    return 'SALE'

def recalcular_campos_auto(sheet_key, row_dict):
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

def col_letter(n):
    result = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result

def get_sheet_config(sheet_key):
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
        return True
    except HttpError as err:
        print(f"❌ Error creando solapa '{sheet_cfg['name']}': {err}")
        return False

def ensure_sheet_headers(sheet_cfg):
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
        return True
    except HttpError as err:
        print(f"❌ Error asegurando headers: {err}")
        return False

def get_all_data(sheet_cfg):
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
                key = (tipo.strip().upper(), descripcion.strip().upper(), fecha_venc_str.strip())
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
# CONTADOR DE NOTIFICACIONES
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
# EVIDENCIAS DE DOCUMENTACIÓN (Cloudinary)
#
# Estructura:  evidencias_flota (madre)
#                 └── Camiones T2 (hija)
#                        └── fotos
#
# IMPORTANTE: el "id" de cada carpeta es su RUTA COMPLETA relativa a la
# carpeta madre (ej: "Camiones T2" o "Camiones T2/AB123CD").
# Antes se usaba solo el nombre, por eso las carpetas no se encontraban.
# ============================================================

def _limpiar_ruta(ruta):
    """Quita barras sobrantes y espacios de una ruta relativa."""
    if not ruta:
        return ''
    partes = [p.strip() for p in str(ruta).replace('\\', '/').split('/') if p.strip()]
    return '/'.join(partes)

def _ruta_absoluta(ruta_relativa):
    """Convierte 'Camiones T2' -> 'evidencias_flota/Camiones T2'."""
    ruta_relativa = _limpiar_ruta(ruta_relativa)
    return f"{CARPETA_RAIZ}/{ruta_relativa}" if ruta_relativa else CARPETA_RAIZ

def _asegurar_carpeta_raiz():
    """Crea la carpeta madre si todavía no existe (no falla si ya existe)."""
    try:
        cloudinary.api.create_folder(CARPETA_RAIZ)
    except Exception as e:
        print(f"ℹ️ Carpeta raíz: {e}")

def _item_imagen(r):
    """Convierte un recurso de Cloudinary al formato que usa la pantalla."""
    url = r.get('secure_url', '')
    miniatura = url.replace('/upload/', '/upload/c_limit,w_300,h_300,q_auto/') if url else ''
    nombre = r.get('display_name') or r['public_id'].split('/')[-1]
    if r.get('format') and not nombre.lower().endswith('.' + r['format'].lower()):
        nombre_visible = f"{nombre}.{r['format']}"
    else:
        nombre_visible = nombre
    return {
        'id': r['public_id'],
        'name': nombre_visible,
        'mimeType': 'image/' + (r.get('format') or 'jpeg'),
        'size': str(r.get('bytes', '')),
        'modifiedTime': r.get('created_at', ''),
        'webViewLink': url,
        'thumbnailLink': miniatura,
        'iconLink': miniatura,
        'secure_url': url,
        'url': url,
    }

def cloudinary_listar(folder_path=None):
    """
    Devuelve (carpetas, imagenes, error) de una carpeta.
    Si algo falla, el error se devuelve para mostrarlo en pantalla
    (antes se ocultaba y la carpeta parecía vacía).
    """
    ruta_rel = _limpiar_ruta(folder_path)
    ruta_abs = _ruta_absoluta(ruta_rel)
    errores = []

    # ---- 1) SUBCARPETAS ----
    carpetas = []
    try:
        try:
            res = cloudinary.api.subfolders(ruta_abs)
        except Exception as e:
            # Si la carpeta madre todavía no existe, la creamos y reintentamos
            if not ruta_rel and ("not found" in str(e).lower() or "does not exist" in str(e).lower()):
                _asegurar_carpeta_raiz()
                res = cloudinary.api.subfolders(ruta_abs)
            else:
                raise
        for c in res.get('folders', []):
            nombre = c.get('name', '')
            ruta_cloud = c.get('path') or f"{ruta_abs}/{nombre}"
            # ruta relativa a la carpeta madre
            if ruta_cloud.startswith(CARPETA_RAIZ + '/'):
                ruta_hija = ruta_cloud[len(CARPETA_RAIZ) + 1:]
            else:
                ruta_hija = f"{ruta_rel}/{nombre}" if ruta_rel else nombre
            carpetas.append({
                'id': ruta_hija,        # RUTA COMPLETA relativa
                'name': nombre,
                'modifiedTime': ''
            })
    except Exception as e:
        print(f"❌ Error listando subcarpetas en Cloudinary ({ruta_abs}): {e}")
        errores.append(f"Subcarpetas: {e}")

    # ---- 2) IMÁGENES ----
    # Cloudinary tiene dos modos de carpetas (fijo y dinámico).
    # Probamos los dos métodos y juntamos los resultados sin repetir.
    imagenes_dict = {}

    # Método A: por carpeta de activos (modo dinámico, el más común hoy)
    try:
        res_a = cloudinary.api.resources_by_asset_folder(ruta_abs, max_results=500)
        for r in res_a.get('resources', []):
            if r.get('resource_type', 'image') == 'image':
                imagenes_dict[r['public_id']] = _item_imagen(r)
    except Exception as e:
        print(f"ℹ️ resources_by_asset_folder no disponible: {e}")

    # Método B: por prefijo (modo clásico)
    try:
        prefix = f"{ruta_abs}/"
        res_b = cloudinary.api.resources(
            type='upload',
            resource_type='image',
            prefix=prefix,
            max_results=500
        )
        for r in res_b.get('resources', []):
            public_id = r['public_id']
            if not public_id.startswith(prefix):
                continue
            if '/' in public_id[len(prefix):]:
                continue  # está en una subcarpeta más profunda
            imagenes_dict.setdefault(public_id, _item_imagen(r))
    except Exception as e:
        print(f"ℹ️ resources por prefijo: {e}")

    imagenes = list(imagenes_dict.values())

    error = ' | '.join(errores) if (errores and not carpetas and not imagenes) else None
    return carpetas, imagenes, error

def cloudinary_crear_carpeta(nombre, parent_path):
    """Crea una carpeta hija dentro de la carpeta actual."""
    nombre = (nombre or '').replace('/', '-').replace('\\', '-').strip()
    if not nombre:
        raise ValueError('El nombre es obligatorio')

    _asegurar_carpeta_raiz()

    ruta_rel = _limpiar_ruta(parent_path)
    ruta_hija_rel = f"{ruta_rel}/{nombre}" if ruta_rel else nombre
    cloudinary.api.create_folder(_ruta_absoluta(ruta_hija_rel))
    return {'id': ruta_hija_rel, 'name': nombre}

def cloudinary_subir_imagen(file_stream, nombre_original, folder_path):
    """Sube una imagen a la carpeta indicada."""
    ruta_abs = _ruta_absoluta(folder_path)
    nombre_base = os.path.splitext(nombre_original or 'foto')[0]

    resultado = cloudinary.uploader.upload(
        file_stream,
        folder=ruta_abs,
        public_id=nombre_base,
        use_filename=True,
        unique_filename=True,
        overwrite=False,
        resource_type='image'
    )
    item = _item_imagen(resultado)
    item['name'] = nombre_original or item['name']
    return item

def cloudinary_eliminar(public_id):
    """Elimina una imagen."""
    try:
        result = cloudinary.uploader.destroy(public_id, invalidate=True)
        return result.get('result') == 'ok'
    except Exception as e:
        print(f"❌ Error eliminando de Cloudinary: {e}")
        return False

def cloudinary_eliminar_carpeta(ruta_rel):
    """Elimina una carpeta (Cloudinary solo permite borrarla si está vacía)."""
    ruta_rel = _limpiar_ruta(ruta_rel)
    if not ruta_rel:
        raise ValueError('No se puede eliminar la carpeta madre')
    cloudinary.api.delete_folder(_ruta_absoluta(ruta_rel))
    return True

def cloudinary_renombrar(public_id, nuevo_nombre):
    """Renombra una imagen."""
    try:
        partes = public_id.split('/')
        if len(partes) > 1:
            carpeta = '/'.join(partes[:-1])
            nuevo_public_id = f"{carpeta}/{nuevo_nombre}"
        else:
            nuevo_public_id = nuevo_nombre
        resultado = cloudinary.uploader.rename(public_id, nuevo_public_id)
        return {'id': resultado['public_id'], 'name': nuevo_nombre}
    except Exception as e:
        print(f"❌ Error renombrando en Cloudinary: {e}")
        return None

def cloudinary_renombrar_carpeta(ruta_rel, nuevo_nombre):
    """Renombra una carpeta."""
    ruta_rel = _limpiar_ruta(ruta_rel)
    nuevo_nombre = (nuevo_nombre or '').replace('/', '-').strip()
    if not ruta_rel or not nuevo_nombre:
        raise ValueError('Datos incompletos')
    partes = ruta_rel.split('/')
    nueva_rel = '/'.join(partes[:-1] + [nuevo_nombre])
    cloudinary.api.rename_folder(_ruta_absoluta(ruta_rel), _ruta_absoluta(nueva_rel))
    return {'id': nueva_rel, 'name': nuevo_nombre}

# ============================================================
# RUTAS
# ============================================================

@app.route('/')
def index():
    return render_template('index.html', sheets=SHEETS, control=CONTROL_DOCUMENTARIO)

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
        non_editable = {sheet_cfg['primary_key']}
        if 'APELLIDO Y NOMBRE' in sheet_cfg['headers']:
            non_editable.add('APELLIDO Y NOMBRE')
        if 'PATENTE' in sheet_cfg['headers']:
            non_editable.add('PATENTE')
        non_editable.add('STATUS')
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
            if h in non_editable:
                new_row[h] = str(current_row.get(h, '') or '')
            elif h in values:
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
# API EVIDENCIAS DE DOCUMENTACIÓN
# ============================================================

def _sin_cache(resp):
    """Evita que el navegador guarde respuestas viejas de las carpetas."""
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    return resp

@app.route('/api/evidencias/root', methods=['GET'])
def api_evidencias_root():
    # La raíz tiene id vacío; el servidor la traduce a la carpeta madre "evidencias_flota"
    return _sin_cache(jsonify({'success': True, 'root': {'id': '', 'name': CARPETA_RAIZ}}))

@app.route('/api/evidencias/listar', methods=['GET'])
def api_evidencias_listar():
    try:
        folder_path = _limpiar_ruta(request.args.get('folder_id', ''))
        carpetas, imagenes, error = cloudinary_listar(folder_path)
        if error:
            return _sin_cache(jsonify({'success': False, 'error': error})), 500
        return _sin_cache(jsonify({
            'success': True,
            'folder_id': folder_path,
            'carpetas': carpetas,
            'imagenes': imagenes
        }))
    except Exception as e:
        print(f"❌ evidencias_listar: {e}")
        return _sin_cache(jsonify({'success': False, 'error': str(e)})), 500

@app.route('/api/evidencias/carpeta', methods=['POST'])
def api_evidencias_crear_carpeta():
    try:
        body = request.get_json() or {}
        nombre = (body.get('nombre') or '').strip()
        parent_id = _limpiar_ruta(body.get('parent_id') or '')

        if not nombre:
            return jsonify({'success': False, 'error': 'El nombre es obligatorio'}), 400

        folder = cloudinary_crear_carpeta(nombre, parent_id)
        return jsonify({'success': True, 'carpeta': folder})
    except Exception as e:
        print(f"❌ evidencias_crear_carpeta: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/evidencias/subir', methods=['POST'])
def api_evidencias_subir():
    try:
        parent_id = _limpiar_ruta(request.form.get('parent_id') or '')

        if 'archivos' not in request.files:
            return jsonify({'success': False, 'error': 'No se enviaron archivos'}), 400
        files = request.files.getlist('archivos')
        if not files:
            return jsonify({'success': False, 'error': 'No se enviaron archivos'}), 400

        subidos = []
        for f in files:
            if not f or not f.filename:
                continue
            file_bytes = f.read()
            info = cloudinary_subir_imagen(file_bytes, f.filename, parent_id)
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
    try:
        body = request.get_json() or {}
        file_id = (body.get('file_id') or '').strip()
        tipo = (body.get('tipo') or 'archivo').strip()
        if not file_id:
            return jsonify({'success': False, 'error': 'file_id es obligatorio'}), 400

        if tipo == 'carpeta':
            try:
                cloudinary_eliminar_carpeta(file_id)
                return jsonify({'success': True})
            except Exception as e:
                msg = str(e)
                if 'not empty' in msg.lower():
                    msg = 'La carpeta no está vacía. Eliminá primero las fotos y subcarpetas que contiene.'
                return jsonify({'success': False, 'error': msg}), 500

        exito = cloudinary_eliminar(file_id)
        if exito:
            return jsonify({'success': True})
        return jsonify({'success': False, 'error': 'No se pudo eliminar de Cloudinary'}), 500
    except Exception as e:
        print(f"❌ evidencias_eliminar: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/evidencias/renombrar', methods=['POST'])
def api_evidencias_renombrar():
    try:
        body = request.get_json() or {}
        file_id = (body.get('file_id') or '').strip()
        nuevo_nombre = (body.get('nuevo_nombre') or '').strip()
        tipo = (body.get('tipo') or 'archivo').strip()
        if not file_id or not nuevo_nombre:
            return jsonify({'success': False, 'error': 'file_id y nuevo_nombre son obligatorios'}), 400

        if tipo == 'carpeta':
            info = cloudinary_renombrar_carpeta(file_id, nuevo_nombre)
            return jsonify({'success': True, 'file': info})

        info = cloudinary_renombrar(file_id, nuevo_nombre)
        if info:
            return jsonify({'success': True, 'file': info})
        return jsonify({'success': False, 'error': 'No se pudo renombrar'}), 500
    except Exception as e:
        print(f"❌ evidencias_renombrar: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/evidencias/thumbnail/<path:file_id>', methods=['GET'])
def api_evidencias_thumbnail(file_id):
    try:
        url = cloudinary.CloudinaryImage(file_id).build_url(secure=True)
        from flask import redirect
        return redirect(url)
    except Exception as e:
        print(f"❌ evidencias_thumbnail: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
