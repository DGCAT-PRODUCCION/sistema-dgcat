import os
import sqlite3
import hashlib
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine, text
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_FILE = "dgcat_gestion.db"

# -----------------------------------------------------------------------------
# CONEXIÓN A BASE DE DATOS (Soporte Híbrido: Supabase PostgreSQL / SQLite Local)
# -----------------------------------------------------------------------------
def get_db_url():
    """
    Detecta si existe la variable SUPABASE_DB_URL en las variables de entorno / secrets.
    Si existe, usa PostgreSQL en Supabase; de lo contrario, usa SQLite local.
    """
    env_url = os.environ.get("SUPABASE_DB_URL")
    if env_url:
        if env_url.startswith("postgres://"):
            env_url = env_url.replace("postgres://", "postgresql://", 1)
        return env_url
    
    try:
        import streamlit as st
        if "SUPABASE_DB_URL" in st.secrets:
            secret_url = st.secrets["SUPABASE_DB_URL"]
            if secret_url.startswith("postgres://"):
                secret_url = secret_url.replace("postgres://", "postgresql://", 1)
            return secret_url
    except Exception:
        pass

    return f"sqlite:///{DB_FILE}"

def get_engine():
    """Retorna el motor SQLAlchemy configurado para la base de datos."""
    db_url = get_db_url()
    if db_url.startswith("sqlite"):
        return create_engine(db_url, connect_args={"check_same_thread": False})
    return create_engine(db_url, pool_pre_ping=True)

def get_connection():
    """Retorna una conexión clásica de SQLite para entornos locales de respaldo."""
    return sqlite3.connect(DB_FILE)

def hash_password(password: str) -> str:
    """Encripta contraseñas usando SHA-256."""
    return hashlib.sha256(password.encode()).hexdigest()

# -----------------------------------------------------------------------------
# LIMPIEZA DE FECHAS SEGURA PARA POSTGRESQL
# -----------------------------------------------------------------------------
def parse_date_safe(val):
    """
    Limpia y convierte fechas con formatos irregulares como '20/032025' o nulos 
    a formato estándar ISO YYYY-MM-DD o None para evitar el error InvalidDatetimeFormat.
    """
    if pd.isna(val) or val is None:
        return None
    val_str = str(val).strip()
    if not val_str or val_str in ['nan', 'None', 'NaT', '']:
        return None
    
    try:
        dt = pd.to_datetime(val_str, errors='coerce', dayfirst=True)
        if pd.notna(dt):
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    
    val_clean = val_str.replace('/', '').replace('-', '').replace(' ', '')
    if len(val_clean) == 8 and val_clean.isdigit():
        try:
            day = int(val_clean[:2])
            month = int(val_clean[2:4])
            year = int(val_clean[4:])
            return f"{year:04d}-{month:02d}-{day:02d}"
        except Exception:
            return None
            
    return None

# -----------------------------------------------------------------------------
# INICIALIZACIÓN Y MIGRACIÓN DE ESTRUCTURAS DE BASE DE DATOS
# -----------------------------------------------------------------------------
def init_db():
    engine = get_engine()
    is_sqlite = engine.url.drivername == 'sqlite'

    with engine.begin() as conn:
        # 1. Tabla de Usuarios
        if is_sqlite:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS usuarios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    nombre_completo TEXT NOT NULL,
                    rol TEXT DEFAULT 'operador',
                    password_plain TEXT
                );
            """))
        else:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    nombre_completo TEXT NOT NULL,
                    rol TEXT DEFAULT 'operador',
                    password_plain TEXT
                );
            """))

        # Crear usuario Administrador inicial (admin / admin123)
        res_usr = conn.execute(text("SELECT COUNT(*) FROM usuarios")).fetchone()
        if res_usr and res_usr[0] == 0:
            conn.execute(text("""
                INSERT INTO usuarios (username, password_hash, nombre_completo, rol, password_plain)
                VALUES (:u, :h, :n, :r, :p)
            """), {
                "u": "admin",
                "h": hash_password("admin123"),
                "n": "Administrador DGCAT",
                "r": "admin",
                "p": "admin123"
            })

        # 2. Catálogos Dinámicos
        conn.execute(text("CREATE TABLE IF NOT EXISTS cat_scg (nombre TEXT UNIQUE);"))
        conn.execute(text("CREATE TABLE IF NOT EXISTS cat_siscat (nombre TEXT UNIQUE);"))
        conn.execute(text("CREATE TABLE IF NOT EXISTS cat_tramite (nombre TEXT UNIQUE);"))

        # Opciones iniciales para SCG
        scg_iniciales = [
            "BANDEJA DE GEOGRAFO", "CON RESPUESTA PREVIA", "CONCLUIDO", 
            "EN ESPERA DE SISTEMAS", "EN OTRA BANDEJA", "GEOG. PATRICIA", 
            "SISTEMAS", "SUBIDO"
        ]
        for item in scg_iniciales:
            if is_sqlite:
                conn.execute(text("INSERT OR IGNORE INTO cat_scg (nombre) VALUES (:n)"), {"n": item})
            else:
                conn.execute(text("INSERT INTO cat_scg (nombre) VALUES (:n) ON CONFLICT DO NOTHING"), {"n": item})

        # Opciones iniciales para SISCAT
        siscat_iniciales = ["ACUSE", "CONCLUIDO", "CORREO", "SISTEMAS", "SUBIDO"]
        for item in siscat_iniciales:
            if is_sqlite:
                conn.execute(text("INSERT OR IGNORE INTO cat_siscat (nombre) VALUES (:n)"), {"n": item})
            else:
                conn.execute(text("INSERT INTO cat_siscat (nombre) VALUES (:n) ON CONFLICT DO NOTHING"), {"n": item})

        # Opciones iniciales para Trámites
        tramites_iniciales = ["CAMBIO DE DESTINO", "CAMBIO DE SUPERFICIE", "DOMINIO PLENO", "ACT. DE MOSAICO", "SENTENCIA"]
        for item in tramites_iniciales:
            if is_sqlite:
                conn.execute(text("INSERT OR IGNORE INTO cat_tramite (nombre) VALUES (:n)"), {"n": item})
            else:
                conn.execute(text("INSERT INTO cat_tramite (nombre) VALUES (:n) ON CONFLICT DO NOTHING"), {"n": item})

        # 3. Tabla Principal de Oficios (Usando TEXT para evitar StringDataRightTruncation)
        if is_sqlite:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS oficios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    id_registro TEXT,
                    estado TEXT,
                    municipio TEXT,
                    ejido TEXT,
                    no_oficio TEXT,
                    dgcat TEXT,
                    fecha_entrega DATE,
                    fecha_recibido DATE,
                    scg TEXT,
                    siscat TEXT,
                    tipo_tramite TEXT,
                    observaciones TEXT,
                    archivo_escaneado TEXT
                );
            """))
        else:
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS oficios (
                    id SERIAL PRIMARY KEY,
                    id_registro TEXT,
                    estado TEXT,
                    municipio TEXT,
                    ejido TEXT,
                    no_oficio TEXT,
                    dgcat TEXT,
                    fecha_entrega DATE,
                    fecha_recibido DATE,
                    scg TEXT,
                    siscat TEXT,
                    tipo_tramite TEXT,
                    observaciones TEXT,
                    archivo_escaneado TEXT
                );
            """))
            # Asegurar la conversión de columnas existentes
            for col in ['id_registro', 'estado', 'municipio', 'ejido', 'no_oficio', 'dgcat', 'scg', 'siscat', 'tipo_tramite', 'observaciones', 'archivo_escaneado']:
                try:
                    conn.execute(text(f"ALTER TABLE oficios ALTER COLUMN {col} TYPE TEXT;"))
                except Exception:
                    pass

        # Carga masiva de ubicaciones desde ESTADOS_TOTAL.xlsx
        try:
            conn.execute(text("SELECT 1 FROM cat_ubicaciones LIMIT 1;"))
        except Exception:
            if os.path.exists("ESTADOS_TOTAL.xlsx"):
                df_e = pd.read_excel("ESTADOS_TOTAL.xlsx")
                df_e['ESTADO'] = df_e['ESTADO'].astype(str).str.strip()
                df_e['MUNICIPIO'] = df_e['MUNICIPIO'].astype(str).str.strip()
                df_e['NUCLEO AGRARIO'] = df_e['NUCLEO AGRARIO'].astype(str).str.strip()
                df_e[['ESTADO', 'MUNICIPIO', 'NUCLEO AGRARIO']].to_sql("cat_ubicaciones", engine, if_exists="replace", index=False)

        # Carga del histórico de oficios
        cursor_check = conn.execute(text("SELECT COUNT(*) FROM oficios")).fetchone()
        if cursor_check and cursor_check[0] == 0:
            ctrl_file = "CONTROL DE ENTRADA Y SALIDA DE LOS OFICIOS DE RESPUESTA.xlsx"
            if os.path.exists(ctrl_file):
                df_c = pd.read_excel(ctrl_file, sheet_name="CONTROL_DE_ENTRADA")
                df_c.columns = [c.strip() for c in df_c.columns]
                
                for _, row in df_c.iterrows():
                    f_ent = parse_date_safe(row.get('FECHA DE ENTREGA', ''))
                    f_rec = parse_date_safe(row.get('FECHA DE RECIBIDO', ''))
                    
                    conn.execute(text("""
                        INSERT INTO oficios (id_registro, estado, municipio, ejido, no_oficio, dgcat, fecha_entrega, fecha_recibido, scg, siscat, observaciones)
                        VALUES (:id_reg, :edo, :mun, :eji, :no_of, :dg, :f_ent, :f_rec, :scg_val, :sis_val, :obs)
                    """), {
                        "id_reg": str(row.get('ID', '')).strip(),
                        "edo": str(row.get('ESTADO', '')).strip(),
                        "mun": str(row.get('MUNICIPIO', '')).strip(),
                        "eji": str(row.get('EJIDO', '')).strip(),
                        "no_of": str(row.get('NO. OFICIO', '')).strip(),
                        "dg": str(row.get('DGCAT', '')).strip(),
                        "f_ent": f_ent,
                        "f_rec": f_rec,
                        "scg_val": str(row.get('SCG', '')).strip(),
                        "sis_val": str(row.get('SISCAT', '')).strip(),
                        "obs": str(row.get('OBSERVACIONES', '')).strip()
                    })

# -----------------------------------------------------------------------------
# FUNCIONES DE AUTENTICACIÓN Y GESTIÓN DE USUARIOS
# -----------------------------------------------------------------------------
def verificar_login(username, password):
    engine = get_engine()
    pwd_hash = hash_password(password)
    with engine.connect() as conn:
        res = conn.execute(
            text("SELECT username, nombre_completo, rol FROM usuarios WHERE LOWER(username) = :u AND password_hash = :p"),
            {"u": username.strip().lower(), "p": pwd_hash}
        ).fetchone()
        return res

def registrar_nuevo_usuario(username, password, nombre_completo, rol="operador"):
    engine = get_engine()
    usr_clean = username.strip().lower()
    pwd_hash = hash_password(password)
    
    with engine.begin() as conn:
        ex = conn.execute(text("SELECT id FROM usuarios WHERE LOWER(username) = :u"), {"u": usr_clean}).fetchone()
        if ex:
            return False, f"El usuario '{usr_clean}' ya está registrado."
        
        conn.execute(text("""
            INSERT INTO usuarios (username, password_hash, nombre_completo, rol, password_plain)
            VALUES (:u, :h, :n, :r, :p)
        """), {
            "u": usr_clean,
            "h": pwd_hash,
            "n": nombre_completo.strip(),
            "r": rol,
            "p": password
        })
    
    enviar_notificacion_correo("NUEVO USUARIO CREADO", f"Se ha creado la cuenta '{usr_clean}' para {nombre_completo} con el rol {rol.upper()}.")
    return True, f"Usuario '{usr_clean}' creado correctamente."

def cambiar_password_usuario(username, nueva_password):
    engine = get_engine()
    usr_clean = username.strip().lower()
    pwd_hash = hash_password(nueva_password)
    
    with engine.begin() as conn:
        conn.execute(text("""
            UPDATE usuarios 
            SET password_hash = :h, password_plain = :p 
            WHERE LOWER(username) = :u
        """), {"h": pwd_hash, "p": nueva_password, "u": usr_clean})
        
    return True, f"Contraseña actualizada correctamente para el usuario '{usr_clean}'."

def eliminar_usuario(username):
    engine = get_engine()
    usr_clean = username.strip().lower()
    
    if usr_clean == "admin":
        return False, "No se puede eliminar al usuario administrador principal."
        
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM usuarios WHERE LOWER(username) = :u"), {"u": usr_clean})
        
    return True, f"Usuario '{usr_clean}' eliminado del sistema."

def eliminar_oficio(id_oficio):
    engine = get_engine()
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM oficios WHERE id = :id"), {"id": id_oficio})
    return True, f"Oficio ID #{id_oficio} eliminado correctamente."

# -----------------------------------------------------------------------------
# ENVÍO DE NOTIFICACIONES POR CORREO SMTP
# -----------------------------------------------------------------------------
def enviar_notificacion_correo(asunto, mensaje_texto):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    smtp_user = os.environ.get("SMTP_USER", "actmosaicocatastral@gmail.com")
    smtp_pass = os.environ.get("SMTP_PASS", "")

    try:
        import streamlit as st
        if "SMTP_PASS" in st.secrets:
            smtp_pass = st.secrets["SMTP_PASS"]
        if "SMTP_USER" in st.secrets:
            smtp_user = st.secrets["SMTP_USER"]
    except Exception:
        pass

    if not smtp_pass:
        return False

    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = "actmosaicocatastral@gmail.com"
    msg['Subject'] = f"🏛️ DGCAT NOTIFICACIÓN: {asunto}"
    msg.attach(MIMEText(mensaje_texto, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        server.quit()
        return True
    except Exception:
        return False

# -----------------------------------------------------------------------------
# GENERACIÓN DE REPORTE EJECUTIVO EN EXCEL CON OPENPYXL
# -----------------------------------------------------------------------------
def generar_excel_ejecutivo(df, filename="Reporte_DGCAT_Ejecutivo.xlsx"):
    wb = Workbook()
    
    ws_sum = wb.active
    ws_sum.title = "Resumen Ejecutivo"
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_det = wb.create_sheet(title="Detalle General")
    ws_det.views.sheetView[0].showGridLines = True

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="10B981", end_color="10B981", fill_type="solid")
    
    font_title = Font(name="Calibri", size=16, bold=True, color="10B981")
    font_regular = Font(name="Calibri", size=10)
    
    border_thin = Side(border_style="thin", color="CCCCCC")
    border_box = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    ws_sum.cell(row=1, column=1, value="DIRECCIÓN GENERAL DE CATASTRO").font = font_title
    ws_sum.cell(row=2, column=1, value=f"REPORTE DE CONTROL DE GESTIÓN | GENERADO: {datetime.now().strftime('%Y-%m-%d %H:%M')}").font = Font(size=10, italic=True, color="666666")

    ws_sum.cell(row=4, column=1, value="Métrica Directiva").font = font_header
    ws_sum.cell(row=4, column=1).fill = fill_header
    ws_sum.cell(row=4, column=2, value="Valor").font = font_header
    ws_sum.cell(row=4, column=2).fill = fill_header

    total = len(df)
    scg_conc = len(df[df['scg'] == 'CONCLUIDO']) if 'scg' in df.columns else 0
    sis_conc = len(df[df['siscat'].isin(['CONCLUIDO', 'SUBIDO'])]) if 'siscat' in df.columns else 0

    metricas = [
        ("Total de Oficios Atendidos", total),
        ("Oficios Concluidos en SCG", scg_conc),
        ("% Eficiencia SCG", f"{round((scg_conc/total*100), 1)}%" if total > 0 else "0%"),
        ("Oficios Procesados SISCAT", sis_conc),
        ("% Avance SISCAT", f"{round((sis_conc/total*100), 1)}%" if total > 0 else "0%")
    ]

    for idx, (m, v) in enumerate(metricas, start=5):
        ws_sum.cell(row=idx, column=1, value=m).font = font_regular
        ws_sum.cell(row=idx, column=2, value=v).font = Font(bold=True)
        ws_sum.cell(row=idx, column=1).border = border_box
        ws_sum.cell(row=idx, column=2).border = border_box

    headers = [
        "ID", "ID REGISTRO", "ESTADO", "MUNICIPIO", "EJIDO", 
        "NO. OFICIO", "DGCAT", "FECHA ENTREGA", "FECHA RECIBIDO", 
        "SCG", "SISCAT", "TIPO TRÁMITE", "OBSERVACIONES", "ARCHIVO ESCANEADO"
    ]
    
    ws_det.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws_det.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cols_df = ['id', 'id_registro', 'estado', 'municipio', 'ejido', 'no_oficio', 'dgcat', 'fecha_entrega', 'fecha_recibido', 'scg', 'siscat', 'tipo_tramite', 'observaciones', 'archivo_escaneado']
    
    for row_idx, row in df.iterrows():
        row_data = [row.get(c, '') for c in cols_df]
        ws_det.append(row_data)
        
        current_row = row_idx + 2
        is_zebra = (row_idx % 2 == 1)
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws_det.cell(row=current_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_box
            if is_zebra:
                cell.fill = zebra_fill
            
            if col_idx in [1, 2, 8, 9, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for ws in [ws_sum, ws_det]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    for l in str(cell.value).split('\n'):
                        if len(l) > max_len:
                            max_len = len(l)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    wb.save(filename)
    return filename
