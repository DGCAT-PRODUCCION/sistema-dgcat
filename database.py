import os
import sqlite3
import hashlib
import pandas as pd
import streamlit as st
from datetime import datetime
from sqlalchemy import create_engine, text
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_FILE = "dgcat_gestion.db"

# -----------------------------------------------------------------------------
# CONEXIÓN A BASE DE DATOS
# -----------------------------------------------------------------------------
def get_db_url():
    env_url = os.environ.get("SUPABASE_DB_URL")
    if env_url:
        if env_url.startswith("postgres://"):
            env_url = env_url.replace("postgres://", "postgresql://", 1)
        return env_url
    
    try:
        if "SUPABASE_DB_URL" in st.secrets:
            secret_url = st.secrets["SUPABASE_DB_URL"]
            if secret_url.startswith("postgres://"):
                secret_url = secret_url.replace("postgres://", "postgresql://", 1)
            return secret_url
    except Exception:
        pass

    return f"sqlite:///{DB_FILE}"

@st.cache_resource
def get_engine():
    db_url = get_db_url()
    if db_url.startswith("sqlite"):
        return create_engine(db_url, connect_args={"check_same_thread": False})
    return create_engine(db_url, pool_pre_ping=True, pool_size=10, max_overflow=20)

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def parse_date_safe(val):
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
# INICIALIZACIÓN DE ESTRUCTURAS Y LIMPIEZA DE DATOS
# -----------------------------------------------------------------------------
def init_db():
    engine = get_engine()
    is_sqlite = engine.url.drivername == 'sqlite'

    with engine.begin() as conn:
        # 1. Tabla Usuarios
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

        siscat_iniciales = ["ACUSE", "CONCLUIDO", "CORREO", "SISTEMAS", "SUBIDO"]
        for item in siscat_iniciales:
            if is_sqlite:
                conn.execute(text("INSERT OR IGNORE INTO cat_siscat (nombre) VALUES (:n)"), {"n": item})
            else:
                conn.execute(text("INSERT INTO cat_siscat (nombre) VALUES (:n) ON CONFLICT DO NOTHING"), {"n": item})

        tramites_iniciales = ["CAMBIO DE DESTINO", "CAMBIO DE SUPERFICIE", "DOMINIO PLENO", "ACT. DE MOSAICO", "SENTENCIA"]
        for item in tramites_iniciales:
            if is_sqlite:
                conn.execute(text("INSERT OR IGNORE INTO cat_tramite (nombre) VALUES (:n)"), {"n": item})
            else:
                conn.execute(text("INSERT INTO cat_tramite (nombre) VALUES (:n) ON CONFLICT DO NOTHING"), {"n": item})

        # 3. Tabla Principal de Oficios
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
                    fecha_entrega TEXT,
                    fecha_recibido TEXT,
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
                    fecha_entrega TEXT,
                    fecha_recibido TEXT,
                    scg TEXT,
                    siscat TEXT,
                    tipo_tramite TEXT,
                    observaciones TEXT,
                    archivo_escaneado TEXT
                );
            """))

        # Carga de catálogos de ubicación
        try:
            conn.execute(text("SELECT 1 FROM cat_ubicaciones LIMIT 1;"))
        except Exception:
            if os.path.exists("ESTADOS_TOTAL.xlsx"):
                df_e = pd.read_excel("ESTADOS_TOTAL.xlsx")
                df_e.columns = [str(c).strip().lower() for c in df_e.columns]
                
                col_edo = 'estado' if 'estado' in df_e.columns else df_e.columns[1]
                col_mun = 'municipio' if 'municipio' in df_e.columns else df_e.columns[3]
                col_eji = 'nucleo agrario' if 'nucleo agrario' in df_e.columns else ('ejido' if 'ejido' in df_e.columns else df_e.columns[5])
                
                df_clean = pd.DataFrame({
                    'estado': df_e[col_edo].astype(str).str.strip(),
                    'municipio': df_e[col_mun].astype(str).str.strip(),
                    'ejido': df_e[col_eji].astype(str).str.strip()
                })
                df_clean.to_sql("cat_ubicaciones", engine, if_exists="replace", index=False)

        # Carga del histórico de datos
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

        # Saneamiento de basura textual ('None', 'NAN', etc.) a NULL real en archivo_escaneado
        conn.execute(text("""
            UPDATE oficios 
            SET archivo_escaneado = NULL 
            WHERE TRIM(UPPER(COALESCE(archivo_escaneado, ''))) IN ('', 'NONE', 'NAN', 'NULL', 'UNDEFINED');
        """))

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
    
    enviar_notificacion_correo_html(
        nombre_completo=nombre_completo.strip(),
        username=usr_clean,
        rol=rol
    )
    
    return True, f"Usuario '{usr_clean}' creado con éxito."

def enviar_notificacion_correo_html(nombre_completo, username, rol):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    smtp_user = os.environ.get("SMTP_USER", "actmosaicocatastral@gmail.com")
    smtp_pass = os.environ.get("SMTP_PASS", "")

    try:
        if "SMTP_PASS" in st.secrets:
            smtp_pass = st.secrets["SMTP_PASS"]
        if "SMTP_USER" in st.secrets:
            smtp_user = st.secrets["SMTP_USER"]
    except Exception:
        pass

    if not smtp_pass:
        return False

    fecha_hora_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    perfil_desc = {
        "operador": "Capturista / Operador (Solo Carga PDF)",
        "supervisor": "Supervisor / Directivo (Acceso Total sin Usuarios)",
        "admin": "Administrador (Acceso Completo y Contraseñas)"
    }.get(rol.lower(), rol.upper())

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f8; margin: 0; padding: 20px; }}
            .card {{ max-width: 650px; margin: 0 auto; background-color: #ffffff; border: 1px solid #047857; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.08); }}
            .header {{ background-color: #047857; color: #ffffff; padding: 24px; text-align: center; }}
            .header h2 {{ margin: 0; font-size: 1.4rem; font-weight: 700; letter-spacing: 0.5px; }}
            .header p {{ margin: 6px 0 0 0; font-size: 0.88rem; opacity: 0.9; }}
            .content {{ padding: 28px; color: #1f2937; }}
            .intro {{ font-size: 0.95rem; margin-bottom: 20px; color: #374151; }}
            .table-details {{ width: 100%; border-collapse: collapse; margin-bottom: 24px; }}
            .table-details td {{ padding: 12px 16px; font-size: 0.92rem; border-bottom: 1px solid #f3f4f6; }}
            .table-details tr:nth-child(odd) {{ background-color: #f9fafb; }}
            .label {{ font-weight: 700; color: #111827; width: 35%; }}
            .value {{ color: #047857; font-weight: 600; }}
            .footer {{ padding: 16px 28px; background-color: #f9fafb; border-top: 1px solid #e5e7eb; font-size: 0.82rem; color: #6b7280; text-align: center; }}
        </style>
    </head>
    <body>
        <div class="card">
            <div class="header">
                <h2>🏛️ DIRECCIÓN GENERAL DE CATASTRO (DGCAT)</h2>
                <p>Notificación Automática de Registro de Usuario</p>
            </div>
            <div class="content">
                <p class="intro">Se ha registrado un nuevo usuario en la plataforma con el siguiente detalle:</p>
                <table class="table-details">
                    <tr><td class="label">Nombre Completo:</td><td>{nombre_completo}</td></tr>
                    <tr><td class="label">Nombre de Usuario:</td><td class="value">{username}</td></tr>
                    <tr><td class="label">Perfil / Rol:</td><td class="value">{perfil_desc}</td></tr>
                    <tr><td class="label">Fecha y Hora:</td><td>{fecha_hora_actual}</td></tr>
                </table>
            </div>
            <div class="footer">Este es un mensaje automático generado por el Sistema DGCAT.</div>
        </div>
    </body>
    </html>
    """

    msg = MIMEMultipart('alternative')
    msg['From'] = smtp_user
    msg['To'] = "actmosaicocatastral@gmail.com"
    msg['Subject'] = f"🔔 Nuevo Registro de Usuario en Sistema DGCAT - {username.upper()}"
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        server.quit()
        return True
    except Exception:
        return False

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
        
    return True, f"Contraseña actualizada para '{usr_clean}'."

def eliminar_usuario(username):
    engine = get_engine()
    usr_clean = username.strip().lower()
    if usr_clean == "admin":
        return False, "No se puede eliminar al administrador principal."
        
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM usuarios WHERE LOWER(username) = :u"), {"u": usr_clean})
        
    return True, f"Usuario '{usr_clean}' eliminado."

def eliminar_oficio(id_oficio):
    engine = get_engine()
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM oficios WHERE id = :id"), {"id": int(id_oficio)})
    return True, f"Oficio ID #{id_oficio} eliminado."

def generar_excel_ejecutivo(df, filename="Reporte_DGCAT_Ejecutivo.xlsx"):
    wb = Workbook()
    
    ws_sum = wb.active
    ws_sum.title = "Resumen Ejecutivo"
    ws_sum.views.sheetView[0].showGridLines = True
    
    ws_det = wb.create_sheet(title="Detalle General")
    ws_det.views.sheetView[0].showGridLines = True

    COLOR_VERDE = "047857"
    COLOR_GRIS = "F3F4F6"
    COLOR_TEXTO = "1F2937"

    font_title = Font(name="Segoe UI", size=16, bold=True, color=COLOR_VERDE)
    font_sub = Font(name="Segoe UI", size=10, italic=True, color="4B5563")
    font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=11, bold=True, color=COLOR_TEXTO)
    font_regular = Font(name="Segoe UI", size=10, color=COLOR_TEXTO)

    fill_header = PatternFill(start_color=COLOR_VERDE, end_color=COLOR_VERDE, fill_type="solid")
    fill_zebra = PatternFill(start_color=COLOR_GRIS, end_color=COLOR_GRIS, fill_type="solid")

    border_thin = Side(border_style="thin", color="D1D5DB")
    border_box = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    ws_sum.cell(row=1, column=1, value="DIRECCIÓN GENERAL DE CATASTRO").font = font_title
    ws_sum.cell(row=2, column=1, value=f"REPORTE DE CONTROL DE GESTIÓN | FECHA: {datetime.now().strftime('%Y-%m-%d')}").font = font_sub

    ws_sum.cell(row=4, column=1, value="Métrica Directiva").font = font_header
    ws_sum.cell(row=4, column=1).fill = fill_header
    ws_sum.cell(row=4, column=2, value="Valor").font = font_header
    ws_sum.cell(row=4, column=2).fill = fill_header

    total = len(df)
    scg_col = 'scg' if 'scg' in df.columns else ('SCG' if 'SCG' in df.columns else None)
    sis_col = 'siscat' if 'siscat' in df.columns else ('SISCAT' if 'SISCAT' in df.columns else None)

    scg_conc = len(df[df[scg_col] == 'CONCLUIDO']) if scg_col else 0
    sis_conc = len(df[df[sis_col].isin(['CONCLUIDO', 'SUBIDO'])]) if sis_col else 0

    metricas = [
        ("Total de Oficios Atendidos", total),
        ("Oficios Concluidos en SCG", scg_conc),
        ("% Eficiencia SCG", f"{round((scg_conc/total*100), 1)}%" if total > 0 else "0%"),
        ("Oficios Procesados SISCAT", sis_conc),
        ("% Avance SISCAT", f"{round((sis_conc/total*100), 1)}%" if total > 0 else "0%")
    ]

    for idx, (m, v) in enumerate(metricas, start=5):
        c1 = ws_sum.cell(row=idx, column=1, value=m)
        c2 = ws_sum.cell(row=idx, column=2, value=v)
        c1.font = font_regular
        c2.font = font_bold
        c1.border = border_box
        c2.border = border_box
        c1.alignment = Alignment(vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")

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
    
    df_lower = df.copy()
    df_lower.columns = [c.lower() for c in df_lower.columns]

    for row_idx, row in df_lower.iterrows():
        row_data = [row.get(c, '') for c in cols_df]
        ws_det.append(row_data)
        
        current_row = row_idx + 2
        is_zebra = (row_idx % 2 == 1)
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws_det.cell(row=current_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_box
            if is_zebra:
                cell.fill = fill_zebra
            
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
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 45)

    wb.save(filename)
    return filename
