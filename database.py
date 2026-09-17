import os
import smtplib
import hashlib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import streamlit as st
from sqlalchemy import create_engine, text

# CONFIGURACIÓN DE CONEXIÓN A BASE DE DATOS (SUPABASE / POSTGRESQL O SQLITE)
def get_db_url():
    if "SUPABASE_DB_URL" in st.secrets:
        return st.secrets["SUPABASE_DB_URL"]
    return "sqlite:///dgcat_gestion.db"

def get_engine():
    return create_engine(get_db_url())

# CORREO DESTINO DE NOTIFICACIONES
CORREO_DESTINO = "actmosaicocatastral@gmail.com"

# CONFIGURACIÓN SMTP DE ENVÍO
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = os.getenv("SMTP_USER", "actmosaicocatastral@gmail.com")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "wycd ksbc mjbc onyw")

def hash_password(password: str) -> str:
    return hashlib.sha256(password.strip().encode()).hexdigest()

def enviar_notificacion_correo(nombre_completo, username, rol):
    """Envía un correo electrónico de notificación a actmosaicocatastral@gmail.com al registrar un usuario."""
    if not SMTP_PASSWORD:
        print(f"[NOTIFICACIÓN LOCAL] Nuevo usuario registrado: {nombre_completo} ({username}) - Rol: {rol}")
        return

    try:
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = CORREO_DESTINO
        msg['Subject'] = f"🔔 Nuevo Registro de Usuario en Sistema DGCAT - {username.upper()}"

        rol_desc = "Administrador (Acceso Completo)" if rol == "admin" else ("Supervisor / Directivo" if rol == "supervisor" else "Capturista / Operador (PDF y Ubicación)")
        fecha_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cuerpo_html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333333;">
            <div style="background-color: #064E3B; padding: 15px; text-align: center; color: white; border-radius: 8px 8px 0 0;">
                <h2>🏛️ DIRECCIÓN GENERAL DE CATASTRO (DGCAT)</h2>
                <p style="margin: 0;">Notificación Automática de Registro de Usuario</p>
            </div>
            <div style="border: 1px solid #10B981; padding: 20px; border-radius: 0 0 8px 8px; background-color: #F8FAFC;">
                <p>Se ha registrado un nuevo usuario en la plataforma con el siguiente detalle:</p>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px; font-weight: bold; width: 35%;">Nombre Completo:</td><td style="padding: 8px;">{nombre_completo}</td></tr>
                    <tr style="background-color: #F1F5F9;"><td style="padding: 8px; font-weight: bold;">Nombre de Usuario:</td><td style="padding: 8px;"><code>{username}</code></td></tr>
                    <tr><td style="padding: 8px; font-weight: bold;">Perfil / Rol:</td><td style="padding: 8px;"><strong style="color: #10B981;">{rol_desc}</strong></td></tr>
                    <tr style="background-color: #F1F5F9;"><td style="padding: 8px; font-weight: bold;">Fecha y Hora:</td><td style="padding: 8px;">{fecha_str}</td></tr>
                </table>
                <br>
                <p style="font-size: 0.85rem; color: #64748B;">Este es un mensaje automático generado por el Sistema de Control de Entrada y Salida de Oficios DGCAT.</p>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(cuerpo_html, 'html'))

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, CORREO_DESTINO, msg.as_string())
        server.quit()
        print(f"✅ Notificación enviada exitosamente a {CORREO_DESTINO}")
    except Exception as e:
        print(f"⚠️ No se pudo enviar el correo de notificación: {e}")

def init_db():
    engine = get_engine()
    
    with engine.begin() as conn:
        # Tabla de usuarios
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id SERIAL PRIMARY KEY,
                username VARCHAR(100) UNIQUE NOT NULL,
                password_hash VARCHAR(255) NOT NULL,
                nombre_completo VARCHAR(255) NOT NULL,
                rol VARCHAR(50) DEFAULT 'operador',
                password_plain TEXT
            );
        """))

        # Inserción de usuarios iniciales si no existen
        res_admin = conn.execute(text("SELECT COUNT(*) FROM usuarios WHERE username = 'admin'")).scalar()
        if res_admin == 0:
            conn.execute(
                text("INSERT INTO usuarios (username, password_hash, nombre_completo, rol, password_plain) VALUES (:u, :p, :n, :r, :pp)"),
                {"u": "admin", "p": hash_password("admin123"), "n": "Administrador DGCAT", "r": "admin", "pp": "admin123"}
            )

        res_op = conn.execute(text("SELECT COUNT(*) FROM usuarios WHERE username = 'operador'")).scalar()
        if res_op == 0:
            conn.execute(
                text("INSERT INTO usuarios (username, password_hash, nombre_completo, rol, password_plain) VALUES (:u, :p, :n, :r, :pp)"),
                {"u": "operador", "p": hash_password("operador123"), "n": "Capturista PDF DGCAT", "r": "operador", "pp": "operador123"}
            )

        # Catálogos
        conn.execute(text("CREATE TABLE IF NOT EXISTS cat_scg (nombre VARCHAR(100) UNIQUE);"))
        conn.execute(text("CREATE TABLE IF NOT EXISTS cat_siscat (nombre VARCHAR(100) UNIQUE);"))
        conn.execute(text("CREATE TABLE IF NOT EXISTS cat_tramite (nombre VARCHAR(100) UNIQUE);"))

        scg_iniciales = ["BANDEJA DE GEOGRAFO", "CON RESPUESTA PREVIA", "CONCLUIDO", "EN ESPERA DE SISTEMAS", "EN OTRA BANDEJA", "GEOG. PATRICIA", "SISTEMAS", "SUBIDO"]
        for item in scg_iniciales:
            conn.execute(text("INSERT INTO cat_scg (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": item})

        siscat_iniciales = ["ACUSE", "CONCLUIDO", "CORREO", "SISTEMAS", "SUBIDO"]
        for item in siscat_iniciales:
            conn.execute(text("INSERT INTO cat_siscat (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": item})

        tramites_iniciales = ["CAMBIO DE DESTINO", "CAMBIO DE SUPERFICIE", "DOMINIO PLENO", "ACT. DE MOSAICO", "SENTENCIA"]
        for item in tramites_iniciales:
            conn.execute(text("INSERT INTO cat_tramite (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": item})

        # Tabla de oficios
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS oficios (
                id SERIAL PRIMARY KEY,
                id_registro VARCHAR(50),
                estado VARCHAR(100),
                municipio VARCHAR(100),
                ejido VARCHAR(200),
                no_oficio VARCHAR(100),
                dgcat VARCHAR(100),
                fecha_entrega DATE,
                fecha_recibido DATE,
                scg VARCHAR(100),
                siscat VARCHAR(100),
                tipo_tramite VARCHAR(100),
                observaciones TEXT,
                archivo_escaneado VARCHAR(255)
            );
        """))

        # Cargar catálogo de ubicaciones
        res_cat = conn.execute(text("SELECT EXISTS (SELECT FROM information_schema.tables WHERE table_name = 'cat_ubicaciones');")).scalar()
        if not res_cat and os.path.exists("ESTADOS_TOTAL.xlsx"):
            df_e = pd.read_excel("ESTADOS_TOTAL.xlsx", engine="openpyxl")
            df_e['ESTADO'] = df_e['ESTADO'].astype(str).str.strip()
            df_e['MUNICIPIO'] = df_e['MUNICIPIO'].astype(str).str.strip()
            df_e['NUCLEO AGRARIO'] = df_e['NUCLEO AGRARIO'].astype(str).str.strip()
            df_e[['ESTADO', 'MUNICIPIO', 'NUCLEO AGRARIO']].to_sql("cat_ubicaciones", engine, if_exists="replace", index=False)

        # Cargar archivo histórico si está vacía la tabla
        res_of = conn.execute(text("SELECT COUNT(*) FROM oficios")).scalar()
        if res_of == 0 and os.path.exists("CONTROL DE ENTRADA Y SALIDA DE LOS OFICIOS DE RESPUESTA.xlsx"):
            df_c = pd.read_excel("CONTROL DE ENTRADA Y SALIDA DE LOS OFICIOS DE RESPUESTA.xlsx", sheet_name="CONTROL_DE_ENTRADA", engine="openpyxl")
            df_c.columns = [c.strip() for c in df_c.columns]
            for _, row in df_c.iterrows():
                conn.execute(text("""
                    INSERT INTO oficios (id_registro, estado, municipio, ejido, no_oficio, dgcat, fecha_entrega, fecha_recibido, scg, siscat, observaciones)
                    VALUES (:id_reg, :edo, :mun, :eji, :no_of, :dg, :f_ent, :f_rec, :scg_val, :sis_val, :obs)
                """), {
                    "id_reg": str(row.get('ID', '')),
                    "edo": str(row.get('ESTADO', '')).strip(),
                    "mun": str(row.get('MUNICIPIO', '')).strip(),
                    "eji": str(row.get('EJIDO', '')).strip(),
                    "no_of": str(row.get('NO. OFICIO', '')).strip(),
                    "dg": str(row.get('DGCAT', '')).strip(),
                    "f_ent": str(row.get('FECHA DE ENTREGA', ''))[:10] if pd.notna(row.get('FECHA DE ENTREGA')) else None,
                    "f_rec": str(row.get('FECHA DE RECIBIDO', ''))[:10] if pd.notna(row.get('FECHA DE RECIBIDO')) else None,
                    "scg_val": str(row.get('SCG', '')).strip(),
                    "sis_val": str(row.get('SISCAT', '')).strip(),
                    "obs": str(row.get('OBSERVACIONES', '')).strip()
                })

def registrar_nuevo_usuario(username, password, nombre_completo, rol="operador"):
    engine = get_engine()
    pwd_hash = hash_password(password)
    try:
        with engine.begin() as conn:
            conn.execute(
                text("INSERT INTO usuarios (username, password_hash, nombre_completo, rol, password_plain) VALUES (:u, :p, :n, :r, :pp)"),
                {"u": username.strip().lower(), "p": pwd_hash, "n": nombre_completo.strip(), "r": rol, "pp": password.strip()}
            )
        enviar_notificacion_correo(nombre_completo.strip(), username.strip().lower(), rol)
        return True, "Cuenta registrada exitosamente en la base de datos."
    except Exception as e:
        return False, f"Error al registrar usuario (Posible duplicado): {e}"

def cambiar_password_usuario(username, nueva_password):
    engine = get_engine()
    pwd_hash = hash_password(nueva_password)
    try:
        with engine.begin() as conn:
            conn.execute(
                text("UPDATE usuarios SET password_hash = :p, password_plain = :pp WHERE LOWER(username) = LOWER(:u)"),
                {"p": pwd_hash, "pp": nueva_password.strip(), "u": username.strip()}
            )
        return True, f"La contraseña de '{username}' fue actualizada correctamente a: {nueva_password.strip()}"
    except Exception as e:
        return False, f"Error al cambiar contraseña: {e}"

def eliminar_usuario(username):
    engine = get_engine()
    try:
        with engine.begin() as conn:
            conn.execute(text("DELETE FROM usuarios WHERE LOWER(username) = LOWER(:u)"), {"u": username.strip()})
        return True, f"El usuario '{username}' ha sido eliminado correctamente."
    except Exception as e:
        return False, f"Error al eliminar usuario: {e}"

def eliminar_oficio(oficio_id):
    """Elimina permanentemente un oficio de la base de datos por su ID."""
    engine = get_engine()
    try:
        with engine.begin() as conn:
            conn.execute(text("DELETE FROM oficios WHERE id = :id"), {"id": oficio_id})
        return True, f"El oficio ID #{oficio_id} ha sido eliminado correctamente de la base de datos."
    except Exception as e:
        return False, f"Error al eliminar el oficio: {e}"

def verificar_login(username, password):
    engine = get_engine()
    pwd_hash = hash_password(password)
    with engine.connect() as conn:
        result = conn.execute(
            text("SELECT username, nombre_completo, rol FROM usuarios WHERE LOWER(username) = LOWER(:u) AND password_hash = :p"),
            {"u": username.strip(), "p": pwd_hash}
        ).fetchone()
        return result

def generar_excel_ejecutivo(df, filename="Reporte_DGCAT_Ejecutivo.xlsx"):
    wb = openpyxl.Workbook()
    
    ws_sum = wb.active
    ws_sum.title = "Resumen Ejecutivo"
    ws_sum.views.sheetView[0].showGridLines = True
    
    DARK_GREEN = "064E3B"
    HEADER_FILL = "0F172A"
    BORDER_COLOR = "CBD5E1"
    
    font_title = Font(name="Calibri", size=15, bold=True, color=DARK_GREEN)
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="475569")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_kpi_num = Font(name="Calibri", size=20, bold=True, color=DARK_GREEN)
    font_kpi_lbl = Font(name="Calibri", size=9, bold=True, color="475569")
    
    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )
    
    ws_sum["A1"] = "DIRECCIÓN GENERAL DE CATASTRO - REGISTRO AGRARIO NACIONAL"
    ws_sum["A1"].font = font_title
    ws_sum["A2"] = "Reporte Ejecutivo de Control de Entrada y Salida de Oficios"
    ws_sum["A2"].font = font_subtitle
    
    total = len(df)
    scg_concluidos = len(df[df['scg'] == 'CONCLUIDO']) if 'scg' in df.columns else 0
    siscat_subidos = len(df[df['siscat'].isin(['CONCLUIDO', 'SUBIDO'])]) if 'siscat' in df.columns else 0
    
    kpis = [("TOTAL OFICIOS", total), ("CONCLUIDOS SCG", scg_concluidos), ("PROCESADOS SISCAT", siscat_subidos)]
    
    col_idx = 1
    for label, val in kpis:
        cell_lbl = ws_sum.cell(row=4, column=col_idx, value=label)
        cell_lbl.font = font_kpi_lbl
        cell_lbl.alignment = Alignment(horizontal="center", vertical="center")
        cell_lbl.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        
        cell_val = ws_sum.cell(row=5, column=col_idx, value=val)
        cell_val.font = font_kpi_num
        cell_val.alignment = Alignment(horizontal="center", vertical="center")
        cell_val.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
        
        ws_sum.merge_cells(start_row=4, start_column=col_idx, end_row=4, end_column=col_idx+1)
        ws_sum.merge_cells(start_row=5, start_column=col_idx, end_row=5, end_column=col_idx+1)
        
        for r in range(4, 6):
            for c in range(col_idx, col_idx+2):
                ws_sum.cell(row=r, column=c).border = thin_border
        col_idx += 3
        
    ws_sum.cell(row=8, column=1, value="Estatus por Bandeja SCG").font = Font(name="Calibri", size=12, bold=True, color=DARK_GREEN)
    ws_sum.cell(row=9, column=1, value="Bandeja").font = font_header
    ws_sum.cell(row=9, column=1).fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")
    ws_sum.cell(row=9, column=2, value="Cantidad").font = font_header
    ws_sum.cell(row=9, column=2).fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")
    
    r_idx = 10
    if 'scg' in df.columns and not df.empty:
        scg_counts = df['scg'].value_counts()
        for b_name, b_cnt in scg_counts.items():
            c1 = ws_sum.cell(row=r_idx, column=1, value=str(b_name))
            c2 = ws_sum.cell(row=r_idx, column=2, value=int(b_cnt))
            c1.border = thin_border
            c2.border = thin_border
            if r_idx % 2 == 0:
                c1.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
                c2.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            r_idx += 1

    ws_det = wb.create_sheet(title="Detalle de Oficios")
    ws_det.views.sheetView[0].showGridLines = True
    
    header_map = {
        'id': 'ID Reg.',
        'id_registro': 'ID Numérico',
        'estado': 'Estado / Entidad',
        'municipio': 'Municipio',
        'ejido': 'Ejido / Núcleo Agrario',
        'no_oficio': 'No. de Oficio',
        'dgcat': 'Folio DGCAT',
        'fecha_entrega': 'Fecha Entrega',
        'fecha_recibido': 'Fecha Recibido',
        'scg': 'Bandeja SCG',
        'siscat': 'Estatus SISCAT',
        'tipo_tramite': 'Tipo de Trámite',
        'observaciones': 'Observaciones',
        'archivo_escaneado': 'Expediente PDF'
    }
    cols = [c for c in df.columns if c in header_map]
    
    ws_det.cell(row=1, column=1, value="REGISTRO DETALLADO DE OFICIOS DE RESPUESTA - DGCAT").font = font_title
    
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws_det.cell(row=3, column=c_idx, value=header_map[col_name])
        cell.font = font_header
        cell.fill = PatternFill(start_color=DARK_GREEN, end_color=DARK_GREEN, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    ws_det.row_dimensions[3].height = 26
    
    for r_offset, row in df.iterrows():
        row_num = r_offset + 4
        is_zebra = (row_num % 2 == 0)
        zebra_fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid") if is_zebra else PatternFill(fill_type=None)
        
        for c_idx, col_name in enumerate(cols, start=1):
            val = row[col_name]
            val_str = "" if pd.isna(val) else str(val).strip()
            cell = ws_det.cell(row=row_num, column=c_idx, value=val_str)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            if is_zebra:
                cell.fill = zebra_fill
            
            if col_name in ['id', 'id_registro', 'fecha_entrega', 'fecha_recibido', 'scg', 'siscat']:
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
