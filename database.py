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
# CONEXIÓN BASE DE DATOS
# -----------------------------------------------------------------------------
def get_db_url():
    env_url = os.environ.get("SUPABASE_DB_URL")
    if not env_url:
        try:
            if "SUPABASE_DB_URL" in st.secrets:
                env_url = st.secrets["SUPABASE_DB_URL"]
        except Exception:
            pass

    if env_url:
        if env_url.startswith("postgres://"):
            env_url = env_url.replace("postgres://", "postgresql://", 1)
        if "db.tcsihzqawhsewjmwbwlc.supabase.co" in env_url:
            env_url = env_url.replace(
                "db.tcsihzqawhsewjmwbwlc.supabase.co:5432", 
                "aws-0-us-east-1.pooler.supabase.com:6543"
            )
        return env_url

    return f"sqlite:///{DB_FILE}"

@st.cache_resource
def get_engine():
    db_url = get_db_url()
    if db_url.startswith("sqlite"):
        return create_engine(db_url, connect_args={"check_same_thread": False})
    
    return create_engine(
        db_url, 
        pool_pre_ping=True, 
        pool_size=5, 
        max_overflow=10,
        connect_args={"connect_timeout": 5}
    )

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def parse_date_safe(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if not val_str or val_str in ['nan', 'None', 'NaT', '']:
        return ""
    
    try:
        dt = pd.to_datetime(val_str, errors='coerce', dayfirst=True)
        if pd.notna(dt):
            return dt.strftime('%d/%m/%Y')
    except Exception:
        pass
            
    return val_str

# -----------------------------------------------------------------------------
# INICIALIZACIÓN ESTRUCTURA DB
# -----------------------------------------------------------------------------
def init_db():
    engine = get_engine()
    is_sqlite = engine.url.drivername == 'sqlite'

    try:
        with engine.begin() as conn:
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
                    "n": "ADMINISTRADOR DGCAT",
                    "r": "admin",
                    "p": "admin123"
                })

            conn.execute(text("CREATE TABLE IF NOT EXISTS cat_scg (nombre TEXT UNIQUE);"))
            conn.execute(text("CREATE TABLE IF NOT EXISTS cat_siscat (nombre TEXT UNIQUE);"))
            conn.execute(text("CREATE TABLE IF NOT EXISTS cat_sistemas_or (nombre TEXT UNIQUE);"))
            conn.execute(text("CREATE TABLE IF NOT EXISTS cat_tramite (nombre TEXT UNIQUE);"))

            for item in ["BANDEJA DE GEOGRAFO", "CON RESPUESTA PREVIA", "CONCLUIDO", "EN ESPERA DE SISTEMAS", "EN OTRA BANDEJA", "GEOG. PATRICIA", "SISTEMAS", "SUBIDO"]:
                conn.execute(text("INSERT OR IGNORE INTO cat_scg (nombre) VALUES (:n)"), {"n": item.upper()})

            for item in ["ACUSE", "CONCLUIDO", "CORREO", "SISTEMAS", "SUBIDO"]:
                conn.execute(text("INSERT OR IGNORE INTO cat_siscat (nombre) VALUES (:n)"), {"n": item.upper()})

            for item in ["SISTEMAS", "OR"]:
                conn.execute(text("INSERT OR IGNORE INTO cat_sistemas_or (nombre) VALUES (:n)"), {"n": item.upper()})

            for item in ["CAMBIO DE DESTINO", "CAMBIO DE SUPERFICIE", "DOMINIO PLENO", "ACT. DE MOSAICO", "SENTENCIA"]:
                conn.execute(text("INSERT OR IGNORE INTO cat_tramite (nombre) VALUES (:n)"), {"n": item.upper()})

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
                        sistemas_or TEXT,
                        tipo_tramite TEXT,
                        observaciones TEXT,
                        archivo_escaneado TEXT
                    );
                """))
                try:
                    conn.execute(text("ALTER TABLE oficios ADD COLUMN sistemas_or TEXT;"))
                except Exception:
                    pass
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
                        sistemas_or TEXT,
                        tipo_tramite TEXT,
                        observaciones TEXT,
                        archivo_escaneado TEXT
                    );
                """))
                try:
                    conn.execute(text("ALTER TABLE oficios ADD COLUMN sistemas_or TEXT;"))
                except Exception:
                    pass

            # -----------------------------------------------------------------
            # NUEVA BASE: SEGUIMIENTO DE UBICACIÓN DE PREDIO (independiente de oficios)
            # -----------------------------------------------------------------
            if is_sqlite:
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS seguimiento_predio (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        dgcat TEXT,
                        estado TEXT,
                        municipio TEXT,
                        ejido TEXT,
                        fecha_registro TEXT,
                        fecha_actualizacion TEXT,
                        observaciones TEXT,
                        archivo_escaneado TEXT,
                        registrado_por TEXT
                    );
                """))
            else:
                conn.execute(text("""
                    CREATE TABLE IF NOT EXISTS seguimiento_predio (
                        id SERIAL PRIMARY KEY,
                        dgcat TEXT,
                        estado TEXT,
                        municipio TEXT,
                        ejido TEXT,
                        fecha_registro TEXT,
                        fecha_actualizacion TEXT,
                        observaciones TEXT,
                        archivo_escaneado TEXT,
                        registrado_por TEXT
                    );
                """))

    except Exception as e:
        print(f"⚠️ Error inicializando base de datos: {e}")

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
    nombre_clean = nombre_completo.strip().upper()
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
            "n": nombre_clean,
            "r": rol,
            "p": password
        })
    return True, f"Usuario '{usr_clean}' creado con éxito."

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
    st.cache_data.clear()
    return True, f"Oficio ID #{id_oficio} eliminado."

def existe_folio_oficio(dgcat, excluir_id=None):
    """Busca si un folio DGCAT ya existe en la tabla oficios. Devuelve el id existente o None."""
    engine = get_engine()
    dgcat_upper = str(dgcat).strip().upper()
    with engine.connect() as conn:
        if excluir_id:
            res = conn.execute(
                text("SELECT id FROM oficios WHERE UPPER(TRIM(dgcat)) = :dg AND id != :ex"),
                {"dg": dgcat_upper, "ex": int(excluir_id)}
            ).fetchone()
        else:
            res = conn.execute(
                text("SELECT id FROM oficios WHERE UPPER(TRIM(dgcat)) = :dg"),
                {"dg": dgcat_upper}
            ).fetchone()
        return res[0] if res else None

def guardar_oficio(datos, id_oficio=None):
    """
    Inserta o actualiza un registro en 'oficios'.
    datos: dict con las llaves de columna -> valor.
    id_oficio: si se provee, actualiza ese registro; si no, inserta uno nuevo.
    Devuelve (ok, mensaje). El commit ocurre dentro del 'with'; el rerun de
    Streamlit se debe llamar DESPUÉS de que esta función retorne, nunca dentro
    del bloque de transacción (st.rerun() lanza una excepción que provocaría
    ROLLBACK si se ejecuta dentro de un 'with engine.begin()').
    """
    engine = get_engine()
    try:
        with engine.begin() as conn:
            if id_oficio:
                set_clause = ", ".join([f"{col} = :{col}" for col in datos.keys()])
                params = dict(datos)
                params["id"] = int(id_oficio)
                conn.execute(text(f"UPDATE oficios SET {set_clause} WHERE id = :id"), params)
            else:
                cols = ", ".join(datos.keys())
                placeholders = ", ".join([f":{c}" for c in datos.keys()])
                conn.execute(text(f"INSERT INTO oficios ({cols}) VALUES ({placeholders})"), datos)
        st.cache_data.clear()
        return True, "Registro guardado correctamente."
    except Exception as e:
        return False, f"Error al guardar en la base de datos: {e}"

# -----------------------------------------------------------------------------
# SEGUIMIENTO DE UBICACIÓN DE PREDIO (base independiente de oficios)
# -----------------------------------------------------------------------------
def guardar_seguimiento_predio(datos, id_registro=None):
    engine = get_engine()
    try:
        with engine.begin() as conn:
            if id_registro:
                set_clause = ", ".join([f"{col} = :{col}" for col in datos.keys()])
                params = dict(datos)
                params["id"] = int(id_registro)
                conn.execute(text(f"UPDATE seguimiento_predio SET {set_clause} WHERE id = :id"), params)
            else:
                cols = ", ".join(datos.keys())
                placeholders = ", ".join([f":{c}" for c in datos.keys()])
                conn.execute(text(f"INSERT INTO seguimiento_predio ({cols}) VALUES ({placeholders})"), datos)
        st.cache_data.clear()
        return True, "Registro de seguimiento guardado correctamente."
    except Exception as e:
        return False, f"Error al guardar seguimiento: {e}"

def eliminar_seguimiento_predio(id_registro):
    engine = get_engine()
    try:
        with engine.begin() as conn:
            conn.execute(text("DELETE FROM seguimiento_predio WHERE id = :id"), {"id": int(id_registro)})
        st.cache_data.clear()
        return True, f"Registro de seguimiento ID #{id_registro} eliminado."
    except Exception as e:
        return False, f"Error al eliminar: {e}"

def existe_folio_seguimiento(dgcat, excluir_id=None):
    engine = get_engine()
    dgcat_upper = str(dgcat).strip().upper()
    with engine.connect() as conn:
        if excluir_id:
            res = conn.execute(
                text("SELECT id FROM seguimiento_predio WHERE UPPER(TRIM(dgcat)) = :dg AND id != :ex"),
                {"dg": dgcat_upper, "ex": int(excluir_id)}
            ).fetchone()
        else:
            res = conn.execute(
                text("SELECT id FROM seguimiento_predio WHERE UPPER(TRIM(dgcat)) = :dg"),
                {"dg": dgcat_upper}
            ).fetchone()
        return res[0] if res else None

# Relación entre cada tabla catálogo y la columna de 'oficios' que la usa,
# para poder mantener sincronizados los registros ya capturados cuando se
# renombra o elimina una opción del catálogo.
CATALOGO_COLUMNA_OFICIOS = {
    "cat_scg": "scg",
    "cat_siscat": "siscat",
    "cat_sistemas_or": "sistemas_or",
    "cat_tramite": "tipo_tramite",
}

def actualizar_opcion_catalogo(tabla, nombre_antiguo, nombre_nuevo):
    engine = get_engine()
    nuevo_clean = nombre_nuevo.strip().upper()
    antiguo_clean = nombre_antiguo.strip().upper()

    try:
        with engine.begin() as conn:
            check = conn.execute(
                text(f"SELECT nombre FROM {tabla} WHERE UPPER(nombre) = :n"),
                {"n": nuevo_clean}
            ).fetchone()

            if check and nuevo_clean != antiguo_clean:
                return False, f"La opción '{nuevo_clean}' ya existe en el catálogo."

            conn.execute(
                text(f"UPDATE {tabla} SET nombre = :nuevo WHERE UPPER(nombre) = :antiguo"),
                {"nuevo": nuevo_clean, "antiguo": antiguo_clean}
            )

            # Propaga el cambio a los oficios que ya usaban el valor anterior,
            # para que no queden registros con un valor de catálogo "huérfano".
            afectados = 0
            columna = CATALOGO_COLUMNA_OFICIOS.get(tabla)
            if columna:
                result = conn.execute(
                    text(f"UPDATE oficios SET {columna} = :nuevo WHERE UPPER({columna}) = :antiguo"),
                    {"nuevo": nuevo_clean, "antiguo": antiguo_clean}
                )
                afectados = result.rowcount or 0

        st.cache_data.clear()
        extra = f" Se actualizaron {afectados} oficio(s) que usaban este valor." if afectados else ""
        return True, f"Opción actualizada a '{nuevo_clean}'.{extra}"
    except Exception as e:
        return False, f"Error al actualizar: {e}"

def eliminar_opcion_catalogo(tabla, nombre_opcion):
    engine = get_engine()
    opcion_clean = nombre_opcion.strip().upper()
    try:
        with engine.begin() as conn:
            conn.execute(
                text(f"DELETE FROM {tabla} WHERE UPPER(nombre) = :n"),
                {"n": opcion_clean}
            )
        st.cache_data.clear()
        return True, f"Opción '{opcion_clean}' eliminada del catálogo."
    except Exception as e:
        return False, f"Error al eliminar: {e}"

def contar_oficios_con_valor_catalogo(tabla, nombre_opcion):
    """Cuenta cuántos oficios ya capturados usan este valor de catálogo,
    para advertir al usuario antes de eliminarlo."""
    columna = CATALOGO_COLUMNA_OFICIOS.get(tabla)
    if not columna:
        return 0
    engine = get_engine()
    opcion_clean = nombre_opcion.strip().upper()
    with engine.connect() as conn:
        res = conn.execute(
            text(f"SELECT COUNT(*) FROM oficios WHERE UPPER({columna}) = :n"),
            {"n": opcion_clean}
        ).fetchone()
        return res[0] if res else 0

# -----------------------------------------------------------------------------
# REPORTE EJECUTIVO EXCEL
# -----------------------------------------------------------------------------
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

    # TITULO DE RESUMEN EJECUTIVO
    ws_sum.cell(row=1, column=1, value="DIRECCIÓN GENERAL DE CATASTRO").font = font_title
    ws_sum.cell(row=2, column=1, value=f"REPORTE DE CONTROL DE GESTIÓN | FECHA: {datetime.now().strftime('%d/%m/%Y')}").font = font_sub

    # ENCABEZADOS DE MÉTRICAS
    ws_sum.cell(row=4, column=1, value="Métrica Directiva").font = font_header
    ws_sum.cell(row=4, column=1).fill = fill_header
    ws_sum.cell(row=4, column=2, value="Valor").font = font_header
    ws_sum.cell(row=4, column=2).fill = fill_header

    total = len(df)
    df_upper = df.copy()
    df_upper.columns = [c.lower() for c in df_upper.columns]

    scg_col = 'scg' if 'scg' in df_upper.columns else None
    sis_col = 'siscat' if 'siscat' in df_upper.columns else None
    pdf_col = 'archivo_escaneado' if 'archivo_escaneado' in df_upper.columns else None
    sis_or_col = 'sistemas_or' if 'sistemas_or' in df_upper.columns else None

    scg_conc = len(df_upper[df_upper[scg_col] == 'CONCLUIDO']) if scg_col else 0
    sis_conc = len(df_upper[df_upper[sis_col].isin(['CONCLUIDO', 'SUBIDO'])]) if sis_col else 0
    pdfs_subidos = len(df_upper[df_upper[pdf_col].notna() & (df_upper[pdf_col] != '') & (df_upper[pdf_col] != 'NONE')]) if pdf_col else 0
    pdfs_pendientes = total - pdfs_subidos
    sistemas_or_count = len(df_upper[df_upper[sis_or_col].notna() & (df_upper[sis_or_col] != '') & (df_upper[sis_or_col] != 'NONE')]) if sis_or_col else 0

    metricas = [
        ("Total de Oficios Atendidos", total),
        ("Oficios Registrados SISTEMAS/OR", sistemas_or_count),
        ("Expedientes PDF Subidos", pdfs_subidos),
        ("Expedientes PDF Pendientes", pdfs_pendientes),
        ("% Digitalización PDF", f"{round((pdfs_subidos/total*100), 1)}%" if total > 0 else "0%"),
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

    # TABLA DE DESGLOSE POR BANDEJA SCG
    start_r = 16
    ws_sum.cell(row=start_r, column=1, value="Bandeja SCG").font = font_header
    ws_sum.cell(row=start_r, column=1).fill = fill_header
    ws_sum.cell(row=start_r, column=2, value="Cantidad").font = font_header
    ws_sum.cell(row=start_r, column=2).fill = fill_header

    if scg_col and not df_upper.empty:
        scg_counts = df_upper['scg'].value_counts()
        for i, (b_name, b_val) in enumerate(scg_counts.items(), start=start_r + 1):
            c1 = ws_sum.cell(row=i, column=1, value=str(b_name).upper())
            c2 = ws_sum.cell(row=i, column=2, value=int(b_val))
            c1.font = font_regular
            c2.font = font_bold
            c1.border = border_box
            c2.border = border_box
            c1.alignment = Alignment(vertical="center")
            c2.alignment = Alignment(horizontal="center", vertical="center")

    # DETALLE GENERAL
    headers = [
        "ID", "ID REGISTRO", "ESTADO", "MUNICIPIO", "EJIDO", 
        "NO. OFICIO", "DGCAT", "FECHA ENTREGA", "FECHA RECIBIDO", 
        "SCG", "SISCAT", "SISTEMAS/OR", "TIPO TRÁMITE", "OBSERVACIONES", "ARCHIVO ESCANEADO"
    ]
    
    ws_det.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws_det.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cols_df = ['id', 'id_registro', 'estado', 'municipio', 'ejido', 'no_oficio', 'dgcat', 'fecha_entrega', 'fecha_recibido', 'scg', 'siscat', 'sistemas_or', 'tipo_tramite', 'observaciones', 'archivo_escaneado']

    for row_idx, row in df_upper.iterrows():
        row_data = [str(row.get(c, '')).upper() if pd.notna(row.get(c, '')) else '' for c in cols_df]
        ws_det.append(row_data)
        
        current_row = row_idx + 2
        is_zebra = (row_idx % 2 == 1)
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws_det.cell(row=current_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_box
            if is_zebra:
                cell.fill = fill_zebra
            
            if col_idx in [1, 2, 8, 9, 10, 11, 12]:
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

# -----------------------------------------------------------------------------
# REPORTE EJECUTIVO EXCEL - SEGUIMIENTO DE UBICACIÓN DE PREDIO
# -----------------------------------------------------------------------------
def generar_excel_seguimiento(df, filename="Reporte_Seguimiento_Predio.xlsx"):
    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Resumen"
    ws_det = wb.create_sheet(title="Detalle")

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
    ws_sum.cell(row=2, column=1, value=f"SEGUIMIENTO DE UBICACIÓN DE PREDIO | FECHA: {datetime.now().strftime('%d/%m/%Y')}").font = font_sub

    df_upper = df.copy()
    df_upper.columns = [c.lower() for c in df_upper.columns]
    total = len(df_upper)
    pdf_col = 'archivo_escaneado' if 'archivo_escaneado' in df_upper.columns else None
    con_pdf = len(df_upper[df_upper[pdf_col].notna() & (df_upper[pdf_col] != '') & (df_upper[pdf_col] != 'NONE')]) if pdf_col else 0

    metricas = [
        ("Total de Predios en Seguimiento", total),
        ("Con Expediente PDF", con_pdf),
        ("Pendientes de PDF", total - con_pdf),
        ("% Digitalización", f"{round((con_pdf/total*100), 1)}%" if total > 0 else "0%"),
    ]
    ws_sum.cell(row=4, column=1, value="Métrica").font = font_header
    ws_sum.cell(row=4, column=1).fill = fill_header
    ws_sum.cell(row=4, column=2, value="Valor").font = font_header
    ws_sum.cell(row=4, column=2).fill = fill_header
    for idx, (m, v) in enumerate(metricas, start=5):
        c1 = ws_sum.cell(row=idx, column=1, value=m)
        c2 = ws_sum.cell(row=idx, column=2, value=v)
        c1.font, c2.font = font_regular, font_bold
        c1.border = c2.border = border_box

    headers = ["ID", "DGCAT/FOLIO", "ESTADO", "MUNICIPIO", "EJIDO", "FECHA REGISTRO", "OBSERVACIONES", "ARCHIVO ESCANEADO", "REGISTRADO POR"]
    ws_det.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws_det.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cols_df = ['id', 'dgcat', 'estado', 'municipio', 'ejido', 'fecha_registro', 'observaciones', 'archivo_escaneado', 'registrado_por']
    for row_idx, row in df_upper.iterrows():
        row_data = [str(row.get(c, '')).upper() if pd.notna(row.get(c, '')) else '' for c in cols_df]
        ws_det.append(row_data)
        current_row = row_idx + 2
        is_zebra = (row_idx % 2 == 1)
        for col_idx in range(1, len(headers) + 1):
            cell = ws_det.cell(row=current_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_box
            if is_zebra:
                cell.fill = fill_zebra

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
