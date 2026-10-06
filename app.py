import streamlit as st
import pandas as pd
import os
import base64
import plotly.express as px
from datetime import date, datetime
from sqlalchemy import text
from database import (
    init_db, 
    verificar_login, 
    registrar_nuevo_usuario, 
    get_engine, 
    generar_excel_ejecutivo,
    generar_excel_seguimiento,
    eliminar_oficio,
    actualizar_opcion_catalogo,
    eliminar_opcion_catalogo,
    contar_oficios_con_valor_catalogo,
    eliminar_usuario,
    cambiar_password_usuario,
    existe_folio_oficio,
    guardar_oficio,
    guardar_seguimiento_predio,
    eliminar_seguimiento_predio,
    existe_folio_seguimiento,
    modificar_usuario,
    agregar_archivo_adjunto,
    obtener_archivos_adjuntos,
    eliminar_archivo_adjunto,
    obtener_pagina,
    obtener_auditoria,
    mostrar_pdf,
)

st.set_page_config(
    page_title="DGCAT - Control de Gestión Institucional",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .header-box { text-align: center; padding: 10px 0 15px 0; margin-bottom: 15px; }
    .header-title { font-size: 2.2rem; font-weight: 800; margin-bottom: 4px; }
    .header-subtitle { color: #10B981; font-size: 1.15rem; font-weight: 600; margin-bottom: 15px; }
    .header-line { height: 3px; background: linear-gradient(90deg, transparent 0%, #10B981 50%, transparent 100%); border: none; margin-bottom: 25px; }

    .stSelectbox label, .stTextInput label, .stDateInput label, .stTextArea label {
        font-weight: 700 !important; font-size: 0.95rem !important;
    }

    .metric-card {
        background: rgba(16, 185, 129, 0.08);
        border: 1px solid rgba(16, 185, 129, 0.3);
        border-top: 4px solid #10B981;
        border-radius: 10px;
        padding: 16px; text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .metric-card h3 { font-size: 0.82rem; margin-bottom: 4px; text-transform: uppercase; opacity: 0.8; }
    .metric-card .number { font-size: 2.1rem; font-weight: 800; color: #10B981; }
    .metric-card .subtitle { font-size: 0.78rem; opacity: 0.75; margin-top: 4px; }

    .session-badge {
        background-color: #10B981; color: #FFFFFF; padding: 10px 18px; border-radius: 8px;
        font-weight: 700; font-size: 0.95rem; margin-bottom: 25px; display: flex;
        justify-content: space-between; align-items: center;
    }

    [data-testid="stSidebarHeader"] img, [data-testid="stImage"] img {
        max-height: 100px !important; object-fit: contain !important;
    }
</style>
""", unsafe_allow_html=True)

UPLOADS_DIR = "uploads"

_MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

def render_documentos_adicionales(modulo, registro_id, key_prefix):
    st.markdown("**📎 Documentos adicionales**")
    adjuntos = obtener_archivos_adjuntos(modulo, registro_id)
    if not adjuntos:
        st.caption("Sin documentos adicionales para este registro.")
        return

    for adj in adjuntos:
        adj_bytes, adj_ruta, adj_error = cargar_pdf_bytes(adj['nombre_archivo'])
        ver_key = f"ver_{key_prefix}_{adj['id']}"

        ac1, ac2, ac3, ac4 = st.columns([3, 1, 1, 1])
        with ac1:
            st.write(f"📄 {adj['nombre_para_mostrar']} — subido {adj['fecha_subida']} por {adj['subido_por']}")
        with ac2:
            if adj_bytes is not None:
                if st.button("👁️ Ver", key=f"btn_{ver_key}"):
                    st.session_state[ver_key] = not st.session_state.get(ver_key, False)
            else:
                st.caption("📌 Favor de subir el oficio.")
        with ac3:
            if adj_bytes is not None:
                st.download_button("📥 Descargar", adj_bytes, file_name=adj['nombre_para_mostrar'], mime="application/pdf", key=f"dl_{key_prefix}_{adj['id']}")
        with ac4:
            if st.button("🗑️ Eliminar", key=f"del_{key_prefix}_{adj['id']}"):
                ok, msg = eliminar_archivo_adjunto(adj['id'])
                if ok:
                    flash("success", msg)
                    st.rerun()
                else:
                    st.error(msg)

        if adj_bytes is not None and st.session_state.get(ver_key, False):
            mostrar_pdf(adj_ruta)

def flash(tipo, mensaje):
    st.session_state.setdefault("_flash", []).append((tipo, mensaje))

def mostrar_flash():
    mensajes = st.session_state.pop("_flash", [])
    for tipo, mensaje in mensajes:
        if tipo == "success":
            st.success(mensaje)
        elif tipo == "error":
            st.error(mensaje)
        elif tipo == "warning":
            st.warning(mensaje)
        else:
            st.info(mensaje)

def es_pdf_valido(datos_archivo):
    if not datos_archivo:
        return False, "No se recibió ningún archivo."
    if len(datos_archivo) < 100:
        return False, "El archivo está vacío o es demasiado pequeño para ser un PDF válido."
    if not datos_archivo.lstrip()[:5] == b"%PDF-":
        return False, "El archivo no tiene un formato PDF válido (falta la firma %PDF-)."
    try:
        import io
        from pypdf import PdfReader
        lector = PdfReader(io.BytesIO(datos_archivo))
        _ = len(lector.pages)
    except Exception:
        return False, "El archivo parece dañado o incompleto y no pudo abrirse como PDF."
    return True, ""

def normalizar_para_mostrar(df, columnas_preservar=('id',)):
    df = df.copy()
    for col in df.columns:
        if col in columnas_preservar:
            continue
        serie = df[col].astype(str).str.upper().str.strip()
        vacíos = serie.isin(['NONE', 'NAN', '', '<NA>'])
        df[col] = serie.where(~vacíos, '—')
    return df

def fecha_larga_es(fecha=None):
    f = fecha or date.today()
    return f"{f.day} de {_MESES_ES[f.month]} de {f.year}"

def resolver_ruta_pdf(nombre_archivo):
    if nombre_archivo is None:
        return None
    nombre_limpio = str(nombre_archivo).strip()
    if not nombre_limpio or nombre_limpio.upper() in ('NONE', 'NAN', ''):
        return None

    ruta_directa = os.path.join(UPLOADS_DIR, nombre_limpio)
    if os.path.exists(ruta_directa) and os.path.isfile(ruta_directa):
        return ruta_directa

    try:
        objetivo = nombre_limpio.lower()
        for f in os.listdir(UPLOADS_DIR):
            if f.lower() == objetivo:
                return os.path.join(UPLOADS_DIR, f)
    except Exception:
        pass
    return None

def cargar_pdf_bytes(nombre_archivo):
    ruta = resolver_ruta_pdf(nombre_archivo)
    if ruta is None:
        return None, None, "no_existe"
    try:
        with open(ruta, "rb") as f:
            contenido = f.read()
        if not contenido:
            return None, ruta, "archivo_vacio"
        return contenido, ruta, None
    except Exception as e:
        return None, ruta, str(e)

os.makedirs(UPLOADS_DIR, exist_ok=True)
init_db()

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""
if "nombre" not in st.session_state:
    st.session_state["nombre"] = ""
if "rol" not in st.session_state:
    st.session_state["rol"] = "operador"

# LOGIN
if not st.session_state["authenticated"]:
    st.markdown("""
    <div class="header-box">
        <div class="header-title">🏛️️ DIRECCIÓN GENERAL DE CATASTRO</div>
        <div class="header-subtitle">Sistema de Control de Entrada y Salida de Oficios de Respuesta</div>
        <hr class="header-line">
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.8, 1])
    with col2:
        st.write("### 🔑 Acceso Institucional")
        usuario_input = st.text_input("Usuario", key="input_usr")
        password_input = st.text_input("Contraseña", type="password", key="input_pwd")
        
        if st.button("Ingresar al Sistema", use_container_width=True, type="primary"):
            if not usuario_input or not password_input:
                st.warning("Ingrese su usuario y contraseña.")
            else:
                user = verificar_login(usuario_input, password_input)
                if user:
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = user[0]
                    st.session_state["nombre"] = user[1]
                    st.session_state["rol"] = user[2]
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas. Verifique sus datos.")
    st.stop()

# PANEL INTERNO
st.markdown(f"""
<div class="session-badge">
    <span>🟢 SESIÓN ACTIVA | <strong>{st.session_state['nombre']}</strong> ({st.session_state['username']})</span>
    <span>PERFIL: <strong>{st.session_state['rol'].upper()}</strong></span>
</div>
""", unsafe_allow_html=True)

if os.path.exists("logo_ran.png"):
    st.sidebar.image("logo_ran.png", use_container_width=True)
elif os.path.exists("Logo_RAN.png"):
    st.sidebar.image("Logo_RAN.png", use_container_width=True)

st.sidebar.title(f"👤 {st.session_state['nombre']}")
st.sidebar.caption(f"ROL: {st.session_state['rol'].upper()}")

if st.sidebar.button("🔒 Cerrar Sesión"):
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""
    st.session_state["nombre"] = ""
    st.session_state["rol"] = "operador"
    st.rerun()

st.sidebar.markdown("---")
engine = get_engine()

ROL_ACTUAL = st.session_state["rol"]

if ROL_ACTUAL == "operador":
    menu_options = ["📍 Seguimiento de Ubicación de Predio"]
elif ROL_ACTUAL == "supervisor":
    menu_options = [
        "📈 Dashboard Ejecutivo",
        "📝 Registro Completo de Oficios",
        "📍 Seguimiento de Ubicación de Predio",
        "🔍 Consulta y Expedientes",
        "🗂️ Consulta de Seguimiento de Predio",
        "⚙️ Gestión de Catálogos",
    ]
else:
    menu_options = [
        "📈 Dashboard Ejecutivo", 
        "📝 Registro Completo de Oficios", 
        "📍 Seguimiento de Ubicación de Predio", 
        "🔍 Consulta y Expedientes", 
        "🗂️️ Consulta de Seguimiento de Predio",
        "⚙️ Gestión de Catálogos",
        "👥 Alta de Usuarios"
    ]

menu = st.sidebar.radio("Menú de Opciones", menu_options)
mostrar_flash()

@st.cache_data(ttl=600)
def get_cat_ubicaciones_df():
    try:
        df = pd.read_sql("SELECT * FROM cat_ubicaciones", engine)
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        col_edo = [c for c in df.columns if 'estado' in c or 'edo' in c][0] if any('estado' in c or 'edo' in c for c in df.columns) else df.columns[0]
        col_mun = [c for c in df.columns if 'muni' in c][0] if any('muni' in c for c in df.columns) else df.columns[1]
        col_eji = [c for c in df.columns if 'nucleo' in c or 'ejido' in c or 'nuc' in c][0] if any('nucleo' in c or 'ejido' in c or 'nuc' in c for c in df.columns) else df.columns[2]
        
        res_df = pd.DataFrame({
            'estado': df[col_edo].astype(str).str.strip().str.upper(),
            'municipio': df[col_mun].astype(str).str.strip().str.upper(),
            'ejido': df[col_eji].astype(str).str.strip().str.upper()
        })
        return res_df
    except Exception:
        return pd.DataFrame(columns=['estado', 'municipio', 'ejido'])

def get_estados():
    df = get_cat_ubicaciones_df()
    if df.empty: return []
    return sorted(list(df['estado'].dropna().unique()))

def get_municipios(estado):
    df = get_cat_ubicaciones_df()
    if df.empty: return []
    filtered = df[df['estado'].str.upper() == str(estado).strip().upper()]
    return sorted(list(filtered['municipio'].dropna().unique()))

def get_ejidos(estado, municipio):
    df = get_cat_ubicaciones_df()
    if df.empty: return []
    filtered = df[
        (df['estado'].str.upper() == str(estado).strip().upper()) & 
        (df['municipio'].str.upper() == str(municipio).strip().upper())
    ]
    return sorted(list(filtered['ejido'].dropna().unique()))

# -----------------------------------------------------------------------------
# 1. DASHBOARD EJECUTIVO
# -----------------------------------------------------------------------------
if menu == "📈 Dashboard Ejecutivo":
    col_titulo, col_refrescar = st.columns([5, 1])
    with col_titulo:
        st.title("🏛️ Tablero de Control Directivo DGCAT")
        st.caption("Monitoreo institucional de oficios de respuesta, expedientes PDF digitalizados, SCG y SISCAT.")
    with col_refrescar:
        st.write("")
        if st.button("🔄 Actualizar ahora", use_container_width=True, help="Fuerza a traer los datos más recientes de la base."):
            st.cache_data.clear()
            st.session_state["_dashboard_recien_actualizado"] = True
            st.rerun()

    if st.session_state.pop("_dashboard_recien_actualizado", False):
        st.success("✅ Información actualizada correctamente.")

    df_raw = pd.read_sql("SELECT * FROM oficios", engine)
    df_raw.columns = [c.lower() for c in df_raw.columns]
    
    if df_raw.empty:
        st.warning("No hay registros disponibles para generar métricas.")
        st.stop()

    with st.expander("🔍 **Filtros de Control Ejecutivo**", expanded=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            estados_opts = ["TODOS"] + sorted(list(df_raw['estado'].dropna().unique()))
            sel_estado = st.selectbox("Estado", estados_opts)
        with f_col2:
            scg_opts = ["TODOS"] + sorted(list(df_raw['scg'].dropna().unique()))
            sel_scg = st.selectbox("Estatus SCG", scg_opts)
        with f_col3:
            siscat_opts = ["TODOS"] + sorted(list(df_raw['siscat'].dropna().unique()))
            sel_siscat = st.selectbox("Estatus SISCAT", siscat_opts)

    df_filtered = df_raw.copy()
    if sel_estado != "TODOS": df_filtered = df_filtered[df_filtered['estado'] == sel_estado]
    if sel_scg != "TODOS": df_filtered = df_filtered[df_filtered['scg'] == sel_scg]
    if sel_siscat != "TODOS": df_filtered = df_filtered[df_filtered['siscat'] == sel_siscat]

    st.markdown("---")
    
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    total_oficios = len(df_filtered)
    concluidos_scg = len(df_filtered[df_filtered['scg'] == 'CONCLUIDO'])
    subidos_siscat = len(df_filtered[df_filtered['siscat'].isin(['CONCLUIDO', 'SUBIDO'])])
    en_sistemas = len(df_filtered[df_filtered['scg'].astype(str).str.contains('SISTEMAS', case=False, na=False)])
    sistemas_or_val = len(df_filtered[df_filtered['sistemas_or'].notna() & (df_filtered['sistemas_or'].astype(str).str.strip() != '') & (df_filtered['sistemas_or'].astype(str).str.upper() != 'NONE')]) if 'sistemas_or' in df_filtered.columns else 0
    
    pct_scg = round((concluidos_scg / total_oficios * 100), 1) if total_oficios > 0 else 0
    pct_siscat = round((subidos_siscat / total_oficios * 100), 1) if total_oficios > 0 else 0

    with kpi1: st.markdown(f'<div class="metric-card"><h3>TOTAL OFICIOS</h3><div class="number">{total_oficios:,}</div><div class="subtitle">Registros Atendidos</div></div>', unsafe_allow_html=True)
    with kpi2: st.markdown(f'<div class="metric-card"><h3>EFICIENCIA SCG</h3><div class="number">{pct_scg}%</div><div class="subtitle">{concluidos_scg:,} Concluidos</div></div>', unsafe_allow_html=True)
    with kpi3: st.markdown(f'<div class="metric-card"><h3>AVANCE SISCAT</h3><div class="number">{pct_siscat}%</div><div class="subtitle">{subidos_siscat:,} Procesados</div></div>', unsafe_allow_html=True)
    with kpi4: st.markdown(f'<div class="metric-card"><h3>EN SISTEMAS</h3><div class="number" style="color:#EF4444;">{en_sistemas:,}</div><div class="subtitle">Pendientes</div></div>', unsafe_allow_html=True)
    with kpi5: st.markdown(f'<div class="metric-card" style="border-top-color:#8B5CF6;"><h3>SISTEMAS / OR</h3><div class="number" style="color:#8B5CF6;">{sistemas_or_val:,}</div><div class="subtitle">Trámite Clasificado</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.subheader("📄 Control Digital de Expedientes PDF")
    has_pdf = df_filtered['archivo_escaneado'].notna() & (df_filtered['archivo_escaneado'] != '') & (df_filtered['archivo_escaneado'] != 'NONE')
    pdf_subidos = len(df_filtered[has_pdf])
    pdf_pendientes = total_oficios - pdf_subidos
    pct_pdf = round((pdf_subidos / total_oficios * 100), 1) if total_oficios > 0 else 0

    pdf_col1, pdf_col2 = st.columns([1, 2])
    with pdf_col1:
        st.markdown(f'<div class="metric-card" style="border-top-color: #3B82F6;"><h3>DIGITALIZACIÓN PDF</h3><div class="number" style="color:#3B82F6;">{pct_pdf}%</div><div class="subtitle">{pdf_subidos:,} Subidos | {pdf_pendientes:,} Pendientes</div></div>', unsafe_allow_html=True)

    with pdf_col2:
        df_pdf_chart = pd.DataFrame({
            'Estatus Expediente': ['CON EXPEDIENTE PDF', 'PENDIENTE DE PDF'],
            'Cantidad': [pdf_subidos, pdf_pendientes]
        })
        fig_pdf = px.pie(
            df_pdf_chart, values='Cantidad', names='Estatus Expediente', hole=0.5, 
            color='Estatus Expediente', color_discrete_map={'CON EXPEDIENTE PDF': '#10B981', 'PENDIENTE DE PDF': '#EF4444'}
        )
        fig_pdf.update_traces(textposition='inside', textinfo='percent+value')
        fig_pdf.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=200, showlegend=True, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", font_color="#1F2937")
        st.plotly_chart(fig_pdf, use_container_width=True)

    st.markdown("---")

    g_col1, g_col2 = st.columns(2)
    with g_col1:
        st.subheader("🍩 Distribución por Bandeja SCG")
        scg_counts = df_filtered['scg'].value_counts().reset_index()
        scg_counts.columns = ['Bandeja', 'Cantidad']
        fig_pie_scg = px.pie(
            scg_counts, values='Cantidad', names='Bandeja',
            color_discrete_sequence=px.colors.qualitative.G10
        )
        fig_pie_scg.update_traces(textinfo='label+value', textposition='inside')
        fig_pie_scg.update_layout(showlegend=True, margin=dict(t=20, b=20, l=20, r=20), paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", font_color="#1F2937")
        st.plotly_chart(fig_pie_scg, use_container_width=True)

    with g_col2:
        st.subheader("📊 Control e Integración en SISCAT")
        siscat_counts = df_filtered['siscat'].value_counts().reset_index()
        siscat_counts.columns = ['Estatus', 'Cantidad']
        
        fig_bar_siscat = px.bar(
            siscat_counts, x='Cantidad', y='Estatus', orientation='h', color='Estatus', 
            text='Cantidad', color_discrete_sequence=px.colors.qualitative.Vivid
        )
        fig_bar_siscat.update_layout(showlegend=False, xaxis_title="Número de Oficios", yaxis_title="", paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", font_color="#1F2937")
        st.plotly_chart(fig_bar_siscat, use_container_width=True)

    st.markdown("---")

    g_col3, g_col4 = st.columns(2)
    with g_col3:
        st.subheader("🇲🇽 Top 10 Estados con Mayor Carga Registrada")
        top_estados = df_filtered['estado'].value_counts().head(10).reset_index()
        top_estados.columns = ['Estado', 'Oficios']
        fig_top_estados = px.bar(
            top_estados, x='Estado', y='Oficios', color='Oficios', text='Oficios', 
            color_continuous_scale='Greens'
        )
        fig_top_estados.update_traces(
            textposition='outside', textfont_color='#111827', textfont_size=13, cliponaxis=False
        )
        fig_top_estados.update_layout(
            xaxis_title="", yaxis_title="Total Oficios",
            coloraxis_showscale=False, xaxis_tickangle=-35,
            margin=dict(t=40, b=10),
            paper_bgcolor='#FFFFFF', plot_bgcolor='#FFFFFF',
            font_color='#1F2937',
            yaxis=dict(range=[0, top_estados['Oficios'].max() * 1.18] if not top_estados.empty else None),
        )
        st.plotly_chart(fig_top_estados, use_container_width=True)

    with g_col4:
        st.subheader("📑 Volumetría por Tipo de Trámite")
        
        def inferir_tramite(row):
            t_actual = str(row.get('tipo_tramite', '')).strip().upper()
            if t_actual and t_actual not in ['NONE', 'NAN', '']:
                return t_actual
            obs = str(row.get('observaciones', '')).upper()
            scg = str(row.get('scg', '')).upper()
            txt = obs + " " + scg
            if 'DESTINO' in txt: return 'CAMBIO DE DESTINO'
            if 'SUPERFICIE' in txt: return 'CAMBIO DE SUPERFICIE'
            if 'DOMINIO' in txt: return 'DOMINIO PLENO'
            if 'MOSAICO' in txt: return 'ACT. DE MOSAICO'
            if 'SENTENCIA' in txt: return 'SENTENCIA'
            return 'PENDIENTE DE CLASIFICAR'

        tramites_series = df_filtered.apply(inferir_tramite, axis=1)
        obs_counts = tramites_series.value_counts().reset_index()
        obs_counts.columns = ['Trámite', 'Cantidad']
        
        fig_tramite_bar = px.bar(
            obs_counts, x='Cantidad', y='Trámite', orientation='h', 
            color='Cantidad', text='Cantidad', color_continuous_scale='blues'
        )
        fig_tramite_bar.update_layout(
            xaxis_title="Número de Oficios", yaxis_title="", 
            coloraxis_showscale=False, yaxis=dict(autorange="reversed"),
            paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", font_color="#1F2937"
        )
        st.plotly_chart(fig_tramite_bar, use_container_width=True)

# -----------------------------------------------------------------------------
# 2. SEGUIMIENTO DE UBICACIÓN DE PREDIO
# -----------------------------------------------------------------------------
elif menu == "📍 Seguimiento de Ubicación de Predio":
    st.title("📍 Seguimiento de Ubicación de Predio")

    df_seg_todo = pd.read_sql("SELECT * FROM seguimiento_predio ORDER BY id DESC", engine)
    df_seg_todo.columns = [c.lower() for c in df_seg_todo.columns]

    modo_seg = st.radio("Modo:", ["➕ Nuevo Registro", "✏ Modificar Registro Existente", "🗑️ Eliminar Registro"], horizontal=True, key="modo_seg")

    if "reset_count_seg" not in st.session_state:
        st.session_state["reset_count_seg"] = 0

    if st.session_state.get("last_modo_seg") != modo_seg:
        st.session_state["last_modo_seg"] = modo_seg
        st.session_state["reset_count_seg"] += 1
        st.rerun()

    cnt_s = st.session_state["reset_count_seg"]

    texto_busqueda_seg = st.text_input(
        "🔎 Buscar por folio DGCAT, Estado o Municipio (opcional):",
        key=f"buscar_seg_{modo_seg}_{cnt_s}"
    )
    df_busqueda_seg = df_seg_todo
    if texto_busqueda_seg.strip() and not df_seg_todo.empty:
        patron = texto_busqueda_seg.strip().upper()
        df_busqueda_seg = df_seg_todo[
            df_seg_todo['dgcat'].astype(str).str.upper().str.contains(patron, na=False) |
            df_seg_todo['estado'].astype(str).str.upper().str.contains(patron, na=False) |
            df_seg_todo['municipio'].astype(str).str.upper().str.contains(patron, na=False)
        ]

    seg_sel = None
    if modo_seg == "➕ Nuevo Registro":
        if texto_busqueda_seg.strip():
            if df_busqueda_seg.empty:
                st.caption("ℹ️ No se encontró ningún registro parecido; puede continuar con la captura.")
            else:
                st.caption("ℹ️ Coincidencias existentes (solo referencia, no bloquea la captura):")
                st.dataframe(df_busqueda_seg[['id', 'dgcat', 'estado', 'municipio']].head(10), use_container_width=True)
    else:
        if df_seg_todo.empty:
            st.warning("No hay registros de seguimiento en la base de datos.")
            st.stop()
        if df_busqueda_seg.empty:
            st.warning("⚠️️ No se encontró ningún registro que coincida con la búsqueda.")
            st.stop()

        df_busqueda_seg = df_busqueda_seg.copy()
        conteo_rep_seg = {}
        etiquetas_seg_sel = []
        for _, fila_tmp in df_busqueda_seg.iterrows():
            base = f"Folio: {str(fila_tmp['dgcat']).upper()} | {str(fila_tmp['estado']).upper()}"
            conteo_rep_seg[base] = conteo_rep_seg.get(base, 0) + 1
            etiquetas_seg_sel.append(base + (f" ({conteo_rep_seg[base]})" if conteo_rep_seg[base] > 1 else ""))
        df_busqueda_seg['display_name'] = etiquetas_seg_sel
        seleccion_seg = st.selectbox("🔍 Seleccione el Registro:", df_busqueda_seg['display_name'].tolist(), key=f"sel_predio_seg_{cnt_s}")
        id_seg_sel = int(df_busqueda_seg.loc[df_busqueda_seg['display_name'] == seleccion_seg, 'id'].iloc[0])
        seg_sel = df_seg_todo[df_seg_todo['id'] == id_seg_sel].iloc[0]

    if modo_seg == "🗑️ Eliminar Registro" and seg_sel is not None:
        st.error(f"⚠️ Eliminar registro de seguimiento **{seg_sel['dgcat']}**.")
        confirmar_seg = st.checkbox("Confirmo que deseo eliminar este registro de forma DEFINITIVA.", key=f"confirmar_del_seg_{cnt_s}")
        if st.button("🚨 ELIMINAR DEFINITIVAMENTE", type="primary", disabled=not confirmar_seg, key=f"btn_del_seg_{cnt_s}"):
            ok, msg = eliminar_seguimiento_predio(int(seg_sel['id']), usuario=st.session_state["username"])
            if ok:
                st.session_state["reset_count_seg"] += 1
                flash("success", f"✅ {msg}")
                st.rerun()
            else:
                st.error(f"❌ {msg}")
        st.stop()

    contexto_id_seg = str(int(seg_sel['id'])) if seg_sel is not None else "nuevo"
    contexto_key_seg = f"{modo_seg}_{contexto_id_seg}_{cnt_s}"

    default_centrales_seg = (seg_sel is not None and str(seg_sel['estado']).strip().upper() == "OFICINAS CENTRALES")
    es_oficinas_centrales = st.checkbox(
        "🏢 Trámite Perteneciente a OFICINAS CENTRALES",
        value=default_centrales_seg,
        key=f"chk_centrales_seg_{contexto_key_seg}"
    )

    estados_list = get_estados()
    col_u1, col_u2, col_u3 = st.columns(3)

    if es_oficinas_centrales:
        estado_sel, municipio_sel, ejido_sel = "OFICINAS CENTRALES", "OFICINAS CENTRALES", "OFICINAS CENTRALES"
        with col_u1: st.text_input("1. Estado *", value="OFICINAS CENTRALES", disabled=True, key=f"dis_se_{contexto_key_seg}")
        with col_u2: st.text_input("2. Municipio *", value="OFICINAS CENTRALES", disabled=True, key=f"dis_sm_{contexto_key_seg}")
        with col_u3: st.text_input("3. Ejido *", value="OFICINAS CENTRALES", disabled=True, key=f"dis_sej_{contexto_key_seg}")
    else:
        def_estado_idx_seg = estados_list.index(seg_sel['estado']) + 1 if (seg_sel is not None and seg_sel['estado'] in estados_list) else 0
        with col_u1:
            estado_sel = st.selectbox("1. Estado *", ["-- Seleccione --"] + estados_list, index=def_estado_idx_seg, key=f"estado_seg_{contexto_key_seg}")

        muns_list = get_municipios(estado_sel) if estado_sel != "-- Seleccione --" else []
        def_mun_idx_seg = muns_list.index(seg_sel['municipio']) + 1 if (seg_sel is not None and seg_sel['municipio'] in muns_list) else 0
        with col_u2:
            municipio_sel = st.selectbox("2. Municipio *", ["-- Seleccione --"] + muns_list, index=def_mun_idx_seg, key=f"muni_seg_{contexto_key_seg}")

        ejidos_list = get_ejidos(estado_sel, municipio_sel) if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --" else []
        def_ejido_idx_seg = ejidos_list.index(seg_sel['ejido']) + 1 if (seg_sel is not None and seg_sel['ejido'] in ejidos_list) else 0
        with col_u3:
            ejido_sel = st.selectbox("3. Ejido *", ["-- Seleccione --"] + ejidos_list, index=def_ejido_idx_seg, key=f"ejido_seg_{contexto_key_seg}")

    st.markdown("---")
    val_dgcat_seg = str(seg_sel['dgcat']) if (seg_sel is not None and pd.notna(seg_sel['dgcat'])) else "DGCAT/100/"
    dgcat_folio = st.text_input("Folio DGCAT *", value=val_dgcat_seg, key=f"dgcat_seg_{contexto_key_seg}")
    dgcat_check = dgcat_folio.strip().upper()

    val_obs_seg = str(seg_sel['observaciones']) if (seg_sel is not None and pd.notna(seg_sel['observaciones'])) else ""

    with st.form(f"form_seguimiento_predio_{contexto_key_seg}", clear_on_submit=True):
        archivo_escaneado = st.file_uploader(
            "Adjuntar/Reemplazar Expediente Escaneado (PDF)" + (" *" if modo_seg == "➕ Nuevo Registro" else ""),
            type=["pdf"],
            key=f"file_seg_{contexto_key_seg}"
        )
        observaciones_capturista = st.text_area("Observaciones (Opcional)", value=val_obs_seg, key=f"obs_seg_{contexto_key_seg}")

        btn_label_seg = "📤 Guardar Seguimiento de Predio" if modo_seg == "➕ Nuevo Registro" else "✏️ Guardar Cambios"
        if st.form_submit_button(btn_label_seg, type="primary"):
            if not es_oficinas_centrales and (estado_sel == "-- Seleccione --" or municipio_sel == "-- Seleccione --" or ejido_sel == "-- Seleccione --"):
                st.error("⚠️ Debe seleccionar Estado, Municipio y Ejido.")
            elif not dgcat_check or dgcat_check == "DGCAT/100/":
                st.error("⚠️ Debe ingresar un folio DGCAT completo.")
            elif archivo_escaneado is None and modo_seg == "➕ Nuevo Registro":
                st.error("⚠️ Debe adjuntar un archivo PDF escaneado.")
            elif archivo_escaneado is not None and not es_pdf_valido(archivo_escaneado.getvalue())[0]:
                _, motivo_pdf = es_pdf_valido(archivo_escaneado.getvalue())
                st.error(f"❌ El archivo no se pudo guardar: {motivo_pdf} Vuelva a intentar con otro archivo.")
            else:
                obs_upper = observaciones_capturista.strip().upper() if observaciones_capturista else ""
                estado_upper = estado_sel.strip().upper()
                municipio_upper = municipio_sel.strip().upper()
                ejido_upper = ejido_sel.strip().upper()
                fecha_hoy = date.today().strftime('%d/%m/%Y')

                nombre_archivo = None
                if archivo_escaneado is not None:
                    dgcat_clean = dgcat_check.replace("/", "_").replace(" ", "")
                    nombre_archivo = f"{dgcat_clean}_{archivo_escaneado.name}"
                    with open(os.path.join(UPLOADS_DIR, nombre_archivo), "wb") as f:
                        f.write(archivo_escaneado.getbuffer())

                datos = {
                    "dgcat": dgcat_check, "estado": estado_upper, "municipio": municipio_upper,
                    "ejido": ejido_upper, "observaciones": obs_upper,
                    "registrado_por": st.session_state["username"],
                }
                if modo_seg == "✏️ Modificar Registro Existente" and seg_sel is not None:
                    datos["fecha_actualizacion"] = fecha_hoy
                    if nombre_archivo:
                        datos["archivo_escaneado"] = nombre_archivo
                    ok, msg = guardar_seguimiento_predio(datos, id_registro=int(seg_sel['id']), usuario=st.session_state["username"])
                else:
                    datos["fecha_registro"] = fecha_hoy
                    datos["fecha_actualizacion"] = fecha_hoy
                    datos["archivo_escaneado"] = nombre_archivo or ""
                    ok, msg = guardar_seguimiento_predio(datos, usuario=st.session_state["username"])

                if ok:
                    st.session_state["reset_count_seg"] += 1
                    flash("success", "✅ Se registró de forma exitosa." if modo_seg == "➕ Nuevo Registro" else f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

# -----------------------------------------------------------------------------
# 3. REGISTRO COMPLETO DE OFICIOS
# -----------------------------------------------------------------------------
elif menu == "📝 Registro Completo de Oficios":
    st.title("📝 Registro y Edición Avanzada de Oficios")

    df_oficios = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", engine)
    df_oficios.columns = [c.lower() for c in df_oficios.columns]

    modo_accion = st.radio("Modo:", ["➕ Nuevo Registro", "✏️ Modificar Registro Existente", "🗑️ Eliminar Registro"], horizontal=True)

    if "reset_count_oficios" not in st.session_state:
        st.session_state["reset_count_oficios"] = 0

    if st.session_state.get("last_modo_oficios") != modo_accion:
        st.session_state["last_modo_oficios"] = modo_accion
        st.session_state["reset_count_oficios"] += 1
        st.rerun()

    cnt_o = st.session_state["reset_count_oficios"]

    texto_busqueda = st.text_input(
        "🔎 Buscar por folio DGCAT, No. de oficio o Estado (opcional):",
        key=f"buscar_oficios_{modo_accion}_{cnt_o}"
    )
    df_busqueda = df_oficios
    if texto_busqueda.strip() and not df_oficios.empty:
        patron = texto_busqueda.strip().upper()
        df_busqueda = df_oficios[
            df_oficios['dgcat'].astype(str).str.upper().str.contains(patron, na=False) |
            df_oficios['no_oficio'].astype(str).str.upper().str.contains(patron, na=False) |
            df_oficios['estado'].astype(str).str.upper().str.contains(patron, na=False)
        ]

    oficio_sel = None
    if modo_accion == "➕ Nuevo Registro":
        if texto_busqueda.strip():
            if df_busqueda.empty:
                st.caption("ℹ️ No se encontró ningún oficio parecido; puede continuar con la captura.")
            else:
                st.caption("ℹ️ Coincidencias existentes (solo referencia, no bloquea la captura):")
                st.dataframe(df_busqueda[['id', 'dgcat', 'no_oficio', 'estado']].head(10), use_container_width=True)
    else:
        if df_oficios.empty:
            st.warning("No hay registros en la base de datos.")
            st.stop()
        if df_busqueda.empty:
            st.warning("⚠️ No se encontró ningún oficio que coincida con la búsqueda.")
            st.stop()

        df_busqueda = df_busqueda.copy()
        conteo_rep_oficio = {}
        etiquetas_oficio = []
        for _, fila_tmp in df_busqueda.iterrows():
            folio_tmp = str(fila_tmp['dgcat']).upper()
            no_of_tmp = str(fila_tmp['no_oficio']) if pd.notna(fila_tmp['no_oficio']) else ''
            base = f"Folio: {folio_tmp}" + (f" | Oficio: {no_of_tmp}" if no_of_tmp and no_of_tmp.upper() != 'NAN' else "")
            conteo_rep_oficio[base] = conteo_rep_oficio.get(base, 0) + 1
            etiquetas_oficio.append(base + (f" ({conteo_rep_oficio[base]})" if conteo_rep_oficio[base] > 1 else ""))
        df_busqueda['display_name'] = etiquetas_oficio
        opciones_oficios = df_busqueda['display_name'].tolist()
        seleccion = st.selectbox("🔍 Seleccione el Oficio:", opciones_oficios, key=f"sel_oficio_{cnt_o}")
        id_oficio_seleccionado = int(df_busqueda.loc[df_busqueda['display_name'] == seleccion, 'id'].iloc[0])
        oficio_sel = df_oficios[df_oficios['id'] == id_oficio_seleccionado].iloc[0]

    if modo_accion == "🗑️️ Eliminar Registro" and oficio_sel is not None:
        st.error(f"⚠️ Eliminar oficio **{oficio_sel['dgcat']}**.")
        confirmar = st.checkbox("Confirmo que deseo eliminar este registro de forma DEFINITIVA.", key=f"confirm_del_of_{cnt_o}")
        if st.button("🚨 ELIMINAR DEFINITIVAMENTE", type="primary", disabled=not confirmar, key=f"btn_del_of_{cnt_o}"):
            ok, msg = eliminar_oficio(int(oficio_sel['id']), usuario=st.session_state["username"])
            if ok:
                st.session_state["reset_count_oficios"] += 1
                flash("success", f"✅ {msg}")
                st.rerun()
            else:
                st.error(f"❌ {msg}")
        st.stop()

    contexto_id = str(int(oficio_sel['id'])) if oficio_sel is not None else "nuevo"
    contexto_key = f"{modo_accion}_{contexto_id}_{cnt_o}"

    default_centrales = (oficio_sel is not None and str(oficio_sel['estado']).strip().upper() == "OFICINAS CENTRALES")
    es_oficinas_centrales_admin = st.checkbox(
        "🏢 Trámite Perteneciente a OFICINAS CENTRALES",
        value=default_centrales,
        key=f"chk_centrales_{contexto_key}"
    )
    estados_list = get_estados()

    col_u1, col_u2, col_u3 = st.columns(3)
    if es_oficinas_centrales_admin:
        estado_sel, municipio_sel, ejido_sel = "OFICINAS CENTRALES", "OFICINAS CENTRALES", "OFICINAS CENTRALES"
        with col_u1: st.text_input("1. Estado", value="OFICINAS CENTRALES", disabled=True, key=f"dis_oe_{contexto_key}")
        with col_u2: st.text_input("2. Municipio", value="OFICINAS CENTRALES", disabled=True, key=f"dis_om_{contexto_key}")
        with col_u3: st.text_input("3. Ejido", value="OFICINAS CENTRALES", disabled=True, key=f"dis_oej_{contexto_key}")
    else:
        def_estado_idx = estados_list.index(oficio_sel['estado']) + 1 if (oficio_sel is not None and oficio_sel['estado'] in estados_list) else 0
        with col_u1:
            estado_sel = st.selectbox("1. Estado", ["-- Seleccione --"] + estados_list, index=def_estado_idx, key=f"estado_sel_{contexto_key}")

        muns_list = get_municipios(estado_sel) if estado_sel != "-- Seleccione --" else []
        def_mun_idx = muns_list.index(oficio_sel['municipio']) + 1 if (oficio_sel is not None and oficio_sel['municipio'] in muns_list) else 0
        with col_u2:
            municipio_sel = st.selectbox("2. Municipio", ["-- Seleccione --"] + muns_list, index=def_mun_idx, key=f"muni_sel_{contexto_key}")

        ejidos_list = get_ejidos(estado_sel, municipio_sel) if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --" else []
        def_ejido_idx = ejidos_list.index(oficio_sel['ejido']) + 1 if (oficio_sel is not None and oficio_sel['ejido'] in ejidos_list) else 0
        with col_u3:
            ejido_sel = st.selectbox("3. Ejido", ["-- Seleccione --"] + ejidos_list, index=def_ejido_idx, key=f"ejido_sel_{contexto_key}")

    val_dgcat = str(oficio_sel['dgcat']) if (oficio_sel is not None and pd.notna(oficio_sel['dgcat'])) else "DGCAT/100/"
    dgcat_folio = st.text_input("Folio DGCAT", value=val_dgcat, key=f"dgcat_folio_{contexto_key}")
    dgcat_check = dgcat_folio.strip().upper()

    scg_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_scg ORDER BY nombre", engine)['nombre'].tolist()
    siscat_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", engine)['nombre'].tolist()
    sistemas_or_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_sistemas_or ORDER BY nombre", engine)['nombre'].tolist()
    tramite_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", engine)['nombre'].tolist()

    val_id_reg = str(oficio_sel['id_registro']) if (oficio_sel is not None and pd.notna(oficio_sel['id_registro'])) else ""
    val_no_oficio = str(oficio_sel['no_oficio']) if (oficio_sel is not None and pd.notna(oficio_sel['no_oficio'])) else ""
    val_obs = str(oficio_sel['observaciones']) if (oficio_sel is not None and pd.notna(oficio_sel['observaciones'])) else ""

    def _parse_fecha(valor):
        if valor is None or (isinstance(valor, float) and pd.isna(valor)):
            return date.today()
        s = str(valor).strip()
        if not s or s.upper() in ('NONE', 'NAN', ''):
            return date.today()
        try:
            return datetime.strptime(s, '%d/%m/%Y').date()
        except Exception:
            return date.today()

    val_f_entrega = _parse_fecha(oficio_sel['fecha_entrega']) if oficio_sel is not None else date.today()
    val_f_recibido = _parse_fecha(oficio_sel['fecha_recibido']) if oficio_sel is not None else date.today()

    RANGO_FECHA_MIN = date(2000, 1, 1)
    RANGO_FECHA_MAX = date(2030, 12, 31)

    with st.form(f"form_oficio_admin_{contexto_key}", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            id_num = st.text_input("ID Numérico", value=val_id_reg, key=f"id_num_{contexto_key}")
            no_oficio = st.text_input("NO. OFICIO", value=val_no_oficio, key=f"no_oficio_{contexto_key}")
            f_entrega = st.date_input("FECHA DE ENTREGA (DD/MM/AAAA)", value=val_f_entrega, min_value=RANGO_FECHA_MIN, max_value=RANGO_FECHA_MAX, format="DD/MM/YYYY", key=f"f_entrega_{contexto_key}")
            scg_sel = st.selectbox("Bandeja SCG", scg_options, index=(scg_options.index(oficio_sel['scg']) if oficio_sel is not None and oficio_sel['scg'] in scg_options else 0), key=f"scg_{contexto_key}")

        with col2:
            f_recibido = st.date_input("FECHA DE RECIBIDO (DD/MM/AAAA)", value=val_f_recibido, min_value=RANGO_FECHA_MIN, max_value=RANGO_FECHA_MAX, format="DD/MM/YYYY", key=f"f_recibido_{contexto_key}")
            siscat_sel = st.selectbox("Estatus SISCAT", siscat_options, index=(siscat_options.index(oficio_sel['siscat']) if oficio_sel is not None and oficio_sel['siscat'] in siscat_options else 0), key=f"siscat_{contexto_key}")
            sistemas_or_sel = st.selectbox("SISTEMAS/OR", sistemas_or_options, index=(sistemas_or_options.index(oficio_sel['sistemas_or']) if oficio_sel is not None and oficio_sel['sistemas_or'] in sistemas_or_options else 0), key=f"sistemas_or_{contexto_key}")
            tipo_tramite = st.selectbox("TIPO DE TRÁMITE", tramite_options, index=(tramite_options.index(oficio_sel['tipo_tramite']) if oficio_sel is not None and oficio_sel['tipo_tramite'] in tramite_options else 0), key=f"tipo_tramite_{contexto_key}")

        observaciones = st.text_area("OBSERVACIONES", value=val_obs, key=f"obs_{contexto_key}")
        archivo_nuevo = st.file_uploader("Subir/Reemplazar PDF Escaneado", type=["pdf"], key=f"file_of_{contexto_key}")

        btn_label = "💾 Guardar Registro" if modo_accion == "➕ Nuevo Registro" else "✏️ Guardar Cambios"
        if st.form_submit_button(btn_label, type="primary"):
            if not es_oficinas_centrales_admin and (estado_sel == "-- Seleccione --" or municipio_sel == "-- Seleccione --" or ejido_sel == "-- Seleccione --"):
                st.error("⚠️ Debe seleccionar Estado, Municipio y Ejido.")
            elif not dgcat_check or dgcat_check == "DGCAT/100/":
                st.error("⚠️ Debe ingresar un folio DGCAT completo.")
            elif "-- Seleccione --" in (scg_sel, siscat_sel, sistemas_or_sel, tipo_tramite):
                faltantes = []
                if scg_sel == "-- Seleccione --": faltantes.append("Bandeja SCG")
                if siscat_sel == "-- Seleccione --": faltantes.append("Estatus SISCAT")
                if sistemas_or_sel == "-- Seleccione --": faltantes.append("SISTEMAS/OR")
                if tipo_tramite == "-- Seleccione --": faltantes.append("Tipo de Trámite")
                st.error(f"⚠️ Este oficio no tiene valor asignado en: **{', '.join(faltantes)}**. Selecciónelo(s) arriba antes de guardar.")
            elif archivo_nuevo is not None and not es_pdf_valido(archivo_nuevo.getvalue())[0]:
                _, motivo_pdf = es_pdf_valido(archivo_nuevo.getvalue())
                st.error(f"❌ El archivo no se pudo guardar: {motivo_pdf} Vuelva a intentar con otro archivo.")
            else:
                id_num_upper = id_num.strip().upper()
                no_oficio_upper = no_oficio.strip().upper()
                obs_upper = observaciones.strip().upper()
                estado_upper = estado_sel.strip().upper()
                municipio_upper = municipio_sel.strip().upper()
                ejido_upper = ejido_sel.strip().upper()
                scg_upper = scg_sel.strip().upper()
                siscat_upper = siscat_sel.strip().upper()
                sistemas_or_upper = sistemas_or_sel.strip().upper()
                tramite_upper = tipo_tramite.strip().upper()

                str_f_entrega = f_entrega.strftime('%d/%m/%Y')
                str_f_recibido = f_recibido.strftime('%d/%m/%Y')

                archivo_final = ""
                if archivo_nuevo is not None:
                    archivo_final = f"{dgcat_check.replace('/', '_')}_{archivo_nuevo.name}"
                    with open(os.path.join(UPLOADS_DIR, archivo_final), "wb") as f:
                        f.write(archivo_nuevo.getbuffer())

                datos = {
                    "id_registro": id_num_upper, "estado": estado_upper, "municipio": municipio_upper,
                    "ejido": ejido_upper, "no_oficio": no_oficio_upper, "dgcat": dgcat_check,
                    "fecha_entrega": str_f_entrega, "fecha_recibido": str_f_recibido,
                    "scg": scg_upper, "siscat": siscat_upper, "sistemas_or": sistemas_or_upper,
                    "tipo_tramite": tramite_upper, "observaciones": obs_upper,
                }
                if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None:
                    if archivo_final:
                        datos["archivo_escaneado"] = archivo_final
                    ok, msg = guardar_oficio(datos, id_oficio=int(oficio_sel['id']), usuario=st.session_state["username"])
                else:
                    datos["archivo_escaneado"] = archivo_final
                    ok, msg = guardar_oficio(datos, usuario=st.session_state["username"])

                if ok:
                    st.session_state["reset_count_oficios"] += 1
                    flash("success", "✅ Se registró de forma exitosa." if modo_accion == "➕ Nuevo Registro" else f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

# -----------------------------------------------------------------------------
# 4. CONSULTA Y EXPEDIENTES
# -----------------------------------------------------------------------------
elif menu == "🔍 Consulta y Expedientes":
    st.title("🔍 Consulta de Expedientes DGCAT y Descarga de PDF")
    st.caption(f"📅 Información actualizada: {fecha_larga_es()}")

    texto_busqueda_consulta = st.text_input("🔎 Buscar por folio DGCAT, No. de oficio o Estado (opcional):", key="busq_consulta_oficios")
    if texto_busqueda_consulta.strip():
        patron = f"%{texto_busqueda_consulta.strip().upper()}%"
        filtro_sql = "WHERE UPPER(dgcat) LIKE :p OR UPPER(no_oficio) LIKE :p OR UPPER(estado) LIKE :p"
        params_filtro = {"p": patron}
    else:
        filtro_sql, params_filtro = "", {}

    col_pag1, col_pag2, col_pag3 = st.columns([1, 1, 2])
    with col_pag1:
        tam_pagina = st.selectbox("Registros por página", [100, 200, 500], index=1, key="tam_pag_oficios")
    if "pagina_oficios" not in st.session_state:
        st.session_state["pagina_oficios"] = 1

    cols_order = ['id', 'id_registro', 'estado', 'municipio', 'ejido', 'no_oficio', 'dgcat', 'fecha_entrega', 'fecha_recibido', 'scg', 'siscat', 'sistemas_or', 'tipo_tramite', 'observaciones', 'archivo_escaneado']
    df_pagina, total_registros = obtener_pagina(
        "oficios", columnas=", ".join(cols_order), filtro_sql=filtro_sql, params=params_filtro,
        order_by="id DESC", page=st.session_state["pagina_oficios"], page_size=tam_pagina
    )
    df_pagina.columns = [c.lower() for c in df_pagina.columns]
    total_paginas = max((total_registros - 1) // tam_pagina + 1, 1)
    st.session_state["pagina_oficios"] = min(st.session_state["pagina_oficios"], total_paginas)

    with col_pag2:
        st.write("")
        st.write("")
        st.caption(f"Página {st.session_state['pagina_oficios']} de {total_paginas} — {total_registros} registro(s) en total")
    with col_pag3:
        st.write("")
        bcol1, bcol2 = st.columns(2)
        with bcol1:
            if st.button("⬅️ Anterior", disabled=(st.session_state["pagina_oficios"] <= 1), use_container_width=True, key="btn_prev_oficios"):
                st.session_state["pagina_oficios"] -= 1
                st.rerun()
        with bcol2:
            if st.button("Siguiente ➡️", disabled=(st.session_state["pagina_oficios"] >= total_paginas), use_container_width=True, key="btn_next_oficios"):
                st.session_state["pagina_oficios"] += 1
                st.rerun()

    df_display = normalizar_para_mostrar(df_pagina)
    st.dataframe(df_display.drop(columns=['id']), use_container_width=True)

    st.markdown("---")
    st.subheader("📁 Visor, Descarga y Documentos Adicionales por Expediente")

    if total_registros == 0:
        st.info("ℹ️ No hay expedientes registrados.")
    else:
        df_selector, _ = obtener_pagina(
            "oficios", columnas="id, dgcat", filtro_sql=filtro_sql, params=params_filtro,
            order_by="id DESC", page=1, page_size=2000
        )
        df_selector['folio_mostrado'] = df_selector['dgcat'].astype(str).str.upper()
        conteo_repetidos = {}
        etiquetas = []
        for folio in df_selector['folio_mostrado']:
            conteo_repetidos[folio] = conteo_repetidos.get(folio, 0) + 1
            etiquetas.append(f"Folio: {folio}" + (f" ({conteo_repetidos[folio]})" if conteo_repetidos[folio] > 1 else ""))
        df_selector['display_name'] = etiquetas
        expediente_elegido = st.selectbox("Seleccione el expediente a consultar:", df_selector['display_name'].tolist(), key="sel_expediente_consulta")
        id_expediente = int(df_selector.loc[df_selector['display_name'] == expediente_elegido, 'id'].iloc[0])

        fila_full, _ = obtener_pagina("oficios", columnas="*", filtro_sql="WHERE id = :id", params={"id": id_expediente}, order_by="id", page=1, page_size=1)
        fila_full.columns = [c.lower() for c in fila_full.columns]
        fila_expediente = fila_full.iloc[0]

        archivo_principal = fila_expediente['archivo_escaneado']

        col_v1, col_v2 = st.columns([2, 1])
        with col_v1:
            st.markdown("**📄 Expediente Principal**")
            pdf_bytes, ruta_resuelta, error_pdf = cargar_pdf_bytes(archivo_principal)
            if error_pdf == "no_existe":
                st.warning("📌 Favor de subir el oficio.")
            elif error_pdf is not None:
                st.error(f"⚠️ No se pudo cargar el archivo ({error_pdf}). Intente volver a subirlo.")
            else:
                mostrar_pdf(ruta_resuelta)
                st.download_button("📥 Descargar PDF Principal", pdf_bytes, file_name=os.path.basename(ruta_resuelta), mime="application/pdf", type="primary", key="dl_principal")

        with col_v2:
            st.markdown("**➕ Agregar documento adicional**")
            nuevo_adjunto = st.file_uploader("Subir otro PDF sobre este mismo expediente", type=["pdf"], key=f"adj_upload_{id_expediente}")
            if nuevo_adjunto is not None and st.button("📤 Guardar documento adicional", key=f"adj_btn_{id_expediente}"):
                es_valido, motivo = es_pdf_valido(nuevo_adjunto.getvalue())
                if not es_valido:
                    st.error(f"❌ No se guardó: {motivo} Vuelva a intentar con otro archivo.")
                else:
                    nombre_adj = f"ADJ_{id_expediente}_{nuevo_adjunto.name}"
                    with open(os.path.join(UPLOADS_DIR, nombre_adj), "wb") as f:
                        f.write(nuevo_adjunto.getbuffer())
                    ok, msg = agregar_archivo_adjunto("oficios", id_expediente, nombre_adj, st.session_state["username"], nombre_original=nuevo_adjunto.name)
                    if ok:
                        flash("success", f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")

        render_documentos_adicionales("oficios", id_expediente, key_prefix="of")

    st.markdown("---")
    df_export, _ = obtener_pagina("oficios", columnas="*", filtro_sql=filtro_sql, params=params_filtro, order_by="id DESC", page=1, page_size=100000)
    df_export.columns = [c.lower() for c in df_export.columns]
    excel_file = generar_excel_ejecutivo(df_export, mostrar_id=False)
    with open(excel_file, "rb") as f:
        st.download_button("📊 Descargar Reporte Ejecutivo en Excel (.xlsx)", f, file_name="Reporte_DGCAT_Ejecutivo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------------------------------------------------------
# 4B. CONSULTA DE SEGUIMIENTO DE UBICACIÓN DE PREDIO
# -----------------------------------------------------------------------------
elif menu == "🗂️ Consulta de Seguimiento de Predio":
    st.title("🗂️ Consulta de Seguimiento de Ubicación de Predio")
    st.caption(f"📅 Información actualizada: {fecha_larga_es()}")

    texto_busqueda_seg_c = st.text_input("🔎 Buscar por folio DGCAT, Estado o Municipio (opcional):", key="busq_consulta_seg")
    if texto_busqueda_seg_c.strip():
        patron_seg = f"%{texto_busqueda_seg_c.strip().upper()}%"
        filtro_sql_seg = "WHERE UPPER(dgcat) LIKE :p OR UPPER(estado) LIKE :p OR UPPER(municipio) LIKE :p"
        params_filtro_seg = {"p": patron_seg}
    else:
        filtro_sql_seg, params_filtro_seg = "", {}

    col_pag1s, col_pag2s, col_pag3s = st.columns([1, 1, 2])
    with col_pag1s:
        tam_pagina_seg = st.selectbox("Registros por página", [100, 200, 500], index=1, key="tam_pag_seg")
    if "pagina_seg" not in st.session_state:
        st.session_state["pagina_seg"] = 1

    cols_order_seg = ['id', 'dgcat', 'estado', 'municipio', 'ejido', 'fecha_registro', 'fecha_actualizacion', 'observaciones', 'archivo_escaneado', 'registrado_por']
    df_pagina_seg, total_seg = obtener_pagina(
        "seguimiento_predio", columnas=", ".join(cols_order_seg), filtro_sql=filtro_sql_seg, params=params_filtro_seg,
        order_by="id DESC", page=st.session_state["pagina_seg"], page_size=tam_pagina_seg
    )
    df_pagina_seg.columns = [c.lower() for c in df_pagina_seg.columns]
    total_paginas_seg = max((total_seg - 1) // tam_pagina_seg + 1, 1)
    st.session_state["pagina_seg"] = min(st.session_state["pagina_seg"], total_paginas_seg)

    with col_pag2s:
        st.write("")
        st.write("")
        st.caption(f"Página {st.session_state['pagina_seg']} de {total_paginas_seg} — {total_seg} registro(s) en total")
    with col_pag3s:
        st.write("")
        bcol1s, bcol2s = st.columns(2)
        with bcol1s:
            if st.button("⬅️ Anterior", disabled=(st.session_state["pagina_seg"] <= 1), use_container_width=True, key="btn_prev_seg"):
                st.session_state["pagina_seg"] -= 1
                st.rerun()
        with bcol2s:
            if st.button("Siguiente ➡️", disabled=(st.session_state["pagina_seg"] >= total_paginas_seg), use_container_width=True, key="btn_next_seg"):
                st.session_state["pagina_seg"] += 1
                st.rerun()

    df_seg_display = normalizar_para_mostrar(df_pagina_seg)
    st.dataframe(df_seg_display.drop(columns=['id']), use_container_width=True)

    st.markdown("---")
    st.subheader("📁 Visor, Descarga y Documentos Adicionales por Predio")

    if total_seg == 0:
        st.info("ℹ️ No hay registros de seguimiento.")
    else:
        df_selector_seg, _ = obtener_pagina(
            "seguimiento_predio", columnas="id, dgcat", filtro_sql=filtro_sql_seg, params=params_filtro_seg,
            order_by="id DESC", page=1, page_size=2000
        )
        df_selector_seg['folio_mostrado'] = df_selector_seg['dgcat'].astype(str).str.upper()
        conteo_repetidos_seg = {}
        etiquetas_seg = []
        for folio in df_selector_seg['folio_mostrado']:
            conteo_repetidos_seg[folio] = conteo_repetidos_seg.get(folio, 0) + 1
            etiquetas_seg.append(f"Folio: {folio}" + (f" ({conteo_repetidos_seg[folio]})" if conteo_repetidos_seg[folio] > 1 else ""))
        df_selector_seg['display_name'] = etiquetas_seg
        predio_elegido = st.selectbox("Seleccione el predio a consultar:", df_selector_seg['display_name'].tolist(), key="sel_predio_consulta")
        id_predio = int(df_selector_seg.loc[df_selector_seg['display_name'] == predio_elegido, 'id'].iloc[0])

        fila_predio_full, _ = obtener_pagina("seguimiento_predio", columnas="*", filtro_sql="WHERE id = :id", params={"id": id_predio}, order_by="id", page=1, page_size=1)
        fila_predio_full.columns = [c.lower() for c in fila_predio_full.columns]
        fila_predio = fila_predio_full.iloc[0]

        archivo_principal_seg = fila_predio['archivo_escaneado']

        col_v1, col_v2 = st.columns([2, 1])
        with col_v1:
            st.markdown("**📄 Expediente Principal**")
            pdf_bytes_seg, ruta_resuelta_seg, error_pdf_seg = cargar_pdf_bytes(archivo_principal_seg)
            if error_pdf_seg == "no_existe":
                st.warning("📌 Favor de subir el oficio.")
            elif error_pdf_seg is not None:
                st.error(f"⚠️ No se pudo cargar el archivo ({error_pdf_seg}). Intente volver a subirlo.")
            else:
                mostrar_pdf(ruta_resuelta_seg)
                st.download_button("📥 Descargar PDF Principal", pdf_bytes_seg, file_name=os.path.basename(ruta_resuelta_seg), mime="application/pdf", type="primary", key="dl_principal_seg")

        with col_v2:
            st.markdown("**➕ Agregar documento adicional**")
            nuevo_adjunto_seg = st.file_uploader("Subir otro PDF sobre este mismo predio", type=["pdf"], key=f"adj_upload_seg_{id_predio}")
            if nuevo_adjunto_seg is not None and st.button("📤 Guardar documento adicional", key=f"adj_btn_seg_{id_predio}"):
                es_valido_seg, motivo_seg = es_pdf_valido(nuevo_adjunto_seg.getvalue())
                if not es_valido_seg:
                    st.error(f"❌ No se guardó: {motivo_seg} Vuelva a intentar con otro archivo.")
                else:
                    nombre_adj_seg = f"ADJ_{id_predio}_{nuevo_adjunto_seg.name}"
                    with open(os.path.join(UPLOADS_DIR, nombre_adj_seg), "wb") as f:
                        f.write(nuevo_adjunto_seg.getbuffer())
                    ok, msg = agregar_archivo_adjunto("seguimiento_predio", id_predio, nombre_adj_seg, st.session_state["username"], nombre_original=nuevo_adjunto_seg.name)
                    if ok:
                        flash("success", f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")

        render_documentos_adicionales("seguimiento_predio", id_predio, key_prefix="seg")

    st.markdown("---")
    df_export_seg, _ = obtener_pagina("seguimiento_predio", columnas="*", filtro_sql=filtro_sql_seg, params=params_filtro_seg, order_by="id DESC", page=1, page_size=100000)
    df_export_seg.columns = [c.lower() for c in df_export_seg.columns]
    excel_file_seg = generar_excel_seguimiento(df_export_seg, mostrar_id=False)
    with open(excel_file_seg, "rb") as f:
        st.download_button("📊 Descargar Reporte de Seguimiento en Excel (.xlsx)", f, file_name="Reporte_Seguimiento_Predio.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_excel_seg")

# -----------------------------------------------------------------------------
# 5. GESTIÓN DE CATÁLOGOS
# -----------------------------------------------------------------------------
elif menu == "⚙️ Gestión de Catálogos":
    st.title("⚙️ Administración de Catálogos Dinámicos")
    
    t1, t2, t3, t4 = st.tabs(["Catálogo SCG", "Catálogo SISCAT", "SISTEMAS/OR", "Tipos de Trámite"])
    
    def render_catalogo_produccion(tabla_db, label_catalogo):
        try:
            with engine.connect() as conn:
                df_opts = pd.read_sql(f"SELECT nombre FROM {tabla_db} ORDER BY nombre", conn)
                options = df_opts['nombre'].tolist() if not df_opts.empty else []
        except Exception:
            df_opts = pd.DataFrame(columns=['nombre'])
            options = []

        st.subheader(f"Opciones Actuales de {label_catalogo}")
        st.dataframe(df_opts, use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        col_crear, _ = st.columns([2, 1])
        with col_crear:
            with st.form(key=f"form_add_{tabla_db}", clear_on_submit=True):
                nueva_opcion = st.text_input(f"Nueva Opción para {label_catalogo}")
                btn_guardar = st.form_submit_button(f"Guardar Opción {label_catalogo}")
                
                if btn_guardar:
                    if nueva_opcion.strip():
                        val_clean = nueva_opcion.strip().upper()
                        try:
                            with engine.begin() as conn:
                                if engine.url.drivername == 'sqlite':
                                    conn.execute(text(f"INSERT OR IGNORE INTO {tabla_db} (nombre) VALUES (:n)"), {"n": val_clean})
                                else:
                                    conn.execute(text(f"INSERT INTO {tabla_db} (nombre) VALUES (:n) ON CONFLICT DO NOTHING"), {"n": val_clean})
                            st.cache_data.clear()
                            flash("success", f"✅ Opción '{val_clean}' guardada con éxito.")
                            st.rerun()
                        except Exception as err:
                            st.error(f"Error al guardar: {err}")
                    else:
                        st.warning("⚠️ Ingrese un nombre válido.")

        st.markdown("---")

        if options:
            st.subheader(f"✏️ Modificar o 🗑️ Eliminar Opción de {label_catalogo}")
            placeholder_opcion = "-- Seleccione un elemento --"
            opciones_con_placeholder = [placeholder_opcion] + options
            
            opcion_sel = st.selectbox(
                f"Seleccione elemento a gestionar:", 
                opciones_con_placeholder, 
                key=f"sel_{tabla_db}"
            )

            with st.form(key=f"form_edit_{tabla_db}", clear_on_submit=True):
                nuevo_nombre = st.text_input(
                    "Nuevo nombre para modificar:", 
                    placeholder="Escriba el nombre en mayúsculas..."
                )
                
                col_b1, col_b2, _ = st.columns([1, 1, 2])
                with col_b1:
                    btn_modificar = st.form_submit_button("✏️ Guardar Modificación", use_container_width=True)
                with col_b2:
                    btn_eliminar = st.form_submit_button("🗑️ Eliminar Opción", type="primary", use_container_width=True)

                if btn_modificar:
                    if opcion_sel == placeholder_opcion:
                        st.warning("⚠️ Debe seleccionar un elemento.")
                    elif not nuevo_nombre.strip():
                        st.warning("⚠️ Escriba el nuevo nombre.")
                    else:
                        nuevo_clean = nuevo_nombre.strip().upper()
                        ok, msg = actualizar_opcion_catalogo(tabla_db, opcion_sel, nuevo_clean)
                        if ok:
                            st.cache_data.clear()
                            flash("success", f"✅ Modificado correctamente a '{nuevo_clean}'.")
                            st.rerun()
                        else:
                            st.error(msg)

                if btn_eliminar:
                    if opcion_sel == placeholder_opcion:
                        st.warning("⚠️ Debe seleccionar un elemento.")
                    else:
                        n_afectados = contar_oficios_con_valor_catalogo(tabla_db, opcion_sel)
                        ok, msg = eliminar_opcion_catalogo(tabla_db, opcion_sel)
                        if ok:
                            aviso = f" ⚠️ {n_afectados} oficio(s) ya capturados seguían usando este valor y ahora quedará fuera del catálogo." if n_afectados else ""
                            flash("success", f"✅ Eliminado correctamente.{aviso}")
                            st.rerun()
                        else:
                            st.error(msg)

    with t1: render_catalogo_produccion("cat_scg", "SCG")
    with t2: render_catalogo_produccion("cat_siscat", "SISCAT")
    with t3: render_catalogo_produccion("cat_sistemas_or", "SISTEMAS/OR")
    with t4: render_catalogo_produccion("cat_tramite", "Tipos de Trámite")

# -----------------------------------------------------------------------------
# 6. ALTA DE USUARIOS
# -----------------------------------------------------------------------------
elif menu == "👥 Alta de Usuarios":
    if ROL_ACTUAL != "admin":
        st.error("⛔ Acceso no autorizado. Esta sección es exclusiva del Administrador.")
        st.stop()

    st.title("👥 Gestión Completa de Usuarios")
    
    col_u1, col_u2 = st.columns([1.1, 1.2])
    with col_u1:
        st.write("### ➕ Registrar Nuevo Usuario")
        with st.form("form_nuevo_usr", clear_on_submit=True):
            reg_nombre = st.text_input("Nombre Completo *")
            reg_user = st.text_input("Usuario Único *")
            reg_pwd = st.text_input("Contraseña *")
            reg_rol = st.selectbox("Perfil *", ["operador", "supervisor", "admin"])
            
            if st.form_submit_button("💾 Crear Cuenta de Usuario", type="primary"):
                if not reg_nombre or not reg_user or not reg_pwd:
                    st.warning("⚠️ Llene todos los campos.")
                else:
                    exito, msg = registrar_nuevo_usuario(reg_user.strip().lower(), reg_pwd, reg_nombre.strip().upper(), reg_rol)
                    if exito:
                        st.cache_data.clear()
                        flash("success", msg)
                        st.rerun()
                    else:
                        st.error(msg)

        st.markdown("---")
        st.write("### ✏️ Modificar / 🗑️ Eliminar Usuario")
        df_users_admin = pd.read_sql("SELECT username, nombre_completo, rol FROM usuarios", engine)

        if df_users_admin.empty:
            st.info("No hay usuarios registrados.")
        else:
            usuario_gestionar = st.selectbox(
                "Seleccione el usuario a gestionar:",
                df_users_admin['username'].tolist(),
                key="sel_usuario_gestionar"
            )
            fila_usr = df_users_admin[df_users_admin['username'] == usuario_gestionar].iloc[0]

            with st.form(f"form_edit_usr_{usuario_gestionar}"):
                edit_nombre = st.text_input("Nombre Completo", value=fila_usr['nombre_completo'])
                edit_rol = st.selectbox(
                    "Perfil",
                    ["operador", "supervisor", "admin"],
                    index=["operador", "supervisor", "admin"].index(fila_usr['rol']) if fila_usr['rol'] in ["operador", "supervisor", "admin"] else 0
                )
                edit_pwd = st.text_input("Nueva Contraseña (dejar en blanco para no cambiarla)", type="password")

                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    btn_mod_usr = st.form_submit_button("✏️ Guardar Cambios", use_container_width=True)
                with col_b2:
                    btn_del_usr = st.form_submit_button("🗑️ Eliminar Usuario", type="primary", use_container_width=True)

                if btn_mod_usr:
                    ok, msg = modificar_usuario(usuario_gestionar, edit_nombre, edit_rol, nueva_password=edit_pwd.strip() or None)
                    if ok:
                        st.cache_data.clear()
                        flash("success", f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")

                if btn_del_usr:
                    if usuario_gestionar.strip().lower() == "admin":
                        st.error("❌ No se puede eliminar al administrador principal.")
                    else:
                        ok, msg = eliminar_usuario(usuario_gestionar)
                        if ok:
                            st.cache_data.clear()
                            if "sel_usuario_gestionar" in st.session_state:
                                del st.session_state["sel_usuario_gestionar"]
                            flash("success", f"✅ {msg}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg}")

    with col_u2:
        st.write("### 📋 Directorio de Usuarios")
        df_users = pd.read_sql("SELECT username, nombre_completo, rol FROM usuarios", engine)
        st.dataframe(df_users, use_container_width=True)
