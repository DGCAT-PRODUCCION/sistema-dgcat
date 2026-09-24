import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import date
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
        <div class="header-title">🏛️ DIRECCIÓN GENERAL DE CATASTRO</div>
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

if st.session_state["rol"] == "operador":
    menu_options = ["📍 Seguimiento de Ubicación de Predio"]
else:
    menu_options = [
        "📈 Dashboard Ejecutivo", 
        "📝 Registro Completo de Oficios", 
        "📍 Seguimiento de Ubicación de Predio", 
        "🔍 Consulta y Expedientes", 
        "🗂️ Consulta de Seguimiento de Predio",
        "⚙️ Gestión de Catálogos",
        "👥 Alta de Usuarios"
    ]

menu = st.sidebar.radio("Menú de Opciones", menu_options)

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
    st.title("🏛️ Tablero de Control Directivo DGCAT")
    st.caption("Monitoreo institucional de oficios de respuesta, expedientes PDF digitalizados, SCG y SISCAT.")
    
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
        fig_pdf.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=200, showlegend=True)
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
        fig_pie_scg.update_layout(showlegend=True, margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_pie_scg, use_container_width=True)

    with g_col2:
        st.subheader("📊 Control e Integración en SISCAT")
        siscat_counts = df_filtered['siscat'].value_counts().reset_index()
        siscat_counts.columns = ['Estatus', 'Cantidad']
        
        fig_bar_siscat = px.bar(
            siscat_counts, x='Cantidad', y='Estatus', orientation='h', color='Estatus', 
            text='Cantidad', color_discrete_sequence=px.colors.qualitative.Vivid
        )
        fig_bar_siscat.update_layout(showlegend=False, xaxis_title="Número de Oficios", yaxis_title="")
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
        fig_top_estados.update_traces(textposition='inside', textfont_color='white')
        fig_top_estados.update_layout(
            xaxis_title="", yaxis_title="Total Oficios", 
            coloraxis_showscale=False, xaxis_tickangle=-35
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
            coloraxis_showscale=False, yaxis=dict(autorange="reversed")
        )
        st.plotly_chart(fig_tramite_bar, use_container_width=True)

# -----------------------------------------------------------------------------
# 2. SEGUIMIENTO DE UBICACIÓN DE PREDIO
# -----------------------------------------------------------------------------
elif menu == "📍 Seguimiento de Ubicación de Predio":
    st.title("📍 Seguimiento de Ubicación de Predio")
    st.caption("Todos los campos marcados con (*) son strictly OBLIGATORIOS.")

    es_oficinas_centrales = st.checkbox("🏢 Trámite Perteneciente a OFICINAS CENTRALES")

    estados_list = get_estados()
    col_u1, col_u2, col_u3 = st.columns(3)

    if es_oficinas_centrales:
        estado_sel, municipio_sel, ejido_sel = "OFICINAS CENTRALES", "OFICINAS CENTRALES", "OFICINAS CENTRALES"
        with col_u1: st.text_input("1. Estado *", value="OFICINAS CENTRALES", disabled=True)
        with col_u2: st.text_input("2. Municipio *", value="OFICINAS CENTRALES", disabled=True)
        with col_u3: st.text_input("3. Ejido *", value="OFICINAS CENTRALES", disabled=True)
    else:
        with col_u1: estado_sel = st.selectbox("1. Estado *", ["-- Seleccione --"] + estados_list)
        muns_list = get_municipios(estado_sel) if estado_sel != "-- Seleccione --" else []
        with col_u2: municipio_sel = st.selectbox("2. Municipio *", ["-- Seleccione --"] + muns_list)
        ejidos_list = get_ejidos(estado_sel, municipio_sel) if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --" else []
        with col_u3: ejido_sel = st.selectbox("3. Ejido *", ["-- Seleccione --"] + ejidos_list)

    st.markdown("---")
    dgcat_folio = st.text_input("Folio DGCAT *", value="DGCAT/100/", key="dgcat_seguimiento")

    dgcat_check = dgcat_folio.strip().upper()
    id_existente = existe_folio_seguimiento(dgcat_check) if dgcat_check and dgcat_check != "DGCAT/100/" else None
    if id_existente:
        st.info(f"ℹ️ Este folio ya existe (ID #{id_existente}). Al guardar se ACTUALIZARÁ ese registro en lugar de crear uno nuevo.")

    with st.form("form_seguimiento_predio"):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st.text_input("Folio DGCAT (confirmado)", value=dgcat_folio, disabled=True)
        with col_f2:
            archivo_escaneado = st.file_uploader("Adjuntar Expediente Escaneado (PDF) *", type=["pdf"])

        observaciones_capturista = st.text_area("Observaciones (Opcional)")

        if st.form_submit_button("📤 Guardar Seguimiento de Predio", type="primary"):
            if not es_oficinas_centrales and (estado_sel == "-- Seleccione --" or municipio_sel == "-- Seleccione --" or ejido_sel == "-- Seleccione --"):
                st.error("⚠️ Debe seleccionar Estado, Municipio y Ejido.")
            elif not dgcat_folio or dgcat_check == "DGCAT/100/":
                st.error("⚠️ Debe ingresar un folio DGCAT completo.")
            elif archivo_escaneado is None and not id_existente:
                st.error("⚠️ Debe adjuntar un archivo PDF escaneado.")
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
                    "dgcat": dgcat_check,
                    "estado": estado_upper,
                    "municipio": municipio_upper,
                    "ejido": ejido_upper,
                    "observaciones": obs_upper,
                    "registrado_por": st.session_state["username"],
                }
                if id_existente:
                    datos["fecha_actualizacion"] = fecha_hoy
                    if nombre_archivo:
                        datos["archivo_escaneado"] = nombre_archivo
                    ok, msg = guardar_seguimiento_predio(datos, id_registro=id_existente)
                else:
                    datos["fecha_registro"] = fecha_hoy
                    datos["fecha_actualizacion"] = fecha_hoy
                    datos["archivo_escaneado"] = nombre_archivo or ""
                    ok, msg = guardar_seguimiento_predio(datos)

                if ok:
                    st.success(f"✅ {msg}")
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

    oficio_sel = None
    if modo_accion in ["✏️ Modificar Registro Existente", "🗑️ Eliminar Registro"]:
        if df_oficios.empty:
            st.warning("No hay registros en la base de datos.")
            st.stop()

        df_oficios['display_name'] = "ID #" + df_oficios['id'].astype(str) + " | Folio: " + df_oficios['dgcat'].astype(str) + " | Oficio: " + df_oficios['no_oficio'].fillna('').astype(str)

        texto_busqueda = st.text_input("🔎 Buscar por folio DGCAT, No. de oficio o Estado (opcional):")
        df_busqueda = df_oficios
        if texto_busqueda.strip():
            patron = texto_busqueda.strip().upper()
            df_busqueda = df_oficios[
                df_oficios['dgcat'].astype(str).str.upper().str.contains(patron, na=False) |
                df_oficios['no_oficio'].astype(str).str.upper().str.contains(patron, na=False) |
                df_oficios['estado'].astype(str).str.upper().str.contains(patron, na=False)
            ]
            if df_busqueda.empty:
                st.warning("⚠️ No se encontró ningún oficio que coincida con la búsqueda.")
                st.stop()

        opciones_oficios = df_busqueda['display_name'].tolist()
        seleccion = st.selectbox("🔍 Seleccione el Oficio:", opciones_oficios)
        id_oficio_seleccionado = int(seleccion.split(" | ")[0].replace("ID #", ""))
        oficio_sel = df_oficios[df_oficios['id'] == id_oficio_seleccionado].iloc[0]

    if modo_accion == "🗑️ Eliminar Registro" and oficio_sel is not None:
        st.error(f"⚠️ Eliminar oficio **{oficio_sel['dgcat']}** (ID #{oficio_sel['id']}).")
        confirmar = st.checkbox("Confirmo que deseo eliminar este registro de forma DEFINITIVA.")
        if st.button("🚨 ELIMINAR DEFINITIVAMENTE", type="primary", disabled=not confirmar):
            ok, msg = eliminar_oficio(int(oficio_sel['id']))
            if ok:
                st.success(f"✅ {msg}")
                st.rerun()
            else:
                st.error(f"❌ {msg}")
        st.stop()

    es_oficinas_centrales_admin = st.checkbox("🏢 Trámite Perteneciente a OFICINAS CENTRALES")
    estados_list = get_estados()

    col_u1, col_u2, col_u3 = st.columns(3)
    if es_oficinas_centrales_admin:
        estado_sel, municipio_sel, ejido_sel = "OFICINAS CENTRALES", "OFICINAS CENTRALES", "OFICINAS CENTRALES"
        with col_u1: st.text_input("1. Estado", value="OFICINAS CENTRALES", disabled=True)
        with col_u2: st.text_input("2. Municipio", value="OFICINAS CENTRALES", disabled=True)
        with col_u3: st.text_input("3. Ejido", value="OFICINAS CENTRALES", disabled=True)
    else:
        def_estado_idx = estados_list.index(oficio_sel['estado']) + 1 if (oficio_sel is not None and oficio_sel['estado'] in estados_list) else 0
        with col_u1: estado_sel = st.selectbox("1. Estado", ["-- Seleccione --"] + estados_list, index=def_estado_idx)

        muns_list = get_municipios(estado_sel) if estado_sel != "-- Seleccione --" else []
        def_mun_idx = muns_list.index(oficio_sel['municipio']) + 1 if (oficio_sel is not None and oficio_sel['municipio'] in muns_list) else 0
        with col_u2: municipio_sel = st.selectbox("2. Municipio", ["-- Seleccione --"] + muns_list, index=def_mun_idx)

        ejidos_list = get_ejidos(estado_sel, municipio_sel) if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --" else []
        def_ejido_idx = ejidos_list.index(oficio_sel['ejido']) + 1 if (oficio_sel is not None and oficio_sel['ejido'] in ejidos_list) else 0
        with col_u3: ejido_sel = st.selectbox("3. Ejido", ["-- Seleccione --"] + ejidos_list, index=def_ejido_idx)

    val_dgcat = str(oficio_sel['dgcat']) if (oficio_sel is not None and pd.notna(oficio_sel['dgcat'])) else "DGCAT/100/"
    dgcat_folio = st.text_input("Folio DGCAT", value=val_dgcat, key="dgcat_registro_completo")

    dgcat_check = dgcat_folio.strip().upper()
    id_excluir = int(oficio_sel['id']) if oficio_sel is not None else None
    id_duplicado = existe_folio_oficio(dgcat_check, excluir_id=id_excluir) if dgcat_check and dgcat_check != "DGCAT/100/" else None
    if modo_accion == "➕ Nuevo Registro" and id_duplicado:
        st.warning(f"⚠️ El folio **{dgcat_check}** ya existe (ID #{id_duplicado}). Cambie a modo 'Modificar Registro Existente' para editarlo, o use un folio distinto.")

    # Cargar opciones con la opción por defecto '-- Seleccione --'
    scg_raw = pd.read_sql("SELECT nombre FROM cat_scg ORDER BY nombre", engine)['nombre'].tolist()
    scg_options = ["-- Seleccione --"] + scg_raw

    siscat_raw = pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", engine)['nombre'].tolist()
    siscat_options = ["-- Seleccione --"] + siscat_raw
    
    try:
        sistemas_or_raw = pd.read_sql("SELECT nombre FROM cat_sistemas_or ORDER BY nombre", engine)['nombre'].tolist()
        if not sistemas_or_raw:
            sistemas_or_raw = ["SISTEMAS", "OR"]
    except Exception:
        sistemas_or_raw = ["SISTEMAS", "OR"]
    sistemas_or_options = ["-- Seleccione --"] + sistemas_or_raw

    tramite_raw = pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", engine)['nombre'].tolist()
    tramite_options = ["-- Seleccione --"] + tramite_raw

    limpiar_al_guardar = True if modo_accion == "➕ Nuevo Registro" else False

    val_id_reg = str(oficio_sel['id_registro']) if (oficio_sel is not None and pd.notna(oficio_sel['id_registro'])) else ""
    val_no_oficio = str(oficio_sel['no_oficio']) if (oficio_sel is not None and pd.notna(oficio_sel['no_oficio'])) else ""
    val_obs = str(oficio_sel['observaciones']) if (oficio_sel is not None and pd.notna(oficio_sel['observaciones'])) else ""

    with st.form("form_oficio_admin", clear_on_submit=limpiar_al_guardar):
        idx_scg = scg_options.index(oficio_sel['scg']) if (oficio_sel is not None and oficio_sel['scg'] in scg_options) else 0
        idx_siscat = siscat_options.index(oficio_sel['siscat']) if (oficio_sel is not None and oficio_sel['siscat'] in siscat_options) else 0
        idx_sistemas_or = sistemas_or_options.index(oficio_sel['sistemas_or']) if (oficio_sel is not None and oficio_sel['sistemas_or'] in sistemas_or_options) else 0
        idx_tramite = tramite_options.index(oficio_sel['tipo_tramite']) if (oficio_sel is not None and oficio_sel['tipo_tramite'] in tramite_options) else 0

        col1, col2 = st.columns(2)
        with col1:
            id_num = st.text_input("ID Numérico", value=val_id_reg)
            no_oficio = st.text_input("NO. OFICIO", value=val_no_oficio)
            f_entrega = st.date_input("FECHA DE ENTREGA (DD/MM/AAAA)", value=date.today(), format="DD/MM/YYYY")
            scg_sel = st.selectbox("Bandeja SCG *", scg_options, index=idx_scg)

        with col2:
            f_recibido = st.date_input("FECHA DE RECIBIDO (DD/MM/AAAA)", value=date.today(), format="DD/MM/YYYY")
            siscat_sel = st.selectbox("Estatus SISCAT *", siscat_options, index=idx_siscat)
            sistemas_or_sel = st.selectbox("SISTEMAS/OR *", sistemas_or_options, index=idx_sistemas_or)
            tipo_tramite = st.selectbox("TIPO DE TRÁMITE *", tramite_options, index=idx_tramite)

        observaciones = st.text_area("OBSERVACIONES", value=val_obs)
        archivo_nuevo = st.file_uploader("Subir/Reemplazar PDF Escaneado", type=["pdf"])

        btn_label = "💾 Guardar Registro" if modo_accion == "➕ Nuevo Registro" else "✏️ Guardar Cambios"
        if st.form_submit_button(btn_label, type="primary"):
            if modo_accion == "➕ Nuevo Registro" and id_duplicado:
                st.error(f"❌ No se guardó: el folio '{dgcat_check}' ya existe (ID #{id_duplicado}). Use 'Modificar Registro Existente'.")
            elif not es_oficinas_centrales_admin and (estado_sel == "-- Seleccione --" or municipio_sel == "-- Seleccione --" or ejido_sel == "-- Seleccione --"):
                st.error("⚠️ Debe seleccionar Estado, Municipio y Ejido.")
            elif scg_sel == "-- Seleccione --" or siscat_sel == "-- Seleccione --" or sistemas_or_sel == "-- Seleccione --" or tipo_tramite == "-- Seleccione --":
                st.error("⚠️ Debe seleccionar una opción válida en Bandeja SCG, Estatus SISCAT, SISTEMAS/OR y Tipo de Trámite.")
            elif not dgcat_check or dgcat_check == "DGCAT/100/":
                st.error("⚠️ Debe ingresar un folio DGCAT completo.")
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
                    ok, msg = guardar_oficio(datos, id_oficio=int(oficio_sel['id']))
                else:
                    datos["archivo_escaneado"] = archivo_final
                    ok, msg = guardar_oficio(datos)

                if ok:
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

# -----------------------------------------------------------------------------
# 4. CONSULTA Y EXPEDIENTES
# -----------------------------------------------------------------------------
elif menu == "🔍 Consulta y Expedientes":
    st.title("🔍 Consulta de Expedientes DGCAT y Descarga de PDF")
    
    df_data = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", engine)
    df_data.columns = [c.lower() for c in df_data.columns]
    
    cols_order = ['id', 'id_registro', 'estado', 'municipio', 'ejido', 'no_oficio', 'dgcat', 'fecha_entrega', 'fecha_recibido', 'scg', 'siscat', 'sistemas_or', 'tipo_tramite', 'observaciones', 'archivo_escaneado']
    cols_presentes = [c for c in cols_order if c in df_data.columns]
    df_display = df_data[cols_presentes]

    for col in df_display.select_dtypes(include=['object']).columns:
        df_display[col] = df_display[col].astype(str).str.upper()

    st.dataframe(df_display, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📁 Visor y Descarga de Archivos PDF (Almacenados en carpeta uploads/)")
    
    df_pdfs = df_display[df_display['archivo_escaneado'].notna() & (df_display['archivo_escaneado'] != '') & (df_display['archivo_escaneado'] != 'NONE')]
    if not df_pdfs.empty:
        col_pdf1, col_pdf2 = st.columns([2, 1])
        with col_pdf1:
            pdf_sel_name = st.selectbox("Seleccione el archivo PDF registrado a consultar:", df_pdfs['archivo_escaneado'].unique())
        with col_pdf2:
            st.write("<br>", unsafe_allow_html=True)
            pdf_path = os.path.join(UPLOADS_DIR, pdf_sel_name)
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    st.download_button("📥 Descargar PDF Escaneado", f, file_name=pdf_sel_name, mime="application/pdf", type="primary")
            else:
                st.warning("⚠️ El archivo no se localizó físicamente en la carpeta uploads/.")
    else:
        st.info("ℹ️ No hay archivos PDF adjuntos registrados.")

    st.markdown("---")
    excel_file = generar_excel_ejecutivo(df_display)
    with open(excel_file, "rb") as f:
        st.download_button("📊 Descargar Reporte Ejecutivo en Excel (.xlsx)", f, file_name="Reporte_DGCAT_Ejecutivo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------------------------------------------------------
# 4B. CONSULTA DE SEGUIMIENTO DE UBICACIÓN DE PREDIO
# -----------------------------------------------------------------------------
elif menu == "🗂️ Consulta de Seguimiento de Predio":
    st.title("🗂️ Consulta de Seguimiento de Ubicación de Predio")

    try:
        df_seg = pd.read_sql("SELECT * FROM seguimiento_predio ORDER BY id DESC", engine)
    except Exception:
        df_seg = pd.DataFrame(columns=[
            'id', 'dgcat', 'estado', 'municipio', 'ejido', 
            'fecha_registro', 'fecha_actualizacion', 'observaciones', 
            'archivo_escaneado', 'registrado_por'
        ])

    df_seg.columns = [c.lower() for c in df_seg.columns]

    cols_order_seg = ['id', 'dgcat', 'estado', 'municipio', 'ejido', 'fecha_registro', 'fecha_actualizacion', 'observaciones', 'archivo_escaneado', 'registrado_por']
    cols_presentes_seg = [c for c in cols_order_seg if c in df_seg.columns]
    df_seg_display = df_seg[cols_presentes_seg]

    for col in df_seg_display.select_dtypes(include=['object']).columns:
        df_seg_display[col] = df_seg_display[col].astype(str).str.upper()

    st.dataframe(df_seg_display, use_container_width=True)

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
                            st.success(f"✅ Opción '{val_clean}' guardada con éxito.")
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
                            st.success(f"✅ Modificado correctamente a '{nuevo_clean}'.")
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
                            st.success(f"✅ Eliminado correctamente.{aviso}")
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
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    with col_u2:
        st.write("### 📋 Directorio de Usuarios")
        df_users = pd.read_sql("SELECT username, nombre_completo, rol FROM usuarios", engine)
        st.dataframe(df_users, use_container_width=True)
