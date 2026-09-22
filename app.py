import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import date, datetime
from sqlalchemy import text
from database import (
    init_db, 
    verificar_login, 
    registrar_nuevo_usuario, 
    get_engine, 
    generar_excel_ejecutivo,
    eliminar_oficio,
    eliminar_usuario,
    cambiar_password_usuario
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

# PANTALLA LOGIN
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

# PANEL PRIVADO
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
    menu_options = ["📄 Carga de Archivo Escaneado (PDF)"]
elif st.session_state["rol"] == "supervisor":
    menu_options = [
        "📈 Dashboard Ejecutivo", 
        "📝 Registro Completo de Oficios", 
        "📄 Carga de Archivo Escaneado (PDF)", 
        "🔍 Consulta y Expedientes", 
        "⚙️ Gestión de Catálogos"
    ]
else:
    menu_options = [
        "📈 Dashboard Ejecutivo", 
        "📝 Registro Completo de Oficios", 
        "📄 Carga de Archivo Escaneado (PDF)", 
        "🔍 Consulta y Expedientes", 
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

def parse_date_for_picker(val):
    if pd.isna(val) or val is None or str(val).strip() in ['', 'None', 'nan', 'NaT']:
        return date.today()
    val_str = str(val).strip()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d', '%d-%m-%Y'):
        try:
            return datetime.strptime(val_str, fmt).date()
        except ValueError:
            pass
    try:
        dt = pd.to_datetime(val_str, errors='coerce', dayfirst=True)
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass
    return date.today()

# -----------------------------------------------------------------------------
# 1. DASHBOARD COMPLETO
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
    
    # METRICAS DIRECTIVAS
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

    # CONTROL DIGITAL PDF
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

    # BANDEJA SCG Y SISCAT
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

    # TOP ESTADOS Y VOLUMETRÍA POR TRÁMITE
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
            textposition='outside', 
            textfont=dict(color='#1F2937', size=12)
        )
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
# 2. CARGA DE ARCHIVO ESCANEADO (PDF)
# -----------------------------------------------------------------------------
elif menu == "📄 Carga de Archivo Escaneado (PDF)":
    st.title("📄 Carga Institucional de Expediente PDF (Operador)")
    st.caption("Todos los campos marcados con (*) son estrictamente OBLIGATORIOS.")
    
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
    
    with st.form("form_carga_operador", clear_on_submit=True):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            dgcat_folio = st.text_input("Folio DGCAT *", value="DGCAT/100/")
        with col_f2:
            archivo_escaneado = st.file_uploader("Adjuntar Expediente Escaneado (PDF) *", type=["pdf"])

        observaciones_capturista = st.text_area("Observaciones (Opcional)")

        if st.form_submit_button("📤 Subir y Vincular Documento PDF", type="primary"):
            if not es_oficinas_centrales and (estado_sel == "-- Seleccione --" or municipio_sel == "-- Seleccione --" or ejido_sel == "-- Seleccione --"):
                st.error("⚠️ Debe seleccionar Estado, Municipio y Ejido.")
            elif not dgcat_folio or dgcat_folio.strip() == "DGCAT/100/":
                st.error("⚠️ Debe ingresar un folio DGCAT completo.")
            elif archivo_escaneado is None:
                st.error("⚠️ Debe adjuntar un archivo PDF escaneado.")
            else:
                dgcat_upper = dgcat_folio.strip().upper()
                obs_upper = observaciones_capturista.strip().upper() if observaciones_capturista else ""
                estado_upper = estado_sel.strip().upper()
                municipio_upper = municipio_sel.strip().upper()
                ejido_upper = ejido_sel.strip().upper()

                dgcat_clean = dgcat_upper.replace("/", "_").replace(" ", "")
                nombre_archivo = f"{dgcat_clean}_{archivo_escaneado.name}"
                
                with open(os.path.join(UPLOADS_DIR, nombre_archivo), "wb") as f:
                    f.write(archivo_escaneado.getbuffer())

                fecha_actual_formatted = date.today().strftime('%d/%m/%Y')

                with engine.begin() as conn:
                    res = conn.execute(text("SELECT id FROM oficios WHERE UPPER(TRIM(dgcat)) = :dg"), {"dg": dgcat_upper}).fetchone()
                    if res:
                        conn.execute(text("""
                            UPDATE oficios 
                            SET estado = :e, municipio = :m, ejido = :ej, archivo_escaneado = :arch, 
                                observaciones = COALESCE(NULLIF(:obs, ''), observaciones)
                            WHERE id = :id
                        """), {
                            "e": estado_upper, "m": municipio_upper, "ej": ejido_upper, 
                            "arch": nombre_archivo, "obs": obs_upper, "id": int(res[0])
                        })
                    else:
                        conn.execute(text("""
                            INSERT INTO oficios (estado, municipio, ejido, dgcat, fecha_entrega, scg, siscat, observaciones, archivo_escaneado)
                            VALUES (:e, :m, :ej, :dg, :fe, 'SUBIDO', 'SUBIDO', :obs, :arch)
                        """), {
                            "e": estado_upper, "m": municipio_upper, "ej": ejido_upper, 
                            "dg": dgcat_upper, "fe": fecha_actual_formatted, 
                            "obs": obs_upper, "arch": nombre_archivo
                        })
                st.cache_data.clear()
                st.success(f"✅ Documento PDF subido exitosamente en uploads/ como: {nombre_archivo}")
                st.rerun()

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
        opciones_oficios = df_oficios['display_name'].tolist()
        seleccion = st.selectbox("🔍 Seleccione el Oficio:", opciones_oficios)
        id_oficio_seleccionado = int(seleccion.split(" | ")[0].replace("ID #", ""))
        oficio_sel = df_oficios[df_oficios['id'] == id_oficio_seleccionado].iloc[0]

    # MODO ELIMINAR
    if modo_accion == "🗑️ Eliminar Registro" and oficio_sel is not None:
        st.error(f"⚠️ ¿Desea eliminar definitivamente el oficio **{oficio_sel['dgcat']}** (ID #{oficio_sel['id']})?")
        if st.button("🚨 ELIMINAR DEFINITIVAMENTE", type="primary"):
            eliminar_oficio(int(oficio_sel['id']))
            st.cache_data.clear()
            st.success("✅ Registro eliminado correctamente.")
            st.rerun()
        st.stop()

    # OPCIONES PARA LOS DESPLEGABLES
    scg_options = pd.read_sql("SELECT nombre FROM cat_scg ORDER BY nombre", engine)['nombre'].tolist()
    siscat_options = pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", engine)['nombre'].tolist()
    tramite_options = pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", engine)['nombre'].tolist()
    estados_list = get_estados()

    # DATOS PRECARGADOS
    val_id_reg = str(oficio_sel['id_registro']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['id_registro'])) else ""
    val_dgcat = str(oficio_sel['dgcat']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['dgcat'])) else "DGCAT/100/"
    val_no_oficio = str(oficio_sel['no_oficio']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['no_oficio'])) else ""
    val_obs = str(oficio_sel['observaciones']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['observaciones'])) else ""

    default_f_entrega = parse_date_for_picker(oficio_sel['fecha_entrega']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None) else date.today()
    default_f_recibido = parse_date_for_picker(oficio_sel['fecha_recibido']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None) else date.today()

    limpiar_al_guardar = True if modo_accion == "➕ Nuevo Registro" else False

    # FORMULARIO UNIFICADO
    with st.form("form_oficio_admin", clear_on_submit=limpiar_al_guardar):
        es_oficinas_centrales_admin = st.checkbox("🏢 Trámite Perteneciente a OFICINAS CENTRALES")

        col_u1, col_u2, col_u3 = st.columns(3)
        if es_oficinas_centrales_admin:
            estado_sel, municipio_sel, ejido_sel = "OFICINAS CENTRALES", "OFICINAS CENTRALES", "OFICINAS CENTRALES"
            with col_u1: st.text_input("1. Estado", value="OFICINAS CENTRALES", disabled=True)
            with col_u2: st.text_input("2. Municipio", value="OFICINAS CENTRALES", disabled=True)
            with col_u3: st.text_input("3. Ejido", value="OFICINAS CENTRALES", disabled=True)
        else:
            def_estado_idx = estados_list.index(oficio_sel['estado']) + 1 if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['estado'] in estados_list) else 0
            with col_u1: estado_sel = st.selectbox("1. Estado", ["-- Seleccione --"] + estados_list, index=def_estado_idx)

            muns_list = get_municipios(estado_sel) if estado_sel != "-- Seleccione --" else []
            def_mun_idx = muns_list.index(oficio_sel['municipio']) + 1 if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['municipio'] in muns_list) else 0
            with col_u2: municipio_sel = st.selectbox("2. Municipio", ["-- Seleccione --"] + muns_list, index=def_mun_idx)

            ejidos_list = get_ejidos(estado_sel, municipio_sel) if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --" else []
            def_ejido_idx = ejidos_list.index(oficio_sel['ejido']) + 1 if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['ejido'] in ejidos_list) else 0
            with col_u3: ejido_sel = st.selectbox("3. Ejido", ["-- Seleccione --"] + ejidos_list, index=def_ejido_idx)

        dgcat_folio = st.text_input("Folio DGCAT", value=val_dgcat)

        col1, col2 = st.columns(2)
        with col1:
            id_num = st.text_input("ID Numérico", value=val_id_reg)
            no_oficio = st.text_input("NO. OFICIO", value=val_no_oficio)
            f_entrega = st.date_input("FECHA DE ENTREGA (DD/MM/AAAA)", value=default_f_entrega, format="DD/MM/YYYY")
            
            scg_idx = scg_options.index(oficio_sel['scg']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['scg'] in scg_options) else 0
            scg_sel = st.selectbox("Bandeja SCG", scg_options, index=scg_idx)

        with col2:
            f_recibido = st.date_input("FECHA DE RECIBIDO (DD/MM/AAAA)", value=default_f_recibido, format="DD/MM/YYYY")
            
            siscat_idx = siscat_options.index(oficio_sel['siscat']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['siscat'] in siscat_options) else 0
            siscat_sel = st.selectbox("Estatus SISCAT", siscat_options, index=siscat_idx)
            
            tramite_idx = tramite_options.index(oficio_sel['tipo_tramite']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and 'tipo_tramite' in oficio_sel and oficio_sel['tipo_tramite'] in tramite_options) else 0
            tipo_tramite = st.selectbox("TIPO DE TRÁMITE", tramite_options, index=tramite_idx)

        observaciones = st.text_area("OBSERVACIONES", value=val_obs)
        archivo_nuevo = st.file_uploader("Subir/Reemplazar PDF Escaneado", type=["pdf"])

        btn_label = "💾 Guardar Registro" if modo_accion == "➕ Nuevo Registro" else "✏️ Guardar Cambios"
        if st.form_submit_button(btn_label, type="primary"):
            id_num_upper = id_num.strip().upper()
            no_oficio_upper = no_oficio.strip().upper()
            dgcat_upper = dgcat_folio.strip().upper()
            obs_upper = observaciones.strip().upper()
            estado_upper = estado_sel.strip().upper()
            municipio_upper = municipio_sel.strip().upper()
            ejido_upper = ejido_sel.strip().upper()
            scg_upper = scg_sel.strip().upper()
            siscat_upper = siscat_sel.strip().upper()
            tramite_upper = tipo_tramite.strip().upper()

            str_f_entrega = f_entrega.strftime('%d/%m/%Y')
            str_f_recibido = f_recibido.strftime('%d/%m/%Y')

            archivo_final = ""
            if archivo_nuevo is not None:
                archivo_final = f"{dgcat_upper.replace('/', '_')}_{archivo_nuevo.name}"
                with open(os.path.join(UPLOADS_DIR, archivo_final), "wb") as f:
                    f.write(archivo_nuevo.getbuffer())

            with engine.begin() as conn:
                if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None:
                    conn.execute(text("""
                        UPDATE oficios 
                        SET id_registro = :id_r, estado = :e, municipio = :m, ejido = :ej, no_oficio = :no_of, 
                            dgcat = :dg, fecha_entrega = :f_ent, fecha_recibido = :f_rec, scg = :scg_val, 
                            siscat = :sis_val, tipo_tramite = :tram, observaciones = :obs, 
                            archivo_escaneado = COALESCE(NULLIF(:arch, ''), archivo_escaneado)
                        WHERE id = :id
                    """), {
                        "id_r": id_num_upper, "e": estado_upper, "m": municipio_upper, "ej": ejido_upper, "no_of": no_oficio_upper,
                        "dg": dgcat_upper, "f_ent": str_f_entrega, "f_rec": str_f_recibido,
                        "scg_val": scg_upper, "sis_val": siscat_upper, "tram": tramite_upper, "obs": obs_upper,
                        "arch": archivo_final, "id": int(oficio_sel['id'])
                    })
                    st.cache_data.clear()
                    st.success("✅ Registro modificado y actualizado correctamente.")
                    st.rerun()
                else:
                    conn.execute(text("""
                        INSERT INTO oficios (id_registro, estado, municipio, ejido, no_oficio, dgcat, fecha_entrega, fecha_recibido, scg, siscat, tipo_tramite, observaciones, archivo_escaneado)
                        VALUES (:id_r, :e, :m, :ej, :no_of, :dg, :f_ent, :f_rec, :scg_val, :sis_val, :tram, :obs, :arch)
                    """), {
                        "id_r": id_num_upper, "e": estado_upper, "m": municipio_upper, "ej": ejido_upper, "no_of": no_oficio_upper,
                        "dg": dgcat_upper, "f_ent": str_f_entrega, "f_rec": str_f_recibido,
                        "scg_val": scg_upper, "sis_val": siscat_upper, "tram": tramite_upper, "obs": obs_upper,
                        "arch": archivo_final
                    })
                    st.cache_data.clear()
                    st.success("✅ Guardado correctamente.")
                    st.rerun()

# -----------------------------------------------------------------------------
# 4. CONSULTA Y EXPEDIENTES
# -----------------------------------------------------------------------------
elif menu == "🔍 Consulta y Expedientes":
    st.title("🔍 Consulta de Expedientes DGCAT y Descarga de PDF")
    
    df_data = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", engine)
    df_data.columns = [c.lower() for c in df_data.columns]
    
    cols_order = ['id', 'id_registro', 'estado', 'municipio', 'ejido', 'no_oficio', 'dgcat', 'fecha_entrega', 'fecha_recibido', 'scg', 'siscat', 'tipo_tramite', 'observaciones', 'archivo_escaneado']
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
# 5. GESTIÓN DE CATÁLOGOS
# -----------------------------------------------------------------------------
elif menu == "⚙️ Gestión de Catálogos":
    st.title("⚙️ Administración de Catálogos Dinámicos")
    
    tab1, tab2, tab3 = st.tabs(["Catálogo SCG", "Catálogo SISCAT", "Tipos de Trámite"])
    
    with tab1:
        st.subheader("Opciones Actuales de SCG")
        df_scg = pd.read_sql("SELECT nombre FROM cat_scg ORDER BY nombre", engine)
        st.dataframe(df_scg, use_container_width=True)
        
        n_scg = st.text_input("Nueva Opción para SCG")
        if st.button("Guardar Opción SCG") and n_scg:
            with engine.begin() as conn:
                conn.execute(text("INSERT INTO cat_scg (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": n_scg.strip().upper()})
            st.cache_data.clear()
            st.success("Guardado.")
            st.rerun()
            
    with tab2:
        st.subheader("Opciones Actuales de SISCAT")
        df_siscat = pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", engine)
        st.dataframe(df_siscat, use_container_width=True)
        
        n_siscat = st.text_input("Nueva Opción para SISCAT")
        if st.button("Guardar Opción SISCAT") and n_siscat:
            with engine.begin() as conn:
                conn.execute(text("INSERT INTO cat_siscat (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": n_siscat.strip().upper()})
            st.cache_data.clear()
            st.success("Guardado.")
            st.rerun()
            
    with tab3:
        st.subheader("Tipos de Trámite Registrados")
        df_tramite = pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", engine)
        st.dataframe(df_tramite, use_container_width=True)
        
        n_tramite = st.text_input("Nuevo Tipo de Trámite")
        if st.button("Guardar Trámite") and n_tramite:
            with engine.begin() as conn:
                conn.execute(text("INSERT INTO cat_tramite (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": n_tramite.strip().upper()})
            st.cache_data.clear()
            st.success("Guardado.")
            st.rerun()

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
                    exito, msg = registrar_nuevo_usuario(reg_user, reg_pwd, reg_nombre, reg_rol)
                    if exito:
                        st.cache_data.clear()
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    with col_u2:
        st.write("### 📋 Directorio de Usuarios")
        df_users = pd.read_sql("SELECT username, nombre_completo, rol, password_plain FROM usuarios", engine)
        df_users.columns = [c.lower() for c in df_users.columns]
        
        df_users_display = df_users.rename(columns={
            'username': 'Usuario',
            'nombre_completo': 'Nombre',
            'rol': 'Perfil',
            'password_plain': 'Contraseña'
        })
        st.dataframe(df_users_display, use_container_width=True)

        st.markdown("---")
        lista_todos_usuarios = df_users['username'].tolist() if not df_users.empty else []
        usr_cambiar_pwd = st.selectbox("Usuario a Modificar:", lista_todos_usuarios)
        nueva_pwd_input = st.text_input("Nueva Contraseña:", key="pwd_reset_input")
        
        if st.button("🔄 Actualizar Contraseña"):
            if nueva_pwd_input and usr_cambiar_pwd:
                exito_pwd, msg_pwd = cambiar_password_usuario(usr_cambiar_pwd, nueva_pwd_input)
                st.cache_data.clear()
                st.success(msg_pwd)
                st.rerun()

        st.markdown("---")
        usuarios_para_borrar = [u for u in lista_todos_usuarios if u.lower() != 'admin' and u.lower() != st.session_state['username'].lower()]
        if usuarios_para_borrar:
            user_a_borrar = st.selectbox("Usuario a eliminar:", usuarios_para_borrar)
            if st.button("❌ Confirmar Eliminación"):
                eliminar_usuario(user_a_borrar)
                st.cache_data.clear()
                st.success("Usuario eliminado.")
                st.rerun()
