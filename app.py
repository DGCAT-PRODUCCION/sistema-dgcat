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

# Helper para cargar ubicaciones compatible PostgreSQL / SQLite
@st.cache_data(ttl=600)
def get_estados():
    df = pd.read_sql("SELECT DISTINCT estado FROM cat_ubicaciones ORDER BY estado", engine)
    df.columns = [c.lower() for c in df.columns]
    return sorted(list(df['estado'].dropna().unique()))

@st.cache_data(ttl=600)
def get_municipios(estado):
    df = pd.read_sql("SELECT DISTINCT municipio FROM cat_ubicaciones WHERE estado = :e ORDER BY municipio", engine, params={"e": estado})
    df.columns = [c.lower() for c in df.columns]
    return sorted(list(df['municipio'].dropna().unique()))

@st.cache_data(ttl=600)
def get_ejidos(estado, municipio):
    df = pd.read_sql("SELECT DISTINCT ejido FROM cat_ubicaciones WHERE estado = :e AND municipio = :m ORDER BY ejido", engine, params={"e": estado, "m": municipio})
    df.columns = [c.lower() for c in df.columns]
    return sorted(list(df['ejido'].dropna().unique()))

# -----------------------------------------------------------------------------
# 1. DASHBOARD COMPLETO CON 6 GRAFICAS
# -----------------------------------------------------------------------------
if menu == "📈 Dashboard Ejecutivo":
    st.title("🏛️ Tablero de Control Directivo DGCAT")
    st.caption("Monitoreo institucional de oficios de respuesta, bandeja de geógrafos, sistemas y estatus SISCAT.")
    
    df_raw = pd.read_sql("SELECT * FROM oficios", engine)
    df_raw.columns = [c.lower() for c in df_raw.columns]
    
    if df_raw.empty:
        st.warning("No hay registros disponibles para generar métricas.")
        st.stop()

    df_raw['fecha_entrega_dt'] = pd.to_datetime(df_raw['fecha_entrega'], errors='coerce')
    df_raw['fecha_recibido_dt'] = pd.to_datetime(df_raw['fecha_recibido'], errors='coerce')
    df_raw['dias_respuesta'] = (df_raw['fecha_recibido_dt'] - df_raw['fecha_entrega_dt']).dt.days

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
    if sel_estado != "TODOS":
        df_filtered = df_filtered[df_filtered['estado'] == sel_estado]
    if sel_scg != "TODOS":
        df_filtered = df_filtered[df_filtered['scg'] == sel_scg]
    if sel_siscat != "TODOS":
        df_filtered = df_filtered[df_filtered['siscat'] == sel_siscat]

    st.markdown("---")
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    total_oficios = len(df_filtered)
    concluidos_scg = len(df_filtered[df_filtered['scg'] == 'CONCLUIDO'])
    subidos_siscat = len(df_filtered[df_filtered['siscat'].isin(['CONCLUIDO', 'SUBIDO'])])
    en_sistemas = len(df_filtered[df_filtered['scg'].astype(str).str.contains('SISTEMAS', case=False, na=False)])
    promedio_dias = round(df_filtered['dias_respuesta'].dropna().mean(), 1) if not df_filtered['dias_respuesta'].dropna().empty else 0
    
    pct_scg = round((concluidos_scg / total_oficios * 100), 1) if total_oficios > 0 else 0
    pct_siscat = round((subidos_siscat / total_oficios * 100), 1) if total_oficios > 0 else 0

    with kpi1:
        st.markdown(f'<div class="metric-card"><h3>TOTAL OFICIOS</h3><div class="number">{total_oficios:,}</div><div class="subtitle">Registros Atendidos</div></div>', unsafe_allow_html=True)
    with kpi2:
        st.markdown(f'<div class="metric-card"><h3>EFICIENCIA SCG</h3><div class="number">{pct_scg}%</div><div class="subtitle">{concluidos_scg:,} Concluidos</div></div>', unsafe_allow_html=True)
    with kpi3:
        st.markdown(f'<div class="metric-card"><h3>AVANCE SISCAT</h3><div class="number">{pct_siscat}%</div><div class="subtitle">{subidos_siscat:,} Procesados</div></div>', unsafe_allow_html=True)
    with kpi4:
        st.markdown(f'<div class="metric-card"><h3>EN SISTEMAS</h3><div class="number" style="color:#EF4444;">{en_sistemas:,}</div><div class="subtitle">Pendientes</div></div>', unsafe_allow_html=True)
    with kpi5:
        st.markdown(f'<div class="metric-card"><h3>PROMEDIO SLA</h3><div class="number" style="color:#8B5CF6;">{promedio_dias}d</div><div class="subtitle">Días Respuesta</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # BLOQUE 1 DE GRÁFICAS
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.subheader("🍩 Distribución por Bandeja SCG")
        df_scg_clean = df_filtered.copy()
        top_scg = df_scg_clean['scg'].value_counts()
        df_scg_clean['scg_grouped'] = df_scg_clean['scg'].apply(lambda x: x if x in top_scg.index[:5] else 'OTRAS BANDEJAS')
        scg_counts = df_scg_clean['scg_grouped'].value_counts().reset_index()
        scg_counts.columns = ['Bandeja', 'Cantidad']
        fig_donut = px.pie(scg_counts, values='Cantidad', names='Bandeja', hole=0.4, color_discrete_sequence=px.colors.qualitative.Bold)
        st.plotly_chart(fig_donut, use_container_width=True)

    with g_col2:
        st.subheader("📊 Control e Integración en SISCAT")
        siscat_counts = df_filtered['siscat'].value_counts().reset_index()
        siscat_counts.columns = ['Estatus', 'Cantidad']
        fig_bar_siscat = px.bar(siscat_counts, x='Cantidad', y='Estatus', orientation='h', color='Estatus', text='Cantidad', color_discrete_sequence=px.colors.qualitative.Vivid)
        fig_bar_siscat.update_layout(showlegend=False, xaxis_title="Número de Oficios", yaxis_title="")
        st.plotly_chart(fig_bar_siscat, use_container_width=True)

    st.markdown("---")

    # BLOQUE 2 DE GRÁFICAS (VOLUMETRÍA Y TOP ESTADOS)
    g_col3, g_col4 = st.columns(2)

    with g_col3:
        st.subheader("Top 10 Estados con Mayor Carga Registrada")
        top_estados = df_filtered['estado'].value_counts().head(10).reset_index()
        top_estados.columns = ['Estado', 'Oficios']
        fig_top_estados = px.bar(top_estados, x='Estado', y='Oficios', color='Oficios', text='Oficios', color_continuous_scale='Greens')
        fig_top_estados.update_layout(xaxis_title="", yaxis_title="Total Oficios", coloraxis_showscale=False)
        st.plotly_chart(fig_top_estados, use_container_width=True)

    with g_col4:
        st.subheader("📑 Volumetría por Tipo de Trámite / Observación")
        obs_counts = df_filtered['observaciones'].value_counts().head(8).reset_index()
        obs_counts.columns = ['Trámite', 'Cantidad']
        fig_obs = px.pie(obs_counts, values='Cantidad', names='Trámite', color_discrete_sequence=px.colors.qualitative.Prism)
        fig_obs.update_traces(textinfo='label+value')
        st.plotly_chart(fig_obs, use_container_width=True)

    st.markdown("---")

    # BLOQUE 3 DE GRÁFICAS (TENDENCIA TEMPORAL Y COBERTURA DE EXPEDIENTES PDF)
    g_col5, g_col6 = st.columns(2)

    with g_col5:
        st.subheader("📈 Tendencia Mensual de Recepción de Oficios")
        df_temp = df_filtered.dropna(subset=['fecha_recibido_dt']).copy()
        if not df_temp.empty:
            df_temp['Mes_Año'] = df_temp['fecha_recibido_dt'].dt.to_period('M').astype(str)
            trend_df = df_temp.groupby('Mes_Año').size().reset_index(name='Total')
            fig_trend = px.line(trend_df, x='Mes_Año', y='Total', markers=True, line_shape='spline', color_discrete_sequence=['#10B981'])
            fig_trend.update_layout(xaxis_title="Mes / Año", yaxis_title="Cantidad de Oficios Recibidos")
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("Sin datos de fecha para generar la tendencia.")

    with g_col6:
        st.subheader("📎 Cobertura de Expedientes Escaneados (PDF)")
        df_filtered['tiene_pdf'] = df_filtered['archivo_escaneado'].apply(lambda x: 'PDF Adjunto' if pd.notna(x) and str(x).strip() != '' else 'Sin Expediente')
        pdf_counts = df_filtered['tiene_pdf'].value_counts().reset_index()
        pdf_counts.columns = ['Estado_PDF', 'Total']
        fig_pdf = px.bar(pdf_counts, x='Estado_PDF', y='Total', color='Estado_PDF', text='Total', color_discrete_map={'PDF Adjunto': '#10B981', 'Sin Expediente': '#EF4444'})
        fig_pdf.update_layout(showlegend=False, xaxis_title="", yaxis_title="Total de Oficios")
        st.plotly_chart(fig_pdf, use_container_width=True)

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
        with col_u1:
            estado_sel = st.selectbox("1. Estado *", ["-- Seleccione --"] + estados_list)
        muns_list = get_municipios(estado_sel) if estado_sel != "-- Seleccione --" else []
        with col_u2:
            municipio_sel = st.selectbox("2. Municipio *", ["-- Seleccione --"] + muns_list)
        ejidos_list = get_ejidos(estado_sel, municipio_sel) if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --" else []
        with col_u3:
            ejido_sel = st.selectbox("3. Ejido *", ["-- Seleccione --"] + ejidos_list)

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
                dgcat_clean = dgcat_folio.strip().replace("/", "_").replace(" ", "")
                nombre_archivo = f"{dgcat_clean}_{archivo_escaneado.name}"
                with open(os.path.join(UPLOADS_DIR, nombre_archivo), "wb") as f:
                    f.write(archivo_escaneado.getbuffer())

                with engine.begin() as conn:
                    res = conn.execute(text("SELECT id FROM oficios WHERE UPPER(TRIM(dgcat)) = :dg"), {"dg": dgcat_folio.strip().upper()}).fetchone()
                    if res:
                        conn.execute(text("""
                            UPDATE oficios 
                            SET estado = :e, municipio = :m, ejido = :ej, archivo_escaneado = :arch, 
                                observaciones = COALESCE(NULLIF(:obs, ''), observaciones)
                            WHERE id = :id
                        """), {
                            "e": estado_sel, "m": municipio_sel, "ej": ejido_sel, 
                            "arch": nombre_archivo, "obs": observaciones_capturista, "id": res[0]
                        })
                    else:
                        conn.execute(text("""
                            INSERT INTO oficios (estado, municipio, ejido, dgcat, fecha_entrega, scg, siscat, observaciones, archivo_escaneado)
                            VALUES (:e, :m, :ej, :dg, :fe, 'SUBIDO', 'SUBIDO', :obs, :arch)
                        """), {
                            "e": estado_sel, "m": municipio_sel, "ej": ejido_sel, 
                            "dg": dgcat_folio.strip().upper(), "fe": str(date.today()), 
                            "obs": observaciones_capturista, "arch": nombre_archivo
                        })
                st.success(f"✅ Documento PDF vinculado al Folio DGCAT: {dgcat_folio.strip().upper()}")

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

    if modo_accion == "🗑️ Eliminar Registro" and oficio_sel is not None:
        st.error(f"⚠️ Eliminar oficio **{oficio_sel['dgcat']}** (ID #{oficio_sel['id']}).")
        if st.checkbox("✔ Confirmo eliminar registro"):
            if st.button("🚨 ELIMINAR REGISTRO", type="primary"):
                eliminar_oficio(oficio_sel['id'])
                st.success("Eliminado correctamente.")
                st.rerun()
        st.stop()

    es_oficinas_centrales_admin = st.checkbox("🏢 Trámite Perteneciente a OFICINAS CENTRALES")

    scg_options = pd.read_sql("SELECT nombre FROM cat_scg ORDER BY nombre", engine)['nombre'].tolist()
    siscat_options = pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", engine)['nombre'].tolist()
    tramite_options = pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", engine)['nombre'].tolist()
    estados_list = get_estados()

    val_id_reg = str(oficio_sel['id_registro']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['id_registro'])) else ""
    val_dgcat = str(oficio_sel['dgcat']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['dgcat'])) else "DGCAT/100/"
    val_no_oficio = str(oficio_sel['no_oficio']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['no_oficio'])) else ""
    val_obs = str(oficio_sel['observaciones']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['observaciones'])) else ""

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

    st.markdown("---")
    dgcat_folio = st.text_input("Folio DGCAT", value=val_dgcat)

    with st.form("form_oficio_admin", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            id_num = st.text_input("ID Numérico", value=val_id_reg)
            no_oficio = st.text_input("NO. OFICIO", value=val_no_oficio)
            f_entrega = st.date_input("FECHA DE ENTREGA", value=date.today())
            scg_sel = st.selectbox("Bandeja SCG", scg_options)

        with col2:
            f_recibido = st.date_input("FECHA DE RECIBIDO", value=date.today())
            siscat_sel = st.selectbox("Estatus SISCAT", siscat_options)
            tipo_tramite = st.selectbox("TIPO DE TRÁMITE", tramite_options)

        observaciones = st.text_area("OBSERVACIONES", value=val_obs)
        archivo_nuevo = st.file_uploader("Subir/Reemplazar PDF Escaneado", type=["pdf"])

        btn_label = "💾 Guardar Registro" if modo_accion == "➕ Nuevo Registro" else "✏️ Guardar Cambios"
        if st.form_submit_button(btn_label, type="primary"):
            archivo_final = ""
            if archivo_nuevo is not None:
                archivo_final = f"{dgcat_folio.strip().replace('/', '_')}_{archivo_nuevo.name}"
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
                        "id_r": id_num, "e": estado_sel, "m": municipio_sel, "ej": ejido_sel, "no_of": no_oficio,
                        "dg": dgcat_folio.strip().upper(), "f_ent": str(f_entrega), "f_rec": str(f_recibido),
                        "scg_val": scg_sel, "sis_val": siscat_sel, "tram": tipo_tramite, "obs": observaciones,
                        "arch": archivo_final, "id": oficio_sel['id']
                    })
                    st.success("✅ Modificado correctamente.")
                else:
                    conn.execute(text("""
                        INSERT INTO oficios (id_registro, estado, municipio, ejido, no_oficio, dgcat, fecha_entrega, fecha_recibido, scg, siscat, tipo_tramite, observaciones, archivo_escaneado)
                        VALUES (:id_r, :e, :m, :ej, :no_of, :dg, :f_ent, :f_rec, :scg_val, :sis_val, :tram, :obs, :arch)
                    """), {
                        "id_r": id_num, "e": estado_sel, "m": municipio_sel, "ej": ejido_sel, "no_of": no_oficio,
                        "dg": dgcat_folio.strip().upper(), "f_ent": str(f_entrega), "f_rec": str(f_recibido),
                        "scg_val": scg_sel, "sis_val": siscat_sel, "tram": tipo_tramite, "obs": observaciones,
                        "arch": archivo_final
                    })
                    st.success("✅ Guardado correctamente.")

# -----------------------------------------------------------------------------
# 4. CONSULTA Y EXPEDIENTES
# -----------------------------------------------------------------------------
elif menu == "🔍 Consulta y Expedientes":
    st.title("🔍 Consulta de Expedientes DGCAT")
    df_data = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", engine)
    df_data.columns = [c.lower() for c in df_data.columns]
    st.dataframe(df_data, use_container_width=True)
    
    excel_file = generar_excel_ejecutivo(df_data)
    with open(excel_file, "rb") as f:
        st.download_button("📊 Descargar Reporte Ejecutivo en Excel (.xlsx)", f, file_name="Reporte_DGCAT_Ejecutivo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

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
                conn.execute(text("INSERT INTO cat_scg (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": n_scg.upper()})
            st.success("Guardado.")
            st.rerun()
            
    with tab2:
        st.subheader("Opciones Actuales de SISCAT")
        df_siscat = pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", engine)
        st.dataframe(df_siscat, use_container_width=True)
        
        n_siscat = st.text_input("Nueva Opción para SISCAT")
        if st.button("Guardar Opción SISCAT") and n_siscat:
            with engine.begin() as conn:
                conn.execute(text("INSERT INTO cat_siscat (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": n_siscat.upper()})
            st.success("Guardado.")
            st.rerun()
            
    with tab3:
        st.subheader("Tipos de Trámite Registrados")
        df_tramite = pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", engine)
        st.dataframe(df_tramite, use_container_width=True)
        
        n_tramite = st.text_input("Nuevo Tipo de Trámite")
        if st.button("Guardar Trámite") and n_tramite:
            with engine.begin() as conn:
                conn.execute(text("INSERT INTO cat_tramite (nombre) VALUES (:n) ON CONFLICT DO NOTHING;"), {"n": n_tramite.upper()})
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
                st.success(msg_pwd)
                st.rerun()

        st.markdown("---")
        usuarios_para_borrar = [u for u in lista_todos_usuarios if u.lower() != 'admin' and u.lower() != st.session_state['username'].lower()]
        if usuarios_para_borrar:
            user_a_borrar = st.selectbox("Usuario a eliminar:", usuarios_para_borrar)
            if st.button("❌ Confirmar Eliminación"):
                eliminar_usuario(user_a_borrar)
                st.success("Usuario eliminado.")
                st.rerun()
