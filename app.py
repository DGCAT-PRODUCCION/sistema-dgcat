import streamlit as st
import pandas as pd
import sqlite3
import os
import plotly.express as px
from datetime import date
from database import init_db, verificar_login, registrar_nuevo_usuario, get_connection, generar_excel_ejecutivo

# Configuración de página institucional
st.set_page_config(
    page_title="DGCAT - Control de Gestión Institucional",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Adaptativos (Compatibles con Light y Dark)
st.markdown("""
<style>
    .header-box {
        text-align: center;
        padding: 10px 0 15px 0;
        margin-bottom: 15px;
    }
    .header-title {
        font-size: 2.2rem;
        font-weight: 800;
        margin-bottom: 4px;
        letter-spacing: 0.8px;
    }
    .header-subtitle {
        color: #10B981;
        font-size: 1.15rem;
        font-weight: 600;
        margin-bottom: 15px;
    }
    .header-line {
        height: 3px;
        background: linear-gradient(90deg, transparent 0%, #10B981 50%, transparent 100%);
        border: none;
        margin-bottom: 25px;
    }

    .stSelectbox label, .stTextInput label, .stDateInput label, .stTextArea label {
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }

    .metric-card {
        background: rgba(16, 185, 129, 0.08);
        border: 1px solid rgba(16, 185, 129, 0.3);
        border-top: 4px solid #10B981;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .metric-card h3 {
        font-size: 0.82rem;
        margin-bottom: 4px;
        text-transform: uppercase;
        letter-spacing: 1px;
        opacity: 0.8;
    }
    .metric-card .number {
        font-size: 2.1rem;
        font-weight: 800;
        color: #10B981;
    }
    .metric-card .subtitle {
        font-size: 0.78rem;
        opacity: 0.75;
        margin-top: 4px;
    }

    .session-badge {
        background-color: #10B981;
        color: #FFFFFF;
        padding: 10px 18px;
        border-radius: 8px;
        font-weight: 700;
        font-size: 0.95rem;
        margin-bottom: 25px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 2px 5px rgba(0,0,0,0.15);
    }

    [data-testid="stSidebarHeader"] img, [data-testid="stImage"] img {
        max-height: 100px !important;
        object-fit: contain !important;
    }
</style>
""", unsafe_allow_html=True)

UPLOADS_DIR = "uploads"
os.makedirs(UPLOADS_DIR, exist_ok=True)

# Inicializar Base de Datos
init_db()

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""
if "nombre" not in st.session_state:
    st.session_state["nombre"] = ""
if "rol" not in st.session_state:
    st.session_state["rol"] = "operador"

# -----------------------------------------------------------------------------
# PANTALLA DE INICIO DE SESIÓN
# -----------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------
# PANEL INTERNO PRIVADO AL INICIAR SESIÓN
# -----------------------------------------------------------------------------
st.markdown(f"""
<div class="session-badge">
    <span>🟢 SESIÓN ACTIVA | <strong>{st.session_state['nombre']}</strong> ({st.session_state['username']})</span>
    <span>PERFIL: <strong>{st.session_state['rol'].upper()}</strong></span>
</div>
""", unsafe_allow_html=True)

# Logo en Sidebar
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
conn = get_connection()

# DEFINICIÓN DINÁMICA DE MENÚ SEGÚN EL ROL DEL USUARIO
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
else:  # admin
    menu_options = [
        "📈 Dashboard Ejecutivo", 
        "📝 Registro Completo de Oficios", 
        "📄 Carga de Archivo Escaneado (PDF)", 
        "🔍 Consulta y Expedientes", 
        "⚙️ Gestión de Catálogos",
        "👥 Alta de Usuarios"
    ]

menu = st.sidebar.radio("Menú de Opciones", menu_options)

# -----------------------------------------------------------------------------
# 1. DASHBOARD EJECUTIVO COMPLETO
# -----------------------------------------------------------------------------
if menu == "📈 Dashboard Ejecutivo":
    st.title("🏛️ Tablero de Control Directivo DGCAT")
    st.caption("Monitoreo institucional de oficios de respuesta, bandeja de geógrafos, sistemas y estatus SISCAT.")
    
    df_raw = pd.read_sql("SELECT * FROM oficios", conn)
    
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
            sel_estado = st.selectbox("Estado / Entidad Federativa", estados_opts)
        with f_col2:
            scg_opts = ["TODOS"] + sorted(list(df_raw['scg'].dropna().unique()))
            sel_scg = st.selectbox("Estatus Bandeja SCG", scg_opts)
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
    en_sistemas = len(df_filtered[df_filtered['scg'].str.contains('SISTEMAS', case=False, na=False)])
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

    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.subheader("🍩 Distribución por Bandeja SCG")
        scg_counts = df_filtered['scg'].value_counts().reset_index()
        scg_counts.columns = ['Bandeja', 'Cantidad']
        fig_donut = px.pie(scg_counts, values='Cantidad', names='Bandeja', hole=0.45, color_discrete_sequence=px.colors.qualitative.Bold)
        fig_donut.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_donut, use_container_width=True)

    with g_col2:
        st.subheader("📊 Control e Integración en SISCAT")
        siscat_counts = df_filtered['siscat'].value_counts().reset_index()
        siscat_counts.columns = ['Estatus', 'Cantidad']
        fig_bar_siscat = px.bar(siscat_counts, x='Cantidad', y='Estatus', orientation='h', color='Estatus', text='Cantidad', color_discrete_sequence=px.colors.qualitative.Vivid)
        fig_bar_siscat.update_layout(showlegend=False, xaxis_title="Número de Oficios", yaxis_title="")
        st.plotly_chart(fig_bar_siscat, use_container_width=True)

    st.markdown("---")

    g_col3, g_col4 = st.columns(2)

    with g_col3:
        st.subheader("Top 10 Estados con Mayor Carga Registrada")
        top_estados = df_filtered['estado'].value_counts().head(10).reset_index()
        top_estados.columns = ['Estado', 'Oficios']
        fig_top_estados = px.bar(top_estados, x='Estado', y='Oficios', color='Oficios', text='Oficios', color_continuous_scale='Greens')
        fig_top_estados.update_layout(xaxis_title="", yaxis_title="Total Oficios", coloraxis_showscale=False)
        st.plotly_chart(fig_top_estados, use_container_width=True)

    with g_col4:
        st.subheader("📑 Volumetría por Tipo de Trámite")
        obs_counts = df_filtered['observaciones'].value_counts().head(8).reset_index()
        obs_counts.columns = ['Trámite', 'Cantidad']
        fig_obs = px.pie(obs_counts, values='Cantidad', names='Trámite', color_discrete_sequence=px.colors.qualitative.Prism)
        fig_obs.update_traces(textinfo='label+value')
        st.plotly_chart(fig_obs, use_container_width=True)

# -----------------------------------------------------------------------------
# 2. MÓDULO OPERADOR / CAPTURISTA
# -----------------------------------------------------------------------------
elif menu == "📄 Carga de Archivo Escaneado (PDF)":
    st.title("📄 Carga Institucional de Expediente PDF (Operador)")
    st.caption("Todos los campos marcados con (*) son estrictamente OBLIGATORIOS.")
    
    es_oficinas_centrales = st.checkbox("🏢 Trámite Perteneciente a OFICINAS CENTRALES", help="Asigna 'OFICINAS CENTRALES' automáticamente en Estado, Municipio y Ejido.")

    estados_list = pd.read_sql("SELECT DISTINCT TRIM(ESTADO) as ESTADO FROM cat_ubicaciones ORDER BY ESTADO", conn)['ESTADO'].tolist()
    
    col_u1, col_u2, col_u3 = st.columns(3)
    
    if es_oficinas_centrales:
        estado_sel = "OFICINAS CENTRALES"
        municipio_sel = "OFICINAS CENTRALES"
        ejido_sel = "OFICINAS CENTRALES"
        
        with col_u1:
            st.text_input("1. Estado *", value="OFICINAS CENTRALES", disabled=True)
        with col_u2:
            st.text_input("2. Municipio *", value="OFICINAS CENTRALES", disabled=True)
        with col_u3:
            st.text_input("3. Ejido / Núcleo Agrario *", value="OFICINAS CENTRALES", disabled=True)
    else:
        with col_u1:
            estado_sel = st.selectbox("1. Estado *", ["-- Seleccione --"] + estados_list)
            
        muns_list = []
        if estado_sel != "-- Seleccione --":
            muns_list = pd.read_sql("SELECT DISTINCT TRIM(MUNICIPIO) as MUNICIPIO FROM cat_ubicaciones WHERE TRIM(ESTADO) = ? ORDER BY MUNICIPIO", conn, params=(estado_sel,))['MUNICIPIO'].tolist()
            
        with col_u2:
            municipio_sel = st.selectbox("2. Municipio *", ["-- Seleccione --"] + muns_list)
            
        ejidos_list = []
        if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --":
            ejidos_list = pd.read_sql("SELECT DISTINCT TRIM(`NUCLEO AGRARIO`) as EJIDO FROM cat_ubicaciones WHERE TRIM(ESTADO) = ? AND TRIM(MUNICIPIO) = ? ORDER BY EJIDO", conn, params=(estado_sel, municipio_sel))['EJIDO'].tolist()
            
        with col_u3:
            ejido_sel = st.selectbox("3. Ejido / Núcleo Agrario *", ["-- Seleccione --"] + ejidos_list)

    st.markdown("---")

    with st.form("form_carga_operador", clear_on_submit=True):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            dgcat_folio = st.text_input("Folio DGCAT * (Ejemplo: DGCAT/100/0316/2026)", value="DGCAT/100/")
        with col_f2:
            archivo_escaneado = st.file_uploader("Adjuntar Expediente Escaneado (PDF) *", type=["pdf"])

        observaciones_capturista = st.text_area("Observaciones (Opcional)")

        if st.form_submit_button("📤 Subir y Vincular Documento PDF", type="primary"):
            if not es_oficinas_centrales and (estado_sel == "-- Seleccione --" or municipio_sel == "-- Seleccione --" or ejido_sel == "-- Seleccione --"):
                st.error("⚠️ Debe seleccionar Estado, Municipio y Ejido (o marcar Oficinas Centrales).")
            elif not dgcat_folio or dgcat_folio.strip() == "DGCAT/100/":
                st.error("⚠️ Debe ingresar un folio DGCAT completo (ejemplo: DGCAT/100/0316/2026).")
            elif archivo_escaneado is None:
                st.error("⚠️ Debe adjuntar un archivo PDF escaneado.")
            else:
                dgcat_clean = dgcat_folio.strip().replace("/", "_").replace(" ", "")
                nombre_archivo = f"{dgcat_clean}_{archivo_escaneado.name}"
                path_destino = os.path.join(UPLOADS_DIR, nombre_archivo)
                
                with open(path_destino, "wb") as f:
                    f.write(archivo_escaneado.getbuffer())

                cursor = conn.cursor()
                cursor.execute("SELECT id FROM oficios WHERE UPPER(TRIM(dgcat)) = ?", (dgcat_folio.strip().upper(),))
                row_ex = cursor.fetchone()
                
                if row_ex:
                    cursor.execute("""
                        UPDATE oficios 
                        SET estado = ?, municipio = ?, ejido = ?, archivo_escaneado = ?, observaciones = COALESCE(NULLIF(?, ''), observaciones)
                        WHERE id = ?
                    """, (estado_sel, municipio_sel, ejido_sel, nombre_archivo, observaciones_capturista, row_ex[0]))
                else:
                    cursor.execute("""
                        INSERT INTO oficios (estado, municipio, ejido, dgcat, fecha_entrega, scg, siscat, observaciones, archivo_escaneado)
                        VALUES (?, ?, ?, ?, ?, 'SUBIDO', 'SUBIDO', ?, ?)
                    """, (estado_sel, municipio_sel, ejido_sel, dgcat_folio.strip().upper(), str(date.today()), observaciones_capturista, nombre_archivo))
                
                conn.commit()
                st.success(f"✅ Documento PDF vinculado exitosamente al Folio DGCAT: {dgcat_folio.strip().upper()}")

# -----------------------------------------------------------------------------
# 3. MÓDULO ADMINISTRADOR / SUPERVISOR (MODIFICAR Y ELIMINAR OFICIOS CON ADVERTENCIA)
# -----------------------------------------------------------------------------
elif menu == "📝 Registro Completo de Oficios":
    st.title("📝 Registro y Edición Avanzada de Oficios")
    st.caption("Seleccione una opción para Crear Nuevo Registro, Modificar un oficio existente o Eliminarlo de la base de datos.")

    from database import eliminar_oficio
    from datetime import datetime

    # Cargar todos los oficios registrados para permitir selección/edición/eliminación
    df_oficios = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", conn)

    modo_accion = st.radio(
        "Modo de Operación:", 
        ["➕ Nuevo Registro", "✏️ Modificar Registro Existente", "🗑️ Eliminar Registro"], 
        horizontal=True
    )

    st.markdown("---")

    oficio_sel = None
    if modo_accion in ["✏️ Modificar Registro Existente", "🗑️ Eliminar Registro"]:
        if df_oficios.empty:
            st.warning("No hay registros en la base de datos para editar o eliminar.")
            st.stop()

        # Crear lista de opciones descriptivas: "ID #10 - DGCAT/100/0316/2026 - NO_OFICIO: 123 - JALISCO"
        df_oficios['display_name'] = "ID #" + df_oficios['id'].astype(str) + " | Folio: " + df_oficios['dgcat'].astype(str) + " | Oficio: " + df_oficios['no_oficio'].fillna('').astype(str) + " | Estado: " + df_oficios['estado'].fillna('').astype(str)
        opciones_oficios = df_oficios['display_name'].tolist()
        
        seleccion = st.selectbox("🔍 Seleccione el Oficio a procesar:", opciones_oficios)
        id_oficio_seleccionado = int(seleccion.split(" | ")[0].replace("ID #", ""))
        oficio_sel = df_oficios[df_oficios['id'] == id_oficio_seleccionado].iloc[0]

    # SECCIÓN: ELIMINAR REGISTRO CON ADVERTENCIA
    if modo_accion == "🗑️ Eliminar Registro" and oficio_sel is not None:
        st.error(f"⚠️ **ADVERTENCIA DE SEGURIDAD**: Está a punto de eliminar el oficio **{oficio_sel['dgcat']}** (ID #{oficio_sel['id']}). Esta acción no se puede deshacer.")
        
        col_del1, col_del2 = st.columns([1, 2])
        with col_del1:
            confirmar_borrado = st.checkbox("✔ Confirmo que deseo ELIMINAR permanentemente este registro.")
        
        if confirmar_borrado:
            if st.button("🚨 ELIMINAR DEFINITIVAMENTE DE LA BASE DE DATOS", type="primary", use_container_width=True):
                exito_del, msg_del = eliminar_oficio(oficio_sel['id'])
                if exito_del:
                    st.success(f"✅ {msg_del}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg_del}")
        st.stop()

    # FORMULARIO PARA AGREGAR O MODIFICAR
    es_oficinas_centrales_admin = st.checkbox("🏢 Trámite Perteneciente a OFICINAS CENTRALES", help="Asigna 'OFICINAS CENTRALES' a Estado, Municipio y Ejido.")

    scg_options = pd.read_sql("SELECT nombre FROM cat_scg ORDER BY nombre", conn)['nombre'].tolist()
    siscat_options = pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", conn)['nombre'].tolist()
    tramite_options = pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", conn)['nombre'].tolist()
    estados_list = pd.read_sql("SELECT DISTINCT TRIM(ESTADO) as ESTADO FROM cat_ubicaciones ORDER BY ESTADO", conn)['ESTADO'].tolist()

    # Precargar valores si estamos en modo modificación
    val_id_reg = str(oficio_sel['id_registro']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['id_registro'])) else ""
    val_dgcat = str(oficio_sel['dgcat']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['dgcat'])) else "DGCAT/100/"
    val_no_oficio = str(oficio_sel['no_oficio']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['no_oficio'])) else ""
    val_obs = str(oficio_sel['observaciones']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['observaciones'])) else ""

    col_u1, col_u2, col_u3 = st.columns(3)
    if es_oficinas_centrales_admin:
        estado_sel = "OFICINAS CENTRALES"
        municipio_sel = "OFICINAS CENTRALES"
        ejido_sel = "OFICINAS CENTRALES"
        with col_u1:
            st.text_input("1. Estado", value="OFICINAS CENTRALES", disabled=True)
        with col_u2:
            st.text_input("2. Municipio", value="OFICINAS CENTRALES", disabled=True)
        with col_u3:
            st.text_input("3. Ejido", value="OFICINAS CENTRALES", disabled=True)
    else:
        def_estado_idx = 0
        if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['estado'] in estados_list:
            def_estado_idx = estados_list.index(oficio_sel['estado']) + 1

        with col_u1:
            estado_sel = st.selectbox("1. Estado", ["-- Seleccione --"] + estados_list, index=def_estado_idx)

        muns_list = []
        def_mun_idx = 0
        if estado_sel != "-- Seleccione --":
            muns_list = pd.read_sql("SELECT DISTINCT TRIM(MUNICIPIO) as MUNICIPIO FROM cat_ubicaciones WHERE TRIM(ESTADO) = ? ORDER BY MUNICIPIO", conn, params=(estado_sel,))['MUNICIPIO'].tolist()
            if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['municipio'] in muns_list:
                def_mun_idx = muns_list.index(oficio_sel['municipio']) + 1

        with col_u2:
            municipio_sel = st.selectbox("2. Municipio", ["-- Seleccione --"] + muns_list, index=def_mun_idx)

        ejidos_list = []
        def_ejido_idx = 0
        if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --":
            ejidos_list = pd.read_sql("SELECT DISTINCT TRIM(`NUCLEO AGRARIO`) as EJIDO FROM cat_ubicaciones WHERE TRIM(ESTADO) = ? AND TRIM(MUNICIPIO) = ? ORDER BY EJIDO", conn, params=(estado_sel, municipio_sel))['EJIDO'].tolist()
            if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['ejido'] in ejidos_list:
                def_ejido_idx = ejidos_list.index(oficio_sel['ejido']) + 1

        with col_u3:
            ejido_sel = st.selectbox("3. Ejido", ["-- Seleccione --"] + ejidos_list, index=def_ejido_idx)

    st.markdown("---")

    dgcat_folio = st.text_input("Folio DGCAT para Vincular (Ejemplo: DGCAT/100/0316/2026)", value=val_dgcat)

    pdf_preexistente = ""
    if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['archivo_escaneado']):
        pdf_preexistente = str(oficio_sel['archivo_escaneado'])

    if dgcat_folio and dgcat_folio.strip() != "DGCAT/100/":
        cursor = conn.cursor()
        cursor.execute("SELECT archivo_escaneado FROM oficios WHERE UPPER(TRIM(dgcat)) = ?", (dgcat_folio.strip().upper(),))
        result = cursor.fetchone()
        if result and result[0]:
            pdf_preexistente = result[0]
            st.success(f"📎 **PDF Detectado:** Se encontró el expediente `{pdf_preexistente}`.")
            path_pdf = os.path.join(UPLOADS_DIR, pdf_preexistente)
            if os.path.exists(path_pdf):
                with open(path_pdf, "rb") as f:
                    st.download_button("👁️ Descargar / Ver PDF Adjunto", f, file_name=pdf_preexistente, mime="application/pdf")

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
        archivo_nuevo = st.file_uploader("Subir/Reemplazar PDF Escaneado (Opcional)", type=["pdf"])

        btn_label = "💾 Guardar Nuevo Registro" if modo_accion == "➕ Nuevo Registro" else "✏️ Guardar Cambios del Registro"
        
        if st.form_submit_button(btn_label, type="primary"):
            # ADVERTENCIA Y CONFIRMACIÓN EN MODO MODIFICACIÓN
            archivo_final = pdf_preexistente
            if archivo_nuevo is not None:
                archivo_final = f"{dgcat_folio.strip().replace('/', '_')}_{archivo_nuevo.name}"
                with open(os.path.join(UPLOADS_DIR, archivo_final), "wb") as f:
                    f.write(archivo_nuevo.getbuffer())

            cursor = conn.cursor()
            
            if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None:
                cursor.execute("""
                    UPDATE oficios 
                    SET id_registro = ?, estado = ?, municipio = ?, ejido = ?, no_oficio = ?, dgcat = ?, fecha_entrega = ?, fecha_recibido = ?, scg = ?, siscat = ?, tipo_tramite = ?, observaciones = ?, archivo_escaneado = ?
                    WHERE id = ?
                """, (id_num, estado_sel, municipio_sel, ejido_sel, no_oficio, dgcat_folio.strip().upper(), str(f_entrega), str(f_recibido), scg_sel, siscat_sel, tipo_tramite, observaciones, archivo_final, oficio_sel['id']))
                conn.commit()
                st.success(f"✅ ¡El oficio **{dgcat_folio.strip().upper()}** (ID #{oficio_sel['id']}) ha sido MODIFICADO exitosamente!")
            else:
                cursor.execute("""
                    INSERT INTO oficios (id_registro, estado, municipio, ejido, no_oficio, dgcat, fecha_entrega, fecha_recibido, scg, siscat, tipo_tramite, observaciones, archivo_escaneado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (id_num, estado_sel, municipio_sel, ejido_sel, no_oficio, dgcat_folio.strip().upper(), str(f_entrega), str(f_recibido), scg_sel, siscat_sel, tipo_tramite, observaciones, archivo_final))
                conn.commit()
                st.success(f"✅ ¡Nuevo registro guardado con éxito con el Folio DGCAT **{dgcat_folio.strip().upper()}**!")

# -----------------------------------------------------------------------------
# 4. CONSULTA Y EXPEDIENTES
# -----------------------------------------------------------------------------
elif menu == "🔍 Consulta y Expedientes":
    st.title("🔍 Consulta de Expedientes DGCAT")
    st.caption("Visualice y descargue el reporte consolidado con formato ejecutivo institucional.")
    
    df_data = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", conn)
    st.dataframe(df_data, use_container_width=True)
    
    st.markdown("---")
    
    excel_file = generar_excel_ejecutivo(df_data, "Reporte_DGCAT_Ejecutivo.xlsx")
    with open(excel_file, "rb") as f:
        st.download_button(
            label="📊 Descargar Reporte Ejecutivo en Excel (.xlsx)",
            data=f,
            file_name="Reporte_DGCAT_Ejecutivo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

# -----------------------------------------------------------------------------
# 5. GESTIÓN DE CATÁLOGOS
# -----------------------------------------------------------------------------
elif menu == "⚙️ Gestión de Catálogos":
    st.title("⚙️ Agregar Opciones a Catálogos")
    t1, t2, t3 = st.tabs(["SCG", "SISCAT", "Trámites"])
    cursor = conn.cursor()
    with t1:
        n_scg = st.text_input("Nueva Opción para SCG")
        if st.button("Guardar Opción SCG") and n_scg:
            cursor.execute("INSERT OR IGNORE INTO cat_scg (nombre) VALUES (?)", (n_scg.upper(),))
            conn.commit()
            st.rerun()
    with t2:
        n_siscat = st.text_input("Nueva Opción para SISCAT")
        if st.button("Guardar Opción SISCAT") and n_siscat:
            cursor.execute("INSERT OR IGNORE INTO cat_siscat (nombre) VALUES (?)", (n_siscat.upper(),))
            conn.commit()
            st.rerun()
    with t3:
        n_tramite = st.text_input("Nuevo Tipo de Trámite")
        if st.button("Guardar Trámite") and n_tramite:
            cursor.execute("INSERT OR IGNORE INTO cat_tramite (nombre) VALUES (?)", (n_tramite.upper(),))
            conn.commit()
            st.rerun()

# -----------------------------------------------------------------------------
# 6. MÓDULO EXCLUSIVO DE ADMINISTRADOR: ALTA DE USUARIOS Y GESTIÓN DE CONTRASEÑAS
# -----------------------------------------------------------------------------
elif menu == "👥 Alta de Usuarios":
    st.title("👥 Gestión Completa de Usuarios (Exclusivo Administrador)")
    st.caption("Desde este panel puedes consultar contraseñas, restablecerlas, crear nuevos usuarios o eliminar cuentas.")

    from database import eliminar_usuario, cambiar_password_usuario

    col_u1, col_u2 = st.columns([1.1, 1.2])

    # SECCIÓN DE REGISTRO
    with col_u1:
        st.write("### ➕ Registrar Nuevo Usuario")
        reg_nombre = st.text_input("Nombre Completo *", key="admin_reg_nom")
        reg_user = st.text_input("Usuario Único *", key="admin_reg_usr")
        reg_pwd = st.text_input("Contraseña *", key="admin_reg_pwd")
        
        # AHORA INCLUYE EL NUEVO ROL DE SUPERVISOR
        reg_rol = st.selectbox(
            "Perfil de Usuario *", 
            ["operador", "supervisor", "admin"], 
            format_func=lambda x: "Capturista / Operador (Solo PDF)" if x=="operador" else ("Supervisor / Directivo (Acceso Total sin Usuarios)" if x=="supervisor" else "Administrador (Acceso Completo y Contraseñas)")
        )
        
        if st.button("💾 Crear Cuenta de Usuario", type="primary", use_container_width=True):
            if not reg_nombre or not reg_user or not reg_pwd:
                st.warning("⚠️ Todos los campos son obligatorios.")
            else:
                exito, msg = registrar_nuevo_usuario(reg_user, reg_pwd, reg_nombre, reg_rol)
                if exito:
                    st.success(f"✅ {msg}")
                    st.info("📧 Notificación enviada a actmosaicocatastral@gmail.com")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

    # SECCIÓN DE CONSULTA, CAMBIO Y ELIMINACIÓN DE CONTRASEÑAS
    with col_u2:
        st.write("### 📋 Directorio y Contraseñas de Usuarios")
        df_users = pd.read_sql("""
            SELECT 
                username as Usuario, 
                nombre_completo as 'Nombre Completo', 
                UPPER(rol) as Perfil, 
                COALESCE(password_plain, '***') as Contraseña 
            FROM usuarios
        """, conn)
        st.dataframe(df_users, use_container_width=True)

        st.markdown("---")
        
        # RESTABLECER CONTRASEÑA DE UN USUARIO
        st.write("### 🔑 Cambiar / Restablecer Contraseña")
        lista_todos_usuarios = df_users['Usuario'].tolist()
        usr_cambiar_pwd = st.selectbox("Seleccione el usuario al que desea cambiarle la contraseña:", lista_todos_usuarios)
        nueva_pwd_input = st.text_input("Nueva Contraseña:", key="pwd_reset_input")
        
        if st.button("🔄 Actualizar Contraseña del Usuario", use_container_width=True):
            if not nueva_pwd_input:
                st.warning("Escriba la nueva contraseña.")
            else:
                exito_pwd, msg_pwd = cambiar_password_usuario(usr_cambiar_pwd, nueva_pwd_input)
                if exito_pwd:
                    st.success(f"✅ {msg_pwd}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg_pwd}")

        st.markdown("---")
        
        # ELIMINAR USUARIO
        st.write("### 🗑️ Eliminar Usuario del Sistema")
        usuarios_para_borrar = [u for u in lista_todos_usuarios if u.lower() != 'admin' and u.lower() != st.session_state['username'].lower()]

        if not usuarios_para_borrar:
            st.info("ℹ️ No hay usuarios secundarios disponibles para eliminar.")
        else:
            user_a_borrar = st.selectbox("Seleccione el usuario que desea eliminar:", usuarios_para_borrar, key="sel_borrar_usr")
            if st.button("❌ Confirmar Eliminación de Usuario", type="primary", use_container_width=True):
                exito_del, msg_del = eliminar_usuario(user_a_borrar)
                if exito_del:
                    st.success(f"✅ {msg_del}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg_del}")

conn.close()