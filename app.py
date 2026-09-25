import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import datetime, date
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
    obtener_oficios_paginados,
    obtener_seguimiento_paginado
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

# CONTROL DE ROL EN FRONTEND
if st.session_state["rol"] == "operador":
    menu_options = ["📍 Seguimiento de Ubicación de Predio"]
elif st.session_state["rol"] == "supervisor":
    menu_options = [
        "📈 Dashboard Ejecutivo", 
        "📝 Registro Completo de Oficios", 
        "📍 Seguimiento de Ubicación de Predio", 
        "🔍 Consulta y Expedientes", 
        "🗂️ Consulta de Seguimiento de Predio",
        "⚙️ Gestión de Catálogos"
    ]
else:  # admin
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

# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES DE UBICACIÓN
# -----------------------------------------------------------------------------
@st.cache_data(ttl=600)
def get_cat_ubicaciones_df():
    try:
        df = pd.read_sql("SELECT * FROM cat_ubicaciones", engine)
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        col_edo = [c for c in df.columns if 'estado' in c or 'edo' in c][0] if any('estado' in c or 'edo' in c for c in df.columns) else df.columns[0]
        col_mun = [c for c in df.columns if 'muni' in c][0] if any('muni' in c for c in df.columns) else df.columns[1]
        col_eji = [c for c in df.columns if 'nucleo' in c or 'ejido' in c or 'nuc' in c][0] if any('nucleo' in c or 'ejido' in c or 'nuc' in c for c in df.columns) else df.columns[2]
        
        return pd.DataFrame({
            'estado': df[col_edo].astype(str).str.strip().str.upper(),
            'municipio': df[col_mun].astype(str).str.strip().str.upper(),
            'ejido': df[col_eji].astype(str).str.strip().str.upper()
        })
    except Exception:
        return pd.DataFrame(columns=['estado', 'municipio', 'ejido'])

def get_estados():
    df = get_cat_ubicaciones_df()
    return [] if df.empty else sorted(list(df['estado'].dropna().unique()))

def get_municipios(estado):
    df = get_cat_ubicaciones_df()
    if df.empty: return []
    filtered = df[df['estado'].str.upper() == str(estado).strip().upper()]
    return sorted(list(filtered['municipio'].dropna().unique()))

def get_ejidos(estado, municipio):
    df = get_cat_ubicaciones_df()
    if df.empty: return []
    filtered = df[(df['estado'].str.upper() == str(estado).strip().upper()) & (df['municipio'].str.upper() == str(municipio).strip().upper())]
    return sorted(list(filtered['ejido'].dropna().unique()))

# -----------------------------------------------------------------------------
# 1. DASHBOARD EJECUTIVO
# -----------------------------------------------------------------------------
if menu == "📈 Dashboard Ejecutivo":
    col_dash_t, col_dash_btn = st.columns([3, 1])
    with col_dash_t:
        st.title("🏛️ Tablero de Control Directivo DGCAT")
    with col_dash_btn:
        st.write("<br>", unsafe_allow_html=True)
        if st.button("🔄 Actualizar ahora", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.success("Información actualizada correctamente.")
            st.rerun()

    st.caption("Monitoreo institucional de oficios de respuesta, expedientes PDF digitalizados, SCG y SISCAT.")
    
    df_raw = pd.read_sql("SELECT * FROM oficios", engine)
    df_raw.columns = [c.lower() for c in df_raw.columns]
    
    if df_raw.empty:
        st.warning("No hay registros disponibles para generar métricas.")
        st.stop()

    with st.expander("🔍 **Filtros de Control Ejecutivo**", expanded=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            sel_estado = st.selectbox("Estado", ["TODOS"] + sorted(list(df_raw['estado'].dropna().unique())))
        with f_col2:
            sel_scg = st.selectbox("Estatus SCG", ["TODOS"] + sorted(list(df_raw['scg'].dropna().unique())))
        with f_col3:
            sel_siscat = st.selectbox("Estatus SISCAT", ["TODOS"] + sorted(list(df_raw['siscat'].dropna().unique())))

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
    sistemas_or_val = len(df_filtered[df_filtered['sistemas_or'].notna() & (df_filtered['sistemas_or'].astype(str).str.strip() != '')]) if 'sistemas_or' in df_filtered.columns else 0
    
    pct_scg = round((concluidos_scg / total_oficios * 100), 1) if total_oficios > 0 else 0
    pct_siscat = round((subidos_siscat / total_oficios * 100), 1) if total_oficios > 0 else 0

    with kpi1: st.markdown(f'<div class="metric-card"><h3>TOTAL OFICIOS</h3><div class="number">{total_oficios:,}</div><div class="subtitle">Registros Atendidos</div></div>', unsafe_allow_html=True)
    with kpi2: st.markdown(f'<div class="metric-card"><h3>EFICIENCIA SCG</h3><div class="number">{pct_scg}%</div><div class="subtitle">{concluidos_scg:,} Concluidos</div></div>', unsafe_allow_html=True)
    with kpi3: st.markdown(f'<div class="metric-card"><h3>AVANCE SISCAT</h3><div class="number">{pct_siscat}%</div><div class="subtitle">{subidos_siscat:,} Procesados</div></div>', unsafe_allow_html=True)
    with kpi4: st.markdown(f'<div class="metric-card"><h3>EN SISTEMAS</h3><div class="number" style="color:#EF4444;">{en_sistemas:,}</div><div class="subtitle">Pendientes</div></div>', unsafe_allow_html=True)
    with kpi5: st.markdown(f'<div class="metric-card" style="border-top-color:#8B5CF6;"><h3>SISTEMAS / OR</h3><div class="number" style="color:#8B5CF6;">{sistemas_or_val:,}</div><div class="subtitle">Trámite Clasificado</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    g_col1, g_col2 = st.columns(2)
    with g_col1:
        st.subheader("🍩 Distribución por Bandeja SCG")
        scg_counts = df_filtered['scg'].value_counts().reset_index()
        scg_counts.columns = ['Bandeja', 'Cantidad']
        fig_pie_scg = px.pie(scg_counts, values='Cantidad', names='Bandeja', color_discrete_sequence=px.colors.qualitative.G10)
        st.plotly_chart(fig_pie_scg, use_container_width=True)

    with g_col2:
        st.subheader("📊 Control e Integración en SISCAT")
        siscat_counts = df_filtered['siscat'].value_counts().reset_index()
        siscat_counts.columns = ['Estatus', 'Cantidad']
        fig_bar_siscat = px.bar(siscat_counts, x='Cantidad', y='Estatus', orientation='h', color='Estatus', text='Cantidad', color_discrete_sequence=px.colors.qualitative.Vivid)
        st.plotly_chart(fig_bar_siscat, use_container_width=True)

# -----------------------------------------------------------------------------
# 2. SEGUIMIENTO DE UBICACIÓN DE PREDIO
# -----------------------------------------------------------------------------
elif menu == "📍 Seguimiento de Ubicación de Predio":
    st.title("📍 Seguimiento de Ubicación de Predio")
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
    dgcat_folio = st.text_input("Folio DGCAT *", value="DGCAT/100/", key="dgcat_seguimiento")

    dgcat_check = dgcat_folio.strip().upper()
    id_existente = existe_folio_seguimiento(dgcat_check) if dgcat_check and dgcat_check != "DGCAT/100/" else None
    if id_existente:
        st.info(f"ℹ️ Este folio ya existe (ID Interno #{id_existente}). Al guardar se ACTUALIZARÁ ese registro.")

    with st.form("form_seguimiento_predio", clear_on_submit=True):
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
                fecha_hoy = date.today().strftime('%d/%m/%Y')
                nombre_archivo = None
                if archivo_escaneado is not None:
                    dgcat_clean = dgcat_check.replace("/", "_").replace(" ", "")
                    nombre_archivo = f"{dgcat_clean}_{archivo_escaneado.name}"
                    with open(os.path.join(UPLOADS_DIR, nombre_archivo), "wb") as f:
                        f.write(archivo_escaneado.getbuffer())

                datos = {
                    "dgcat": dgcat_check,
                    "estado": estado_sel.strip().upper(),
                    "municipio": municipio_sel.strip().upper(),
                    "ejido": ejido_sel.strip().upper(),
                    "observaciones": observaciones_capturista.strip().upper() if observaciones_capturista else "",
                    "registrado_por": st.session_state["username"],
                }
                if id_existente:
                    datos["fecha_actualizacion"] = fecha_hoy
                    if nombre_archivo: datos["archivo_escaneado"] = nombre_archivo
                    ok, msg = guardar_seguimiento_predio(datos, id_registro=id_existente)
                else:
                    datos["fecha_registro"] = fecha_hoy
                    datos["fecha_actualizacion"] = fecha_hoy
                    datos["archivo_escaneado"] = nombre_archivo or ""
                    ok, msg = guardar_seguimiento_predio(datos)

                if ok:
                    st.cache_data.clear()
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

# -----------------------------------------------------------------------------
# 3. REGISTRO Y EDICIÓN AVANZADA DE OFICIOS
# -----------------------------------------------------------------------------
elif menu == "📝 Registro Completo de Oficios":
    st.title("📝 Registro y Edición Avanzada de Oficios")

    df_oficios = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", engine)
    df_oficios.columns = [c.lower() for c in df_oficios.columns]

    modo_accion = st.radio("Modo:", ["➕ Nuevo Registro", "✏️ Modificar Registro Existente", "🗑️ Eliminar Registro"], horizontal=True)

    # Limpiar estado si se selecciona un Nuevo Registro para evitar heredar folios anteriores
    if modo_accion == "➕ Nuevo Registro":
        if "dgcat_registro_completo" in st.session_state and st.session_state["dgcat_registro_completo"] != "DGCAT/100/":
            st.session_state["dgcat_registro_completo"] = "DGCAT/100/"

    oficio_sel = None
    if modo_accion in ["✏️ Modificar Registro Existente", "🗑️ Eliminar Registro"]:
        if df_oficios.empty:
            st.warning("No hay registros en la base de datos.")
            st.stop()

        df_oficios['display_name'] = "Folio: " + df_oficios['dgcat'].astype(str) + " | Oficio: " + df_oficios['no_oficio'].fillna('').astype(str) + " | Estado: " + df_oficios['estado'].fillna('').astype(str)

        texto_busqueda = st.text_input("🔎 Buscar por folio DGCAT, No. de oficio o Estado:")
        df_busqueda = df_oficios
        if texto_busqueda.strip():
            patron = texto_busqueda.strip().upper()
            df_busqueda = df_oficios[
                df_oficios['dgcat'].astype(str).str.upper().str.contains(patron, na=False) |
                df_oficios['no_oficio'].astype(str).str.upper().str.contains(patron, na=False) |
                df_oficios['estado'].astype(str).str.upper().str.contains(patron, na=False)
            ]
            if df_busqueda.empty:
                st.warning("⚠️ No se encontró ningún oficio que coincida.")
                st.stop()

        opciones_oficios = df_busqueda['display_name'].tolist()
        seleccion = st.selectbox("🔍 Seleccione el Oficio:", opciones_oficios)
        idx_match = df_busqueda[df_busqueda['display_name'] == seleccion].index[0]
        oficio_sel = df_busqueda.loc[idx_match]

    if modo_accion == "🗑️ Eliminar Registro" and oficio_sel is not None:
        st.error(f"⚠️ Eliminar oficio **{oficio_sel['dgcat']}**.")
        confirmar = st.checkbox("Confirmo que deseo eliminar este registro de forma DEFINITIVA.")
        if st.button("🚨 ELIMINAR DEFINITIVAMENTE", type="primary", disabled=not confirmar):
            ok, msg = eliminar_oficio(int(oficio_sel['id']), usuario_actual=st.session_state['username'])
            if ok:
                st.cache_data.clear()
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
        def_estado_idx = estados_list.index(oficio_sel['estado']) + 1 if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['estado'] in estados_list) else 0
        with col_u1: estado_sel = st.selectbox("1. Estado", ["-- Seleccione --"] + estados_list, index=def_estado_idx)

        muns_list = get_municipios(estado_sel) if estado_sel != "-- Seleccione --" else []
        def_mun_idx = muns_list.index(oficio_sel['municipio']) + 1 if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['municipio'] in muns_list) else 0
        with col_u2: municipio_sel = st.selectbox("2. Municipio", ["-- Seleccione --"] + muns_list, index=def_mun_idx)

        ejidos_list = get_ejidos(estado_sel, municipio_sel) if estado_sel != "-- Seleccione --" and municipio_sel != "-- Seleccione --" else []
        def_ejido_idx = ejidos_list.index(oficio_sel['ejido']) + 1 if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['ejido'] in ejidos_list) else 0
        with col_u3: ejido_sel = st.selectbox("3. Ejido", ["-- Seleccione --"] + ejidos_list, index=def_ejido_idx)

    # Asignación correcta del Folio DGCAT según el modo
    if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None:
        val_dgcat = str(oficio_sel['dgcat']) if pd.notna(oficio_sel['dgcat']) else "DGCAT/100/"
    else:
        val_dgcat = "DGCAT/100/"

    dgcat_folio = st.text_input("Folio DGCAT", value=val_dgcat, key="dgcat_registro_completo")

    dgcat_check = dgcat_folio.strip().upper()
    id_excluir = int(oficio_sel['id']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None) else None
    id_duplicado = existe_folio_oficio(dgcat_check, excluir_id=id_excluir) if dgcat_check and dgcat_check != "DGCAT/100/" else None
    
    if modo_accion == "➕ Nuevo Registro" and id_duplicado:
        st.warning(f"⚠️ El folio **{dgcat_check}** ya existe. Cambie a 'Modificar Registro Existente' para editarlo.")

    # Catálogos obligatorios con '-- Seleccione --'
    scg_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_scg ORDER BY nombre", engine)['nombre'].tolist()
    siscat_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_siscat ORDER BY nombre", engine)['nombre'].tolist()
    
    try:
        sistemas_or_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_sistemas_or ORDER BY nombre", engine)['nombre'].tolist()
    except Exception:
        sistemas_or_options = ["-- Seleccione --", "SISTEMAS", "OR"]

    tramite_options = ["-- Seleccione --"] + pd.read_sql("SELECT nombre FROM cat_tramite ORDER BY nombre", engine)['nombre'].tolist()

    limpiar_al_guardar = True if modo_accion == "➕ Nuevo Registro" else False

    val_id_reg = str(oficio_sel['id_registro']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['id_registro'])) else ""
    val_no_oficio = str(oficio_sel['no_oficio']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['no_oficio'])) else ""
    val_obs = str(oficio_sel['observaciones']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and pd.notna(oficio_sel['observaciones'])) else ""

    with st.form("form_oficio_admin", clear_on_submit=limpiar_al_guardar):
        idx_scg = scg_options.index(oficio_sel['scg']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['scg'] in scg_options) else 0
        idx_siscat = siscat_options.index(oficio_sel['siscat']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['siscat'] in siscat_options) else 0
        idx_sistemas_or = sistemas_or_options.index(oficio_sel['sistemas_or']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and 'sistemas_or' in oficio_sel and oficio_sel['sistemas_or'] in sistemas_or_options) else 0
        idx_tramite = tramite_options.index(oficio_sel['tipo_tramite']) if (modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None and oficio_sel['tipo_tramite'] in tramite_options) else 0

        col1, col2 = st.columns(2)
        with col1:
            id_num = st.text_input("ID Numérico", value=val_id_reg)
            no_oficio = st.text_input("NO. OFICIO", value=val_no_oficio)
            f_entrega = st.date_input("FECHA DE ENTREGA", value=date.today(), format="DD/MM/YYYY")
            scg_sel = st.selectbox("Bandeja SCG *", scg_options, index=idx_scg)

        with col2:
            f_recibido = st.date_input("FECHA DE RECIBIDO", value=date.today(), format="DD/MM/YYYY")
            siscat_sel = st.selectbox("Estatus SISCAT *", siscat_options, index=idx_siscat)
            sistemas_or_sel = st.selectbox("SISTEMAS/OR *", sistemas_or_options, index=idx_sistemas_or)
            tipo_tramite = st.selectbox("TIPO DE TRÁMITE *", tramite_options, index=idx_tramite)

        observaciones = st.text_area("OBSERVACIONES", value=val_obs)
        archivo_nuevo = st.file_uploader("Subir/Reemplazar PDF Escaneado", type=["pdf"])

        btn_label = "💾 Guardar Registro" if modo_accion == "➕ Nuevo Registro" else "✏️ Guardar Cambios"
        if st.form_submit_button(btn_label, type="primary"):
            if modo_accion == "➕ Nuevo Registro" and id_duplicado:
                st.error(f"❌ No se guardó: el folio '{dgcat_check}' ya existe.")
            elif not es_oficinas_centrales_admin and (estado_sel == "-- Seleccione --" or municipio_sel == "-- Seleccione --" or ejido_sel == "-- Seleccione --"):
                st.error("⚠️ Debe seleccionar Estado, Municipio y Ejido.")
            elif scg_sel == "-- Seleccione --" or siscat_sel == "-- Seleccione --" or sistemas_or_sel == "-- Seleccione --" or tipo_tramite == "-- Seleccione --":
                st.error("⚠️ Debe seleccionar una opción válida en los catálogos marcados con *.")
            elif not dgcat_check or dgcat_check == "DGCAT/100/":
                st.error("⚠️ Debe ingresar un folio DGCAT completo.")
            else:
                str_f_entrega = f_entrega.strftime('%Y-%m-%d') if f_entrega else None
                str_f_recibido = f_recibido.strftime('%Y-%m-%d') if f_recibido else None

                archivo_final = ""
                if archivo_nuevo is not None:
                    archivo_final = f"{dgcat_check.replace('/', '_')}_{archivo_nuevo.name}"
                    with open(os.path.join(UPLOADS_DIR, archivo_final), "wb") as f:
                        f.write(archivo_nuevo.getbuffer())

                datos = {
                    "id_registro": id_num.strip().upper(), "estado": estado_sel.strip().upper(), 
                    "municipio": municipio_sel.strip().upper(), "ejido": ejido_sel.strip().upper(), 
                    "no_oficio": no_oficio.strip().upper(), "dgcat": dgcat_check,
                    "fecha_entrega": str_f_entrega, "fecha_recibido": str_f_recibido,
                    "scg": scg_sel.strip().upper(), "siscat": siscat_sel.strip().upper(), 
                    "sistemas_or": sistemas_or_sel.strip().upper(), "tipo_tramite": tipo_tramite.strip().upper(), 
                    "observaciones": observaciones.strip().upper(), "archivo_escaneado": archivo_final
                }
                
                if modo_accion == "✏️ Modificar Registro Existente" and oficio_sel is not None:
                    ok, msg = guardar_oficio(datos, id_oficio=int(oficio_sel['id']), usuario_actual=st.session_state['username'])
                else:
                    ok, msg = guardar_oficio(datos, usuario_actual=st.session_state['username'])

                if ok:
                    st.cache_data.clear()
                    st.success(f"✅ {msg}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")

# -----------------------------------------------------------------------------
# 4. CONSULTA DE EXPEDIENTES DGCAT Y DESCARGA
# -----------------------------------------------------------------------------
elif menu == "🔍 Consulta y Expedientes":
    st.title("🔍 Consulta de Expedientes DGCAT y Descarga de PDF")
    
    fecha_actual_str = datetime.now().strftime('%d de %B de %Y')
    st.markdown(f"**Información actualizada:** {fecha_actual_str}")
    
    col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
    with col_f1:
        txt_buscar = st.text_input("🔎 Buscar en catálogo:")
    with col_f2:
        page_size = st.selectbox("Registros por página:", [200, 300, 500], index=0)
    with col_f3:
        page_num = st.number_input("Página:", min_value=1, value=1, step=1)

    df_data, total_recs, total_pags = obtener_oficios_paginados(page=page_num, page_size=page_size, busqueda=txt_buscar)
    st.caption(f"Mostrando página {page_num} de {total_pags} | Total de registros encontrados: {total_recs:,}")

    if not df_data.empty:
        cols_order = ['id_registro', 'estado', 'municipio', 'ejido', 'no_oficio', 'dgcat', 'fecha_entrega', 'fecha_recibido', 'scg', 'siscat', 'sistemas_or', 'tipo_tramite', 'observaciones', 'archivo_escaneado']
        cols_presentes = [c for c in cols_order if c in df_data.columns]
        df_display = df_data[cols_presentes].copy()

        # Formatear visualmente las fechas a DD/MM/AAAA para el usuario
        for col_fecha in ['fecha_entrega', 'fecha_recibido']:
            if col_fecha in df_display.columns:
                df_display[col_fecha] = pd.to_datetime(df_display[col_fecha], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')

        for col in df_display.select_dtypes(include=['object']).columns:
            df_display[col] = df_display[col].astype(str).str.upper()

        st.dataframe(df_display, use_container_width=True)
    else:
        st.warning("No se encontraron registros con los filtros aplicados.")

    st.markdown("---")
    st.subheader("📁 Visor y Descarga de Expedientes PDF")
    
    if not df_data.empty and 'archivo_escaneado' in df_data.columns:
        df_pdfs = df_data[df_data['archivo_escaneado'].notna() & (df_data['archivo_escaneado'] != '') & (df_data['archivo_escaneado'] != 'NONE')]
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
                    st.warning("⚠️ Favor de subir el oficio.")
        else:
            st.info("ℹ️ Favor de subir el oficio.")

    st.markdown("---")
    if st.button("📊 Generar Reporte Ejecutivo Excel"):
        df_exp = pd.read_sql("SELECT * FROM oficios ORDER BY id DESC", engine)
        excel_file = generar_excel_ejecutivo(df_exp)
        with open(excel_file, "rb") as f:
            st.download_button("📥 Descargar Excel", f, file_name="Reporte_DGCAT_Ejecutivo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------------------------------------------------------
# 5. CONSULTA DE SEGUIMIENTO DE PREDIO
# -----------------------------------------------------------------------------
elif menu == "🗂️ Consulta de Seguimiento de Predio":
    st.title("🗂️ Consulta de Seguimiento de Ubicación de Predio")

    fecha_actual_str = datetime.now().strftime('%d de %B de %Y')
    st.markdown(f"**Información actualizada:** {fecha_actual_str}")

    col_s1, col_s2, col_s3 = st.columns([2, 1, 1])
    with col_s1:
        txt_buscar_seg = st.text_input("🔎 Buscar seguimiento:")
    with col_s2:
        page_size_seg = st.selectbox("Registros por página:", [200, 300, 500], key="ps_seg")
    with col_s3:
        page_num_seg = st.number_input("Página:", min_value=1, value=1, step=1, key="pn_seg")

    df_seg, total_recs_s, total_pags_s = obtener_seguimiento_paginado(page=page_num_seg, page_size=page_size_seg, busqueda=txt_buscar_seg)
    st.caption(f"Mostrando página {page_num_seg} de {total_pags_s} | Total de registros: {total_recs_s:,}")

    if not df_seg.empty:
        cols_order_seg = ['dgcat', 'estado', 'municipio', 'ejido', 'fecha_registro', 'fecha_actualizacion', 'observaciones', 'archivo_escaneado', 'registrado_por']
        cols_presentes_seg = [c for c in cols_order_seg if c in df_seg.columns]
        df_seg_display = df_seg[cols_presentes_seg].copy()

        for col in df_seg_display.select_dtypes(include=['object']).columns:
            df_seg_display[col] = df_seg_display[col].astype(str).str.upper()

        st.dataframe(df_seg_display, use_container_width=True)

        st.markdown("---")
        st.subheader("📁 Archivo PDF de Seguimiento")
        df_seg_pdf = df_seg[df_seg['archivo_escaneado'].notna() & (df_seg['archivo_escaneado'] != '')]
        if not df_seg_pdf.empty:
            pdf_seg_sel = st.selectbox("Seleccionar expediente de seguimiento:", df_seg_pdf['archivo_escaneado'].unique())
            pdf_seg_path = os.path.join(UPLOADS_DIR, pdf_seg_sel)
            if os.path.exists(pdf_seg_path):
                with open(pdf_seg_path, "rb") as f:
                    st.download_button("📥 Descargar PDF", f, file_name=pdf_seg_sel, mime="application/pdf")
            else:
                st.warning("⚠️ Favor de subir el oficio.")
        else:
            st.info("ℹ️ Favor de subir el oficio.")
    else:
        st.warning("No hay registros de seguimiento.")

# -----------------------------------------------------------------------------
# 6. GESTIÓN DE CATÁLOGOS
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
        
        col_crear, _ = st.columns([2, 1])
        with col_crear:
            with st.form(key=f"form_add_{tabla_db}", clear_on_submit=True):
                nueva_opcion = st.text_input(f"Nueva Opción para {label_catalogo}")
                if st.form_submit_button(f"Guardar Opción {label_catalogo}"):
                    if nueva_opcion.strip():
                        val_clean = nueva_opcion.strip().upper()
                        try:
                            with engine.begin() as conn:
                                conn.execute(text(f"INSERT INTO {tabla_db} (nombre) VALUES (:n) ON CONFLICT DO NOTHING"), {"n": val_clean})
                            st.cache_data.clear()
                            st.success(f"✅ Opción '{val_clean}' guardada.")
                            st.rerun()
                        except Exception as err:
                            st.error(f"Error: {err}")

        st.markdown("---")
        if options:
            st.subheader(f"✏️ Modificar o 🗑️ Eliminar Opción")
            opcion_sel = st.selectbox(f"Seleccione elemento:", ["-- Seleccione --"] + options, key=f"sel_{tabla_db}")

            with st.form(key=f"form_edit_{tabla_db}", clear_on_submit=True):
                nuevo_nombre = st.text_input("Nuevo nombre para modificar:")
                col_b1, col_b2, _ = st.columns([1, 1, 2])
                with col_b1: btn_mod = st.form_submit_button("✏️ Guardar Modificación")
                with col_b2: btn_eli = st.form_submit_button("🗑️ Eliminar Opción", type="primary")

                if btn_mod and opcion_sel != "-- Seleccione --" and nuevo_nombre.strip():
                    ok, msg = actualizar_opcion_catalogo(tabla_db, opcion_sel, nuevo_nombre.strip().upper())
                    if ok:
                        st.cache_data.clear()
                        st.success(f"✅ {msg}")
                        st.rerun()
                if btn_eli and opcion_sel != "-- Seleccione --":
                    ok, msg = eliminar_opcion_catalogo(tabla_db, opcion_sel)
                    if ok:
                        st.cache_data.clear()
                        st.success(f"✅ {msg}")
                        st.rerun()

    with t1: render_catalogo_produccion("cat_scg", "SCG")
    with t2: render_catalogo_produccion("cat_siscat", "SISCAT")
    with t3: render_catalogo_produccion("cat_sistemas_or", "SISTEMAS/OR")
    with t4: render_catalogo_produccion("cat_tramite", "Tipos de Trámite")

# -----------------------------------------------------------------------------
# 7. GESTIÓN COMPLETA DE USUARIOS
# -----------------------------------------------------------------------------
elif menu == "👥 Alta de Usuarios":
    if st.session_state["rol"] != "admin":
        st.error("⛔ ACCESO NO AUTORIZADO: Este módulo requiere privilegios de Administrador.")
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
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    with col_u2:
        st.write("### 📋 Directorio y Control de Usuarios")
        df_users = pd.read_sql("SELECT username, nombre_completo, rol FROM usuarios", engine)
        st.dataframe(df_users, use_container_width=True)

        st.markdown("---")
        st.write("### 🔑 Modificar / Restablecer Contraseña")
        lista_usr = df_users['username'].tolist() if not df_users.empty else []
        usr_mod = st.selectbox("Usuario a modificar:", ["-- Seleccione --"] + lista_usr)
        pwd_nueva = st.text_input("Nueva contraseña:", type="password")

        if st.button("🔄 Actualizar Contraseña") and usr_mod != "-- Seleccione --" and pwd_nueva:
            ok, msg = cambiar_password_usuario(usr_mod, pwd_nueva)
            if ok:
                st.success(f"✅ {msg}")
                st.rerun()

        st.markdown("---")
        st.write("### 🗑️ Eliminar Usuario")
        usr_del = st.selectbox("Usuario a eliminar:", ["-- Seleccione --"] + [u for u in lista_usr if u.lower() != 'admin' and u.lower() != st.session_state['username'].lower()])
        if st.button("❌ Eliminar Usuario", type="primary") and usr_del != "-- Seleccione --":
            ok, msg = eliminar_usuario(usr_del)
            if ok:
                st.success(f"✅ {msg}")
                st.rerun()
