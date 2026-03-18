"""
Dashboard Interactivo de Visualización de Programación 2026
Versión Streamlit - PRODUCCIÓN
Optimizado para Streamlit Cloud
Autor: IST - Equipo de Especialidades Técnicas
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import duckdb
import re
import io

# Configuración de página
st.set_page_config(
    page_title="Dashboard Programación HO IST 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado
st.markdown("""
    <style>
    .main {background-color: #F8F9FA;}
    h1 {color: #2E86AB;}
    .stMetric {background-color: white; padding: 10px; border-radius: 5px;}
    .detalle-section {
        background-color: #f0f7fb;
        padding: 15px;
        border-radius: 8px;
        margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================================
# SISTEMA DE AUTENTICACIÓN
# ============================================================================

def check_password():
    """Retorna True si el usuario ingresó las credenciales correctas."""

    def password_entered():
        """Revisa si las credenciales son correctas."""
        if (
            st.session_state["username"] == st.secrets["credentials"]["username"]
            and st.session_state["password"] == st.secrets["credentials"]["password"]
        ):
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Eliminar contraseña de session_state
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    # Página de login con estilo premium
    st.markdown("""
        <style>
        .login-container {
            max-width: 450px;
            margin: 50px auto;
            padding: 40px;
            background-color: white;
            border-radius: 20px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            border: 1px solid #E0E0E0;
        }
        .login-header {
            text-align: center;
            margin-bottom: 30px;
        }
        .login-title {
            color: #2E86AB;
            font-size: 24px;
            font-weight: bold;
            margin-top: 10px;
        }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<div class="login-header">', unsafe_allow_html=True)
        st.markdown('<h1>📊</h1>', unsafe_allow_html=True)
        st.markdown('<div class="login-title">Acceso Programación HO 2026</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.text_input("Usuario", key="username", placeholder="Ingrese su RUT/ID")
        st.text_input("Contraseña", type="password", key="password", placeholder="••••••••")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Ingresar al Dashboard", use_container_width=True, on_click=password_entered):
            pass # on_click maneja la lógica

        if "password_correct" in st.session_state and not st.session_state["password_correct"]:
            st.error("❌ Credenciales incorrectas. Intente nuevamente.")
        
        st.markdown('<div style="text-align: center; margin-top: 20px; color: #888; font-size: 12px;">© 2026 IST - Especialidades Técnicas</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    return False

# Solo continuar si está autenticado
if not check_password():
    st.stop()

def construir_url_exportacion(url_sheet):
    """Construye la URL de exportación CSV a partir de la URL de Google Sheets"""
    # Extraer spreadsheet ID
    match_id = re.search(r'/d/([a-zA-Z0-9_-]+)', url_sheet)
    if not match_id:
        st.error("❌ No se pudo extraer el ID del spreadsheet de la URL configurada en secrets.toml")
        st.stop()
    spreadsheet_id = match_id.group(1)

    # Extraer gid (ID de pestaña)
    match_gid = re.search(r'gid=(\d+)', url_sheet)
    gid = match_gid.group(1) if match_gid else '0'

    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"

# ============================================================================
# FUNCIONES DE CARGA Y PROCESAMIENTO
# ============================================================================

def normalizar_columnas(df):
    """
    Normaliza nombres de columnas del CSV de Google Sheets para que coincidan
    con los nombres esperados por el código (con tildes y caracteres especiales).
    Google Sheets exporta CSV sin tildes en los encabezados.
    """
    mapeo_columnas = {
        'Identificador unico (ID) centro de trabajo (CT)': 'Identificador único (ID) centro de trabajo (CT)',
        'Fecha de Evaluacion Cualitativa 2026': 'Fecha de Evaluación Cualitativa 2026',
        'Fecha de Evaluacion Cuantitativa 2026': 'Fecha de Evaluación Cuantitativa 2026',
        'Fecha de ultima Evaluacion Cualitativa': 'Fecha de última Evaluación Cualitativa',
        'Fecha de ultima Evaluacion Cuantitativa': 'Fecha de última Evaluación Cuantitativa',
        'Fecha de ultima Evaluacion Vigilancia de Salud': 'Fecha de última Evaluación Vigilancia de Salud',
        'Motivo de programacion': 'Motivo de programación',
        'Origen de Inclusion': 'Origen de Inclusión',
        'N de Trabajadores(as) CT': 'N° de Trabajadores(as) CT',
        'N trabajadores que deben ingresar a Vigilancia de Salud Hombres': 'N° trabajadores que deben ingresar a Vigilancia de Salud Hombres',
        'N trabajadores que deben ingresar a Vigilancia de Salud Mujeres': 'N° trabajadores que deben ingresar a Vigilancia de Salud Mujeres',
        'N trabajadores en Vigilancia de Salud Hombres': 'N° trabajadores en Vigilancia de Salud Hombres',
        'N trabajadores en Vigilancia de Salud Mujeres': 'N° trabajadores en Vigilancia de Salud Mujeres',
    }
    df = df.rename(columns=mapeo_columnas)
    return df

def parsear_fecha_flexible(serie):
    """
    Parsea una serie de fechas que puede tener formatos mixtos, priorizando convención chilena:
    - DD-MM-YYYY (formato Excel original)
    - DD/MM/YYYY (formato común nacional / Google Sheets export)
    - YYYY-MM-DD (formato ISO)
    """
    # 1. Intentar formato DD-MM-YYYY primero
    resultado = pd.to_datetime(serie, format='%d-%m-%Y', errors='coerce')

    # 2. Intentar formato DD/MM/YYYY o DD/MM/YY (Nacional)
    mascara_nulos = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara_nulos.any():
        resultado[mascara_nulos] = pd.to_datetime(serie[mascara_nulos], dayfirst=True, errors='coerce')

    return resultado

@st.cache_data(ttl=300)
def cargar_datos():
    """Carga datos desde Google Sheets y los procesa con DuckDB en memoria"""

    try:
        # Leer URL desde secrets
        url_sheet = st.secrets["gsheets"]["url"]
        export_url = construir_url_exportacion(url_sheet)

        # Descargar CSV desde Google Sheets
        df_2026 = pd.read_csv(export_url)

        # Normalizar nombres de columnas (CSV no tiene tildes)
        df_2026 = normalizar_columnas(df_2026)

        # Cargar en DuckDB en memoria para mayor velocidad
        con = duckdb.connect(':memory:')
        con.register('programacion_raw', df_2026)

        # Usar DuckDB para la consulta inicial (aprovecha optimización columnar)
        df_2026 = con.execute("SELECT * FROM programacion_raw").fetchdf()
        con.close()

        # Convertir fechas con parser flexible (maneja formatos mixtos del CSV)
        df_2026['Fecha de Evaluación Cualitativa 2026'] = parsear_fecha_flexible(
            df_2026['Fecha de Evaluación Cualitativa 2026']
        )
        df_2026['Fecha de Evaluación Cuantitativa 2026'] = parsear_fecha_flexible(
            df_2026['Fecha de Evaluación Cuantitativa 2026']
        )

        return df_2026

    except KeyError:
        st.error("❌ No se encontró la URL de Google Sheets en secrets.toml")
        st.info("Configura el archivo `.streamlit/secrets.toml` con la sección [gsheets] y la clave `url`.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Error al cargar datos desde Google Sheets: {str(e)}")
        st.stop()


def normalizar_columnas_seguimiento(df):
    """Normaliza columnas del CSV seguimiento (Google Sheets quita las tildes)."""
    mapeo = {
        'Identificador unico (ID) centro de trabajo (CT)': 'Identificador único (ID) centro de trabajo (CT)',
        'Fecha de Evaluacion Cualitativa 2026': 'Fecha de Evaluación Cualitativa 2026',
        'Fecha de Evaluacion Cuantitativa 2026': 'Fecha de Evaluación Cuantitativa 2026',
        'Fecha de Evaluacion Vigilancia de Salud 2026': 'Fecha de Evaluación Vigilancia de Salud 2026',
        'Numero de trabajadores evaluados 2026 Hombres': 'Número de trabajadores evaluados 2026 Hombres',
        'Numero de trabajadores evaluados 2026 Mujeres': 'Número de trabajadores evaluados 2026 Mujeres',
        'N de Trabajadores CT': 'N° de Trabajadores CT',
        'Grupo Act. Economica': 'Grupo Act. Económica',
        'Protocolo': 'Protocolo_SUSESO_Interno', # Oculto al usuario
        'Codigo Europeo': 'CODIGO_EUROPEO_Interno', # Oculto al usuario
        'Programa': 'Protocolo', # Usar nombre legible como 'Protocolo'
    }
    return df.rename(columns=mapeo)


@st.cache_data(ttl=300)
def cargar_datos_seguimiento():
    """Carga datos de seguimiento desde la hoja SeguimientoHO del Google Sheet.
    Retorna DataFrame vacío si la hoja no existe o no tiene datos aún.
    No detiene la app (soft-fail) para no bloquear la vista de programación."""
    try:
        url_seg = st.secrets["gsheets"].get("seguimiento")
        if not url_seg:
            return pd.DataFrame()
        match_id = re.search(r'/d/([a-zA-Z0-9_-]+)', url_seg)
        if not match_id:
            return pd.DataFrame()
        sid = match_id.group(1)
        match_gid = re.search(r'gid=(\d+)', url_seg)
        gid = match_gid.group(1) if match_gid else '0'
        export_url = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"
        df = pd.read_csv(export_url, dtype=str)
        if df.empty or len(df.columns) < 3:
            return pd.DataFrame()
        df.columns = df.columns.str.strip()
        df = normalizar_columnas_seguimiento(df)
        for col in ['Fecha de Evaluación Cualitativa 2026',
                    'Fecha de Evaluación Cuantitativa 2026',
                    'Fecha de Evaluación Vigilancia de Salud 2026']:
            if col in df.columns:
                df[col] = parsear_fecha_flexible(df[col])
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data
def preparar_datos_eventos(df):
    """Prepara datos en formato largo para visualización - VERSIÓN CORREGIDA"""
    
    # Separar cualitativas y cuantitativas
    cuali = df[df['Fecha de Evaluación Cualitativa 2026'].notna()].copy()
    cuali['tipo'] = 'Cualitativa'
    cuali['fecha'] = cuali['Fecha de Evaluación Cualitativa 2026']
    
    cuanti = df[df['Fecha de Evaluación Cuantitativa 2026'].notna()].copy()
    cuanti['tipo'] = 'Cuantitativa'
    cuanti['fecha'] = cuanti['Fecha de Evaluación Cuantitativa 2026']
    
    # Columnas base que siempre deben existir
    columnas_base = ['fecha', 'tipo', 'Protocolo', 'Region Sucursal', 'Agente',
                     'Nivel de riesgo', 'Comuna CT', 'NOMBRE SUCURSAL', 'Rut Empleador o Rut trabajador(a)','Nombre empleador',
                     'AnexoSUSESO', 'Identificador único (ID) centro de trabajo (CT)',
                     'Gerencia - Cuentas Nacionales', 'Faena Marítimo - Portuaria', 'Holding']

    # Columnas opcionales si existen
    columnas_opcionales = ['Motivo de programación', 'Faena Codelco']
    
    columnas_finales = columnas_base.copy()
    for col in columnas_opcionales:
        if col in df.columns:
            columnas_finales.append(col)
    
    df_eventos = pd.concat([cuali[columnas_finales], cuanti[columnas_finales]], ignore_index=True)
    
    # FIX CRÍTICO: Convertir columnas a string ANTES de cualquier operación
    df_eventos['Comuna CT'] = df_eventos['Comuna CT'].fillna('Sin Información').astype(str)
    df_eventos['AnexoSUSESO'] = df_eventos['AnexoSUSESO'].fillna('Sin Información').astype(str)
    df_eventos['Identificador único (ID) centro de trabajo (CT)'] = df_eventos['Identificador único (ID) centro de trabajo (CT)'].fillna('Sin ID').astype(str)
    df_eventos['Protocolo'] = df_eventos['Protocolo'].fillna('Sin Protocolo').astype(str)
    df_eventos['Region Sucursal'] = df_eventos['Region Sucursal'].fillna('Sin Región').astype(str)
    df_eventos['Agente'] = df_eventos['Agente'].fillna('Sin Agente').astype(str)
    df_eventos['Nivel de riesgo'] = df_eventos['Nivel de riesgo'].fillna('Sin Nivel').astype(str)
    df_eventos['NOMBRE SUCURSAL'] = df_eventos['NOMBRE SUCURSAL'].fillna('Sin Sucursal').astype(str)
    df_eventos['Nombre empleador'] = df_eventos['Nombre empleador'].fillna('Sin Empleador').astype(str)
    df_eventos['Gerencia - Cuentas Nacionales'] = df_eventos['Gerencia - Cuentas Nacionales'].fillna('Sin Gerente').astype(str)
    df_eventos['Holding'] = df_eventos['Holding'].fillna('Sin Holding').astype(str)

    # Convertir columnas opcionales a string si existen
    if 'Faena Codelco' in df_eventos.columns:
        df_eventos['Faena Codelco'] = df_eventos['Faena Codelco'].fillna('Sin Faena').astype(str)
    
    if 'Faena Marítimo - Portuaria' in df_eventos.columns:
        df_eventos['Faena Marítimo - Portuaria'] = df_eventos['Faena Marítimo - Portuaria'].fillna('Sin Información').astype(str)
    
    if 'Motivo de programación' in df_eventos.columns:
        df_eventos['Motivo de programación'] = df_eventos['Motivo de programación'].fillna('Sin Motivo').astype(str)
    
    # Extraer información temporal de forma robusta
    df_eventos['mes'] = df_eventos['fecha'].dt.month
    df_eventos['dia'] = df_eventos['fecha'].dt.day
    
    # Nombres de meses en español
    nombres_meses = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    
    # Mapear nombres de meses y manejar NaN
    df_eventos['mes_nombre'] = df_eventos['mes'].map(nombres_meses).fillna('Sin Mes').astype(str)
    
    # FIX CRÍTICO: NO usar Categorical (causa errores con filtros dinámicos)
    # Eliminar filas con fechas inválidas
    df_eventos = df_eventos[df_eventos['fecha'].notna()].copy()
    
    return df_eventos

def aplicar_filtros(df, anexo_suseso, protocolo, region, tipo, mes, faena_codelco, gerente, maritimo_portuario, holding, nombre_empleador):
    """Aplica los filtros seleccionados de forma flexible para Programación y Seguimiento"""
    if df.empty:
        return df
        
    df_filtrado = df.copy()

    # Filtro Anexo
    if anexo_suseso != 'Todos' and 'AnexoSUSESO' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['AnexoSUSESO'] == anexo_suseso].copy()

    # Filtro Gerente/Gerencia
    if gerente != 'Todos':
        col_ger = 'Gerencia - Cuentas Nacionales' if 'Gerencia - Cuentas Nacionales' in df_filtrado.columns else 'Gerencia'
        if col_ger in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado[col_ger] == gerente].copy()

    if holding != 'Todos' and 'Holding' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['Holding'] == holding].copy()
    
    if nombre_empleador != 'Todos' and 'Nombre empleador' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['Nombre empleador'] == nombre_empleador].copy()
    
    if maritimo_portuario != 'Todos' and 'Faena Marítimo - Portuaria' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['Faena Marítimo - Portuaria'] == maritimo_portuario].copy()

    # Filtro Protocolo/Programa
    if protocolo != 'Todos':
        # Priorizar la columna de texto 'Protocolo' (antiguo 'Programa')
        if 'Protocolo' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['Protocolo'] == protocolo].copy()
    
    if region != 'Todas' and 'Region Sucursal' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['Region Sucursal'] == region].copy()
    
    if tipo != 'Todas' and 'tipo' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['tipo'] == tipo].copy()
    
    if mes != 'Todos':
        meses_es_a_num = {
            'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4,
            'Mayo': 5, 'Junio': 6, 'Julio': 7, 'Agosto': 8,
            'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
        }
        mes_num = meses_es_a_num[mes]
        # En df_seg no hay 'mes' directo, debemos extraerlo de las fechas si es necesario
        if 'mes' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['mes'] == mes_num].copy()
        else:
            # Para df_seg, filtramos si alguna de las fechas de evaluación o programación cae en ese mes
            cols_fecha = [c for c in df_filtrado.columns if 'Fecha' in c]
            if cols_fecha:
                mask = pd.Series(False, index=df_filtrado.index)
                for col_f in cols_fecha:
                    if pd.api.types.is_datetime64_any_dtype(df_filtrado[col_f]):
                        mask |= (df_filtrado[col_f].dt.month == mes_num)
                df_filtrado = df_filtrado[mask].copy()
    
    if faena_codelco != 'Todos':
        col_faena = 'Faena Codelco' if 'Faena Codelco' in df_filtrado.columns else None
        if col_faena:
            df_filtrado = df_filtrado[df_filtrado[col_faena] == faena_codelco].copy()
    
    return df_filtrado

def es_protocolo_plaguicidas(protocolo):
    """Verifica si el protocolo es de plaguicidas"""
    if pd.isna(protocolo) or protocolo == 'Sin Protocolo':
        return False
    return 'PLAGUICIDAS' in str(protocolo).upper()

def contar_evaluaciones(df, protocolo_seleccionado):
    """
    Cuenta evaluaciones según el protocolo:
    - Plaguicidas: cuenta por Centro de Trabajo único
    - Otros: cuenta cada evaluación individual
    """
    if len(df) == 0:
        return 0
    
    # Si hay un protocolo específico seleccionado y es plaguicidas
    if protocolo_seleccionado != 'Todos' and es_protocolo_plaguicidas(protocolo_seleccionado):
        # Contar centros de trabajo únicos, excluyendo 'Sin ID'
        return df[df['Identificador único (ID) centro de trabajo (CT)'] != 'Sin ID']['Identificador único (ID) centro de trabajo (CT)'].nunique()
    
    # Para otros protocolos o cuando está en "Todos"
    return len(df)

def grafico_barras_mensuales(df, protocolo_seleccionado):
    """Genera gráfico de barras mensuales con Plotly - VERSIÓN CORREGIDA"""
    if len(df) == 0:
        return None
    
    # Determinar si es protocolo de plaguicidas
    es_plaguicidas = (protocolo_seleccionado != 'Todos' and 
                      es_protocolo_plaguicidas(protocolo_seleccionado))
    
    # FIX CRÍTICO: Usar observed=False para evitar errores con categorías vacías
    if es_plaguicidas:
        conteo = df.groupby(['mes', 'tipo'], observed=False)['Identificador único (ID) centro de trabajo (CT)'].nunique().reset_index(name='cantidad')
    else:
        conteo = df.groupby(['mes', 'tipo'], observed=False).size().reset_index(name='cantidad')
    
    # Agregar nombre del mes
    nombres_meses_es = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    conteo['mes_nombre'] = conteo['mes'].map(nombres_meses_es).fillna('Sin Mes')
    conteo = conteo.sort_values('mes')
    conteo['mes_nombre'] = conteo['mes_nombre'].astype(str)
    
    # Filtrar filas con cantidad 0
    conteo = conteo[conteo['cantidad'] > 0]
    
    if len(conteo) == 0:
        return None
    
    # Título dinámico
    titulo = 'Carga Mensual de Centros de Trabajo (Plaguicidas)' if es_plaguicidas else 'Carga Mensual de Evaluaciones'
    
    # Orden correcto de meses
    orden_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 
                   'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    
    fig = px.bar(
        conteo,
        x='mes_nombre',
        y='cantidad',
        color='tipo',
        title=titulo,
        labels={'mes_nombre': 'Mes', 'cantidad': 'Cantidad'},
        color_discrete_map={'Cualitativa': '#2E86AB', 'Cuantitativa': '#A23B72'},
        barmode='group',
        height=500,
        category_orders={"mes_nombre": orden_meses}
    )
    
    fig.update_traces(texttemplate='%{y}', textposition='outside')
    fig.update_layout(
        xaxis_tickangle=-45,
        xaxis_title='Mes',
        yaxis_title='Cantidad de Centros de Trabajo' if es_plaguicidas else 'Cantidad de Evaluaciones'
    )
    
    return fig

def grafico_programado_vs_realizado(df_prog, df_seg, protocolo_seleccionado):
    """Gráfico de barras agrupadas: evaluaciones programadas vs realizadas por mes."""
    if df_prog is None or len(df_prog) == 0:
        return None

    nombres_meses_es = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    orden_meses = list(nombres_meses_es.values())

    es_plaguicidas = (protocolo_seleccionado != 'Todos' and
                      es_protocolo_plaguicidas(protocolo_seleccionado))
    id_col = 'Identificador único (ID) centro de trabajo (CT)'

    # ── Programadas ──────────────────────────────────────────────────────────
    if es_plaguicidas:
        prog = (df_prog.groupby(['mes', 'tipo'], observed=False)[id_col]
                .nunique().reset_index(name='cantidad'))
    else:
        prog = (df_prog.groupby(['mes', 'tipo'], observed=False)
                .size().reset_index(name='cantidad'))

    prog['mes_nombre'] = prog['mes'].map(nombres_meses_es).fillna('Sin Mes')
    prog['serie'] = prog['tipo'].map({
        'Cualitativa':  'Cuali. programada',
        'Cuantitativa': 'Cuanti. programada'
    })
    prog = prog[prog['cantidad'] > 0][['mes', 'mes_nombre', 'serie', 'cantidad']]

    # ── Realizadas ───────────────────────────────────────────────────────────
    frames_real = []
    if df_seg is not None and not df_seg.empty:
        col_fc = 'Fecha de Evaluación Cualitativa 2026'
        col_ec = 'Estado Cualitativa'
        col_fq = 'Fecha de Evaluación Cuantitativa 2026'
        col_eq = 'Estado Cuantitativa'

        if col_fc in df_seg.columns and col_ec in df_seg.columns:
            df_rc = df_seg[df_seg[col_ec].str.startswith('Realizada', na=False)].copy()
            df_rc['_fecha'] = pd.to_datetime(df_rc[col_fc], dayfirst=True, errors='coerce')
            df_rc['mes'] = df_rc['_fecha'].dt.month
            df_rc = df_rc.dropna(subset=['mes'])
            df_rc['mes'] = df_rc['mes'].astype(int)
            if es_plaguicidas:
                r = df_rc.groupby('mes')[id_col].nunique().reset_index(name='cantidad')
            else:
                r = df_rc.groupby('mes').size().reset_index(name='cantidad')
            r['serie'] = 'Cuali. realizada'
            frames_real.append(r)

        if col_fq in df_seg.columns and col_eq in df_seg.columns:
            df_rq = df_seg[df_seg[col_eq].str.startswith('Realizada', na=False)].copy()
            df_rq['_fecha'] = pd.to_datetime(df_rq[col_fq], dayfirst=True, errors='coerce')
            df_rq['mes'] = df_rq['_fecha'].dt.month
            df_rq = df_rq.dropna(subset=['mes'])
            df_rq['mes'] = df_rq['mes'].astype(int)
            if es_plaguicidas:
                r = df_rq.groupby('mes')[id_col].nunique().reset_index(name='cantidad')
            else:
                r = df_rq.groupby('mes').size().reset_index(name='cantidad')
            r['serie'] = 'Cuanti. realizada'
            frames_real.append(r)

    # ── Combinar ─────────────────────────────────────────────────────────────
    if frames_real:
        real_all = pd.concat(frames_real, ignore_index=True)
        real_all['mes_nombre'] = real_all['mes'].map(nombres_meses_es).fillna('Sin Mes')
        real_all = real_all[real_all['cantidad'] > 0][['mes', 'mes_nombre', 'serie', 'cantidad']]
        df_chart = pd.concat([prog, real_all], ignore_index=True)
    else:
        df_chart = prog

    if len(df_chart) == 0:
        return None

    color_map = {
        'Cuali. programada':  '#AED6F1',
        'Cuali. realizada':   '#2E86AB',
        'Cuanti. programada': '#F1A9C9',
        'Cuanti. realizada':  '#A23B72',
    }
    titulo = ('CT Programados vs Realizados por Mes (Plaguicidas)'
              if es_plaguicidas else
              'Evaluaciones Programadas vs Realizadas por Mes')

    fig = px.bar(
        df_chart,
        x='mes_nombre',
        y='cantidad',
        color='serie',
        title=titulo,
        labels={'mes_nombre': 'Mes', 'cantidad': 'Cantidad', 'serie': ''},
        color_discrete_map=color_map,
        barmode='group',
        height=500,
        category_orders={
            'mes_nombre': orden_meses,
            'serie': ['Cuali. programada', 'Cuali. realizada',
                      'Cuanti. programada', 'Cuanti. realizada']
        }
    )
    fig.update_traces(texttemplate='%{y}', textposition='outside')
    fig.update_layout(
        xaxis_tickangle=-45,
        xaxis_title='Mes',
        yaxis_title='CT únicos' if es_plaguicidas else 'Evaluaciones',
        legend_title=''
    )
    return fig


def grafico_top_protocolos(df):
    """Genera gráfico de top protocolos - VERSIÓN CORREGIDA"""
    if len(df) == 0:
        return None
    
    # Excluir 'Sin Protocolo'
    df_filtered = df[df['Protocolo'] != 'Sin Protocolo']
    
    if len(df_filtered) == 0:
        return None
    
    top10 = df_filtered['Protocolo'].value_counts().head(10).reset_index()
    top10.columns = ['Protocolo', 'Cantidad']
    
    fig = px.bar(
        top10,
        x='Cantidad',
        y='Protocolo',
        title='Top 10 Protocolos',
        orientation='h',
        color_discrete_sequence=['#F39C12'],
        height=500
    )
    fig.update_traces(texttemplate='%{x}', textposition='outside')
    fig.update_layout(yaxis={'categoryorder': 'total ascending'})
    
    return fig

def mostrar_resumen_detallado(df_filtrado, protocolo_seleccionado, seccion='tab1', df_seg=None):
    """Muestra un resumen detallado de las evaluaciones filtradas - VERSIÓN CORREGIDA"""
    if len(df_filtrado) == 0:
        st.info("No hay evaluaciones para mostrar con los filtros seleccionados")
        return
    
    es_plaguicidas = (protocolo_seleccionado != 'Todos' and 
                      es_protocolo_plaguicidas(protocolo_seleccionado))
    
    if es_plaguicidas:
        st.markdown("### 📋 Detalle de Centros de Trabajo - Plaguicidas")
        st.info("ℹ️ Para plaguicidas, se muestra un informe único por Centro de Trabajo que engloba todos los agentes")
    else:
        st.markdown("### 📋 Detalle de Evaluaciones")
    
    # Resumen por dimensiones clave
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Por Región")
        if es_plaguicidas:
            region_df = df_filtrado[df_filtrado['Identificador único (ID) centro de trabajo (CT)'] != 'Sin ID']
            if len(region_df) > 0:
                region_counts = region_df.groupby('Region Sucursal')['Identificador único (ID) centro de trabajo (CT)'].nunique().reset_index()
                region_counts.columns = ['Región', 'Cantidad CT']
                st.dataframe(region_counts, use_container_width=True, hide_index=True)
            else:
                st.info("No hay datos de CT válidos")
        else:
            region_counts = df_filtrado['Region Sucursal'].value_counts().reset_index()
            region_counts.columns = ['Región', 'Cantidad']
            st.dataframe(region_counts, use_container_width=True, hide_index=True)
    
    with col2:
        if es_plaguicidas:
            st.markdown("#### Agentes por CT")
            agentes_df = df_filtrado[df_filtrado['Identificador único (ID) centro de trabajo (CT)'] != 'Sin ID']
            if len(agentes_df) > 0:
                agentes_por_ct = agentes_df.groupby('Identificador único (ID) centro de trabajo (CT)')['Agente'].count().reset_index()
                agentes_por_ct.columns = ['Centro de Trabajo', 'Cantidad Agentes']
                agentes_por_ct = agentes_por_ct.sort_values('Cantidad Agentes', ascending=False).head(10)
                st.dataframe(agentes_por_ct, use_container_width=True, hide_index=True)
            else:
                st.info("No hay datos de CT válidos")
        else:
            st.markdown("#### Por Agente")
            agente_df = df_filtrado[df_filtrado['Agente'] != 'Sin Agente']
            if len(agente_df) > 0:
                agente_counts = agente_df['Agente'].value_counts().head(10).reset_index()
                agente_counts.columns = ['Agente', 'Cantidad']
                st.dataframe(agente_counts, use_container_width=True, hide_index=True)
            else:
                st.info("No hay datos de agentes")
    
    st.markdown("---")
    
    # Tabla detallada completa
    if es_plaguicidas:
        st.markdown("#### Listado de Centros de Trabajo (Agrupado)")
        
        df_valid = df_filtrado[df_filtrado['Identificador único (ID) centro de trabajo (CT)'] != 'Sin ID']
        
        if len(df_valid) == 0:
            st.warning("No hay centros de trabajo válidos para mostrar")
            return
        
        df_agrupado = df_valid.groupby('Identificador único (ID) centro de trabajo (CT)').agg({
            'fecha': 'first',
            'tipo': 'first',
            'Nombre empleador': 'first',
            'NOMBRE SUCURSAL': 'first',
            'Protocolo': 'first',
            'Region Sucursal': 'first',
            'Comuna CT': 'first',
            'Agente': lambda x: ', '.join(sorted(set(str(a) for a in x if str(a) != 'Sin Agente'))),
            'AnexoSUSESO': 'first',
            'Gerencia - Cuentas Nacionales': 'first',
            'Faena Marítimo - Portuaria': 'first'
        }).reset_index()
        
        df_agrupado['fecha'] = pd.to_datetime(df_agrupado['fecha'], dayfirst=True).dt.strftime('%d-%m-%Y')
        df_agrupado['Cantidad Agentes'] = df_agrupado['Agente'].apply(lambda x: len(x.split(', ')) if x else 0)
        
        df_agrupado.columns = ['ID Centro de Trabajo', 'Fecha', 'Tipo', 'Nombre empleador', 
                               'Sucursal', 'Protocolo', 'Región', 'Comuna', 'Agentes Evaluados', 
                               'Anexo SUSESO', 'Gerente', 'Marítimo Portuario', 'Cantidad Agentes']
        
        st.dataframe(df_agrupado, use_container_width=True, height=400, hide_index=True)
        
        # Preparar Excel en memoria
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_agrupado.to_excel(writer, index=False, sheet_name='Detalle_Plaguicidas')
        
        st.download_button(
            label="📥 Descargar Detalle en Excel",
            data=buffer.getvalue(),
            file_name=f'detalle_plaguicidas_ct_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key=f'download_btn_{seccion}'
        )
    else:
        st.markdown("#### Listado Completo de Evaluaciones")

        columnas_detalle = [
            'fecha', 'tipo', 'Gerencia - Cuentas Nacionales', 'Region Sucursal',
            'Rut Empleador o Rut trabajador(a)', 'Nombre empleador',
            'Identificador único (ID) centro de trabajo (CT)', 'NOMBRE SUCURSAL', 'Agente',
            'Protocolo', 'Comuna CT', 'Nivel de riesgo', 'AnexoSUSESO', 'Faena Marítimo - Portuaria'
        ]
        nombres_columnas = [
            'Fecha', 'Tipo', 'Gerente', 'Región',
            'Rut Empleador o Rut trabajador(a)', 'Nombre empleador',
            'Identificador único (ID) centro de trabajo (CT)', 'Sucursal', 'Agente',
            'Protocolo', 'Comuna', 'Nivel de Riesgo', 'Anexo SUSESO', 'Marítimo Portuario'
        ]

        if 'Motivo de programación' in df_filtrado.columns:
            columnas_detalle.append('Motivo de programación')
            nombres_columnas.append('Motivo de programación')

        if 'Faena Codelco' in df_filtrado.columns:
            columnas_detalle.append('Faena Codelco')
            nombres_columnas.append('Faena Codelco')

        df_tabla = df_filtrado[columnas_detalle].copy()
        df_tabla['fecha'] = df_tabla['fecha'].dt.strftime('%d-%m-%Y')
        df_tabla.columns = nombres_columnas
        df_tabla = df_tabla.sort_values('Fecha')

        # ── Join con seguimiento: agregar 4 columnas de estado/fecha ─────────
        if df_seg is not None and not df_seg.empty:
            id_col_j  = 'Identificador único (ID) centro de trabajo (CT)'
            seg_cols  = [id_col_j, 'Protocolo']
            agente_seg = 'AGENTE' if 'AGENTE' in df_seg.columns else (
                          'Agente' if 'Agente' in df_seg.columns else None)
            extra_seg = ['Estado Cualitativa', 'Fecha de Evaluación Cualitativa 2026',
                         'Estado Cuantitativa', 'Fecha de Evaluación Cuantitativa 2026']
            # Mantener sólo columnas que existen en df_seg
            extra_seg = [c for c in extra_seg if c in df_seg.columns]

            if agente_seg and extra_seg:
                seg_cols.append(agente_seg)
                df_seg_join = df_seg[seg_cols + extra_seg].copy()

                # Normalizar clave de agente a minúsculas para join robusto
                df_tabla['_agente_key'] = df_tabla['Agente'].str.strip().str.lower()
                df_seg_join['_agente_key'] = df_seg_join[agente_seg].str.strip().str.lower()
                df_seg_join = df_seg_join.drop(columns=[agente_seg])

                df_tabla = df_tabla.merge(
                    df_seg_join,
                    left_on=[id_col_j, 'Protocolo', '_agente_key'],
                    right_on=[id_col_j, 'Protocolo', '_agente_key'],
                    how='left'
                )
                df_tabla = df_tabla.drop(columns=['_agente_key'])

                # Renombrar columnas de seguimiento
                rename_seg = {
                    'Estado Cualitativa':                    'Estado Cuali.',
                    'Fecha de Evaluación Cualitativa 2026':  'Fecha real Cuali.',
                    'Estado Cuantitativa':                   'Estado Cuanti.',
                    'Fecha de Evaluación Cuantitativa 2026': 'Fecha real Cuanti.',
                }
                df_tabla = df_tabla.rename(columns={k: v for k, v in rename_seg.items() if k in df_tabla.columns})
        
        st.dataframe(df_tabla, use_container_width=True, height=400, hide_index=True)
        
        # Preparar Excel en memoria
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_tabla.to_excel(writer, index=False, sheet_name='Detalle_Evaluaciones')
            
        st.download_button(
            label="📥 Descargar Detalle en Excel",
            data=buffer.getvalue(),
            file_name=f'detalle_evaluaciones_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key=f'download_btn_{seccion}'
        )

# ============================================================================
# INTERFAZ PRINCIPAL
# ============================================================================

st.title("📊 Dashboard HO 2026 — Programación y Seguimiento")
st.markdown("### IST Organismo de Seguridad y Salud del Trabajo - Higiene Ocupacional")
st.markdown("---")

# Cargar datos con manejo de errores
try:
    with st.spinner('Cargando datos... ⏳'):
        df = cargar_datos()
        df_eventos = preparar_datos_eventos(df)
    
    # Validar que hay datos
    if len(df_eventos) == 0:
        st.error("❌ No hay datos válidos para mostrar. Verifica el archivo de entrada.")
        st.stop()
    
    # ── Sidebar — Filtros Bidireccionales (Cross-filtering) ──────────────────
    # Cada filtro muestra sólo las opciones compatibles con TODOS los demás
    # filtros activos. La actualización es inmediata en un solo click.
    st.sidebar.header("🔍 Filtros")

    _base = df_eventos  # dataset completo, nunca se modifica

    # ── 1. Defaults e inicialización de session_state ─────────────────────────
    _defaults = {
        'ho_gerente':   'Todos',
        'ho_holding':   'Todos',
        'ho_empleador': 'Todos',
        'ho_protocolo': 'Todos',
        'ho_region':    'Todas',
        'ho_anexo':     'Todos',
        'ho_tipo':      'Todas',
        'ho_mes':       'Todos',
        'ho_faena_cod': 'Todos',
        'ho_maritimo':  'Todos',
    }
    for k, v in _defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # Reset pendiente (del botón) — ejecutar ANTES de cualquier widget
    if st.session_state.get('_ho_reset', False):
        for k, v in _defaults.items():
            st.session_state[k] = v
        st.session_state['_ho_reset'] = False

    # ── 2. Definición de filtros ──────────────────────────────────────────────
    # (key_session, columna_df, valor_excluido, valor_todos)
    _defs = [
        ('ho_gerente',   'Gerencia - Cuentas Nacionales', 'Sin Gerente',     'Todos'),
        ('ho_holding',   'Holding',                        'Sin Holding',     'Todos'),
        ('ho_empleador', 'Nombre empleador',               'Sin Empleador',   'Todos'),
        ('ho_protocolo', 'Protocolo',                      'Sin Protocolo',   'Todos'),
        ('ho_region',    'Region Sucursal',                'Sin Región',      'Todas'),
        ('ho_anexo',     'AnexoSUSESO',                    'Sin Información', 'Todos'),
        ('ho_tipo',      'tipo',                           None,              'Todas'),
        ('ho_mes',       'mes_nombre',                     None,              'Todos'),
    ]
    if 'Faena Codelco' in _base.columns:
        _defs.append(('ho_faena_cod', 'Faena Codelco', 'Sin Faena', 'Todos'))
    if 'Faena Marítimo - Portuaria' in _base.columns:
        _defs.append(('ho_maritimo', 'Faena Marítimo - Portuaria', 'Sin Información', 'Todos'))

    # ── 3. Funciones auxiliares ───────────────────────────────────────────────
    def _sin_uno(excluir_key):
        """Retorna _base filtrado por TODOS los filtros excepto el indicado."""
        dff = _base
        for key, col, _, all_val in _defs:
            if key == excluir_key:
                continue
            val = st.session_state.get(key, all_val)
            if val != all_val and col in dff.columns:
                dff = dff[dff[col] == val]
        return dff

    def _opciones(key, col, excl_val):
        dff = _sin_uno(key)
        if col not in dff.columns:
            return []
        vals = dff[col].dropna().unique()
        if excl_val:
            vals = [x for x in vals if x != excl_val]
        return sorted(vals)

    # ── 4. Pase de validación ANTES de renderizar widgets ─────────────────────
    # Si un valor activo ya no aparece en las opciones disponibles (porque
    # otro filtro lo excluyó), se resetea a "Todos" automáticamente.
    # Se itera hasta estabilidad (máx. N veces).
    for _ in range(len(_defs)):
        changed = False
        for key, col, excl_val, all_val in _defs:
            val = st.session_state.get(key, all_val)
            if val == all_val:
                continue
            available = _opciones(key, col, excl_val)
            if val not in available:
                st.session_state[key] = all_val
                changed = True
        if not changed:
            break

    # ── 5. Widgets con key= (session_state ya está limpio y consistente) ──────
    meses_espanol = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

    gerente = st.sidebar.selectbox(
        "Gerencia - Cuentas Nacionales",
        ['Todos'] + _opciones('ho_gerente', 'Gerencia - Cuentas Nacionales', 'Sin Gerente'),
        key='ho_gerente'
    )
    holding = st.sidebar.selectbox(
        "Holding",
        ['Todos'] + _opciones('ho_holding', 'Holding', 'Sin Holding'),
        key='ho_holding'
    )
    nombre_empleador = st.sidebar.selectbox(
        "Nombre empleador",
        ['Todos'] + _opciones('ho_empleador', 'Nombre empleador', 'Sin Empleador'),
        key='ho_empleador'
    )
    protocolo = st.sidebar.selectbox(
        "Protocolo o Programa",
        ['Todos'] + _opciones('ho_protocolo', 'Protocolo', 'Sin Protocolo'),
        key='ho_protocolo'
    )
    region = st.sidebar.selectbox(
        "Región",
        ['Todas'] + _opciones('ho_region', 'Region Sucursal', 'Sin Región'),
        key='ho_region'
    )
    anexo_suseso = st.sidebar.selectbox(
        "Anexo 4 - Protocolos Ministeriales o Anexo 5 No Ministerial",
        ['Todos'] + _opciones('ho_anexo', 'AnexoSUSESO', 'Sin Información'),
        key='ho_anexo'
    )
    tipo = st.sidebar.selectbox(
        "Tipo de Evaluación",
        ['Todas'] + _opciones('ho_tipo', 'tipo', None),
        key='ho_tipo'
    )
    # Mes: orden natural de meses
    meses_disp = [m for m in meses_espanol if m in _opciones('ho_mes', 'mes_nombre', None)]
    mes = st.sidebar.selectbox(
        "Mes",
        ['Todos'] + meses_disp,
        key='ho_mes'
    )

    # Faenas condicionales
    if 'Faena Codelco' in _base.columns:
        faena_codelco = st.sidebar.selectbox(
            "Faena Codelco",
            ['Todos'] + _opciones('ho_faena_cod', 'Faena Codelco', 'Sin Faena'),
            key='ho_faena_cod'
        )
    else:
        faena_codelco = 'Todos'

    if 'Faena Marítimo - Portuaria' in _base.columns:
        maritimo_portuario = st.sidebar.selectbox(
            "Faena Marítimo - Portuaria",
            ['Todos'] + _opciones('ho_maritimo', 'Faena Marítimo - Portuaria', 'Sin Información'),
            key='ho_maritimo'
        )
    else:
        maritimo_portuario = 'Todos'

    # ── 6. Resultado final ────────────────────────────────────────────────────
    df_filtrado = _base.copy()
    for key, col, _, all_val in _defs:
        val = st.session_state.get(key, all_val)
        if val != all_val and col in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado[col] == val]

    # Contador y reseteo
    st.sidebar.markdown("---")
    st.sidebar.caption(f"🔎 **{len(df_filtrado):,}** registros con los filtros actuales")

    if st.sidebar.button("🔄 Resetear Filtros"):
        st.session_state['_ho_reset'] = True
        st.rerun()
    
    # Cargar datos de seguimiento (soft-fail)
    df_seg_raw = cargar_datos_seguimiento()
    
    # Asegurar que df_seg tenga 'Faena Codelco' para que el filtro lateral funcione
    if not df_seg_raw.empty and 'Faena Codelco' not in df_seg_raw.columns:
        # Intentar traer Faena Codelco desde df_eventos usando el ID-CT
        if 'Faena Codelco' in df_eventos.columns:
            mapeo_faena = df_eventos[['Identificador único (ID) centro de trabajo (CT)', 'Faena Codelco']].drop_duplicates().copy()
            df_seg_raw = df_seg_raw.merge(
                mapeo_faena, 
                on='Identificador único (ID) centro de trabajo (CT)', 
                how='left'
            )
    
    # Aplicar los MISMOS filtros laterales a Seguimiento (reutiliza variables de los selectbox)
    df_seg = aplicar_filtros(df_seg_raw, anexo_suseso, protocolo, region, tipo, mes, faena_codelco, gerente, maritimo_portuario, holding, nombre_empleador) if not df_seg_raw.empty else pd.DataFrame()

    # Métricas
    col1, col2, col3, col4 = st.columns(4)
    
    total_evaluaciones = contar_evaluaciones(df_filtrado, protocolo)
    df_cuali = df_filtrado[df_filtrado['tipo'] == 'Cualitativa']
    df_cuanti = df_filtrado[df_filtrado['tipo'] == 'Cuantitativa']
    cuali_count = contar_evaluaciones(df_cuali, protocolo)
    cuanti_count = contar_evaluaciones(df_cuanti, protocolo)
    
    with col1:
        label = "Total Centros de Trabajo" if (protocolo != 'Todos' and es_protocolo_plaguicidas(protocolo)) else "Total Evaluaciones"
        st.metric(label, f"{total_evaluaciones:,}")
    
    with col2:
        st.metric("Cualitativas", f"{cuali_count:,}")
    
    with col3:
        st.metric("Cuantitativas", f"{cuanti_count:,}")
    
    with col4:
        if len(df_filtrado) > 0:
            mes_counts = df_filtrado['mes_nombre'].value_counts()
            if len(mes_counts) > 0:
                mes_max = mes_counts.idxmax()
                cantidad_max = mes_counts.max()
                st.metric("Mes con Mayor Carga", f"{mes_max}", f"{cantidad_max:,} eval.")
            else:
                st.metric("Mes con Mayor Carga", "N/A")
        else:
            st.metric("Mes con Mayor Carga", "N/A")
    
    st.markdown("---")
    
    # Tabs — Programación (1-3) + Seguimiento (4-6)
    tab1, tab2 = st.tabs([
        "📊 Programación y Avance", "🏥 Vigilancia de Salud"
    ])

    with tab1:
        # ── Avance de seguimiento (si está disponible) ─────────────────────
        if not df_seg.empty and 'Estado Cualitativa' in df_seg.columns:
            # Estados que indican que la evaluación NO estaba en la programación oficial:
            # "Realizada fuera de programa" → estaba en el archivo pero sin fecha asignada
            # "Realizada - No programada"   → no estaba en la programación (filas VS(2))
            ESTADOS_FUERA = {'Realizada fuera de programa', 'Realizada - No programada'}

            # Realizadas (incluye todas las variantes de "Realizada*")
            rc_t1 = int(df_seg['Estado Cualitativa'].str.startswith('Realizada', na=False).sum())
            rq_t1 = int(df_seg['Estado Cuantitativa'].str.startswith('Realizada', na=False).sum()) if 'Estado Cuantitativa' in df_seg.columns else 0

            # Conteo interno: evaluaciones realizadas que NO estaban en la programación oficial
            np_cuali_t1  = int(df_seg['Estado Cualitativa'].isin(ESTADOS_FUERA).sum())
            np_cuanti_t1 = int(df_seg['Estado Cuantitativa'].isin(ESTADOS_FUERA).sum()) if 'Estado Cuantitativa' in df_seg.columns else 0
            np_total_t1  = np_cuali_t1 + np_cuanti_t1

            pa_t1 = int((df_seg['Estado Cualitativa'] == 'Pendiente atrasada').sum())
            es_plag_t1 = ('Programa' in df_seg.columns and
                          any('PLAGUICIDAS' in str(p).upper() for p in df_seg['Programa'].unique()))
            if es_plag_t1:
                id_col = 'Identificador único (ID) centro de trabajo (CT)'
                pt_t1 = int(df_seg[id_col].nunique())
                rt_t1 = int(df_seg[df_seg['Estado Cualitativa'].str.startswith('Realizada', na=False)][id_col].nunique())
            else:
                # Denominador: sólo filas con estado programado (excluye fuera de programa y no programadas)
                pc_t1 = int((df_seg['Estado Cualitativa'].notna() &
                             (~df_seg['Estado Cualitativa'].isin(ESTADOS_FUERA))).sum())
                pq_t1 = int((df_seg['Estado Cuantitativa'].notna() &
                             (~df_seg['Estado Cuantitativa'].isin(ESTADOS_FUERA))).sum()) if 'Estado Cuantitativa' in df_seg.columns else 0
                pt_t1 = pc_t1 + pq_t1
                rt_t1 = rc_t1 + rq_t1
            pct_t1  = round(rt_t1 / pt_t1 * 100, 1) if pt_t1 > 0 else 0
            pend_t1 = pt_t1 - rt_t1

            # Pendientes por tipo (sólo disponible fuera del modo Plaguicidas)
            if not es_plag_t1:
                pend_cuali_t1  = pc_t1 - rc_t1
                pend_cuanti_t1 = pq_t1 - rq_t1
                pa_cuanti_t1   = int((df_seg['Estado Cuantitativa'] == 'Pendiente atrasada').sum()) if 'Estado Cuantitativa' in df_seg.columns else 0
            else:
                pend_cuali_t1 = pend_cuanti_t1 = pa_cuanti_t1 = None

            st.markdown("##### ✅ Avance del seguimiento")
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Realizadas",               f"{rt_t1:,}")
            a2.metric("Cualitativas realizadas",  f"{rc_t1:,}",  f"de {cuali_count:,} programadas")
            a3.metric("Cuantitativas realizadas", f"{rq_t1:,}",  f"de {cuanti_count:,} programadas")
            a4.metric("% Avance",                 f"{pct_t1}%")
            st.progress(pct_t1 / 100)

            b1, b2, b3, b4 = st.columns(4)
            b1.metric("Pendientes",               f"{pend_t1:,}",                                          delta_color="inverse")
            b2.metric("Cuali. pendientes",         f"{pend_cuali_t1:,}"  if pend_cuali_t1  is not None else "—", delta_color="inverse")
            b3.metric("Cuanti. pendientes",        f"{pend_cuanti_t1:,}" if pend_cuanti_t1 is not None else "—", delta_color="inverse")
            b4.metric("Atrasadas",                 f"{pa_t1:,}",                                           delta_color="inverse")

            # Conteo interno: evaluaciones realizadas fuera de la programación oficial
            if np_total_t1 > 0:
                partes = []
                if np_cuali_t1  > 0: partes.append(f"{np_cuali_t1:,} cuali.")
                if np_cuanti_t1 > 0: partes.append(f"{np_cuanti_t1:,} cuanti.")
                st.caption(
                    f"ℹ️ Las realizadas incluyen **{np_total_t1:,}** evaluaciones fuera de la "
                    f"programación oficial ({' / '.join(partes)}): "
                    f"sin fecha asignada o no incluidas en el programa."
                )
            st.markdown("---")

        # ── Programación mensual vs realizado ───────────────────────────────
        fig_barras = grafico_programado_vs_realizado(df_filtrado, df_seg if not df_seg.empty else None, protocolo)
        if fig_barras:
            st.plotly_chart(fig_barras, use_container_width=True)
            with st.expander("📋 Ver Detalle de Evaluaciones", expanded=True):
                mostrar_resumen_detallado(df_filtrado, protocolo, seccion='tab1',
                                          df_seg=df_seg if not df_seg.empty else None)
        else:
            st.warning("No hay datos para mostrar con los filtros seleccionados")

    with tab2:
        if df_seg.empty:
            st.info("Sin datos de seguimiento disponibles.")
        else:
            st.markdown("#### Vigilancia de Salud — Evaluados 2026")
            COL_VS = 'Fecha de Evaluación Vigilancia de Salud 2026'
            COL_H  = 'Número de trabajadores evaluados 2026 Hombres'
            COL_M  = 'Número de trabajadores evaluados 2026 Mujeres'
            df_vs = df_seg[df_seg[COL_VS].notna()].copy() if COL_VS in df_seg.columns else pd.DataFrame()
            v1, v2, v3 = st.columns(3)
            v1.metric("Registros con VS", f"{len(df_vs):,}")
            if COL_H in df_vs.columns and not df_vs.empty:
                total_h = pd.to_numeric(df_vs[COL_H], errors='coerce').sum()
                total_m = pd.to_numeric(df_vs[COL_M], errors='coerce').sum() if COL_M in df_vs.columns else 0
                v2.metric("Trabajadores evaluados H", f"{int(total_h):,}")
                v3.metric("Trabajadoras evaluadas M", f"{int(total_m):,}")
            if not df_vs.empty:
                cols_show = [c for c in ['Region Sucursal', 'Nombre Empleador',
                                          'Identificador único (ID) centro de trabajo (CT)',
                                          'Protocolo', 'Agente', # Solo nombres legibles
                                          COL_VS, COL_H, COL_M, 'Observaciones']
                             if c in df_vs.columns]
                df_vs_show = df_vs[cols_show].copy()
                if COL_VS in df_vs_show.columns:
                    df_vs_show[COL_VS] = df_vs_show[COL_VS].dt.strftime('%d-%m-%Y')
                st.dataframe(df_vs_show, use_container_width=True, height=400, hide_index=True)
            else:
                st.info("No hay registros de Vigilancia de Salud en el seguimiento.")

    st.markdown("---")
    st.caption("Versión Producción - Preparado por Diego Vicente Contreras")

except Exception as e:
    st.error(f"❌ Error al cargar o procesar los datos: {str(e)}")
    st.exception(e)
    
    # Información de debug
    with st.expander("🔍 Información de Debug"):
        try:
            url_sheet = st.secrets["gsheets"]["url"]
            st.write("**Fuente de datos:** Google Sheets")
            st.write("**URL configurada:**", url_sheet)
            st.write("**URL de exportación:**", construir_url_exportacion(url_sheet))
        except Exception:
            st.write("No se pudo leer la configuración de secrets.toml")
