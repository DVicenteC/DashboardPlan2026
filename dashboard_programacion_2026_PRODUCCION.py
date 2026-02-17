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
            st.session_state["username"] == "70015580-3"
            and st.session_state["password"] == "IST"
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
    Parsea una serie de fechas que puede tener formatos mixtos:
    - DD-MM-YYYY (formato Excel original)
    - M/D/YYYY o MM/DD/YYYY (formato Google Sheets export)
    - YYYY-MM-DD (formato ISO)
    """
    # Intentar formato DD-MM-YYYY primero
    resultado = pd.to_datetime(serie, format='%d-%m-%Y', errors='coerce')

    # Para las que fallaron, intentar formato M/D/YYYY (US)
    mascara_nulos = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara_nulos.any():
        resultado[mascara_nulos] = pd.to_datetime(serie[mascara_nulos], format='%m/%d/%Y', errors='coerce')

    # Para las que aún fallaron, intentar formato mixto automático
    mascara_nulos2 = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara_nulos2.any():
        resultado[mascara_nulos2] = pd.to_datetime(serie[mascara_nulos2], dayfirst=True, errors='coerce')

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
                     'Gerencia - Cuentas Nacionales', 'Faena Marítimo - Portuaria']

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

def aplicar_filtros(df, anexo_suseso, protocolo, region, tipo, mes, faena_codelco, gerente, maritimo_portuario):
    """Aplica los filtros seleccionados - VERSIÓN CORREGIDA"""
    df_filtrado = df.copy()
    
    # FIX CRÍTICO: Usar .copy() después de cada filtro
    if anexo_suseso != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['AnexoSUSESO'] == anexo_suseso].copy()
    
    if gerente != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Gerencia - Cuentas Nacionales'] == gerente].copy()
    
    if maritimo_portuario != 'Todos' and 'Faena Marítimo - Portuaria' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['Faena Marítimo - Portuaria'] == maritimo_portuario].copy()

    if protocolo != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Protocolo'] == protocolo].copy()
    
    if region != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Region Sucursal'] == region].copy()
    
    if tipo != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['tipo'] == tipo].copy()
    
    if mes != 'Todos':
        meses_es_a_num = {
            'Enero': 1, 'Febrero': 2, 'Marzo': 3, 'Abril': 4,
            'Mayo': 5, 'Junio': 6, 'Julio': 7, 'Agosto': 8,
            'Septiembre': 9, 'Octubre': 10, 'Noviembre': 11, 'Diciembre': 12
        }
        mes_num = meses_es_a_num[mes]
        df_filtrado = df_filtrado[df_filtrado['mes'] == mes_num].copy()
    
    if faena_codelco != 'Todos' and 'Faena Codelco' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['Faena Codelco'] == faena_codelco].copy()
    
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

def mostrar_resumen_detallado(df_filtrado, protocolo_seleccionado, seccion='tab1'):
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
        
        df_agrupado['fecha'] = pd.to_datetime(df_agrupado['fecha']).dt.strftime('%d-%m-%Y')
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
        
        columnas_detalle = ['fecha', 'tipo', 'Rut Empleador o Rut trabajador(a)', 'Nombre empleador','Identificador único (ID) centro de trabajo (CT)', 'NOMBRE SUCURSAL', 'Agente', 
                            'Protocolo', 'Region Sucursal', 'Comuna CT', 'Nivel de riesgo', 'AnexoSUSESO', 'Gerencia - Cuentas Nacionales', 'Faena Marítimo - Portuaria']
        
        nombres_columnas = ['Fecha', 'Tipo', 'Rut Empleador o Rut trabajador(a)', 'Nombre empleador','Identificador único (ID) centro de trabajo (CT)', 'Sucursal', 'Agente', 
                            'Protocolo', 'Región', 'Comuna', 'Nivel de Riesgo', 'Anexo SUSESO', 'Gerente', 'Marítimo Portuario']
        
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

st.title("📊 Dashboard de Programación de Evaluaciones 2026")
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
    
    # Sidebar - Filtros
    st.sidebar.header("🔍 Filtros")
    
    # Excluir valores 'Sin Información'
    anexos_unicos = sorted([x for x in df_eventos['AnexoSUSESO'].unique() if x != 'Sin Información'])
    anexo_suseso = st.sidebar.selectbox(
        "Anexo 4 - Protocolos Ministeriales o Anexo 5 No Ministerial",
        ['Todos'] + anexos_unicos
    )
    
    protocolos_unicos = sorted([x for x in df_eventos['Protocolo'].unique() if x != 'Sin Protocolo'])
    protocolo = st.sidebar.selectbox(
        "Protocolo o Programa",
        ['Todos'] + protocolos_unicos
    )
    
    regiones_unicas = sorted([x for x in df_eventos['Region Sucursal'].unique() if x != 'Sin Región'])
    region = st.sidebar.selectbox(
        "Región",
        ['Todas'] + regiones_unicas
    )

    gerentes_unicos = sorted([x for x in df_eventos['Gerencia - Cuentas Nacionales'].unique() if x != 'Sin Gerente'])
    gerente = st.sidebar.selectbox(
        "Gerencia - Cuentas Nacionales",
        ['Todos'] + gerentes_unicos
    )
    
    tipo = st.sidebar.selectbox(
        "Tipo de Evaluación",
        ['Todas', 'Cualitativa', 'Cuantitativa']
    )
    
    meses_espanol = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    mes = st.sidebar.selectbox(
        "Mes",
        ['Todos'] + meses_espanol
    )
    
    if 'Faena Codelco' in df_eventos.columns:
        faenas_unicas = sorted([x for x in df_eventos['Faena Codelco'].unique() if x != 'Sin Faena'])
        faena_codelco = st.sidebar.selectbox(
            "Faena Codelco",
            ['Todos'] + faenas_unicas
        )
    else:
        faena_codelco = 'Todos'
    
    if 'Faena Marítimo - Portuaria' in df_eventos.columns:
        maritimo_portuario_unicos = sorted([x for x in df_eventos['Faena Marítimo - Portuaria'].unique() if x != 'Sin Información'])
        maritimo_portuario = st.sidebar.selectbox(
            "Faena Marítimo - Portuaria",
            ['Todos'] + maritimo_portuario_unicos
        )
    else:
        maritimo_portuario = 'Todos'
    
    if st.sidebar.button("🔄 Resetear Filtros"):
        st.rerun()
    
    # Aplicar filtros
    df_filtrado = aplicar_filtros(df_eventos, anexo_suseso, protocolo, region, tipo, mes, faena_codelco, gerente, maritimo_portuario)
    
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
    
    # Tabs
    tab1, tab2, tab3 = st.tabs(["📊 Resumen Mensual", "📈 Top Protocolos", "🔍 Detalle Completo"])
    
    with tab1:
        fig_barras = grafico_barras_mensuales(df_filtrado, protocolo)
        if fig_barras:
            st.plotly_chart(fig_barras, use_container_width=True)
            with st.expander("📋 Ver Detalle de Evaluaciones", expanded=True):
                mostrar_resumen_detallado(df_filtrado, protocolo, seccion='tab1')
        else:
            st.warning("No hay datos para mostrar con los filtros seleccionados")
    
    with tab2:
        fig_protocolos = grafico_top_protocolos(df_filtrado)
        if fig_protocolos:
            st.plotly_chart(fig_protocolos, use_container_width=True)
            with st.expander("📋 Ver Detalle de Evaluaciones", expanded=False):
                mostrar_resumen_detallado(df_filtrado, protocolo, seccion='tab2')
        else:
            st.warning("No hay datos para mostrar con los filtros seleccionados")
    
    with tab3:
        mostrar_resumen_detallado(df_filtrado, protocolo, seccion='tab3')
    
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
