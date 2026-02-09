# Dashboard Programación de Evaluaciones IST 2026

Dashboard interactivo para la visualización y seguimiento de la programación de evaluaciones de Higiene Ocupacional 2026 del Instituto de Seguridad del Trabajo (IST).

## Características

- **Fuente de datos en tiempo real:** Lee directamente desde Google Sheets (actualización cada 5 minutos)
- **Motor de consulta en memoria:** Usa DuckDB para carga y procesamiento rápido
- **Filtros interactivos:** Anexo SUSESO, Protocolo, Región, Tipo de Evaluación, Mes y Faena Codelco
- **Visualizaciones:** Carga mensual por tipo, Top 10 Protocolos, detalle completo con exportación CSV
- **Lógica especial para Plaguicidas:** Conteo por Centro de Trabajo único en vez de evaluaciones individuales

## Requisitos

- Python 3.9+
- Entorno conda recomendado (`dash`)

## Instalación

```bash
conda activate dash
pip install -r requirements.txt
```

## Configuración

Crear el archivo `.streamlit/secrets.toml` con la URL de la hoja de Google Sheets:

```toml
[gsheets]
url = "https://docs.google.com/spreadsheets/d/TU_SPREADSHEET_ID/edit?gid=TU_GID#gid=TU_GID"
```

> **Importante:** La hoja debe tener acceso "Cualquiera con el enlace puede ver".

## Ejecución local

```bash
streamlit run dashboard_programacion_2026_PRODUCCION.py
```

## Deploy en Streamlit Cloud

1. Conectar el repositorio desde [share.streamlit.io](https://share.streamlit.io)
2. Configurar los secrets en **Settings > Secrets** con el contenido de `secrets.toml`
3. El archivo `secrets.toml` **nunca** debe subirse al repositorio

## Estructura del proyecto

```
DashboardPlan2026/
├── dashboard_programacion_2026_PRODUCCION.py   # Aplicación principal
├── requirements.txt                             # Dependencias
├── .streamlit/
│   └── secrets.toml                             # URL Google Sheets (local, no en git)
├── .gitignore
└── README.md
```

## Stack tecnológico

| Componente | Tecnología |
|---|---|
| Frontend | Streamlit 1.31.0 |
| Datos | Google Sheets (CSV export) |
| Motor en memoria | DuckDB 0.10.0 |
| Procesamiento | Pandas 2.1.4 |
| Visualización | Plotly 5.18.0 |

## Autor

Diego Vicente Contreras - Equipo de Especialidades Técnicas, IST
