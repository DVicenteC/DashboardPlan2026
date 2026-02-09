# Dashboard Programación de Evaluaciones 2026

Dashboard interactivo para la visualización y seguimiento de la programación de evaluaciones

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

## ConfiguraciónEjecución local

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

| Componente       | Tecnología                |
| ---------------- | -------------------------- |
| Frontend         | Streamlit 1.31.0           |
| Datos            | Google Sheets (CSV export) |
| Motor en memoria | DuckDB 0.10.0              |
| Procesamiento    | Pandas 2.1.4               |
| Visualización   | Plotly 5.18.0              |

## Autor

Diego Vicente Contreras - Equipo de Especialidades Técnicas, IST
