"""
App de Streamlit para mostrar el dashboard de Mi Tienda.

Esta página no dibuja nada por sí sola: usa el script dashboard_mitienda.py.
- Si ya existe dashboard_mitienda.html en el repositorio, lo muestra directo (es lo más rápido).
- Si no existe, lo arma al abrir la página con el archivo Mi_Tienda.xlsx.
"""
import os

import streamlit as st

import dashboard_mitienda as dm

st.set_page_config(page_title="Dashboard Mi Tienda", layout="wide")

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_HTML = os.path.join(CARPETA, "dashboard_mitienda.html")
ARCHIVO_EXCEL = os.path.join(CARPETA, "Mi_Tienda.xlsx")


@st.cache_data(show_spinner="Armando el dashboard (puede tardar un minuto)...")
def cargar_pagina():
    if os.path.exists(ARCHIVO_HTML):
        with open(ARCHIVO_HTML, encoding="utf-8") as f:
            return f.read()
    datos = dm.preparar(dm.cargar_excel(ARCHIVO_EXCEL))
    return dm.construir_html(datos, "archivo Excel")


pagina = cargar_pagina()
if hasattr(st, "iframe"):          # versiones nuevas de Streamlit
    st.iframe(pagina, height=1000)
else:                              # versiones viejas
    import streamlit.components.v1 as components
    components.html(pagina, height=1000, scrolling=True)
