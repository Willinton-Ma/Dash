"""
DASHBOARD MI TIENDA - versión dinámica en Streamlit
===================================================

Usa las funciones de dashboard_mitienda.py (lectura de datos, limpieza y gráficos de Seaborn).
Cada vez que se mueve un filtro de la barra lateral, los datos se vuelven a filtrar
y los gráficos de Seaborn se dibujan de nuevo con lo que quedó seleccionado.

Para correrlo en el computador:   streamlit run app.py
"""
import base64

import pandas as pd
import streamlit as st

import dashboard_mitienda as dm

dm.configurar_estilo(1.2)  # letras un poco más grandes porque cada gráfico ocupa media pantalla

st.set_page_config(page_title="Dashboard Mi Tienda", page_icon="🛒", layout="wide")

# Un poco de estilo para que las tarjetas se vean parecidas a las del HTML
st.markdown(
    """
    <style>
    [data-testid="stMetric"]{background:#fff;border-radius:14px;padding:14px 16px;
        border-left:5px solid #0F766E;box-shadow:0 1px 3px rgba(15,23,42,.10)}
    [data-testid="stMetricLabel"] p{color:#64748B;font-size:13px;text-transform:uppercase}
    [data-testid="stMetricValue"]{color:#0F172A}
    .encabezado{background:linear-gradient(120deg,#115E59,#0F766E 60%,#14B8A6);color:#fff;
        padding:22px 28px;border-radius:16px;margin-bottom:18px}
    .encabezado h1{margin:0;font-size:30px;color:#fff;padding:0}
    .encabezado p{margin:6px 0 0;opacity:.92;font-size:15px}
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# Ayudas para que funcione en versiones nuevas y viejas de Streamlit
# ----------------------------------------------------------------------------
def mostrar_imagen(png_base64):
    imagen = base64.b64decode(png_base64)
    try:
        st.image(imagen, width="stretch")
    except TypeError:
        st.image(imagen, use_container_width=True)


def mostrar_tabla(df, **kwargs):
    try:
        st.dataframe(df, width="stretch", hide_index=True, **kwargs)
    except TypeError:
        st.dataframe(df, use_container_width=True, hide_index=True, **kwargs)


# ----------------------------------------------------------------------------
# Lectura de datos (con memoria para que no lea el Excel en cada clic)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo los datos...")
def cargar(origen, url):
    if origen == "Base de datos":
        tablas = dm.cargar_bd(url)
    else:
        tablas = dm.cargar_excel(dm.ARCHIVO_EXCEL)
    return dm.preparar(tablas)


def filtrar(datos, ini, fin, barrios, puestos, tipos):
    """Aplica los filtros de la barra lateral y devuelve (facturas, detalle) ya filtrados."""
    fac, det = datos["fac"], datos["det"]
    f = fac[(fac["fecha"] >= pd.Timestamp(ini)) & (fac["fecha"] < pd.Timestamp(fin) + pd.Timedelta(days=1))]
    if barrios:
        f = f[f["barrio"].isin(barrios)]
    if puestos:
        f = f[f["puesto"].isin(puestos)]
    d = det[det["id_factura"].isin(f["id_factura"])]
    if tipos:
        d = d[d["tipo"].isin(tipos)]
        f = f[f["id_factura"].isin(d["id_factura"])]
    return f, d


@st.cache_data(show_spinner="Dibujando los gráficos con Seaborn...")
def dibujar(origen, url, ini, fin, barrios, puestos, tipos):
    datos = cargar(origen, url)
    f, d = filtrar(datos, ini, fin, barrios, puestos, tipos)
    anios = f["anio"].unique()
    anio = int(anios[0]) if len(anios) == 1 else None
    graficos = dm.graficos_del_periodo(anio, f, d, ultima_fecha=datos["fac"]["fecha"].max())
    graficos.update(dm.graficos_estaticos(datos["emp"]))
    return graficos


# ----------------------------------------------------------------------------
# Barra lateral: de dónde vienen los datos y los filtros
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("Datos")
    origen = st.radio("Fuente", ["Excel", "Base de datos"], horizontal=True,
                      help="Excel usa el archivo Mi_Tienda.xlsx. Base de datos lee las tablas directamente "
                           "de MySQL, PostgreSQL o SQL Server.")
    url = ""
    if origen == "Base de datos":
        try:
            por_defecto = st.secrets.get("BD_URL", "")
        except Exception:  # no hay archivo de secretos
            por_defecto = ""
        url = st.text_input("Cadena de conexión", value=por_defecto, type="password",
                            placeholder="postgresql+psycopg2://usuario:clave@servidor:5432/mitienda")

try:
    if origen == "Base de datos" and not url:
        st.info("Escribe la cadena de conexión en la barra lateral para leer de la base de datos.")
        st.stop()
    datos = cargar(origen, url)
except Exception as error:  # si falla la conexión se avisa y no se cae la página
    st.error(f"No se pudieron leer los datos: {error}")
    st.stop()

fac, det = datos["fac"], datos["det"]
fecha_min, fecha_max = fac["fecha"].min().date(), fac["fecha"].max().date()
barrios_todos = sorted(fac["barrio"].dropna().unique())
puestos_todos = sorted(fac["puesto"].dropna().unique())
tipos_todos = sorted(det["tipo"].dropna().unique())


def restablecer():
    st.session_state["rango"] = (fecha_min, fecha_max)
    st.session_state["barrios"] = []
    st.session_state["puestos"] = []
    st.session_state["tipos"] = []


with st.sidebar:
    st.header("Filtros")
    rango = st.date_input("Rango de fechas", value=(fecha_min, fecha_max), min_value=fecha_min,
                          max_value=fecha_max, key="rango")
    barrios = st.multiselect("Barrio del cliente", barrios_todos, key="barrios", placeholder="Todos")
    puestos = st.multiselect("Puesto del empleado", puestos_todos, key="puestos", placeholder="Todos")
    tipos = st.multiselect("Tipo de producto", tipos_todos, key="tipos", placeholder="Todos")
    st.button("Restablecer filtros", on_click=restablecer)
    st.caption("Si no eliges nada en un filtro, se muestran todos.")

# el selector de fechas devuelve un solo valor mientras el usuario no ha elegido la fecha final
if isinstance(rango, (tuple, list)) and len(rango) == 2:
    ini, fin = rango
else:
    st.info("Elige también la fecha final del rango.")
    st.stop()

f, d = filtrar(datos, ini, fin, barrios, puestos, tipos)

# ----------------------------------------------------------------------------
# Encabezado e indicadores
# ----------------------------------------------------------------------------
st.markdown(
    f'<div class="encabezado"><h1>🛒 Dashboard Mi Tienda</h1>'
    f"<p>Ventas, productos, clientes y equipo · Datos desde: "
    f'{"base de datos" if origen == "Base de datos" else "archivo Excel"} · '
    f"Del {ini:%d/%m/%Y} al {fin:%d/%m/%Y}</p></div>",
    unsafe_allow_html=True,
)

if len(f) < 5:
    st.warning("Con estos filtros quedan muy pocas facturas para dibujar los gráficos. "
               "Amplía el rango de fechas o quita algún filtro.")
    st.stop()

# comparación con el mismo periodo del año anterior (si hay datos de ese año)
delta = None
ini_prev, fin_prev = pd.Timestamp(ini) - pd.DateOffset(years=1), pd.Timestamp(fin) - pd.DateOffset(years=1)
if ini_prev.date() >= fecha_min:
    f_prev, _ = filtrar(datos, ini_prev.date(), fin_prev.date(), barrios, puestos, tipos)
    if f_prev["total"].sum() > 0:
        delta = (f["total"].sum() - f_prev["total"].sum()) / f_prev["total"].sum() * 100

mejor_mes = f.groupby(f["fecha"].dt.to_period("M"))["total"].sum().idxmax()
mejor_dia = f.groupby("dow")["total"].sum().idxmax()

fila1 = st.columns(4)
fila1[0].metric("Ventas totales", "$ " + dm.a_millones(f["total"].sum()),
                delta=None if delta is None else f"{delta:+.1f} %".replace(".", ","),
                help="La flecha compara con el mismo periodo del año anterior.")
fila1[1].metric("Facturas", dm.entero(len(f)))
fila1[2].metric("Ticket promedio", dm.pesos(f["total"].mean()))
fila1[3].metric("Clientes que compraron", dm.entero(f["id_cliente"].nunique()))
fila2 = st.columns(4)
fila2[0].metric("Unidades vendidas", dm.entero(d["cantidad"].sum()))
fila2[1].metric("Mejor mes", f"{dm.MESES[mejor_mes.month - 1]} {mejor_mes.year}")
fila2[2].metric("Mejor día", dm.DIAS[mejor_dia])
fila2[3].metric("Productos agotados", datos["avisos"]["agotados"],
                help="Productos con stock en cero (no depende de los filtros).")

# ----------------------------------------------------------------------------
# Gráficos por secciones
# ----------------------------------------------------------------------------
graficos = dibujar(origen, url, ini, fin, tuple(barrios), tuple(puestos), tuple(tipos))

pestanas = st.tabs([titulo for _, titulo, _ in dm.SECCIONES] + ["Facturas", "Notas"])

for pestana, (_, _, lista) in zip(pestanas, dm.SECCIONES):
    with pestana:
        columnas = st.columns(2)
        for i, (clave, titulo, explicacion) in enumerate(lista):
            with columnas[i % 2].container(border=True):
                st.markdown(f"#### {titulo}")
                st.caption(explicacion)
                mostrar_imagen(graficos[clave])

with pestanas[-2]:
    st.subheader("Facturas del periodo")
    buscar = st.text_input("Buscar por cliente, empleado, barrio o puesto", placeholder="Ejemplo: Aranjuez")
    tabla = f.sort_values("fecha", ascending=False)[
        ["id_factura", "fecha", "cliente_et", "empleado_et", "barrio", "puesto", "total"]]
    tabla.columns = ["Factura", "Fecha", "Cliente", "Empleado", "Barrio del cliente", "Puesto", "Total"]
    if buscar:
        mascara = tabla.drop(columns=["Fecha", "Total"]).astype(str).apply(
            lambda col: col.str.contains(buscar, case=False, regex=False)).any(axis=1)
        tabla = tabla[mascara]
    st.caption(f"{len(tabla)} facturas")
    mostrar_tabla(tabla, column_config={
        "Fecha": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
        "Total": st.column_config.NumberColumn(format="$ %d"),
    })
    st.download_button("Descargar estas facturas (CSV)", tabla.to_csv(index=False).encode("utf-8-sig"),
                       file_name="facturas_filtradas.csv", mime="text/csv")

with pestanas[-1]:
    av = datos["avisos"]
    st.subheader("Notas sobre los datos")
    if av["ids_arreglados"]:
        st.write("- El Excel traía los IDs repetidos (cada hoja estaba pegada dos veces y los números empezaban de nuevo). "
                 "Se separó cada bloque como un lote distinto; con eso el total de cada factura coincide con "
                 f"cantidad × precio de su pedido en el {av['coinciden']:.0f} % de los casos.")
    st.write(f"- Cada una de las {av['facturas']} facturas tiene un pedido con un producto; las gráficas de productos, "
             "tipos y proveedores salen de esos pedidos.")
    st.write(f"- {av['sin_nombre']} productos no tenían nombre en la base de datos; se muestran como \"Producto sin nombre\".")
    st.write("- El último mes con datos puede estar incompleto; por eso no se dibuja en la gráfica de ventas por mes.")
    st.write("- Los salarios no cambian con los filtros de ventas.")
