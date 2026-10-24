"""
DASHBOARD MI TIENDA  (HTML + Python + Seaborn)
==============================================

Extensión de la tarea del dólar: ahora con la base de datos de Mi Tienda
se arma un dashboard parecido al de Power BI, pero hecho con Python.

Cómo funciona (resumen):
  1. Se leen las tablas (del Excel o directamente de una base de datos).
  2. Se limpian y se unen (facturas + clientes + empleados + pedidos + productos...).
  3. Para cada filtro de año (Todos, 2024, 2025, 2026) se dibujan los gráficos con Seaborn.
  4. Todo se mete en UN solo archivo HTML (sin internet, sin servidor) con botones
     para cambiar el año, buscador de facturas y gráficos que se pueden ampliar.

Instalación de librerías:
  pip install pandas openpyxl seaborn matplotlib sqlalchemy
  # y, según la base de datos que se vaya a usar:
  pip install pymysql          # MySQL
  pip install psycopg2-binary  # PostgreSQL
  pip install pyodbc           # SQL Server

Formas de ejecutarlo:
  python dashboard_mitienda.py                      -> lee el Excel y genera el HTML
  python dashboard_mitienda.py --cargar-bd          -> sube las 6 hojas del Excel a la base de datos
  python dashboard_mitienda.py --origen bd          -> lee TODO desde la base de datos y genera el HTML
  python dashboard_mitienda.py --origen bd --bd-url "postgresql+psycopg2://usuario:clave@localhost:5432/mitienda"
"""

import argparse
import base64
import html
import io
import os
import re
from datetime import datetime

import matplotlib

matplotlib.use("Agg")  # para dibujar sin abrir ventanas
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib.ticker import FuncFormatter, MaxNLocator
from sqlalchemy import create_engine, text

# ----------------------------------------------------------------------------
# 1. CONFIGURACIÓN (lo único que hay que tocar para cambiar de base de datos)
# ----------------------------------------------------------------------------
CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(CARPETA, "Mi_Tienda.xlsx")
ARCHIVO_HTML = os.path.join(CARPETA, "dashboard_mitienda.html")

# Cadena de conexión. Se puede cambiar aquí o con la variable de entorno MITIENDA_BD_URL.
# Ejemplos (descomentar el que se use):
#   MySQL:       "mysql+pymysql://root:clave@localhost:3306/mitienda"
#   PostgreSQL:  "postgresql+psycopg2://postgres:clave@localhost:5432/mitienda"
#   SQL Server:  "mssql+pyodbc://usuario:clave@localhost/mitienda?driver=ODBC+Driver+17+for+SQL+Server"
#   SQLite (solo para probar sin instalar nada): "sqlite:///mitienda.db"
BD_URL = os.getenv("MITIENDA_BD_URL", "mysql+pymysql://root:@localhost:3306/mitienda")

HOJAS = ["Clientes", "Empleados", "Facturas", "Pedidos", "Productos", "Proveedores"]

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
DIAS_CORTO = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

# Colores del dashboard
BASE = "#0F766E"      # verde azulado
ACENTO = "#F59E0B"    # naranja para resaltar lo más alto
GRIS = "#94A3B8"
PALETA_ANIOS = {2024: GRIS, 2025: BASE, 2026: ACENTO}

sns.set_theme(
    style="whitegrid",
    font_scale=0.85,
    rc={"grid.color": "#E5E7EB", "axes.edgecolor": "#CBD5E1", "axes.labelcolor": "#334155",
        "xtick.color": "#475569", "ytick.color": "#475569"},
)


# ----------------------------------------------------------------------------
# 2. LECTURA DE DATOS (Excel o base de datos)
# ----------------------------------------------------------------------------
def cargar_excel(ruta):
    """Lee las 6 hojas del Excel y las devuelve en un diccionario."""
    datos = pd.read_excel(ruta, sheet_name=None)
    return {h: datos[h] for h in HOJAS}


def cargar_bd(url):
    """Lee las 6 tablas directamente desde la base de datos (MySQL, PostgreSQL o SQL Server)."""
    motor = create_engine(url)
    tablas = {}
    with motor.connect() as conexion:
        for hoja in HOJAS:
            tablas[hoja] = pd.read_sql(text(f"SELECT * FROM {hoja.lower()}"), conexion)
    motor.dispose()
    return tablas


def subir_excel_a_bd(ruta, url):
    """Toma las hojas del Excel y las guarda como tablas en la base de datos."""
    motor = create_engine(url)
    datos = arreglar_ids(cargar_excel(ruta))
    for hoja in HOJAS:
        tabla = datos[hoja]
        tabla.to_sql(hoja.lower(), motor, if_exists="replace", index=False)
        print(f"  tabla '{hoja.lower()}' cargada: {len(tabla)} filas")
    motor.dispose()


# ----------------------------------------------------------------------------
# 3. LIMPIEZA Y UNIÓN DE TABLAS
# ----------------------------------------------------------------------------
def corto(texto, n=30):
    return texto if len(texto) <= n else texto[: n - 1] + "…"


def arreglar_ids(t):
    """
    El Excel trae cada hoja "pegada dos veces": los IDs vuelven a empezar en 1 en la segunda mitad
    (por ejemplo, hay dos clientes con id 1 y dos facturas con id 1). Si se unen las tablas así,
    cada venta se duplica. Aquí se separa cada bloque como un "lote" y se le suma un salto a los IDs
    para que queden únicos (1 a 1600 en facturas, 1 a 1000 en clientes, etc.).
    Si los IDs ya son únicos (por ejemplo, datos que vienen de la base de datos) no se cambia nada.
    """
    t = {k: v.copy() for k, v in t.items()}
    saltos = {}

    # dimensiones: clientes, empleados y productos
    for hoja, col in [("Clientes", "id_cliente"), ("Empleados", "id_empleado"), ("Productos", "id_producto")]:
        d = t[hoja]
        lote = d.groupby(col).cumcount()
        saltos[col] = int(d[col].max()) if (lote > 0).any() else 0
        d[col] = d[col] + lote * saltos[col]

    # facturas y pedidos: cada pedido se queda con la factura de su mismo lote
    fac, ped = t["Facturas"], t["Pedidos"]
    lote_f = fac.groupby("id_factura").cumcount()
    lote_p = ped.groupby("id_factura").cumcount()
    salto_fac = int(fac["id_factura"].max()) if (lote_f > 0).any() else 0
    fac["id_cliente"] = fac["id_cliente"] + lote_f * saltos["id_cliente"]
    fac["id_empleado"] = fac["id_empleado"] + lote_f * saltos["id_empleado"]
    fac["id_factura"] = fac["id_factura"] + lote_f * salto_fac
    ped["id_producto"] = ped["id_producto"] + lote_p * saltos["id_producto"]
    ped["id_factura"] = ped["id_factura"] + lote_p * salto_fac
    if (ped["id_pedido"].duplicated()).any():
        ped["id_pedido"] = ped["id_pedido"] + ped.groupby("id_pedido").cumcount() * int(ped["id_pedido"].max())
    t["_ids_arreglados"] = bool((lote_f > 0).any())
    return t


def preparar(t):
    """Limpia las tablas y arma dos tablas grandes: facturas (fac) y detalle de pedidos (det)."""
    t = arreglar_ids(t)
    cli, emp, fac, ped, pro, prv = (t[h].copy() for h in HOJAS)
    avisos = {"ids_arreglados": t["_ids_arreglados"]}

    # --- clientes: nombre completo y barrio (viene dentro de la dirección)
    cli["cliente"] = cli["nombre"] + " " + cli["apellido"].str.split().str[0]
    cli["cliente_et"] = cli["cliente"] + " (#" + cli["id_cliente"].astype(str) + ")"
    cli["barrio"] = cli["direccion"].str.split(",").str[-2].str.strip()

    # --- empleados
    emp["empleado"] = emp["nombre"] + " " + emp["apellido"]
    emp["empleado_et"] = emp["empleado"] + " (#" + emp["id_empleado"].astype(str) + ")"
    emp["fecha_contratacion"] = pd.to_datetime(emp["fecha_contratacion"])

    # --- productos: hay algunos sin nombre, se les pone uno para no perderlos
    avisos["sin_nombre"] = int(pro["nombre"].isna().sum())
    sin = pro["nombre"].isna()
    pro.loc[sin, "nombre"] = "Producto sin nombre (Lote #" + pro.loc[sin, "id_producto"].astype(str) + ")"
    pro["limpio"] = pro["nombre"].str.replace(r"\s*\(Lote #\d+\)", "", regex=True)
    pro["tipo"] = pro["limpio"].str.split().str[0]
    pro.loc[sin, "tipo"] = "Sin nombre"
    pro["producto"] = pro["limpio"].map(lambda n: corto(n, 22)) + " (#" + pro["id_producto"].astype(str) + ")"

    # --- facturas
    fac["fecha"] = pd.to_datetime(fac["fecha"])
    fac["anio"] = fac["fecha"].dt.year
    fac["mes"] = fac["fecha"].dt.month
    fac["dow"] = fac["fecha"].dt.dayofweek
    fac["hora"] = fac["fecha"].dt.hour
    fac = fac.merge(cli[["id_cliente", "cliente_et", "barrio"]], on="id_cliente", how="left")
    fac = fac.merge(emp[["id_empleado", "empleado_et", "puesto"]], on="id_empleado", how="left")

    # --- detalle: pedidos + fecha de la factura + producto + proveedor
    ped["ingreso"] = ped["cantidad"] * ped["precio_unitario"]
    det = ped.merge(fac[["id_factura", "fecha", "anio"]], on="id_factura", how="inner")
    det = det.merge(pro[["id_producto", "producto", "tipo", "id_proveedor"]], on="id_producto", how="left")
    det = det.merge(prv[["id_proveedor", "nombre"]].rename(columns={"nombre": "proveedor"}),
                    on="id_proveedor", how="left")

    # --- datos para los avisos del final de la página
    avisos["facturas"] = len(fac)
    avisos["facturas_con_detalle"] = int(ped["id_factura"].nunique())
    suma = ped.groupby("id_factura")["ingreso"].sum()
    comparar = fac.set_index("id_factura")["total"].reindex(suma.index)
    avisos["coinciden"] = float((suma == comparar).mean() * 100)
    avisos["agotados"] = int((pro["stock"] == 0).sum())

    return {"cli": cli, "emp": emp, "fac": fac, "det": det, "pro": pro, "avisos": avisos}


# ----------------------------------------------------------------------------
# 4. FORMATOS Y AYUDAS PARA LOS GRÁFICOS
# ----------------------------------------------------------------------------
def pesos(v):
    return "$ " + f"{v:,.0f}".replace(",", ".")


def a_millones(v, pos=None):
    return f"{v / 1e6:.1f}".replace(".", ",") + " M"


def a_miles(v, pos=None):
    return f"{v / 1e3:,.0f}".replace(",", ".") + " mil"


def entero(v, pos=None):
    return f"{v:,.0f}".replace(",", ".")


def formato_auto(maximo):
    return a_millones if maximo >= 1e6 else a_miles


def a_imagen(fig):
    """Convierte una figura de matplotlib en texto base64 para pegarla en el HTML."""
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=110, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode("ascii")


def barras_horizontales(tabla, x, y, fmt=None):
    """Gráfico de barras horizontales con el valor al lado y la barra más grande resaltada."""
    fmt = fmt or formato_auto(tabla[x].max())
    fig, ax = plt.subplots(figsize=(6.4, 0.36 * len(tabla) + 1.0))
    sns.barplot(data=tabla, x=x, y=y, orient="h", color=BASE, ax=ax)
    ax.containers[0][0].set_color(ACENTO)
    ax.bar_label(ax.containers[0], labels=[fmt(v) for v in tabla[x]], padding=3, fontsize=8, color="#334155")
    ax.xaxis.set_major_locator(MaxNLocator(5))
    ax.xaxis.set_major_formatter(FuncFormatter(fmt))
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.margins(x=0.16)
    sns.despine(ax=ax, left=True, bottom=True)
    return a_imagen(fig)


def top(df, columna_grupo, columna_valor, n=10, funcion="sum"):
    resultado = df.groupby(columna_grupo)[columna_valor].agg(funcion).nlargest(n)
    return resultado.reset_index()


# ----------------------------------------------------------------------------
# 5. LOS GRÁFICOS (cada filtro de año genera su propio juego)
# ----------------------------------------------------------------------------
def graficos_estaticos(emp):
    """Gráficos que no dependen del año (los salarios no cambian con las ventas)."""
    np.random.seed(1)  # para que los puntos dispersos salgan siempre igual
    fig, ax = plt.subplots(figsize=(6.4, 4.4))
    orden = emp.groupby("puesto")["salario"].median().sort_values(ascending=False).index
    sns.boxplot(data=emp, x="salario", y="puesto", order=orden, color="#99D5CF",
                fliersize=0, linewidth=1, ax=ax)
    sns.stripplot(data=emp, x="salario", y="puesto", order=orden, color="#134E4A",
                  size=3, alpha=0.6, ax=ax)
    ax.xaxis.set_major_formatter(FuncFormatter(a_millones))
    ax.set_xlabel("Salario mensual")
    ax.set_ylabel("")
    sns.despine(ax=ax, left=True)
    return {"salarios": a_imagen(fig)}


def graficos_del_periodo(anio, fac, det):
    """Dibuja todos los gráficos para un año (o para todos si anio es None)."""
    f = fac if anio is None else fac[fac["anio"] == anio]
    d = det if anio is None else det[det["anio"] == anio]
    g = {}

    # 1) Ventas por mes ------------------------------------------------------
    mensual = f.groupby(["anio", "mes"], as_index=False)["total"].sum()
    ultima = fac["fecha"].max()  # si el último mes quedó a medias, no se dibuja (se vería como una caída falsa)
    if ultima.day < ultima.days_in_month:
        mensual = mensual[~((mensual["anio"] == ultima.year) & (mensual["mes"] == ultima.month))]
    fig, ax = plt.subplots(figsize=(6.4, 3.4))
    if anio is None:
        sns.lineplot(data=mensual, x="mes", y="total", hue="anio", palette=PALETA_ANIOS,
                     marker="o", linewidth=2, ax=ax)
        ax.legend(title="Año", frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.1))
    else:
        sns.lineplot(data=mensual, x="mes", y="total", color=BASE, marker="o", linewidth=2.2, ax=ax)
        mejor = mensual.loc[mensual["total"].idxmax()]
        ax.scatter([mejor["mes"]], [mejor["total"]], color=ACENTO, s=90, zorder=5)
    ax.set_xticks(range(1, 13))
    ax.set_xticklabels(MESES)
    ax.set_ylim(bottom=0)  # que el eje arranque en cero para no exagerar las subidas y bajadas
    ax.yaxis.set_major_formatter(FuncFormatter(a_millones))
    ax.set_xlabel("")
    ax.set_ylabel("")
    sns.despine(ax=ax)
    g["mes"] = a_imagen(fig)

    # 2) Ventas acumuladas ---------------------------------------------------
    acum = f.groupby(f["fecha"].dt.normalize())["total"].sum().sort_index().cumsum().reset_index()
    fig, ax = plt.subplots(figsize=(6.4, 3.4))
    sns.lineplot(data=acum, x="fecha", y="total", color=BASE, linewidth=2, ax=ax)
    ax.fill_between(acum["fecha"], acum["total"], alpha=0.15, color=BASE)
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3 if anio is None else 1))
    if anio is None:
        ax.xaxis.set_major_formatter(FuncFormatter(
            lambda v, pos: f"{MESES[mdates.num2date(v).month - 1]} {mdates.num2date(v).year % 100}"))
    else:
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, pos: MESES[mdates.num2date(v).month - 1]))
    ax.tick_params(axis="x", labelsize=8)
    ax.yaxis.set_major_formatter(FuncFormatter(formato_auto(acum["total"].max())))
    ax.set_xlabel("")
    ax.set_ylabel("")
    sns.despine(ax=ax)
    g["acumulado"] = a_imagen(fig)

    # 3) Ventas por día de la semana ----------------------------------------
    por_dia = f.groupby("dow", as_index=False)["total"].sum()
    por_dia["dia"] = por_dia["dow"].map(lambda i: DIAS[i])
    fig, ax = plt.subplots(figsize=(6.4, 3.4))
    sns.barplot(data=por_dia, x="dia", y="total", order=DIAS, color=BASE, ax=ax)
    ax.containers[0][int(por_dia.loc[por_dia["total"].idxmax(), "dow"])].set_color(ACENTO)
    ax.yaxis.set_major_formatter(FuncFormatter(formato_auto(por_dia["total"].max())))
    ax.tick_params(axis="x", labelsize=8)
    ax.set_xlabel("")
    ax.set_ylabel("")
    sns.despine(ax=ax)
    g["dia"] = a_imagen(fig)

    # 4) Mapa de calor: día de la semana x hora -----------------------------
    calor = f.pivot_table(index="dow", columns="hora", values="total", aggfunc="sum", fill_value=0)
    calor = calor.reindex(index=range(7), columns=range(24), fill_value=0)
    fig, ax = plt.subplots(figsize=(6.4, 3.4))
    sns.heatmap(calor, cmap="YlGnBu", linewidths=0.4, linecolor="white", xticklabels=2, ax=ax,
                cbar_kws={"format": FuncFormatter(formato_auto(calor.values.max())), "shrink": 0.85})
    ax.set_yticklabels(DIAS_CORTO, rotation=0)
    ax.set_xlabel("Hora del día")
    ax.set_ylabel("")
    g["calor"] = a_imagen(fig)

    # 5) Distribución del valor de las facturas -----------------------------
    fig, ax = plt.subplots(figsize=(6.4, 3.4))
    sns.histplot(f["total"], bins=25, kde=True, color=BASE, edgecolor="white", ax=ax)
    promedio, mediana = f["total"].mean(), f["total"].median()
    ax.axvline(promedio, color=ACENTO, linestyle="--", linewidth=1.8, label=f"Promedio {pesos(promedio)}")
    ax.axvline(mediana, color="#475569", linestyle=":", linewidth=1.8, label=f"Mediana {pesos(mediana)}")
    ax.legend(frameon=False, fontsize=8)
    ax.xaxis.set_major_formatter(FuncFormatter(a_miles))
    ax.set_xlabel("Valor de la factura")
    ax.set_ylabel("Cantidad de facturas")
    sns.despine(ax=ax)
    g["distribucion"] = a_imagen(fig)

    # 6) Top productos por ingresos / 7) por unidades -----------------------
    g["prod_ingreso"] = barras_horizontales(top(d, "producto", "ingreso"), "ingreso", "producto")
    g["prod_unidades"] = barras_horizontales(top(d, "producto", "cantidad"), "cantidad", "producto", fmt=entero)

    # 8) Ingresos por tipo de producto --------------------------------------
    g["tipo"] = barras_horizontales(top(d, "tipo", "ingreso", n=10), "ingreso", "tipo")

    # 9) Mejores clientes / 10) Ventas por barrio ---------------------------
    g["clientes"] = barras_horizontales(top(f, "cliente_et", "total"), "total", "cliente_et")
    g["barrios"] = barras_horizontales(top(f, "barrio", "total"), "total", "barrio")

    # 11) Valor de factura por puesto (diagrama de cajas) -------------------
    fig, ax = plt.subplots(figsize=(6.4, 4.4))
    orden = f.groupby("puesto")["total"].median().sort_values(ascending=False).index
    sns.boxplot(data=f, x="total", y="puesto", order=orden, color="#99D5CF",
                fliersize=2, linewidth=1, ax=ax)
    ax.xaxis.set_major_formatter(FuncFormatter(a_miles))
    ax.set_xlabel("Valor de la factura")
    ax.set_ylabel("")
    sns.despine(ax=ax, left=True)
    g["puestos"] = a_imagen(fig)

    # 12) Mejores empleados / 13) Mejores proveedores -----------------------
    g["empleados"] = barras_horizontales(top(f, "empleado_et", "total"), "total", "empleado_et")
    g["proveedores"] = barras_horizontales(top(d, "proveedor", "ingreso"), "ingreso", "proveedor")
    return g


# Orden de las secciones y qué gráfico va en cada una: (clave, título, explicación)
SECCIONES = [
    ("ventas", "Ventas", [
        ("mes", "Ventas por mes", "Cuánto se vendió cada mes (el último mes incompleto no se dibuja). En un año específico, el punto naranja es el mejor mes."),
        ("acumulado", "Ventas acumuladas", "Cómo va creciendo el total vendido con el tiempo."),
        ("dia", "Ventas por día de la semana", "El día que más vende aparece en naranja."),
        ("calor", "Mapa de calor: día y hora", "Los colores más oscuros son los momentos con más ventas."),
        ("distribucion", "Valor de las facturas", "La mayoría de facturas se concentra en valores bajos o medios."),
    ]),
    ("productos", "Productos", [
        ("prod_ingreso", "Top 10 productos por ingresos", "Los que más plata dejaron."),
        ("prod_unidades", "Top 10 productos por unidades", "Los que más se llevaron los clientes."),
        ("tipo", "Ingresos por tipo de producto", "Se agrupan por la primera palabra del nombre (Arroz, Leche, etc.)."),
    ]),
    ("clientes", "Clientes", [
        ("clientes", "Top 10 clientes", "Los que más han comprado en el periodo."),
        ("barrios", "Ventas por barrio", "De qué barrios vienen las compras (según la dirección del cliente)."),
    ]),
    ("equipo", "Equipo y proveedores", [
        ("empleados", "Top 10 empleados por ventas", "Quiénes más facturaron."),
        ("puestos", "Valor de factura por puesto", "Diagrama de cajas: la línea del centro es la mediana de cada puesto."),
        ("proveedores", "Top 10 proveedores por ingresos", "Proveedores de los productos más vendidos."),
        ("salarios", "Salarios por puesto", "Cada punto es un empleado. Este gráfico no cambia con el año."),
    ]),
]


# ----------------------------------------------------------------------------
# 6. INDICADORES (las tarjetas de arriba) Y TABLA DE FACTURAS
# ----------------------------------------------------------------------------
def indicadores(anio, fac, det, agotados):
    f = fac if anio is None else fac[fac["anio"] == anio]
    d = det if anio is None else det[det["anio"] == anio]
    ventas = f["total"].sum()

    # comparación con el mismo periodo del año anterior (solo cuando se elige un año)
    delta = None
    if anio is not None and (fac["anio"] == anio - 1).any():
        limite = f["fecha"].dt.dayofyear.max()
        previo = fac[(fac["anio"] == anio - 1) & (fac["fecha"].dt.dayofyear <= limite)]["total"].sum()
        delta = (ventas - previo) / previo * 100

    mejor = f.groupby(f["fecha"].dt.to_period("M"))["total"].sum().idxmax()
    tarjetas = [
        ("Ventas totales", "$ " + a_millones(ventas), delta),
        ("Facturas", f"{len(f):,}".replace(",", "."), None),
        ("Ticket promedio", pesos(f["total"].mean()), None),
        ("Clientes que compraron", f"{f['id_cliente'].nunique():,}".replace(",", "."), None),
        ("Unidades vendidas", f"{d['cantidad'].sum():,}".replace(",", "."), None),
        ("Mejor mes", f"{MESES[mejor.month - 1]} {mejor.year}", None),
        ("Productos agotados", str(agotados), None),
    ]
    html_tarjetas = ""
    for titulo, valor, cambio in tarjetas:
        extra = ""
        if cambio is not None:
            clase = "sube" if cambio >= 0 else "baja"
            flecha = "▲" if cambio >= 0 else "▼"
            extra = (f'<span class="delta {clase}">{flecha} {abs(cambio):.1f}'.replace(".", ",")
                     + " % vs. mismo periodo del año anterior</span>")
        html_tarjetas += (f'<div class="kpi"><span class="kpi-t">{titulo}</span>'
                          f'<span class="kpi-v">{valor}</span>{extra}</div>')
    return html_tarjetas


def tabla_facturas(anio, fac, n=25):
    f = fac if anio is None else fac[fac["anio"] == anio]
    ultimas = f.sort_values("fecha", ascending=False).head(n)
    filas = ""
    for _, r in ultimas.iterrows():
        filas += (f"<tr><td>{r['id_factura']}</td><td>{r['fecha']:%d/%m/%Y %H:%M}</td>"
                  f"<td>{html.escape(str(r['cliente_et']))}</td><td>{html.escape(str(r['empleado_et']))}</td>"
                  f"<td class='num'>{pesos(r['total'])}</td></tr>")
    return ("<table><thead><tr><th>Factura</th><th>Fecha</th><th>Cliente</th><th>Empleado</th>"
            f"<th class='num'>Total</th></tr></thead><tbody>{filas}</tbody></table>")


# ----------------------------------------------------------------------------
# 7. ARMADO DEL HTML
# ----------------------------------------------------------------------------
CSS = """
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;font-family:"Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif;background:#F1F5F9;color:#0F172A}
header{background:linear-gradient(120deg,#115E59,#0F766E 60%,#14B8A6);color:#fff;padding:28px 32px 20px}
header h1{margin:0;font-size:28px}
header p{margin:6px 0 0;opacity:.9;font-size:14px}
.barra{position:sticky;top:0;z-index:10;background:#fff;border-bottom:1px solid #E2E8F0;padding:10px 32px;
 display:flex;flex-wrap:wrap;gap:14px;align-items:center;justify-content:space-between;box-shadow:0 2px 6px rgba(15,23,42,.05)}
.grupo{display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.grupo span.et{font-size:12px;color:#64748B;margin-right:4px;text-transform:uppercase;letter-spacing:.04em}
.boton{border:1px solid #CBD5E1;background:#fff;color:#334155;padding:6px 14px;border-radius:999px;cursor:pointer;font-size:14px}
.boton:hover{background:#F0FDFA;border-color:#0F766E}
.boton.activo{background:#0F766E;color:#fff;border-color:#0F766E}
main{max-width:1400px;margin:0 auto;padding:24px 32px 40px}
.panel{display:none}
.panel.activo{display:block}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:28px}
.kpi{background:#fff;border-radius:14px;padding:16px 18px;box-shadow:0 1px 3px rgba(15,23,42,.08);display:flex;flex-direction:column;gap:4px;border-left:5px solid #0F766E}
.kpi-t{font-size:12px;color:#64748B;text-transform:uppercase;letter-spacing:.04em}
.kpi-v{font-size:24px;font-weight:700}
.delta{font-size:11px;font-weight:600;line-height:1.3}
.delta.sube{color:#15803D}.delta.baja{color:#B91C1C}
section[data-sec]{scroll-margin-top:72px}
h2{font-size:20px;margin:34px 0 14px;padding-bottom:6px;border-bottom:3px solid #14B8A6;display:inline-block}
.rejilla{display:grid;grid-template-columns:repeat(auto-fit,minmax(460px,1fr));gap:18px}
.grafico{margin:0;background:#fff;border-radius:14px;padding:16px;box-shadow:0 1px 3px rgba(15,23,42,.08)}
.grafico h3{margin:0;font-size:16px}
.grafico p{margin:2px 0 10px;font-size:13px;color:#64748B}
.grafico img{width:100%;height:auto;display:block;cursor:zoom-in;border-radius:6px}
.tabla-caja{background:#fff;border-radius:14px;padding:16px;box-shadow:0 1px 3px rgba(15,23,42,.08);overflow-x:auto}
.tabla-caja input{width:100%;max-width:360px;padding:8px 12px;border:1px solid #CBD5E1;border-radius:8px;margin-bottom:12px;font-size:14px}
table{border-collapse:collapse;width:100%;font-size:13px}
th{text-align:left;background:#F0FDFA;color:#115E59;padding:8px 10px;position:sticky;top:0}
td{padding:7px 10px;border-top:1px solid #E2E8F0}
.num{text-align:right}
tbody tr:hover{background:#F8FAFC}
.notas{margin-top:34px;background:#FEF9C3;border:1px solid #FDE68A;border-radius:12px;padding:14px 20px;font-size:13px;color:#713F12}
.notas ul{margin:6px 0 0;padding-left:20px}
footer{text-align:center;color:#64748B;font-size:12px;padding:0 0 30px}
#visor{position:fixed;inset:0;background:rgba(15,23,42,.82);display:none;align-items:center;justify-content:center;z-index:50;flex-direction:column;gap:12px;padding:20px}
#visor.abierto{display:flex}
#visor img{max-width:96vw;max-height:84vh;background:#fff;border-radius:10px}
#visor a{color:#fff;font-size:14px}
@media(max-width:560px){.rejilla{grid-template-columns:1fr}header,.barra,main{padding-left:16px;padding-right:16px}}
"""

JS = """
const botonesAnio = document.querySelectorAll('[data-filtro]');
const paneles = document.querySelectorAll('.panel');
function mostrar(valor){
  paneles.forEach(p => p.classList.toggle('activo', p.dataset.anio === valor));
  botonesAnio.forEach(b => b.classList.toggle('activo', b.dataset.filtro === valor));
  document.getElementById('buscador').value = '';
  filtrarTabla();
}
botonesAnio.forEach(b => b.addEventListener('click', () => mostrar(b.dataset.filtro)));

document.querySelectorAll('[data-ir]').forEach(b => b.addEventListener('click', () => {
  const destino = document.querySelector('.panel.activo [data-sec="' + b.dataset.ir + '"]');
  if (destino) destino.scrollIntoView({behavior:'smooth', block:'start'});
}));

function filtrarTabla(){
  const q = document.getElementById('buscador').value.toLowerCase();
  document.querySelectorAll('.panel.activo tbody tr').forEach(tr => {
    tr.style.display = tr.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
}
document.getElementById('buscador').addEventListener('input', filtrarTabla);

const visor = document.getElementById('visor');
const visorImg = document.getElementById('visor-img');
const descarga = document.getElementById('visor-descarga');
document.addEventListener('click', e => {
  if (e.target.matches('.grafico img')) {
    visorImg.src = e.target.src;
    descarga.href = e.target.src;
    descarga.download = (e.target.alt || 'grafico').replace(/\\s+/g, '_') + '.png';
    visor.classList.add('abierto');
  } else if (e.target === visor || e.target === visorImg) {
    visor.classList.remove('abierto');
  }
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') visor.classList.remove('abierto'); });
"""


def tarjeta_grafico(titulo, explicacion, imagen):
    t = html.escape(titulo)
    return (f'<figure class="grafico"><h3>{t}</h3><p>{html.escape(explicacion)}</p>'
            f'<img src="data:image/png;base64,{imagen}" alt="{t}" loading="lazy"></figure>')


def construir_html(datos, origen):
    fac, det, emp = datos["fac"], datos["det"], datos["emp"]
    av = datos["avisos"]
    anios = sorted(fac["anio"].unique())
    filtros = [("todos", None, "Todos")] + [(str(a), int(a), str(a)) for a in anios]

    estaticos = graficos_estaticos(emp)
    paneles = ""
    for clave, anio, _ in filtros:
        print(f"  dibujando gráficos de: {clave} ...")
        g = graficos_del_periodo(anio, fac, det)
        g.update(estaticos)

        cuerpo = f'<div class="kpis">{indicadores(anio, fac, det, av["agotados"])}</div>'
        for sec_id, sec_titulo, graficos in SECCIONES:
            cuerpo += f'<section data-sec="{sec_id}"><h2>{sec_titulo}</h2><div class="rejilla">'
            cuerpo += "".join(tarjeta_grafico(t, e, g[k]) for k, t, e in graficos)
            cuerpo += "</div></section>"
        cuerpo += ('<section data-sec="facturas"><h2>Últimas facturas</h2><div class="tabla-caja">'
                   + tabla_facturas(anio, fac) + "</div></section>")
        activo = " activo" if clave == "todos" else ""
        paneles += f'<div class="panel{activo}" data-anio="{clave}">{cuerpo}</div>'

    botones_anio = "".join(
        f'<button class="boton{" activo" if c == "todos" else ""}" data-filtro="{c}">{n}</button>'
        for c, _, n in filtros)
    botones_sec = "".join(f'<button class="boton" data-ir="{s}">{t}</button>' for s, t, _ in SECCIONES)
    botones_sec += '<button class="boton" data-ir="facturas">Facturas</button>'

    nota_ids = ""
    if av["ids_arreglados"]:
        nota_ids = ("<li>El Excel traía los IDs repetidos (cada hoja estaba pegada dos veces y los números empezaban de nuevo). "
                    "Se separó cada bloque como un lote distinto; con eso el total de cada factura coincide con "
                    f"cantidad × precio de su pedido en el {av['coinciden']:.0f} % de los casos.</li>")
    notas = f"""
    <div class="notas"><strong>Notas sobre los datos</strong><ul>
      {nota_ids}
      <li>Cada una de las {av['facturas']} facturas tiene un pedido con un producto; las gráficas de productos, tipos y proveedores salen de esos pedidos.</li>
      <li>{av['sin_nombre']} productos no tenían nombre en la base de datos; se muestran como "Producto sin nombre".</li>
    </ul></div>"""

    return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard Mi Tienda</title><style>{CSS}</style></head>
<body>
<header><h1>Dashboard Mi Tienda</h1>
<p>Ventas, productos, clientes y equipo · Datos desde: {html.escape(origen)} · Generado el {datetime.now():%d/%m/%Y %H:%M}</p></header>
<div class="barra">
  <div class="grupo"><span class="et">Año</span>{botones_anio}</div>
  <div class="grupo"><span class="et">Ir a</span>{botones_sec}</div>
  <div class="grupo"><input id="buscador" type="search" placeholder="Buscar en las últimas facturas..."
       style="padding:6px 12px;border:1px solid #CBD5E1;border-radius:999px;font-size:14px;min-width:230px"></div>
</div>
<main>{paneles}{notas}</main>
<footer>Hecho con Python · pandas · Seaborn · Haz clic en cualquier gráfico para verlo grande.</footer>
<div id="visor"><img id="visor-img" alt=""><a id="visor-descarga" href="#">Descargar gráfico (PNG)</a></div>
<script>{JS}</script></body></html>"""


# ----------------------------------------------------------------------------
# 8. PROGRAMA PRINCIPAL
# ----------------------------------------------------------------------------
def main():
    pa = argparse.ArgumentParser(description="Dashboard de Mi Tienda con Seaborn")
    pa.add_argument("--origen", choices=["excel", "bd"], default="excel", help="de dónde se leen los datos")
    pa.add_argument("--excel", default=ARCHIVO_EXCEL, help="ruta del archivo Excel")
    pa.add_argument("--bd-url", default=BD_URL, help="cadena de conexión a la base de datos")
    pa.add_argument("--salida", default=ARCHIVO_HTML, help="archivo HTML que se genera")
    pa.add_argument("--cargar-bd", action="store_true", help="sube el Excel a la base de datos y termina")
    a = pa.parse_args()

    if a.cargar_bd:
        print("Subiendo el Excel a la base de datos...")
        subir_excel_a_bd(a.excel, a.bd_url)
        print("Listo. Ahora se puede correr con:  --origen bd")
        return

    if a.origen == "bd":
        print("Leyendo datos desde la base de datos...")
        tablas, origen = cargar_bd(a.bd_url), "base de datos"
    else:
        print("Leyendo datos desde el Excel...")
        tablas, origen = cargar_excel(a.excel), "archivo Excel"

    datos = preparar(tablas)
    print("Armando el dashboard...")
    pagina = construir_html(datos, origen)
    with open(a.salida, "w", encoding="utf-8") as archivo:
        archivo.write(pagina)
    print(f"Dashboard generado: {a.salida} ({os.path.getsize(a.salida) / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
