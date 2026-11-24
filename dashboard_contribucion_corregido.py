
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(
    page_title="Contribución SENATI",
    page_icon="📊",
    layout="wide"
)

BASE_DIR = Path(__file__).resolve().parent

# El archivo Excel está en la carpeta BASE_CONTRI (un nivel por encima
# de "files (8)" donde puede estar guardado este .py).
# En Windows puede estar oculto el sufijo .xlsx en el Explorador.
_CANDIDATOS_EXCEL = [
    BASE_DIR.parent / "BASE_CONTRI.xlsx",
    BASE_DIR.parent / "BASE_CONTRI(1).xlsx",
    BASE_DIR / "BASE_CONTRI.xlsx",
    BASE_DIR / "BASE_CONTRI(1).xlsx",
]

ARCHIVO = next((ruta for ruta in _CANDIDATOS_EXCEL if ruta.is_file()), None)

# Como respaldo, acepta otros nombres que comiencen con BASE_CONTRI.
if ARCHIVO is None:
    for carpeta in (BASE_DIR.parent, BASE_DIR):
        encontrados = sorted(carpeta.glob("BASE_CONTRI*.xlsx"))
        if encontrados:
            ARCHIVO = encontrados[0]
            break

if ARCHIVO is None:
    raise FileNotFoundError(
        "No se encontró el archivo Excel BASE_CONTRI.xlsx. "
        f"Se buscó en: {BASE_DIR.parent} y {BASE_DIR}. "
        "Coloca el Excel en la carpeta BASE_CONTRI, junto a la carpeta files (8)."
    )

MESES = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",7:"JUL",8:"AGO",9:"SET",10:"OCT",11:"NOV",12:"DIC"}
MES_INV = {v:k for k,v in MESES.items()}
MES_ORDEN = list(MESES.values())

# Paleta institucional SENATI (azul); ajustar aquí si se cuenta con el código oficial del manual de marca
# Altos de los gráficos (px). Cada fila usa el mismo alto a izquierda y derecha para que
# los títulos y los bordes queden alineados; la suma está calibrada con la columna central.
ALTO_FILA1=456    # ingresos mensuales | contribución por sector
ALTO_FILA2=396    # ingresos anuales   | contribución por origen
HUESO="#F5F0E1"   # fondo de los gráficos (tono hueso)
AZUL="#003B8E"; AZUL2="#0A2F6B"; AZUL3="#2F6FD0"; NARANJA="#E98236"; VERDE="#B9DDB0"

import re, unicodedata

# Nombres de zonal en BASE -> nombre usado en PRESUPUESTO
EQUIV_ZONAL = {
    "CAJAMARCA SAN MARTIN AMAZONAS": "CAJAMARCA",
    "CUSCO APURIMAC MADRE DE DIOS": "CUSCO APURIMAC",
    "JUNIN PASCO HUANCAVELICA": "JUNIN PASCO",
}

# Zonales fuera del filtro y de los cálculos; su presupuesto solo suma al presupuesto anual nacional
ZONAL_OCULTAS = {"DIRECCION NACIONAL"}

def norm_txt(v):
    """Unifica nombres: sin tildes, mayúsculas, sin guiones/puntos, sin espacios dobles ni prefijo 'ZONAL'."""
    if pd.isna(v):
        return "(Sin dato)"
    t = unicodedata.normalize("NFKD", str(v)).encode("ascii", "ignore").decode()
    t = re.sub(r"[^A-Za-z0-9]+", " ", t).upper().strip()
    t = re.sub(r"^(DIRECCION )?ZONAL ", "", t)
    return EQUIV_ZONAL.get(t, t) or "(Sin dato)"

@st.cache_data
def cargar():
    base = pd.read_excel(ARCHIVO, sheet_name="BASE")
    p = pd.read_excel(ARCHIVO, sheet_name="PRESUPUESTO")

    for c in ["CONTRIBUCION","TOTAL"]:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)
    base["AÑO"] = pd.to_numeric(base["AÑO"], errors="coerce")
    # FECHA viene mezclada (fechas y textos dd/mm/aaaa): se unifica a fecha
    _f = base["FECHA"]
    _txt = _f.map(lambda v: isinstance(v, str))
    base["FECHA"] = pd.to_datetime(_f.where(~_txt), errors="coerce").fillna(
        pd.to_datetime(_f.where(_txt), format="%d/%m/%Y", errors="coerce"))
    base["PERIODO"] = pd.to_datetime(base["PERIODO"], errors="coerce")
    for c in ["INTERES","COSTAS"]:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)
    base["MES_CONT_DT"] = pd.to_datetime(base["MES CONT."], errors="coerce")
    base["MES_NUM"] = base["MES_CONT_DT"].dt.month
    base["MES"] = base["MES_NUM"].map(MESES)

    p["AÑO"] = pd.to_numeric(p["AÑO"], errors="coerce")
    for c in ["INGRESOS","PRESUPUESTOS"]:
        p[c] = pd.to_numeric(
            p[c].astype(str).str.replace(r"[^0-9.\-]","",regex=True),
            errors="coerce"
        ).fillna(0)
    p["MES"] = p["MES"].astype(str).str.strip().str.upper().str[:3].replace({"SEP":"SET"})
    p["MES_NUM"] = p["MES"].map(MES_INV)
    p["TIPO_NORM"] = p["TIPO"].astype(str).str.upper().str.replace("Ó","O")
    base["TIPO_NORM"] = base["TIPO"].astype(str).str.upper().str.replace("Ó","O")

    for c in ["ZONAL","ORIGEN","SECTOR","TIPO","ESTADO","CONTRIBUYENTES"]:
        if c in base:
            base[c] = base[c].fillna("(Sin dato)").astype(str)
    for c in ["ZONAL","TIPO"]:
        p[c] = p[c].fillna("(Sin dato)").astype(str)

    # Misma escritura de ZONAL en ambas hojas, para que el filtro cruce
    # Nombre largo de cada zonal tal como figura en BASE (para mostrar en tablas)
    _orig = base["ZONAL"].astype(str).str.strip()
    nombres_zonal = _orig.groupby(_orig.map(norm_txt)).agg(lambda x: x.value_counts().index[0]).to_dict()
    base["ZONAL"] = base["ZONAL"].map(norm_txt)
    p["ZONAL"] = p["ZONAL"].map(norm_txt)

    # Presupuesto de las zonales excluidas (Dirección Nacional): se guarda aparte porque el
    # Tablero de Control lo cuenta solo en el presupuesto ANUAL nacional, no en el acumulado.
    dn = (p[p["ZONAL"].isin(ZONAL_OCULTAS)]
          .groupby(["AÑO","TIPO_NORM"], as_index=False)["PRESUPUESTOS"].sum())

    # Zonales excluidas del resto de cálculos
    base = base[~base["ZONAL"].isin(ZONAL_OCULTAS)].copy()
    p = p[~p["ZONAL"].isin(ZONAL_OCULTAS)].copy()

    # INGRESOS: para los años que existen en BASE se toma el TOTAL pagado de BASE
    # (fuente de detalle); PRESUPUESTO aporta el presupuesto y el histórico anterior.
    claves = ["ZONAL","TIPO_NORM","AÑO","MES_NUM"]
    ing = base.groupby(claves, as_index=False)["TOTAL"].sum()
    anios_base = set(base["AÑO"].dropna().unique())
    p = p.merge(ing, on=claves, how="outer")
    en_base = p["AÑO"].isin(anios_base)
    p.loc[en_base, "INGRESOS"] = p.loc[en_base, "TOTAL"].fillna(0)
    p["INGRESOS"] = p["INGRESOS"].fillna(0)
    p["PRESUPUESTOS"] = p["PRESUPUESTOS"].fillna(0)
    p["MES"] = p["MES_NUM"].map(MESES)
    p["TIPO"] = p["TIPO"].fillna(p["TIPO_NORM"])
    p = p.drop(columns="TOTAL")

    return base, p, dn, nombres_zonal

def fmt(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")

def money(x):
    return fmt(x)

def pct(x):
    return f"{x:.1f}%"

def card(title, value, color=AZUL):
    st.markdown(f"""
    <div style="border:1px solid #777;background:#fff;margin-bottom:8px;">
      <div style="background:{color};color:#fff;text-align:center;font-weight:800;padding:5px;font-size:12px;">{title}</div>
      <div style="text-align:center;font-size:19px;font-weight:800;padding:9px 3px;color:#222;">{value}</div>
    </div>""", unsafe_allow_html=True)

def titulo(texto):
    """Título de sección con franja en el azul institucional."""
    st.markdown(
        f"<div style='background:{AZUL};color:#fff;font-weight:800;font-size:17px;letter-spacing:.3px;"
        f"padding:8px 14px;margin:0 0 6px 0;border-left:6px solid {AZUL3};'>{texto}</div>",
        unsafe_allow_html=True)

def tabla_html(df, enteros=(), derecha=(), alto=None, fijas=None):
    """Tabla con encabezado azul y letras claras (mismo formato que las tarjetas).
    alto: si se indica (px), la tabla se desplaza dentro de ese alto con el encabezado fijo.
    fijas: {columna: ancho_px} de las primeras columnas que quedan ancladas a la izquierda."""
    import html as _html
    global _TABLA_N
    _TABLA_N = globals().get("_TABLA_N", 0) + 1
    k = f"tbs{_TABLA_N}"
    cols = list(df.columns)
    num = [pd.api.types.is_numeric_dtype(df[c]) for c in cols]
    der = [n or c in derecha for c, n in zip(cols, num)]

    css = [
        f".{k}{{border-collapse:collapse;width:100%;}}",
        f".{k} th{{background:{AZUL2};color:#fff;font-weight:800;font-size:12px;padding:7px 10px;"
        f"border:1px solid {AZUL2};white-space:nowrap;text-align:left;{'position:sticky;top:0;z-index:1;' if alto else ''}}}",
        f".{k} td{{background:#fff;color:#111;font-weight:600;font-size:13px;padding:6px 10px;"
        f"border:1px solid #999;white-space:nowrap;text-align:left;}}",
        f".{k} .r{{text-align:right;}}",
    ]
    izq = 0
    for c, w in (fijas or {}).items():
        n = cols.index(c) + 1
        css.append(f".{k} th:nth-child({n}),.{k} td:nth-child({n}){{position:sticky;left:{izq}px;"
                   f"min-width:{w}px;max-width:{w}px;overflow:hidden;text-overflow:ellipsis;}}")
        css.append(f".{k} th:nth-child({n}){{z-index:3;}}")
        izq += w

    h = "".join(f"<th{' class=r' if d else ''}>{_html.escape(str(c))}</th>" for c, d in zip(cols, der))
    filas = []
    for r in df.itertuples(index=False):
        celdas = ""
        for v, c, n, d in zip(r, cols, num, der):
            if n:
                v = "" if pd.isna(v) else (f"{int(v)}" if c in enteros else money(v))
            else:
                v = "" if pd.isna(v) else _html.escape(str(v))
            celdas += f"<td{' class=r' if d else ''}>{v}</td>"
        filas.append(f"<tr>{celdas}</tr>")
    caja = f"overflow:auto;max-height:{alto}px;" if alto else "overflow-x:auto;"
    st.markdown(
        f"<style>{''.join(css)}</style><div style='{caja}margin-bottom:14px;'><table class='{k}'>"
        f"<thead><tr>{h}</tr></thead><tbody>{''.join(filas)}</tbody></table></div>",
        unsafe_allow_html=True)

def filt_common(df, years, months, zonals, types):
    x=df.copy()
    if years: x=x[x["AÑO"].isin(years)]
    if months: x=x[x["MES_NUM"].isin(months)]
    if zonals: x=x[x["ZONAL"].isin(zonals)]
    if types: x=x[x["TIPO_NORM"].isin(types)]
    return x

def bar_mensual(g):

    g = (
        g.groupby("MES_NUM", as_index=False)["INGRESOS"]
        .sum()
        .set_index("MES_NUM")
        .reindex(range(1, 13), fill_value=0)
        .reset_index()
    )

    g["MES"] = g["MES_NUM"].map(MESES)
    g["VALOR"] = g["INGRESOS"] / 1000

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=g["MES"],
            y=g["VALOR"],
            text=[f"{v:,.0f}" for v in g["VALOR"]],
            textposition="outside",
            textfont=dict(
                color="#000000",
                size=15
            ),
            marker=dict(
                color=VERDE
            ),
            cliponaxis=False,
            hovertemplate="<b>%{x}</b><br>%{y:,.0f} miles de soles<extra></extra>"
        )
    )

    fig.update_layout(
        height=ALTO_FILA1,
        margin=dict(
            l=10,
            r=15,
            t=35,
            b=35
        ),
        plot_bgcolor=HUESO,
        paper_bgcolor=HUESO,
        showlegend=False,
        font=dict(
            family="Arial",
            color="#000000"
        ),
        xaxis=dict(
            showgrid=False,
            tickfont=dict(
                color="#000000",
                size=10
            ),
            linecolor="#000000"
        ),
        yaxis=dict(
            visible=False,
            showgrid=False,
            zeroline=False
        ),
        uniformtext=dict(
            minsize=9,
            mode="show"
        )
    )

    return fig

def bar_anual(g):

    g = (
        g.groupby("AÑO", as_index=False)["INGRESOS"]
        .sum()
        .sort_values("AÑO")
    )

    g["VALOR"] = g["INGRESOS"] / 1000
    g["AÑO"] = g["AÑO"].astype(int).astype(str)

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=g["AÑO"],
            y=g["VALOR"],
            text=[f"{v:,.0f}" for v in g["VALOR"]],
            textposition="outside",
            textfont=dict(
                color="#000000",
                size=15
            ),
            marker=dict(
                color=NARANJA
            ),
            cliponaxis=False,
            hovertemplate="<b>%{x}</b><br>%{y:,.0f} miles de soles<extra></extra>"
        )
    )

    fig.update_layout(
        height=ALTO_FILA2,
        margin=dict(
            l=10,
            r=15,
            t=35,
            b=35
        ),
        plot_bgcolor=HUESO,
        paper_bgcolor=HUESO,
        showlegend=False,
        font=dict(
            family="Arial",
            color="#000000"
        ),
        xaxis=dict(
            type="category",
            showgrid=False,
            tickfont=dict(
                color="#000000",
                size=12
            ),
            linecolor="#000000"
        ),
        yaxis=dict(
            visible=False,
            showgrid=False,
            zeroline=False
        ),
        uniformtext=dict(
            minsize=9,
            mode="show"
        )
    )

    return fig


def donut_cumplimiento(actual_acum,budget_acum):
    valor = actual_acum/budget_acum*100 if budget_acum else 0
    fig=go.Figure(go.Pie(values=[min(valor,100),max(100-min(valor,100),0)],
                         labels=["Cumplimiento","Pendiente"],hole=.58,
                         marker=dict(colors=[AZUL3,"#E5E7EB"]),
                         textinfo="none"))
    fig.add_annotation(text=f"<b>{valor:.0f}%</b>",x=.5,y=.58,showarrow=False,font=dict(size=22,color="#555"))
    fig.update_layout(height=250,margin=dict(l=0,r=0,t=8,b=8),showlegend=False,paper_bgcolor=HUESO)
    return fig

def bar_sector(base):

    if base.empty:
        fig = go.Figure()

        fig.add_annotation(
            text="No existen datos para los filtros seleccionados",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
            font=dict(
                color="#C00000",
                size=16
            )
        )

        fig.update_layout(
            height=ALTO_FILA1,
            plot_bgcolor=HUESO,
            paper_bgcolor=HUESO,
            xaxis=dict(visible=False),
            yaxis=dict(visible=False)
        )

        return fig

    g = (
        base
        .groupby("SECTOR", as_index=False)["TOTAL"]
        .sum()
        .sort_values("TOTAL", ascending=True)
    )

    total = g["TOTAL"].sum()

    g["PORC"] = (
        g["TOTAL"] / total * 100
        if total > 0
        else 0
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=g["PORC"],
            y=g["SECTOR"],
            orientation="h",

            text=[f"<b><i>{v:.2f}%</i></b>" for v in g["PORC"]],
            textposition="outside",

            textfont=dict(
                color="#000000",
                size=13
            ),

            marker=dict(
                color=AZUL3
            ),

            cliponaxis=False,

            hovertemplate=(
                "<b>%{y}</b><br>"
                "%{x:.2f}%<extra></extra>"
            )
        )
    )

    fig.update_layout(
        height=ALTO_FILA1,

        margin=dict(
            l=10,
            r=85,
            t=10,
            b=10
        ),

        plot_bgcolor=HUESO,
        paper_bgcolor=HUESO,

        showlegend=False,

        xaxis=dict(
            visible=False,
            showgrid=False,
            zeroline=False,
            range=[
                0,
                max(g["PORC"].max() * 1.20, 5)
            ]
        ),

        yaxis=dict(
            showgrid=False,
            tickfont=dict(
                color="#000000",
                size=11
            ),
            automargin=True
        )
    )

    return fig

def pie_origen(base):

    if base.empty:
        fig = go.Figure()
        fig.add_annotation(text="No existen datos para los filtros seleccionados",
                           x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False,
                           font=dict(color="#C00000", size=16))
        fig.update_layout(height=ALTO_FILA2, plot_bgcolor=HUESO, paper_bgcolor=HUESO,
                          xaxis=dict(visible=False), yaxis=dict(visible=False))
        return fig

    g = (
        base.groupby("ORIGEN", as_index=False)["TOTAL"]
        .sum()
    )

    fig = go.Figure(
        go.Pie(
            values=g["TOTAL"],
            labels=g["ORIGEN"],

            hole=0.48,

            textinfo="percent",

            texttemplate="%{percent:.2%}",

            textposition="inside",

            textfont=dict(
                color="#000000",
                size=15
            ),

            marker=dict(
                colors=[
                    "#F2B08A",
                    "#8EA3DB",
                    "#A6A6A6"
                ]
            ),

            hovertemplate=(
                "<b>%{label}</b><br>"
                "%{percent:.2%}<extra></extra>"
            )
        )
    )

    fig.update_layout(
        height=ALTO_FILA2,

        margin=dict(
            l=5,
            r=5,
            t=5,
            b=5
        ),

        paper_bgcolor=HUESO,

        showlegend=True,

        legend=dict(
            orientation="v",
            x=0.75,
            y=0.5,
            font=dict(
                color="#000000",
                size=11
            )
        )
    )

    return fig

def a_excel(df, hoja, cols_monto=(), cols_fecha=(), cols_entero=()):
    """Devuelve el DataFrame como archivo .xlsx con números y fechas reales (aptos para cálculo)."""
    import io
    from openpyxl.utils import get_column_letter
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name=hoja, index=False)
        ws = xw.sheets[hoja]
        formatos = {}
        formatos.update({c: "#,##0.00" for c in cols_monto})
        formatos.update({c: "dd/mm/yyyy" for c in cols_fecha})
        formatos.update({c: "0" for c in cols_entero})
        for i, col in enumerate(df.columns, start=1):
            letra = get_column_letter(i)
            fmt = formatos.get(col)
            if fmt:
                for (celda,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    celda.number_format = fmt
            ancho = 45 if col in ("CONTRIBUYENTES", "DIVISION") else 16
            ws.column_dimensions[letra].width = ancho
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def proyeccion_cierre(pres, anio, M, zonals, types):
    """Proyección de cierre de año con corte al mes M (acumulado ENE..M).
    Contribución Regular: % de cumplimiento acumulado actual + cambio histórico de ese % entre
    el mes M y diciembre en los 4 años previos (promedio = cambio típico; conservador = peor cambio).
    Recuperación de Deuda: ritmo mensual actual extendido a 12 meses."""
    d = filt_common(pres, [], [], zonals, types)
    es_cr = d["TIPO_NORM"] == "CONTRIBUCION REGULAR"
    cr, rd = d[es_cr], d[~es_cr]

    def acum(x, y, m=12):
        r = x[(x["AÑO"] == y) & (x["MES_NUM"] <= m)]
        return r["INGRESOS"].sum(), r["PRESUPUESTOS"].sum()

    i_m, p_m = acum(cr, anio, M)
    _, p_f = acum(cr, anio)
    cambios, anios_hist = [], []
    for y in range(anio - 4, anio):
        im, pm = acum(cr, y, M)
        i_f, pf = acum(cr, y)
        if pm and pf and i_f:
            cambios.append(i_f / pf - im / pm)
            anios_hist.append(y)

    if p_m and p_f:
        if not cambios:
            return None
        cumpl = i_m / p_m
        cr_prom = (cumpl + sum(cambios) / len(cambios)) * p_f
        cr_cons = (cumpl + min(cambios)) * p_f
    else:
        cr_prom = cr_cons = i_m  # sin presupuesto de CR en la selección

    rd_proy = acum(rd, anio, M)[0] * 12 / M
    ppto_anual = d[d["AÑO"] == anio]["PRESUPUESTOS"].sum()
    return dict(prom=cr_prom + rd_proy, cons=cr_cons + rd_proy,
                ppto_anual=ppto_anual, anios_hist=anios_hist)


# -------------------- ACCESO CON CLAVE (opcional) --------------------
# Si la variable de entorno DASHBOARD_CLAVE está definida, se pide esa clave antes de mostrar datos.
import os, hmac
_CLAVE = os.environ.get("DASHBOARD_CLAVE", "").strip()
if _CLAVE and not st.session_state.get("acceso_ok"):
    st.markdown(f"<div style='background:{AZUL};color:#fff;font-weight:800;font-size:22px;text-align:center;padding:10px;'>CONTRIBUCIÓN SENATI</div>", unsafe_allow_html=True)
    _intento = st.text_input("Clave de acceso", type="password")
    if _intento:
        if hmac.compare_digest(_intento, _CLAVE):
            st.session_state["acceso_ok"] = True
            st.rerun()
        else:
            st.error("Clave incorrecta.")
    st.stop()

base,pres,dn_ppto,NOMBRE_ZONAL=cargar()

# -------------------- SIDEBAR --------------------
with st.sidebar:
    st.markdown(f"<div style='background:{AZUL};color:white;text-align:center;font-weight:900;padding:8px;'>FILTROS</div>",unsafe_allow_html=True)
    years_all=sorted(pres["AÑO"].dropna().astype(int).unique())
    default_year=[max(years_all)] if years_all else []
    years=st.multiselect("AÑO",years_all,default=default_year)
    months=st.multiselect("MES",MES_ORDEN,default=MES_ORDEN)
    month_nums=[MES_INV[m] for m in months]
    zonals=st.multiselect("ZONAL",sorted(z for z in pres["ZONAL"].unique() if z not in ZONAL_OCULTAS),default=[])
    types=st.multiselect("TIPO",sorted(pres["TIPO_NORM"].unique()),default=[])
    st.markdown("---")
    st.caption("Los filtros ZONAL y TIPO se aplican a presupuesto/ingresos. ORIGEN y SECTOR pertenecen a la BASE.")
    sin_cruce = sorted(set(pres["ZONAL"]) - set(base["ZONAL"]) - ZONAL_OCULTAS)
    if sin_cruce:
        with st.expander("⚠️ Zonales sin datos en BASE"):
            st.write("En PRESUPUESTO pero no en BASE:", sin_cruce)
            st.write("En BASE pero no en PRESUPUESTO:", sorted(set(base["ZONAL"]) - set(pres["ZONAL"])))
    if st.button("Restablecer filtros",use_container_width=True):
        st.rerun()

########## -------------------- DATA FILTRADA --------------------
# -------------------- FILTROS DE BASE --------------------

with st.sidebar:

    origins = st.multiselect(
        "ORIGEN",
        sorted(base["ORIGEN"].unique()),
        default=sorted(base["ORIGEN"].unique())
    )

    sectors = st.multiselect(
        "SECTOR",
        sorted(base["SECTOR"].unique()),
        default=sorted(base["SECTOR"].unique())
    )


# -------------------- BASE PARA GRÁFICOS DE CONTRIBUCIÓN --------------------

bp = base.copy()

# Año
if years:
    bp = bp[bp["AÑO"].isin(years)]

# Mes
if month_nums:
    bp = bp[bp["MES_NUM"].isin(month_nums)]

# Zonal
if zonals:
    bp = bp[bp["ZONAL"].isin(zonals)]

# Tipo
if types:
    bp = bp[bp["TIPO_NORM"].isin(types)]

# Origen
if origins:
    bp = bp[bp["ORIGEN"].isin(origins)]

# Sector
if sectors:
    bp = bp[bp["SECTOR"].isin(sectors)]


# -------------------- PRESUPUESTO --------------------

pp = filt_common(
    pres,
    years,
    month_nums,
    zonals,
    types
)

# -------------------- FUENTE PRINCIPAL --------------------
# Mensual/anual: PRESUPUESTO contiene el histórico 2018-2026 y reproduce los valores de referencia.
# Cuando se filtra ORIGEN o SECTOR, no existe esa dimensión en PRESUPUESTO; por ello se mantiene
# el gráfico presupuestal sin alterar por esas dos selecciones.
if not years:
    years_calc=years_all
else:
    years_calc=years

pp_chart=filt_common(pres,years_calc,month_nums,zonals,types)
monthly=pp_chart.groupby("MES_NUM",as_index=False)["INGRESOS"].sum()
annual=pp_chart.groupby("AÑO",as_index=False)["INGRESOS"].sum()
# Gráfico de ingresos anuales: siempre todos los años, sin aplicar ningún filtro
annual_total=pres.groupby("AÑO",as_index=False)["INGRESOS"].sum().sort_values("AÑO")

# Para KPIs de ejecución se toma el mayor año seleccionado y los meses seleccionados.
anio=max(years_calc) if years_calc else max(years_all)
actual_anio=pp_chart[pp_chart["AÑO"]==anio]
actual_acum=actual_anio["INGRESOS"].sum()
budget_acum=actual_anio["PRESUPUESTOS"].sum()
# Presupuesto anual completo (12 meses), respetando los filtros ZONAL y TIPO
budget_full=filt_common(pres,[anio],[],zonals,types)["PRESUPUESTOS"].sum()
# + presupuesto anual de Dirección Nacional, solo en la vista nacional (sin filtro de zonal)
dn_anual=0
if not zonals:
    _dn=dn_ppto[dn_ppto["AÑO"]==anio]
    if types: _dn=_dn[_dn["TIPO_NORM"].isin(types)]
    dn_anual=_dn["PRESUPUESTOS"].sum()
budget_full+=dn_anual

promedio=actual_anio.groupby("MES_NUM")["INGRESOS"].sum()
promedio=promedio[promedio>0].mean() if len(promedio[promedio>0]) else 0
mes_mayor_num=int(actual_anio.groupby("MES_NUM")["INGRESOS"].sum().idxmax()) if not actual_anio.empty else None
mes_mayor=MESES.get(mes_mayor_num,"-")
mes_mayor_val=actual_anio.groupby("MES_NUM")["INGRESOS"].sum().max() if not actual_anio.empty else 0

# Superávit/déficit vs presupuesto anual: presupuesto anual menos acumulado ejecutado.
saldo_pendiente=budget_full-actual_acum
cumpl=actual_acum/budget_acum*100 if budget_acum else 0
saldo_vs_acum=actual_acum-budget_acum
# Mes de corte: último mes seleccionado con ingresos en el año de análisis.
# Las comparaciones con otros años usan solo los meses seleccionados hasta ese corte.
_ing_mes=actual_anio.groupby("MES_NUM")["INGRESOS"].sum()
_ing_mes=_ing_mes[_ing_mes>0]
mes_corte=int(_ing_mes.index.max()) if len(_ing_mes) else None
meses_comp=[m for m in (month_nums or list(range(1,13))) if mes_corte and m<=mes_corte]
vs_anterior=None
prev=filt_common(pres,[anio-1],meses_comp,zonals,types)["INGRESOS"].sum() if meses_comp else 0
if prev: vs_anterior=(actual_acum/prev-1)*100
vs_2023=None
base23=filt_common(pres,[2023],meses_comp,zonals,types)["INGRESOS"].sum() if meses_comp else 0
if base23: vs_2023=(actual_acum/base23-1)*100
dif_2023=actual_acum-base23 if base23 else None
var_interanual=actual_acum-prev if prev else None

# -------------------- PROYECCIÓN DE CIERRE --------------------
proy=None
ing_corte=0
if mes_corte:
    ing_corte=filt_common(pres,[anio],list(range(1,mes_corte+1)),zonals,types)["INGRESOS"].sum()
    proy=proyeccion_cierre(pres,int(anio),mes_corte,zonals,types)
    if proy: proy["ppto_anual"]+=dn_anual

# Avance de recaudación (acumulado al corte / presupuesto anual) vs avance temporal (mes de corte / 12)
avance_pp=None
if proy and proy["ppto_anual"]:
    avance_pp=(ing_corte/proy["ppto_anual"]-mes_corte/12)*100

# Sector con mayor crecimiento interanual (TOTAL pagado, mismos meses y filtros)
sector_top,sector_crec=None,None
if meses_comp:
    bs=base[base["MES_NUM"].isin(meses_comp)]
    if zonals: bs=bs[bs["ZONAL"].isin(zonals)]
    if types: bs=bs[bs["TIPO_NORM"].isin(types)]
    if origins: bs=bs[bs["ORIGEN"].isin(origins)]
    if sectors: bs=bs[bs["SECTOR"].isin(sectors)]
    _a=bs[bs["AÑO"]==anio].groupby("SECTOR")["TOTAL"].sum()
    _b=bs[bs["AÑO"]==anio-1].groupby("SECTOR")["TOTAL"].sum()
    _c=(_a/_b[_b>0]-1).dropna()*100
    if len(_c):
        sector_top,sector_crec=_c.idxmax(),_c.max()

# -------------------- HEADER --------------------
st.markdown(f"""<div style="background:{AZUL};color:white;text-align:center;font-size:25px;font-weight:900;padding:8px;margin-bottom:8px;">CONTRIBUCIÓN SENATI</div>""",unsafe_allow_html=True)

tab_dash, tab_det, tab_contrib = st.tabs(["📊 Dashboard", "📋 Detalle de pagos", "🏢 Pagos por contribuyente"])

with tab_dash:
    c1,c2,c3=st.columns([1.65,.75,1.65])

    with c1:
        titulo("INGRESOS MENSUALES (MILES SOLES)")
        st.plotly_chart(bar_mensual(monthly), use_container_width=True, theme=None, config={"displayModeBar": False})
        titulo("INGRESOS ANUALES (MILES SOLES)")
        st.plotly_chart(bar_anual(annual_total), use_container_width=True, theme=None, config={"displayModeBar": False})


    with c2:
        card("PROMEDIO MENSUAL",money(promedio))
        card("MES MAYOR RECAUD. (MILES SOLES)",f"{mes_mayor}<br>{money(mes_mayor_val)}",AZUL2)
        card("ACUMULADO A LA FECHA",money(actual_acum),AZUL2)
        card("SUPERÁVIT/DÉFICIT VS PRES.",money(saldo_vs_acum),AZUL2)
        card("SUPERÁVIT/DÉFICIT VS PRES. %",pct((actual_acum/budget_acum-1)*100 if budget_acum else 0),AZUL2)
        card(f"CRECIMIENTO/DISM. VS EJEC {anio-1}", "N/D" if vs_anterior is None else pct(vs_anterior),AZUL2)
        card("CRECIMIENTO/DISM. VS EJEC 2023","N/D" if vs_2023 is None else pct(vs_2023),AZUL2)
        st.markdown('<div style="border:1px solid #777;background:white;"><div style="background:#003B8E;color:white;text-align:center;font-weight:800;padding:6px;">CUMPLIMIENTO PRESUPUESTO</div></div>',unsafe_allow_html=True)
        st.plotly_chart(donut_cumplimiento(actual_acum,budget_acum), use_container_width=True, theme=None, config={"displayModeBar": False})
        st.markdown(f"<div style='text-align:center;font-weight:800;margin-top:-10px;'>{cumpl:.1f}% acumulado</div>",unsafe_allow_html=True)

    with c3:
        titulo("CONTRIBUCIÓN POR SECTOR")
        if bp.empty:
            anios_base = sorted(base["AÑO"].dropna().astype(int).unique())
            st.warning(f"La hoja BASE solo tiene registros de los años: {', '.join(map(str, anios_base))}. "
                       "Sector y origen solo pueden mostrarse para esos años.")
        st.plotly_chart(bar_sector(bp), use_container_width=True, theme=None, config={"displayModeBar": False})
        titulo("CONTRIBUCIÓN POR ORIGEN DE APORTE")
        st.plotly_chart(pie_origen(bp), use_container_width=True, theme=None, config={"displayModeBar": False})


    # -------------------- PROYECCIÓN DE CIERRE DE AÑO --------------------
    titulo(f"PROYECCIÓN DE CIERRE DE AÑO {int(anio)}" + (f" (corte: {MESES[mes_corte]})" if mes_corte else ""))
    if proy is None:
        st.info("No hay datos suficientes para proyectar el cierre con los filtros seleccionados.")
    else:
        pa=proy["ppto_anual"]
        f1=st.columns(4)
        with f1[0]: card("INGRESO PROYECTADO — PROMEDIO (S/)",money(proy["prom"]))
        with f1[1]: card("INGRESO PROYECTADO — CONSERVADOR (S/)",money(proy["cons"]))
        with f1[2]: card("% CUMPLIMIENTO PROYECTADO — PROMEDIO",pct(proy["prom"]/pa*100) if pa else "N/D")
        with f1[3]: card("% CUMPLIMIENTO PROYECTADO — CONSERVADOR",pct(proy["cons"]/pa*100) if pa else "N/D")
        f2=st.columns(4)
        with f2[0]: card("BRECHA PROYECTADA CIERRE — PROMEDIO (S/)",money(proy["prom"]-pa),AZUL2)
        with f2[1]: card("BRECHA PROYECTADA CIERRE — CONSERVADOR (S/)",money(proy["cons"]-pa),AZUL2)
        with f2[2]: card("AVANCE RECAUDACIÓN VS. AVANCE TEMPORAL","N/D" if avance_pp is None else f"{avance_pp:+.1f} p.p.",AZUL2)
        with f2[3]: card(f"VARIACIÓN INTERANUAL VS {int(anio)-1} (S/)","N/D" if var_interanual is None else money(var_interanual),AZUL2)
        f3=st.columns(4)
        with f3[0]: card("SECTOR CON MAYOR CRECIMIENTO",sector_top or "N/D",AZUL2)
        with f3[1]: card("CRECIMIENTO INTERANUAL DEL SECTOR","N/D" if sector_crec is None else pct(sector_crec),AZUL2)
        with f3[2]: card("BRECHA PRESUPUESTAL ACTUAL (S/)",money(saldo_vs_acum),AZUL2)
        with f3[3]: card("PRESUPUESTO ANUAL (S/)",money(pa),AZUL2)
        _h=", ".join(map(str,proy["anios_hist"])) or "sin historia"
        st.caption(
            "Ambos escenarios parten del % de cumplimiento acumulado de Contribución Regular al mes de corte y le aplican "
            f"cómo cambió ese % entre ese mes y diciembre en los años {_h} (Promedio = cambio típico; Conservador = peor "
            "cambio observado). Recuperación de Deuda se proyecta aparte con el ritmo mensual actual. "
            "El escenario Conservador es el más prudente para comprometer cifras."
            + (f" El presupuesto anual incluye S/ {money(dn_anual)} de Dirección Nacional, que no forma parte del presupuesto acumulado." if dn_anual else "")
        )

    # -------------------- CUMPLIMIENTO ACUMULADO POR ZONAL --------------------
    titulo(f"CUMPLIMIENTO ACUMULADO POR ZONAL {int(anio)}")
    cz = actual_anio.groupby("ZONAL", as_index=False)[["INGRESOS","PRESUPUESTOS"]].sum().sort_values("ZONAL")
    if cz.empty:
        st.info("No existen datos para los filtros seleccionados.")
    else:
        def _fila_cumpl(nombre, ing, ppto, total=False):
            p = ing / ppto * 100 if ppto else None
            if p is None:
                icono, fondo, txt = "", "#fff", "N/D"
            else:
                if p >= 100:
                    icono, t, tope = "<span style='color:#1E8E3E;'>&#9650;</span>", min((p - 100) / 20, 1), (99, 190, 123)
                elif p >= 95:
                    icono, t, tope = "<span style='color:#E0A100;'>&#9654;</span>", (100 - p) / 15, (248, 105, 107)
                else:
                    icono, t, tope = "<span style='color:#D93025;'>&#9660;</span>", min((100 - p) / 15, 1), (248, 105, 107)
                fondo = "rgb(%d,%d,%d)" % tuple(round(255 + (c - 255) * t) for c in tope)
                txt = f"{p:.1f}%"
            base_td = "padding:6px 12px;border:1px solid #999;font-size:13px;color:#111;white-space:nowrap;"
            peso = "font-weight:800;" if total else "font-weight:600;"
            bg = "background:#EDE8D8;" if total else "background:#fff;"
            borde = f"border-top:2px solid {AZUL2};" if total else ""
            return (f"<tr><td style='{base_td}{peso}{bg}{borde}'>{nombre}</td>"
                    f"<td style='{base_td}{peso}{bg}{borde}text-align:right;'>{money(round(ing))[:-3]}</td>"
                    f"<td style='{base_td}{peso}{bg}{borde}text-align:right;'>{money(round(ppto))[:-3]}</td>"
                    f"<td style='{base_td}font-weight:800;{borde}background:{fondo};'>"
                    f"<div style='display:flex;justify-content:space-between;gap:14px;'><span>{icono}</span><span>{txt}</span></div></td></tr>")
        _th = f"background:{AZUL2};color:#fff;font-weight:800;font-size:12px;padding:8px 12px;border:1px solid {AZUL2};"
        _filas = "".join(_fila_cumpl(NOMBRE_ZONAL.get(r.ZONAL, r.ZONAL.title()), r.INGRESOS, r.PRESUPUESTOS) for r in cz.itertuples())
        _filas += _fila_cumpl("TOTAL" if zonals else "TOTAL NACIONAL", cz["INGRESOS"].sum(), cz["PRESUPUESTOS"].sum(), total=True)
        st.markdown(
            "<table style='border-collapse:collapse;width:100%;margin-bottom:6px;'><thead><tr>"
            f"<th style='{_th}text-align:left;'>Zonal</th><th style='{_th}text-align:right;'>Ingreso Acum. (S/)</th>"
            f"<th style='{_th}text-align:right;'>Presupuesto Acum. (S/)</th><th style='{_th}text-align:right;'>% Cumplimiento</th>"
            f"</tr></thead><tbody>{_filas}</tbody></table>", unsafe_allow_html=True)
        st.caption("Ingreso y presupuesto de los meses seleccionados (contribución regular + recuperación de deuda, según el filtro TIPO). "
                   "▲ cumple o supera el presupuesto · ▶ entre 95% y 100% · ▼ por debajo de 95%.")

    # -------------------- CONTROL DE RECAUDACIÓN POR ZONAL: CR vs RD --------------------
    titulo(f"CONTROL DE RECAUDACIÓN POR ZONAL {int(anio)} — CONTRIBUCIÓN REGULAR vs RECUPERACIÓN DE DEUDA")
    if actual_anio.empty:
        st.info("No existen datos para los filtros seleccionados.")
    else:
        _es_cr = actual_anio["TIPO_NORM"] == "CONTRIBUCION REGULAR"
        ctl = pd.DataFrame({
            "CR": actual_anio[_es_cr].groupby("ZONAL")["INGRESOS"].sum(),
            "RD": actual_anio[~_es_cr].groupby("ZONAL")["INGRESOS"].sum(),
            "PPTO": actual_anio.groupby("ZONAL")["PRESUPUESTOS"].sum(),
        }).fillna(0).sort_index()
        ctl["TOTAL"] = ctl["CR"] + ctl["RD"]
        _max_total = ctl["TOTAL"].max() or 1

        def _ent(v):
            return money(round(v))[:-3]

        def _fila_ctl(nombre, cr, rd, ppto, total=False):
            tot = cr + rd
            p = tot / ppto * 100 if ppto else None
            if p is None:
                icono, fondo, txt = "", "#fff", "N/D"
            else:
                if p >= 100:
                    icono, t, tope = "<span style='color:#1E8E3E;'>&#9650;</span>", min((p - 100) / 20, 1), (99, 190, 123)
                elif p >= 95:
                    icono, t, tope = "<span style='color:#E0A100;'>&#9654;</span>", (100 - p) / 15, (248, 105, 107)
                else:
                    icono, t, tope = "<span style='color:#D93025;'>&#9660;</span>", min((100 - p) / 15, 1), (248, 105, 107)
                fondo = "rgb(%d,%d,%d)" % tuple(round(255 + (c - 255) * t) for c in tope)
                txt = f"{p:.1f}%"
            bg = "#EDE8D8" if total else "#fff"
            e = ("padding:6px 12px;border:1px solid #999;font-size:13px;color:#111;white-space:nowrap;"
                 + ("font-weight:800;" if total else "font-weight:600;")
                 + (f"border-top:2px solid {AZUL2};" if total else ""))
            d = e + "text-align:right;"
            pc_cr = f"{cr / tot * 100:.1f}%" if tot else "N/D"
            pc_rd = f"{rd / tot * 100:.1f}%" if tot else "N/D"
            if total:
                barra = f"background:{bg};"
            else:
                w = max(0, min(100, tot / _max_total * 100))
                barra = f"background:linear-gradient(90deg,#9DB7E8 {w:.1f}%,#fff {w:.1f}%);"
            return (f"<tr><td style='{e}background:{bg};'>{nombre}</td>"
                    f"<td style='{d}background:{bg};'>{_ent(cr)}</td>"
                    f"<td style='{d}background:{bg};'>{_ent(rd)}</td>"
                    f"<td style='{d}{barra}'>{_ent(tot)}</td>"
                    f"<td style='{d}background:{bg};'>{pc_cr}</td>"
                    f"<td style='{d}background:{bg};'>{pc_rd}</td>"
                    f"<td style='{d}background:{bg};'>{_ent(ppto)}</td>"
                    f"<td style='{e}font-weight:800;background:{fondo};'>"
                    f"<div style='display:flex;justify-content:space-between;gap:14px;'><span>{icono}</span><span>{txt}</span></div></td>"
                    f"<td style='{d}background:{bg};'>{_ent(tot - ppto)}</td></tr>")

        _thc = f"background:{AZUL2};color:#fff;font-weight:800;font-size:12px;padding:8px 12px;border:1px solid {AZUL2};text-align:right;"
        _enc = ["Contribución Regular (S/)","Recuperación de Deuda (S/)","Total (S/)","% Contribución Regular",
                "% Recuperación de Deuda","Presupuesto Total (S/)","% Cumplimiento","Sup/Def Total"]
        _fc = "".join(_fila_ctl(NOMBRE_ZONAL.get(z, z.title()), r.CR, r.RD, r.PPTO) for z, r in ctl.iterrows())
        _fc += _fila_ctl("TOTAL" if zonals else "TOTAL NACIONAL", ctl["CR"].sum(), ctl["RD"].sum(), ctl["PPTO"].sum(), total=True)
        st.markdown(
            "<div style='overflow-x:auto;'><table style='border-collapse:collapse;width:100%;margin-bottom:6px;'><thead><tr>"
            f"<th style='{_thc}text-align:left;'>Zonal</th>" + "".join(f"<th style='{_thc}'>{c}</th>" for c in _enc)
            + f"</tr></thead><tbody>{_fc}</tbody></table></div>", unsafe_allow_html=True)
        st.caption("Acumulado de los meses seleccionados. Sup/Def Total = total recaudado menos presupuesto total.")

    # -------------------- TABLA RESUMEN --------------------
    with st.expander("Ver detalle de cálculos"):
        st.write(f"**Año de análisis:** {anio}")
        st.write(f"**Acumulado ejecutado:** S/ {money(actual_acum)}")
        st.write(f"**Presupuesto acumulado:** S/ {money(budget_acum)}")
        st.write(f"**Presupuesto anual:** S/ {money(budget_full)}")
        st.write(f"**Brecha (ejecutado - presupuesto acumulado):** S/ {money(saldo_vs_acum)}")
        st.write(f"**Saldo por recaudar del presupuesto anual:** S/ {money(saldo_pendiente)}")
        st.write(f"**Variación acumulada vs año anterior:** {'N/D' if vs_anterior is None else pct(vs_anterior)}")
        st.write(f"**Variación acumulada vs 2023:** {'N/D' if vs_2023 is None else pct(vs_2023)}")
        st.write(f"**Desglose por TIPO ({anio}, meses y zonal seleccionados):**")
        tabla_html(actual_anio.groupby("TIPO",as_index=False)[["INGRESOS","PRESUPUESTOS"]].sum())
        sin_mes=pres[(pres["AÑO"]==anio)&(pres["MES_NUM"].isna())]
        if not sin_mes.empty:
            st.warning(f"Hay {len(sin_mes)} filas de PRESUPUESTO {anio} con MES no reconocido (quedan fuera del acumulado): "
                       f"S/ {money(sin_mes['PRESUPUESTOS'].sum())} de presupuesto.")
        tabla_html(annual, enteros=["AÑO"])


# -------------------- HOJA: DETALLE DE PAGOS --------------------
with tab_det:
    titulo("DETALLE DE PAGOS")
    st.caption("Registros de la hoja BASE según los filtros de la barra lateral (año, mes, zonal, tipo, origen y sector).")
    buscar = st.text_input("Buscar por RUC o contribuyente", "")
    det = bp
    if buscar.strip():
        q = buscar.strip().upper()
        det = det[det["CONTRIBUYENTES"].str.upper().str.contains(q, regex=False)
                  | det["RUC"].astype(str).str.contains(q, regex=False)]
    cols = ["BANCO","FECHA","REGISTRO","RUC","TIPO","CONTRIBUYENTES","RDD","PERIODO","AÑO DEC.",
            "CONTRIBUCION","INTERES","COSTAS","TOTAL","MES CONT.","AÑO","ZONAL","ORIGEN","SECTOR","DIVISION","ESTADO"]
    cols = [c for c in cols if c in det.columns]

    # Controles de orden y paginación (la tabla se muestra por páginas)
    FILAS_PAG = 200
    o1,o2,o3 = st.columns([2,1,1])
    with o1: orden = st.selectbox("Ordenar por", cols, index=cols.index("FECHA"), key="det_orden")
    with o2: sentido = st.selectbox("Sentido", ["Descendente","Ascendente"], key="det_sentido")
    det = det.sort_values(orden, ascending=(sentido=="Ascendente"), kind="stable", na_position="last")
    n_pag = max(1, -(-len(det)//FILAS_PAG))
    with o3: pag = st.number_input(f"Página (de {n_pag:,})".replace(",", "."), min_value=1, max_value=n_pag, value=1, step=1, key="det_pag")

    k1,k2,k3,k4 = st.columns(4)
    with k1: card("REGISTROS", f"{len(det):,}".replace(",", "."))
    with k2: card("CONTRIBUYENTES", f"{det['RUC'].nunique():,}".replace(",", "."))
    with k3: card("CONTRIBUCIÓN (S/)", money(det["CONTRIBUCION"].sum()), AZUL2)
    with k4: card("TOTAL PAGADO (S/)", money(det["TOTAL"].sum()), AZUL2)

    vista = det[cols]
    pagina = vista.iloc[(pag-1)*FILAS_PAG : pag*FILAS_PAG].copy()
    for c in ["FECHA"]:
        pagina[c] = pagina[c].dt.strftime("%d-%m-%y")
    for c in ["PERIODO","MES CONT."]:
        _d = pd.to_datetime(pagina[c], errors="coerce")
        pagina[c] = _d.map(lambda d: "" if pd.isna(d) else f"{MESES[d.month].lower()}-{d.year % 100:02d}")
    tabla_html(pagina, enteros=["REGISTRO","RUC","AÑO","AÑO DEC."], derecha=["FECHA","PERIODO","MES CONT."], alto=620)
    st.caption(f"Mostrando {len(pagina):,} de {len(vista):,} registros. La descarga en Excel incluye todos.".replace(",", "."))

    # El Excel se genera a pedido porque la tabla puede tener más de 100 mil filas
    if st.button("Preparar descarga en Excel", key="prep_det"):
        with st.spinner("Generando archivo..."):
            st.session_state["xlsx_det"] = a_excel(
                vista, "Detalle de pagos",
                cols_monto=["CONTRIBUCION","INTERES","COSTAS","TOTAL"],
                cols_fecha=["FECHA","PERIODO","MES CONT."],
                cols_entero=["REGISTRO","RUC","AÑO","AÑO DEC."],
            )
            st.session_state["xlsx_det_filas"] = len(vista)
    if "xlsx_det" in st.session_state:
        st.download_button(
            f"Descargar Excel ({st.session_state['xlsx_det_filas']:,} filas)".replace(",", "."),
            st.session_state["xlsx_det"],
            file_name="detalle_pagos.xlsx", mime=XLSX_MIME, key="dl_det",
        )
        st.caption("Si cambias los filtros, vuelve a pulsar «Preparar descarga en Excel» para actualizar el archivo.")

# -------------------- HOJA: PAGOS POR CONTRIBUYENTE --------------------
with tab_contrib:
    titulo("PAGOS POR CONTRIBUYENTE (TOTAL POR MES CONTABLE)")
    st.caption("Suma de TOTAL por registro y mes contable, según los filtros de la barra lateral. "
               "Una celda vacía significa que no hubo pago en ese mes.")
    buscar2 = st.text_input("Buscar por registro, RUC o contribuyente", "", key="buscar_contrib")
    dc = bp
    if buscar2.strip():
        q = buscar2.strip().upper()
        dc = dc[dc["CONTRIBUYENTES"].str.upper().str.contains(q, regex=False)
                | dc["RUC"].astype(str).str.contains(q, regex=False)
                | dc["REGISTRO"].astype(str).str.contains(q, regex=False)]

    if dc.empty:
        st.info("No existen datos para los filtros seleccionados.")
    else:
        pv = dc.pivot_table(index=["REGISTRO","CONTRIBUYENTES"], columns="MES_CONT_DT",
                            values="TOTAL", aggfunc="sum", sort=True)
        pv = pv.where(pv != 0)   # meses sin pago quedan vacíos
        pv.columns = [f"{MESES[c.month].lower()}-{c.year % 100:02d}" for c in pv.columns]
        meses_cols = list(pv.columns)
        pv["Total general"] = pv[meses_cols].sum(axis=1)
        pv = pv.reset_index()

        t1,t2,t3 = st.columns(3)
        with t1: card("REGISTROS", f"{len(pv):,}".replace(",", "."))
        with t2: card("MESES CON PAGOS", str(len(meses_cols)))
        with t3: card("TOTAL GENERAL (S/)", money(pv["Total general"].sum()), AZUL2)

        # Controles de orden y paginación
        FILAS_PV = 200
        p1,p2,p3 = st.columns([2,1,1])
        with p1: orden_pv = st.selectbox("Ordenar por", list(pv.columns), index=0, key="pv_orden")
        with p2: sentido_pv = st.selectbox("Sentido", ["Ascendente","Descendente"], key="pv_sentido")
        pv = pv.sort_values(orden_pv, ascending=(sentido_pv=="Ascendente"), kind="stable", na_position="last")
        n_pv = max(1, -(-len(pv)//FILAS_PV))
        with p3: pag_pv = st.number_input(f"Página (de {n_pv:,})".replace(",", "."), min_value=1, max_value=n_pv, value=1, step=1, key="pv_pag")
        pag_df = pv.iloc[(pag_pv-1)*FILAS_PV : pag_pv*FILAS_PV]
        tabla_html(pag_df, enteros=["REGISTRO"], alto=620, fijas={"REGISTRO": 110, "CONTRIBUYENTES": 300})
        st.caption(f"Mostrando {len(pag_df):,} de {len(pv):,} registros. La descarga en Excel incluye todos.".replace(",", "."))

        with st.expander("Totales por mes"):
            tot = pv[meses_cols + ["Total general"]].sum().to_frame("TOTAL (S/)").T
            tabla_html(tot)

        st.download_button(
            "Descargar tabla (Excel)",
            a_excel(pv, "Pagos por contribuyente",
                    cols_monto=meses_cols + ["Total general"], cols_entero=["REGISTRO"]),
            file_name="pagos_por_contribuyente.xlsx", mime=XLSX_MIME, key="dl_contrib",
        )

st.caption("Fuente: BASE_CONTRI(1).xlsx | BASE para contribución por sector/origen y PRESUPUESTO para ingresos, presupuesto y ejecución.")
