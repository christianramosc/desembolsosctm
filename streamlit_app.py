# -*- coding: utf-8 -*-
"""
streamlit_app.py · @analyticsdotpy
Desglose por cliente del Reporte de Operaciones (ReportePagos_*.pdf) de Autoexpress Inbursa.

- Sube uno o varios PDFs y ve el desglose con gráficas, conciliación y tarjetas por cliente.
- Descarga el PDF visual (el mismo que genera el script de Colab) y un CSV.
- Incluye un apartado de donaciones (SPEI y pago con tarjeta) configurable desde Secrets.

Ejecutar en local:  streamlit run streamlit_app.py
"""

import io
import zipfile
from decimal import Decimal
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

import desglose_reporte_pagos as drp

st.set_page_config(
    page_title="Desglose de Reporte de Pagos · @analyticsdotpy",
    page_icon="📊",
    layout="wide",
)

CARPETA = Path(__file__).parent
EJEMPLO = CARPETA / "ejemplo" / "ReportePagos_EJEMPLO_10desembolsos.pdf"
INSTAGRAM_URL = "https://www.instagram.com/analyticsdotpy/"

COLORES = {k: "#" + v.hexval()[2:].upper() for k, v in drp.COLOR_CONCEPTO.items()}
NOMBRES = drp.NOMBRE_CONCEPTO
EMOJI = {"neto": "🏦", "seguro": "🛡️", "gap_financiado": "🚗", "gap_contado": "🚙",
         "comision": "％", "subsidio": "🏷️"}


# =============================================================================
# Estilos
# =============================================================================
st.markdown(
    """
    <style>
      .block-container {padding-top: 2rem;}
      [data-testid='stMetricValue'] {font-size: clamp(1.15rem, 1.9vw, 1.75rem); font-weight: 700;}
      .hero {background:#1B2446; border-radius:16px; padding:22px 26px; display:flex;
             align-items:center; gap:18px; margin-bottom:.6rem;}
      .hero .logo {background:#FFD43B; border-radius:14px; width:58px; height:58px; flex-shrink:0;
                   display:flex; align-items:center; justify-content:center; font-size:30px;}
      .hero h1 {color:#FFFFFF; font-size:1.65rem; margin:0; padding:0; line-height:1.2;}
      .hero p {color:#C9D1E6; margin:.25rem 0 0 0; font-size:.95rem;}
      .hero .handle {margin-left:auto; font-weight:700; color:#FFFFFF; background:#263159;
                     padding:7px 14px; border-radius:999px; white-space:nowrap; text-decoration:none;}
      .hero .handle span {color:#FFD43B;}
      .marca span {color:#C99A06;}
      .marca {font-weight:700; color:#1B2446;}
      .ubicacion {color:#6B7280; font-size:.95rem; margin:.2rem 0 .8rem 0;}
      .pie {color:#6B7280; font-size:.85rem; text-align:center; margin-top:2rem;}
      .clabe-ok {color:#079455; font-size:.8rem;}
      @media (max-width: 640px) {
        .hero {flex-wrap:wrap; padding:18px;}
        .hero .handle {margin-left:0;}
        .hero h1 {font-size:1.35rem;}
      }
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# Donaciones (los datos viven en Secrets, no en el código)
# =============================================================================
def leer_donaciones():
    try:
        return {k: v for k, v in st.secrets["donaciones"].items()}
    except Exception:
        return {}


def clabe_valida(clabe: str) -> bool:
    """Verifica el dígito de control de una CLABE de 18 dígitos (pesos 3, 7, 1)."""
    if len(clabe) != 18 or not clabe.isdigit():
        return False
    pesos = [3, 7, 1] * 6
    suma = sum((int(d) * p) % 10 for d, p in zip(clabe[:17], pesos))
    return (10 - suma % 10) % 10 == int(clabe[17])


def botones_tarjeta(cfg):
    """Devuelve [(etiqueta, url)] con los servicios de pago con tarjeta configurados."""
    botones = []
    mp = cfg.get("mercado_pago")
    if isinstance(mp, str) and mp:
        botones.append(("💳 Donar con Mercado Pago", mp))
    elif mp:  # varios montos: {"$50" = "url", "$100" = "url"}
        botones += [(f"💳 Donar {monto}".replace("$", "\\$"), url) for monto, url in mp.items() if url]
    for clave, etiqueta in (("stripe", "💳 Donar con tarjeta (Stripe)"),
                            ("paypal", "🅿️ Donar con PayPal"),
                            ("kofi", "☕ Invítame un café (Ko-fi)")):
        if cfg.get(clave):
            botones.append((etiqueta, cfg[clave]))
    return botones


def seccion_donaciones():
    cfg = leer_donaciones()
    st.markdown("### 💛 Apoya este proyecto")
    st.write(
        "Esta herramienta es **gratuita**. Si te ahorró tiempo y quieres apoyar para que siga "
        "creciendo con nuevas funciones, puedes hacer una donación voluntaria del monto que quieras."
    )
    clabe = str(cfg.get("clabe", "")).replace(" ", "")
    botones = botones_tarjeta(cfg)

    if not clabe and not botones:
        st.info("Aún no hay métodos de donación configurados. Agrega tu CLABE y/o tus links de pago en "
                "**Settings → Secrets** de Streamlit.", icon="⚙️")
        return

    col_spei, col_tarjeta = st.columns(2, gap="large")
    with col_spei:
        with st.container(border=True):
            st.markdown("#### 🏦 Transferencia SPEI")
            st.caption("Desde la app de tu banco; normalmente sin comisión.")
            if clabe:
                if cfg.get("beneficiario"):
                    st.markdown(f"**Beneficiario:** {cfg['beneficiario']}")
                if cfg.get("banco"):
                    st.markdown(f"**Banco:** {cfg['banco']}")
                st.markdown("**CLABE** (toca el ícono para copiar):")
                st.code(clabe, language=None)
                if not clabe_valida(clabe):
                    st.warning("La CLABE configurada no pasa la validación de 18 dígitos. Revísala en Secrets.")
                st.caption("Concepto sugerido: *Donación analyticsdotpy*")
            else:
                st.caption("No disponible por ahora.")
    with col_tarjeta:
        with st.container(border=True):
            st.markdown("#### 💳 Tarjeta de crédito o débito")
            st.caption("Pago seguro en la página del procesador. Tus datos de tarjeta nunca pasan por esta app.")
            if botones:
                for etiqueta, url in botones:
                    st.link_button(etiqueta, url, width="stretch")
            else:
                st.caption("No disponible por ahora.")
    st.markdown(f"¿Ideas o sugerencias? Escríbeme en Instagram: "
                f"[<span class='marca'>@analytics<span>dotpy</span></span>]({INSTAGRAM_URL})",
                unsafe_allow_html=True)


# =============================================================================
# Procesamiento (en memoria; no se guarda ningún archivo)
# =============================================================================
@st.cache_data(show_spinner=False, max_entries=50)
def procesar_pdf(nombre: str, contenido: bytes):
    datos = drp.calcular(drp.extraer_reporte(io.BytesIO(contenido)))
    salida = io.BytesIO()
    drp.generar_pdf_desglose(datos, salida, nombre_origen=nombre)
    return datos, salida.getvalue(), drp.nombre_salida(datos, nombre)


def f(x):
    return float(x)


def md(valor):
    """Monto formateado para texto Markdown (escapa el $ para que no se lea como fórmula)."""
    return drp.fmt(valor).replace("$", "\\$")


def fmt_tabla(v):
    return "—" if v is None or pd.isna(v) else drp.fmt(Decimal(str(round(v, 2))))


def tabla_resumen(datos):
    visibles = datos["calculo"]["retenciones_visibles"]
    filas = []
    for c in datos["clientes"]:
        fila = {"#": c["num"], "Folio": c["folio"], "Titular": c["titular"],
                "Monto financiado": f(c["financiado"])}
        for clave, _, _ in visibles:
            fila[NOMBRES[clave]] = f(c[clave])
        fila["Total retenciones"] = f(c["total_retenciones"])
        fila["Transferencia neta"] = f(c["neto"])
        fila["Llega a la agencia"] = 100 - f(c["pct_retenido"])
        filas.append(fila)
    return pd.DataFrame(filas)


def tabla_conciliacion(datos):
    calc = datos["calculo"]
    filas = []
    for (clave, _, _), fila in zip(drp.CONCEPTOS, calc["conciliacion"]):
        if clave != "financiado" and fila["suma"] == 0 and not fila["reportado"]:
            continue
        filas.append((fila["concepto"], fila["suma"], fila["reportado"]))
    if calc["cancelaciones_aplicadas"] != 0 or datos["totales"].get("cancelaciones"):
        filas.append(("Saldo de cancelaciones de créditos", calc["cancelaciones_aplicadas"],
                      datos["totales"].get("cancelaciones")))
    filas.append(("Total de transferencias bancarias", calc["total_calculado"], calc["reportado_transferencias"]))
    salida = []
    for concepto, suma, reportado in filas:
        if reportado is None:
            estado, dif = "⚪ Sin dato", None
        else:
            dif = f(suma - reportado)
            estado = "✅ Cuadra" if suma == reportado else "❌ Revisar"
        salida.append({"Concepto": concepto, "Suma del desglose": f(suma),
                       "Total del reporte": None if reportado is None else f(reportado),
                       "Diferencia": dif, "Estatus": estado})
    return pd.DataFrame(salida)


def grafica_dona(datos):
    calc = datos["calculo"]
    items = [("neto", f(calc["suma_netos"]))]
    for clave, _, _ in calc["retenciones_visibles"]:
        monto = -f(sum((c[clave] for c in datos["clientes"]), Decimal("0")))
        if monto > 0:
            items.append((clave, monto))
    base_total = f(calc["suma_financiado"]) or 1
    df = pd.DataFrame([{"Concepto": NOMBRES[k], "Monto": m, "Porcentaje": m / base_total} for k, m in items])
    escala = alt.Scale(domain=[NOMBRES[k] for k, _ in items], range=[COLORES[k] for k, _ in items])
    arco = alt.Chart(df).mark_arc(innerRadius=78, outerRadius=118, cornerRadius=3, padAngle=0.012).encode(
        theta=alt.Theta("Monto:Q", stack=True),
        color=alt.Color("Concepto:N", scale=escala, sort=None,
                        legend=alt.Legend(orient="bottom", columns=2, title=None, labelFontSize=12)),
        order=alt.Order("orden:Q"),
        tooltip=["Concepto", alt.Tooltip("Monto:Q", format="$,.2f"), alt.Tooltip("Porcentaje:Q", format=".1%")],
    ).transform_window(orden="row_number()")
    pct = items[0][1] / base_total
    centro = alt.Chart(pd.DataFrame({"t": [f"{pct:.1%}"]})).mark_text(
        fontSize=28, fontWeight="bold", color=COLORES["neto"], dy=-6).encode(text="t:N")
    sub = alt.Chart(pd.DataFrame({"t": ["llega a la agencia"]})).mark_text(
        fontSize=12, color="#6B7280", dy=18).encode(text="t:N")
    return (arco + centro + sub).properties(height=330)


def grafica_barras(datos):
    visibles = datos["calculo"]["retenciones_visibles"]
    claves = ["neto"] + [k for k, _, _ in visibles]
    filas, orden_clientes = [], []
    for c in datos["clientes"]:
        etiqueta = f"{c['num']} · {c['titular']}"
        orden_clientes.append(etiqueta)
        for i, clave in enumerate(claves):
            monto = f(c["neto"]) if clave == "neto" else -f(c[clave])
            if monto > 0:
                filas.append({"Cliente": etiqueta, "Concepto": NOMBRES[clave], "Monto": monto, "orden": i})
    df = pd.DataFrame(filas)
    escala = alt.Scale(domain=[NOMBRES[k] for k in claves], range=[COLORES[k] for k in claves])
    alto = max(160, 38 * len(orden_clientes))
    return alt.Chart(df).mark_bar(size=20, cornerRadius=3).encode(
        y=alt.Y("Cliente:N", sort=orden_clientes, title=None, axis=alt.Axis(labelLimit=220)),
        x=alt.X("Monto:Q", title=None, stack="zero", axis=alt.Axis(format="$,.0f", grid=True)),
        color=alt.Color("Concepto:N", scale=escala, legend=None),  # mismos colores que la dona
        order=alt.Order("orden:Q"),
        tooltip=["Cliente", "Concepto", alt.Tooltip("Monto:Q", format="$,.2f")],
    ).properties(height=alto)


def tarjetas_clientes(datos):
    visibles = datos["calculo"]["retenciones_visibles"]
    clientes = datos["clientes"]
    por_fila = 3
    for i in range(0, len(clientes), por_fila):
        cols = st.columns(por_fila)
        for col, c in zip(cols, clientes[i:i + por_fila]):
            with col.container(border=True):
                st.markdown(f"**👤 {c['titular']}**  \n"
                            f"<span style='color:#6B7280;font-size:.85rem'>#{c['num']} · Número de folio "
                            f"{c['folio']}</span>", unsafe_allow_html=True)
                lineas = [f"💵 Monto financiado: **{md(c['financiado'])}**"]
                for clave, _, _ in visibles:
                    if c[clave] != 0:
                        lineas.append(f"{EMOJI[clave]} {NOMBRES[clave]}: :red[{md(c[clave])}]")
                st.markdown("  \n".join(lineas))
                pct = 100 - f(c["pct_retenido"])
                st.progress(min(max(pct / 100, 0.0), 1.0), text=f"{pct:.1f}% llega a la agencia")
                st.markdown(f"🏦 Transferencia neta: :green[**{md(c['neto'])}**]")


def mostrar_reporte(nombre, datos, pdf_bytes, nombre_pdf, clave):
    calc = datos["calculo"]
    st.markdown(f"<div class='ubicacion'>📍 <b>{datos['pdv'] or 'PDV no identificado'}"
                f"{' - ' + datos['pdv_num'] if datos['pdv_num'] else ''}</b> &nbsp;·&nbsp; "
                f"📅 {datos['fecha_texto'] or 'Fecha no identificada'} &nbsp;·&nbsp; 📄 {nombre}</div>",
                unsafe_allow_html=True)

    if calc["reportado_transferencias"] is None:
        st.warning(f"No se encontró el Total de Transferencias Bancarias en el PDF. "
                   f"Total calculado: **{md(calc['total_calculado'])}**.", icon="⚠️")
    elif datos["cuadra"]:
        st.success(f"**El desglose cuadra.** La suma por cliente ({md(calc['total_calculado'])}) es igual "
                   f"al Total de Transferencias Bancarias del reporte.", icon="✅")
    else:
        st.error(f"**El desglose no cuadra.** Diferencia de {md(calc['diferencia_final'])}. "
                 f"Revisa la pestaña de conciliación.", icon="❌")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("💵 Total financiado", drp.fmt(calc["suma_financiado"]), border=True)
    m2.metric("🧾 Total retenciones", drp.fmt(calc["suma_retenciones"]), border=True)
    m3.metric("🏦 Transferencias", drp.fmt(calc["total_calculado"]), border=True)
    n = len(datos["clientes"])
    m4.metric("👥 Créditos", f"{n} {'cliente' if n == 1 else 'clientes'}", border=True)

    g1, g2 = st.columns([2, 3], gap="large")
    with g1:
        st.markdown("##### 🍩 ¿A dónde va el dinero financiado?")
        st.altair_chart(grafica_dona(datos), width="stretch")
    with g2:
        st.markdown("##### 📊 Composición por cliente")
        st.altair_chart(grafica_barras(datos), width="stretch")

    t_res, t_con, t_det = st.tabs(["📋 Resumen por cliente", "🔎 Conciliación", "👤 Detalle por cliente"])
    with t_res:
        df = tabla_resumen(datos)
        montos = [c for c in df.columns if c not in ("#", "Folio", "Titular", "Llega a la agencia")]
        st.dataframe(
            df.style.format(fmt_tabla, subset=montos),
            hide_index=True, width="stretch",
            column_config={"Llega a la agencia": st.column_config.ProgressColumn(
                "Llega a la agencia", min_value=0, max_value=100, format="%.1f%%")},
        )
    with t_con:
        dfc = tabla_conciliacion(datos)
        st.dataframe(dfc.style.format(fmt_tabla, subset=["Suma del desglose", "Total del reporte", "Diferencia"]),
                     hide_index=True, width="stretch")
    with t_det:
        tarjetas_clientes(datos)

    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("⬇️ Descargar PDF del desglose", pdf_bytes, file_name=nombre_pdf,
                       mime="application/pdf", type="primary", width="stretch", key=f"pdf_{clave}")
    csv = tabla_resumen(datos).to_csv(index=False).encode("utf-8-sig")
    d2.download_button("⬇️ Descargar CSV", csv, file_name=nombre_pdf.replace(".pdf", ".csv"),
                       mime="text/csv", width="stretch", key=f"csv_{clave}")


# =============================================================================
# Interfaz
# =============================================================================
st.markdown(
    f"""
    <div class="hero">
      <div class="logo">📊</div>
      <div>
        <h1>Desglose por cliente</h1>
        <p>Sube el Reporte de Operaciones de Autoexpress Inbursa y obtén el desglose de cada crédito,
           la conciliación contra el total de transferencias y un PDF listo para compartir.</p>
      </div>
      <a class="handle" href="{INSTAGRAM_URL}" target="_blank">@analytics<span>dotpy</span></a>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown(f"## <span class='marca'>@analytics<span>dotpy</span></span>", unsafe_allow_html=True)
    st.caption("Análisis de datos con Python")
    st.markdown("**Cómo usarla**\n1. Sube uno o varios `ReportePagos_*.pdf`.\n"
                "2. Revisa las gráficas, la conciliación y el detalle.\n3. Descarga el PDF o el CSV.")
    st.divider()
    st.markdown("💛 ¿Te sirvió? En la pestaña **Apoya el proyecto** puedes hacer una donación voluntaria.")
    st.link_button("📸 Sígueme en Instagram", INSTAGRAM_URL, width="stretch")

tab_app, tab_apoyo = st.tabs(["📄 Desglose", "💛 Apoya el proyecto"])

with tab_app:
    archivos = st.file_uploader("Sube uno o varios Reportes de Operaciones (PDF)", type=["pdf"],
                                accept_multiple_files=True)
    usar_ejemplo = False
    if not archivos and EJEMPLO.exists():
        st.caption("¿No tienes un reporte a la mano?")
        usar_ejemplo = st.toggle("Probar con un reporte de ejemplo (datos ficticios)")

    entradas = [(a.name, a.getvalue()) for a in archivos or []]
    if usar_ejemplo:
        entradas = [(EJEMPLO.name, EJEMPLO.read_bytes())]

    if not entradas:
        st.info("Tus archivos se procesan en memoria y no se guardan en ningún lado.", icon="🔒")
    else:
        resultados = []
        with st.spinner("Leyendo y desglosando..."):
            for nombre, contenido in entradas:
                try:
                    resultados.append((nombre, *procesar_pdf(nombre, contenido)))
                except Exception as e:
                    st.error(f"No se pudo procesar **{nombre}**: {e}", icon="❌")

        if len(resultados) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                for _, _, pdf_bytes, nombre_pdf in resultados:
                    z.writestr(nombre_pdf, pdf_bytes)
            st.download_button(f"⬇️ Descargar los {len(resultados)} PDFs (ZIP)", zip_buf.getvalue(),
                               file_name="Desgloses_ReportePagos.zip", mime="application/zip")
            pestañas = st.tabs([f"📄 {r[0]}" for r in resultados])
            for i, (pestaña, r) in enumerate(zip(pestañas, resultados)):
                with pestaña:
                    mostrar_reporte(*r, clave=i)
        elif resultados:
            mostrar_reporte(*resultados[0], clave=0)

with tab_apoyo:
    seccion_donaciones()

st.markdown(
    f"<div class='pie'>Hecho por <a href='{INSTAGRAM_URL}' target='_blank' style='text-decoration:none'>"
    f"<span class='marca'>@analytics<span>dotpy</span></span></a> · Herramienta independiente, "
    f"no oficial de Autoexpress Inbursa · Tus archivos no se guardan.</div>",
    unsafe_allow_html=True,
)
