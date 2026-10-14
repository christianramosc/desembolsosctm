# -*- coding: utf-8 -*-
"""
desglose_reporte_pagos.py
--------------------------------------------------------------------------
Desglose por cliente del "Reporte de Operaciones" (ReportePagos_*.pdf)
de Autoexpress Inbursa.

Qué hace:
  1. Lee el PDF del reporte de pagos (una o varias páginas).
  2. Extrae cada operación: número de folio, titular, monto financiado y retenciones.
  3. Calcula por cliente: total de retenciones y transferencia neta.
  4. Concilia contra los totales del reporte (concepto por concepto) y
     verifica que la suma de las transferencias por cliente sea igual al
     "Total de Transferencias Bancarias".
  5. Genera un PDF visual de desglose (carta vertical) con iconos, gráfica de dona,
     barras de composición por cliente y tarjetas, y lo descarga.

Uso en Google Colab:
  - Pega todo el script en una celda y ejecútala.
  - Sube uno o varios PDFs del reporte cuando te lo pida.
  - Se descarga un PDF de desglose por cada reporte.

Uso local (fuera de Colab):
  python desglose_reporte_pagos.py ReportePagos_xxx.pdf [otro.pdf ...]

Autor: @analyticsdotpy
--------------------------------------------------------------------------
"""

import os
import re
import sys
import subprocess
import unicodedata
from decimal import Decimal
from datetime import datetime
from xml.sax.saxutils import escape


# ==========================================================================
# 0. DEPENDENCIAS (se instalan solas si faltan)
# ==========================================================================
def _asegurar_paquetes():
    faltantes = []
    for modulo in ("pdfplumber", "reportlab"):
        try:
            __import__(modulo)
        except ImportError:
            faltantes.append(modulo)
    if faltantes:
        print(f"Instalando: {', '.join(faltantes)} ...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", *faltantes], check=True)


_asegurar_paquetes()

import pdfplumber  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER  # noqa: E402
from reportlab.lib.pagesizes import letter  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import inch  # noqa: E402
from reportlab.pdfgen import canvas as rl_canvas  # noqa: E402
from reportlab.pdfbase.pdfmetrics import stringWidth  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether,
)
from reportlab.platypus.flowables import HRFlowable, Flowable  # noqa: E402


# ==========================================================================
# 1. CONFIGURACIÓN
# ==========================================================================
# (clave interna, etiqueta para el desglose, etiqueta del total en el PDF de Inbursa)
CONCEPTOS = [
    ("financiado",     "Monto financiado",                "Total de Financiaciones"),
    ("subsidio",       "Retención por subsidio",          "Total de Retenciones por Subsidio"),
    ("seguro",         "Retención seguro",                "Total de Retenciones Seguro"),
    ("gap_contado",    "Retención GAP contado",           "Total de Retenciones de GAP Contado"),
    ("gap_financiado", "Retención GAP financiado",        "Total de Retenciones de GAP Financiado"),
    ("comision",       "Retención comisión por apertura", "Total de Retenciones Comisión por Apertura"),
]
RETENCIONES = [c for c in CONCEPTOS if c[0] != "financiado"]
# Etiquetas cortas para las tarjetas por cliente
ETIQ_CORTA = {"subsidio": "por subsidio", "seguro": "seguro", "gap_contado": "GAP contado",
              "gap_financiado": "GAP financiado", "comision": "Comisión por apertura"}
ETIQ_CANCELACIONES = "Saldo de Cancelaciones de Créditos"
ETIQ_TRANSFERENCIAS = "Total de Transferencias Bancarias"

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

# Paleta
NAVY = colors.HexColor("#1B2446")
NAVY_SUAVE = colors.HexColor("#C9D1E6")
GRIS_FONDO = colors.HexColor("#F3F5F9")
GRIS_LINEA = colors.HexColor("#DDE2EC")
GRIS_TEXTO = colors.HexColor("#6B7280")
TEXTO = colors.HexColor("#1F2937")
ROJO = colors.HexColor("#B42318")
VERDE = colors.HexColor("#067647")
VERDE_FONDO = colors.HexColor("#E7F6EE")
ROJO_FONDO = colors.HexColor("#FDECEA")

# Marca personal: aparece en el encabezado, en el pie de cada página y en los metadatos del PDF
MARCA_HANDLE = "@analyticsdotpy"
NAVY_MARCA = colors.HexColor("#263159")
AMARILLO_PY = colors.HexColor("#FFD43B")
AZUL_PY = colors.HexColor("#3776AB")


# ==========================================================================
# 2. UTILIDADES
# ==========================================================================
NUM = r"-?[\d,]+\.\d{2}"
RE_FILA = re.compile(
    rf"^(\d{{3}})\s+(\S+)\s+(.+?)\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s*$"
)
RE_TOTAL = re.compile(rf"^((?:Total|Saldo)\s+de\s+.+?)\s+({NUM})\s*$", re.IGNORECASE)
RE_ENCABEZADO = re.compile(
    r"^(?P<pdv>.+?)\s+-\s+(?P<num>\d+)\s+.*?,\s*a\s+(?P<fecha>\d{1,2}\s+de\s+\w+\s+de\s+\d{4})",
    re.IGNORECASE,
)


def a_decimal(texto):
    return Decimal(texto.replace(",", "").strip())


def normalizar(texto):
    sin_acentos = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", sin_acentos).strip().lower()


def fmt(valor, con_signo_pesos=True):
    """282262 -> $282,262.00 ; -15060.07 -> -$15,060.07"""
    valor = Decimal(valor)
    txt = f"{abs(valor):,.2f}"
    if con_signo_pesos:
        txt = "$" + txt
    return ("-" if valor < 0 else "") + txt


def fecha_desde_texto(texto):
    m = re.match(r"(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})", texto.strip(), re.IGNORECASE)
    if not m:
        return None
    mes = MESES.get(normalizar(m.group(2)))
    if not mes:
        return None
    return datetime(int(m.group(3)), mes, int(m.group(1)))


# ==========================================================================
# 3. LECTURA DEL REPORTE
# ==========================================================================
def extraer_reporte(origen):
    """origen: ruta o archivo tipo bytes (BytesIO). Devuelve un dict con todo."""
    lineas = []
    with pdfplumber.open(origen) as pdf:
        for pagina in pdf.pages:
            lineas.extend((pagina.extract_text() or "").splitlines())

    mapa_totales = {normalizar(etq): clave for clave, _, etq in CONCEPTOS}
    mapa_totales[normalizar(ETIQ_CANCELACIONES)] = "cancelaciones"
    mapa_totales[normalizar(ETIQ_TRANSFERENCIAS)] = "transferencias"

    datos = {"pdv": None, "pdv_num": None, "fecha_texto": None, "fecha": None,
             "clientes": [], "totales": {}}

    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue

        m = RE_FILA.match(linea)
        if m:
            num, folio, titular, *valores = m.groups()
            cliente = {"num": num, "folio": folio, "titular": titular.strip()}
            for (clave, _, _), valor in zip(CONCEPTOS, valores):
                cliente[clave] = a_decimal(valor)
            datos["clientes"].append(cliente)
            continue

        m = RE_TOTAL.match(linea)
        if m:
            clave = mapa_totales.get(normalizar(m.group(1)))
            if clave:
                datos["totales"][clave] = a_decimal(m.group(2))
            continue

        if datos["pdv"] is None:
            m = RE_ENCABEZADO.match(linea)
            if m:
                datos["pdv"] = m.group("pdv").strip()
                datos["pdv_num"] = m.group("num")
                datos["fecha_texto"] = m.group("fecha")
                datos["fecha"] = fecha_desde_texto(m.group("fecha"))

    if not datos["clientes"]:
        raise ValueError("No se encontró ninguna operación en el PDF. "
                         "¿Es un Reporte de Operaciones de Autoexpress Inbursa?")
    return datos


# ==========================================================================
# 4. CÁLCULOS Y CONCILIACIÓN
# ==========================================================================
def calcular(datos):
    clientes = datos["clientes"]
    for c in clientes:
        c["total_retenciones"] = sum((c[k] for k, _, _ in RETENCIONES), Decimal("0"))
        c["neto"] = c["financiado"] + c["total_retenciones"]
        c["pct_retenido"] = (-c["total_retenciones"] / c["financiado"] * 100) if c["financiado"] else Decimal("0")

    totales = datos["totales"]

    # Conciliación concepto por concepto
    conciliacion = []
    for clave, etiqueta, _ in CONCEPTOS:
        suma = sum((c[clave] for c in clientes), Decimal("0"))
        reportado = totales.get(clave)
        conciliacion.append({
            "concepto": etiqueta, "suma": suma, "reportado": reportado,
            "diferencia": (suma - reportado) if reportado is not None else None,
        })

    suma_netos = sum((c["neto"] for c in clientes), Decimal("0"))
    cancelaciones = totales.get("cancelaciones", Decimal("0"))
    reportado_transf = totales.get("transferencias")

    # Las cancelaciones no vienen por cliente; se aplican al total.
    # Se asume que vienen con signo (negativas restan), igual que las retenciones.
    total_calculado = suma_netos + cancelaciones
    if (reportado_transf is not None and cancelaciones != 0
            and total_calculado != reportado_transf
            and suma_netos - cancelaciones == reportado_transf):
        total_calculado = suma_netos - cancelaciones
        cancelaciones = -cancelaciones  # se muestra con el signo con que se aplicó

    datos["calculo"] = {
        "suma_financiado": sum((c["financiado"] for c in clientes), Decimal("0")),
        "suma_retenciones": sum((c["total_retenciones"] for c in clientes), Decimal("0")),
        "suma_netos": suma_netos,
        "cancelaciones_aplicadas": cancelaciones,
        "total_calculado": total_calculado,
        "reportado_transferencias": reportado_transf,
        "diferencia_final": (total_calculado - reportado_transf) if reportado_transf is not None else None,
        "conciliacion": conciliacion,
        # Retenciones con movimiento en al menos un cliente (las demás no se muestran en las tarjetas)
        "retenciones_visibles": [r for r in RETENCIONES if any(c[r[0]] != 0 for c in clientes)],
    }
    datos["cuadra"] = datos["calculo"]["diferencia_final"] == 0
    return datos


# ==========================================================================
# 5. PDF DE DESGLOSE (diseño visual con iconos y gráficas)
# ==========================================================================
# Color, icono y nombre de cada concepto: se usan igual en la dona, las barras y las tarjetas
COLOR_CONCEPTO = {
    "neto": colors.HexColor("#079455"),
    "seguro": colors.HexColor("#3776AB"),
    "gap_financiado": colors.HexColor("#F79009"),
    "gap_contado": colors.HexColor("#C4320A"),
    "comision": colors.HexColor("#7A5AF8"),
    "subsidio": colors.HexColor("#0E9384"),
}
ICONO_CONCEPTO = {
    "financiado": "billete", "neto": "banco", "seguro": "escudo", "gap_financiado": "auto",
    "gap_contado": "auto", "comision": "porcentaje", "subsidio": "etiqueta",
}
NOMBRE_CONCEPTO = {
    "financiado": "Monto financiado", "neto": "Transferencia a la agencia", "seguro": "Seguro",
    "gap_financiado": "GAP financiado", "gap_contado": "GAP contado",
    "comision": "Comisión por apertura", "subsidio": "Subsidio",
}


def aclarar(color, f=0.85):
    """Mezcla un color con blanco (f=0 color original, f=1 blanco)."""
    return colors.Color(color.red + (1 - color.red) * f, color.green + (1 - color.green) * f,
                        color.blue + (1 - color.blue) * f)


# ---------------------------------------------------------------- iconos ---
def dibujar_icono(c, tipo, x, y, s, color=colors.white, detalle=NAVY):
    """Dibuja un icono vectorial en el cuadro (x, y, s, s). 'detalle' es el color de huecos/líneas."""
    def P(px, py):
        return x + px * s, y + py * s

    def poligono(puntos, fill=1, stroke=0):
        p = c.beginPath()
        p.moveTo(*P(*puntos[0]))
        for pt in puntos[1:]:
            p.lineTo(*P(*pt))
        p.close()
        c.drawPath(p, stroke=stroke, fill=fill)

    def linea(puntos, ancho):
        c.setLineWidth(s * ancho)
        p = c.beginPath()
        p.moveTo(*P(*puntos[0]))
        for pt in puntos[1:]:
            p.lineTo(*P(*pt))
        c.drawPath(p, stroke=1, fill=0)

    c.saveState()
    c.setFillColor(color)
    c.setStrokeColor(color)
    c.setLineCap(1)
    c.setLineJoin(1)

    if tipo == "billete":
        c.roundRect(*P(0.04, 0.22), s * 0.92, s * 0.56, s * 0.08, stroke=0, fill=1)
        c.setFillColor(detalle)
        c.circle(*P(0.5, 0.5), s * 0.14, stroke=0, fill=1)
        c.circle(*P(0.19, 0.5), s * 0.05, stroke=0, fill=1)
        c.circle(*P(0.81, 0.5), s * 0.05, stroke=0, fill=1)
    elif tipo == "escudo":
        p = c.beginPath()
        p.moveTo(*P(0.5, 0.96))
        p.lineTo(*P(0.88, 0.82))
        p.lineTo(*P(0.88, 0.52))
        p.curveTo(*P(0.88, 0.26), *P(0.7, 0.12), *P(0.5, 0.03))
        p.curveTo(*P(0.3, 0.12), *P(0.12, 0.26), *P(0.12, 0.52))
        p.lineTo(*P(0.12, 0.82))
        p.close()
        c.drawPath(p, stroke=0, fill=1)
        c.setStrokeColor(detalle)
        linea([(0.32, 0.5), (0.46, 0.36), (0.68, 0.64)], 0.09)
    elif tipo == "auto":
        poligono([(0.2, 0.5), (0.33, 0.78), (0.67, 0.78), (0.8, 0.5)])
        c.roundRect(*P(0.04, 0.26), s * 0.92, s * 0.3, s * 0.08, stroke=0, fill=1)
        c.setFillColor(detalle)
        poligono([(0.3, 0.55), (0.38, 0.71), (0.48, 0.71), (0.48, 0.55)])
        poligono([(0.52, 0.55), (0.52, 0.71), (0.62, 0.71), (0.7, 0.55)])
        for cx in (0.27, 0.73):
            c.setFillColor(detalle)
            c.circle(*P(cx, 0.24), s * 0.16, stroke=0, fill=1)
            c.setFillColor(color)
            c.circle(*P(cx, 0.24), s * 0.1, stroke=0, fill=1)
    elif tipo == "porcentaje":
        linea([(0.24, 0.16), (0.76, 0.84)], 0.1)
        c.setLineWidth(s * 0.09)
        c.circle(*P(0.27, 0.72), s * 0.13, stroke=1, fill=0)
        c.circle(*P(0.73, 0.28), s * 0.13, stroke=1, fill=0)
    elif tipo == "etiqueta":
        poligono([(0.1, 0.9), (0.52, 0.9), (0.92, 0.5), (0.52, 0.1), (0.1, 0.52)])
        c.setFillColor(detalle)
        c.circle(*P(0.3, 0.7), s * 0.09, stroke=0, fill=1)
    elif tipo == "banco":
        poligono([(0.04, 0.66), (0.5, 0.95), (0.96, 0.66)])
        for cx in (0.16, 0.43, 0.7):
            c.rect(*P(cx, 0.27), s * 0.14, s * 0.33, stroke=0, fill=1)
        c.rect(*P(0.06, 0.08), s * 0.88, s * 0.13, stroke=0, fill=1)
    elif tipo == "persona":
        c.circle(*P(0.5, 0.72), s * 0.2, stroke=0, fill=1)
        p = c.beginPath()
        p.moveTo(*P(0.12, 0.05))
        p.curveTo(*P(0.12, 0.5), *P(0.88, 0.5), *P(0.88, 0.05))
        p.close()
        c.drawPath(p, stroke=0, fill=1)
    elif tipo == "documento":
        poligono([(0.18, 0.04), (0.18, 0.96), (0.6, 0.96), (0.82, 0.74), (0.82, 0.04)])
        c.setStrokeColor(detalle)
        for yy in (0.56, 0.4, 0.24):
            linea([(0.32, yy), (0.68, yy)], 0.07)
    elif tipo == "recibo":
        poligono([(0.18, 0.96), (0.82, 0.96), (0.82, 0.1), (0.71, 0.03), (0.61, 0.1), (0.5, 0.03),
                  (0.39, 0.1), (0.29, 0.03), (0.18, 0.1)])
        c.setStrokeColor(detalle)
        for yy in (0.74, 0.56, 0.38):
            linea([(0.32, yy), (0.68, yy)], 0.07)
    elif tipo == "check":
        linea([(0.18, 0.5), (0.41, 0.26), (0.82, 0.72)], 0.14)
    elif tipo == "cruz":
        linea([(0.25, 0.25), (0.75, 0.75)], 0.14)
        linea([(0.25, 0.75), (0.75, 0.25)], 0.14)
    elif tipo == "grafica":
        for cx, h in ((0.1, 0.4), (0.4, 0.64), (0.7, 0.88)):
            c.roundRect(*P(cx, 0.06), s * 0.2, s * h, s * 0.04, stroke=0, fill=1)
    elif tipo == "dona":
        c.circle(*P(0.5, 0.5), s * 0.46, stroke=0, fill=1)
        c.setFillColor(detalle)
        c.wedge(*P(0.04, 0.04), *P(0.96, 0.96), 90, -110, stroke=0, fill=1)
        c.setFillColor(color)
        c.wedge(*P(0.1, 0.1), *P(0.9, 0.9), 90, -110, stroke=0, fill=1)
        c.setFillColor(detalle)
        c.circle(*P(0.5, 0.5), s * 0.2, stroke=0, fill=1)
    elif tipo == "pin":
        c.circle(*P(0.5, 0.64), s * 0.3, stroke=0, fill=1)
        poligono([(0.235, 0.5), (0.765, 0.5), (0.5, 0.03)])
        c.setFillColor(detalle)
        c.circle(*P(0.5, 0.64), s * 0.11, stroke=0, fill=1)
    elif tipo == "calendario":
        c.roundRect(*P(0.08, 0.06), s * 0.84, s * 0.76, s * 0.1, stroke=0, fill=1)
        linea([(0.3, 0.76), (0.3, 0.96)], 0.09)
        linea([(0.7, 0.76), (0.7, 0.96)], 0.09)
        c.setFillColor(detalle)
        c.rect(*P(0.08, 0.6), s * 0.84, s * 0.05, stroke=0, fill=1)
        for cx in (0.2, 0.43, 0.66):
            for cy in (0.38, 0.16):
                c.rect(*P(cx, cy), s * 0.14, s * 0.13, stroke=0, fill=1)
    elif tipo == "lupa":
        c.setLineWidth(s * 0.11)
        c.circle(*P(0.42, 0.58), s * 0.27, stroke=1, fill=0)
        linea([(0.62, 0.38), (0.88, 0.12)], 0.15)
    elif tipo == "lista":
        for yy in (0.74, 0.5, 0.26):
            c.roundRect(*P(0.06, yy - 0.08), s * 0.16, s * 0.16, s * 0.03, stroke=0, fill=1)
            linea([(0.36, yy), (0.92, yy)], 0.1)
    c.restoreState()


class Icono(Flowable):
    """Icono dentro de un círculo (o cuadro redondeado) de color, listo para usarse en tablas."""

    def __init__(self, tipo, size=16, fondo=NAVY, color=colors.white, forma="circulo",
                 pad=0.24, detalle=None):
        super().__init__()
        self.tipo, self.size, self.fondo, self.color = tipo, size, fondo, color
        self.forma, self.pad = forma, pad
        self.detalle = detalle or fondo or colors.white

    def wrap(self, *_):
        return self.size, self.size

    def draw(self):
        c, s = self.canv, self.size
        if self.fondo is not None:
            c.setFillColor(self.fondo)
            if self.forma == "circulo":
                c.circle(s / 2, s / 2, s / 2, stroke=0, fill=1)
            else:
                c.roundRect(0, 0, s, s, s * 0.24, stroke=0, fill=1)
        inset = s * self.pad
        dibujar_icono(c, self.tipo, inset, inset, s - 2 * inset, self.color, self.detalle)


class BarraComposicion(Flowable):
    """Barra horizontal apilada con extremos redondeados (neto + retenciones)."""

    def __init__(self, segmentos, ancho, alto=6):
        super().__init__()
        self.segmentos = [(float(v), col) for v, col in segmentos if v > 0]
        self.ancho, self.alto = ancho, alto

    def wrap(self, *_):
        return self.ancho, self.alto

    def draw(self):
        c = self.canv
        total = sum(v for v, _ in self.segmentos) or 1
        c.saveState()
        p = c.beginPath()
        p.roundRect(0, 0, self.ancho, self.alto, self.alto / 2)
        c.clipPath(p, stroke=0, fill=0)
        c.setFillColor(GRIS_LINEA)
        c.rect(0, 0, self.ancho, self.alto, stroke=0, fill=1)
        x = 0
        for v, col in self.segmentos:
            w = self.ancho * v / total
            c.setFillColor(col)
            c.rect(x, 0, w, self.alto, stroke=0, fill=1)
            x += w
            c.setStrokeColor(colors.white)
            c.setLineWidth(0.8)
            if x < self.ancho - 0.5:
                c.line(x, 0, x, self.alto)
        c.restoreState()


class Chip(Flowable):
    """Etiqueta redondeada pequeña (p. ej. el número de operación #001)."""

    def __init__(self, texto, fondo=AMARILLO_PY, color=NAVY, size=7.5):
        super().__init__()
        self.texto, self.fondo, self.color, self.size = texto, fondo, color, size
        self.w = stringWidth(texto, "Helvetica-Bold", size) + 12
        self.h = size + 7

    def wrap(self, *_):
        return self.w, self.h

    def draw(self):
        c = self.canv
        c.setFillColor(self.fondo)
        c.roundRect(0, 0, self.w, self.h, self.h / 2, stroke=0, fill=1)
        c.setFillColor(self.color)
        c.setFont("Helvetica-Bold", self.size)
        c.drawCentredString(self.w / 2, 3.6, self.texto)


def _segmentos_cliente(cliente, visibles):
    segs = [(cliente["neto"], COLOR_CONCEPTO["neto"])]
    segs += [(-cliente[k], COLOR_CONCEPTO[k]) for k, _, _ in visibles if cliente[k] < 0]
    return segs


class PanelDona(Flowable):
    """Dona de '¿A dónde va el dinero financiado?' con leyenda, montos y barras de porcentaje."""

    def __init__(self, datos, ancho):
        super().__init__()
        calc = datos["calculo"]
        clientes = datos["clientes"]
        self.base = float(calc["suma_financiado"]) or 1.0
        self.items = [("neto", float(calc["suma_netos"]))]
        for clave, _, _ in calc["retenciones_visibles"]:
            monto = -float(sum((c[clave] for c in clientes), Decimal("0")))
            if monto > 0:
                self.items.append((clave, monto))
        self.ancho = ancho
        self.alto = max(104, 23 * len(self.items) + 14)

    def wrap(self, *_):
        return self.ancho, self.alto

    def draw(self):
        c, W, H = self.canv, self.ancho, self.alto
        c.saveState()
        c.setFillColor(GRIS_FONDO)
        c.roundRect(0, 0, W, H, 8, stroke=0, fill=1)

        # Dona
        R, r = 50, 32
        cx, cy = 22 + R, H / 2
        total = sum(m for _, m in self.items) or 1.0
        angulo = 90.0
        c.setStrokeColor(GRIS_FONDO)
        c.setLineWidth(1.6)
        for clave, monto in self.items:
            ext = -360.0 * monto / total
            c.setFillColor(COLOR_CONCEPTO[clave])
            if abs(ext) >= 359.99:
                c.circle(cx, cy, R, stroke=0, fill=1)
            else:
                c.wedge(cx - R, cy - R, cx + R, cy + R, angulo, ext, stroke=1, fill=1)
            angulo += ext
        c.setFillColor(GRIS_FONDO)
        c.circle(cx, cy, r, stroke=0, fill=1)
        pct_neto = 100 * self.items[0][1] / self.base
        c.setFillColor(COLOR_CONCEPTO["neto"])
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString(cx, cy + 1, f"{pct_neto:.1f}%")
        c.setFillColor(GRIS_TEXTO)
        c.setFont("Helvetica", 6.5)
        c.drawCentredString(cx, cy - 9, "llega a la agencia")

        # Leyenda
        x0 = cx + R + 30
        x1 = W - 16
        fila = 23
        y = H / 2 + (len(self.items) * fila) / 2 - 15
        for clave, monto in self.items:
            col = COLOR_CONCEPTO[clave]
            pct = 100 * monto / self.base
            c.setFillColor(col)
            c.circle(x0 + 8, y + 3, 8, stroke=0, fill=1)
            dibujar_icono(c, ICONO_CONCEPTO[clave], x0 + 3.5, y - 1.5, 9, colors.white, col)
            c.setFillColor(TEXTO)
            c.setFont("Helvetica-Bold" if clave == "neto" else "Helvetica", 8.3)
            c.drawString(x0 + 22, y + 4, NOMBRE_CONCEPTO[clave])
            c.setFont("Helvetica-Bold", 8.3)
            c.setFillColor(col if clave == "neto" else TEXTO)
            c.drawRightString(x1 - 42, y + 4, fmt(monto))
            c.setFont("Helvetica", 7.5)
            c.setFillColor(GRIS_TEXTO)
            c.drawRightString(x1, y + 4, f"{pct:.1f}%")
            # barra de proporción
            bx, bw = x0 + 22, x1 - (x0 + 22)
            c.setFillColor(GRIS_LINEA)
            c.roundRect(bx, y - 4, bw, 3.5, 1.75, stroke=0, fill=1)
            c.setFillColor(col)
            c.roundRect(bx, y - 4, max(bw * pct / 100, 3.5), 3.5, 1.75, stroke=0, fill=1)
            y -= fila
        c.restoreState()


# ---------------------------------------------------------------- estilos ---
def _estilos():
    base = dict(fontName="Helvetica", fontSize=8.5, leading=11, textColor=TEXTO)

    def E(nombre, **kw):
        return ParagraphStyle(nombre, **{**base, **kw})

    return {
        "titulo": E("titulo", fontName="Helvetica-Bold", fontSize=17, leading=20, textColor=colors.white),
        "subtitulo": E("subtitulo", fontSize=8.5, leading=11, textColor=NAVY_SUAVE),
        "enc_der": E("enc_der", fontSize=8.5, leading=11, textColor=colors.white),
        "enc_der_b": E("enc_der_b", fontName="Helvetica-Bold", fontSize=9, leading=11, textColor=colors.white),
        "seccion": E("seccion", fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=NAVY),
        "normal": E("normal"),
        "negrita": E("negrita", fontName="Helvetica-Bold"),
        "der": E("der", alignment=TA_RIGHT),
        "der_b": E("der_b", fontName="Helvetica-Bold", alignment=TA_RIGHT),
        "gris": E("gris", textColor=GRIS_TEXTO),
        "mini": E("mini", fontSize=7, leading=9, textColor=GRIS_TEXTO),
        "kpi_lbl": E("kpi_lbl", fontName="Helvetica-Bold", fontSize=6.5, leading=8, textColor=GRIS_TEXTO),
        "kpi_val": E("kpi_val", fontName="Helvetica-Bold", fontSize=12.5, leading=15),
        "card_nombre": E("card_nombre", fontName="Helvetica-Bold", fontSize=9.5, leading=11.5,
                         textColor=colors.white),
        "card_sub": E("card_sub", fontSize=7, leading=9, textColor=NAVY_SUAVE),
        "chip": E("chip", fontName="Helvetica-Bold", fontSize=7.5, leading=9, textColor=NAVY,
                  alignment=TA_CENTER),
        "th": E("th", fontName="Helvetica-Bold", fontSize=7.5, leading=9, textColor=colors.white),
        "th_der": E("th_der", fontName="Helvetica-Bold", fontSize=7.5, leading=9, textColor=colors.white,
                    alignment=TA_RIGHT),
        "marca_enc": E("marca_enc", fontName="Helvetica-Bold", fontSize=8.5, leading=10,
                       textColor=colors.white),
        "banner": E("banner", fontSize=9, leading=12.5),
    }


SIN_PAD = [("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
           ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]


def _icono_texto(icono, texto, ancho, sep=6):
    """Celda compuesta: icono + texto (Paragraph) alineados al centro."""
    t = Table([[icono, texto]], colWidths=[icono.size + sep, ancho - icono.size - sep])
    t.setStyle(TableStyle(SIN_PAD))
    return t


# --------------------------------------------------------------- canvas ---
class _CanvasNumerado(rl_canvas.Canvas):
    """Pie de página con marca y 'Página X de Y'; encabezado corto a partir de la página 2."""
    pie_izquierdo = ""
    encabezado_corrido = ""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._paginas = []

    def showPage(self):
        self._paginas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._paginas)
        for estado in self._paginas:
            self.__dict__.update(estado)
            self._decorar(total)
            super().showPage()
        super().save()

    def _decorar(self, total):
        ancho, alto = letter
        izq, der = 0.6 * inch, ancho - 0.6 * inch
        # Encabezado corrido (páginas 2 en adelante)
        if self._pageNumber > 1 and self.encabezado_corrido:
            self.setFillColor(NAVY)
            self.roundRect(izq, alto - 0.42 * inch, 3, 9, 1.5, stroke=0, fill=1)
            self.setFont("Helvetica-Bold", 7.5)
            self.drawString(izq + 7, alto - 0.42 * inch + 1.5, self.encabezado_corrido)
            self.setStrokeColor(GRIS_LINEA)
            self.setLineWidth(0.6)
            self.line(izq, alto - 0.48 * inch, der, alto - 0.48 * inch)
        # Pie
        y = 0.38 * inch
        self.setStrokeColor(GRIS_LINEA)
        self.setLineWidth(0.6)
        self.line(izq, 0.55 * inch, der, 0.55 * inch)
        self.setFont("Helvetica", 7)
        self.setFillColor(GRIS_TEXTO)
        self.drawString(izq, y, self.pie_izquierdo)
        pagina = f"Página {self._pageNumber} de {total}"
        self.drawRightString(der, y, pagina)
        x = der - stringWidth(pagina, "Helvetica", 7) - 8
        self.setFillColor(GRIS_LINEA)
        self.drawRightString(x, y, "|")
        x -= stringWidth("|", "Helvetica", 7) + 8
        self.setFont("Helvetica-Bold", 7)
        w_py = stringWidth("dotpy", "Helvetica-Bold", 7)
        if MARCA_HANDLE.endswith("dotpy"):
            self.setFillColor(colors.HexColor("#C99A06"))
            self.drawRightString(x, y, "dotpy")
            self.setFillColor(AZUL_PY)
            self.drawRightString(x - w_py, y, MARCA_HANDLE[:-5])
        else:
            self.setFillColor(AZUL_PY)
            self.drawRightString(x, y, MARCA_HANDLE)


# ------------------------------------------------------------- bloques ---
def _marca_html():
    """@analyticsdotpy con 'dotpy' resaltado en amarillo."""
    h = escape(MARCA_HANDLE)
    if h.endswith("dotpy"):
        return f"{h[:-5]}<font color='#FFD43B'>dotpy</font>"
    return h


def _encabezado(datos, s, ancho):
    logo = Icono("grafica", 40, fondo=AMARILLO_PY, color=NAVY, forma="cuadrado", pad=0.24)
    titulo = [Paragraph("DESGLOSE POR CLIENTE", s["titulo"]), Spacer(1, 2),
              Paragraph("Basado en el Reporte de Operaciones de Autoexpress Inbursa", s["subtitulo"])]
    pdv = escape(datos["pdv"] or "PDV no identificado")
    num = f" - {datos['pdv_num']}" if datos["pdv_num"] else ""
    w_der = ancho * 0.36
    der = Table([
        [_icono_texto(Icono("pin", 15, fondo=NAVY_MARCA, pad=0.2), Paragraph(f"{pdv}{num}", s["enc_der_b"]),
                      w_der - 28)],
        [_icono_texto(Icono("calendario", 15, fondo=NAVY_MARCA, pad=0.22),
                      Paragraph(escape(datos["fecha_texto"] or "Fecha no identificada"), s["enc_der"]),
                      w_der - 28)],
    ], colWidths=[w_der - 28])
    der.setStyle(TableStyle(SIN_PAD + [("BOTTOMPADDING", (0, 0), (-1, 0), 6)]))
    marca = Paragraph(_marca_html(), s["marca_enc"])
    t = Table([[logo, titulo, der], [marca, "", ""]],
              colWidths=[56, ancho - 56 - (w_der - 28) - 28, w_der])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("BACKGROUND", (0, 1), (-1, 1), NAVY_MARCA),
        ("SPAN", (0, 1), (-1, 1)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 14), ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("LEFTPADDING", (1, 0), (1, 0), 4),
        ("TOPPADDING", (0, 0), (-1, 0), 12), ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("TOPPADDING", (0, 1), (-1, 1), 5), ("BOTTOMPADDING", (0, 1), (-1, 1), 6),
    ]))
    return t


def _kpis(datos, s, ancho):
    calc = datos["calculo"]
    n = len(datos["clientes"])
    items = [
        ("billete", "TOTAL FINANCIADO", fmt(calc["suma_financiado"]), NAVY),
        ("recibo", "TOTAL RETENCIONES", fmt(calc["suma_retenciones"]), ROJO),
        ("banco", "TRANSFERENCIAS", fmt(calc["total_calculado"]), VERDE),
        ("persona", "CRÉDITOS", f"{n} {'cliente' if n == 1 else 'clientes'}", AZUL_PY),
    ]
    gap = 7
    w = (ancho - gap * 3) / 4
    celdas = []
    for icono, lbl, val, col in items:
        # El tamaño de letra se ajusta solo para que montos de millones quepan en una línea
        disponible = w - 20 - 28 - 7 - 2
        tam = min(12.5, disponible / stringWidth(val, "Helvetica-Bold", 1))
        texto = [Paragraph(lbl, s["kpi_lbl"]), Spacer(1, 2),
                 Paragraph(val, ParagraphStyle("v", parent=s["kpi_val"], textColor=col,
                                               fontSize=tam, leading=tam + 2.5))]
        celdas.append(_icono_texto(Icono(icono, 28, fondo=aclarar(col, 0.86), color=col, pad=0.24),
                                   texto, w - 20, sep=7))
    t = Table([[celdas[0], "", celdas[1], "", celdas[2], "", celdas[3]]],
              colWidths=[w, gap, w, gap, w, gap, w])
    estilo = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
              ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]
    for i, col in enumerate((0, 2, 4, 6)):
        estilo += [("BACKGROUND", (col, 0), (col, 0), GRIS_FONDO),
                   ("LINEBELOW", (col, 0), (col, 0), 2.5, items[i][3])]
    for col in (1, 3, 5):
        estilo += [("LEFTPADDING", (col, 0), (col, 0), 0), ("RIGHTPADDING", (col, 0), (col, 0), 0)]
    t.setStyle(TableStyle(estilo))
    return t


def _seccion(titulo, icono, s, ancho):
    enc = _icono_texto(Icono(icono, 18, fondo=NAVY, pad=0.25), Paragraph(titulo, s["seccion"]), ancho, sep=7)
    return [Spacer(1, 10), enc,
            HRFlowable(width=ancho, thickness=0.8, color=GRIS_LINEA, spaceBefore=4, spaceAfter=7)]


def _tarjeta(cliente, datos, s, w):
    visibles = datos["calculo"]["retenciones_visibles"]
    w_lbl, w_val = w * 0.6, w * 0.4
    rojo_der = ParagraphStyle("rd", parent=s["der"], textColor=ROJO)
    gris_der = ParagraphStyle("gd", parent=s["der"], textColor=GRIS_TEXTO)

    encabezado = Table([[Icono("persona", 26, fondo=NAVY_MARCA, pad=0.22),
                         [Paragraph(escape(cliente["titular"]), s["card_nombre"]),
                          Paragraph(f"Número de folio {escape(cliente['folio'])}", s["card_sub"])],
                         Chip(f"#{cliente['num']}")]],
                       colWidths=[34, w - 34 - 54, 36])
    encabezado.setStyle(TableStyle(SIN_PAD + [("ALIGN", (2, 0), (2, 0), "RIGHT")]))

    def fila_concepto(clave, etiqueta, estilo_lbl, icono_col):
        ic = Icono(ICONO_CONCEPTO[clave], 15, fondo=aclarar(icono_col, 0.84), color=icono_col, pad=0.2)
        return _icono_texto(ic, Paragraph(etiqueta, estilo_lbl), w_lbl - 18)

    filas = [[encabezado, ""],
             [fila_concepto("financiado", "Monto financiado", s["negrita"], NAVY),
              Paragraph(fmt(cliente["financiado"]), s["der_b"])]]
    for clave, _, _ in visibles:
        valor = cliente[clave]
        col = COLOR_CONCEPTO[clave]
        filas.append([fila_concepto(clave, f"(-) {NOMBRE_CONCEPTO[clave]}", s["gris"], col),
                      Paragraph(fmt(valor), rojo_der if valor else gris_der)])
    filas.append([Paragraph("Total retenciones", s["negrita"]),
                  Paragraph(fmt(cliente["total_retenciones"]),
                            ParagraphStyle("tr", parent=s["der_b"], textColor=ROJO))])
    pct_neto = 100 - float(cliente["pct_retenido"])
    filas.append([[BarraComposicion(_segmentos_cliente(cliente, visibles), w - 18, 7), Spacer(1, 3),
                   Paragraph(f"<font color='#079455'><b>{pct_neto:.1f}%</b> llega a la agencia</font>"
                             f"&nbsp;&nbsp;·&nbsp;&nbsp;{float(cliente['pct_retenido']):.1f}% retenido",
                             s["mini"])], ""])
    neto_lbl = ParagraphStyle("nl", parent=s["negrita"], textColor=VERDE, fontSize=9.5, leading=12)
    neto_val = ParagraphStyle("nv", parent=s["der_b"], textColor=VERDE, fontSize=12, leading=14)
    filas.append([_icono_texto(Icono("banco", 20, fondo=VERDE, pad=0.22),
                               Paragraph("Transferencia neta", neto_lbl), w_lbl - 18, sep=7),
                  Paragraph(fmt(cliente["neto"]), neto_val)])

    n = len(filas)
    t = Table(filas, colWidths=[w_lbl, w_val])
    t.setStyle(TableStyle([
        ("SPAN", (0, 0), (1, 0)), ("SPAN", (0, n - 2), (1, n - 2)),
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("BACKGROUND", (0, n - 1), (-1, n - 1), VERDE_FONDO),
        ("BOX", (0, 0), (-1, -1), 0.6, GRIS_LINEA),
        ("LINEBELOW", (0, 1), (-1, n - 4), 0.4, GRIS_LINEA),
        ("LINEABOVE", (0, n - 3), (-1, n - 3), 0.8, TEXTO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 9), ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("TOPPADDING", (0, 0), (-1, 0), 8), ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, n - 2), (-1, n - 2), 7), ("BOTTOMPADDING", (0, n - 2), (-1, n - 2), 6),
        ("TOPPADDING", (0, n - 1), (-1, n - 1), 7), ("BOTTOMPADDING", (0, n - 1), (-1, n - 1), 7),
    ]))
    return t


def _rejilla_tarjetas(datos, s, ancho):
    """Tarjetas de cliente en 2 columnas."""
    gap = 12
    w = (ancho - gap) / 2
    tarjetas = [_tarjeta(c, datos, s, w) for c in datos["clientes"]]
    bloques = []
    for i in range(0, len(tarjetas), 2):
        par = tarjetas[i:i + 2]
        if len(par) == 1:
            par.append("")
        fila = Table([[par[0], "", par[1]]], colWidths=[w, gap, w])
        fila.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                  ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 12)]))
        bloques.append(fila)
    return bloques


def _tabla_resumen(datos, s, ancho):
    calc = datos["calculo"]
    visibles = calc["retenciones_visibles"]
    cw = [0.05, 0.12, 0.25, 0.15, 0.13, 0.15, 0.15]
    w_barra = ancho * cw[6] - 12
    enc = [Paragraph(x, s["th"]) for x in ("#", "Folio", "Titular del crédito")] + \
          [Paragraph(x, s["th_der"]) for x in ("Financiado", "Retenciones", "Neto")] + \
          [Paragraph("Composición", s["th"])]
    filas = [enc]
    for c in datos["clientes"]:
        filas.append([c["num"], c["folio"], Paragraph(escape(c["titular"]), s["normal"]),
                      fmt(c["financiado"]), fmt(c["total_retenciones"]), fmt(c["neto"]),
                      BarraComposicion(_segmentos_cliente(c, visibles), w_barra, 7)])
    ultima_cliente = len(filas) - 1
    if calc["cancelaciones_aplicadas"] != 0:
        filas.append(["", "", Paragraph("Saldo de cancelaciones de créditos", s["gris"]),
                      "", "", fmt(calc["cancelaciones_aplicadas"]), ""])
    total_segs = [(calc["suma_netos"], COLOR_CONCEPTO["neto"])] + [
        (-sum((c[k] for c in datos["clientes"]), Decimal("0")), COLOR_CONCEPTO[k]) for k, _, _ in visibles]
    filas.append(["", "", Paragraph("<b>TOTAL</b>", s["normal"]),
                  fmt(calc["suma_financiado"]), fmt(calc["suma_retenciones"]), fmt(calc["total_calculado"]),
                  BarraComposicion(total_segs, w_barra, 7)])

    t = Table(filas, colWidths=[ancho * x for x in cw], repeatRows=1)
    n = len(filas)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 1), (-1, -1), 8.3),
        ("TEXTCOLOR", (0, 1), (-1, -1), TEXTO),
        ("TEXTCOLOR", (0, 1), (1, -1), GRIS_TEXTO),
        ("ALIGN", (3, 1), (5, -1), "RIGHT"),
        ("TEXTCOLOR", (4, 1), (4, -1), ROJO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 1), (-1, ultima_cliente), 0.4, GRIS_LINEA),
        ("LINEABOVE", (0, n - 1), (-1, n - 1), 1, NAVY),
        ("BACKGROUND", (0, n - 1), (-1, n - 1), VERDE_FONDO),
        ("FONTNAME", (3, n - 1), (-1, n - 1), "Helvetica-Bold"),
        ("TEXTCOLOR", (5, 1), (5, -1), VERDE), ("FONTNAME", (5, 1), (5, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]
    for i in range(1, ultima_cliente + 1):
        if i % 2 == 0:
            estilo.append(("BACKGROUND", (0, i), (-1, i), GRIS_FONDO))
    t.setStyle(TableStyle(estilo))
    return t


def _leyenda_colores(datos, s, ancho):
    """Pequeña leyenda de colores de las barras de composición."""
    claves = ["neto"] + [k for k, _, _ in datos["calculo"]["retenciones_visibles"]]
    partes = [f"<font color='{COLOR_CONCEPTO[k].hexval().replace('0x', '#')}'>&#9632;</font>&nbsp;"
              f"{NOMBRE_CONCEPTO[k] if k != 'neto' else 'Neto a la agencia'}" for k in claves]
    return Paragraph("Composición:&nbsp;&nbsp;" + "&nbsp;&nbsp;&nbsp;".join(partes), s["mini"])


def _tabla_conciliacion(datos, s, ancho):
    calc = datos["calculo"]
    cw = [0.36, 0.18, 0.18, 0.13, 0.15]
    enc = [Paragraph("Concepto", s["th"])] + \
          [Paragraph(x, s["th_der"]) for x in ("Suma del desglose", "Total del reporte", "Diferencia")] + \
          [Paragraph("Estatus", s["th"])]
    filas = [enc]

    def _fila(clave_icono, color_icono, concepto, suma, reportado, negrita=False):
        if reportado is None:
            dif_txt, estado, ok = "—", "Sin dato", None
        else:
            dif = suma - reportado
            dif_txt, ok = fmt(dif), dif == 0
            estado = "Cuadra" if ok else "Revisar"
        col_est = VERDE if ok else (ROJO if ok is False else GRIS_TEXTO)
        est = _icono_texto(Icono("check" if ok else "cruz", 13, fondo=col_est, pad=0.2),
                           Paragraph(f"<b>{estado}</b>", ParagraphStyle("e", parent=s["normal"],
                                                                       textColor=col_est)),
                           ancho * cw[4] - 12, sep=5)
        lbl = _icono_texto(Icono(clave_icono, 15, fondo=aclarar(color_icono, 0.84), color=color_icono, pad=0.2),
                           Paragraph(f"<b>{concepto}</b>" if negrita else concepto, s["normal"]),
                           ancho * cw[0] - 12)
        filas.append([lbl, fmt(suma), fmt(reportado) if reportado is not None else "—", dif_txt, est])

    omitidos = []
    for (clave, _, _), fila in zip(CONCEPTOS, calc["conciliacion"]):
        # Se omiten los conceptos que vienen en cero tanto en el desglose como en el reporte
        if clave != "financiado" and fila["suma"] == 0 and not fila["reportado"]:
            omitidos.append(NOMBRE_CONCEPTO[clave].lower() if clave != "gap_contado" else "GAP contado")
            continue
        _fila(ICONO_CONCEPTO[clave], COLOR_CONCEPTO.get(clave, NAVY), fila["concepto"], fila["suma"],
              fila["reportado"])
    if calc["cancelaciones_aplicadas"] != 0 or datos["totales"].get("cancelaciones"):
        _fila("documento", GRIS_TEXTO, "Saldo de cancelaciones de créditos", calc["cancelaciones_aplicadas"],
              datos["totales"].get("cancelaciones"))
    else:
        omitidos.append("cancelaciones de créditos")
    _fila("banco", VERDE, "Total de transferencias bancarias", calc["total_calculado"],
          calc["reportado_transferencias"], negrita=True)

    t = Table(filas, colWidths=[ancho * x for x in cw])
    n = len(filas)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 1), (-1, -1), 8.3),
        ("TEXTCOLOR", (0, 1), (-1, -1), TEXTO),
        ("ALIGN", (1, 1), (3, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 1), (-1, n - 2), 0.4, GRIS_LINEA),
        ("LINEABOVE", (0, n - 1), (-1, n - 1), 1, NAVY),
        ("FONTNAME", (1, n - 1), (3, n - 1), "Helvetica-Bold"),
        ("BACKGROUND", (0, n - 1), (-1, n - 1), GRIS_FONDO),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    if not omitidos:
        return [t]
    nota = Paragraph("Conceptos en $0.00 en todo el reporte (no se muestran): " + ", ".join(omitidos) + ".",
                     s["mini"])
    return [t, Spacer(1, 4), nota]


def _banner_resultado(datos, s, ancho):
    calc = datos["calculo"]
    if calc["reportado_transferencias"] is None:
        texto = ("<b>ATENCIÓN:</b> no se encontró el Total de Transferencias Bancarias en el PDF, "
                 f"por lo que no se pudo validar. Total calculado: <b>{fmt(calc['total_calculado'])}</b>.")
        fondo, color, icono = ROJO_FONDO, ROJO, "cruz"
    elif datos["cuadra"]:
        texto = (f"<b>EL DESGLOSE CUADRA.</b> La suma de las transferencias por cliente "
                 f"(<b>{fmt(calc['total_calculado'])}</b>) es igual al Total de Transferencias Bancarias "
                 f"del reporte (<b>{fmt(calc['reportado_transferencias'])}</b>).")
        fondo, color, icono = VERDE_FONDO, VERDE, "check"
    else:
        texto = (f"<b>EL DESGLOSE NO CUADRA.</b> Total calculado <b>{fmt(calc['total_calculado'])}</b> vs. "
                 f"reporte <b>{fmt(calc['reportado_transferencias'])}</b>. Diferencia: "
                 f"<b>{fmt(calc['diferencia_final'])}</b>. Revisa la tabla de conciliación.")
        fondo, color, icono = ROJO_FONDO, ROJO, "cruz"
    p = Paragraph(texto, ParagraphStyle("bn", parent=s["banner"], textColor=color))
    t = Table([[Icono(icono, 26, fondo=color, pad=0.22), p]], colWidths=[40, ancho - 40])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), fondo),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("LINEBEFORE", (0, 0), (0, 0), 3, color),
                           ("LEFTPADDING", (0, 0), (0, 0), 12), ("LEFTPADDING", (1, 0), (1, 0), 4),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                           ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
    return t


def generar_pdf_desglose(datos, destino, nombre_origen=""):
    """destino: ruta de salida o archivo tipo bytes (BytesIO)."""
    s = _estilos()
    margen = 0.6 * inch
    doc = SimpleDocTemplate(destino, pagesize=letter, leftMargin=margen, rightMargin=margen,
                            topMargin=0.55 * inch, bottomMargin=0.72 * inch,
                            title="Desglose por cliente - Reporte de Operaciones",
                            author=MARCA_HANDLE, creator=f"desglose_reporte_pagos.py · {MARCA_HANDLE}")
    ancho = letter[0] - 2 * margen
    calc = datos["calculo"]

    historia = [_encabezado(datos, s, ancho), Spacer(1, 10), _kpis(datos, s, ancho)]

    # 1) Panorama: a dónde va el dinero
    panorama = _seccion("¿A dónde va el dinero financiado?", "dona", s, ancho) + [PanelDona(datos, ancho)]
    if calc["cancelaciones_aplicadas"] != 0:
        panorama += [Spacer(1, 4), Paragraph(
            f"Además, el reporte aplica un saldo de cancelaciones de créditos de "
            f"<b>{fmt(calc['cancelaciones_aplicadas'])}</b> al total de transferencias.", s["mini"])]
    historia.append(KeepTogether(panorama))

    # 2) Resumen por cliente
    historia.append(KeepTogether(_seccion("Resumen por cliente", "lista", s, ancho)
                                 + [_tabla_resumen(datos, s, ancho), Spacer(1, 4),
                                    _leyenda_colores(datos, s, ancho)]))

    # 3) Conciliación
    historia.append(KeepTogether(
        _seccion("Conciliación contra el reporte de Inbursa", "lupa", s, ancho)
        + _tabla_conciliacion(datos, s, ancho) + [Spacer(1, 8), _banner_resultado(datos, s, ancho)]
    ))

    # 4) Detalle por cliente (el título va junto con la primera fila de tarjetas)
    tarjetas = _rejilla_tarjetas(datos, s, ancho)
    historia.append(KeepTogether(_seccion("Detalle por cliente", "persona", s, ancho) + tarjetas[:1]))
    historia += tarjetas[1:]

    try:
        from zoneinfo import ZoneInfo
        ahora = datetime.now(ZoneInfo("America/Mexico_City"))
    except Exception:
        ahora = datetime.now()
    pie = f"Generado el {ahora:%d/%m/%Y %H:%M}"
    if nombre_origen:
        pie += f"  ·  Fuente: {nombre_origen}"
    num = f" - {datos['pdv_num']}" if datos["pdv_num"] else ""
    # Subclase por documento: así varios usuarios pueden generar PDFs a la vez (p. ej. en Streamlit)
    canvas_doc = type("_CanvasDoc", (_CanvasNumerado,), {
        "pie_izquierdo": pie,
        "encabezado_corrido": (f"DESGLOSE POR CLIENTE  ·  {datos['pdv'] or ''}{num}  ·  "
                               f"{datos['fecha_texto'] or ''}"),
    })
    doc.build(historia, canvasmaker=canvas_doc)


# ==========================================================================
# 6. FLUJO PRINCIPAL
# ==========================================================================
def nombre_salida(datos, nombre_origen):
    fecha = datos["fecha"].strftime("%Y-%m-%d") if datos["fecha"] else "sin_fecha"
    pdv = datos["pdv_num"] or os.path.splitext(os.path.basename(nombre_origen))[0]
    return f"Desglose_ReportePagos_{pdv}_{fecha}.pdf"


def imprimir_resumen(datos):
    calc = datos["calculo"]
    print("\n" + "=" * 78)
    print(f"{datos['pdv']} - {datos['pdv_num']}   |   {datos['fecha_texto']}")
    print("=" * 78)
    print(f"{'#':<4}{'Folio':<11}{'Titular':<24}{'Financiado':>13}{'Retenciones':>13}{'Neto':>13}")
    print("-" * 78)
    for c in datos["clientes"]:
        print(f"{c['num']:<4}{c['folio']:<11}{c['titular'][:23]:<24}"
              f"{fmt(c['financiado'], False):>13}{fmt(c['total_retenciones'], False):>13}"
              f"{fmt(c['neto'], False):>13}")
    if calc["cancelaciones_aplicadas"] != 0:
        print(f"{'':<39}{'Cancelaciones':>26}{fmt(calc['cancelaciones_aplicadas'], False):>13}")
    print("-" * 78)
    print(f"{'TOTAL':<39}{fmt(calc['suma_financiado'], False):>13}"
          f"{fmt(calc['suma_retenciones'], False):>13}{fmt(calc['total_calculado'], False):>13}")
    if calc["reportado_transferencias"] is not None:
        estado = "CUADRA" if datos["cuadra"] else f"NO CUADRA (dif. {fmt(calc['diferencia_final'])})"
        print(f"Total de Transferencias Bancarias del reporte: {fmt(calc['reportado_transferencias'])}  ->  {estado}")
    print()


def procesar(ruta_pdf, carpeta_salida="."):
    nombre = os.path.basename(ruta_pdf)
    datos = calcular(extraer_reporte(ruta_pdf))
    imprimir_resumen(datos)
    salida = os.path.join(carpeta_salida, nombre_salida(datos, nombre))
    generar_pdf_desglose(datos, salida, nombre_origen=nombre)
    print(f"PDF generado: {salida}")
    print(MARCA_HANDLE)
    return salida


def main():
    try:
        from google.colab import files  # type: ignore
        en_colab = True
    except ImportError:
        en_colab = False

    if en_colab:
        print("Sube uno o varios PDFs del Reporte de Operaciones (ReportePagos_*.pdf):")
        subidos = files.upload()
        for nombre in subidos:
            if not nombre.lower().endswith(".pdf"):
                print(f"Se omite {nombre}: no es PDF.")
                continue
            try:
                salida = procesar(nombre)
                files.download(salida)
            except Exception as e:
                print(f"ERROR con {nombre}: {e}")
    else:
        rutas = [a for a in sys.argv[1:] if a.lower().endswith(".pdf")]
        if not rutas:
            print("Uso: python desglose_reporte_pagos.py ReportePagos_xxx.pdf [otro.pdf ...]")
            return
        for ruta in rutas:
            try:
                procesar(ruta)
            except Exception as e:
                print(f"ERROR con {ruta}: {e}")


if __name__ == "__main__":
    main()
