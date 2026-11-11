# -*- coding: utf-8 -*-
"""
===============================================================================
 app_ems_suelos_e050_pro.py
 Sistema Automatizador de Estudios de Mecánica de Suelos (EMS - NTE E.050):
 Capacidad Portante, Perfiles SUCS en AutoCAD (.DXF) e Informe en Word (.DOCX)
-------------------------------------------------------------------------------
 Normativa de referencia (Perú):
   - RNE NTE E.050 Suelos y Cimentaciones
   - RNE NTE E.030 Diseño Sismorresistente (2018)
   - RNE NTE E.060 Concreto Armado (Tabla 4.4 - exposición a sulfatos) / ACI 318
   - NTP 339.127, 339.128, 339.129, 339.131, 339.134, 339.135, 339.150,
     339.152, 339.162, 339.171, 339.177, 339.178 / ASTM equivalentes
 100 % offline. Librerías gratuitas: customtkinter, openpyxl, pandas, ezdxf,
 python-docx, Pillow.

 Uso:
   python app_ems_suelos_e050_pro.py              -> abre la interfaz gráfica
   python app_ems_suelos_e050_pro.py --demo DIR   -> genera el estudio de
        ejemplo (DOCX + DXF + XLSX) en la carpeta DIR sin abrir la interfaz.
===============================================================================
"""

import math
import os
import sys
import datetime
import random
import subprocess
import tempfile
import unicodedata

try:
    import pandas as pd
except Exception:  # pandas solo se usa para importar Excel
    pd = None

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series

import ezdxf
from ezdxf.enums import TextEntityAlignment, MTextEntityAlignment

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from PIL import Image, ImageDraw, ImageFont, ImageOps

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import customtkinter as ctk
    GUI_OK = True
except Exception:  # permite usar el modo --demo en servidores sin pantalla
    tk = ttk = filedialog = messagebox = ctk = None
    GUI_OK = False


# =============================================================================
# 0. CONSTANTES GENERALES
# =============================================================================
APP_TITULO = ("Sistema Automatizador de Estudios de Mecánica de Suelos (EMS - NTE E.050): "
              "Capacidad Portante, Perfiles SUCS en AutoCAD (.DXF) e Informe en Word (.DOCX)")
APP_CORTO = "EMS PRO · NTE E.050"
VERSION = "1.0"

DF_MATRIZ = [1.20, 1.50, 1.80, 2.00]          # profundidades de desplante (m)
B_CORRIDO = 0.60                              # cimiento corrido (m)
B_ZAPATAS = [1.20, 1.50, 2.00, 2.50]          # zapatas (m)
METODOS = ["Terzaghi", "Meyerhof", "Vesic"]
GAMMA_W = 1.00                                # t/m3
SMAX_E050 = 2.54                              # cm (1")
DISTORSION_MAX = 1.0 / 500.0                  # NTE E.050

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def fecha_larga(d=None):
    d = d or datetime.date.today()
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


# =============================================================================
# 1. BASE DE DATOS SUCS (descripción, Es típico, Poisson, patrón CAD)
#    Es en kg/cm2 (rangos según Bowles, "Foundation Analysis and Design")
# =============================================================================
SUCS_DB = {
    "GW": dict(nombre="Grava bien graduada",
               desc="Grava bien graduada, mezcla de grava y arena con poco o ningún fino",
               Es=1500, Es_r="1000 - 2000", mu=0.30, phi=38, c=0.00, patron="grava_bg",
               aci=32, rgb=(236, 196, 140), aashto="A-1-a(0)"),
    "GP": dict(nombre="Grava mal graduada",
               desc="Grava mal graduada, mezcla de grava y arena con poco o ningún fino",
               Es=1200, Es_r="800 - 1600", mu=0.30, phi=36, c=0.00, patron="grava_mg",
               aci=32, rgb=(226, 186, 128), aashto="A-1-a(0)"),
    "GM": dict(nombre="Grava limosa",
               desc="Grava limosa, mezcla de grava, arena y limo",
               Es=800, Es_r="500 - 1200", mu=0.30, phi=33, c=0.00, patron="grava_lim",
               aci=34, rgb=(214, 190, 150), aashto="A-1-b(0)"),
    "GC": dict(nombre="Grava arcillosa",
               desc="Grava arcillosa, mezcla de grava, arena y arcilla",
               Es=600, Es_r="400 - 1000", mu=0.30, phi=30, c=0.05, patron="grava_arc",
               aci=34, rgb=(205, 170, 135), aashto="A-2-6(0)"),
    "SW": dict(nombre="Arena bien graduada",
               desc="Arena bien graduada, arena con grava, con poco o ningún fino",
               Es=600, Es_r="500 - 800", mu=0.30, phi=35, c=0.00, patron="arena_bg",
               aci=42, rgb=(250, 230, 160), aashto="A-1-b(0)"),
    "SP": dict(nombre="Arena mal graduada",
               desc="Arena mal graduada, arena con grava, con poco o ningún fino",
               Es=350, Es_r="200 - 500", mu=0.30, phi=32, c=0.00, patron="arena_mg",
               aci=42, rgb=(250, 238, 185), aashto="A-3(0)"),
    "SM": dict(nombre="Arena limosa",
               desc="Arena limosa, mezcla de arena y limo",
               Es=150, Es_r="100 - 250", mu=0.30, phi=30, c=0.00, patron="arena_lim",
               aci=40, rgb=(240, 220, 170), aashto="A-2-4(0)"),
    "SC": dict(nombre="Arena arcillosa",
               desc="Arena arcillosa, mezcla de arena y arcilla",
               Es=120, Es_r="80 - 200", mu=0.30, phi=27, c=0.10, patron="arena_arc",
               aci=40, rgb=(230, 205, 160), aashto="A-2-6(0)"),
    "ML": dict(nombre="Limo inorgánico de baja plasticidad",
               desc="Limo inorgánico, arena muy fina, polvo de roca, limo arenoso o arcilloso de baja plasticidad",
               Es=100, Es_r="50 - 200", mu=0.35, phi=28, c=0.05, patron="limo",
               aci=52, rgb=(215, 215, 175), aashto="A-4(0)"),
    "CL": dict(nombre="Arcilla inorgánica de baja plasticidad",
               desc="Arcilla inorgánica de baja a media plasticidad, arcilla con grava, arenosa o limosa",
               Es=200, Es_r="100 - 400", mu=0.35, phi=22, c=0.20, patron="arcilla",
               aci=22, rgb=(205, 170, 150), aashto="A-6(5)"),
    "OL": dict(nombre="Limo orgánico de baja plasticidad",
               desc="Limo orgánico y arcilla limosa orgánica de baja plasticidad",
               Es=40, Es_r="20 - 60", mu=0.40, phi=18, c=0.08, patron="organico",
               aci=84, rgb=(170, 185, 150), aashto="A-5(5)"),
    "MH": dict(nombre="Limo inorgánico de alta plasticidad",
               desc="Limo inorgánico, suelo limoso o arenoso fino micáceo o diatomáceo, limo elástico",
               Es=80, Es_r="50 - 150", mu=0.40, phi=20, c=0.15, patron="limo_ap",
               aci=52, rgb=(200, 200, 160), aashto="A-5(10)"),
    "CH": dict(nombre="Arcilla inorgánica de alta plasticidad",
               desc="Arcilla inorgánica de alta plasticidad, arcilla grasa",
               Es=120, Es_r="50 - 250", mu=0.45, phi=15, c=0.30, patron="arcilla_ap",
               aci=22, rgb=(190, 150, 135), aashto="A-7-6(20)"),
    "OH": dict(nombre="Arcilla orgánica de alta plasticidad",
               desc="Arcilla orgánica de media a alta plasticidad, limo orgánico",
               Es=30, Es_r="15 - 50", mu=0.45, phi=12, c=0.10, patron="organico",
               aci=84, rgb=(150, 170, 135), aashto="A-7-5(20)"),
    "Pt": dict(nombre="Turba",
               desc="Turba y otros suelos altamente orgánicos",
               Es=10, Es_r="5 - 20", mu=0.40, phi=10, c=0.05, patron="turba",
               aci=94, rgb=(120, 140, 100), aashto="A-8"),
    "Roca": dict(nombre="Roca",
                 desc="Macizo rocoso / basamento rocoso",
                 Es=50000, Es_r="20000 - 300000", mu=0.20, phi=45, c=2.00, patron="roca",
                 aci=8, rgb=(190, 190, 195), aashto="-"),
    "RELL": dict(nombre="Relleno no controlado",
                 desc="Material de relleno heterogéneo no controlado (desmonte, restos orgánicos)",
                 Es=50, Es_r="30 - 100", mu=0.35, phi=24, c=0.00, patron="relleno",
                 aci=250, rgb=(175, 160, 150), aashto="-"),
}
SUCS_LISTA = ["GW", "GP", "GM", "GC", "SW", "SP", "SM", "SC", "ML", "CL", "OL", "MH", "CH", "OH",
              "Pt", "Roca", "RELL", "GP-GM", "GW-GM", "SP-SM", "SW-SM", "SM-SC", "SC-SM", "CL-ML",
              "GC-GM", "GP-GC", "SP-SC"]

ACI_RGB = {1: (220, 0, 0), 2: (230, 200, 0), 3: (0, 150, 0), 4: (0, 170, 200), 5: (0, 60, 220),
           6: (200, 0, 200), 7: (0, 0, 0), 8: (110, 110, 110), 30: (240, 120, 0)}


def sucs_base(s):
    """Devuelve la clave de la base de datos para un símbolo SUCS (incluye duales)."""
    s = (s or "").strip()
    if not s:
        return "SM"
    up = s.upper()
    if up.startswith("RELL") or up.startswith("REL"):
        return "RELL"
    if up.startswith("ROC"):
        return "Roca"
    if up.startswith("PT"):
        return "Pt"
    k = up.split("-")[0].split("/")[0].strip()[:2]
    return k if k in SUCS_DB else "SM"


# =============================================================================
# 2. UTILIDADES NUMÉRICAS
# =============================================================================
def num(x, default=None):
    """Convierte texto/celda a float. Acepta coma decimal. 'NP', '' -> default."""
    if x is None:
        return default
    if isinstance(x, (int, float)):
        if isinstance(x, float) and math.isnan(x):
            return default
        return float(x)
    s = str(x).strip().replace(",", ".")
    if s == "" or s.upper() in ("NP", "N.P.", "N.P", "-", "NAN", "NONE", "S/M"):
        return default
    try:
        return float(s)
    except ValueError:
        return default


def fmt(x, nd=2, np_txt="NP"):
    if x is None:
        return np_txt
    return f"{x:.{nd}f}"


def interp(x, xs, ys):
    if x <= xs[0]:
        return ys[0]
    if x >= xs[-1]:
        return ys[-1]
    for i in range(1, len(xs)):
        if x <= xs[i]:
            t = (x - xs[i - 1]) / (xs[i] - xs[i - 1])
            return ys[i - 1] + t * (ys[i] - ys[i - 1])
    return ys[-1]


def orden_natural(cid):
    dig = "".join(ch for ch in str(cid) if ch.isdigit())
    return (int(dig) if dig else 9999, str(cid))


# =============================================================================
# 3. MOTOR GEOTÉCNICO - CAPACIDAD PORTANTE (Terzaghi, Meyerhof, Vesic)
# =============================================================================
def factores_N(metodo, phi_deg):
    """Factores de capacidad de carga Nc, Nq, Nγ."""
    phi = math.radians(phi_deg)
    if metodo == "Terzaghi":
        if phi_deg < 0.01:
            return 5.70, 1.00, 0.00
        a = math.exp((0.75 * math.pi - phi / 2.0) * math.tan(phi))
        Nq = a ** 2 / (2.0 * math.cos(math.radians(45.0) + phi / 2.0) ** 2)
        Nc = (Nq - 1.0) / math.tan(phi)
        # Nγ de Terzaghi (aproximación de Coduto 2001, error < 3 % frente a la tabla original)
        Ng = 2.0 * (Nq + 1.0) * math.tan(phi) / (1.0 + 0.4 * math.sin(4.0 * phi))
        return Nc, Nq, Ng
    # Meyerhof y Vesic comparten Nq (Prandtl-Reissner) y Nc
    if phi_deg < 0.01:
        return 5.14, 1.00, 0.00
    Nq = math.exp(math.pi * math.tan(phi)) * math.tan(math.radians(45.0) + phi / 2.0) ** 2
    Nc = (Nq - 1.0) / math.tan(phi)
    if metodo == "Meyerhof":
        Ng = (Nq - 1.0) * math.tan(1.4 * phi)
    else:  # Vesic (1973)
        Ng = 2.0 * (Nq + 1.0) * math.tan(phi)
    return Nc, Nq, Ng


def factores_forma_profundidad(metodo, tipo, BL, Df, B, phi_deg, Nc, Nq):
    """Factores de forma (s) y profundidad (d): sc, sq, sγ, dc, dq, dγ."""
    phi = math.radians(phi_deg)
    if metodo == "Terzaghi":
        if tipo == "Continua":
            sc, sg = 1.0, 1.0
        elif tipo == "Cuadrada":
            sc, sg = 1.3, 0.8
        else:  # rectangular
            sc, sg = 1.0 + 0.3 * BL, 1.0 - 0.2 * BL
        return sc, 1.0, sg, 1.0, 1.0, 1.0
    if metodo == "Meyerhof":
        Kp = math.tan(math.radians(45.0) + phi / 2.0) ** 2
        sc = 1.0 + 0.2 * Kp * BL
        sq = sg = (1.0 + 0.1 * Kp * BL) if phi_deg > 10 else 1.0
        dc = 1.0 + 0.2 * math.sqrt(Kp) * Df / B
        dq = dg = (1.0 + 0.1 * math.sqrt(Kp) * Df / B) if phi_deg > 10 else 1.0
        return sc, sq, sg, dc, dq, dg
    # Vesic: forma De Beer (1970), profundidad Hansen (1970)
    sc = 1.0 + BL * (Nq / Nc) if Nc > 0 else 1.0
    sq = 1.0 + BL * math.tan(phi)
    sg = max(0.6, 1.0 - 0.4 * BL)
    k = Df / B if Df / B <= 1.0 else math.atan(Df / B)
    dc = 1.0 + 0.4 * k
    dq = 1.0 + 2.0 * math.tan(phi) * (1.0 - math.sin(phi)) ** 2 * k
    return sc, sq, sg, dc, dq, 1.0


def gamma_saturado(g):
    return g + 0.15 if g < 2.05 else g + 0.08


def correccion_nivel_freatico(g_sobre, g_base, Df, B, nf):
    """Devuelve (q sobrecarga t/m2, γ para el término de Nγ, descripción del caso)."""
    if nf is None or nf <= 0:
        return g_sobre * Df, g_base, "N.F. no encontrado: sin corrección"
    g_sub_s = gamma_saturado(g_sobre) - GAMMA_W
    g_sub_b = gamma_saturado(g_base) - GAMMA_W
    if nf <= Df:
        q = g_sobre * nf + g_sub_s * (Df - nf)
        return q, g_sub_b, "Caso I (N.F. sobre el nivel de cimentación): γ' en q y en Nγ"
    if nf < Df + B:
        d = nf - Df
        g2 = g_sub_b + (d / B) * (g_base - g_sub_b)
        return g_sobre * Df, g2, "Caso II (Df < N.F. < Df+B): γ ponderado en el término Nγ"
    return g_sobre * Df, g_base, "Caso III (N.F. ≥ Df+B): el N.F. no influye"


def capacidad_portante(phi_deg, c_kgcm2, g_base, g_sobre, Df, B, L, tipo, metodo,
                       falla="General", FS=3.0, nf=None):
    """
    qu = c·Nc·Sc·dc + γ·Df·Nq·Sq·dq + 0.5·γ·B·Nγ·Sγ·dγ    (t/m2 internamente)
    qadm = qu / FS     (FS = 3.0 estático, NTE E.050)
    Falla local (Terzaghi): tan φ* = 2/3·tan φ ; c* = 2/3·c
    """
    phi_u, c_u = phi_deg, c_kgcm2
    if falla.startswith("Local"):
        phi_u = math.degrees(math.atan((2.0 / 3.0) * math.tan(math.radians(phi_deg))))
        c_u = (2.0 / 3.0) * c_kgcm2
    c_t = c_u * 10.0                                  # kg/cm2 -> t/m2
    Nc, Nq, Ng = factores_N(metodo, phi_u)
    BL = 0.0 if (tipo == "Continua" or not L) else min(1.0, B / L)
    sc, sq, sg, dc, dq, dg = factores_forma_profundidad(metodo, tipo, BL, Df, B, phi_u, Nc, Nq)
    q, g2, caso = correccion_nivel_freatico(g_sobre, g_base, Df, B, nf)
    t1 = c_t * Nc * sc * dc
    t2 = q * Nq * sq * dq
    t3 = 0.5 * g2 * B * Ng * sg * dg
    qu = t1 + t2 + t3
    qadm = qu / FS
    return dict(metodo=metodo, falla=falla, phi=phi_deg, c=c_kgcm2, phi_u=phi_u, c_u=c_u,
                Nc=Nc, Nq=Nq, Ng=Ng, sc=sc, sq=sq, sg=sg, dc=dc, dq=dq, dg=dg, BL=BL,
                q=q, g2=g2, caso_nf=caso, t1=t1, t2=t2, t3=t3,
                qu_t=qu, qu=qu / 10.0, qadm_t=qadm, qadm=qadm / 10.0, FS=FS,
                Df=Df, B=B, L=L, tipo=tipo, g_base=g_base, g_sobre=g_sobre)


# =============================================================================
# 4. ASENTAMIENTO ELÁSTICO (Schleicher / Steinbrenner) Y DISTORSIÓN ANGULAR
# =============================================================================
# Factor de influencia If para cimentación rígida (Bowles), en función de L/B
IF_LB = [1.0, 1.5, 2.0, 5.0, 10.0, 100.0]
IF_RIGIDO = [0.82, 1.06, 1.20, 1.70, 2.10, 3.40]


def factor_influencia(tipo, relLB=1.0):
    if tipo == "Continua":
        return 2.10
    if tipo == "Cuadrada":
        return 0.82
    return interp(max(1.0, relLB), IF_LB, IF_RIGIDO)


def asentamiento_elastico(q_t, B, Es_kgcm2, mu, If):
    """Se = q·B·(1-μ²)/Es·If  -> devuelve cm (q en t/m2, B en m, Es en kg/cm2)."""
    Es_t = Es_kgcm2 * 10.0
    return q_t * B * (1.0 - mu ** 2) / Es_t * If * 100.0


def q_por_asentamiento(Smax_cm, B, Es_kgcm2, mu, If):
    """Presión (t/m2) que produce el asentamiento máximo tolerable."""
    Es_t = Es_kgcm2 * 10.0
    return (Smax_cm / 100.0) * Es_t / (B * (1.0 - mu ** 2) * If)


# =============================================================================
# 5. CLASIFICACIÓN AUTOMÁTICA SUCS (ASTM D2487 / NTP 339.134) Y AASHTO (NTP 339.135)
# =============================================================================
TAMICES = [("3\"", 76.2), ("2\"", 50.8), ("1 1/2\"", 38.1), ("1\"", 25.4), ("3/4\"", 19.05),
           ("3/8\"", 9.525), ("N°4", 4.75), ("N°10", 2.00), ("N°20", 0.850), ("N°40", 0.425),
           ("N°60", 0.250), ("N°140", 0.106), ("N°200", 0.075)]


def curva_granulometrica(grava, arena, finos):
    """Curva granulométrica (% que pasa) a partir de %G, %S, %F por interpolación log."""
    g = grava or 0.0
    s = arena or 0.0
    f = finos or 0.0
    tot = g + s + f
    if tot <= 0:
        return [100.0] * len(TAMICES)
    g, s, f = 100 * g / tot, 100 * s / tot, 100 * f / tot
    anc_d = [math.log10(76.2), math.log10(19.05), math.log10(4.75), math.log10(0.425), math.log10(0.075)]
    anc_p = [100.0, 100.0 - 0.55 * g, 100.0 - g, f + 0.45 * s, f]
    out = []
    for _, d in TAMICES:
        out.append(round(max(0.0, min(100.0, interp(math.log10(d), anc_d, anc_p))), 1))
    return out


def diametro_efectivo(pasa, pct):
    """Diámetro (mm) correspondiente a un % que pasa (interpolación semilog)."""
    ds = [d for _, d in TAMICES]
    for i in range(len(ds) - 1):
        p1, p2 = pasa[i], pasa[i + 1]
        if p1 >= pct >= p2 and p1 != p2:
            t = (p1 - pct) / (p1 - p2)
            return 10 ** (math.log10(ds[i]) + t * (math.log10(ds[i + 1]) - math.log10(ds[i])))
    return None


def coeficientes_curva(e):
    pasa = curva_granulometrica(e.get("grava"), e.get("arena"), e.get("finos"))
    D10, D30, D60 = (diametro_efectivo(pasa, p) for p in (10, 30, 60))
    cu = e.get("cu")
    cc = e.get("cc")
    if cu is None and D10 and D60:
        cu = D60 / D10
    if cc is None and D10 and D30 and D60:
        cc = D30 ** 2 / (D10 * D60)
    return pasa, D10, D30, D60, cu, cc


def clasificar_sucs(grava, arena, finos, ll, ip, cu=None, cc=None, organico=False):
    g, s, f = grava or 0.0, arena or 0.0, finos or 0.0
    ipv = ip if ip is not None else 0.0
    llv = ll if ll is not None else 0.0
    ipA = 0.73 * (llv - 20.0)
    if f >= 50.0:
        if ll is None or ip is None or ipv <= 0:
            return "ML"
        if llv < 50:
            if organico:
                return "OL"
            if ipv > 7 and ipv >= ipA:
                return "CL"
            if 4 <= ipv <= 7 and ipv >= ipA:
                return "CL-ML"
            return "ML"
        if organico:
            return "OH"
        return "CH" if ipv >= ipA else "MH"
    pre = "G" if g > s else "S"

    def tipo_fino():
        if ll is None or ip is None or ipv < 4 or ipv < ipA:
            return "M"
        if ipv > 7 and ipv >= ipA:
            return "C"
        return "CM"

    def graduacion():
        if cu is None or cc is None:
            return "P"
        lim = 4.0 if pre == "G" else 6.0
        return "W" if (cu >= lim and 1.0 <= cc <= 3.0) else "P"

    if f < 5.0:
        return pre + graduacion()
    if f > 12.0:
        t = tipo_fino()
        return pre + "C-" + pre + "M" if t == "CM" else pre + t
    t = tipo_fino()
    t = "M" if t == "CM" else t
    return pre + graduacion() + "-" + pre + t


def clasificar_aashto(p10, p40, p200, ll, ip):
    llv = ll if ll is not None else 0.0
    ipv = ip if ip is not None else 0.0
    F = p200

    def gi_total():
        a = max(0.0, min(40.0, F - 35.0))
        b = max(0.0, min(40.0, F - 15.0))
        c = max(0.0, min(20.0, llv - 40.0))
        d = max(0.0, min(20.0, ipv - 10.0))
        return max(0, round(0.2 * a + 0.005 * a * c + 0.01 * b * d))

    if F <= 35:
        if p10 <= 50 and p40 <= 30 and F <= 15 and ipv <= 6:
            return "A-1-a(0)"
        if p40 <= 50 and F <= 25 and ipv <= 6:
            return "A-1-b(0)"
        if p40 > 50 and F <= 10 and (ip is None or ipv == 0):
            return "A-3(0)"
        if llv <= 40 and ipv <= 10:
            return "A-2-4(0)"
        if llv > 40 and ipv <= 10:
            return "A-2-5(0)"
        gi = max(0, round(0.01 * max(0, min(40, F - 15)) * max(0, min(20, ipv - 10))))
        return f"A-2-6({gi})" if llv <= 40 else f"A-2-7({gi})"
    gi = gi_total()
    if llv <= 40 and ipv <= 10:
        return f"A-4({gi})"
    if llv > 40 and ipv <= 10:
        return f"A-5({gi})"
    if llv <= 40:
        return f"A-6({gi})"
    return f"A-7-5({gi})" if ipv <= llv - 30 else f"A-7-6({gi})"


def autoclasificar_estrato(e):
    if sucs_base(e.get("sucs")) in ("RELL", "Roca", "Pt"):
        return e.get("sucs"), e.get("aashto")
    pasa, D10, D30, D60, cu, cc = coeficientes_curva(e)
    s = clasificar_sucs(e.get("grava"), e.get("arena"), e.get("finos"), e.get("ll"), e.get("ip"), cu, cc)
    a = clasificar_aashto(pasa[7], pasa[9], pasa[12], e.get("ll"), e.get("ip"))
    return s, a


# =============================================================================
# 6. PARÁMETROS SÍSMICOS NTE E.030-2018 Y AGRESIVIDAD QUÍMICA (E.060 / ACI 318)
# =============================================================================
Z_E030 = {1: 0.10, 2: 0.25, 3: 0.35, 4: 0.45}
S_E030 = {4: {"S0": 0.80, "S1": 1.00, "S2": 1.05, "S3": 1.10},
          3: {"S0": 0.80, "S1": 1.00, "S2": 1.15, "S3": 1.20},
          2: {"S0": 0.80, "S1": 1.00, "S2": 1.20, "S3": 1.40},
          1: {"S0": 0.80, "S1": 1.00, "S2": 1.60, "S3": 2.00}}
TP_E030 = {"S0": 0.30, "S1": 0.40, "S2": 0.60, "S3": 1.00}
TL_E030 = {"S0": 3.00, "S1": 2.50, "S2": 2.00, "S3": 1.60}
U_E030 = {"A": 1.50, "B": 1.30, "C": 1.00}
PERFIL_DESC = {"S0": "Roca dura (Vs > 1500 m/s)",
               "S1": "Roca o suelos muy rígidos (500 < Vs ≤ 1500 m/s)",
               "S2": "Suelos intermedios (180 < Vs ≤ 500 m/s)",
               "S3": "Suelos blandos (Vs ≤ 180 m/s)"}
CATEGORIA_DESC = {"A": "Edificaciones esenciales", "B": "Edificaciones importantes",
                  "C": "Edificaciones comunes"}


def factor_C(T, Tp, Tl):
    if T < Tp:
        return 2.5
    if T < Tl:
        return 2.5 * Tp / T
    return 2.5 * Tp * Tl / T ** 2


def parametros_sismicos(zona, perfil, categoria, pisos, h_piso=2.80):
    zona = int(zona)
    Z = Z_E030[zona]
    S = S_E030[zona][perfil]
    Tp, Tl = TP_E030[perfil], TL_E030[perfil]
    U = U_E030.get(categoria, 1.0)
    hn = max(1, int(pisos)) * h_piso
    Ct = 60 if int(pisos) <= 5 else 45
    T = hn / Ct
    C = factor_C(T, Tp, Tl)
    return dict(zona=zona, perfil=perfil, Z=Z, S=S, Tp=Tp, Tl=Tl, U=U, categoria=categoria,
                hn=hn, Ct=Ct, T=T, C=C, ZUCS=Z * U * C * S)


GRADOS = ["Despreciable", "Moderado", "Severo", "Muy Severo"]


def agresividad_quimica(so4, cl, sst, ph):
    """Grado de ataque al concreto y recomendación de cemento."""
    so4 = so4 or 0.0
    cl = cl or 0.0
    sst = sst or 0.0
    ph = ph if ph is not None else 7.0
    # Sulfatos solubles en agua (E.060 Tabla 4.4 / ACI 318): 0.1 % = 1000 ppm
    if so4 < 1000:
        gs = 0
    elif so4 < 2000:
        gs = 1
    elif so4 < 20000:
        gs = 2
    else:
        gs = 3
    cemento = ["Portland Tipo I (o Tipo GU / adicionado IP, ICo)",
               "Portland Tipo II o adicionado MS (IP(MS), IS(MS), ICo(MS))",
               "Portland Tipo V o adicionado HS",
               "Portland Tipo V + puzolana (adicionado HS)"][gs]
    ac = [0.55, 0.50, 0.45, 0.45][gs]
    fc = [210, 280, 315, 315][gs]
    # Cloruros (criterio práctico usual en EMS: > 6000 ppm perjudicial)
    gc = 0 if cl <= 1000 else 1 if cl <= 6000 else 2 if cl <= 15000 else 3
    # Sales solubles totales (> 15000 ppm perjudicial: lixiviación)
    gst = 0 if sst <= 5000 else 1 if sst <= 15000 else 2 if sst <= 30000 else 3
    # pH (suelo ácido)
    gph = 0 if ph >= 6.5 else 1 if ph >= 5.5 else 2 if ph >= 4.5 else 3
    gtot = max(gs, gc, gst, gph)
    if gc >= 2 or gst >= 2 or gph >= 2:
        ac = min(ac, 0.45)
        fc = max(fc, 315)
    medidas = []
    if gs >= 1:
        medidas.append("Emplear cemento resistente a sulfatos según lo indicado.")
    if gc >= 2:
        medidas.append("Cloruros perjudiciales: recubrimiento mínimo de 7.5 cm, concreto de baja "
                       "permeabilidad y aditivo inhibidor de corrosión.")
    elif gc == 1:
        medidas.append("Cloruros moderados: respetar recubrimiento de 7.5 cm (E.060) y buena compactación.")
    if gst >= 2:
        medidas.append("Sales solubles perjudiciales: impermeabilizar caras en contacto con el suelo "
                       "(emulsión asfáltica) y colocar solado.")
    if gph >= 1:
        medidas.append("Suelo ácido: concreto denso y protección superficial con pintura bituminosa.")
    if not medidas:
        medidas.append("No se requieren medidas especiales; usar buenas prácticas de curado y compactación.")
    return dict(so4=so4, cl=cl, sst=sst, ph=ph, g_so4=GRADOS[gs], g_cl=GRADOS[gc], g_sst=GRADOS[gst],
                g_ph=GRADOS[gph], grado=GRADOS[gtot], cemento=cemento, ac=ac, fc=fc, medidas=medidas)


def coeficientes_empuje(phi_deg, c=0.0):
    phi = math.radians(phi_deg)
    Ka = math.tan(math.radians(45) - phi / 2) ** 2
    Kp = math.tan(math.radians(45) + phi / 2) ** 2
    K0 = 1.0 - math.sin(phi)
    mu = math.tan(2.0 / 3.0 * phi)
    return dict(Ka=Ka, Kp=Kp, K0=K0, mu=mu)


# =============================================================================
# 7. ANÁLISIS INTEGRAL (matriz Df x B, envolvente de calicatas, caso de diseño)
# =============================================================================
CFG_DEFECTO = dict(metodo="Terzaghi", falla="General", tipo="Cuadrada", relLB=1.5, FS=3.0,
                   es_auto=True, Es=150.0, mu=0.30, smax=SMAX_E050, luz=5.0, kdif=0.75,
                   calicata="Envolvente", manual=False, phi=30.0, c=0.0, gamma=1.70,
                   Df=1.50, B=1.50)


def calicatas_de(estratos):
    ids = sorted({e["cal"] for e in estratos}, key=orden_natural)
    return ids


def estratos_de(estratos, cal):
    return sorted([e for e in estratos if e["cal"] == cal], key=lambda e: e["ini"])


def estrato_en(estr_cal, z):
    for e in estr_cal:
        if e["ini"] <= z < e["fin"]:
            return e
    return estr_cal[-1] if estr_cal else None


def gamma_promedio(estr_cal, Df):
    acum, h = 0.0, 0.0
    for e in estr_cal:
        a, b = e["ini"], min(e["fin"], Df)
        if b > a:
            acum += (e.get("gamma") or 1.6) * (b - a)
            h += b - a
    if Df > h and estr_cal:
        acum += (estr_cal[-1].get("gamma") or 1.6) * (Df - h)
        h = Df
    return acum / h if h > 0 else 1.6


def dims_cimiento(tipo, B, relLB):
    if tipo == "Continua":
        return None, "Continua"
    if tipo == "Cuadrada":
        return B, "Cuadrada"
    return B * relLB, "Rectangular"


def analizar_punto(estr_cal, nf, Df, B, tipo, cfg):
    base = estrato_en(estr_cal, Df + 0.005)
    if cfg["manual"]:
        phi, c, g_base, g_sobre = cfg["phi"], cfg["c"], cfg["gamma"], cfg["gamma"]
    else:
        phi = base.get("phi") if base.get("phi") is not None else SUCS_DB[sucs_base(base["sucs"])]["phi"]
        c = base.get("c") if base.get("c") is not None else 0.0
        g_base = base.get("gamma") or 1.7
        g_sobre = gamma_promedio(estr_cal, Df)
    L, tnom = dims_cimiento(tipo, B, cfg["relLB"])
    res = {m: capacidad_portante(phi, c, g_base, g_sobre, Df, B, L, tipo, m, cfg["falla"], cfg["FS"], nf)
           for m in METODOS}
    rc = res[cfg["metodo"]]
    db = SUCS_DB[sucs_base(base["sucs"])]
    if cfg["es_auto"]:
        Es, mu = db["Es"], db["mu"]
    else:
        Es, mu = cfg["Es"], cfg["mu"]
    If = factor_influencia(tipo, cfg["relLB"])
    se_corte = asentamiento_elastico(rc["qadm_t"], B, Es, mu, If)
    q_s = q_por_asentamiento(cfg["smax"], B, Es, mu, If)
    # Distorsión angular: δ = kdif·Se ≤ L/500  ->  Se límite = (L/500)/kdif
    se_lim_dist = (cfg["luz"] * DISTORSION_MAX / cfg["kdif"]) * 100.0
    q_d = q_por_asentamiento(se_lim_dist, B, Es, mu, If)
    q_dis_t = min(rc["qadm_t"], q_s, q_d)
    se_dis = asentamiento_elastico(q_dis_t, B, Es, mu, If)
    delta = cfg["kdif"] * se_dis
    alfa = (delta / 100.0) / cfg["luz"] if cfg["luz"] > 0 else 0.0
    if q_dis_t >= rc["qadm_t"] - 1e-9:
        gob = "Corte"
    elif q_s <= q_d:
        gob = "Asentamiento"
    else:
        gob = "Distorsión angular"
    return dict(cal=base["cal"], estrato=base["est"], sucs=base["sucs"], phi=phi, c=c, g_base=g_base,
                g_sobre=g_sobre, Df=Df, B=B, L=L, tipo=tnom, res=res, Es=Es, mu=mu, If=If,
                qadm_corte=rc["qadm"], qu=rc["qu"], se_corte=se_corte, q_asent=q_s / 10.0,
                q_dist=q_d / 10.0, se_lim_dist=se_lim_dist,
                qadm=q_dis_t / 10.0, se=se_dis, delta=delta, alfa=alfa, gobierna=gob,
                ok_se=se_dis <= cfg["smax"] + 1e-9, ok_alfa=alfa <= DISTORSION_MAX + 1e-12,
                es_relleno=sucs_base(base["sucs"]) in ("RELL", "Pt", "OL", "OH"))


def analizar_envolvente(estratos, cal_info, nf, Df, B, tipo, cfg):
    cals = calicatas_de(estratos)
    if cfg["calicata"] not in ("Envolvente", "", None) and cfg["calicata"] in cals:
        cals = [cfg["calicata"]]
    peor = None
    for cid in cals:
        ec = estratos_de(estratos, cid)
        if not ec:
            continue
        nfc = (cal_info.get(cid, {}).get("nf") if cal_info.get(cid, {}).get("nf") is not None else nf)
        r = analizar_punto(ec, nfc, Df, B, tipo, cfg)
        if peor is None or r["qadm"] < peor["qadm"]:
            peor = r
    return peor


def calcular_todo(proyecto, cal_info, estratos, cfg=None):
    cfg = dict(CFG_DEFECTO, **(cfg or {}))
    nf = num(proyecto.get("nf"))
    nf = nf if (nf is not None and nf > 0) else None
    filas = []
    for Df in DF_MATRIZ:
        combos = [(B_CORRIDO, "Continua")] + [(b, cfg["tipo"]) for b in B_ZAPATAS]
        for B, tipo in combos:
            r = analizar_envolvente(estratos, cal_info, nf, Df, B, tipo, cfg)
            if r:
                filas.append(r)
    diseno = analizar_envolvente(estratos, cal_info, nf, cfg["Df"], cfg["B"], cfg["tipo"], cfg)
    corrido = analizar_envolvente(estratos, cal_info, nf, cfg["Df"], B_CORRIDO, "Continua", cfg)
    sismo = parametros_sismicos(proyecto.get("zona", 4), proyecto.get("perfil", "S2"),
                                proyecto.get("categoria", "C"), num(proyecto.get("pisos"), 2))
    quim = agresividad_quimica(num(proyecto.get("so4")), num(proyecto.get("cl")),
                               num(proyecto.get("sst")), num(proyecto.get("ph")))
    empujes = []
    for e in sorted(estratos, key=lambda x: (orden_natural(x["cal"]), x["ini"])):
        phi = e.get("phi") if e.get("phi") is not None else SUCS_DB[sucs_base(e["sucs"])]["phi"]
        k = coeficientes_empuje(phi, e.get("c") or 0)
        empujes.append(dict(cal=e["cal"], est=e["est"], sucs=e["sucs"], phi=phi, c=e.get("c") or 0,
                            gamma=e.get("gamma") or 1.7, **k))
    return dict(cfg=cfg, nf=nf, filas=filas, diseno=diseno, corrido=corrido, sismo=sismo,
                quim=quim, empujes=empujes, especiales=problemas_especiales(proyecto, estratos, nf))


def problemas_especiales(proyecto, estratos, nf):
    zona = int(proyecto.get("zona", 4))
    out = []
    # Licuación
    sus = [e for e in estratos if nf and e["fin"] > nf and sucs_base(e["sucs"]) in ("SP", "SM", "SW", "ML")]
    if sus and zona >= 3:
        out.append(("Licuación de suelos", True,
                    f"Se identificaron estratos granulares finos ({', '.join(sorted({e['sucs'] for e in sus}))}) "
                    f"bajo el nivel freático en zona sísmica {zona}. Se recomienda una evaluación específica del "
                    "potencial de licuación (Seed & Idriss / NCEER) con ensayos SPT o DPL complementarios."))
    else:
        out.append(("Licuación de suelos", False,
                    "No se identifican arenas finas sueltas o limos no plásticos saturados dentro de la "
                    "profundidad explorada; el potencial de licuación se considera bajo."))
    # Colapso
    col = [e for e in estratos if (e.get("gamma") or 2) < 1.55 and (e.get("w") or 99) < 10
           and sucs_base(e["sucs"]) in ("SM", "ML", "SC", "SP")]
    if col:
        out.append(("Suelos colapsables", True,
                    "Los estratos " + ", ".join(f"{e['cal']}/{e['est']}" for e in col) +
                    " presentan baja densidad y baja humedad; son potencialmente colapsables. Verificar con "
                    "ensayo de potencial de colapso (ASTM D5333) antes de cimentar sobre ellos."))
    else:
        out.append(("Suelos colapsables", False,
                    "No se identifican suelos de baja densidad seca y baja humedad susceptibles de colapso "
                    "al nivel de cimentación."))
    # Expansividad
    exp = [e for e in estratos if (e.get("ip") or 0) >= 20]
    if exp:
        ipmax = max(e.get("ip") or 0 for e in exp)
        nivel = "medio" if ipmax < 25 else "alto" if ipmax < 35 else "muy alto"
        out.append(("Suelos expansivos", True,
                    f"Se registran suelos finos con IP de hasta {ipmax:.0f} %, potencial de expansión {nivel}. "
                    "Verificar con ensayo de expansión libre/controlada (ASTM D4546) y evitar cimentar sobre ellos."))
    else:
        out.append(("Suelos expansivos", False,
                    "Los suelos finos encontrados presentan índice plástico menor a 20 %; el potencial de "
                    "expansión se considera bajo."))
    # Relleno
    rel = [e for e in estratos if sucs_base(e["sucs"]) == "RELL"]
    if rel:
        pmax = max(e["fin"] for e in rel)
        out.append(("Rellenos no controlados", True,
                    f"Se encontró relleno no controlado hasta {pmax:.2f} m. No debe usarse como suelo de apoyo; "
                    "deberá ser eliminado íntegramente (NTE E.050) y reemplazado, de ser necesario, por relleno "
                    "controlado o falso cimiento de concreto ciclópeo."))
    # Nivel freático
    if nf:
        out.append(("Nivel freático", True,
                    f"Se detectó el nivel freático a {nf:.2f} m. Prever bombeo/deprimido durante la excavación "
                    "y considerar la subpresión en el diseño de las estructuras enterradas."))
    else:
        out.append(("Nivel freático", False, "No se encontró nivel freático hasta la profundidad explorada."))
    return [(t, d, r) for t, r, d in out]


# =============================================================================
# 8. DATOS DE EJEMPLO (3 calicatas de 3.00 m con estratos SM, GP y CL + fotos)
# =============================================================================
def proyecto_ejemplo():
    return dict(nombre="Vivienda Multifamiliar de 4 Pisos con Azotea",
                solicitante="Inversiones Inmobiliarias Los Andes S.A.C. (ejemplo)",
                distrito="Carabayllo", provincia="Lima", region="Lima",
                categoria="C", pisos="4", zona="4", perfil="S2", nf="2.60",
                so4="850", cl="420", sst="1650", ph="7.40",
                ingeniero="Ing. Civil Responsable (ejemplo)", cip="000000",
                area="160.00", tipo_cim="Zapatas aisladas y cimientos corridos")


def estratos_ejemplo():
    rell = ("Relleno no controlado conformado por limo arenoso con restos de raíces, fragmentos de "
            "ladrillo y material orgánico; color marrón oscuro, suelto, húmedo.")
    sm = ("Arena limosa con escasa grava, de grano fino a medio, color beige, medianamente densa, "
          "ligeramente húmeda, no plástica.")
    gp = ("Grava mal graduada con arena, de forma subredondeada, TM 3\", color gris, densa, húmeda; "
          "presencia aislada de bolonería (≤ 15 %).")
    cl = ("Arcilla inorgánica de baja plasticidad con arena, color marrón claro, consistencia media a "
          "firme, húmeda.")

    def R(cal, ini, fin):
        return dict(cal=cal, est="E-1", ini=ini, fin=fin, muestra="-", sucs="RELL", aashto="-",
                    grava=15, arena=45, finos=40, w=8.5, ll=None, ip=None, gamma=1.45, phi=24.0, c=0.0,
                    cu=None, cc=None, desc=rell)

    def SM_(cal, est, mm, ini, fin, w, g, phi):
        return dict(cal=cal, est=est, ini=ini, fin=fin, muestra=mm, sucs="SM", aashto="A-2-4(0)",
                    grava=8, arena=64, finos=28, w=w, ll=None, ip=None, gamma=g, phi=phi, c=0.02,
                    cu=None, cc=None, desc=sm)

    def GP_(cal, est, mm, ini, fin, w, g, phi):
        return dict(cal=cal, est=est, ini=ini, fin=fin, muestra=mm, sucs="GP", aashto="A-1-a(0)",
                    grava=66, arena=30, finos=4, w=w, ll=None, ip=None, gamma=g, phi=phi, c=0.0,
                    cu=3.1, cc=0.8, desc=gp)

    return [
        R("C-1", 0.00, 0.40),
        SM_("C-1", "E-2", "M-1", 0.40, 1.60, 6.8, 1.68, 30.5),
        GP_("C-1", "E-3", "M-2", 1.60, 3.00, 4.2, 1.98, 36.0),
        R("C-2", 0.00, 0.30),
        dict(cal="C-2", est="E-2", ini=0.30, fin=1.50, muestra="M-1", sucs="CL", aashto="A-6(7)",
             grava=2, arena=30, finos=68, w=18.4, ll=34.0, ip=14.0, gamma=1.78, phi=18.0, c=0.22,
             cu=None, cc=None, desc=cl),
        GP_("C-2", "E-3", "M-2", 1.50, 3.00, 4.6, 2.00, 36.5),
        R("C-3", 0.00, 0.50),
        SM_("C-3", "E-2", "M-1", 0.50, 1.70, 7.2, 1.66, 30.0),
        GP_("C-3", "E-3", "M-2", 1.70, 3.00, 4.0, 1.97, 35.5),
    ]


def info_ejemplo():
    hoy = datetime.date.today().strftime("%d/%m/%Y")
    return {"C-1": dict(este=279850.0, norte=8681200.0, cota=245.30, nf=None, fecha=hoy, fotos=[]),
            "C-2": dict(este=279872.0, norte=8681214.0, cota=245.55, nf=None, fecha=hoy, fotos=[]),
            "C-3": dict(este=279893.0, norte=8681229.0, cota=245.80, nf=None, fecha=hoy, fotos=[])}


def generar_fotos_ejemplo(carpeta, estratos, cal_info):
    """Crea fotografías sintéticas de las calicatas (para probar el Anexo Fotográfico)."""
    os.makedirs(carpeta, exist_ok=True)
    for cid in calicatas_de(estratos):
        ec = estratos_de(estratos, cid)
        fotos = []
        for k, vista in enumerate(["Vista general de la excavación", "Detalle de la estratigrafía"]):
            W, H = 1200, 900
            img = Image.new("RGB", (W, H), (150, 185, 225))
            d = ImageDraw.Draw(img)
            d.rectangle([0, 230, W, H], fill=(160, 140, 110))
            x1, x2, y0 = (300, 900, 260) if k == 0 else (150, 1050, 120)
            esc = (H - y0 - 60) / max(e["fin"] for e in ec)
            rnd = random.Random(sum(map(ord, cid)) + k)
            for e in ec:
                rgb = SUCS_DB[sucs_base(e["sucs"])]["rgb"]
                ya, yb = y0 + e["ini"] * esc, y0 + e["fin"] * esc
                d.rectangle([x1, ya, x2, yb], fill=tuple(max(0, v - 40) for v in rgb))
                for _ in range(int((yb - ya) * (x2 - x1) / 900)):
                    px, py = rnd.uniform(x1, x2), rnd.uniform(ya, yb)
                    r = rnd.uniform(2, 9) if sucs_base(e["sucs"]) in ("GP", "GW") else rnd.uniform(1, 2.5)
                    tono = rnd.randint(-35, 35)
                    col = tuple(max(0, min(255, v - 40 + tono)) for v in rgb)
                    d.ellipse([px - r, py - r, px + r, py + r], fill=col)
                d.text((x2 + 12, (ya + yb) / 2 - 8), f"{e['est']} {e['sucs']}", fill=(255, 255, 255),
                       font=_fuente_pil(26, True))
            # wincha
            d.rectangle([x1 - 40, y0, x1 - 20, H - 60], fill=(250, 210, 30))
            for i in range(0, int(ec[-1]["fin"] * 10) + 1):
                yy = y0 + i * 0.1 * esc
                d.line([x1 - 40, yy, x1 - 28 if i % 5 else x1 - 20, yy], fill=(0, 0, 0), width=2)
            d.rectangle([0, 0, W, 70], fill=(30, 30, 30))
            d.text((20, 18), f"Calicata {cid} - {vista}", fill=(255, 255, 255), font=_fuente_pil(32, True))
            p = os.path.join(carpeta, f"{cid}_foto{k + 1}.jpg")
            img.save(p, quality=88)
            fotos.append(dict(path=p, desc=f"Calicata {cid}: {vista.lower()}."))
        cal_info.setdefault(cid, {})["fotos"] = fotos
    return cal_info


# =============================================================================
# 9. IMPORTACIÓN DESDE EXCEL Y PLANTILLA
# =============================================================================
ALIAS = {
    "cal": ["calicata", "cal", "punto", "sondaje", "calicatan"],
    "est": ["estrato", "est", "capa", "horizonte"],
    "ini": ["profini", "profundidadinicial", "desde", "ini", "profinicial", "de"],
    "fin": ["proffin", "profundidadfinal", "hasta", "fin", "proffinal", "a"],
    "muestra": ["muestra", "mx", "nmuestra"],
    "sucs": ["sucs", "clasificacionsucs", "clasifsucs"],
    "aashto": ["aashto", "clasificacionaashto", "clasifaashto"],
    "grava": ["grava", "gravas", "%grava", "pgrava"],
    "arena": ["arena", "arenas", "parena"],
    "finos": ["finos", "pasa200", "n200", "pfinos"],
    "w": ["humedad", "w", "wn", "humedadnatural", "contenidodehumedad"],
    "ll": ["ll", "limiteliquido"],
    "lp": ["lp", "limiteplastico"],
    "ip": ["ip", "indiceplastico", "indicedeplasticidad"],
    "gamma": ["gamma", "pesoespecifico", "pesounitario", "densidad", "pesovolumetrico"],
    "phi": ["phi", "angulodefriccion", "angulo", "friccion", "angulofriccion"],
    "c": ["c", "cohesion"],
    "cu": ["cu"],
    "cc": ["cc"],
    "desc": ["descripcion", "descripcionlitologica", "litologia", "desc"],
    "este": ["este", "x", "coordeste"],
    "norte": ["norte", "y", "coordnorte"],
    "cota": ["cota", "elevacion", "z", "altitud"],
    "nf": ["nf", "nivelfreatico"],
}


def _norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    s = s.split("(")[0]
    return "".join(ch for ch in s if ch.isalnum() or ch == "%")


def _mapear_columnas(cols):
    m = {}
    for col in cols:
        n = _norm(col)
        for key, al in ALIAS.items():
            if key in m:
                continue
            if n in al or n.replace("%", "") in al:
                m[key] = col
                break
    return m


def importar_excel(path):
    if pd is None:
        raise RuntimeError("La librería pandas no está instalada.")
    hojas = pd.read_excel(path, sheet_name=None)
    estratos, info = [], {}
    for nombre, df in hojas.items():
        df = df.dropna(how="all")
        m = _mapear_columnas(df.columns)
        if "cal" in m and "ini" in m and "fin" in m:
            for _, row in df.iterrows():
                cid = str(row[m["cal"]]).strip()
                if not cid or cid.lower() == "nan":
                    continue
                g = lambda k, d=None: num(row[m[k]], d) if k in m else d
                s = lambda k, d="": (str(row[m[k]]).strip() if k in m and str(row[m[k]]) != "nan" else d)
                ll, lp, ip = g("ll"), g("lp"), g("ip")
                if ip is None and ll is not None and lp is not None:
                    ip = ll - lp
                e = dict(cal=cid, est=s("est", "E-?"), ini=g("ini", 0.0), fin=g("fin", 0.0),
                         muestra=s("muestra", "-"), sucs=s("sucs", "SM"), aashto=s("aashto", "-"),
                         grava=g("grava"), arena=g("arena"), finos=g("finos"), w=g("w"), ll=ll, ip=ip,
                         gamma=g("gamma", 1.70), phi=g("phi"), c=g("c", 0.0), cu=g("cu"), cc=g("cc"),
                         desc=s("desc", ""))
                if e["phi"] is None:
                    e["phi"] = SUCS_DB[sucs_base(e["sucs"])]["phi"]
                if e["fin"] > e["ini"]:
                    estratos.append(e)
                    d = info.setdefault(cid, dict(este=None, norte=None, cota=None, nf=None, fotos=[],
                                                  fecha=datetime.date.today().strftime("%d/%m/%Y")))
                    for k in ("este", "norte", "cota", "nf"):
                        if g(k) is not None:
                            d[k] = g(k)
        elif "cal" in m and ("este" in m or "cota" in m):
            for _, row in df.iterrows():
                cid = str(row[m["cal"]]).strip()
                d = info.setdefault(cid, dict(este=None, norte=None, cota=None, nf=None, fotos=[],
                                              fecha=datetime.date.today().strftime("%d/%m/%Y")))
                for k in ("este", "norte", "cota", "nf"):
                    if k in m and num(row[m[k]]) is not None:
                        d[k] = num(row[m[k]])
    if not estratos:
        raise ValueError("No se encontró una hoja con columnas Calicata, Prof_Ini y Prof_Fin.")
    return estratos, info


def crear_plantilla_excel(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Estratigrafia"
    cab = ["Calicata", "Estrato", "Prof_Ini", "Prof_Fin", "Muestra", "SUCS", "AASHTO", "Grava", "Arena",
           "Finos", "Humedad", "LL", "LP", "IP", "Gamma", "Phi", "Cohesion", "Cu", "Cc", "Descripcion"]
    ws.append(cab)
    for e in estratos_ejemplo():
        lp = (e["ll"] - e["ip"]) if (e["ll"] is not None and e["ip"] is not None) else "NP"
        ws.append([e["cal"], e["est"], e["ini"], e["fin"], e["muestra"], e["sucs"], e["aashto"], e["grava"],
                   e["arena"], e["finos"], e["w"], e["ll"] if e["ll"] is not None else "NP", lp,
                   e["ip"] if e["ip"] is not None else "NP", e["gamma"], e["phi"], e["c"], e["cu"], e["cc"],
                   e["desc"]])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
    for i, w in enumerate([10, 9, 9, 9, 9, 8, 10, 8, 8, 8, 9, 6, 6, 6, 8, 7, 9, 6, 6, 80], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws2 = wb.create_sheet("Calicatas")
    ws2.append(["Calicata", "Este", "Norte", "Cota", "NF"])
    for cid, d in info_ejemplo().items():
        ws2.append([cid, d["este"], d["norte"], d["cota"], "NP"])
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
    ws3 = wb.create_sheet("Instrucciones")
    for t in ["Unidades: profundidades en m, humedad/LL/LP/IP/Grava/Arena/Finos en %, Gamma en g/cm3,",
              "Phi en grados sexagesimales, Cohesion en kg/cm2. Escriba NP para suelos no plásticos.",
              "SUCS admitidos: GW GP GM GC SW SP SM SC ML CL OL MH CH OH Pt Roca RELL (relleno) y duales (SP-SM).",
              "La hoja 'Calicatas' (opcional) contiene coordenadas UTM, cota del terreno y N.F. por calicata."]:
        ws3.append([t])
    ws3.column_dimensions["A"].width = 120
    wb.save(path)


# =============================================================================
# 10. LIENZOS DE DIBUJO (DXF con ezdxf  y  PNG con Pillow) - misma geometría
#     Unidades: milímetros de papel, eje Y hacia arriba.
# =============================================================================
CAPA_CUADROS = "EMS-CUADROS"
CAPA_SUCS = "EMS-SIMBOLOGIA-SUCS"
CAPA_COTAS = "EMS-COTAS"
CAPA_TEXTOS = "EMS-TEXTOS"

ALINEA_DXF = {"BL": TextEntityAlignment.BOTTOM_LEFT, "BC": TextEntityAlignment.BOTTOM_CENTER,
              "BR": TextEntityAlignment.BOTTOM_RIGHT, "ML": TextEntityAlignment.MIDDLE_LEFT,
              "MC": TextEntityAlignment.MIDDLE_CENTER, "MR": TextEntityAlignment.MIDDLE_RIGHT,
              "TL": TextEntityAlignment.TOP_LEFT, "TC": TextEntityAlignment.TOP_CENTER,
              "TR": TextEntityAlignment.TOP_RIGHT}

_FUENTES = {}


def _fuente_pil(px, negrita=False):
    key = (int(px), negrita)
    if key in _FUENTES:
        return _FUENTES[key]
    nombres = (["arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf"] if negrita
               else ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf"])
    rutas = ["", "C:/Windows/Fonts/", "/usr/share/fonts/truetype/dejavu/",
             "/usr/share/fonts/truetype/liberation/", "/Library/Fonts/"]
    f = None
    for n in nombres:
        for r in rutas:
            try:
                f = ImageFont.truetype(r + n, max(6, int(px)))
                break
            except Exception:
                continue
        if f:
            break
    if f is None:
        try:
            f = ImageFont.load_default(size=max(6, int(px)))
        except Exception:
            f = ImageFont.load_default()
    _FUENTES[key] = f
    return f


class LienzoDXF:
    def __init__(self, doc):
        self.doc = doc
        self.msp = doc.modelspace()

    def linea(self, p1, p2, capa=CAPA_CUADROS, color=None, tipo=None, grosor=None, esc_tipo=2.5):
        at = {"layer": capa}
        if color is not None:
            at["color"] = color
        if tipo:
            at["linetype"] = tipo
            at["ltscale"] = esc_tipo
        if grosor:
            at["lineweight"] = grosor
        self.msp.add_line(p1, p2, dxfattribs=at)

    def rect(self, x1, y1, x2, y2, capa=CAPA_CUADROS, grosor=None, color=None):
        at = {"layer": capa}
        if grosor:
            at["lineweight"] = grosor
        if color is not None:
            at["color"] = color
        self.msp.add_lwpolyline([(x1, y1), (x2, y1), (x2, y2), (x1, y2)], close=True, dxfattribs=at)

    def circulo(self, c, r, capa=CAPA_SUCS, color=None):
        at = {"layer": capa}
        if color is not None:
            at["color"] = color
        self.msp.add_circle(c, r, dxfattribs=at)

    def relleno(self, pts, color=7, capa=CAPA_COTAS):
        h = self.msp.add_hatch(color=color, dxfattribs={"layer": capa})
        h.paths.add_polyline_path(pts, is_closed=True)

    def fondo(self, x1, y1, x2, y2, rgb):
        pass  # en CAD se mantiene sólo el patrón vectorial

    def texto(self, s, x, y, h=2.0, al="BL", capa=CAPA_TEXTOS, rot=0, negrita=False, color=None):
        at = {"layer": capa, "height": h, "style": "EMS_B" if negrita else "EMS", "rotation": rot}
        if color is not None:
            at["color"] = color
        t = self.msp.add_text(str(s), dxfattribs=at)
        t.set_placement((x, y), align=ALINEA_DXF[al])

    def mtexto(self, s, x, y, h=1.8, ancho=60, capa=CAPA_TEXTOS):
        m = self.msp.add_mtext(str(s), dxfattribs={"layer": capa, "char_height": h, "width": ancho,
                                                  "style": "EMS"})
        m.set_location((x, y), attachment_point=MTextEntityAlignment.TOP_LEFT)

    def primitivas(self, prims, capa=CAPA_SUCS, color=None):
        for p in prims:
            if p[0] == "l":
                self.linea((p[1], p[2]), (p[3], p[4]), capa=capa, color=color)
            else:
                self.circulo((p[1], p[2]), p[3], capa=capa, color=color)


class LienzoPIL:
    def __init__(self, ancho_mm, alto_mm, x0=0.0, ytop=0.0, px_mm=8.0):
        self.s = px_mm
        self.x0, self.ytop = x0, ytop
        self.img = Image.new("RGB", (int(ancho_mm * px_mm) + 2, int(alto_mm * px_mm) + 2), "white")
        self.d = ImageDraw.Draw(self.img)

    def P(self, x, y):
        return ((x - self.x0) * self.s, (self.ytop - y) * self.s)

    def _rgb(self, color):
        return ACI_RGB.get(color, (0, 0, 0)) if color is not None else (0, 0, 0)

    def linea(self, p1, p2, capa=None, color=None, tipo=None, grosor=None, esc_tipo=2.5):
        w = max(1, int((grosor or 18) / 100.0 * self.s))
        a, b = self.P(*p1), self.P(*p2)
        col = self._rgb(color)
        if tipo:
            L = math.hypot(b[0] - a[0], b[1] - a[1])
            if L < 1:
                return
            dash, gap = 3.0 * self.s, 1.2 * self.s
            t = 0.0
            while t < L:
                t2 = min(L, t + dash)
                self.d.line([(a[0] + (b[0] - a[0]) * t / L, a[1] + (b[1] - a[1]) * t / L),
                             (a[0] + (b[0] - a[0]) * t2 / L, a[1] + (b[1] - a[1]) * t2 / L)], fill=col, width=w)
                t = t2 + gap
        else:
            self.d.line([a, b], fill=col, width=w)

    def rect(self, x1, y1, x2, y2, capa=None, grosor=None, color=None):
        for p, q in (((x1, y1), (x2, y1)), ((x2, y1), (x2, y2)), ((x2, y2), (x1, y2)), ((x1, y2), (x1, y1))):
            self.linea(p, q, grosor=grosor, color=color)

    def circulo(self, c, r, capa=None, color=None):
        x, y = self.P(*c)
        rr = max(1.0, r * self.s)
        self.d.ellipse([x - rr, y - rr, x + rr, y + rr], outline=self._rgb(color), width=max(1, int(self.s / 6)))

    def relleno(self, pts, color=7, capa=None):
        self.d.polygon([self.P(*p) for p in pts], fill=self._rgb(color))

    def fondo(self, x1, y1, x2, y2, rgb):
        a, b = self.P(x1, y2), self.P(x2, y1)
        self.d.rectangle([a, b], fill=rgb)

    def texto(self, s, x, y, h=2.0, al="BL", capa=None, rot=0, negrita=False, color=None):
        f = _fuente_pil(h * self.s * 1.40, negrita)
        s = str(s)
        bbox = f.getbbox(s)
        tw, th = bbox[2] - bbox[0], h * self.s
        col = self._rgb(color)
        if rot == 90:
            tmp = Image.new("RGBA", (int(tw) + 4, int(th * 1.6) + 4), (255, 255, 255, 0))
            ImageDraw.Draw(tmp).text((2, 2 - bbox[1] + th * 0.15), s, font=f, fill=col)
            tmp = tmp.rotate(90, expand=True)
            px, py = self.P(x, y)
            ox = {"L": 0, "C": tmp.width / 2, "R": tmp.width}[al[1]] if False else tmp.width / 2
            oy = {"L": tmp.height, "C": tmp.height / 2, "R": 0}[al[1]]
            self.img.paste(tmp, (int(px - ox), int(py - oy)), tmp)
            return
        px, py = self.P(x, y)
        dx = {"L": 0, "C": -tw / 2, "R": -tw}[al[1]]
        dy = {"B": -th, "M": -th / 2, "T": 0}[al[0]]
        self.d.text((px + dx - bbox[0], py + dy - bbox[1]), s, font=f, fill=col)

    def mtexto(self, s, x, y, h=1.8, ancho=60, capa=None):
        f = _fuente_pil(h * self.s * 1.40)
        maxw = ancho * self.s
        lineas, actual = [], ""
        for pal in str(s).split():
            prueba = (actual + " " + pal).strip()
            if f.getlength(prueba) <= maxw:
                actual = prueba
            else:
                if actual:
                    lineas.append(actual)
                actual = pal
        if actual:
            lineas.append(actual)
        px, py = self.P(x, y)
        for i, ln in enumerate(lineas):
            self.d.text((px, py + i * h * self.s * 1.65), ln, font=f, fill=(0, 0, 0))

    def primitivas(self, prims, capa=None, color=None):
        for p in prims:
            if p[0] == "l":
                self.linea((p[1], p[2]), (p[3], p[4]), grosor=13)
            else:
                x, y = self.P(p[1], p[2])
                rr = max(1.0, p[3] * self.s)
                if p[3] < 0.35:
                    self.d.ellipse([x - rr, y - rr, x + rr, y + rr], fill=(0, 0, 0))
                else:
                    self.d.ellipse([x - rr, y - rr, x + rr, y + rr], outline=(0, 0, 0), width=max(1, int(self.s / 6)))


# -----------------------------------------------------------------------------
# 10.1 Patrones vectoriales SUCS (devuelven primitivas 'l' líneas y 'c' círculos)
# -----------------------------------------------------------------------------
def _recortar(x1, y1, x2, y2, xmin, ymin, xmax, ymax):
    """Recorte de segmento a rectángulo (Liang-Barsky)."""
    dx, dy = x2 - x1, y2 - y1
    u1, u2 = 0.0, 1.0
    for p, q in ((-dx, x1 - xmin), (dx, xmax - x1), (-dy, y1 - ymin), (dy, ymax - y1)):
        if abs(p) < 1e-12:
            if q < 0:
                return None
        else:
            t = q / p
            if p < 0:
                u1 = max(u1, t)
            else:
                u2 = min(u2, t)
    if u1 > u2:
        return None
    return (x1 + u1 * dx, y1 + u1 * dy, x1 + u2 * dx, y1 + u2 * dy)


def generar_patron(sucs, xmin, ymin, xmax, ymax):
    base = sucs_base(sucs)
    pat = SUCS_DB[base]["patron"]
    P = []
    rng = random.Random(sum(map(ord, base)) * 7 + int(ymin * 13))
    gravas_hechas = []

    def lin(x1, y1, x2, y2):
        s = _recortar(x1, y1, x2, y2, xmin, ymin, xmax, ymax)
        if s and math.hypot(s[2] - s[0], s[3] - s[1]) > 0.05:
            P.append(("l",) + s)

    def circ(cx, cy, r):
        if cx - r >= xmin + 0.2 and cx + r <= xmax - 0.2 and cy - r >= ymin + 0.2 and cy + r <= ymax - 0.2:
            P.append(("c", cx, cy, r))
            return True
        return False

    def puntos(paso, r=0.22, jit=0.30, evitar=True):
        y, fila = ymin + paso / 2, 0
        while y < ymax:
            x = xmin + paso / 2 + (paso / 2 if fila % 2 else 0)
            while x < xmax:
                px = x + rng.uniform(-jit, jit) * paso
                py = y + rng.uniform(-jit, jit) * paso
                if not (evitar and any(math.hypot(px - gx, py - gy) < gr + 0.5 for gx, gy, gr in gravas_hechas)):
                    circ(px, py, r)
                x += paso
            y += paso
            fila += 1

    def horizontales(esp, seg=None, hueco=1.5, desfase=0.0):
        y, fila = ymin + esp / 2 + desfase, 0
        while y < ymax:
            if seg is None:
                lin(xmin, y, xmax, y)
            else:
                x = xmin + (seg / 2 if fila % 2 else 0)
                while x < xmax:
                    lin(x, y, min(xmax, x + seg), y)
                    x += seg + hueco
            y += esp
            fila += 1

    def diagonales(esp, signo=1):
        w, h = xmax - xmin, ymax - ymin
        paso = esp * math.sqrt(2)
        k = -w - h
        while k < w + h:
            if signo > 0:
                lin(xmin + k, ymin, xmin + k + h, ymax)
            else:
                lin(xmin + k, ymax, xmin + k + h, ymin)
            k += paso

    def gravas(sx, sy, radios):
        y, fila, i = ymin + sy / 2, 0, 0
        while y < ymax:
            x = xmin + sx / 2 + (sx / 2 if fila % 2 else 0)
            while x < xmax:
                r = radios[i % len(radios)]
                cx = x + rng.uniform(-0.25, 0.25) * sx
                cy = y + rng.uniform(-0.2, 0.2) * sy
                if circ(cx, cy, r):
                    gravas_hechas.append((cx, cy, r))
                x += sx
                i += 1
            y += sy
            fila += 1

    if pat == "grava_bg":
        gravas(5.0, 4.2, [0.7, 1.4, 1.0, 0.5, 1.2])
        puntos(2.6, jit=0.25)
    elif pat == "grava_mg":
        gravas(4.6, 4.0, [1.1])
    elif pat == "grava_lim":
        gravas(5.6, 4.8, [1.1, 0.8])
        puntos(2.0)
    elif pat == "grava_arc":
        gravas(5.6, 4.8, [1.1, 0.8])
        diagonales(5.0, 1)
    elif pat == "arena_bg":
        puntos(1.6)
    elif pat == "arena_mg":
        puntos(2.6, jit=0.15)
    elif pat == "arena_lim":
        puntos(2.2)
        horizontales(4.4, seg=2.0, hueco=3.6, desfase=1.1)
    elif pat == "arena_arc":
        puntos(2.4)
        diagonales(5.0, 1)
    elif pat == "limo":
        horizontales(3.0, seg=4.0, hueco=1.4)
        y = ymin + 3.0
        while y < ymax - 0.5:
            x = xmin + 1.5
            while x < xmax:
                circ(x, y, 0.22)
                x += 3.0
            y += 3.0
    elif pat == "limo_ap":
        horizontales(2.0, seg=3.0, hueco=1.0)
        puntos(4.0, jit=0.1, evitar=False)
    elif pat == "arcilla":
        horizontales(2.0)
    elif pat == "arcilla_ap":
        horizontales(1.5)
        diagonales(6.0, -1)
    elif pat == "organico":
        horizontales(3.0)
        y = ymin + 1.5
        while y < ymax - 1:
            x = xmin + 2.0
            while x < xmax - 1:
                circ(x, y, 0.55)
                x += 5.0
            y += 6.0
    elif pat == "turba":
        y = ymin + 2.5
        fila = 0
        while y < ymax - 1:
            x = xmin + 2.5 + (2.5 if fila % 2 else 0)
            while x < xmax - 2:
                lin(x - 1.5, y, x + 1.5, y)
                lin(x, y, x, y + 1.6)
                lin(x, y, x - 1.0, y + 1.3)
                lin(x, y, x + 1.0, y + 1.3)
                x += 5.0
            y += 4.0
            fila += 1
    elif pat == "roca":
        y, fila = ymin, 0
        while y < ymax:
            lin(xmin, y, xmax, y)
            x = xmin + (4.0 if fila % 2 else 0.0)
            while x < xmax:
                lin(x, y, x, min(ymax, y + 4.0))
                x += 8.0
            y += 4.0
            fila += 1
    elif pat == "relleno":
        diagonales(3.0, 1)
        diagonales(3.0, -1)
    return P


# -----------------------------------------------------------------------------
# 10.2 Registro de exploración (formato oficial por calicata)
# -----------------------------------------------------------------------------
COLS_REG = [("PROF.|(m)", 14), ("MUESTRA", 14), ("SÍMBOLO|GRÁFICO", 24), ("SUCS", 16),
            ("AASHTO", 18), ("DESCRIPCIÓN DEL MATERIAL", 72), ("N.F.", 14)]
ANCHO_REG = sum(w for _, w in COLS_REG)
ESC_V = 50.0  # mm de papel por metro de profundidad (Escala vertical 1:20)
H_TIT, H_INF, H_CAB, H_PIE = 8.0, 26.0, 12.0, 8.0


def alto_registro(prof):
    return H_TIT + H_INF + H_CAB + prof * ESC_V + H_PIE


def _recortar_txt(s, n):
    s = str(s or "")
    return s if len(s) <= n else s[:n - 1] + "…"


def dibujar_registro(lz, x0, ytop, cid, ec, info, proyecto, nf, Df):
    prof = max(e["fin"] for e in ec)
    H = alto_registro(prof)
    W = ANCHO_REG
    lz.rect(x0, ytop - H, x0 + W, ytop, grosor=50)
    # --- Título
    y = ytop - H_TIT
    lz.linea((x0, y), (x0 + W, y), grosor=35)
    lz.texto(f"REGISTRO DE EXPLORACIÓN - PERFIL ESTRATIGRÁFICO  {cid}", x0 + W / 2, y + H_TIT / 2, 3.0,
             "MC", negrita=True)
    # --- Datos generales
    y2 = y - H_INF
    lz.linea((x0, y2), (x0 + W, y2), grosor=35)
    xm = x0 + 108
    lz.linea((xm, y), (xm, y2))
    ubic = f"{proyecto.get('distrito', '')} - {proyecto.get('provincia', '')} - {proyecto.get('region', '')}"
    lz.texto("PROYECTO:", x0 + 2, y - 4.0, 1.9, "BL", negrita=True)
    lz.mtexto(_recortar_txt(proyecto.get("nombre", ""), 120), x0 + 21, y - 1.6, 1.8, 85)
    lz.texto("UBICACIÓN:", x0 + 2, y - 12.0, 1.9, "BL", negrita=True)
    lz.texto(_recortar_txt(ubic, 60), x0 + 21, y - 12.0, 1.8, "BL")
    lz.texto("SOLICITANTE:", x0 + 2, y - 17.5, 1.9, "BL", negrita=True)
    lz.texto(_recortar_txt(proyecto.get("solicitante", ""), 58), x0 + 24, y - 17.5, 1.8, "BL")
    lz.texto("MÉTODO: Excavación a cielo abierto (NTP 339.162)", x0 + 2, y - 23.0, 1.8, "BL")
    nf_txt = f"{nf:.2f} m" if (nf and nf <= prof) else "No encontrado"
    est = info.get("este")
    nor = info.get("norte")
    cota = info.get("cota")
    datos = [("CALICATA:", cid), ("PROF. TOTAL:", f"{prof:.2f} m"), ("NIVEL FREÁTICO:", nf_txt),
             ("ESTE / NORTE:", f"{est:.0f} / {nor:.0f}" if (est and nor) else "-"),
             ("COTA TERRENO:", f"{cota:.2f} m.s.n.m." if cota else "-"),
             ("FECHA:", info.get("fecha") or datetime.date.today().strftime("%d/%m/%Y"))]
    for i, (a, b) in enumerate(datos):
        yy = y - 4.0 - i * 3.85
        lz.texto(a, xm + 2, yy, 1.8, "BL", negrita=True)
        lz.texto(b, xm + 30, yy, 1.8, "BL")
    # --- Cabecera de columnas
    y3 = y2 - H_CAB
    lz.linea((x0, y3), (x0 + W, y3), grosor=35)
    xs = [x0]
    for _, w in COLS_REG:
        xs.append(xs[-1] + w)
    ybot = y3 - prof * ESC_V
    for i, (nom, w) in enumerate(COLS_REG):
        cx = xs[i] + w / 2
        partes = nom.split("|")
        if len(partes) == 1:
            lz.texto(partes[0], cx, y2 - H_CAB / 2, 2.0, "MC", negrita=True)
        else:
            lz.texto(partes[0], cx, y2 - H_CAB / 2 + 2.0, 2.0, "MC", negrita=True)
            lz.texto(partes[1], cx, y2 - H_CAB / 2 - 2.0, 2.0, "MC", negrita=True)
        if i > 0:
            lz.linea((xs[i], y2), (xs[i], ybot))
    # --- Escala de profundidad cada 0.20 m
    n = int(round(prof / 0.20))
    for k in range(n + 1):
        z = k * 0.20
        yy = y3 - z * ESC_V
        mayor = abs(z - round(z)) < 1e-6
        lz.linea((xs[1] - (5.0 if mayor else 2.5), yy), (xs[1], yy), capa=CAPA_COTAS)
        lz.texto(f"{z:.2f}", xs[0] + 1.2, yy, 1.6 if not mayor else 1.8, "ML", capa=CAPA_COTAS,
                 negrita=mayor)
    # --- Estratos
    for e in ec:
        ya, yb = y3 - e["ini"] * ESC_V, y3 - e["fin"] * ESC_V
        ymid = (ya + yb) / 2
        lz.linea((xs[1], yb), (xs[6], yb), grosor=25)
        lz.fondo(xs[2], yb, xs[3], ya, SUCS_DB[sucs_base(e["sucs"])]["rgb"])
        lz.primitivas(generar_patron(e["sucs"], xs[2], yb, xs[3], ya), capa=CAPA_SUCS,
                      color=SUCS_DB[sucs_base(e["sucs"])]["aci"])
        lz.texto(e.get("muestra") or "-", (xs[1] + xs[2]) / 2, ymid, 2.0, "MC")
        lz.texto(e["sucs"], (xs[3] + xs[4]) / 2, ymid, 2.4, "MC", negrita=True)
        lz.texto(e.get("aashto") or "-", (xs[4] + xs[5]) / 2, ymid, 1.8, "MC")
        lz.texto(f"{e['est']}: {e['ini']:.2f} - {e['fin']:.2f} m", xs[5] + 1.5, ya - 1.2, 1.7, "TL",
                 negrita=True)
        lz.mtexto(e.get("desc", ""), xs[5] + 1.5, ya - 4.2, 1.6, 68.5)
    # --- Nivel freático
    cxn = (xs[6] + xs[7]) / 2
    if nf and nf <= prof:
        yn = y3 - nf * ESC_V
        lz.relleno([(cxn - 2.6, yn + 4.2), (cxn + 2.6, yn + 4.2), (cxn, yn)], color=5)
        lz.linea((cxn - 4.0, yn), (cxn + 4.0, yn), capa=CAPA_COTAS, color=5, grosor=35)
        lz.linea((cxn - 3.0, yn - 1.0), (cxn + 3.0, yn - 1.0), capa=CAPA_COTAS, color=5)
        lz.linea((cxn - 1.8, yn - 2.0), (cxn + 1.8, yn - 2.0), capa=CAPA_COTAS, color=5)
        lz.texto("N.F.", cxn, yn + 6.5, 1.8, "BC", capa=CAPA_COTAS, negrita=True, color=5)
        lz.texto(f"-{nf:.2f}", cxn, yn - 3.5, 1.6, "TC", capa=CAPA_COTAS, color=5)
        lz.linea((xs[2], yn), (xs[3], yn), capa=CAPA_COTAS, color=5, tipo="DASHED")
    else:
        lz.texto("N.F. NO ENCONTRADO", cxn, (y3 + ybot) / 2, 2.2, "MC", capa=CAPA_COTAS, rot=90, color=5)
    # --- Profundidad de cimentación recomendada
    if Df and Df <= prof:
        yd = y3 - Df * ESC_V
        lz.linea((xs[1], yd), (xs[5], yd), capa=CAPA_COTAS, color=1, tipo="DASHDOT", grosor=35, esc_tipo=1.5)
        lz.texto(f"Df = {Df:.2f} m", xs[1] + 0.8, yd + 0.8, 1.6, "BL", capa=CAPA_COTAS, color=1, negrita=True)
    # --- Pie
    lz.linea((x0, ybot), (x0 + W, ybot), grosor=35)
    lz.texto("Escala vertical 1:20  |  Muestreo y descripción: NTP 339.150 / ASTM D2488  |  "
             "Clasificación: NTP 339.134 (SUCS) - NTP 339.135 (AASHTO)", x0 + 2, ybot - H_PIE / 2, 1.6, "ML")
    return H


# -----------------------------------------------------------------------------
# 10.3 Tabla genérica, leyenda, perfil de correlación y membrete
# -----------------------------------------------------------------------------
def dibujar_tabla(lz, x, ytop, titulo, cab, anchos, filas, h_fila=5.5, h_cab=8.0, th=1.7):
    W = sum(anchos)
    y = ytop
    lz.rect(x, y - 7.0, x + W, y, grosor=35)
    lz.texto(titulo, x + W / 2, y - 3.5, 2.4, "MC", negrita=True)
    y -= 7.0
    lz.rect(x, y - h_cab, x + W, y, grosor=25)
    cx = x
    for c, w in zip(cab, anchos):
        partes = c.split("|")
        for j, p in enumerate(partes):
            off = (len(partes) - 1) * 1.3 - j * 2.6
            lz.texto(p, cx + w / 2, y - h_cab / 2 + off, th, "MC", negrita=True)
        cx += w
    y -= h_cab
    for f in filas:
        cx = x
        for v, w in zip(f, anchos):
            lz.texto(_recortar_txt(v, int(w / (th * 0.62))), cx + w / 2, y - h_fila / 2, th, "MC")
            cx += w
        lz.linea((x, y - h_fila), (x + W, y - h_fila))
        y -= h_fila
    cx = x
    for w in anchos[:-1]:
        cx += w
        lz.linea((cx, ytop - 7.0), (cx, y))
    lz.linea((x, ytop - 7.0), (x, y))
    lz.linea((x + W, ytop - 7.0), (x + W, y))
    return ytop - y


def dibujar_leyenda(lz, x, ytop, sucs_usados, ancho=120):
    n = len(sucs_usados)
    cols = 2
    filas = math.ceil(n / cols)
    H = 9.0 + filas * 11.0
    lz.rect(x, ytop - H, x + ancho, ytop, grosor=35)
    lz.texto("LEYENDA - SIMBOLOGÍA SUCS", x + ancho / 2, ytop - 4.0, 2.4, "MC", negrita=True)
    for i, s in enumerate(sucs_usados):
        c, f = i % cols, i // cols
        xx = x + 4 + c * (ancho / cols)
        yy = ytop - 10.0 - f * 11.0
        lz.rect(xx, yy - 8, xx + 14, yy, capa=CAPA_SUCS)
        lz.primitivas(generar_patron(s, xx, yy - 8, xx + 14, yy), capa=CAPA_SUCS,
                      color=SUCS_DB[sucs_base(s)]["aci"])
        lz.texto(s, xx + 16, yy - 2.6, 2.0, "ML", negrita=True)
        lz.texto(_recortar_txt(SUCS_DB[sucs_base(s)]["nombre"], 34), xx + 16, yy - 6.0, 1.6, "ML")
    return H


def dibujar_correlacion(lz, x, ytop, ancho, estratos, cal_info, nf, Df):
    cals = calicatas_de(estratos)
    prof_max = max(e["fin"] for e in estratos)
    cotas = [cal_info.get(c, {}).get("cota") for c in cals]
    usar_cota = all(ct for ct in cotas) and len(cals) > 1
    cmax = max(cotas) if usar_cota else 0.0
    cmin = min(cotas) if usar_cota else 0.0
    alto_dib = (prof_max + (cmax - cmin)) * ESC_V
    H = 12.0 + 10.0 + alto_dib + 16.0
    lz.rect(x, ytop - H, x + ancho, ytop, grosor=35)
    lz.texto("PERFIL LONGITUDINAL DE CORRELACIÓN ESTRATIGRÁFICA", x + ancho / 2, ytop - 5.0, 2.8, "MC",
             negrita=True)
    lz.texto("Esc. vertical 1:20 - Esc. horizontal: indicada por distancias", x + ancho / 2, ytop - 9.5, 1.7,
             "MC")
    y0 = ytop - 22.0  # nivel de la cota máxima
    # distancias horizontales
    xs_cal = []
    dist = [0.0]
    for i in range(1, len(cals)):
        a, b = cal_info.get(cals[i - 1], {}), cal_info.get(cals[i], {})
        if a.get("este") and b.get("este") and a.get("norte") and b.get("norte"):
            dist.append(dist[-1] + math.hypot(b["este"] - a["este"], b["norte"] - a["norte"]))
        else:
            dist.append(dist[-1] + 20.0)
    x_ini, x_fin = x + 34, x + ancho - 30
    span = dist[-1] if dist[-1] > 0 else 1.0
    for dd in dist:
        xs_cal.append(x_ini + (x_fin - x_ini) * (dd / span if len(cals) > 1 else 0.5))
    wcol = 12.0
    # eje de cotas / profundidades
    lz.linea((x + 22, y0), (x + 22, y0 - alto_dib), capa=CAPA_COTAS)
    paso = 0.5
    k = 0
    while k * paso <= prof_max + (cmax - cmin) + 1e-6:
        yy = y0 - k * paso * ESC_V
        lz.linea((x + 20, yy), (x + 22, yy), capa=CAPA_COTAS)
        etq = f"{cmax - k * paso:.2f}" if usar_cota else f"{k * paso:.2f}"
        lz.texto(etq, x + 19, yy, 1.5, "MR", capa=CAPA_COTAS)
        k += 1
    lz.texto("COTA (m.s.n.m.)" if usar_cota else "PROF. (m)", x + 3, y0 + 4.5, 1.6, "BL", capa=CAPA_COTAS,
             negrita=True)
    tops = []
    for i, cid in enumerate(cals):
        ec = estratos_de(estratos, cid)
        top = y0 - ((cmax - cotas[i]) * ESC_V if usar_cota else 0.0)
        tops.append(top)
        xc = xs_cal[i]
        for e in ec:
            ya, yb = top - e["ini"] * ESC_V, top - e["fin"] * ESC_V
            lz.rect(xc - wcol / 2, yb, xc + wcol / 2, ya, capa=CAPA_CUADROS)
            lz.primitivas(generar_patron(e["sucs"], xc - wcol / 2, yb, xc + wcol / 2, ya), capa=CAPA_SUCS,
                          color=SUCS_DB[sucs_base(e["sucs"])]["aci"])
            lz.texto(e["sucs"], xc + wcol / 2 + 1.2, (ya + yb) / 2, 1.8, "ML", negrita=True)
        pf = ec[-1]["fin"]
        lz.texto(cid, xc, top + 3.0, 2.6, "BC", negrita=True)
        lz.texto(f"Prof. {pf:.2f} m", xc, top - pf * ESC_V - 2.0, 1.6, "TC")
        if i > 0:
            lz.texto(f"d = {dist[i] - dist[i - 1]:.1f} m", (xs_cal[i - 1] + xc) / 2, top + 6.5, 1.6, "BC",
                     capa=CAPA_COTAS)
    # Líneas de correlación (punteadas) entre calicatas consecutivas
    for i in range(1, len(cals)):
        ea, eb = estratos_de(estratos, cals[i - 1]), estratos_de(estratos, cals[i])
        xa, xb = xs_cal[i - 1] + wcol / 2, xs_cal[i] - wcol / 2
        lz.linea((xa, tops[i - 1]), (xb, tops[i]), capa=CAPA_CUADROS, grosor=35)
        usados = set()
        for e1 in ea[1:]:
            par = None
            for e2 in eb[1:]:
                if sucs_base(e2["sucs"]) == sucs_base(e1["sucs"]) and id(e2) not in usados:
                    par = e2
                    break
            if par is None:
                idx = ea.index(e1)
                par = eb[idx] if idx < len(eb) and id(eb[idx]) not in usados else None
            if par is None:
                continue
            usados.add(id(par))
            lz.linea((xa, tops[i - 1] - e1["ini"] * ESC_V), (xb, tops[i] - par["ini"] * ESC_V),
                     capa=CAPA_COTAS, tipo="DASHED", color=8)
        # nivel freático y cota de cimentación
        if nf:
            lz.linea((xa, tops[i - 1] - nf * ESC_V), (xb, tops[i] - nf * ESC_V), capa=CAPA_COTAS, color=5,
                     tipo="DASHED")
        if Df:
            lz.linea((xs_cal[i - 1] - wcol / 2, tops[i - 1] - Df * ESC_V), (xs_cal[i] + wcol / 2, tops[i] - Df * ESC_V),
                     capa=CAPA_COTAS, color=1, tipo="DASHDOT", grosor=35, esc_tipo=1.5)
    if Df:
        lz.texto(f"Df = {Df:.2f} m (profundidad de cimentación recomendada)", xs_cal[0] - wcol / 2,
                 tops[0] - Df * ESC_V - 2.0, 1.7, "TL", capa=CAPA_COTAS, color=1, negrita=True)
    if nf:
        xn = xs_cal[-1] + wcol / 2 + 9
        yn = tops[-1] - nf * ESC_V
        lz.relleno([(xn - 2.0, yn + 3.2), (xn + 2.0, yn + 3.2), (xn, yn)], color=5)
        lz.texto(f"N.F. -{nf:.2f} m", xn, yn + 4.0, 1.6, "BC", capa=CAPA_COTAS, color=5)
    lz.texto("Líneas punteadas: continuidad estratigráfica inferida entre calicatas.", x + 3, ytop - H + 4.0,
             1.6, "BL")
    return H


def dibujar_membrete(lz, x, ytop, W, proyecto):
    H = 56.0
    lz.rect(x, ytop - H, x + W, ytop, grosor=50)
    lz.rect(x, ytop - 12, x + W, ytop, grosor=35)
    lz.texto("ESTUDIO DE MECÁNICA DE SUELOS CON FINES DE CIMENTACIÓN", x + W / 2, ytop - 4.0, 2.6, "MC",
             negrita=True)
    lz.texto("NTE E.050 - SUELOS Y CIMENTACIONES / NTE E.030 - DISEÑO SISMORRESISTENTE", x + W / 2, ytop - 8.8,
             1.8, "MC")
    filas = [("PROYECTO:", proyecto.get("nombre", "")),
             ("SOLICITANTE:", proyecto.get("solicitante", "")),
             ("UBICACIÓN:", f"Dist. {proyecto.get('distrito', '')} - Prov. {proyecto.get('provincia', '')} - "
                            f"Reg. {proyecto.get('region', '')}"),
             ("CONTENIDO:", "Registros de exploración, perfil de correlación y cuadro resumen de ensayos"),
             ("RESPONSABLE:", f"{proyecto.get('ingeniero', '')}   CIP N° {proyecto.get('cip', '')}")]
    for i, (a, b) in enumerate(filas):
        yy = ytop - 16.5 - i * 5.6
        lz.texto(a, x + 2, yy, 1.9, "ML", negrita=True)
        lz.texto(_recortar_txt(b, 92), x + 25, yy, 1.8, "ML")
    yb = ytop - 44.0
    lz.linea((x, yb), (x + W, yb), grosor=35)
    celdas = [("LÁMINA:", "EMS-01"), ("ESCALA:", "V 1:20 / H indicada"), ("FECHA:", fecha_larga()),
              ("FIRMA Y SELLO:", "")]
    wc = W / 4
    for i, (a, b) in enumerate(celdas):
        cx = x + i * wc
        if i:
            lz.linea((cx, yb), (cx, ytop - H))
        lz.texto(a, cx + 2, yb - 3.5, 1.7, "ML", negrita=True)
        lz.texto(b, cx + wc / 2, yb - 8.5, 2.6 if a == "LÁMINA:" else 1.9, "MC", negrita=(a == "LÁMINA:"))
    return H


# =============================================================================
# 11. EXPORTACIÓN A AUTOCAD (.DXF) - LÁMINA EMS-01
# =============================================================================
def exportar_dxf(path, proyecto, cal_info, estratos, calc):
    doc = ezdxf.new("R2010", setup=True)
    doc.units = ezdxf.units.MM
    doc.header["$INSUNITS"] = 4
    doc.header["$LWDISPLAY"] = 1
    for nombre, color in ((CAPA_CUADROS, 7), (CAPA_SUCS, 32), (CAPA_COTAS, 1), (CAPA_TEXTOS, 7)):
        if nombre not in doc.layers:
            doc.layers.add(nombre, color=color)
    if "EMS" not in doc.styles:
        doc.styles.add("EMS", font="arial.ttf")
    if "EMS_B" not in doc.styles:
        doc.styles.add("EMS_B", font="arialbd.ttf")
    lz = LienzoDXF(doc)
    cals = calicatas_de(estratos)
    nf = calc["nf"]
    dis = calc["diseno"]
    Df = calc["cfg"]["Df"]
    M = 15.0
    gap = 12.0
    logs_w = len(cals) * ANCHO_REG + (len(cals) - 1) * gap
    prof_max = max(e["fin"] for e in estratos)
    h_log = alto_registro(prof_max)
    der_w = 300.0
    izq_w = max(240.0, logs_w - der_w - gap)
    ancho_total = max(logs_w, izq_w + gap + der_w)
    # ---- Contenido - fila 1: registros
    x0, ytop = M, 0.0
    for i, cid in enumerate(cals):
        ec = estratos_de(estratos, cid)
        nfc = cal_info.get(cid, {}).get("nf") or nf
        dibujar_registro(lz, x0 + i * (ANCHO_REG + gap), ytop, cid, ec, cal_info.get(cid, {}), proyecto, nfc, Df)
    # ---- fila 2 izquierda: correlación
    y2 = ytop - h_log - gap
    h_cor = dibujar_correlacion(lz, x0, y2, izq_w, estratos, cal_info, nf, Df)
    # ---- fila 2 derecha: cuadro resumen + capacidad portante + leyenda + membrete
    xr = x0 + izq_w + gap
    cab = ["CALICATA", "MUESTRA", "PROF.|(m)", "SUCS", "AASHTO", "w|(%)", "LL|(%)", "IP|(%)", "γ|(g/cm³)",
           "φ|(°)", "c|(kg/cm²)"]
    anchos = [24, 22, 30, 22, 30, 22, 22, 22, 28, 22, 30]
    escala = der_w / sum(anchos)
    anchos = [a * escala for a in anchos]
    filas = []
    for e in sorted(estratos, key=lambda e: (orden_natural(e["cal"]), e["ini"])):
        filas.append([e["cal"], e.get("muestra") or "-", f"{e['ini']:.2f}-{e['fin']:.2f}", e["sucs"],
                      e.get("aashto") or "-", fmt(e.get("w"), 1, "-"), fmt(e.get("ll"), 0), fmt(e.get("ip"), 0),
                      fmt(e.get("gamma"), 2, "-"), fmt(e.get("phi"), 1, "-"), fmt(e.get("c"), 2, "-")])
    h_tab = dibujar_tabla(lz, xr, y2, "CUADRO RESUMEN DE ENSAYOS DE LABORATORIO", cab, anchos, filas)
    # capacidad portante recomendada
    y3 = y2 - h_tab - 6.0
    s = calc["sismo"]
    q = calc["quim"]
    cfg = calc["cfg"]
    rec = [["Profundidad de cimentación Df", f"{Df:.2f} m"],
           ["Tipo de cimentación / ancho B", f"Zapata {dis['tipo'].lower()} B = {dis['B']:.2f} m"],
           ["Estrato de apoyo (calicata crítica)", f"{dis['sucs']} ({dis['cal']} / {dis['estrato']})"],
           ["Teoría / tipo de falla / FS", f"{cfg['metodo']} / Corte {cfg['falla'].lower()} / FS = {cfg['FS']:.1f}"],
           ["q admisible por corte", f"{dis['qadm_corte']:.2f} kg/cm²"],
           [f"q admisible por asentamiento (S ≤ {cfg['smax']:.2f} cm)", f"{dis['q_asent']:.2f} kg/cm²"],
           [f"q admisible por distorsión (α ≤ 1/500, L = {cfg['luz']:.1f} m)", f"{dis['q_dist']:.2f} kg/cm²"],
           ["CAPACIDAD ADMISIBLE DE DISEÑO", f"{dis['qadm']:.2f} kg/cm²"],
           ["Asentamiento elástico esperado", f"{dis['se']:.2f} cm"],
           ["Cimiento corrido B = 0.60 m (qadm)",
            f"{calc['corrido']['qadm']:.2f} kg/cm²" if calc.get("corrido") else "-"],
           ["Parámetros E.030", f"Z={s['Z']:.2f}  S={s['S']:.2f}  Tp={s['Tp']:.1f} s  TL={s['Tl']:.1f} s  ({s['perfil']})"],
           ["Agresividad / cemento", f"{q['grado']} - {q['cemento'].split(' (')[0]}"]]
    h_rec = dibujar_tabla(lz, xr, y3, "CAPACIDAD PORTANTE RECOMENDADA (NTE E.050)", ["PARÁMETRO", "VALOR"],
                          [der_w * 0.48, der_w * 0.52], rec, h_fila=5.2, h_cab=6.0)
    y4 = y3 - h_rec - 6.0
    usados = []
    for e in estratos:
        if e["sucs"] not in usados:
            usados.append(e["sucs"])
    h_ley = dibujar_leyenda(lz, xr, y4, usados, ancho=der_w)
    y5 = y4 - h_ley - 6.0
    h_mem = dibujar_membrete(lz, xr, y5, der_w, proyecto)
    yfin = min(y2 - h_cor, y5 - h_mem) - M
    # ---- Marco general
    lz.rect(M / 2 - 2, yfin, M + ancho_total + M / 2 + 2, ytop + M, grosor=70)
    lz.rect(M / 2 + 1, yfin + 3, M + ancho_total + M / 2 - 1, ytop + M - 3, grosor=25)
    lz.texto("LÁMINA EMS-01 - REGISTROS DE EXPLORACIÓN, PERFIL ESTRATIGRÁFICO Y CAPACIDAD PORTANTE",
             M, ytop + 5.0, 3.5, "BL", negrita=True)
    doc.set_modelspace_vport(height=(ytop + M - yfin) * 1.05, center=((M + ancho_total) / 2, (ytop + yfin) / 2))
    doc.saveas(path)
    return path


def imagen_registro_png(path, cid, ec, info, proyecto, nf, Df, px_mm=7.0):
    prof = max(e["fin"] for e in ec)
    H = alto_registro(prof)
    lz = LienzoPIL(ANCHO_REG + 4, H + 4, x0=-2, ytop=2, px_mm=px_mm)
    dibujar_registro(lz, 0, 0, cid, ec, info, proyecto, nf, Df)
    lz.img.save(path)
    return path


# =============================================================================
# 12. EXPORTACIÓN A EXCEL (.XLSX) - 4 HOJAS
# =============================================================================
AZUL = "1F4E79"
BORDE = Border(*(Side(style="thin", color="808080"),) * 4)


def _xl_titulo(ws, fila, texto, ncols, proyecto):
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
    c = ws.cell(fila, 1, texto)
    c.font = Font(bold=True, size=14, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=AZUL)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[fila].height = 24
    ws.merge_cells(start_row=fila + 1, start_column=1, end_row=fila + 1, end_column=ncols)
    c = ws.cell(fila + 1, 1, f"Proyecto: {proyecto.get('nombre', '')}  |  Solicitante: {proyecto.get('solicitante', '')}"
                             f"  |  {proyecto.get('distrito', '')} - {proyecto.get('provincia', '')} - "
                             f"{proyecto.get('region', '')}  |  Fecha: {fecha_larga()}")
    c.font = Font(italic=True, size=9)
    c.alignment = Alignment(horizontal="center")


def _xl_cabecera(ws, fila, cab, col0=1, color=AZUL):
    for j, t in enumerate(cab):
        c = ws.cell(fila, col0 + j, t)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.row_dimensions[fila].height = 32


def _xl_fila(ws, fila, vals, col0=1, fmt_num="0.00", fill=None, negrita=False):
    for j, v in enumerate(vals):
        c = ws.cell(fila, col0 + j, v)
        c.border = BORDE
        c.font = Font(size=9, bold=negrita)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=isinstance(v, str) and len(v) > 30)
        if isinstance(v, float):
            c.number_format = fmt_num
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)


def _xl_anchos(ws, anchos):
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def exportar_excel(path, proyecto, cal_info, estratos, calc):
    wb = openpyxl.Workbook()
    est_ord = sorted(estratos, key=lambda e: (orden_natural(e["cal"]), e["ini"]))
    cfg = calc["cfg"]
    # ---------------------------------------------------------------- Hoja 1
    ws = wb.active
    ws.title = "1_Resumen_Laboratorio_SUCS"
    _xl_titulo(ws, 1, "RESUMEN DE ENSAYOS DE LABORATORIO Y CLASIFICACIÓN SUCS / AASHTO", 20, proyecto)
    cab = ["Calicata", "Estrato", "Muestra", "Prof. ini (m)", "Prof. fin (m)", "Espesor (m)", "SUCS", "AASHTO",
           "% Grava", "% Arena", "% Finos", "w (%)", "LL (%)", "LP (%)", "IP (%)", "γ (g/cm³)", "φ (°)",
           "c (kg/cm²)", "Es típico (kg/cm²)", "Descripción litológica"]
    _xl_cabecera(ws, 4, cab)
    r = 5
    for e in est_ord:
        lp = (e["ll"] - e["ip"]) if (e.get("ll") is not None and e.get("ip") is not None) else "NP"
        _xl_fila(ws, r, [e["cal"], e["est"], e.get("muestra") or "-", e["ini"], e["fin"], None, e["sucs"],
                         e.get("aashto") or "-", e.get("grava"), e.get("arena"), e.get("finos"), e.get("w"),
                         e["ll"] if e.get("ll") is not None else "NP", lp,
                         e["ip"] if e.get("ip") is not None else "NP", e.get("gamma"), e.get("phi"), e.get("c"),
                         SUCS_DB[sucs_base(e["sucs"])]["Es"], e.get("desc", "")])
        ws.cell(r, 6, f"=E{r}-D{r}").number_format = "0.00"
        ws.cell(r, 6).border = BORDE
        ws.cell(r, 20).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 42
        r += 1
    # Granulometría
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = ws.cell(r, 1, "ANÁLISIS GRANULOMÉTRICO POR TAMIZADO (NTP 339.128 / ASTM D422) - % QUE PASA")
    c.font = Font(bold=True, size=11, color=AZUL)
    r += 1
    muestras = [e for e in est_ord if sucs_base(e["sucs"]) not in ("RELL", "Roca")]
    cabg = ["Tamiz", "Abertura (mm)"] + [f"{e['cal']} {e.get('muestra') or e['est']}" for e in muestras]
    _xl_cabecera(ws, r, cabg)
    r_g0 = r
    curvas = [coeficientes_curva(e) for e in muestras]
    for i, (tam, d) in enumerate(TAMICES):
        r += 1
        _xl_fila(ws, r, [tam, d] + [cv[0][i] for cv in curvas], fmt_num="0.0")
        ws.cell(r, 2).number_format = "0.000"
    for nombre, idx, f_ in (("D10 (mm)", 1, "0.000"), ("D30 (mm)", 2, "0.000"), ("D60 (mm)", 3, "0.000"),
                            ("Cu = D60/D10", 4, "0.00"), ("Cc = D30²/(D10·D60)", 5, "0.00")):
        r += 1
        _xl_fila(ws, r, [nombre, ""] + [(float(cv[idx]) if cv[idx] is not None else "-") for cv in curvas],
                 fmt_num=f_, fill="DDEBF7", negrita=True)
    r += 1
    _xl_fila(ws, r, ["SUCS (auto)", ""] + [autoclasificar_estrato(e)[0] for e in muestras], fill="FCE4D6",
             negrita=True)
    r += 1
    _xl_fila(ws, r, ["AASHTO (auto)", ""] + [autoclasificar_estrato(e)[1] for e in muestras], fill="FCE4D6",
             negrita=True)
    if muestras:
        ch = ScatterChart()
        ch.title = "Curvas granulométricas"
        ch.style = 13
        ch.x_axis.title = "Abertura (mm) - escala log"
        ch.y_axis.title = "% que pasa"
        ch.x_axis.scaling.logBase = 10
        ch.x_axis.scaling.orientation = "maxMin"
        ch.x_axis.scaling.min = 0.05
        ch.x_axis.scaling.max = 100
        ch.y_axis.scaling.min = 0
        ch.y_axis.scaling.max = 100
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        xref = Reference(ws, min_col=2, min_row=r_g0 + 1, max_row=r_g0 + len(TAMICES))
        for j in range(len(muestras)):
            yref = Reference(ws, min_col=3 + j, min_row=r_g0, max_row=r_g0 + len(TAMICES))
            se = Series(yref, xref, title_from_data=True)
            se.marker.symbol = "circle"
            se.smooth = False
            ch.series.append(se)
        ch.height, ch.width = 9.5, 18
        ws.add_chart(ch, f"A{r + 3}")
    _xl_anchos(ws, [11, 11, 10, 10, 10, 10, 9, 11, 9, 9, 9, 8, 8, 8, 8, 10, 8, 10, 12, 70])
    ws.freeze_panes = "A5"
    # ---------------------------------------------------------------- Hoja 2
    ws = wb.create_sheet("2_Capacidad_Portante_Terz_Meyer")
    _xl_titulo(ws, 1, "CAPACIDAD PORTANTE - TERZAGHI / MEYERHOF / VESIC (NTE E.050)", 26, proyecto)
    datos = [("Tipo de falla", f"Corte {cfg['falla']}"), ("Factor de seguridad FS", cfg["FS"]),
             ("Teoría de diseño", cfg["metodo"]), ("Calicata analizada", cfg["calicata"]),
             ("Nivel freático (m)", calc["nf"] if calc["nf"] else "No encontrado"),
             ("Relación L/B (rectangular)", cfg["relLB"])]
    for i, (a, b) in enumerate(datos):
        ws.cell(4 + i, 1, a).font = Font(bold=True, size=9)
        ws.cell(4 + i, 3, b).font = Font(size=9)
    ws.cell(4, 6, "qu = c·Nc·Sc·dc + γ·Df·Nq·Sq·dq + 0.5·γ·B·Nγ·Sγ·dγ     qadm = qu / FS").font = \
        Font(bold=True, italic=True, size=10, color=AZUL)
    ws.cell(5, 6, "Falla local: tan φ* = 2/3·tan φ ; c* = 2/3·c   |   c en kg/cm² ×10 = t/m²; γ en g/cm³ = t/m³").font = \
        Font(italic=True, size=9)
    ws.cell(6, 6, "FS en la celda C5: las columnas 'qadm' son fórmulas vinculadas (modifique FS y se recalcula).").font = \
        Font(italic=True, size=9, color="C00000")
    r = 11
    for metodo in METODOS:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=24)
        c = ws.cell(r, 1, f"MÉTODO DE {metodo.upper()}")
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="C55A11" if metodo == "Meyerhof" else
                             ("548235" if metodo == "Vesic" else "2F5597"))
        r += 1
        cab = ["Df (m)", "B (m)", "L (m)", "Tipo", "Calicata crítica", "Estrato", "SUCS", "φ (°)", "c (kg/cm²)",
               "γ base (t/m³)", "γ sobre (t/m³)", "φ usado (°)", "Nc", "Nq", "Nγ", "Sc", "Sq", "Sγ", "dc", "dq",
               "dγ", "q = γDf (t/m²)", "qu (kg/cm²)", "qadm (kg/cm²)"]
        _xl_cabecera(ws, r, cab)
        for f in calc["filas"]:
            r += 1
            x = f["res"][metodo]
            _xl_fila(ws, r, [f["Df"], f["B"], (f["L"] if f["L"] else "∞"), f["tipo"], f["cal"], f["estrato"],
                             f["sucs"], float(f["phi"]), float(f["c"]), float(f["g_base"]), float(f["g_sobre"]),
                             x["phi_u"], x["Nc"], x["Nq"], x["Ng"], x["sc"], x["sq"], x["sg"], x["dc"], x["dq"],
                             x["dg"], x["q"], x["qu"], None])
            ws.cell(r, 24, f"=W{r}/$C$5").number_format = "0.00"
            ws.cell(r, 24).border = BORDE
            ws.cell(r, 24).font = Font(bold=True, size=9, color="C00000")
        r += 3
    _xl_anchos(ws, [8, 8, 8, 12, 10, 8, 8, 7, 9, 9, 9, 9, 8, 8, 8, 7, 7, 7, 7, 7, 7, 10, 10, 11])
    # ---------------------------------------------------------------- Hoja 3
    ws = wb.create_sheet("3_Calculo_Asentamientos")
    _xl_titulo(ws, 1, "ASENTAMIENTO ELÁSTICO INMEDIATO (Schleicher / Steinbrenner) Y DISTORSIÓN ANGULAR", 18, proyecto)
    ws.cell(4, 1, "Se = q·B·(1-μ²)/Es · If      |   S máx = 2.54 cm (1\")   |   Distorsión angular α = δ/L ≤ 1/500"
                  f"   |   δ ≈ {cfg['kdif']:.2f}·Se (asentamiento diferencial)").font = Font(bold=True, italic=True, color=AZUL)
    ws.cell(5, 1, f"Teoría de corte: {cfg['metodo']}  |  Es y μ: {'según SUCS del estrato de apoyo' if cfg['es_auto'] else 'manual'}"
                  f"  |  Luz entre columnas L = {cfg['luz']:.2f} m  |  If: cimentación rígida (Bowles)").font = Font(size=9)
    cab = ["Df (m)", "B (m)", "Tipo", "Calicata", "SUCS", "qadm corte (kg/cm²)", "Es (kg/cm²)", "μ", "If",
           "Se con qadm corte (cm)", "q adm por asentamiento (kg/cm²)", "q adm por distorsión (kg/cm²)",
           "qadm DISEÑO (kg/cm²)", "Gobierna",
           "Se diseño (cm)", "δ diferencial (cm)", "α = δ/L", "Verif. Se ≤ 2.54", "Verif. α ≤ 1/500"]
    _xl_cabecera(ws, 7, cab)
    r = 7
    for f in calc["filas"]:
        r += 1
        _xl_fila(ws, r, [f["Df"], f["B"], f["tipo"], f["cal"], f["sucs"], f["qadm_corte"], float(f["Es"]), f["mu"],
                         f["If"], f["se_corte"], f["q_asent"], f["q_dist"], f["qadm"], f["gobierna"], f["se"], f["delta"],
                         f"1/{(1 / f['alfa']):.0f}" if f["alfa"] > 0 else "-",
                         "CUMPLE" if f["ok_se"] else "NO CUMPLE", "CUMPLE" if f["ok_alfa"] else "NO CUMPLE"],
                 fill=None)
        ws.cell(r, 13).font = Font(bold=True, size=9, color="C00000")
        for col in (18, 19):
            cc = ws.cell(r, col)
            cc.fill = PatternFill("solid", fgColor="C6EFCE" if cc.value == "CUMPLE" else "FFC7CE")
    _xl_anchos(ws, [8, 8, 12, 9, 8, 12, 10, 7, 7, 12, 14, 14, 13, 14, 10, 11, 10, 12, 12])
    ws.freeze_panes = "A8"
    # ---------------------------------------------------------------- Hoja 4
    ws = wb.create_sheet("4_Agresividad_Quimica_y_Empujes")
    _xl_titulo(ws, 1, "AGRESIVIDAD QUÍMICA (E.060 / ACI 318), PARÁMETROS SÍSMICOS E.030 Y EMPUJES", 10, proyecto)
    q = calc["quim"]
    _xl_cabecera(ws, 4, ["Elemento químico", "Resultado (ppm)", "Despreciable", "Moderado", "Severo",
                         "Muy severo", "Grado de ataque", "Norma"])
    filasq = [("Sulfatos solubles SO₄", q["so4"], "0 - 1000", "1000 - 2000", "2000 - 20000", "> 20000", q["g_so4"], "NTP 339.178"),
              ("Cloruros Cl⁻", q["cl"], "≤ 1000", "1000 - 6000", "6000 - 15000", "> 15000", q["g_cl"], "NTP 339.177"),
              ("Sales solubles totales", q["sst"], "≤ 5000", "5000 - 15000", "15000 - 30000", "> 30000", q["g_sst"], "NTP 339.152"),
              ("pH", q["ph"], "≥ 6.5", "5.5 - 6.5", "4.5 - 5.5", "< 4.5", q["g_ph"], "NTP 339.176")]
    r = 4
    for f in filasq:
        r += 1
        _xl_fila(ws, r, list(f))
        ws.cell(r, 7).fill = PatternFill("solid", fgColor={"Despreciable": "C6EFCE", "Moderado": "FFEB9C",
                                                          "Severo": "F8CBAD", "Muy Severo": "FFC7CE"}[f[6]])
    r += 2
    for a, b in (("GRADO DE ATAQUE GLOBAL", q["grado"]), ("CEMENTO RECOMENDADO", q["cemento"]),
                 ("Relación a/c máxima", q["ac"]), ("f'c mínimo (kg/cm²)", q["fc"])):
        ws.cell(r, 1, a).font = Font(bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(r, 2, b).font = Font(bold=True, color="C00000")
        r += 1
    for m_ in q["medidas"]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.cell(r, 1, "• " + m_).font = Font(size=9)
        r += 1
    r += 1
    s = calc["sismo"]
    _xl_cabecera(ws, r, ["Parámetro E.030", "Valor", "Descripción"])
    for a, b, d_ in (("Zona sísmica", s["zona"], "NTE E.030-2018, Anexo I"), ("Z", s["Z"], "Factor de zona"),
                     ("Perfil de suelo", s["perfil"], PERFIL_DESC[s["perfil"]]), ("S", s["S"], "Factor de suelo"),
                     ("Tp (s)", s["Tp"], "Periodo de plataforma"), ("TL (s)", s["Tl"], "Periodo inicio desplaz. constante"),
                     ("U", s["U"], f"Categoría {s['categoria']}: {CATEGORIA_DESC.get(s['categoria'], '')}"),
                     ("T estimado (s)", s["T"], f"T = hn/CT = {s['hn']:.2f}/{s['Ct']}"),
                     ("C", s["C"], "Factor de amplificación sísmica")):
        r += 1
        _xl_fila(ws, r, [a, float(b) if isinstance(b, (int, float)) and not isinstance(b, bool) else b, d_])
    r += 2
    _xl_cabecera(ws, r, ["Calicata", "Estrato", "SUCS", "φ (°)", "c (kg/cm²)", "γ (t/m³)", "Ka (Rankine)",
                         "Kp (Rankine)", "K0 (Jaky)", "μ = tan(2φ/3)"])
    for em in calc["empujes"]:
        r += 1
        _xl_fila(ws, r, [em["cal"], em["est"], em["sucs"], float(em["phi"]), float(em["c"]), float(em["gamma"]),
                         em["Ka"], em["Kp"], em["K0"], em["mu"]], fmt_num="0.000")
    _xl_anchos(ws, [26, 14, 18, 14, 14, 14, 16, 14, 12, 14])
    wb.save(path)
    return path


# =============================================================================
# 13. EXPORTACIÓN A WORD (.DOCX) - INFORME TÉCNICO EMS (Art. 16 NTE E.050)
# =============================================================================
def _sombrear(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def _campo(parrafo, instr, texto="Actualice el campo (clic derecho > Actualizar campo)"):
    def run_fld(tipo):
        r = parrafo.add_run()
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), tipo)
        r._r.append(f)
        return r
    run_fld("begin")
    r = parrafo.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    r._r.append(it)
    run_fld("separate")
    parrafo.add_run(texto)
    run_fld("end")


def _tabla_word(doc, cab, filas, anchos=None, fs=8.5, color="1F4E79", negrita_ultima=False):
    t = doc.add_table(rows=1, cols=len(cab))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, c in enumerate(cab):
        cel = t.rows[0].cells[j]
        cel.text = ""
        p = cel.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(c))
        run.bold = True
        run.font.size = Pt(fs)
        run.font.color.rgb = RGBColor(255, 255, 255)
        _sombrear(cel, color)
    for i, f in enumerate(filas):
        cells = t.add_row().cells
        for j, v in enumerate(f):
            cells[j].text = ""
            p = cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (len(str(v)) < 40) else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(v))
            run.font.size = Pt(fs)
            if negrita_ultima and i == len(filas) - 1:
                run.bold = True
            if i % 2 == 1:
                _sombrear(cells[j], "EAF1FB")
    if anchos:
        for row in t.rows:
            for j, w in enumerate(anchos):
                row.cells[j].width = Cm(w)
    doc.add_paragraph()
    return t


def _par(doc, texto, negrita=False, cursiva=False, al=WD_ALIGN_PARAGRAPH.JUSTIFY, fs=10.5, color=None, espacio=6):
    p = doc.add_paragraph()
    p.alignment = al
    p.paragraph_format.space_after = Pt(espacio)
    r = p.add_run(texto)
    r.bold, r.italic = negrita, cursiva
    r.font.size = Pt(fs)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return p


def _eq(doc, texto, fs=11):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(texto)
    r.italic = True
    r.font.size = Pt(fs)
    r.font.name = "Cambria Math"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Cambria Math")
    return p


def _vinetas(doc, items, fs=10.5):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(it)
        r.font.size = Pt(fs)


def _preparar_imagen(path, tmpdir, idx, max_px=1600):
    try:
        im = Image.open(path)
        im = ImageOps.exif_transpose(im)
        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        im.thumbnail((max_px, max_px))
        out = os.path.join(tmpdir, f"img_{idx}.jpg")
        im.save(out, "JPEG", quality=85)
        return out, im.size
    except Exception:
        return None, None


def exportar_word(path, proyecto, cal_info, estratos, calc):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.5)
    sec.top_margin, sec.bottom_margin = Cm(2.5), Cm(2.0)
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for lvl, tam in ((1, 13), (2, 11.5), (3, 10.5)):
        h = doc.styles[f"Heading {lvl}"]
        h.font.name = "Arial"
        h.font.size = Pt(tam)
        h.font.bold = True
        h.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        h.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    # Actualizar campos (índice) al abrir
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    doc.settings.element.append(uf)
    # Encabezado y pie
    hp = sec.header.paragraphs[0]
    hp.text = f"ESTUDIO DE MECÁNICA DE SUELOS CON FINES DE CIMENTACIÓN - {proyecto.get('nombre', '')}"
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in hp.runs:
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = fp.add_run(f"{proyecto.get('ingeniero', '')} - CIP N° {proyecto.get('cip', '')}    |    Página ")
    rr.font.size = Pt(8)
    _campo(fp, "PAGE", "1")
    sec.different_first_page_header_footer = True

    cfg = calc["cfg"]
    dis = calc["diseno"]
    s = calc["sismo"]
    q = calc["quim"]
    nf = calc["nf"]
    cals = calicatas_de(estratos)
    est_ord = sorted(estratos, key=lambda e: (orden_natural(e["cal"]), e["ini"]))
    ubic = f"Distrito de {proyecto.get('distrito', '')}, Provincia de {proyecto.get('provincia', '')}, " \
           f"Región {proyecto.get('region', '')}"
    Df = cfg["Df"]

    # ------------------------------------------------------------- PORTADA
    for _ in range(3):
        doc.add_paragraph()
    _par(doc, "ESTUDIO DE MECÁNICA DE SUELOS", True, al=WD_ALIGN_PARAGRAPH.CENTER, fs=24, color="1F4E79")
    _par(doc, "CON FINES DE CIMENTACIÓN", True, al=WD_ALIGN_PARAGRAPH.CENTER, fs=18, color="1F4E79")
    _par(doc, "(NTE E.050 Suelos y Cimentaciones - Reglamento Nacional de Edificaciones)", False, True,
         WD_ALIGN_PARAGRAPH.CENTER, 11)
    doc.add_paragraph()
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for a, b in (("PROYECTO", proyecto.get("nombre", "")), ("SOLICITANTE", proyecto.get("solicitante", "")),
                 ("UBICACIÓN", ubic), ("CATEGORÍA (E.030)", f"{s['categoria']} - {CATEGORIA_DESC.get(s['categoria'], '')}"),
                 ("N° DE PISOS", str(proyecto.get("pisos", ""))),
                 ("RESPONSABLE", f"{proyecto.get('ingeniero', '')} - CIP N° {proyecto.get('cip', '')}"),
                 ("FECHA", fecha_larga())):
        c = t.add_row().cells
        c[0].text, c[1].text = a, b
        c[0].paragraphs[0].runs[0].bold = True
        _sombrear(c[0], "DDEBF7")
        c[0].width, c[1].width = Cm(4.5), Cm(11.5)
    for _ in range(6):
        doc.add_paragraph()
    _par(doc, f"{proyecto.get('provincia', '')} - PERÚ, {datetime.date.today().year}", True,
         al=WD_ALIGN_PARAGRAPH.CENTER, fs=12)
    doc.add_page_break()

    # ------------------------------------------------------------- ÍNDICE
    _par(doc, "ÍNDICE", True, al=WD_ALIGN_PARAGRAPH.CENTER, fs=14, color="1F4E79")
    _campo(doc.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u')
    _par(doc, "(Si el índice no se muestra, haga clic derecho sobre él y elija \"Actualizar campo\").", cursiva=True,
         fs=8.5, al=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # ------------------------------------------------------------- RESUMEN
    doc.add_heading("RESUMEN DE LAS CONDICIONES DE CIMENTACIÓN (Art. 16 NTE E.050)", 1)
    filas_res = [
        ["Tipo de cimentación", proyecto.get("tipo_cim") or "Zapatas aisladas conectadas con vigas y cimientos corridos"],
        ["Estrato de apoyo de la cimentación", f"{SUCS_DB[sucs_base(dis['sucs'])]['nombre']} ({dis['sucs']})"],
        ["Parámetros de diseño de la cimentación",
         f"Df = {Df:.2f} m ; φ = {dis['phi']:.1f}° ; c = {dis['c']:.2f} kg/cm² ; γ = {dis['g_base']:.2f} g/cm³"],
        ["Capacidad admisible de diseño", f"qadm = {dis['qadm']:.2f} kg/cm² (zapata {dis['tipo'].lower()} B = {dis['B']:.2f} m)"],
        ["Cimiento corrido (B = 0.60 m)", f"qadm = {calc['corrido']['qadm']:.2f} kg/cm²" if calc.get("corrido") else "-"],
        ["Factor de seguridad por corte", f"FS = {cfg['FS']:.1f} (cargas estáticas)"],
        ["Asentamiento diferencial máximo aceptable", f"{SMAX_E050:.2f} cm (1\") ; distorsión angular α ≤ 1/500"],
        ["Asentamiento esperado", f"Se = {dis['se']:.2f} cm"],
        ["Parámetros sísmicos (E.030)", f"Zona {s['zona']} (Z = {s['Z']:.2f}) ; Perfil {s['perfil']} (S = {s['S']:.2f}) ; "
                                        f"Tp = {s['Tp']:.1f} s ; TL = {s['Tl']:.1f} s"],
        ["Agresividad del suelo a la cimentación", f"{q['grado']} - {q['cemento']}"],
        ["Nivel freático", f"{nf:.2f} m" if nf else "No se encontró hasta la profundidad explorada"],
        ["Problemas especiales", "; ".join(t_ for t_, d_, rel in calc["especiales"] if rel) or "Ninguno relevante"],
    ]
    _tabla_word(doc, ["Concepto", "Valor / Recomendación"], filas_res, [6.0, 10.0], fs=9)
    doc.add_page_break()

    # ------------------------------------------------------------- 1. GENERALIDADES
    doc.add_heading("1. GENERALIDADES", 1)
    doc.add_heading("1.1 Objetivo del estudio", 2)
    _par(doc, f"El presente Estudio de Mecánica de Suelos (EMS) tiene por objetivo determinar las características "
              f"físico-mecánicas del subsuelo del terreno destinado al proyecto \"{proyecto.get('nombre', '')}\", "
              f"mediante trabajos de exploración de campo y ensayos de laboratorio, con el fin de establecer el tipo y "
              f"profundidad de cimentación, la capacidad admisible de carga, los asentamientos probables, los parámetros "
              f"sísmicos de diseño y la agresividad química del suelo al concreto, de acuerdo con la Norma Técnica E.050 "
              f"\"Suelos y Cimentaciones\" del Reglamento Nacional de Edificaciones.")
    doc.add_heading("1.2 Ubicación y descripción del proyecto", 2)
    _par(doc, f"El terreno en estudio se ubica en el {ubic}. El proyecto contempla una edificación de "
              f"{proyecto.get('pisos', '')} pisos, de categoría {s['categoria']} ({CATEGORIA_DESC.get(s['categoria'], '').lower()}) "
              f"según la NTE E.030, con un área aproximada de {proyecto.get('area') or '-'} m².")
    filas_ub = [["Región", proyecto.get("region", "")], ["Provincia", proyecto.get("provincia", "")],
                ["Distrito", proyecto.get("distrito", "")], ["Solicitante", proyecto.get("solicitante", "")],
                ["N° de pisos", str(proyecto.get("pisos", ""))], ["Área del terreno (m²)", str(proyecto.get("area") or "-")]]
    for cid in cals:
        inf = cal_info.get(cid, {})
        if inf.get("este"):
            filas_ub.append([f"Coordenadas UTM {cid} (WGS-84)", f"E {inf['este']:.0f} - N {inf['norte']:.0f}"])
    _tabla_word(doc, ["Dato", "Descripción"], filas_ub, [6.0, 10.0], fs=9)
    doc.add_heading("1.3 Normatividad aplicada", 2)
    _vinetas(doc, [
        "Reglamento Nacional de Edificaciones - Norma Técnica E.050 \"Suelos y Cimentaciones\".",
        "Reglamento Nacional de Edificaciones - Norma Técnica E.030 \"Diseño Sismorresistente\" (2018).",
        "Reglamento Nacional de Edificaciones - Norma Técnica E.060 \"Concreto Armado\" (Tabla 4.4).",
        "NTP 339.162 - Guía normalizada para caracterización de campo con fines de diseño (calicatas) / ASTM D420.",
        "NTP 339.150 - Descripción e identificación de suelos (procedimiento visual-manual) / ASTM D2488.",
        "NTP 339.127 - Método de ensayo para determinar el contenido de humedad de un suelo / ASTM D2216.",
        "NTP 339.128 - Método de ensayo para el análisis granulométrico / ASTM D422.",
        "NTP 339.129 - Determinación del límite líquido, límite plástico e índice de plasticidad / ASTM D4318.",
        "NTP 339.131 - Determinación del peso específico relativo de las partículas sólidas / ASTM D854.",
        "NTP 339.134 - Clasificación de suelos con propósitos de ingeniería (SUCS) / ASTM D2487.",
        "NTP 339.135 - Clasificación de suelos para uso en vías de transporte (AASHTO) / AASHTO M145.",
        "NTP 339.171 - Ensayo de corte directo en condiciones consolidadas drenadas / ASTM D3080.",
        "NTP 339.152, 339.177, 339.178 - Sales solubles totales, cloruros y sulfatos solubles en suelos.",
    ])
    doc.add_heading("1.4 Metodología y alcances", 2)
    _par(doc, "El estudio se desarrolló en tres etapas: (i) trabajos de campo, consistentes en la excavación de "
              f"{len(cals)} calicatas a cielo abierto, registro de la estratigrafía y toma de muestras alteradas e "
              "inalteradas; (ii) ensayos de laboratorio estándar y especiales; y (iii) trabajos de gabinete, que incluyen "
              "el análisis de la cimentación por resistencia al corte (Terzaghi, Meyerhof y Vesic) y por asentamiento "
              "(teoría elástica), la evaluación sísmica y química, y la elaboración del presente informe, de los perfiles "
              "estratigráficos (Lámina EMS-01) y de las hojas de cálculo.")
    _par(doc, "Los resultados y recomendaciones del presente estudio son válidos únicamente para el área explorada y "
              "para el proyecto descrito. Cualquier cambio en el tipo de estructura, número de pisos o ubicación obliga a "
              "verificar las conclusiones aquí expuestas (NTE E.050).")
    doc.add_page_break()

    # ------------------------------------------------------------- 2. GEOLOGÍA Y SISMICIDAD
    doc.add_heading("2. GEOLOGÍA, GEOMORFOLOGÍA Y SISMICIDAD DEL ÁREA EN ESTUDIO", 1)
    gruesos = [e for e in estratos if sucs_base(e["sucs"]) in ("GW", "GP", "GM", "GC", "SW", "SP", "SM", "SC")]
    tipo_dep = "depósitos aluviales cuaternarios" if gruesos else "depósitos coluvio-aluviales y lacustres de grano fino"
    doc.add_heading("2.1 Geología local", 2)
    _par(doc, f"Regionalmente, el área se emplaza sobre {tipo_dep}, conformados por gravas, arenas y finos en proporciones "
              "variables, depositados por la dinámica fluvial y torrencial del Cuaternario. Superficialmente se presenta "
              "una capa de relleno antrópico no controlado. Según la exploración realizada, la secuencia estratigráfica "
              "típica está formada por: " +
              "; ".join(f"{e['sucs']} ({SUCS_DB[sucs_base(e['sucs'])]['nombre'].lower()})"
                        for e in estratos_de(estratos, cals[0])) + ".")
    doc.add_heading("2.2 Geomorfología", 2)
    _par(doc, "El terreno presenta una topografía sensiblemente plana a ligeramente inclinada, sin evidencias de "
              "procesos geodinámicos activos (deslizamientos, erosión de laderas, flujos de detritos) que comprometan la "
              "estabilidad de la edificación proyectada. El drenaje superficial es moderado.")
    doc.add_heading("2.3 Sismicidad y parámetros sísmicos (NTE E.030-2018)", 2)
    _par(doc, "El territorio peruano se encuentra en una de las regiones de más alta actividad sísmica del mundo, debido "
              "a la subducción de la Placa de Nazca bajo la Placa Sudamericana. De acuerdo con la NTE E.030, el área "
              f"de estudio se ubica en la Zona Sísmica {s['zona']} y el suelo de cimentación corresponde al perfil "
              f"{s['perfil']}: {PERFIL_DESC[s['perfil']]}.")
    _tabla_word(doc, ["Factor", "Símbolo", "Valor", "Referencia E.030"], [
        ["Factor de zona", "Z", f"{s['Z']:.2f}", "Tabla N° 1"],
        ["Factor de uso", "U", f"{s['U']:.2f}", f"Tabla N° 5 (Cat. {s['categoria']})"],
        ["Factor de suelo", "S", f"{s['S']:.2f}", "Tabla N° 3"],
        ["Periodo de plataforma", "Tp (s)", f"{s['Tp']:.2f}", "Tabla N° 4"],
        ["Periodo de desplazamiento constante", "TL (s)", f"{s['Tl']:.2f}", "Tabla N° 4"],
        ["Periodo fundamental estimado", "T (s)", f"{s['T']:.3f}", f"T = hn/CT = {s['hn']:.2f}/{s['Ct']}"],
        ["Factor de amplificación sísmica", "C", f"{s['C']:.3f}", "Art. 14"],
        ["Producto Z·U·C·S", "ZUCS", f"{s['ZUCS']:.3f}", "Cortante basal V = ZUCS·P/R"],
    ], [5.5, 2.2, 2.0, 6.3], fs=9)
    _par(doc, "Factor de amplificación sísmica (Art. 14 E.030):", negrita=True)
    _eq(doc, "T < Tp :  C = 2.5        Tp < T < TL :  C = 2.5·(Tp / T)        T > TL :  C = 2.5·(Tp·TL / T²)")
    filas_c = []
    for T in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.8, 1.0, 1.2, 1.5, 2.0, 2.5, 3.0, 4.0]:
        C = factor_C(T, s["Tp"], s["Tl"]) if T > 0 else 2.5
        filas_c.append([f"{T:.2f}", f"{C:.3f}", f"{s['Z'] * s['U'] * C * s['S']:.3f}"])
    _tabla_word(doc, ["T (s)", "C", "Z·U·C·S (Sa/g para R = 1)"], filas_c, [3.0, 3.0, 6.0], fs=8.5)
    doc.add_page_break()

    # ------------------------------------------------------------- 3. INVESTIGACIÓN DE CAMPO
    doc.add_heading("3. INVESTIGACIÓN DE CAMPO", 1)
    doc.add_heading("3.1 Calicatas a cielo abierto (NTP 339.162)", 2)
    _par(doc, f"Se excavaron {len(cals)} calicatas a cielo abierto, distribuidas en el área del terreno de manera que "
              "representen las condiciones del subsuelo bajo la edificación proyectada, de acuerdo con el número mínimo "
              "de puntos de investigación de la NTE E.050. En cada calicata se registró la estratigrafía, se describió "
              "el suelo por el procedimiento visual-manual (NTP 339.150) y se obtuvieron muestras representativas que "
              "fueron embaladas, identificadas y trasladadas al laboratorio.")
    filas_cal = []
    for cid in cals:
        ec = estratos_de(estratos, cid)
        inf = cal_info.get(cid, {})
        nfc = inf.get("nf") or nf
        filas_cal.append([cid, f"{inf['este']:.0f}" if inf.get("este") else "-",
                          f"{inf['norte']:.0f}" if inf.get("norte") else "-",
                          f"{inf['cota']:.2f}" if inf.get("cota") else "-", f"{ec[-1]['fin']:.2f}",
                          f"{nfc:.2f}" if (nfc and nfc <= ec[-1]["fin"]) else "No encontrado",
                          str(sum(1 for e in ec if e.get("muestra") not in (None, "", "-")))])
    _tabla_word(doc, ["Calicata", "Este (m)", "Norte (m)", "Cota (msnm)", "Prof. (m)", "N.F. (m)", "N° muestras"],
                filas_cal, fs=9)
    doc.add_heading("3.2 Profundidad de investigación", 2)
    zmin = Df + 1.5 * dis["B"]
    pexp = min(estratos_de(estratos, c)[-1]["fin"] for c in cals)
    comp = estratos_de(estratos, cals[0])[-1]
    _par(doc, f"Según la NTE E.050, para edificaciones sin sótano la profundidad mínima de investigación es "
              f"p = Df + z, con z = 1.5·B. Para Df = {Df:.2f} m y B = {dis['B']:.2f} m: p = {Df:.2f} + 1.5 × {dis['B']:.2f} "
              f"= {zmin:.2f} m. La profundidad explorada mínima fue de {pexp:.2f} m.")
    if pexp + 1e-6 < zmin:
        if sucs_base(comp["sucs"]) in ("GW", "GP", "GM", "Roca"):
            _par(doc, f"Si bien la profundidad explorada es menor, las calicatas finalizaron en un estrato competente de "
                      f"{SUCS_DB[sucs_base(comp['sucs'])]['nombre'].lower()} ({comp['sucs']}) denso, cuya continuidad en "
                      "profundidad es conocida regionalmente, por lo que se considera representativa (NTE E.050). Se "
                      "recomienda verificar la continuidad del estrato durante la excavación de las zapatas.")
        else:
            _par(doc, "La profundidad explorada es menor a la requerida; se recomienda complementar la exploración con "
                      "ensayos de penetración (DPL / SPT) hasta la profundidad p indicada.", color="C00000")
    doc.add_heading("3.3 Muestreo", 2)
    _par(doc, "Se obtuvieron muestras alteradas en bolsas herméticas para ensayos estándar (humedad, granulometría, "
              "límites de consistencia) y muestras para ensayos especiales (corte directo y análisis químico). Las "
              "muestras se identificaron con el código de la calicata, el número de muestra y la profundidad.")
    doc.add_page_break()

    # ------------------------------------------------------------- 4. LABORATORIO
    doc.add_heading("4. ENSAYOS DE LABORATORIO ESTÁNDAR Y ESPECIALES", 1)
    doc.add_heading("4.1 Ensayos realizados", 2)
    _tabla_word(doc, ["Ensayo", "Norma NTP", "Norma ASTM"], [
        ["Contenido de humedad natural", "339.127", "D2216"], ["Análisis granulométrico por tamizado", "339.128", "D422"],
        ["Límite líquido y límite plástico", "339.129", "D4318"], ["Peso específico relativo de sólidos", "339.131", "D854"],
        ["Clasificación SUCS", "339.134", "D2487"], ["Clasificación AASHTO", "339.135", "D3282"],
        ["Corte directo (CD)", "339.171", "D3080"], ["Sales solubles totales", "339.152", "-"],
        ["Cloruros solubles", "339.177", "-"], ["Sulfatos solubles", "339.178", "-"]], [8.0, 3.5, 3.5], fs=9)
    doc.add_heading("4.2 Resumen de ensayos estándar", 2)
    filas_lab = []
    for e in est_ord:
        filas_lab.append([e["cal"], e.get("muestra") or "-", f"{e['ini']:.2f}-{e['fin']:.2f}", fmt(e.get("grava"), 0, "-"),
                          fmt(e.get("arena"), 0, "-"), fmt(e.get("finos"), 0, "-"), fmt(e.get("w"), 1, "-"),
                          fmt(e.get("ll"), 0), fmt(e.get("ip"), 0), e["sucs"], e.get("aashto") or "-"])
    _tabla_word(doc, ["Cal.", "Muestra", "Prof. (m)", "%G", "%S", "%F", "w %", "LL", "IP", "SUCS", "AASHTO"],
                filas_lab, fs=8)
    doc.add_heading("4.3 Análisis granulométrico (% que pasa)", 2)
    muestras = [e for e in est_ord if sucs_base(e["sucs"]) not in ("RELL", "Roca")]
    curvas = [coeficientes_curva(e) for e in muestras]
    filas_g = []
    for i, (tam, d) in enumerate(TAMICES):
        filas_g.append([tam, f"{d:.3f}"] + [f"{cv[0][i]:.1f}" for cv in curvas])
    filas_g.append(["Cu", ""] + [fmt(cv[4], 2, "-") for cv in curvas])
    filas_g.append(["Cc", ""] + [fmt(cv[5], 2, "-") for cv in curvas])
    _tabla_word(doc, ["Tamiz", "mm"] + [f"{e['cal']} {e.get('muestra') or e['est']}" for e in muestras], filas_g, fs=7.5)
    doc.add_heading("4.4 Ensayo de corte directo (NTP 339.171) y peso unitario", 2)
    filas_cd = []
    for e in est_ord:
        if sucs_base(e["sucs"]) in ("RELL",):
            continue
        filas_cd.append([e["cal"], e.get("muestra") or "-", e["sucs"], fmt(e.get("gamma"), 2, "-"), fmt(e.get("phi"), 1, "-"),
                         fmt(e.get("c"), 2, "-"), f"{SUCS_DB[sucs_base(e['sucs'])]['Es']}", f"{SUCS_DB[sucs_base(e['sucs'])]['mu']:.2f}"])
    _tabla_word(doc, ["Calicata", "Muestra", "SUCS", "γ (g/cm³)", "φ (°)", "c (kg/cm²)", "Es (kg/cm²)", "μ"], filas_cd,
                fs=8.5)
    _par(doc, "Los módulos de elasticidad (Es) y coeficientes de Poisson (μ) corresponden a valores típicos según el tipo "
              "de suelo SUCS (Bowles, Foundation Analysis and Design), adoptados del lado de la seguridad.", cursiva=True, fs=9)
    doc.add_page_break()

    # ------------------------------------------------------------- 5. PERFILES
    doc.add_heading("5. PERFILES ESTRATIGRÁFICOS", 1)
    _par(doc, "A continuación se describen los perfiles estratigráficos de cada calicata. El detalle gráfico completo, "
              "con la simbología SUCS, el perfil longitudinal de correlación y el cuadro resumen, se presenta en la "
              "Lámina EMS-01 (archivo AutoCAD .DXF).")
    tmpdir = tempfile.mkdtemp(prefix="ems_")
    for i, cid in enumerate(cals):
        ec = estratos_de(estratos, cid)
        inf = cal_info.get(cid, {})
        doc.add_heading(f"5.{i + 1} Calicata {cid}", 2)
        for e in ec:
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run(f"{e['est']} ({e['ini']:.2f} - {e['fin']:.2f} m) - {e['sucs']}: ")
            r.bold = True
            r.font.size = Pt(9.5)
            r2 = p.add_run(e.get("desc", ""))
            r2.font.size = Pt(9.5)
        png = os.path.join(tmpdir, f"registro_{cid}.png")
        imagen_registro_png(png, cid, ec, inf, proyecto, inf.get("nf") or nf, Df)
        doc.add_picture(png, width=Cm(15.0))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _par(doc, f"Figura 5.{i + 1}: Registro de exploración y perfil estratigráfico de la calicata {cid}.", cursiva=True,
             fs=9, al=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_page_break()

    # ------------------------------------------------------------- 6. ANÁLISIS DE LA CIMENTACIÓN
    doc.add_heading("6. ANÁLISIS DE LA CIMENTACIÓN", 1)
    doc.add_heading("6.1 Tipo y profundidad de cimentación", 2)
    _par(doc, f"Considerando la estratigrafía encontrada y las cargas de una edificación de {proyecto.get('pisos', '')} "
              f"pisos, se recomienda una cimentación superficial de {(proyecto.get('tipo_cim') or 'zapatas aisladas conectadas y cimientos corridos').lower()}, "
              f"con una profundidad mínima de desplante Df = {Df:.2f} m medida desde el nivel del terreno natural, "
              f"apoyada sobre el estrato de {SUCS_DB[sucs_base(dis['sucs'])]['nombre'].lower()} ({dis['sucs']}). "
              "Se deberá retirar íntegramente el material de relleno y no cimentar sobre suelos orgánicos.")
    doc.add_heading("6.2 Parámetros de diseño (calicata crítica)", 2)
    _tabla_word(doc, ["Parámetro", "Valor"], [
        ["Calicata / estrato crítico", f"{dis['cal']} / {dis['estrato']} ({dis['sucs']})"],
        ["Ángulo de fricción interna φ", f"{dis['phi']:.2f}°"], ["Cohesión c", f"{dis['c']:.3f} kg/cm² = {dis['c'] * 10:.2f} t/m²"],
        ["Peso unitario del estrato de apoyo γ", f"{dis['g_base']:.3f} t/m³"],
        ["Peso unitario promedio sobre Df", f"{dis['g_sobre']:.3f} t/m³"],
        ["Profundidad de desplante Df", f"{Df:.2f} m"], ["Ancho de cimentación B", f"{dis['B']:.2f} m"],
        ["Tipo de cimentación", f"{dis['tipo']}" + (f" (L = {dis['L']:.2f} m)" if dis["L"] else "")],
        ["Tipo de falla", f"Corte {cfg['falla'].lower()}"], ["Factor de seguridad", f"FS = {cfg['FS']:.1f}"],
        ["Nivel freático", f"{nf:.2f} m" if nf else "No encontrado"]], [7.0, 9.0], fs=9)
    for metodo in ("Terzaghi", "Meyerhof"):
        x = dis["res"][metodo]
        doc.add_heading(f"6.{3 if metodo == 'Terzaghi' else 4} Capacidad de carga por corte - Teoría de {metodo}", 2)
        if metodo == "Terzaghi":
            _par(doc, "Karl Terzaghi (1943) propuso la ecuación de capacidad última para cimentaciones superficiales "
                      "(Df ≤ B), considerando falla general por corte. La forma general utilizada es:")
        else:
            _par(doc, "G. G. Meyerhof (1963) amplió la ecuación incorporando factores de forma y de profundidad, que "
                      "consideran la resistencia al corte del suelo sobre el nivel de cimentación:")
        _eq(doc, "qu = c·Nc·Sc·dc + γ·Df·Nq·Sq·dq + ½·γ·B·Nγ·Sγ·dγ")
        _eq(doc, "qadm = qu / FS")
        if cfg["falla"].startswith("Local"):
            _par(doc, "Por tratarse de falla local (suelo suelto/compresible), se reducen los parámetros de resistencia:")
            _eq(doc, f"tan φ* = (2/3)·tan φ = (2/3)·tan({x['phi']:.2f}°)  →  φ* = {x['phi_u']:.2f}°")
            _eq(doc, f"c* = (2/3)·c = (2/3)·{x['c']:.3f} = {x['c_u']:.3f} kg/cm²")
        _par(doc, "Factores de capacidad de carga:", negrita=True)
        if metodo == "Terzaghi":
            _eq(doc, "Nq = e^{2(3π/4 − φ/2)·tanφ} / [2·cos²(45° + φ/2)]")
            _eq(doc, "Nc = (Nq − 1)·cot φ          Nγ = 2·(Nq + 1)·tan φ / (1 + 0.4·sen 4φ)")
        else:
            _eq(doc, "Nq = e^{π·tanφ}·tan²(45° + φ/2)")
            _eq(doc, "Nc = (Nq − 1)·cot φ          Nγ = (Nq − 1)·tan(1.4·φ)")
        _eq(doc, f"Para φ = {x['phi_u']:.2f}°:   Nc = {x['Nc']:.2f}   ;   Nq = {x['Nq']:.2f}   ;   Nγ = {x['Ng']:.2f}")
        _par(doc, "Factores de forma y profundidad:", negrita=True)
        if metodo == "Terzaghi":
            _eq(doc, "Corrida: Sc = 1.0, Sγ = 1.0  |  Cuadrada: Sc = 1.3, Sγ = 0.8  |  Rectangular: Sc = 1 + 0.3·B/L, Sγ = 1 − 0.2·B/L")
            _eq(doc, "Sq = 1.0 ; dc = dq = dγ = 1.0 (Terzaghi no considera la resistencia sobre Df)")
        else:
            _eq(doc, "Kp = tan²(45° + φ/2) ;  Sc = 1 + 0.2·Kp·(B/L) ;  Sq = Sγ = 1 + 0.1·Kp·(B/L)")
            _eq(doc, "dc = 1 + 0.2·√Kp·(Df/B) ;  dq = dγ = 1 + 0.1·√Kp·(Df/B)")
        _eq(doc, f"Sc = {x['sc']:.3f} ; Sq = {x['sq']:.3f} ; Sγ = {x['sg']:.3f} ; dc = {x['dc']:.3f} ; dq = {x['dq']:.3f} ; dγ = {x['dg']:.3f}")
        _par(doc, f"Sobrecarga efectiva al nivel de desplante ({x['caso_nf']}):", negrita=True)
        _eq(doc, f"q = γ·Df = {x['q']:.3f} t/m²   ;   γ (término Nγ) = {x['g2']:.3f} t/m³")
        _par(doc, "Reemplazando valores (unidades en t/m²):", negrita=True)
        _eq(doc, f"qu = {x['c_u'] * 10:.2f}·{x['Nc']:.2f}·{x['sc']:.3f}·{x['dc']:.3f} + {x['q']:.3f}·{x['Nq']:.2f}·{x['sq']:.3f}·{x['dq']:.3f}"
                 f" + 0.5·{x['g2']:.3f}·{x['B']:.2f}·{x['Ng']:.2f}·{x['sg']:.3f}·{x['dg']:.3f}")
        _eq(doc, f"qu = {x['t1']:.2f} + {x['t2']:.2f} + {x['t3']:.2f} = {x['qu_t']:.2f} t/m² = {x['qu']:.2f} kg/cm²")
        _eq(doc, f"qadm = {x['qu']:.2f} / {x['FS']:.1f} = {x['qadm']:.2f} kg/cm²", fs=12)
    doc.add_heading("6.5 Asentamiento elástico inmediato (Teoría de la Elasticidad)", 2)
    _par(doc, "El asentamiento inmediato se calcula con la expresión de Schleicher-Steinbrenner para un medio elástico "
              "semi-infinito, homogéneo e isótropo, con factor de influencia If para cimentación rígida (Bowles):")
    _eq(doc, "Se = q·B·(1 − μ²)·If / Es")
    _par(doc, "Donde: q = presión de trabajo (qadm), B = ancho de la cimentación, μ = coeficiente de Poisson, Es = módulo "
              "de elasticidad del suelo e If = factor de forma (cuadrada 0.82; L/B = 1.5: 1.06; L/B = 2: 1.20; corrida 2.10).")
    rc = dis["res"][cfg["metodo"]]
    _eq(doc, f"Se = {rc['qadm_t']:.2f}·{dis['B']:.2f}·(1 − {dis['mu']:.2f}²)·{dis['If']:.2f} / {dis['Es'] * 10:.0f} = "
             f"{dis['se_corte'] / 100:.4f} m = {dis['se_corte']:.2f} cm")
    _par(doc, f"Presión admisible para el asentamiento tolerable S máx = {cfg['smax']:.2f} cm:", negrita=True)
    _eq(doc, f"q(S máx) = S máx·Es / [B·(1 − μ²)·If] = {cfg['smax'] / 100:.4f}·{dis['Es'] * 10:.0f} / "
             f"[{dis['B']:.2f}·(1 − {dis['mu']:.2f}²)·{dis['If']:.2f}] = {dis['q_asent']:.2f} kg/cm²")
    doc.add_heading("6.6 Distorsión angular", 2)
    _par(doc, f"La NTE E.050 limita la distorsión angular a α = δ/L ≤ 1/500 para edificaciones convencionales. Con un "
              f"asentamiento diferencial δ ≈ {cfg['kdif']:.2f}·Se y una luz entre columnas L = {cfg['luz']:.2f} m, el "
              "asentamiento total límite y la presión asociada son:")
    _eq(doc, f"Se lím = (L/500)/{cfg['kdif']:.2f} = ({cfg['luz']:.2f}/500)/{cfg['kdif']:.2f} = {dis['se_lim_dist']:.2f} cm"
             f"   →   q(α) = {dis['q_dist']:.2f} kg/cm²")
    _eq(doc, f"qadm diseño = mín(qadm corte ; q(S máx) ; q(α)) = mín({dis['qadm_corte']:.2f} ; {dis['q_asent']:.2f} ; "
             f"{dis['q_dist']:.2f}) = {dis['qadm']:.2f} kg/cm²  →  gobierna: {dis['gobierna'].lower()}", fs=12)
    _par(doc, f"Verificación con la presión de diseño: Se = {dis['se']:.2f} cm ; δ = {dis['delta']:.2f} cm:")
    _eq(doc, f"α = δ / L = {dis['delta'] / 100:.4f} / {cfg['luz']:.2f} = 1/{(1 / dis['alfa']) if dis['alfa'] > 0 else 0:.0f}"
             f"  {'≤' if dis['ok_alfa'] else '>'} 1/500  →  {'CUMPLE' if dis['ok_alfa'] else 'NO CUMPLE'}")
    doc.add_heading("6.7 Matriz de capacidad admisible vs. asentamientos", 2)
    filas_m = []
    for f in calc["filas"]:
        filas_m.append([f"{f['Df']:.2f}", f"{f['B']:.2f}", f["tipo"][:4] + ".", f["cal"],
                        f"{f['res']['Terzaghi']['qadm']:.2f}", f"{f['res']['Meyerhof']['qadm']:.2f}",
                        f"{f['res']['Vesic']['qadm']:.2f}", f"{f['se_corte']:.2f}", f"{min(f['q_asent'], f['q_dist']):.2f}",
                        f"{f['qadm']:.2f}", f"{f['se']:.2f}"])
    _tabla_word(doc, ["Df (m)", "B (m)", "Tipo", "Cal.", "qadm Terz.", "qadm Meyer.", "qadm Vesic", "Se corte (cm)",
                      "q(S, α)", "qadm DISEÑO", "Se (cm)"], filas_m, fs=7.5)
    _par(doc, f"Valores en kg/cm². qadm de diseño = mínimo entre la capacidad por corte ({cfg['metodo']}, FS = "
              f"{cfg['FS']:.1f}), la presión que produce el asentamiento máximo tolerable de {cfg['smax']:.2f} cm "
              f"y la que cumple la distorsión angular 1/500 (L = {cfg['luz']:.2f} m). "
              "Se reporta la calicata más desfavorable en cada combinación.", cursiva=True, fs=9)
    doc.add_heading("6.8 Capacidad admisible de diseño", 2)
    _par(doc, f"Para una cimentación con zapata {dis['tipo'].lower()} de B = {dis['B']:.2f} m desplantada a "
              f"Df = {Df:.2f} m, se adopta una capacidad admisible de diseño de:", espacio=2)
    _par(doc, f"qadm = {dis['qadm']:.2f} kg/cm²", True, al=WD_ALIGN_PARAGRAPH.CENTER, fs=16, color="C00000")
    if calc.get("corrido"):
        _par(doc, f"Para cimientos corridos de B = 0.60 m a la misma profundidad: qadm = {calc['corrido']['qadm']:.2f} kg/cm².",
             al=WD_ALIGN_PARAGRAPH.CENTER, negrita=True)
    if dis["es_relleno"]:
        _par(doc, "ADVERTENCIA: el estrato al nivel de Df es relleno u orgánico; profundizar la cimentación hasta el "
                  "estrato natural competente.", True, color="C00000")
    doc.add_page_break()

    # ------------------------------------------------------------- 7. AGRESIVIDAD
    doc.add_heading("7. AGRESIVIDAD QUÍMICA DEL SUELO A LA CIMENTACIÓN", 1)
    _par(doc, "Se ensayaron muestras representativas del suelo de cimentación para determinar el contenido de sulfatos "
              "solubles, cloruros solubles, sales solubles totales y pH. La evaluación se realiza con la Tabla 4.4 de la "
              "NTE E.060 (ACI 318) para sulfatos y criterios usuales de la práctica geotécnica peruana para los demás.")
    _tabla_word(doc, ["Parámetro", "Resultado", "Grado de ataque", "Criterio de perjuicio"], [
        ["Sulfatos solubles SO₄ (ppm)", f"{q['so4']:.0f}", q["g_so4"], "> 1000 ppm (0.10 %)"],
        ["Cloruros Cl⁻ (ppm)", f"{q['cl']:.0f}", q["g_cl"], "> 6000 ppm"],
        ["Sales solubles totales (ppm)", f"{q['sst']:.0f}", q["g_sst"], "> 15000 ppm"],
        ["pH", f"{q['ph']:.2f}", q["g_ph"], "< 5.5"]], [5.5, 3.0, 3.5, 4.0], fs=9)
    _tabla_word(doc, ["Exposición a sulfatos", "SO₄ en suelo (% peso)", "Tipo de cemento", "a/c máx.", "f'c mín. (kg/cm²)"], [
        ["Despreciable", "0.00 ≤ SO₄ < 0.10", "Tipo I / sin restricción", "-", "-"],
        ["Moderado", "0.10 ≤ SO₄ < 0.20", "Tipo II, IP(MS), IS(MS), I(PM)(MS)", "0.50", "280"],
        ["Severo", "0.20 ≤ SO₄ ≤ 2.00", "Tipo V", "0.45", "315"],
        ["Muy severo", "SO₄ > 2.00", "Tipo V + puzolana", "0.45", "315"]], fs=8.5, color="7F6000")
    _par(doc, f"Conclusión: el grado de ataque global del suelo al concreto es {q['grado'].upper()}. Se recomienda "
              f"emplear cemento {q['cemento']}, con relación agua/cemento máxima de {q['ac']:.2f} y f'c mínimo de "
              f"{q['fc']} kg/cm² en los elementos de cimentación.", negrita=True)
    _vinetas(doc, q["medidas"])
    doc.add_page_break()

    # ------------------------------------------------------------- 8. PROBLEMAS ESPECIALES / CONCLUSIONES
    doc.add_heading("8. PROBLEMAS ESPECIALES, MUROS DE CONTENCIÓN Y CONCLUSIONES", 1)
    doc.add_heading("8.1 Problemas especiales de cimentación (Cap. 6 NTE E.050)", 2)
    for tit, desc, _rel in calc["especiales"]:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(tit + ": ")
        r.bold = True
        p.add_run(desc)
    doc.add_heading("8.2 Parámetros para diseño de muros de contención y calzaduras", 2)
    _par(doc, "Coeficientes de empuje de tierras según Rankine (activo y pasivo), Jaky (reposo) y coeficiente de fricción "
              "suelo-concreto μ = tan(2φ/3):")
    _eq(doc, "Ka = tan²(45° − φ/2)     Kp = tan²(45° + φ/2)     K0 = 1 − sen φ")
    _tabla_word(doc, ["Cal.", "Estrato", "SUCS", "γ (t/m³)", "φ (°)", "c (kg/cm²)", "Ka", "Kp", "K0", "μ"],
                [[em["cal"], em["est"], em["sucs"], f"{em['gamma']:.2f}", f"{em['phi']:.1f}", f"{em['c']:.2f}",
                  f"{em['Ka']:.3f}", f"{em['Kp']:.3f}", f"{em['K0']:.3f}", f"{em['mu']:.3f}"] for em in calc["empujes"]],
                fs=8)
    _par(doc, "Calzaduras: en excavaciones colindantes con edificaciones existentes se deberán ejecutar calzaduras de "
              "concreto ciclópeo por paños alternados de 1.00 a 1.50 m de ancho, conforme a la NTE E.050 (Cap. 7), "
              "diseñadas con los parámetros anteriores y bajo responsabilidad del ingeniero estructural.")
    doc.add_heading("8.3 Conclusiones", 2)
    con = [
        f"Se ejecutaron {len(cals)} calicatas a cielo abierto hasta una profundidad máxima de "
        f"{max(e['fin'] for e in estratos):.2f} m.",
        "El perfil estratigráfico está conformado superficialmente por relleno no controlado, seguido de " +
        ", ".join(sorted({SUCS_DB[sucs_base(e['sucs'])]['nombre'].lower() + f" ({e['sucs']})" for e in estratos
                          if sucs_base(e['sucs']) != 'RELL'})) + ".",
        (f"El nivel freático se detectó a {nf:.2f} m de profundidad." if nf else
         "No se detectó nivel freático hasta la profundidad explorada."),
        f"La cimentación se apoyará sobre {SUCS_DB[sucs_base(dis['sucs'])]['nombre'].lower()} ({dis['sucs']}) a una "
        f"profundidad mínima Df = {Df:.2f} m.",
        f"La capacidad admisible de diseño es qadm = {dis['qadm']:.2f} kg/cm² para zapatas de B = {dis['B']:.2f} m"
        + (f", y qadm = {calc['corrido']['qadm']:.2f} kg/cm² para cimientos corridos de B = 0.60 m." if calc.get("corrido") else "."),
        f"El asentamiento elástico esperado es de {dis['se']:.2f} cm (máximo tolerable {cfg['smax']:.2f} cm) y la "
        f"distorsión angular es 1/{(1 / dis['alfa']) if dis['alfa'] > 0 else 0:.0f} (límite 1/500); gobierna el "
        f"diseño: {dis['gobierna'].lower()}.",
        f"Parámetros sísmicos E.030: Z = {s['Z']:.2f}, U = {s['U']:.2f}, S = {s['S']:.2f}, Tp = {s['Tp']:.1f} s, "
        f"TL = {s['Tl']:.1f} s (perfil {s['perfil']}).",
        f"La agresividad del suelo al concreto es {q['grado'].lower()}; usar {q['cemento']}.",
    ]
    for i, c_ in enumerate(con, 1):
        _par(doc, f"{i}. {c_}", espacio=3)
    doc.add_heading("8.4 Recomendaciones (obligatorias según NTE E.050)", 2)
    recs = [
        "Eliminar íntegramente el material de relleno y la capa orgánica superficial; no cimentar sobre ellos.",
        f"Desplantar la cimentación a una profundidad mínima de {Df:.2f} m desde el terreno natural, sobre el estrato "
        f"natural de {dis['sucs']}. Si al excavar se encuentra material diferente al descrito, comunicar al profesional "
        "responsable del EMS.",
        "Colocar un solado de concreto f'c = 100 kg/cm² de 0.10 m de espesor como base de las zapatas.",
        "Conectar las zapatas aisladas mediante vigas de cimentación para controlar asentamientos diferenciales.",
        "Compactar el fondo de las excavaciones con equipo manual o mecánico antes de vaciar el solado.",
        "Los rellenos que se coloquen sobre la cimentación deberán ser controlados, compactados en capas de 0.20 m al "
        "95 % de la Máxima Densidad Seca del Proctor Modificado (NTP 339.141).",
        f"Emplear cemento {q['cemento']} en todos los elementos en contacto con el suelo, con a/c ≤ {q['ac']:.2f}.",
        "Proteger las excavaciones contra la humedad y el ingreso de agua; prever drenaje perimetral.",
        "Las conclusiones de este estudio solo son válidas para el terreno y proyecto descritos (NTE E.050).",
    ]
    for i, c_ in enumerate(recs, 1):
        _par(doc, f"{i}. {c_}", espacio=3)
    doc.add_paragraph()
    _par(doc, "_______________________________", al=WD_ALIGN_PARAGRAPH.CENTER)
    _par(doc, f"{proyecto.get('ingeniero', '')}", True, al=WD_ALIGN_PARAGRAPH.CENTER, espacio=0)
    _par(doc, f"Ingeniero Civil - CIP N° {proyecto.get('cip', '')}", al=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # ------------------------------------------------------------- 9. ANEXO FOTOGRÁFICO
    doc.add_heading("9. ANEXO FOTOGRÁFICO DE LAS CALICATAS", 1)
    fotos = []
    for cid in cals:
        for f in cal_info.get(cid, {}).get("fotos", []):
            fotos.append((cid, f))
    if not fotos:
        _par(doc, "No se vincularon fotografías de las calicatas.", cursiva=True)
    n = 0
    for i in range(0, len(fotos), 2):
        par = fotos[i:i + 2]
        t = doc.add_table(rows=len(par), cols=1)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, (cid, f) in enumerate(par):
            cel = t.rows[j].cells[0]
            p = cel.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ruta, tam = _preparar_imagen(f["path"], tmpdir, n)
            n += 1
            if ruta:
                w_cm = 15.0
                if tam[1] / tam[0] * w_cm > 9.5:
                    w_cm = 9.5 * tam[0] / tam[1]
                p.add_run().add_picture(ruta, width=Cm(w_cm))
            else:
                p.add_run(f"[No se pudo cargar la imagen: {os.path.basename(f['path'])}]")
            cp = cel.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr = cp.add_run(f"Fotografía N° {n}: ")
            rr.bold = True
            rr.font.size = Pt(9)
            r2 = cp.add_run(f.get("desc") or f"Calicata {cid}.")
            r2.font.size = Pt(9)
        if i + 2 < len(fotos):
            doc.add_page_break()
    doc.save(path)
    return path


# =============================================================================
# 14. UTILIDADES DE SISTEMA
# =============================================================================
def abrir_archivo(path):
    try:
        if hasattr(os, "startfile"):
            os.startfile(path)  # Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def modo_demo(carpeta):
    os.makedirs(carpeta, exist_ok=True)
    proyecto = proyecto_ejemplo()
    estratos = estratos_ejemplo()
    info = info_ejemplo()
    generar_fotos_ejemplo(os.path.join(carpeta, "fotos_ejemplo"), estratos, info)
    calc = calcular_todo(proyecto, info, estratos)
    p1 = exportar_word(os.path.join(carpeta, "Informe_EMS_E050.docx"), proyecto, info, estratos, calc)
    p2 = exportar_dxf(os.path.join(carpeta, "Lamina_EMS-01_Perfiles.dxf"), proyecto, info, estratos, calc)
    p3 = exportar_excel(os.path.join(carpeta, "Hojas_Calculo_EMS.xlsx"), proyecto, info, estratos, calc)
    d = calc["diseno"]
    print(f"qadm diseño = {d['qadm']:.2f} kg/cm²  (Df={d['Df']:.2f} m, B={d['B']:.2f} m, {d['cal']}, {d['sucs']}), "
          f"Se = {d['se']:.2f} cm")
    for p in (p1, p2, p3):
        print("Generado:", p)


# =============================================================================
# 15. INTERFAZ GRÁFICA (CustomTkinter)
# =============================================================================
_BaseCTk = ctk.CTk if GUI_OK else object
_BaseTop = ctk.CTkToplevel if GUI_OK else object

CAMPOS_ESTRATO = [("cal", "Calicata"), ("est", "Estrato"), ("ini", "Prof. inicial (m)"), ("fin", "Prof. final (m)"),
                  ("muestra", "Muestra"), ("sucs", "Clasificación SUCS"), ("aashto", "Clasificación AASHTO"),
                  ("grava", "% Grava"), ("arena", "% Arena"), ("finos", "% Finos (pasa N°200)"),
                  ("w", "Humedad natural w (%)"), ("ll", "LL (%)  [vacío = NP]"), ("ip", "IP (%)  [vacío = NP]"),
                  ("gamma", "Peso específico γ (g/cm³)"), ("phi", "Ángulo de fricción φ (°)"),
                  ("c", "Cohesión c (kg/cm²)"), ("cu", "Cu (opcional)"), ("cc", "Cc (opcional)"),
                  ("desc", "Descripción litológica")]
NUMERICOS = {"ini", "fin", "grava", "arena", "finos", "w", "ll", "ip", "gamma", "phi", "c", "cu", "cc"}

COLS_TABLA = [("cal", "Calicata", 62), ("est", "Estrato", 56), ("prof", "Prof. (m)", 90), ("muestra", "Muestra", 58),
              ("sucs", "SUCS", 58), ("aashto", "AASHTO", 78), ("w", "w (%)", 54), ("ll", "LL (%)", 52),
              ("ip", "IP (%)", 52), ("gamma", "γ (g/cm³)", 70), ("phi", "φ (°)", 50), ("c", "c (kg/cm²)", 72),
              ("desc", "Descripción litológica", 520)]

COLS_CALC = [("Df", "Df (m)", 52), ("B", "B (m)", 50), ("tipo", "Tipo", 82), ("cal", "Cal.", 46), ("sucs", "SUCS", 50),
             ("phi", "φ (°)", 48), ("c", "c", 46), ("Nc", "Nc", 56), ("Nq", "Nq", 56), ("Ng", "Nγ", 56),
             ("qu", "qu (kg/cm²)", 80), ("qadmc", "qadm corte", 78), ("se", "ΔH corte (cm)", 88),
             ("qs", "q(ΔH, α)", 80), ("qadm", "qadm DISEÑO", 90), ("sed", "ΔH (cm)", 62), ("ok", "Verif.", 70)]


class DialogoEstrato(_BaseTop):
    def __init__(self, master, datos, al_aceptar):
        super().__init__(master)
        self.title("Datos del estrato")
        self.geometry("560x720")
        self.al_aceptar = al_aceptar
        self.vars = {}
        fr = ctk.CTkScrollableFrame(self, width=520, height=620)
        fr.pack(fill="both", expand=True, padx=10, pady=(10, 4))
        for i, (k, lab) in enumerate(CAMPOS_ESTRATO):
            ctk.CTkLabel(fr, text=lab, anchor="w").grid(row=i, column=0, sticky="w", padx=6, pady=3)
            v = datos.get(k)
            txt = "" if v is None else (f"{v:g}" if isinstance(v, float) else str(v))
            var = tk.StringVar(value=txt)
            self.vars[k] = var
            if k == "sucs":
                w = ctk.CTkComboBox(fr, values=SUCS_LISTA, variable=var, width=280)
            else:
                w = ctk.CTkEntry(fr, textvariable=var, width=280)
            w.grid(row=i, column=1, sticky="ew", padx=6, pady=3)
        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(fill="x", padx=10, pady=8)
        ctk.CTkButton(bf, text="Autocompletar φ, c, γ típicos del SUCS", command=self._tipicos,
                      fg_color="#555").pack(side="left", padx=4)
        ctk.CTkButton(bf, text="Aceptar", command=self._ok, width=110).pack(side="right", padx=4)
        ctk.CTkButton(bf, text="Cancelar", command=self.destroy, width=110, fg_color="#777").pack(side="right", padx=4)
        self.after(150, self._modal)

    def _modal(self):
        try:
            self.grab_set()
            self.focus_force()
        except Exception:
            pass

    def _tipicos(self):
        db = SUCS_DB[sucs_base(self.vars["sucs"].get())]
        if not self.vars["phi"].get().strip():
            self.vars["phi"].set(str(db["phi"]))
        if not self.vars["c"].get().strip():
            self.vars["c"].set(str(db["c"]))
        if not self.vars["gamma"].get().strip():
            self.vars["gamma"].set("1.80")
        if not self.vars["aashto"].get().strip():
            self.vars["aashto"].set(db["aashto"])
        if not self.vars["desc"].get().strip():
            self.vars["desc"].set(db["desc"] + ".")

    def _ok(self):
        d = {}
        for k, var in self.vars.items():
            s = var.get().strip()
            d[k] = num(s) if k in NUMERICOS else s
        if not d["cal"] or d["ini"] is None or d["fin"] is None or d["fin"] <= d["ini"]:
            messagebox.showerror("Datos incompletos", "Indique la calicata y profundidades válidas (fin > inicio).",
                                 parent=self)
            return
        d["sucs"] = d["sucs"] or "SM"
        d["gamma"] = d["gamma"] or 1.70
        d["c"] = d["c"] or 0.0
        if d["phi"] is None:
            d["phi"] = float(SUCS_DB[sucs_base(d["sucs"])]["phi"])
        self.al_aceptar(d)
        self.destroy()


class DialogoCalicata(_BaseTop):
    def __init__(self, master, cid, info, al_aceptar):
        super().__init__(master)
        self.title(f"Datos de la calicata {cid}")
        self.geometry("420x300")
        self.al_aceptar, self.cid = al_aceptar, cid
        self.vars = {}
        for i, (k, lab) in enumerate((("este", "Este UTM (m)"), ("norte", "Norte UTM (m)"),
                                      ("cota", "Cota del terreno (msnm)"), ("nf", "N.F. propio (m) [vacío = general]"),
                                      ("fecha", "Fecha de excavación"))):
            ctk.CTkLabel(self, text=lab).grid(row=i, column=0, sticky="w", padx=10, pady=5)
            v = info.get(k)
            var = tk.StringVar(value="" if v is None else (f"{v:g}" if isinstance(v, float) else str(v)))
            self.vars[k] = var
            ctk.CTkEntry(self, textvariable=var, width=200).grid(row=i, column=1, padx=10, pady=5)
        ctk.CTkButton(self, text="Aceptar", command=self._ok).grid(row=6, column=1, pady=14)
        self.after(150, lambda: (self.grab_set(), self.focus_force()))

    def _ok(self):
        d = {k: (num(v.get()) if k != "fecha" else v.get().strip()) for k, v in self.vars.items()}
        self.al_aceptar(self.cid, d)
        self.destroy()


class AppEMS(_BaseCTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title(f"{APP_CORTO}  -  {APP_TITULO}")
        self.geometry("1280x820")
        self.minsize(1100, 700)
        self.estratos = []
        self.cal_info = {}
        self.calc = None
        self._job = None
        self.carpeta_salida = os.path.join(os.path.expanduser("~"), "EMS_E050_Resultados")
        self._construir_cabecera()
        self.tabs = ctk.CTkTabview(self, corner_radius=10)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=(0, 4))
        self.t1 = self.tabs.add("1. Proyecto, Sismicidad E.030 y Química del Suelo")
        self.t2 = self.tabs.add("2. Calicatas, Estratigrafía y Laboratorio")
        self.t3 = self.tabs.add("3. Capacidad Portante y Asentamientos")
        self._estilo_tablas()
        self._construir_tab1()
        self._construir_tab2()
        self._construir_tab3()
        self._construir_barra()
        self._actualizar_resumen_tab1()

    # ------------------------------------------------------------------ cabecera
    def _construir_cabecera(self):
        fr = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1F4E79", "#14324f"))
        fr.pack(fill="x")
        ctk.CTkLabel(fr, text="Sistema Automatizador de Estudios de Mecánica de Suelos (EMS - NTE E.050)",
                     font=ctk.CTkFont(size=19, weight="bold"), text_color="white").pack(side="left", padx=14, pady=(8, 0))
        der = ctk.CTkFrame(fr, fg_color="transparent")
        der.pack(side="right", padx=10)
        ctk.CTkLabel(der, text="Tema:", text_color="white").pack(side="left", padx=4)
        self.tema = ctk.CTkOptionMenu(der, values=["Oscuro", "Claro"], width=100, command=self._cambiar_tema)
        self.tema.pack(side="left", pady=8)
        sub = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1F4E79", "#14324f"))
        sub.pack(fill="x")
        ctk.CTkLabel(sub, text="Capacidad Portante (Terzaghi · Meyerhof · Vesic)  ·  Perfiles SUCS en AutoCAD (.DXF)  ·  "
                                "Informe en Word (.DOCX)  ·  Hojas de cálculo (.XLSX)  ·  100 % offline",
                     text_color="#cfe2f3", font=ctk.CTkFont(size=12)).pack(side="left", padx=14, pady=(0, 8))

    def _cambiar_tema(self, val):
        ctk.set_appearance_mode("dark" if val == "Oscuro" else "light")
        self._estilo_tablas()

    def _estilo_tablas(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        oscuro = ctk.get_appearance_mode() == "Dark"
        bg, fg, sel = ("#2b2b2b", "#e8e8e8", "#1f6aa5") if oscuro else ("#ffffff", "#111111", "#9fc5e8")
        hb = "#1F4E79"
        st.configure("EMS.Treeview", background=bg, fieldbackground=bg, foreground=fg, rowheight=24,
                     font=("Segoe UI", 10), borderwidth=0)
        st.map("EMS.Treeview", background=[("selected", sel)], foreground=[("selected", "white" if oscuro else "black")])
        st.configure("EMS.Treeview.Heading", background=hb, foreground="white", font=("Segoe UI", 10, "bold"),
                     relief="flat")
        st.map("EMS.Treeview.Heading", background=[("active", "#2a6099")])
        for tv in (getattr(self, "tree", None), getattr(self, "tree_calc", None)):
            if tv is not None:
                tv.tag_configure("par", background="#333333" if oscuro else "#eef3fa")
                tv.tag_configure("impar", background=bg)
                tv.tag_configure("mal", foreground="#ff6b6b" if oscuro else "#c00000")

    # ------------------------------------------------------------------ TAB 1
    def _construir_tab1(self):
        sc = ctk.CTkScrollableFrame(self.t1)
        sc.pack(fill="both", expand=True)
        sc.grid_columnconfigure((0, 1), weight=1)
        self.v = {}

        def seccion(parent, titulo, fila, col):
            f = ctk.CTkFrame(parent, corner_radius=10)
            f.grid(row=fila, column=col, sticky="nsew", padx=8, pady=8)
            ctk.CTkLabel(f, text=titulo, font=ctk.CTkFont(size=15, weight="bold"),
                         text_color=("#1F4E79", "#7fb3e6")).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 4))
            f.grid_columnconfigure(1, weight=1)
            return f

        def campo(f, fila, clave, etiqueta, valor="", valores=None, ancho=320):
            ctk.CTkLabel(f, text=etiqueta, anchor="w").grid(row=fila, column=0, sticky="w", padx=10, pady=4)
            var = tk.StringVar(value=valor)
            if valores:
                w = ctk.CTkOptionMenu(f, values=valores, variable=var, width=ancho)
            else:
                w = ctk.CTkEntry(f, textvariable=var, width=ancho)
            w.grid(row=fila, column=1, sticky="ew", padx=10, pady=4)
            var.trace_add("write", lambda *a: self._cambio_datos())
            self.v[clave] = var
            return var

        f1 = seccion(sc, "Datos generales del proyecto", 0, 0)
        campo(f1, 1, "nombre", "Nombre del proyecto", "")
        campo(f1, 2, "solicitante", "Solicitante / Entidad", "")
        campo(f1, 3, "distrito", "Distrito", "")
        campo(f1, 4, "provincia", "Provincia", "")
        campo(f1, 5, "region", "Región", "")
        campo(f1, 6, "area", "Área del terreno (m²)", "")
        campo(f1, 7, "tipo_cim", "Tipo de cimentación prevista", "Zapatas aisladas y cimientos corridos")

        f2 = seccion(sc, "Sismicidad NTE E.030-2018", 0, 1)
        campo(f2, 1, "categoria", "Categoría de la edificación", "C", ["A", "B", "C"])
        campo(f2, 2, "pisos", "Número de pisos", "2")
        campo(f2, 3, "zona", "Zona sísmica", "4", ["1", "2", "3", "4"])
        campo(f2, 4, "perfil", "Perfil de suelo", "S2", ["S0", "S1", "S2", "S3"])
        campo(f2, 5, "nf", "Nivel freático (m)  [0 o vacío = no encontrado]", "")
        self.lbl_sismo = ctk.CTkLabel(f2, text="", justify="left", anchor="w", font=ctk.CTkFont(family="Consolas", size=12))
        self.lbl_sismo.grid(row=6, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 10))

        f3 = seccion(sc, "Química del suelo (E.060 / ACI 318)", 1, 0)
        campo(f3, 1, "so4", "Sulfatos solubles SO₄ (ppm)", "0")
        campo(f3, 2, "cl", "Cloruros Cl⁻ (ppm)", "0")
        campo(f3, 3, "sst", "Sales solubles totales (ppm)", "0")
        campo(f3, 4, "ph", "pH", "7.0")
        self.lbl_quim = ctk.CTkLabel(f3, text="", justify="left", anchor="w", wraplength=520,
                                     font=ctk.CTkFont(size=12))
        self.lbl_quim.grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=(6, 10))

        f4 = seccion(sc, "Profesional responsable", 1, 1)
        campo(f4, 1, "ingeniero", "Ingeniero Civil responsable", "")
        campo(f4, 2, "cip", "N° CIP", "")
        ctk.CTkLabel(f4, text="Carpeta de salida:", anchor="w").grid(row=3, column=0, sticky="w", padx=10, pady=4)
        self.lbl_carpeta = ctk.CTkLabel(f4, text=self.carpeta_salida, anchor="w", wraplength=330)
        self.lbl_carpeta.grid(row=3, column=1, sticky="w", padx=10)
        ctk.CTkButton(f4, text="Cambiar carpeta de salida…", command=self._elegir_carpeta).grid(
            row=4, column=1, sticky="w", padx=10, pady=(4, 10))

    def _elegir_carpeta(self):
        d = filedialog.askdirectory(initialdir=self.carpeta_salida)
        if d:
            self.carpeta_salida = d
            self.lbl_carpeta.configure(text=d)

    def get_proyecto(self):
        return {k: v.get() for k, v in self.v.items()}

    def _actualizar_resumen_tab1(self):
        p = self.get_proyecto()
        try:
            s = parametros_sismicos(p["zona"], p["perfil"], p["categoria"], num(p["pisos"], 1))
            self.lbl_sismo.configure(text=f"Z = {s['Z']:.2f}   U = {s['U']:.2f}   S = {s['S']:.2f}\n"
                                          f"Tp = {s['Tp']:.2f} s   TL = {s['Tl']:.2f} s\n"
                                          f"T ≈ {s['T']:.3f} s   C = {s['C']:.3f}   ZUCS = {s['ZUCS']:.3f}\n"
                                          f"{PERFIL_DESC[s['perfil']]}")
        except Exception as ex:
            self.lbl_sismo.configure(text=f"Revise los datos sísmicos ({ex})")
        q = agresividad_quimica(num(p["so4"]), num(p["cl"]), num(p["sst"]), num(p["ph"]))
        col = {"Despreciable": "#2e8b57", "Moderado": "#d9a400", "Severo": "#e06c00", "Muy Severo": "#d00000"}[q["grado"]]
        self.lbl_quim.configure(text=f"Grado de ataque: {q['grado']}  (SO₄: {q['g_so4']}, Cl⁻: {q['g_cl']}, "
                                     f"SST: {q['g_sst']}, pH: {q['g_ph']})\nCemento: {q['cemento']}\n"
                                     f"a/c máx = {q['ac']:.2f}   f'c mín = {q['fc']} kg/cm²", text_color=col)

    def _cambio_datos(self):
        if hasattr(self, "lbl_quim"):
            self._actualizar_resumen_tab1()
        self._programar_calculo()

    # ------------------------------------------------------------------ TAB 2
    def _construir_tab2(self):
        top = ctk.CTkFrame(self.t2, fg_color="transparent")
        top.pack(fill="x", pady=(2, 4))
        ctk.CTkButton(top, text="1. Importar Datos de Calicatas y Estratigrafía desde Excel (.XLSX)",
                      command=self.importar_excel, height=36, fg_color="#1f6aa5").pack(side="left", padx=4)
        ctk.CTkButton(top, text="Cargar Estudio Completo de Ejemplo (3 Calicatas de 3.00 m con Estratos SM, GP y CL + Fotos)",
                      command=self.cargar_ejemplo, height=36, fg_color="#7b3fa0", hover_color="#5e2f7a").pack(side="left", padx=4)
        ctk.CTkButton(top, text="Plantilla Excel", command=self.guardar_plantilla, height=36, width=120,
                      fg_color="#4a4a4a").pack(side="left", padx=4)
        bar = ctk.CTkFrame(self.t2, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 4))
        for txt, cmd, col in (("➕ Añadir estrato", self.anadir_estrato, "#2e8b57"),
                              ("✎ Editar estrato", self.editar_estrato, "#1f6aa5"),
                              ("🗑 Eliminar estrato", self.eliminar_estrato, "#b03a2e"),
                              ("⚙ Clasificar SUCS / AASHTO automático", self.autoclasificar, "#8a6d00"),
                              ("📍 Datos de la calicata (UTM, cota)", self.editar_calicata, "#555555"),
                              ("📷 Vincular fotos a la calicata seleccionada", self.vincular_fotos, "#c0392b")):
            ctk.CTkButton(bar, text=txt, command=cmd, fg_color=col, height=30).pack(side="left", padx=4)
        fr = ctk.CTkFrame(self.t2)
        fr.pack(fill="both", expand=True)
        cols = [c for c, _, _ in COLS_TABLA]
        self.tree = ttk.Treeview(fr, columns=cols, show="headings", style="EMS.Treeview", selectmode="browse")
        for c, t, w in COLS_TABLA:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w" if c == "desc" else "center", stretch=(c == "desc"))
        ys = ttk.Scrollbar(fr, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(fr, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        fr.grid_rowconfigure(0, weight=1)
        fr.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", lambda e: self.editar_estrato())
        self.lbl_cal = ctk.CTkLabel(self.t2, text="Sin datos. Importe un Excel o cargue el ejemplo.", anchor="w",
                                    justify="left")
        self.lbl_cal.pack(fill="x", padx=6, pady=4)
        self._estilo_tablas()

    def refrescar_tabla(self):
        self.tree.delete(*self.tree.get_children())
        self.estratos.sort(key=lambda e: (orden_natural(e["cal"]), e["ini"]))
        for i, e in enumerate(self.estratos):
            vals = [e["cal"], e["est"], f"{e['ini']:.2f} - {e['fin']:.2f}", e.get("muestra") or "-", e["sucs"],
                    e.get("aashto") or "-", fmt(e.get("w"), 1, "-"), fmt(e.get("ll"), 0), fmt(e.get("ip"), 0),
                    fmt(e.get("gamma"), 2, "-"), fmt(e.get("phi"), 1, "-"), fmt(e.get("c"), 2, "-"), e.get("desc", "")]
            self.tree.insert("", "end", iid=str(i), values=vals, tags=("par" if i % 2 else "impar",))
        for cid in calicatas_de(self.estratos):
            self.cal_info.setdefault(cid, dict(este=None, norte=None, cota=None, nf=None, fotos=[],
                                               fecha=datetime.date.today().strftime("%d/%m/%Y")))
        txt = []
        for cid in calicatas_de(self.estratos):
            inf = self.cal_info[cid]
            ec = estratos_de(self.estratos, cid)
            txt.append(f"{cid}: {len(ec)} estratos, prof. {ec[-1]['fin']:.2f} m, {len(inf.get('fotos', []))} foto(s)")
        self.lbl_cal.configure(text="   |   ".join(txt) if txt else "Sin datos.")
        vals = ["Envolvente"] + calicatas_de(self.estratos)
        self.om_cal.configure(values=vals)
        if self.v3["calicata"].get() not in vals:
            self.v3["calicata"].set("Envolvente")
        self._programar_calculo()

    def _fila_sel(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def anadir_estrato(self):
        base = {}
        i = self._fila_sel()
        if i is not None:
            e = self.estratos[i]
            sig = [x for x in self.estratos if x["cal"] == e["cal"]]
            base = dict(cal=e["cal"], est=f"E-{len(sig) + 1}", ini=max(x["fin"] for x in sig), muestra=f"M-{len(sig)}")
        DialogoEstrato(self, base, self._agregar)

    def _agregar(self, d):
        self.estratos.append(d)
        self.refrescar_tabla()

    def editar_estrato(self):
        i = self._fila_sel()
        if i is None:
            messagebox.showinfo("Editar", "Seleccione un estrato de la tabla.")
            return

        def ok(d, i=i):
            self.estratos[i] = d
            self.refrescar_tabla()
        DialogoEstrato(self, self.estratos[i], ok)

    def eliminar_estrato(self):
        i = self._fila_sel()
        if i is None:
            return
        e = self.estratos[i]
        if messagebox.askyesno("Eliminar", f"¿Eliminar el estrato {e['est']} de {e['cal']}?"):
            del self.estratos[i]
            self.refrescar_tabla()

    def autoclasificar(self):
        if not self.estratos:
            return
        cambios = []
        for e in self.estratos:
            if e.get("grava") is None and e.get("finos") is None:
                continue
            s, a = autoclasificar_estrato(e)
            if s != e["sucs"] or a != e.get("aashto"):
                cambios.append(f"{e['cal']} {e['est']}: {e['sucs']} → {s} ; {e.get('aashto')} → {a}")
            e["sucs"], e["aashto"] = s, a
        self.refrescar_tabla()
        messagebox.showinfo("Clasificación automática (NTP 339.134 / 339.135)",
                            "\n".join(cambios) if cambios else "Las clasificaciones coinciden con los datos de laboratorio.")

    def editar_calicata(self):
        i = self._fila_sel()
        if i is None:
            messagebox.showinfo("Calicata", "Seleccione un estrato de la calicata a editar.")
            return
        cid = self.estratos[i]["cal"]

        def ok(cid, d):
            self.cal_info.setdefault(cid, {}).update(d)
            self.refrescar_tabla()
        DialogoCalicata(self, cid, self.cal_info.get(cid, {}), ok)

    def vincular_fotos(self):
        i = self._fila_sel()
        if i is None:
            messagebox.showinfo("Fotos", "Seleccione un estrato de la calicata a la que desea vincular fotos.")
            return
        cid = self.estratos[i]["cal"]
        rutas = filedialog.askopenfilenames(title=f"Fotografías de la calicata {cid}",
                                            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Todos", "*.*")])
        if not rutas:
            return
        fotos = self.cal_info.setdefault(cid, {}).setdefault("fotos", [])
        for r in rutas:
            fotos.append(dict(path=r, desc=f"Calicata {cid}: {os.path.splitext(os.path.basename(r))[0]}."))
        self.refrescar_tabla()
        self.estado(f"{len(rutas)} foto(s) vinculadas a {cid}.")

    def importar_excel(self):
        p = filedialog.askopenfilename(title="Excel de calicatas", filetypes=[("Excel", "*.xlsx *.xls")])
        if not p:
            return
        try:
            est, info = importar_excel(p)
        except Exception as ex:
            messagebox.showerror("Importar Excel", f"No se pudo importar:\n{ex}")
            return
        self.estratos = est
        for cid, d in info.items():
            fotos = self.cal_info.get(cid, {}).get("fotos", [])
            d["fotos"] = fotos
            self.cal_info[cid] = d
        self.refrescar_tabla()
        self.estado(f"Importados {len(est)} estratos de {len(info)} calicatas desde {os.path.basename(p)}.")

    def guardar_plantilla(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Plantilla_Calicatas_EMS.xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            crear_plantilla_excel(p)
            abrir_archivo(p)

    def cargar_ejemplo(self):
        for k, val in proyecto_ejemplo().items():
            if k in self.v:
                self.v[k].set(val)
        self.estratos = estratos_ejemplo()
        self.cal_info = info_ejemplo()
        try:
            generar_fotos_ejemplo(os.path.join(self.carpeta_salida, "fotos_ejemplo"), self.estratos, self.cal_info)
        except Exception as ex:
            self.estado(f"No se pudieron crear las fotos de ejemplo: {ex}")
        self.refrescar_tabla()
        self.recalcular()
        self.estado("Estudio de ejemplo cargado: 3 calicatas (C-1, C-2, C-3) de 3.00 m, estratos RELL/SM/GP/CL y 6 fotos.")

    # ------------------------------------------------------------------ TAB 3
    def _construir_tab3(self):
        izq = ctk.CTkScrollableFrame(self.t3, width=330)
        izq.pack(side="left", fill="y", padx=(0, 6))
        der = ctk.CTkFrame(self.t3)
        der.pack(side="left", fill="both", expand=True)
        self.v3 = {}

        def fila(lab, clave, valor, valores=None, seg=False):
            ctk.CTkLabel(izq, text=lab, anchor="w").pack(fill="x", padx=8, pady=(6, 0))
            var = tk.StringVar(value=valor)
            if seg:
                w = ctk.CTkSegmentedButton(izq, values=valores, variable=var)
            elif valores:
                w = ctk.CTkOptionMenu(izq, values=valores, variable=var)
            else:
                w = ctk.CTkEntry(izq, textvariable=var)
            w.pack(fill="x", padx=8, pady=2)
            var.trace_add("write", lambda *a: self._programar_calculo())
            self.v3[clave] = var
            return w

        ctk.CTkLabel(izq, text="Parámetros de cálculo", font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(4, 2))
        fila("Tipo de falla", "falla", "General", ["General", "Local (Punzonamiento)"], seg=True)
        fila("Tipo de cimentación", "tipo", "Cuadrada", ["Cuadrada", "Rectangular", "Continua"], seg=True)
        fila("Relación L/B (zapata rectangular)", "relLB", "1.50")
        fila("Teoría de diseño (gobernante)", "metodo", "Terzaghi", METODOS, seg=True)
        fila("Factor de seguridad FS (E.050 estático = 3.0)", "FS", "3.0")
        self.om_cal = fila("Calicata analizada", "calicata", "Envolvente", ["Envolvente"])
        fila("Módulo Es y Poisson μ", "es_modo", "Automático (SUCS)", ["Automático (SUCS)", "Manual"], seg=True)
        fila("Es manual (kg/cm²)", "Es", "150")
        fila("μ manual", "mu", "0.30")
        fila("Asentamiento máximo tolerable (cm)", "smax", f"{SMAX_E050:.2f}")
        fila("Luz entre columnas para distorsión (m)", "luz", "5.00")
        fila("Factor diferencial δ/Se (0.50 - 0.75)", "kdif", "0.75")
        fila("Df de diseño (m)", "Df", "1.50")
        fila("B de diseño (m)", "B", "1.50")
        self.chk_manual = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(izq, text="Usar φ, c, γ manuales (ignorar estratos)", variable=self.chk_manual,
                        command=self._programar_calculo).pack(fill="x", padx=8, pady=(10, 2))
        fila("φ manual (°)", "phi", "30")
        fila("c manual (kg/cm²)", "c", "0.00")
        fila("γ manual (g/cm³)", "gamma", "1.70")
        ctk.CTkButton(izq, text="↻ Recalcular ahora", command=self.recalcular).pack(fill="x", padx=8, pady=12)

        self.lbl_dis = ctk.CTkLabel(der, text="Cargue datos de calicatas para calcular.",
                                    font=ctk.CTkFont(size=16, weight="bold"), text_color=("#c00000", "#ff8c69"),
                                    anchor="w", justify="left", wraplength=860)
        self.lbl_dis.pack(fill="x", padx=8, pady=(6, 2))
        trf = ctk.CTkFrame(der)
        trf.pack(fill="both", expand=True, padx=4, pady=4)
        cols = [c for c, _, _ in COLS_CALC]
        self.tree_calc = ttk.Treeview(trf, columns=cols, show="headings", style="EMS.Treeview", height=12)
        for c, t, w in COLS_CALC:
            self.tree_calc.heading(c, text=t)
            self.tree_calc.column(c, width=w, anchor="center", stretch=False)
        ys = ttk.Scrollbar(trf, orient="vertical", command=self.tree_calc.yview)
        xs = ttk.Scrollbar(trf, orient="horizontal", command=self.tree_calc.xview)
        self.tree_calc.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree_calc.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        trf.grid_rowconfigure(0, weight=1)
        trf.grid_columnconfigure(0, weight=1)
        self.txt_det = ctk.CTkTextbox(der, height=190, font=ctk.CTkFont(family="Consolas", size=12))
        self.txt_det.pack(fill="x", padx=4, pady=(2, 4))
        self._estilo_tablas()

    def get_cfg(self):
        g = lambda k, d: num(self.v3[k].get(), d)
        return dict(metodo=self.v3["metodo"].get(),
                    falla="Local" if self.v3["falla"].get().startswith("Local") else "General",
                    tipo=self.v3["tipo"].get(), relLB=max(1.0, g("relLB", 1.5)), FS=max(1.0, g("FS", 3.0)),
                    es_auto=self.v3["es_modo"].get().startswith("Auto"), Es=max(1.0, g("Es", 150)),
                    mu=min(0.49, max(0.0, g("mu", 0.30))), smax=max(0.1, g("smax", SMAX_E050)),
                    luz=max(0.5, g("luz", 5.0)), kdif=min(1.0, max(0.25, g("kdif", 0.75))), calicata=self.v3["calicata"].get(), manual=self.chk_manual.get(),
                    phi=g("phi", 30.0), c=g("c", 0.0), gamma=g("gamma", 1.7),
                    Df=max(0.3, g("Df", 1.5)), B=max(0.3, g("B", 1.5)))

    def _programar_calculo(self):
        if not hasattr(self, "tree_calc"):
            return
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(300, self.recalcular)

    def recalcular(self):
        self._job = None
        if not self.estratos:
            return
        try:
            self.calc = calcular_todo(self.get_proyecto(), self.cal_info, self.estratos, self.get_cfg())
        except Exception as ex:
            self.lbl_dis.configure(text=f"Error de cálculo: {ex}")
            return
        c = self.calc
        cfg = c["cfg"]
        self.tree_calc.delete(*self.tree_calc.get_children())
        for i, f in enumerate(c["filas"]):
            x = f["res"][cfg["metodo"]]
            ok = f["ok_se"] and f["ok_alfa"] and not f["es_relleno"]
            tags = ["par" if i % 2 else "impar"]
            if not ok:
                tags.append("mal")
            self.tree_calc.insert("", "end", values=[
                f"{f['Df']:.2f}", f"{f['B']:.2f}", f["tipo"], f["cal"], f["sucs"], f"{x['phi_u']:.1f}", f"{x['c_u']:.2f}",
                f"{x['Nc']:.2f}", f"{x['Nq']:.2f}", f"{x['Ng']:.2f}", f"{f['qu']:.2f}", f"{f['qadm_corte']:.2f}",
                f"{f['se_corte']:.2f}", f"{min(f['q_asent'], f['q_dist']):.2f}", f"{f['qadm']:.2f}", f"{f['se']:.2f}",
                "CUMPLE" if ok else "REVISAR"], tags=tags)
        d = c["diseno"]
        self.lbl_dis.configure(text=f"qadm DISEÑO = {d['qadm']:.2f} kg/cm²   ·   Df = {d['Df']:.2f} m   ·   "
                                    f"Zapata {d['tipo'].lower()} B = {d['B']:.2f} m   ·   ΔH = {d['se']:.2f} cm   ·   "
                                    f"Gobierna: {d['gobierna']}   ·   Calicata crítica {d['cal']} ({d['sucs']})")
        x = d["res"][cfg["metodo"]]
        det = [f"DESARROLLO ({cfg['metodo']}, corte {cfg['falla'].lower()}) - {x['caso_nf']}",
               f"φ = {x['phi']:.2f}° → φ usado = {x['phi_u']:.2f}° ; c usado = {x['c_u']:.3f} kg/cm² ; "
               f"γ base = {d['g_base']:.3f} t/m³ ; γ sobre Df = {d['g_sobre']:.3f} t/m³",
               f"Nc = {x['Nc']:.3f}  Nq = {x['Nq']:.3f}  Nγ = {x['Ng']:.3f}  |  Sc = {x['sc']:.3f}  Sq = {x['sq']:.3f}  "
               f"Sγ = {x['sg']:.3f}  |  dc = {x['dc']:.3f}  dq = {x['dq']:.3f}  dγ = {x['dg']:.3f}",
               f"qu = c·Nc·Sc·dc + q·Nq·Sq·dq + 0.5·γ·B·Nγ·Sγ·dγ = {x['t1']:.2f} + {x['t2']:.2f} + {x['t3']:.2f} = "
               f"{x['qu_t']:.2f} t/m² = {x['qu']:.3f} kg/cm²",
               f"qadm corte = qu/FS = {x['qu']:.3f}/{cfg['FS']:.1f} = {x['qadm']:.3f} kg/cm²",
               f"Se = q·B·(1-μ²)·If/Es = {x['qadm_t']:.2f}·{d['B']:.2f}·(1-{d['mu']:.2f}²)·{d['If']:.2f}/{d['Es'] * 10:.0f} "
               f"= {d['se_corte']:.2f} cm   (Es = {d['Es']:.0f} kg/cm²)",
               f"q por asentamiento (S = {cfg['smax']:.2f} cm) = {d['q_asent']:.3f} ; q por distorsión (Se ≤ "
               f"{d['se_lim_dist']:.2f} cm) = {d['q_dist']:.3f}  →  qadm = {d['qadm']:.3f} kg/cm²",
               f"Distorsión: δ = {cfg['kdif']:.2f}·Se = {d['delta']:.2f} cm ; α = δ/L = 1/{(1 / d['alfa']) if d['alfa'] > 0 else 0:.0f} "
               f"({'CUMPLE' if d['ok_alfa'] else 'NO CUMPLE'} 1/500)",
               f"Terzaghi: {d['res']['Terzaghi']['qadm']:.2f}  |  Meyerhof: {d['res']['Meyerhof']['qadm']:.2f}  |  "
               f"Vesic: {d['res']['Vesic']['qadm']:.2f} kg/cm² (qadm por corte)"]
        if d["es_relleno"]:
            det.append("⚠ El estrato al nivel de Df es relleno/orgánico: profundice la cimentación.")
        self.txt_det.delete("1.0", "end")
        self.txt_det.insert("1.0", "\n".join(det))

    # ------------------------------------------------------------------ barra inferior
    def _construir_barra(self):
        bar = ctk.CTkFrame(self, corner_radius=0)
        bar.pack(fill="x", side="bottom")
        fb = ctk.CTkFrame(bar, fg_color="transparent")
        fb.pack(fill="x", padx=8, pady=(6, 2))
        fb.grid_columnconfigure((0, 1, 2), weight=1, uniform="b")
        ctk.CTkButton(fb, text="1. Generar Informe Completo de Estudio de Mecánica\nde Suelos (EMS - E.050) en Word (.DOCX)",
                      command=lambda: self.exportar("docx"), height=52, fg_color="#1f6aa5", hover_color="#144870",
                      font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, sticky="ew", padx=4)
        ctk.CTkButton(fb, text="2. Dibujar Lámina de Perfiles Estratigráficos de\nCalicatas (SUCS) en AutoCAD (.DXF)",
                      command=lambda: self.exportar("dxf"), height=52, fg_color="#d9822b", hover_color="#a8611c",
                      font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=1, sticky="ew", padx=4)
        ctk.CTkButton(fb, text="3. Exportar Hojas de Laboratorio, Granulometría y Matrices\nde Capacidad Portante en Excel (.XLSX)",
                      command=lambda: self.exportar("xlsx"), height=52, fg_color="#2e8b57", hover_color="#1f5f3b",
                      font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=2, sticky="ew", padx=4)
        self.lbl_estado = ctk.CTkLabel(bar, text="Listo.", anchor="w")
        self.lbl_estado.pack(fill="x", padx=12, pady=(0, 4))

    def estado(self, txt):
        self.lbl_estado.configure(text=txt)
        self.update_idletasks()

    def exportar(self, tipo):
        if not self.estratos:
            messagebox.showwarning("Sin datos", "Primero importe las calicatas o cargue el estudio de ejemplo.")
            return
        self.recalcular()
        if not self.calc:
            return
        os.makedirs(self.carpeta_salida, exist_ok=True)
        nombres = {"docx": ("Informe_EMS_E050.docx", "Word", "*.docx"),
                   "dxf": ("Lamina_EMS-01_Perfiles.dxf", "AutoCAD DXF", "*.dxf"),
                   "xlsx": ("Hojas_Calculo_EMS.xlsx", "Excel", "*.xlsx")}
        nom, desc, pat = nombres[tipo]
        p = filedialog.asksaveasfilename(initialdir=self.carpeta_salida, initialfile=nom, defaultextension="." + tipo,
                                         filetypes=[(desc, pat)])
        if not p:
            return
        self.estado(f"Generando {os.path.basename(p)}…")
        try:
            pr = self.get_proyecto()
            if tipo == "docx":
                exportar_word(p, pr, self.cal_info, self.estratos, self.calc)
            elif tipo == "dxf":
                exportar_dxf(p, pr, self.cal_info, self.estratos, self.calc)
            else:
                exportar_excel(p, pr, self.cal_info, self.estratos, self.calc)
        except PermissionError:
            messagebox.showerror("Archivo en uso", "Cierre el archivo si está abierto en otro programa e intente de nuevo.")
            self.estado("Error: archivo en uso.")
            return
        except Exception as ex:
            messagebox.showerror("Error al exportar", str(ex))
            self.estado(f"Error: {ex}")
            return
        self.estado(f"Archivo generado: {p}")
        abrir_archivo(p)


# =============================================================================
# 16. PUNTO DE ENTRADA
# =============================================================================
def main():
    if len(sys.argv) >= 2 and sys.argv[1] == "--demo":
        carpeta = sys.argv[2] if len(sys.argv) >= 3 else os.path.join(os.getcwd(), "EMS_demo")
        modo_demo(carpeta)
        return
    if not GUI_OK:
        print("No se pudo cargar customtkinter/tkinter. Instale: pip install customtkinter")
        print("Puede generar el ejemplo sin interfaz con:  python app_ems_suelos_e050_pro.py --demo carpeta")
        sys.exit(1)
    app = AppEMS()
    app.mainloop()


if __name__ == "__main__":
    main()
