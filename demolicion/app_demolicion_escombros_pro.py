# -*- coding: utf-8 -*-
"""
===============================================================================
 app_demolicion_escombros_pro.py
 Generador Pro de Expedientes de Demolición, Cálculo de Esponjamiento,
 Flota de Volquetes y Planos en AutoCAD (.DXF)
-------------------------------------------------------------------------------
 Aplicación de escritorio 100 % offline (CustomTkinter) para:
   * Metrado de demolición según el Reglamento Nacional de Metrados
     (RD N° 073-2010-VIVIENDA).
   * Cálculo de volumen en banco, volumen esponjado, peso, viajes de volquete,
     ciclo de transporte a escombrera y flota necesaria.
   * Presupuesto de demolición y eliminación (S/) con horas-máquina.
   * Memoria Descriptiva + Plan de Seguridad y Salud en el Trabajo (Norma
     G.050 del RNE) + Plan de Mitigación de Polvo y Ruido en Word (.docx).
   * Lámina D-01 de demolición en AutoCAD (.dxf) con ezdxf.
   * Planilla de metrados, esponjamiento y presupuesto en Excel (.xlsx).
 Marco normativo: Ley N° 29090, D.S. N° 029-2019-VIVIENDA, Norma G.050 (RNE),
 D.S. N° 011-2019-TR, Ley N° 29783, D.S. N° 003-2013-VIVIENDA (RCD) y su
 modificatoria D.S. N° 019-2016-VIVIENDA, RD N° 073-2010-VIVIENDA.
===============================================================================
"""

import io
import math
import os
import re
import subprocess
import sys
import datetime
import unicodedata

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import customtkinter as ctk
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import ezdxf
from ezdxf.enums import TextEntityAlignment
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont, ImageOps

APP_TITULO = ("Generador Pro de Expedientes de Demolición, Cálculo de Esponjamiento, "
              "Flota de Volquetes y Planos en AutoCAD (.DXF)")
APP_VERSION = "1.0.0"

# =============================================================================
# 1. CATÁLOGO TÉCNICO DE PARTIDAS DE DEMOLICIÓN Y DESMONTAJE (CAPECO / RNM)
# -----------------------------------------------------------------------------
#  modo:  "vol"    -> Metrado (m³)  = N° x Largo x Ancho x Alto
#         "muro"   -> Metrado (m²)  = N° x Largo x Alto ; Vb = m² x espesor
#                     (Ancho/Espesor de la tabla; si es 0 se usa "esp")
#         "planta" -> Metrado (m²)  = N° x Largo x Ancho ; Vb = m² x esp
#                     (si usa_h=True y Alto>0, Alto es el espesor real)
#         "und"    -> Metrado (und) = N° veces ; Vb = und x esp (m³/und)
#  fe   : Factor de Esponjamiento        peso : Peso específico suelto (t/m³)
#  esp  : Espesor equivalente (m) o factor (m³/m²) o volumen por unidad (m³/und)
#  cu   : Costo unitario de demolición (S/) para método Mixto (base)
#  rman : Rendimiento manual por cuadrilla (und/día)
#  rmec : Rendimiento mecánico por equipo (und/día)
#  rhid : Rendimiento de retroexcavadora con martillo hidráulico (m³ banco/h)
#  fase : Fase de la secuencia descendente G.050 (1 a 6)
# =============================================================================
CATALOGO = [
    dict(cod="DM-01", grupo="Desmontajes",
         nombre="Desmontaje de cobertura liviana (calamina / aluzinc / eternit) sobre tijerales",
         und="m²", modo="planta", usa_h=False, fe=1.15, peso=0.30, esp=0.03, cu=9.80,
         equipo="Herramientas manuales, andamio y línea de vida", rman=60, rmec=60, rhid=0, fase=1),
    dict(cod="DM-02", grupo="Desmontajes",
         nombre="Desmontaje de puertas (hojas, marcos y accesorios) - acopio recuperable",
         und="und", modo="und", usa_h=False, fe=1.15, peso=0.45, esp=0.12, cu=25.00,
         equipo="Herramientas manuales", rman=16, rmec=16, rhid=0, fase=1),
    dict(cod="DM-03", grupo="Desmontajes",
         nombre="Desmontaje de ventanas (vidrios, marcos y rejas) - acopio recuperable",
         und="und", modo="und", usa_h=False, fe=1.15, peso=0.50, esp=0.08, cu=28.00,
         equipo="Herramientas manuales", rman=14, rmec=14, rhid=0, fase=1),
    dict(cod="DM-04", grupo="Desmontajes",
         nombre="Desmontaje de aparatos sanitarios y griferías - acopio recuperable",
         und="und", modo="und", usa_h=False, fe=1.15, peso=0.60, esp=0.10, cu=22.00,
         equipo="Herramientas manuales", rman=12, rmec=12, rhid=0, fase=1),
    dict(cod="DM-05", grupo="Losas aligeradas",
         nombre="Demolición de losa aligerada h=0.17 m (0.12 m³/m² descontando ladrillo)",
         und="m²", modo="planta", usa_h=False, fe=1.35, peso=1.50, esp=0.12, cu=36.50,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=8, rmec=45, rhid=5.0, fase=2),
    dict(cod="DM-06", grupo="Losas aligeradas",
         nombre="Demolición de losa aligerada h=0.20 m (0.14 m³/m² descontando ladrillo)",
         und="m²", modo="planta", usa_h=False, fe=1.35, peso=1.50, esp=0.14, cu=41.00,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=7, rmec=40, rhid=5.0, fase=2),
    dict(cod="DM-07", grupo="Losas aligeradas",
         nombre="Demolición de losa aligerada h=0.25 m (0.16 m³/m² descontando ladrillo)",
         und="m²", modo="planta", usa_h=False, fe=1.35, peso=1.50, esp=0.16, cu=46.50,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=6, rmec=35, rhid=5.0, fase=2),
    dict(cod="DM-08", grupo="Concreto armado",
         nombre="Demolición de vigas de concreto armado",
         und="m³", modo="vol", usa_h=False, fe=1.40, peso=1.60, esp=1.00, cu=245.00,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=1.2, rmec=12, rhid=3.0, fase=3),
    dict(cod="DM-09", grupo="Albañilería",
         nombre="Demolición de tabiques de ladrillo pandereta e=0.12 m (no portantes)",
         und="m²", modo="muro", usa_h=False, fe=1.30, peso=1.40, esp=0.12, cu=14.50,
         equipo="Martillo demoledor eléctrico / comba y cincel", rman=18, rmec=70, rhid=7.0, fase=4),
    dict(cod="DM-10", grupo="Albañilería",
         nombre="Demolición de muros de ladrillo KK de soga e=0.15 m",
         und="m²", modo="muro", usa_h=False, fe=1.30, peso=1.40, esp=0.15, cu=18.50,
         equipo="Martillo demoledor eléctrico / Retroexcavadora", rman=14, rmec=60, rhid=6.0, fase=4),
    dict(cod="DM-11", grupo="Albañilería",
         nombre="Demolición de muros de ladrillo KK de cabeza e=0.25 m (portantes)",
         und="m²", modo="muro", usa_h=False, fe=1.30, peso=1.40, esp=0.25, cu=27.00,
         equipo="Martillo demoledor eléctrico / Retroexcavadora", rman=9, rmec=40, rhid=6.0, fase=5),
    dict(cod="DM-12", grupo="Concreto armado",
         nombre="Demolición de columnas de concreto armado",
         und="m³", modo="vol", usa_h=False, fe=1.40, peso=1.60, esp=1.00, cu=235.00,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=1.3, rmec=14, rhid=3.0, fase=5),
    dict(cod="DM-13", grupo="Concreto armado",
         nombre="Demolición de placas / muros de concreto armado",
         und="m³", modo="vol", usa_h=False, fe=1.40, peso=1.60, esp=1.00, cu=225.00,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=1.4, rmec=15, rhid=3.2, fase=5),
    dict(cod="DM-14", grupo="Concreto armado",
         nombre="Demolición de escaleras de concreto armado",
         und="m³", modo="vol", usa_h=False, fe=1.40, peso=1.60, esp=1.00, cu=250.00,
         equipo="Martillo demoledor eléctrico", rman=1.2, rmec=10, rhid=3.0, fase=5),
    dict(cod="DM-15", grupo="Adobe / Tapial",
         nombre="Demolición de muros de adobe o tapial e=0.40 m (por m²)",
         und="m²", modo="muro", usa_h=False, fe=1.25, peso=1.30, esp=0.40, cu=22.00,
         equipo="Herramientas manuales / Retroexcavadora", rman=10, rmec=50, rhid=9.0, fase=5),
    dict(cod="DM-16", grupo="Adobe / Tapial",
         nombre="Demolición de muros de adobe o tapial (por m³)",
         und="m³", modo="vol", usa_h=False, fe=1.25, peso=1.30, esp=1.00, cu=55.00,
         equipo="Herramientas manuales / Retroexcavadora", rman=4, rmec=30, rhid=9.0, fase=5),
    dict(cod="DM-17", grupo="Pisos y veredas",
         nombre="Demolición de pisos de concreto, falsos pisos y contrapisos e=0.10 m",
         und="m²", modo="planta", usa_h=True, fe=1.30, peso=1.55, esp=0.10, cu=16.50,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=15, rmec=80, rhid=6.0, fase=6),
    dict(cod="DM-18", grupo="Pisos y veredas",
         nombre="Demolición de veredas y pisos exteriores de concreto e=0.15 m",
         und="m²", modo="planta", usa_h=True, fe=1.30, peso=1.55, esp=0.15, cu=21.00,
         equipo="Martillo demoledor / Retroexcavadora con picotón", rman=12, rmec=70, rhid=6.0, fase=6),
    dict(cod="DM-19", grupo="Cimentaciones",
         nombre="Demolición de cimientos corridos de concreto ciclópeo",
         und="m³", modo="vol", usa_h=False, fe=1.35, peso=1.55, esp=1.00, cu=145.00,
         equipo="Retroexcavadora con picotón / Martillo demoledor", rman=1.8, rmec=18, rhid=4.0, fase=6),
    dict(cod="DM-20", grupo="Cimentaciones",
         nombre="Demolición de sobrecimientos de concreto ciclópeo",
         und="m³", modo="vol", usa_h=False, fe=1.35, peso=1.55, esp=1.00, cu=150.00,
         equipo="Martillo demoledor / Retroexcavadora", rman=2.0, rmec=18, rhid=4.0, fase=6),
    dict(cod="DM-21", grupo="Cimentaciones",
         nombre="Demolición de zapatas y vigas de cimentación de concreto armado",
         und="m³", modo="vol", usa_h=False, fe=1.40, peso=1.60, esp=1.00, cu=260.00,
         equipo="Retroexcavadora con picotón (martillo hidráulico)", rman=1.0, rmec=12, rhid=2.8, fase=6),
]
CAT = {c["cod"]: c for c in CATALOGO}
CAT_ETIQUETAS = [f'{c["cod"]} | {c["nombre"]} [{c["und"]}]' for c in CATALOGO]

FASES = {
    1: ("Desmontaje previo", "Retiro de coberturas livianas, puertas, ventanas, aparatos sanitarios, "
                             "instalaciones y elementos recuperables."),
    2: ("Techos y losas", "Demolición de techos y losas de arriba hacia abajo, empezando por la azotea."),
    3: ("Vigas", "Demolición de vigas una vez liberadas de la carga de la losa."),
    4: ("Muros no portantes", "Demolición de tabiques y muros de soga sin función portante."),
    5: ("Muros portantes y columnas", "Demolición de muros portantes, placas, escaleras y columnas por niveles."),
    6: ("Cimentaciones y pisos", "Demolición de pisos, falsos pisos, sobrecimientos, cimientos y zapatas."),
}

TRAMITES = [
    "Licencia de Edificación - Modalidad Demolición Total",
    "Demolición Parcial por Remodelación",
    "Expediente Técnico de Obra Pública",
]
SISTEMAS = [
    "Albañilería confinada con losas aligeradas",
    "Pórticos de concreto armado con losas aligeradas",
    "Sistema dual / muros de concreto armado",
    "Adobe / Tapial (casona o local antiguo)",
    "Estructura metálica con cobertura liviana",
]
METODOS = [
    "Mixto: Manual en colindancias + Mecánico con Maquinaria",
    "100% Manual con Martillos Eléctricos/Neumáticos",
    "100% Mecánico",
]
VOLQUETES = ["15 m³", "10 m³", "6 m³"]
TARIFA_VOLQUETE = {15: 245.00, 10: 195.00, 6: 150.00}   # S/ por hora-máquina
CATEGORIAS_FOTO = ["Fachada principal", "Interiores", "Colindancia izquierda",
                   "Colindancia derecha", "Colindancia fondo", "Azotea / Techos", "Otros"]

CHECKLIST_G050 = [
    "Licencia / autorización municipal de demolición vigente y en obra",
    "Corte y bloqueo de energía eléctrica verificado por el concesionario",
    "Cierre de agua potable y sellado de desagüe contra roedores",
    "Purgado y retiro de la red de gas natural / GLP",
    "Acta de constatación del estado de las edificaciones colindantes",
    "Cerco provisional H=2.40 m y bandeja peatonal instalados",
    "Malla raschel antipolvo H=4.00 m en colindancias",
    "Apuntalamiento preventivo de muros colindantes instalado",
    "Señalización, cinta de peligro y vigía de tránsito",
    "Líneas de vida y puntos de anclaje certificados",
    "Personal con EPP, SCTR vigente y charla de 5 minutos",
    "Sistema de riego con agua nebulizada operativo (PM10)",
    "Botiquín, extintor PQS y plan de respuesta ante emergencias",
    "Escombrera autorizada y manifiesto de residuos RCD",
]


# =============================================================================
# 2. UTILIDADES GENERALES
# =============================================================================
def quitar_acentos(txt):
    txt = str(txt)
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")


def fnum(valor, defecto=0.0):
    """Convierte texto (con coma decimal o unidades) a float."""
    if valor is None:
        return defecto
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and math.isnan(valor):
            return defecto
        return float(valor)
    s = str(valor).strip().replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    return float(m.group(0)) if m else defecto


def soles(x):
    return f"S/ {x:,.2f}"


def fmt(x, d=2):
    return f"{x:,.{d}f}"


def cortar(txt, n):
    txt = str(txt)
    return txt if len(txt) <= n else txt[: n - 1] + "…"


def capacidad_volquete(texto):
    v = int(fnum(texto, 15))
    return v if v in TARIFA_VOLQUETE else 15


def abrir_archivo(ruta):
    """Abre el archivo generado con la aplicación predeterminada del sistema."""
    try:
        if hasattr(os, "startfile"):
            os.startfile(ruta)                       # Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta])
    except Exception as exc:                        # pragma: no cover
        print("No se pudo abrir el archivo:", exc)


def nombre_seguro(txt, maximo=40):
    s = quitar_acentos(txt)
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s[:maximo] or "Proyecto"


# =============================================================================
# 3. MOTOR DE CÁLCULO DE DEMOLICIÓN, ESPONJAMIENTO Y MAQUINARIA
# =============================================================================
def config_metodo(metodo):
    m = quitar_acentos(metodo).lower()
    if "manual" in m and "100" in m:
        return dict(clave="MANUAL", fm=1.00, fcu=1.20,
                    cargador="Minicargador frontal 70 HP (0.5 m³)", rend_carg=30.0, tarifa_carg=150.0)
    if "mecanico" in m and "100" in m:
        return dict(clave="MECANICO", fm=0.00, fcu=0.85,
                    cargador="Cargador frontal 125 HP (2.5 yd³)", rend_carg=70.0, tarifa_carg=260.0)
    return dict(clave="MIXTO", fm=0.30, fcu=1.00,
                cargador="Cargador frontal 125 HP (2.5 yd³)", rend_carg=70.0, tarifa_carg=260.0)


def nueva_fila(cod, piso="", desc="", n=1, largo=0.0, ancho=0.0, alto=0.0, metrado_directo=0.0):
    return dict(cod=cod, piso=piso, desc=desc, n=n, largo=largo, ancho=ancho, alto=alto,
                md=metrado_directo)


def calcular_fila(f, cap):
    """Calcula metrado, volumen en banco, volumen esponjado, peso y viajes de una fila."""
    c = CAT[f["cod"]]
    n, L, A, H = fnum(f["n"], 1), fnum(f["largo"]), fnum(f["ancho"]), fnum(f["alto"])
    md = fnum(f.get("md", 0))
    modo = c["modo"]
    if md > 0:
        metrado = md
    elif modo == "vol":
        metrado = n * L * A * H
    elif modo == "muro":
        metrado = n * L * H
    elif modo == "planta":
        metrado = n * L * A
    else:
        metrado = n
    if modo == "vol":
        vb = metrado
        esp = 1.0
    elif modo == "muro":
        esp = A if A > 0 else c["esp"]
        vb = metrado * esp
    elif modo == "planta":
        esp = H if (c["usa_h"] and H > 0) else c["esp"]
        vb = metrado * esp
    else:
        esp = c["esp"]
        vb = metrado * esp
    ve = vb * c["fe"]
    r = dict(f)
    r.update(und=c["und"], partida=c["nombre"], grupo=c["grupo"], fase=c["fase"],
             metrado=metrado, esp_eq=esp, vb=vb, fe=c["fe"], pct=(c["fe"] - 1) * 100,
             ve=ve, peso=ve * c["peso"], viajes=ve / cap if cap else 0)
    return r


def calcular_todo(datos, filas):
    """Calcula todo el expediente: metrados, flota, horas-máquina, presupuesto y cronograma."""
    cap = capacidad_volquete(datos.get("cap_volquete", "15 m³"))
    met = config_metodo(datos.get("metodo", METODOS[0]))
    fm = met["fm"]
    dist = max(fnum(datos.get("distancia"), 10.0), 0.1)
    jornada = max(fnum(datos.get("jornada"), 8.0), 1.0)
    cuadrillas = max(int(fnum(datos.get("cuadrillas"), 2)), 1)
    area_t = fnum(datos.get("area_terreno"), 0.0)
    calc = [calcular_fila(f, cap) for f in filas]

    vb = sum(r["vb"] for r in calc)
    ve = sum(r["ve"] for r in calc)
    peso = sum(r["peso"] for r in calc)
    viajes = math.ceil(ve / cap - 1e-9) if ve > 0 else 0

    # --- Ciclo del volquete ----------------------------------------------------
    t_carga = cap / met["rend_carg"] * 60.0
    t_ida = dist / 30.0 * 60.0
    t_desc = 4.0
    t_ret = dist / 45.0 * 60.0
    t_man = 4.0
    ciclo = t_carga + t_ida + t_desc + t_ret + t_man
    min_ef = jornada * 60.0 * 0.83
    viajes_vol_dia = max(1, int(min_ef // ciclo))
    n_opt = max(1, math.ceil(ciclo / (t_carga + t_man)))
    n_vol = max(1, min(n_opt, 6, math.ceil(viajes / viajes_vol_dia) if viajes else 1))
    dias_elim = math.ceil(viajes / (n_vol * viajes_vol_dia)) if viajes else 0
    flota = dict(cap=cap, dist=dist, t_carga=t_carga, t_ida=t_ida, t_desc=t_desc, t_ret=t_ret,
                 t_man=t_man, ciclo=ciclo, min_ef=min_ef, viajes_vol_dia=viajes_vol_dia,
                 n_opt=n_opt, n_vol=n_vol, dias_elim=dias_elim, viajes=viajes,
                 cargador=met["cargador"], rend_carg=met["rend_carg"],
                 costo_viaje=ciclo / 60.0 * TARIFA_VOLQUETE[cap], tarifa_vol=TARIFA_VOLQUETE[cap])

    # --- Duración por fases --------------------------------------------------------
    dias_fase = {k: 0.0 for k in FASES}
    for r in calc:
        c = CAT[r["cod"]]
        if c["fase"] == 1:
            d = r["metrado"] / c["rman"] / cuadrillas
        else:
            d = r["metrado"] * fm / c["rman"] / cuadrillas + r["metrado"] * (1 - fm) / c["rmec"]
        dias_fase[c["fase"]] += d
    dias_fase = {k: (math.ceil(v) if v > 0 else 0) for k, v in dias_fase.items()}
    dias_demol = sum(dias_fase.values())
    cron = []
    dia = 1
    cron.append(("Fase 0", "Obras provisionales, corte de suministros, cerco y apuntalamiento", dia, dia + 1))
    dia += 2
    inicio_elim = None
    for k in sorted(FASES):
        if dias_fase[k] <= 0:
            continue
        cron.append((f"Fase {k}", FASES[k][0], dia, dia + dias_fase[k] - 1))
        dia += dias_fase[k]
        if k == 1:
            inicio_elim = dia
    if inicio_elim is None:
        inicio_elim = 3
    fin_demol = dia - 1
    fin_elim = inicio_elim + max(dias_elim, 1) - 1 if viajes else inicio_elim
    fin_elim = max(fin_elim, fin_demol)
    cron.append(("Eliminación", f"Carguío y eliminación con {n_vol} volquete(s) de {cap} m³ "
                                f"(en paralelo a la demolición)", inicio_elim, fin_elim))
    cron.append(("Cierre", "Limpieza final, nivelación del terreno y entrega", fin_elim + 1, fin_elim + 1))
    plazo = fin_elim + 1

    # --- Horas-máquina --------------------------------------------------------------
    h_hid = sum(r["vb"] * (1 - fm) / CAT[r["cod"]]["rhid"] for r in calc if CAT[r["cod"]]["rhid"] > 0)
    h_mart = sum(r["vb"] * fm / (CAT[r["cod"]]["rhid"] * 0.12) for r in calc if CAT[r["cod"]]["rhid"] > 0)
    h_carg = ve / met["rend_carg"] * 1.20
    h_vol = viajes * ciclo / 60.0
    h_cist = dias_demol * 2.0
    hm = [
        dict(eq="Retroexcavadora 416 con martillo hidráulico (picotón)", h=h_hid, tarifa=280.0,
             nota="Incluido en C.U. de partidas 02"),
        dict(eq="Martillo demoledor eléctrico / neumático 30 kg", h=h_mart, tarifa=25.0,
             nota="Incluido en C.U. de partidas 02"),
        dict(eq=met["cargador"], h=h_carg, tarifa=met["tarifa_carg"], nota="Valorizado en partida 03.02"),
        dict(eq=f"Camión volquete {cap} m³", h=h_vol, tarifa=TARIFA_VOLQUETE[cap], nota="Valorizado en partida 03.03"),
        dict(eq="Camión cisterna 2,000 gal (riego nebulizado PM10)", h=h_cist, tarifa=165.0,
             nota="Valorizado en partida 04.03"),
    ]
    for e in hm:
        e["parcial"] = e["h"] * e["tarifa"]

    # --- Presupuesto -----------------------------------------------------------------
    lado = math.sqrt(area_t) if area_t > 0 else 10.0
    frente = round(lado, 2)
    perim_col = round(3 * lado, 2)
    personal = cuadrillas * 5 + 3
    pres = []
    s1 = [
        ("01.01", "Cerco perimétrico provisional H=2.40 m (postes de madera y triplay/calamina)", "m", frente + 2.0, 85.00),
        ("01.02", "Malla raschel antipolvo H=4.00 m en colindancias (postes F.G. 2\")", "m²", perim_col * 4.0, 12.50),
        ("01.03", "Cartel de obra 3.60 x 2.40 m", "und", 1, 950.00),
        ("01.04", "Corte, clausura y bloqueo de suministros (luz, agua, desagüe, gas)", "glb", 1, 850.00),
        ("01.05", "Apuntalamiento preventivo de muros colindantes (tornapuntas, soleras y muertos)", "m", perim_col, 68.00),
        ("01.06", "Bandeja protectora peatonal / pasadizo cubierto en vereda", "m", frente, 120.00),
    ]
    pres.append(("01", "OBRAS PROVISIONALES Y TRABAJOS PRELIMINARES", s1))
    s2 = []
    agrup = {}
    for r in calc:
        agrup.setdefault(r["cod"], 0.0)
        agrup[r["cod"]] += r["metrado"]
    for i, (cod, m) in enumerate(sorted(agrup.items(), key=lambda kv: (CAT[kv[0]]["fase"], kv[0])), 1):
        c = CAT[cod]
        fcu = 1.0 if c["fase"] == 1 else met["fcu"]
        s2.append((f"02.{i:02d}", c["nombre"], c["und"], m, round(c["cu"] * fcu, 2)))
    pres.append(("02", "DEMOLICIONES Y DESMONTAJES", s2))
    cu_carg = (h_carg * met["tarifa_carg"] / ve) if ve > 0 else 0
    s3 = [
        ("03.01", "Acarreo interno de escombros con carretilla/buggy D<=30 m", "m³", ve, 22.00),
        ("03.02", f"Carguío de desmonte con {met['cargador']}", "m³", ve, round(cu_carg, 2)),
        ("03.03", f"Eliminación con volquete {cap} m³ a escombrera autorizada D={dist:.1f} km", "viaje",
         viajes, round(flota["costo_viaje"], 2)),
        ("03.04", "Derecho de disposición final en escombrera autorizada (D.S. 003-2013-VIVIENDA)", "m³", ve, 6.50),
    ]
    pres.append(("03", "ACARREO, CARGUÍO Y ELIMINACIÓN DE RCD", s3))
    s4 = [
        ("04.01", "Equipos de protección personal (EPP) G.050", "trab", personal, 480.00),
        ("04.02", "Señalización temporal de seguridad, cintas y conos", "glb", 1, 1200.00),
        ("04.03", "Riego con agua nebulizada / cisterna para control de PM10", "h", h_cist, 165.00),
        ("04.04", "Elaboración, implementación y administración del Plan SST", "glb", 1, 1500.00),
        ("04.05", "Líneas de vida, anclajes y arneses certificados", "glb", 1, 1350.00),
        ("04.06", "Monitoreo de ruido y material particulado PM10", "und", 2, 650.00),
    ]
    pres.append(("04", "SEGURIDAD Y SALUD EN EL TRABAJO (NORMA G.050) Y MITIGACIÓN", s4))
    cd = 0.0
    secciones = []
    for cod, tit, items in pres:
        its = [dict(item=i, desc=d, und=u, metrado=round(m, 2), cu=cu, parcial=round(round(m, 2) * cu, 2))
               for i, d, u, m, cu in items]
        sub = sum(x["parcial"] for x in its)
        cd += sub
        secciones.append(dict(cod=cod, titulo=tit, items=its, subtotal=sub))
    gg = cd * 0.10
    ut = cd * 0.08
    sub_t = cd + gg + ut
    igv = sub_t * 0.18
    total = sub_t + igv
    presupuesto = dict(secciones=secciones, cd=cd, gg=gg, ut=ut, subtotal=sub_t, igv=igv, total=total)

    grupos = {}
    for r in calc:
        g = grupos.setdefault(r["grupo"], dict(vb=0.0, ve=0.0, peso=0.0))
        g["vb"] += r["vb"]
        g["ve"] += r["ve"]
        g["peso"] += r["peso"]

    return dict(filas=calc, vb=vb, ve=ve, peso=peso, viajes=viajes, cap=cap, metodo=met,
                flota=flota, dias_fase=dias_fase, dias_demol=dias_demol, cron=cron, plazo=plazo,
                hm=hm, presupuesto=presupuesto, grupos=grupos, personal=personal,
                frente=frente, perim_col=perim_col, agrupado=agrup)


# =============================================================================
# 4. CASO DE EJEMPLO E IMPORTACIÓN DESDE EXCEL / REVIT
# =============================================================================
def caso_ejemplo():
    """Demolición de edificación de 2 pisos + azotea (12 partidas)."""
    datos = dict(
        tramite=TRAMITES[0],
        proyecto="Demolición total de vivienda unifamiliar de 2 pisos + azotea",
        propietario="Inversiones Los Álamos S.A.C.",
        ubicacion="Av. Los Próceres N° 1250, Mz. B Lt. 14, Urb. Santa Rosa",
        municipalidad="Municipalidad Distrital de Santiago de Surco - Lima",
        n_pisos="2", area_terreno="160.00", area_techada="240.00",
        sistema=SISTEMAS[0], metodo=METODOS[0], cap_volquete="15 m³",
        distancia="18.5", escombrera="Escombrera autorizada - Relleno de RCD Lurín",
        ingeniero="Ing. Carlos Alberto Mendoza Ríos", cip="CIP N° 98765",
        jornada="8", cuadrillas="2",
        fecha=datetime.date.today().strftime("%d/%m/%Y"),
    )
    filas = [
        nueva_fila("DM-01", "Azotea", "Cobertura de calamina sobre tijerales de madera (lavandería)", 1, 6.00, 4.50, 0),
        nueva_fila("DM-02", "Piso 1 y 2", "Puertas contraplacadas y metálicas con marcos", 14),
        nueva_fila("DM-03", "Piso 1 y 2", "Ventanas de fierro y vidrio con rejas", 10),
        nueva_fila("DM-06", "Techo Piso 2", "Losa aligerada h=0.20 m techo 2° piso (azotea)", 1, 12.00, 10.00, 0.20),
        nueva_fila("DM-06", "Techo Piso 1", "Losa aligerada h=0.20 m techo 1° piso", 1, 12.00, 10.00, 0.20),
        nueva_fila("DM-08", "Piso 1 y 2", "Vigas V-101 / V-201 (0.25 x 0.40 bajo losa)", 2, 44.00, 0.25, 0.40),
        nueva_fila("DM-12", "Piso 1 y 2", "Columnas C-1 0.25 x 0.40 (12 por nivel)", 24, 0.25, 0.40, 2.60),
        nueva_fila("DM-14", "Piso 1 - 2", "Escalera de concreto armado de 2 tramos", 1, 7.00, 1.10, 0.20),
        nueva_fila("DM-10", "Piso 1 y 2", "Muros de soga e=0.15 ejes A-D interiores", 2, 38.00, 0.15, 2.40),
        nueva_fila("DM-11", "Piso 1 y 2", "Muros de cabeza e=0.25 ejes 1 y 5 (colindancias, picado manual)", 2, 20.00, 0.25, 2.40),
        nueva_fila("DM-17", "Piso 1", "Falsos pisos y pisos de concreto e=0.10 m", 1, 12.00, 10.00, 0.10),
        nueva_fila("DM-19", "Cimentación", "Cimientos corridos de concreto ciclópeo 0.50 x 0.80", 1, 64.00, 0.50, 0.80),
    ]
    return datos, filas


def inferir_codigo(texto):
    t = quitar_acentos(texto).lower()
    for c in CATALOGO:                                   # código explícito
        if c["cod"].lower() in t:
            return c["cod"]
    reglas = [
        (("cobertura", "calamina", "aluzinc", "eternit", "teja", "tijeral"), "DM-01"),
        (("puerta",), "DM-02"),
        (("ventana", "mampara", "reja"), "DM-03"),
        (("sanitari", "inodoro", "lavatorio", "griferia", "ducha"), "DM-04"),
        (("viga de cimentacion", "zapata", "cimentacion de c"), "DM-21"),
        (("sobrecimiento",), "DM-20"),
        (("cimiento", "ciclopeo"), "DM-19"),
        (("vereda", "exterior"), "DM-18"),
        (("falso piso", "contrapiso", "piso"), "DM-17"),
        (("escalera",), "DM-14"),
        (("columna",), "DM-12"),
        (("placa", "muro de concreto", "muro de c.a", "muro estructural"), "DM-13"),
        (("viga",), "DM-08"),
        (("adobe", "tapial"), "DM-15"),
        (("pandereta", "tabique"), "DM-09"),
        (("cabeza", "e=0.25", "e = 0.25"), "DM-11"),
        (("soga", "muro", "albanil", "ladrillo", "wall"), "DM-10"),
    ]
    if any(k in t for k in ("losa", "aligerad", "techo", "floor", "slab")):
        if "0.25" in t or "h=25" in t or " 25" in t:
            return "DM-07"
        if "0.17" in t or "h=17" in t or " 17" in t:
            return "DM-05"
        return "DM-06"
    for claves, cod in reglas:
        if any(k in t for k in claves):
            return cod
    return None


def _norm(h):
    return quitar_acentos(h).lower().strip()


def mapear_columnas(cols):
    mapa = {}
    for col in cols:
        h = _norm(col)
        if not h or h.startswith("unnamed"):
            continue
        if "piso" in h or "sector" in h or "nivel" in h or h == "level":
            mapa.setdefault("piso", col)
        elif "veces" in h or h.startswith("n°") or h.startswith("nro") or "cantidad" in h \
                or "recuento" in h or h == "count" or h.startswith("n ") or h == "n":
            mapa.setdefault("n", col)
        elif "codigo" in h or "partida" in h or h.startswith("tipo") or "familia" in h or "categoria" in h \
                or "elemento" in h:
            mapa.setdefault("cod", col)
        elif "descrip" in h or "comentario" in h or "nombre" in h or "marca" in h:
            mapa.setdefault("desc", col)
        elif "largo" in h or "longitud" in h or "length" in h:
            mapa.setdefault("largo", col)
        elif "ancho" in h or "espesor" in h or "width" in h or "grosor" in h or "thickness" in h:
            mapa.setdefault("ancho", col)
        elif "alto" in h or "altura" in h or "height" in h or "peralte" in h:
            mapa.setdefault("alto", col)
        elif ("volumen" in h and "banco" not in h and "esponj" not in h) or h.startswith("area") \
                or "metrado directo" in h or h == "metrado" or h == "volume":
            mapa.setdefault("md", col)
    return mapa


def importar_excel(ruta):
    """Importa una tabla de elementos a demoler (Excel propio o tabla de planificación de Revit)."""
    crudo = pd.read_excel(ruta, sheet_name=0, header=None)
    fila_enc = 0
    claves = ("piso", "nivel", "descrip", "largo", "ancho", "alto", "veces", "partida", "tipo",
              "volumen", "area", "longitud", "familia")
    for i in range(min(20, len(crudo))):
        valores = [_norm(v) for v in crudo.iloc[i].tolist() if isinstance(v, str)]
        if sum(any(k in v for k in claves) for v in valores) >= 2:
            fila_enc = i
            break
    df = pd.read_excel(ruta, sheet_name=0, header=fila_enc)
    mapa = mapear_columnas(df.columns)
    filas, omitidas = [], 0
    for _, row in df.iterrows():
        get = lambda k, d=None: row[mapa[k]] if k in mapa and not (isinstance(row[mapa[k]], float)
                                                                     and math.isnan(row[mapa[k]])) else d
        texto_cod = str(get("cod", "") or "")
        desc = str(get("desc", "") or "")
        if not texto_cod.strip() and not desc.strip():
            continue
        cod = texto_cod.strip() if texto_cod.strip() in CAT else inferir_codigo(texto_cod + " " + desc)
        if cod is None:
            omitidas += 1
            continue
        f = nueva_fila(cod, str(get("piso", "") or ""), desc or texto_cod,
                       fnum(get("n", 1), 1) or 1, fnum(get("largo", 0)), fnum(get("ancho", 0)),
                       fnum(get("alto", 0)), fnum(get("md", 0)))
        if f["md"] == 0 and f["largo"] == 0 and CAT[cod]["modo"] != "und":
            omitidas += 1
            continue
        filas.append(f)
    return filas, omitidas, mapa


def exportar_plantilla(ruta):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Elementos_a_Demoler"
    enc = ["Piso/Sector", "Tipo de Elemento / Partida", "Descripción", "N° Veces",
           "Largo (m)", "Ancho/Espesor (m)", "Alto (m)", "Metrado Directo"]
    ws.append(enc)
    _, filas = caso_ejemplo()
    for f in filas:
        ws.append([f["piso"], f["cod"], f["desc"], f["n"], f["largo"], f["ancho"], f["alto"], ""])
    for i, w in enumerate([16, 26, 60, 10, 12, 18, 10, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    ws2 = wb.create_sheet("Catalogo_Partidas")
    ws2.append(["Código", "Partida", "Und", "FE", "Peso suelto (t/m³)", "Espesor/Factor eq.", "C.U. (S/)"])
    for c in CATALOGO:
        ws2.append([c["cod"], c["nombre"], c["und"], c["fe"], c["peso"], c["esp"], c["cu"]])
    ws2.column_dimensions["B"].width = 80
    wb.save(ruta)


# =============================================================================
# 5. EXPORTACIÓN A EXCEL (.XLSX) - 3 HOJAS PROFESIONALES
# =============================================================================
_AZUL = "1F4E78"
_AZUL_CL = "DDEBF7"
_NARANJA = "C55A11"
_GRIS = "F2F2F2"
_VERDE = "E2EFDA"
_FINO = Side(style="thin", color="7F7F7F")
_BORDE = Border(left=_FINO, right=_FINO, top=_FINO, bottom=_FINO)


def _xl_encabezado(ws, titulo, datos, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1, titulo)
    c.font = Font(bold=True, size=14, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=_AZUL)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    info = [
        ("PROYECTO / TRÁMITE:", f'{datos.get("proyecto", "")} - {datos.get("tramite", "")}'),
        ("PROPIETARIO / ENTIDAD:", datos.get("propietario", "")),
        ("UBICACIÓN:", f'{datos.get("ubicacion", "")} - {datos.get("municipalidad", "")}'),
        ("RESPONSABLE:", f'{datos.get("ingeniero", "")} - {datos.get("cip", "")}    FECHA: {datos.get("fecha", "")}'),
    ]
    for i, (k, v) in enumerate(info, 2):
        ws.cell(i, 1, k).font = Font(bold=True, size=9)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=ncols)
        ws.cell(i, 2, v).font = Font(size=9)
    return 7


def _xl_tabla_enc(ws, fila, encabezados, color=_AZUL):
    for j, h in enumerate(encabezados, 1):
        c = ws.cell(fila, j, h)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _BORDE
    ws.row_dimensions[fila].height = 32


def _xl_celda(ws, fila, col, valor, fmtn=None, bold=False, fill=None, wrap=False, align=None):
    c = ws.cell(fila, col, valor)
    c.border = _BORDE
    c.font = Font(bold=bold, size=9)
    if fmtn:
        c.number_format = fmtn
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(wrap_text=wrap, vertical="center",
                            horizontal=align or ("right" if isinstance(valor, (int, float)) else "left"))
    return c


def _xl_anchos(ws, anchos):
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _xl_impresion(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


def exportar_excel(datos, res, ruta):
    wb = openpyxl.Workbook()
    filas = sorted(res["filas"], key=lambda r: (r["fase"], r["cod"]))
    N2 = "#,##0.00"
    N3 = "#,##0.000"

    # ---------------- HOJA 1: METRADOS RNM ----------------
    ws = wb.active
    ws.title = "1_Metrados_Demolicion_RNM"
    enc = ["Item", "Código", "Piso / Sector", "Partida (RNM)", "Descripción", "N° Veces",
           "Largo (m)", "Ancho / Esp. (m)", "Alto (m)", "Und", "Metrado"]
    f0 = _xl_encabezado(ws, "PLANILLA DE METRADOS DE DEMOLICIÓN Y DESMONTAJE - RD N° 073-2010-VIVIENDA (RNM)",
                        datos, len(enc))
    _xl_tabla_enc(ws, f0, enc)
    fila = f0 + 1
    fase_act = None
    ini = fila
    for i, r in enumerate(filas, 1):
        if r["fase"] != fase_act:
            fase_act = r["fase"]
            ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=len(enc))
            _xl_celda(ws, fila, 1, f"FASE {fase_act}: {FASES[fase_act][0].upper()}", bold=True, fill=_AZUL_CL)
            fila += 1
        vals = [i, r["cod"], r["piso"], r["partida"], r["desc"], fnum(r["n"], 1), fnum(r["largo"]),
                fnum(r["ancho"]), fnum(r["alto"]), r["und"], round(r["metrado"], 2)]
        for j, v in enumerate(vals, 1):
            _xl_celda(ws, fila, j, v, N2 if j >= 6 and j != 10 else None, wrap=j in (4, 5),
                      align="center" if j in (1, 2, 10) else None)
        fila += 1
    fila += 1
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=len(enc))
    _xl_celda(ws, fila, 1, "RESUMEN DE METRADOS POR PARTIDA", bold=True, fill=_AZUL_CL)
    fila += 1
    _xl_tabla_enc(ws, fila, ["Código", "", "", "Partida", "", "", "", "", "", "Und", "Metrado total"])
    fila += 1
    for cod, m in sorted(res["agrupado"].items(), key=lambda kv: (CAT[kv[0]]["fase"], kv[0])):
        _xl_celda(ws, fila, 1, cod, align="center")
        ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=9)
        _xl_celda(ws, fila, 2, CAT[cod]["nombre"])
        _xl_celda(ws, fila, 10, CAT[cod]["und"], align="center")
        _xl_celda(ws, fila, 11, round(m, 2), N2, bold=True)
        fila += 1
    fila += 1
    ws.cell(fila, 1, "Criterio RNM: m³ = N° x L x A x H ; muros (m²) = N° x L x H ; losas, pisos y coberturas (m²) "
                     "= N° x L x A ; desmontajes (und) = N° de piezas. No se descuentan vanos menores a 0.25 m².").font = \
        Font(italic=True, size=8)
    _xl_anchos(ws, [6, 9, 14, 48, 46, 9, 10, 12, 9, 7, 13])
    ws.freeze_panes = ws.cell(f0 + 1, 1)
    _xl_impresion(ws)

    # ---------------- HOJA 2: ESPONJAMIENTO Y VOLQUETES ----------------
    ws = wb.create_sheet("2_Calc_Esponjamiento_Volquetes")   # Excel limita a 31 caracteres
    enc = ["Item", "Código", "Piso / Sector", "Descripción", "Und", "Metrado", "Esp. / Factor eq.",
           "Vol. en Banco Vb (m³)", "FE", "% Esponj.", "Vol. Esponjado Ve (m³)", "Peso suelto (t/m³)",
           "Peso (t)", f"Viajes volquete {res['cap']} m³"]
    f0 = _xl_encabezado(ws, "CÁLCULO DE VOLUMEN DE ESCOMBROS, FACTOR DE ESPONJAMIENTO Y FLOTA DE VOLQUETES",
                        datos, len(enc))
    _xl_tabla_enc(ws, f0, enc, _NARANJA)
    fila = f0 + 1
    ini = fila
    for i, r in enumerate(filas, 1):
        vals = [i, r["cod"], r["piso"], r["desc"], r["und"], round(r["metrado"], 2), round(r["esp_eq"], 3),
                round(r["vb"], 3), r["fe"], r["pct"] / 100.0, round(r["ve"], 3), CAT[r["cod"]]["peso"],
                round(r["peso"], 2), round(r["viajes"], 2)]
        fm = [None, None, None, None, None, N2, N3, N3, "0.00", "0%", N3, "0.00", N2, N2]
        for j, v in enumerate(vals, 1):
            _xl_celda(ws, fila, j, v, fm[j - 1], wrap=j == 4, align="center" if j in (1, 2, 5) else None)
        fila += 1
    fin = fila - 1
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=7)
    _xl_celda(ws, fila, 1, "TOTALES", bold=True, fill=_VERDE, align="right")
    for j, col in ((8, "H"), (11, "K"), (13, "M"), (14, "N")):
        _xl_celda(ws, fila, j, f"=SUM({col}{ini}:{col}{fin})", N2, bold=True, fill=_VERDE)
    fila += 1
    _xl_celda(ws, fila, 13, "N° viajes (redondeo superior):", bold=True)
    _xl_celda(ws, fila, 14, res["viajes"], "0", bold=True, fill="FFF2CC")
    fila += 2
    fl = res["flota"]
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=7)
    _xl_celda(ws, fila, 1, "CÁLCULO DEL CICLO DEL VOLQUETE Y DIMENSIONAMIENTO DE FLOTA", bold=True, fill=_AZUL_CL)
    fila += 1
    parametros = [
        ("Capacidad del volquete", fl["cap"], "m³"),
        ("Distancia a la escombrera autorizada", fl["dist"], "km"),
        ("Equipo de carguío", fl["cargador"], ""),
        ("Rendimiento de carguío", fl["rend_carg"], "m³/h"),
        ("Tiempo de carguío (Cap / Rend x 60)", fl["t_carga"], "min"),
        ("Tiempo de ida cargado a 30 km/h", fl["t_ida"], "min"),
        ("Tiempo de descarga en escombrera", fl["t_desc"], "min"),
        ("Tiempo de retorno vacío a 45 km/h", fl["t_ret"], "min"),
        ("Maniobras, espera y acomodo", fl["t_man"], "min"),
        ("CICLO TOTAL DEL VOLQUETE", fl["ciclo"], "min"),
        ("Minutos efectivos por día (jornada x 60 x 0.83)", fl["min_ef"], "min"),
        ("Viajes por volquete por día", fl["viajes_vol_dia"], "viajes"),
        ("N° óptimo de volquetes (Ciclo / (t carga + maniobra))", fl["n_opt"], "und"),
        ("N° de volquetes asignados", fl["n_vol"], "und"),
        ("Volumen esponjado total a eliminar Ve", res["ve"], "m³"),
        ("N° total de viajes = ⌈Ve / Cap⌉", res["viajes"], "viajes"),
        ("DÍAS EFECTIVOS DE ELIMINACIÓN", fl["dias_elim"], "días"),
        ("Costo por viaje (ciclo x tarifa volquete)", fl["costo_viaje"], "S/"),
    ]
    for k, v, u in parametros:
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=5)
        _xl_celda(ws, fila, 1, k, bold=k.isupper())
        _xl_celda(ws, fila, 6, round(v, 2) if isinstance(v, float) else v, N2 if isinstance(v, float) else None,
                  bold=k.isupper())
        _xl_celda(ws, fila, 7, u, align="center")
        fila += 1
    fila += 1
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=7)
    _xl_celda(ws, fila, 1, "RESUMEN DE ESCOMBROS POR TIPO DE MATERIAL (RCD)", bold=True, fill=_AZUL_CL)
    fila += 1
    _xl_tabla_enc(ws, fila, ["Grupo", "", "", "", "Vb (m³)", "Ve (m³)", "Peso (t)"], _NARANJA)
    fila += 1
    for g, v in res["grupos"].items():
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=4)
        _xl_celda(ws, fila, 1, g)
        _xl_celda(ws, fila, 5, round(v["vb"], 2), N2)
        _xl_celda(ws, fila, 6, round(v["ve"], 2), N2)
        _xl_celda(ws, fila, 7, round(v["peso"], 2), N2)
        fila += 1
    _xl_anchos(ws, [6, 9, 14, 46, 7, 11, 11, 13, 7, 9, 13, 11, 11, 13])
    ws.freeze_panes = ws.cell(f0 + 1, 1)
    _xl_impresion(ws)

    # ---------------- HOJA 3: PRESUPUESTO Y HORAS-MÁQUINA ----------------
    ws = wb.create_sheet("3_Presupuesto_y_Horas_Maquina")
    enc = ["Item", "Descripción", "Und", "Metrado", "C.U. (S/)", "Parcial (S/)"]
    f0 = _xl_encabezado(ws, "PRESUPUESTO DE DEMOLICIÓN, ELIMINACIÓN Y SEGURIDAD G.050 - HORAS-MÁQUINA", datos, len(enc))
    _xl_tabla_enc(ws, f0, enc, "375623")
    fila = f0 + 1
    pres = res["presupuesto"]
    filas_sub = []
    for s in pres["secciones"]:
        _xl_celda(ws, fila, 1, s["cod"], bold=True, fill=_VERDE, align="center")
        ws.merge_cells(start_row=fila, start_column=2, end_row=fila, end_column=5)
        _xl_celda(ws, fila, 2, s["titulo"], bold=True, fill=_VERDE)
        fsec = fila
        fila += 1
        ini = fila
        for it in s["items"]:
            _xl_celda(ws, fila, 1, it["item"], align="center")
            _xl_celda(ws, fila, 2, it["desc"], wrap=True)
            _xl_celda(ws, fila, 3, it["und"], align="center")
            _xl_celda(ws, fila, 4, it["metrado"], N2)
            _xl_celda(ws, fila, 5, it["cu"], N2)
            _xl_celda(ws, fila, 6, f"=ROUND(D{fila}*E{fila},2)", N2)
            fila += 1
        _xl_celda(ws, fsec, 6, f"=SUM(F{ini}:F{fila - 1})", N2, bold=True, fill=_VERDE)
        filas_sub.append(fsec)
    fila += 1
    resumen = [
        ("COSTO DIRECTO", "=" + "+".join(f"F{x}" for x in filas_sub)),
        ("GASTOS GENERALES (10 % C.D.)", f"=ROUND(F{fila}*0.10,2)"),
        ("UTILIDAD (8 % C.D.)", f"=ROUND(F{fila}*0.08,2)"),
        ("SUB TOTAL", f"=F{fila}+F{fila + 1}+F{fila + 2}"),
        ("IGV (18 %)", f"=ROUND(F{fila + 3}*0.18,2)"),
        ("PRESUPUESTO TOTAL", f"=F{fila + 3}+F{fila + 4}"),
    ]
    for k, v in resumen:
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=5)
        tot = k == "PRESUPUESTO TOTAL"
        _xl_celda(ws, fila, 1, k, bold=True, align="right", fill="FFF2CC" if tot else None)
        _xl_celda(ws, fila, 6, v, '"S/ "#,##0.00', bold=True, fill="FFF2CC" if tot else None)
        fila += 1
    ws.cell(fila, 1, f"Son: {numero_a_letras(pres['total'])}").font = Font(italic=True, size=9, bold=True)
    fila += 2
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=6)
    _xl_celda(ws, fila, 1, "CUADRO DE HORAS-MÁQUINA (H-M) DE EQUIPOS", bold=True, fill=_AZUL_CL)
    fila += 1
    _xl_tabla_enc(ws, fila, ["N°", "Equipo", "H-M", "Tarifa (S/ h)", "Parcial (S/)", "Observación"], "375623")
    fila += 1
    for i, e in enumerate(res["hm"], 1):
        _xl_celda(ws, fila, 1, i, align="center")
        _xl_celda(ws, fila, 2, e["eq"], wrap=True)
        _xl_celda(ws, fila, 3, round(e["h"], 2), N2)
        _xl_celda(ws, fila, 4, e["tarifa"], N2)
        _xl_celda(ws, fila, 5, f"=ROUND(C{fila}*D{fila},2)", N2)
        _xl_celda(ws, fila, 6, e["nota"], wrap=True)
        fila += 1
    fila += 1
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=6)
    _xl_celda(ws, fila, 1, f"CRONOGRAMA DE EJECUCIÓN - PLAZO TOTAL: {res['plazo']} DÍAS CALENDARIO", bold=True,
              fill=_AZUL_CL)
    fila += 1
    _xl_tabla_enc(ws, fila, ["Fase", "Actividad", "Día inicio", "Día fin", "Duración (días)", ""], "375623")
    fila += 1
    for fa, act, d1, d2 in res["cron"]:
        _xl_celda(ws, fila, 1, fa, align="center")
        _xl_celda(ws, fila, 2, act, wrap=True)
        _xl_celda(ws, fila, 3, d1, "0")
        _xl_celda(ws, fila, 4, d2, "0")
        _xl_celda(ws, fila, 5, d2 - d1 + 1, "0")
        fila += 1
    _xl_anchos(ws, [9, 70, 9, 13, 15, 32])
    ws.freeze_panes = ws.cell(f0 + 1, 1)
    _xl_impresion(ws)

    wb.properties.title = "Planilla de Metrados de Demolición"
    wb.properties.creator = datos.get("ingeniero", "")
    wb.save(ruta)
    return ruta


def numero_a_letras(monto):
    """Convierte un monto a letras (formato de presupuesto peruano)."""
    unidades = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ",
                "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO",
                "DIECINUEVE", "VEINTE", "VEINTIUNO", "VEINTIDÓS", "VEINTITRÉS", "VEINTICUATRO",
                "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"]
    decenas = ["", "", "", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
    centenas = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS",
                "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"]

    def menor_mil(n):
        if n == 0:
            return ""
        if n == 100:
            return "CIEN"
        c, r = divmod(n, 100)
        txt = centenas[c]
        if r:
            if r < 30:
                t = unidades[r]
            else:
                d, u = divmod(r, 10)
                t = decenas[d] + (" Y " + unidades[u] if u else "")
            txt = (txt + " " + t).strip()
        return txt

    entero = int(monto)
    cent = int(round((monto - entero) * 100))
    if cent == 100:
        entero, cent = entero + 1, 0
    if entero == 0:
        letras = "CERO"
    else:
        mill, resto = divmod(entero, 1_000_000)
        miles, uni = divmod(resto, 1000)
        partes = []
        if mill:
            partes.append("UN MILLÓN" if mill == 1 else menor_mil(mill) + " MILLONES")
        if miles:
            txt_m = menor_mil(miles)
            txt_m = txt_m[:-3] + "ÚN" if txt_m.endswith("VEINTIUNO") else (
                txt_m[:-1] if txt_m.endswith("UNO") else txt_m)
            partes.append("MIL" if miles == 1 else txt_m + " MIL")
        if uni:
            partes.append(menor_mil(uni))
        letras = " ".join(partes)
    return f"{letras} CON {cent:02d}/100 SOLES"


# =============================================================================
# 6. GRÁFICOS E IMÁGENES CON PILLOW (para el informe Word)
# =============================================================================
def _fuente(tam, negrita=False):
    candidatas = ["arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"] if negrita else \
                 ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf"]
    for nombre in candidatas:
        try:
            return ImageFont.truetype(nombre, tam)
        except Exception:
            continue
    try:
        return ImageFont.load_default(size=tam)
    except Exception:
        return ImageFont.load_default()


def _texto_centrado(draw, xy, txt, fuente, color="black"):
    x, y = xy
    bb = draw.textbbox((0, 0), txt, font=fuente)
    draw.text((x - (bb[2] - bb[0]) / 2, y - (bb[3] - bb[1]) / 2), txt, font=fuente, fill=color)


def grafico_volumenes(res):
    """Gráfico de barras Vb vs Ve por tipo de material (PNG en memoria)."""
    W, H = 1700, 820
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    f_t, f_n, f_s = _fuente(34, True), _fuente(22), _fuente(20)
    _texto_centrado(d, (W / 2, 40), "VOLUMEN DE ESCOMBROS POR TIPO DE MATERIAL (m³)", f_t, "#1F4E78")
    grupos = list(res["grupos"].items()) or [("Sin datos", dict(vb=0, ve=0, peso=0))]
    x0, y0, x1, y1 = 130, 110, W - 60, H - 150
    vmax = max([g["ve"] for _, g in grupos] + [1.0])
    paso = 10 ** math.floor(math.log10(vmax)) if vmax > 0 else 1
    tope = math.ceil(vmax / paso) * paso
    for k in range(6):
        y = y1 - (y1 - y0) * k / 5
        d.line([(x0, y), (x1, y)], fill="#D9D9D9", width=2)
        d.text((x0 - 110, y - 12), f"{tope * k / 5:,.1f}", font=f_s, fill="#404040")
    d.line([(x0, y0), (x0, y1), (x1, y1)], fill="black", width=3)
    n = len(grupos)
    ancho_g = (x1 - x0) / n
    bw = min(ancho_g * 0.32, 110)
    for i, (nom, g) in enumerate(grupos):
        cx = x0 + ancho_g * (i + 0.5)
        for j, (val, col) in enumerate(((g["vb"], "#2E75B6"), (g["ve"], "#ED7D31"))):
            h = (y1 - y0) * val / tope if tope else 0
            bx = cx - bw - 4 if j == 0 else cx + 4
            d.rectangle([bx, y1 - h, bx + bw, y1], fill=col, outline="black")
            _texto_centrado(d, (bx + bw / 2, y1 - h - 16), f"{val:,.1f}", f_s)
        _texto_centrado(d, (cx, y1 + 28), cortar(nom, 22), f_n)
    for j, (txt, col) in enumerate((("Volumen en banco Vb (m³)", "#2E75B6"), ("Volumen esponjado Ve (m³)", "#ED7D31"))):
        lx = W / 2 - 420 + j * 440
        d.rectangle([lx, H - 75, lx + 40, H - 45], fill=col, outline="black")
        d.text((lx + 52, H - 75), txt, font=f_n, fill="black")
    bio = io.BytesIO()
    img.save(bio, "PNG")
    bio.seek(0)
    return bio


def imagen_sin_foto(categoria):
    W, H = 1600, 1000
    img = Image.new("RGB", (W, H), "#EDEDED")
    d = ImageDraw.Draw(img)
    d.rectangle([20, 20, W - 20, H - 20], outline="#7F7F7F", width=6)
    d.line([(20, 20), (W - 20, H - 20)], fill="#BFBFBF", width=4)
    d.line([(W - 20, 20), (20, H - 20)], fill="#BFBFBF", width=4)
    d.rectangle([W / 2 - 560, H / 2 - 110, W / 2 + 560, H / 2 + 110], fill="#FFFFFF", outline="#7F7F7F", width=3)
    _texto_centrado(d, (W / 2, H / 2 - 45), f"FOTOGRAFÍA: {categoria.upper()}", _fuente(48, True), "#1F4E78")
    _texto_centrado(d, (W / 2, H / 2 + 45), "(Adjuntar registro fotográfico del estado actual)", _fuente(36), "#404040")
    bio = io.BytesIO()
    img.save(bio, "JPEG", quality=85)
    bio.seek(0)
    return bio


def preparar_foto(ruta):
    im = Image.open(ruta)
    im = ImageOps.exif_transpose(im).convert("RGB")
    im.thumbnail((1600, 1600))
    bio = io.BytesIO()
    im.save(bio, "JPEG", quality=85)
    bio.seek(0)
    return bio


# =============================================================================
# 7. EXPORTACIÓN A WORD (.DOCX): MEMORIA + PLAN SST G.050 + MITIGACIÓN
# =============================================================================
class _Doc:
    """Ayudante para construir el expediente en python-docx."""

    def __init__(self, datos):
        self.d = datos
        self.doc = Document()
        sec = self.doc.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
        sec.left_margin, sec.right_margin = Cm(2.5), Cm(2.0)
        sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)
        st = self.doc.styles["Normal"]
        st.font.name = "Arial"
        st.font.size = Pt(10.5)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        st.paragraph_format.space_after = Pt(4)
        st.paragraph_format.line_spacing = 1.15
        for nombre, tam, color in (("Heading 1", 13, "1F4E78"), ("Heading 2", 11.5, "2E75B6"),
                                   ("Heading 3", 10.5, "404040")):
            h = self.doc.styles[nombre]
            h.font.name = "Arial"
            h.font.size = Pt(tam)
            h.font.bold = True
            h.font.color.rgb = RGBColor.from_string(color)
            h.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(5)
        self._cabecera_pie()

    def _cabecera_pie(self):
        sec = self.doc.sections[0]
        sec.different_first_page_header_footer = True
        p = sec.header.paragraphs[0]
        p.text = cortar(f'{self.d.get("tramite", "")} | {self.d.get("proyecto", "")}', 110)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for r in p.runs:
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        pf = sec.footer.paragraphs[0]
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pf.add_run(f'{self.d.get("ingeniero", "")} - {self.d.get("cip", "")}   |   Página ')
        r.font.size = Pt(8)
        self._campo(pf, "PAGE")
        r = pf.add_run(" de ")
        r.font.size = Pt(8)
        self._campo(pf, "NUMPAGES")

    @staticmethod
    def _campo(par, instr):
        run = par.add_run()
        run.font.size = Pt(8)
        f1 = OxmlElement("w:fldChar")
        f1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        f2 = OxmlElement("w:fldChar")
        f2.set(qn("w:fldCharType"), "end")
        run._r.append(f1)
        run._r.append(it)
        run._r.append(f2)

    def h1(self, t):
        self.doc.add_heading(t, level=1)

    def h2(self, t):
        self.doc.add_heading(t, level=2)

    def h3(self, t):
        self.doc.add_heading(t, level=3)

    def p(self, t, bold=False, italic=False, align="j", size=None):
        par = self.doc.add_paragraph()
        r = par.add_run(t)
        r.bold, r.italic = bold, italic
        if size:
            r.font.size = Pt(size)
        par.alignment = {"j": WD_ALIGN_PARAGRAPH.JUSTIFY, "c": WD_ALIGN_PARAGRAPH.CENTER,
                         "l": WD_ALIGN_PARAGRAPH.LEFT, "r": WD_ALIGN_PARAGRAPH.RIGHT}[align]
        return par

    def pk(self, clave, valor):
        par = self.doc.add_paragraph()
        par.add_run(clave).bold = True
        par.add_run(valor)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return par

    def vinetas(self, items, estilo="List Bullet"):
        for it in items:
            par = self.doc.add_paragraph(style=estilo)
            if isinstance(it, tuple):
                par.add_run(it[0]).bold = True
                par.add_run(it[1])
            else:
                par.add_run(it)
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            par.paragraph_format.space_after = Pt(2)

    def numeros(self, items):
        """Lista numerada que reinicia en 1 en cada bloque."""
        for i, it in enumerate(items, 1):
            par = self.doc.add_paragraph()
            par.paragraph_format.left_indent = Cm(0.9)
            par.paragraph_format.first_line_indent = Cm(-0.6)
            par.paragraph_format.space_after = Pt(2)
            par.add_run(f"{i}.  ").bold = True
            par.add_run(it)
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    @staticmethod
    def sombrear(celda, color):
        tcPr = celda._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    def tabla(self, enc, filas, anchos=None, tam=8.5, color="1F4E78", total=False, alinear_num=True):
        t = self.doc.add_table(rows=1, cols=len(enc))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, h in enumerate(enc):
            c = t.rows[0].cells[j]
            c.text = ""
            r = c.paragraphs[0].add_run(str(h))
            r.bold = True
            r.font.size = Pt(tam)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.sombrear(c, color)
        for i, fila in enumerate(filas):
            cells = t.add_row().cells
            ultima = total and i == len(filas) - 1
            for j, v in enumerate(fila):
                cells[j].text = ""
                r = cells[j].paragraphs[0].add_run(str(v))
                r.font.size = Pt(tam)
                r.bold = ultima
                es_num = alinear_num and bool(re.fullmatch(r"[S/\s\-]*[\d,]+(\.\d+)?\s*%?", str(v)))
                cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT if es_num else WD_ALIGN_PARAGRAPH.LEFT
                if ultima:
                    self.sombrear(cells[j], "FFF2CC")
                elif i % 2 == 1:
                    self.sombrear(cells[j], "F2F2F2")
        if anchos:
            t.autofit = False
            for j, w in enumerate(anchos):
                t.columns[j].width = Cm(w)
            for row in t.rows:
                for j, w in enumerate(anchos):
                    row.cells[j].width = Cm(w)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    def salto(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def imagen(self, flujo, ancho_cm=15.5, pie=None):
        self.doc.add_picture(flujo, width=Cm(ancho_cm))
        self.doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if pie:
            self.p(pie, italic=True, align="c", size=9)


def exportar_word(datos, res, fotos, checklist, ruta):
    D = _Doc(datos)
    d = datos
    fl = res["flota"]
    pres = res["presupuesto"]
    met = res["metodo"]
    hoy = d.get("fecha") or datetime.date.today().strftime("%d/%m/%Y")
    n_pisos = int(fnum(d.get("n_pisos"), 2))

    # ------------------------------ PORTADA ------------------------------
    for _ in range(3):
        D.p("")
    D.p(d.get("municipalidad", "").upper(), bold=True, align="c", size=14)
    D.p("")
    D.p("EXPEDIENTE TÉCNICO DE DEMOLICIÓN", bold=True, align="c", size=22)
    D.p(d.get("tramite", "").upper(), bold=True, align="c", size=13)
    D.p("")
    D.p("MEMORIA DESCRIPTIVA DE DEMOLICIÓN\nPLAN DE SEGURIDAD Y SALUD EN EL TRABAJO (NORMA G.050)\n"
        "PLAN DE MITIGACIÓN DE POLVO Y RUIDO\nCÁLCULO DE ESCOMBROS Y PLAN DE EVACUACIÓN DE RCD",
        bold=True, align="c", size=12)
    D.p("")
    D.tabla(["DATOS GENERALES", ""], [
        ["Proyecto", d.get("proyecto", "")],
        ["Propietario / Entidad", d.get("propietario", "")],
        ["Ubicación", d.get("ubicacion", "")],
        ["Municipalidad / Entidad", d.get("municipalidad", "")],
        ["Área de terreno", f'{fnum(d.get("area_terreno")):,.2f} m²'],
        ["Área techada a demoler", f'{fnum(d.get("area_techada")):,.2f} m²'],
        ["N° de pisos a demoler", str(n_pisos)],
        ["Profesional responsable", f'{d.get("ingeniero", "")} - {d.get("cip", "")}'],
        ["Fecha", hoy],
    ], anchos=[5.5, 11], tam=10, alinear_num=False)
    for _ in range(4):
        D.p("")
    D.p("_______________________________________", align="c")
    D.p(f'{d.get("ingeniero", "")}\nIngeniero Civil - {d.get("cip", "")}', align="c", bold=True)
    D.salto()

    # ------------------------------ ÍNDICE ------------------------------
    D.h1("ÍNDICE")
    indice = [
        "1. Memoria descriptiva del estado actual y alcance de la demolición",
        "2. Corte preventivo de suministros",
        "3. Metodología y secuencia técnica de demolición",
        "4. Protección de propiedades vecinas y vía pública",
        "5. Cálculo de volumen de escombros, esponjamiento y plan de evacuación (RCD)",
        "6. Plan de Seguridad y Salud en el Trabajo en demoliciones (Norma G.050)",
        "7. Plan de mitigación de polvo y ruido",
        "8. Presupuesto, horas-máquina y cronograma de ejecución",
        "9. Conclusiones, recomendaciones y responsabilidad profesional",
        "10. Anexo fotográfico del inmueble y colindancias",
    ]
    for it in indice:
        D.p(it, align="l", size=11)
    D.p("")
    D.p("Base legal y normativa aplicable", bold=True)
    D.vinetas([
        "Ley N° 29090, Ley de Regulación de Habilitaciones Urbanas y de Edificaciones, y sus modificatorias.",
        "D.S. N° 029-2019-VIVIENDA, Reglamento de Licencias de Habilitación Urbana y Licencias de Edificación.",
        "Reglamento Nacional de Edificaciones (RNE) - Norma G.050 Seguridad Durante la Construcción.",
        "D.S. N° 011-2019-TR, Reglamento de Seguridad y Salud en el Trabajo para el Sector Construcción.",
        "Ley N° 29783, Ley de Seguridad y Salud en el Trabajo, y su Reglamento D.S. N° 005-2012-TR.",
        "D.S. N° 003-2013-VIVIENDA, Reglamento para la Gestión y Manejo de los Residuos de las Actividades "
        "de la Construcción y Demolición, modificado por D.S. N° 019-2016-VIVIENDA.",
        "D. Leg. N° 1278, Ley de Gestión Integral de Residuos Sólidos, y su Reglamento D.S. N° 014-2017-MINAM.",
        "RD N° 073-2010-VIVIENDA/VMCS-DNC, Norma Técnica Metrados para Obras de Edificación y Habilitaciones Urbanas.",
        "D.S. N° 085-2003-PCM, Estándares Nacionales de Calidad Ambiental para Ruido.",
        "D.S. N° 003-2017-MINAM, Estándares de Calidad Ambiental (ECA) para Aire.",
        "Ordenanzas municipales del distrito sobre horarios de obra, ocupación de vía pública y ruidos molestos.",
    ])
    D.salto()

    # ------------------------------ 1. MEMORIA ------------------------------
    D.h1("1. MEMORIA DESCRIPTIVA DEL ESTADO ACTUAL Y ALCANCE DE LA DEMOLICIÓN")
    D.h2("1.1 Generalidades")
    D.p(f'La presente Memoria Descriptiva forma parte del expediente "{d.get("tramite", "")}" correspondiente al '
        f'proyecto "{d.get("proyecto", "")}", solicitado por {d.get("propietario", "")}, para el inmueble ubicado en '
        f'{d.get("ubicacion", "")}, jurisdicción de la {d.get("municipalidad", "")}. El documento describe el estado '
        'actual de la edificación existente, el alcance de los trabajos de demolición y desmontaje, la metodología '
        'y secuencia constructiva, las medidas de protección de las propiedades colindantes y de la vía pública, el '
        'cálculo del volumen de residuos de construcción y demolición (RCD) y su plan de evacuación, así como el Plan '
        'de Seguridad y Salud en el Trabajo de acuerdo con la Norma G.050 del Reglamento Nacional de Edificaciones.')
    D.h2("1.2 Objeto")
    D.p('El objeto de los trabajos es la demolición de la edificación existente, dejando el terreno libre, limpio y '
        'nivelado para la futura ejecución del proyecto de edificación, garantizando en todo momento la estabilidad de '
        'las construcciones vecinas, la seguridad de los trabajadores y de los transeúntes, y la correcta disposición '
        'final de los residuos en una escombrera autorizada conforme al D.S. N° 003-2013-VIVIENDA.')
    D.h2("1.3 Ubicación, linderos y accesos")
    D.p(f'El predio se ubica en {d.get("ubicacion", "")}, con un área de terreno de {fnum(d.get("area_terreno")):,.2f} m². '
        'El acceso vehicular para los volquetes y el equipo de carguío se realizará por la vía principal frente al '
        'predio, previa autorización municipal de uso temporal de la vía pública y con vigía de tránsito permanente. '
        'Los linderos laterales y de fondo colindan con edificaciones existentes, cuyas condiciones se registrarán '
        'en el Acta de Constatación previa al inicio de obra (ver Capítulo 4 y Anexo Fotográfico).')
    D.h2("1.4 Descripción del estado actual de la edificación")
    D.p(f'La edificación existente cuenta con {n_pisos} piso(s) y azotea, con un área techada total a demoler de '
        f'{fnum(d.get("area_techada")):,.2f} m². El sistema estructural existente corresponde a: {d.get("sistema", "")}. '
        'De la inspección ocular realizada se aprecia que la edificación presenta desgaste propio de su antigüedad: '
        'fisuras en tabiquería, humedad en muros de primer piso, eflorescencias y acero de refuerzo con signos de '
        'corrosión en zonas puntuales. No se evidencian elementos con riesgo de colapso inminente; sin embargo, '
        'previo al inicio de los trabajos el Ingeniero Responsable verificará la estabilidad de voladizos, parapetos, '
        'escaleras y tanques elevados.')
    D.tabla(["Característica", "Descripción"], [
        ["Sistema estructural existente", d.get("sistema", "")],
        ["N° de pisos a demoler", f"{n_pisos} piso(s) + azotea"],
        ["Área de terreno", f'{fnum(d.get("area_terreno")):,.2f} m²'],
        ["Área techada a demoler", f'{fnum(d.get("area_techada")):,.2f} m²'],
        ["Método de demolición", d.get("metodo", "")],
        ["Volumen en banco total (Vb)", f'{res["vb"]:,.2f} m³'],
        ["Volumen esponjado total (Ve)", f'{res["ve"]:,.2f} m³'],
        ["Peso estimado de escombros", f'{res["peso"]:,.2f} t'],
        ["Escombrera autorizada", f'{d.get("escombrera", "")} (D = {fl["dist"]:.1f} km)'],
        ["Plazo estimado", f'{res["plazo"]} días calendario'],
    ], anchos=[6, 10.5], alinear_num=False)
    D.h2("1.5 Alcance de la demolición por pisos y ambientes")
    D.p('El alcance comprende la totalidad de los elementos que se listan a continuación, agrupados por piso o sector, '
        'según el metrado realizado de acuerdo con la Norma Técnica de Metrados (RD N° 073-2010-VIVIENDA):')
    por_piso = {}
    for r in res["filas"]:
        por_piso.setdefault(r["piso"] or "General", []).append(r)
    filas_t = []
    for piso, lst in por_piso.items():
        for r in lst:
            filas_t.append([piso, cortar(r["desc"] or r["partida"], 70), r["und"], fmt(r["metrado"])])
    D.tabla(["Piso / Sector", "Elemento a demoler / desmontar", "Und", "Metrado"], filas_t,
            anchos=[3, 10, 1.5, 2.2])
    D.salto()

    # ------------------------------ 2. CORTE DE SUMINISTROS ------------------------------
    D.h1("2. CORTE PREVENTIVO DE SUMINISTROS")
    D.p('Antes de iniciar cualquier trabajo de desmontaje o demolición, y conforme a la Norma G.050 (demoliciones), '
        'se ejecutará la clausura definitiva y verificada de todas las acometidas de servicios públicos. Ningún trabajo '
        'de demolición podrá iniciarse sin el registro firmado de corte de suministros por el Ingeniero Responsable y '
        'el Supervisor de Seguridad.')
    D.h2("2.1 Energía eléctrica")
    D.vinetas([
        "Solicitud a la empresa concesionaria (Luz del Sur / Enel / Electro regional) del retiro del medidor y corte "
        "de la acometida desde la red pública.",
        "Aplicación del procedimiento de bloqueo y etiquetado (LOTO): candado y tarjeta de 'PELIGRO - NO OPERAR' en "
        "el tablero general, verificación de ausencia de tensión con detector y puesta a tierra temporal.",
        "Retiro de conductores, tableros y luminarias como parte de la Fase 1 de desmontaje.",
        "La energía para herramientas eléctricas se tomará de un tablero provisional de obra con interruptor "
        "diferencial de 30 mA y puesta a tierra, o de un grupo electrógeno.",
    ])
    D.h2("2.2 Agua potable")
    D.vinetas([
        "Cierre de la llave de paso y solicitud a la EPS (SEDAPAL u otra) del cierre o retiro del medidor.",
        "Se mantendrá un punto de agua provisional exclusivamente para el riego con agua nebulizada (control de PM10).",
        "Vaciado de tanques elevados y cisternas antes de su demolición para evitar sobrecargas y colapsos.",
    ])
    D.h2("2.3 Desagüe")
    D.vinetas([
        "Sellado hermético de la conexión domiciliaria de desagüe con tapón y mortero, para impedir el ingreso de "
        "escombros a la red pública y la salida de roedores.",
        "Desratización y desinsectación del inmueble con 72 horas de anticipación al inicio de la demolición.",
    ])
    D.h2("2.4 Gas natural / GLP")
    D.vinetas([
        "Solicitud a la empresa distribuidora (Cálidda u otra) del corte y retiro del medidor de gas natural.",
        "Purgado de la red interna con gas inerte o aire, verificación con explosímetro y sellado de las tuberías.",
        "Retiro previo de balones de GLP, termas y artefactos a gas fuera del área de trabajo.",
    ])
    D.h2("2.5 Telecomunicaciones y otros")
    D.p('Se coordinará con las empresas de telefonía, cable e internet el retiro de las acometidas aéreas y '
        'subterráneas. Se verificará la ausencia de cables aéreos de media tensión cercanos a la fachada; de existir, '
        'se mantendrá la distancia mínima de seguridad establecida por el Código Nacional de Electricidad.')
    D.tabla(["Servicio", "Acción", "Verificación", "Responsable"], [
        ["Electricidad", "Corte, retiro de medidor y LOTO", "Detector de tensión", "Ing. Residente"],
        ["Agua potable", "Cierre y retiro de medidor", "Prueba de presión nula", "Maestro de obra"],
        ["Desagüe", "Sellado de conexión", "Inspección visual", "Maestro de obra"],
        ["Gas", "Corte, purgado y sellado", "Explosímetro 0 % LEL", "Supervisor SST"],
        ["Telecomunicaciones", "Retiro de acometidas", "Inspección visual", "Ing. Residente"],
    ], anchos=[3.2, 5, 4, 4.3], alinear_num=False)

    # ------------------------------ 3. METODOLOGÍA ------------------------------
    D.h1("3. METODOLOGÍA Y SECUENCIA TÉCNICA DE DEMOLICIÓN")
    D.h2("3.1 Método de demolición adoptado")
    texto_met = {
        "MIXTO": "Se adopta un método MIXTO: demolición manual con martillos demoledores eléctricos y herramientas "
                 "manuales en una franja de 1.00 m adyacente a los muros medianeros y en los elementos de los pisos "
                 "superiores, y demolición mecánica con retroexcavadora equipada con martillo hidráulico (picotón) en "
                 "la zona central y en los niveles inferiores, una vez que la altura de la edificación lo permita con "
                 "seguridad. Se estima que el 30 % del volumen se demolerá manualmente y el 70 % mecánicamente.",
        "MANUAL": "Se adopta un método 100 % MANUAL con martillos demoledores eléctricos/neumáticos, combas, cinceles "
                  "y cortadoras de disco, debido a la cercanía de las edificaciones colindantes y a las restricciones "
                  "de acceso de maquinaria. Los escombros se evacuarán por manga o chute hacia el volquete.",
        "MECANICO": "Se adopta un método 100 % MECÁNICO con retroexcavadora con martillo hidráulico (picotón) y "
                    "cargador frontal. No obstante, de acuerdo con la Norma G.050, en la franja de 1.00 m adyacente a "
                    "los muros medianeros vecinos solo se permite el picado manual sin percusión pesada, por lo que "
                    "dicha franja se tratará con herramientas manuales.",
    }[met["clave"]]
    D.p(texto_met)
    D.h2("3.2 Principios generales (Norma G.050)")
    D.vinetas([
        "La demolición se ejecutará en orden DESCENDENTE estricto, nivel por nivel, de arriba hacia abajo, sin "
        "trabajar simultáneamente en niveles superpuestos.",
        "Se prohíbe la demolición por zapa (socavación en la base) y el vuelco de muros completos hacia la vía "
        "pública o hacia predios colindantes.",
        "No se permitirá acumular escombros sobre losas o entrepisos en una altura que supere su capacidad portante; "
        "el desmonte se evacuará diariamente por manga/chute o por los vanos de los niveles inferiores.",
        "Ningún elemento estructural se demolerá sin haber retirado previamente las cargas que soporta.",
        "Se mantendrán arriostrados los muros que queden libres en altura y se demolerán el mismo día.",
        "Los trabajos se suspenderán con vientos fuertes, lluvias intensas o sismos, debiendo revisarse la "
        "estabilidad de todos los elementos antes de reanudarlos.",
        "Toda la zona de trabajo estará señalizada, con acceso restringido al personal autorizado y con un vigía "
        "permanente en la vía pública durante las maniobras de volquetes.",
    ])
    D.h2("3.3 Secuencia de demolición paso a paso")
    detalles_fase = {
        1: "Retiro de coberturas livianas (calamina, aluzinc, fibrocemento) y tijerales desde la azotea, usando "
           "arnés de cuerpo completo anclado a línea de vida. Desmontaje de puertas, ventanas, vidrios, rejas, "
           "aparatos sanitarios, griferías, luminarias, tableros, tuberías expuestas y todo elemento recuperable. "
           "Los vidrios se retirarán enteros y se acopiarán en caballetes. El material recuperable se clasificará "
           "y almacenará en un área de acopio limpio para su comercialización o reutilización.",
        2: "Demolición de techos desde la azotea hacia los niveles inferiores. Se inicia por parapetos y tanques "
           "elevados, luego la losa aligerada por paños, picando la losa desde el centro del paño hacia los apoyos, "
           "cortando el acero con oxicorte o cizalla. Se retirarán los ladrillos de techo y el concreto, dejando "
           "las vigas sin carga. El personal trabajará sobre zonas no demolidas y anclado a línea de vida.",
        3: "Demolición de vigas una vez liberadas de la losa. Se apuntalarán previamente las vigas en voladizo o de "
           "gran luz. El corte se realiza en los extremos, junto a la columna, controlando la caída del elemento "
           "mediante descenso controlado o fragmentación in situ.",
        4: "Demolición de tabiques y muros no portantes de soga, de arriba hacia abajo, por hiladas, sin dejar "
           "paños libres de altura mayor a 1.50 m sin arriostrar. Se prohíbe el volteo de muros.",
        5: "Demolición de muros portantes de cabeza, placas, escaleras y columnas, nivel por nivel. En la franja "
           "de 1.00 m colindante con muros medianeros vecinos se realizará picado manual sin percusión pesada. "
           "Las columnas se debilitarán en la base y se descenderán en dirección controlada hacia el interior del "
           "predio. La escalera se demolerá al final de cada nivel, ya que constituye la vía de evacuación.",
        6: "Demolición de pisos, falsos pisos, contrapisos y veredas; luego sobrecimientos, cimientos corridos y "
           "zapatas con retroexcavadora con picotón o martillos demoledores. Antes de excavar junto a medianeros se "
           "verificará la cimentación vecina para no descalzarla. Finalmente se realizará la limpieza, relleno de "
           "oquedades y nivelación del terreno.",
    }
    filas_f = []
    for k in sorted(FASES):
        D.h3(f"Fase {k}: {FASES[k][0]}")
        D.p(detalles_fase[k])
        filas_f.append([f"Fase {k}", FASES[k][0], f'{res["dias_fase"].get(k, 0)} día(s)'])
    D.tabla(["Fase", "Actividad", "Duración estimada"], filas_f, anchos=[2.5, 9, 4])
    D.h2("3.4 Equipos y herramientas")
    D.tabla(["Equipo / herramienta", "Uso"], [
        ["Retroexcavadora 416 con martillo hidráulico (picotón)", "Demolición mecánica de concreto y cimentaciones"],
        ["Martillo demoledor eléctrico / neumático 30 kg", "Picado manual en colindancias y pisos superiores"],
        [met["cargador"], "Carguío de desmonte a volquetes"],
        [f"Camión volquete {res['cap']} m³", "Transporte de RCD a escombrera autorizada"],
        ["Camión cisterna / nebulizadores", "Riego con agua nebulizada para control de PM10"],
        ["Equipo de oxicorte y cizalla", "Corte de acero de refuerzo"],
        ["Manga / chute de evacuación de polietileno", "Descenso controlado de desmonte desde pisos altos"],
        ["Andamios certificados, puntales, líneas de vida", "Acceso y protección de trabajos en altura"],
    ], anchos=[7.5, 9], alinear_num=False)
    D.salto()

    # ------------------------------ 4. PROTECCIÓN VECINOS ------------------------------
    D.h1("4. PROTECCIÓN DE PROPIEDADES VECINAS Y VÍA PÚBLICA")
    D.h2("4.1 Acta de constatación previa")
    D.p('Antes del inicio de los trabajos, el propietario y el Ingeniero Responsable, de preferencia con la '
        'participación de un Notario Público o Juez de Paz, levantarán un Acta de Constatación del estado de las '
        'edificaciones colindantes (izquierda, derecha y fondo), con registro fotográfico y fílmico de fisuras, '
        'humedades y deterioros existentes. Se colocarán testigos de yeso fechados en las fisuras existentes para '
        'monitorear su evolución durante la obra.')
    D.h2("4.2 Protocolo de picado manual en colindancias")
    D.p('En una franja de 1.00 m medida desde los muros medianeros vecinos, la demolición será exclusivamente manual, '
        'con martillos demoledores eléctricos livianos, combas y cinceles, sin percusión pesada ni uso de martillo '
        'hidráulico, cortando primero la unión entre los elementos propios y el muro vecino para independizarlos. '
        'Los elementos se fragmentarán en piezas pequeñas que caigan hacia el interior del predio. Se prohíbe '
        'apoyar escombros o equipos contra los muros vecinos.')
    D.h2("4.3 Apuntalamiento preventivo de muros colindantes")
    D.p(f'Se instalará apuntalamiento preventivo en una longitud estimada de {res["perim_col"]:,.2f} m de muros '
        'colindantes, conformado por soleras verticales de madera de 2" x 8" adosadas al muro vecino, tornapuntas de '
        'madera tornillo de 4" x 4" o puntales metálicos telescópicos inclinados a 45°, separados como máximo cada '
        '2.40 m, con cuñas de madera dura para el ajuste y "muertos" de concreto f\'c = 140 kg/cm² o durmientes '
        'anclados con estacas de fierro. El apuntalamiento se mantendrá hasta que la nueva edificación brinde el '
        'confinamiento lateral definitivo (ver Lámina D-01).')
    D.h2("4.4 Malla raschel antipolvo perimetral")
    D.p(f'En los linderos colindantes se instalará malla raschel de polietileno de alta densidad (80 % sombra) de '
        f'H = 4.00 m, sobre postes de tubo F.G. de 2" cada 3.00 m, con cable tensor superior e inferior, en un área '
        f'estimada de {res["perim_col"] * 4:,.2f} m². La malla se elevará conforme a la altura de trabajo para '
        'retener partículas y fragmentos.')
    D.h2("4.5 Cerco provisional y bandeja peatonal")
    D.p(f'En el frente del predio se instalará un cerco perimétrico provisional de H = 2.40 m en una longitud '
        f'aproximada de {res["frente"] + 2:,.2f} m, con postes de madera de 3" x 3" cada 2.40 m empotrados en dados de '
        'concreto, paneles de triplay de 18 mm o calamina galvanizada pintada, puerta de ingreso vehicular y puerta '
        'peatonal independiente. Sobre la vereda se instalará una bandeja protectora peatonal (pasadizo cubierto) '
        'con techo de tablones y malla, con iluminación nocturna y señalización reflectiva.')
    D.h2("4.6 Gestión del tránsito y uso de la vía pública")
    D.vinetas([
        "Autorización municipal de interferencia de vía para la permanencia del volquete durante el carguío.",
        "Vigía con paletas PARE/SIGA, chaleco reflectivo y radio durante el ingreso y salida de volquetes.",
        "Volquetes con tolva cubierta con toldo o malla para evitar la caída de material en la vía pública.",
        "Lavado de llantas y limpieza diaria de la vereda y calzada frente a la obra.",
        "Prohibido el acopio de escombros en la vía pública o en las bermas.",
    ])
    D.salto()

    # ------------------------------ 5. CÁLCULO DE ESCOMBROS ------------------------------
    D.h1("5. CÁLCULO DE VOLUMEN DE ESCOMBROS, ESPONJAMIENTO Y PLAN DE EVACUACIÓN DE RCD")
    D.h2("5.1 Criterios de cálculo")
    D.vinetas([
        ("Volumen sólido en banco: ", "Vb = Σ (N° x L x An x H); en partidas metradas por m² se multiplica por el "
                                      "espesor equivalente (losas aligeradas 0.12 a 0.16 m³/m², descontando el "
                                      "volumen de los ladrillos de techo)."),
        ("Volumen suelto esponjado: ", "Ve = Vb x FE, donde FE es el Factor de Esponjamiento del material."),
        ("Peso de escombros: ", "P = Ve x γ suelto (t/m³)."),
        ("Número de viajes: ", f"N = ⌈Ve / Capacidad del volquete⌉ = ⌈{res['ve']:,.2f} / {res['cap']}⌉ = "
                               f"{res['viajes']} viajes."),
    ])
    D.h2("5.2 Cuadro detallado de volumen de escombros y factores de esponjamiento")
    filas_e = []
    for r in sorted(res["filas"], key=lambda x: (x["fase"], x["cod"])):
        filas_e.append([r["cod"], cortar(r["desc"] or r["partida"], 42), r["und"], fmt(r["metrado"]),
                        fmt(r["vb"]), f'{r["pct"]:.0f}%', fmt(r["ve"]), fmt(r["peso"]), fmt(r["viajes"])])
    filas_e.append(["", "TOTAL", "", "", fmt(res["vb"]), "", fmt(res["ve"]), fmt(res["peso"]), str(res["viajes"])])
    D.tabla(["Cód.", "Descripción", "Und", "Metrado", "Vb (m³)", "% Esp.", "Ve (m³)", "Peso (t)", "Viajes"],
            filas_e, anchos=[1.3, 5.6, 1, 1.6, 1.6, 1.2, 1.6, 1.6, 1.3], tam=7.5, color="C55A11", total=True)
    D.imagen(grafico_volumenes(res), 15.5, "Figura 1. Volumen en banco vs. volumen esponjado por tipo de material.")
    D.h2("5.3 Ciclo del volquete y dimensionamiento de la flota")
    D.tabla(["Parámetro", "Valor"], [
        ["Capacidad del volquete", f'{fl["cap"]} m³'],
        ["Distancia a escombrera autorizada", f'{fl["dist"]:.2f} km'],
        ["Equipo de carguío / rendimiento", f'{fl["cargador"]} / {fl["rend_carg"]:.0f} m³/h'],
        ["Tiempo de carguío", f'{fl["t_carga"]:.1f} min'],
        ["Ida cargado a 30 km/h", f'{fl["t_ida"]:.1f} min'],
        ["Descarga en escombrera", f'{fl["t_desc"]:.1f} min'],
        ["Retorno vacío a 45 km/h", f'{fl["t_ret"]:.1f} min'],
        ["Maniobras y espera", f'{fl["t_man"]:.1f} min'],
        ["Ciclo total", f'{fl["ciclo"]:.1f} min'],
        ["Viajes por volquete por día", f'{fl["viajes_vol_dia"]}'],
        ["N° de volquetes asignados", f'{fl["n_vol"]} (óptimo teórico: {fl["n_opt"]})'],
        ["Total de viajes", f'{res["viajes"]}'],
        ["Días efectivos de eliminación", f'{fl["dias_elim"]}'],
    ], anchos=[8, 8.5], alinear_num=False)
    D.h2("5.4 Plan de evacuación y gestión de RCD (D.S. N° 003-2013-VIVIENDA)")
    D.vinetas([
        "Segregación en la fuente: (a) material recuperable (puertas, ventanas, aparatos sanitarios, acero de "
        "refuerzo, calaminas); (b) residuos inertes de concreto, ladrillo y adobe; (c) residuos peligrosos "
        "(fibrocemento con asbesto, aceites, pinturas) que se manejarán por una EO-RS autorizada.",
        "El acero de refuerzo se separará y se comercializará como chatarra; la madera recuperable se reutilizará.",
        "Las planchas de fibrocemento antiguas que pudieran contener asbesto se retirarán enteras, humedecidas, "
        "embaladas en polietileno y dispuestas en un relleno de seguridad autorizado.",
        f'Los RCD inertes se transportarán en volquetes con tolva cubierta hacia la escombrera autorizada '
        f'"{d.get("escombrera", "")}", a una distancia de {fl["dist"]:.1f} km.',
        "Cada viaje contará con su Manifiesto de Manejo de Residuos y la guía de remisión correspondiente; la "
        "escombrera emitirá la constancia de recepción que se archivará en el cuaderno de obra.",
        "Se presentará a la Municipalidad el Plan de Manejo de Residuos y el reporte final de volúmenes dispuestos.",
    ])
    D.salto()

    # ------------------------------ 6. PLAN SST G.050 ------------------------------
    D.h1("6. PLAN DE SEGURIDAD Y SALUD EN EL TRABAJO EN DEMOLICIONES (NORMA G.050)")
    D.h2("6.1 Objetivos")
    D.vinetas([
        "Garantizar condiciones de seguridad y salud a todo el personal, visitantes y terceros durante la demolición.",
        "Identificar los peligros, evaluar los riesgos y establecer los controles (IPERC) de cada actividad.",
        "Cumplir la Norma G.050, la Ley N° 29783 y el D.S. N° 011-2019-TR.",
        "Lograr cero accidentes incapacitantes, cero daños a propiedades vecinas y cero incidentes ambientales.",
    ])
    D.h2("6.2 Organización y responsabilidades")
    D.tabla(["Cargo", "Responsabilidades principales"], [
        ["Ingeniero Responsable", "Dirección técnica, aprobación de la secuencia, verificación de estabilidad y "
                                  "firma de permisos de trabajo de alto riesgo."],
        ["Supervisor / Prevencionista SST", "Implementación del Plan SST, charlas diarias, inspecciones, ATS, "
                                            "investigación de incidentes y control de EPP."],
        ["Maestro de obra", "Ejecución de la secuencia aprobada, control del personal y orden y limpieza."],
        ["Operadores de equipo", "Check-list diario del equipo, operación segura y respeto de zonas de exclusión."],
        ["Vigía / controlador de tránsito", "Control del ingreso y salida de volquetes y protección peatonal."],
        ["Trabajadores", "Uso correcto del EPP, cumplimiento de procedimientos y reporte de actos y condiciones "
                         "inseguras."],
    ], anchos=[5, 11.5], alinear_num=False)
    D.h2("6.3 Identificación de peligros, evaluación de riesgos y controles (IPERC)")
    D.tabla(["Actividad", "Peligro", "Riesgo", "Nivel", "Medidas de control"], [
        ["Desmontaje de coberturas", "Trabajo en altura", "Caída a distinto nivel", "ALTO",
         "Línea de vida, arnés con doble línea, tablones de tránsito, permiso de trabajo en altura."],
        ["Demolición de losas", "Colapso de elementos", "Aplastamiento", "ALTO",
         "Secuencia descendente, demolición por paños, apuntalamiento, zona de exclusión."],
        ["Uso de martillos demoledores", "Vibración y ruido", "Hipoacusia, trastornos músculo-esqueléticos", "MEDIO",
         "Protección auditiva doble, rotación de personal cada 2 h, guantes antivibración."],
        ["Demolición mecánica", "Equipo en movimiento", "Atropello, golpes", "ALTO",
         "Radio de exclusión de 1.5 veces el alcance del brazo, alarma de retroceso, vigía."],
        ["Corte de acero", "Oxicorte / chispas", "Quemaduras, incendio", "MEDIO",
         "Permiso de trabajo en caliente, extintor PQS, careta de soldador, retiro de combustibles."],
        ["Generación de polvo", "Material particulado PM10", "Afección respiratoria", "MEDIO",
         "Riego nebulizado, respirador N95 / media cara con filtros P100, malla raschel."],
        ["Carguío y transporte", "Volquetes y cargador", "Atropello, choque", "ALTO",
         "Vigía, señalización, velocidad máxima 10 km/h en obra, tolva cubierta."],
        ["Trabajos junto a medianeros", "Muro vecino inestable", "Colapso, daño a terceros", "ALTO",
         "Picado manual a 1.00 m, apuntalamiento, testigos de yeso, monitoreo diario."],
        ["Instalaciones existentes", "Energía residual / gas", "Electrocución, explosión", "ALTO",
         "Corte verificado, LOTO, purgado y medición con explosímetro."],
        ["Manipulación de escombros", "Cargas pesadas, filos", "Lumbalgia, cortes", "MEDIO",
         "Carga máxima 25 kg por persona, guantes de cuero, uso de carretillas."],
    ], anchos=[3, 2.8, 3, 1.4, 6.3], tam=7.5, alinear_num=False)
    D.h2("6.4 Equipos de protección personal (EPP) específicos para demolición")
    D.tabla(["EPP", "Especificación", "Uso obligatorio"], [
        ["Casco de seguridad", "Tipo I clase E con barbiquejo (ANSI Z89.1)", "Todo el personal en obra"],
        ["Lentes de seguridad", "Policarbonato antiempañante (ANSI Z87.1)", "Todo el personal"],
        ["Careta facial", "Malla o policarbonato", "Picado, corte y uso de martillos"],
        ["Protección auditiva", "Tapones + orejeras, NRR >= 25 dB", "Exposición > 85 dB(A)"],
        ["Respirador", "Media cara con filtros P100 / N95", "Zonas con polvo"],
        ["Guantes", "Cuero reforzado / antivibración / anticorte", "Manipulación y martillos"],
        ["Calzado de seguridad", "Punta de acero y plantilla antiperforación", "Todo el personal"],
        ["Arnés de cuerpo completo", "Doble línea de anclaje con absorbedor (ANSI Z359)", "Trabajos a más de 1.80 m"],
        ["Chaleco reflectivo", "Clase 2 alta visibilidad", "Vigías y personal en vía pública"],
        ["Ropa de trabajo", "Overol o pantalón y camisa manga larga", "Todo el personal"],
    ], anchos=[3.8, 7, 5.7], alinear_num=False)
    D.h2("6.5 Trabajos en altura y líneas de vida")
    D.p('Todo trabajo a más de 1.80 m de altura requiere permiso escrito de trabajo en altura, arnés de cuerpo '
        'completo con doble línea de anclaje y absorbedor de impacto, conectado a una línea de vida horizontal de '
        'cable de acero de 1/2" con anclajes certificados capaces de resistir 2,270 kg (5,000 lb) por trabajador, '
        'instalada en elementos estructurales que no serán demolidos en esa jornada. Los bordes de losa se '
        'protegerán con barandas de 1.00 m de altura con travesaño intermedio y rodapié, y los vanos con tapas o '
        'barandas. Los andamios serán certificados, con plataformas completas y accesos seguros.')
    D.h2("6.6 Señalización")
    D.vinetas([
        "Señales de obligación (uso de casco, protección auditiva, respirador, arnés).",
        "Señales de advertencia (peligro caída de objetos, zona de demolición, maquinaria en movimiento).",
        "Señales de prohibición (prohibido el ingreso a personas no autorizadas, prohibido fumar).",
        "Señales de evacuación y emergencia (rutas de escape, punto de reunión, extintor, botiquín).",
        "Cinta de peligro roja y amarilla, conos y mallas de seguridad delimitando las zonas de exclusión.",
        "Cartel de obra con número de licencia, profesional responsable y horario de trabajo.",
    ])
    D.h2("6.7 Control de ruido")
    D.p('Se medirá el nivel de presión sonora con sonómetro tipo II. Los trabajadores expuestos a niveles '
        'superiores a 85 dB(A) usarán doble protección auditiva y se rotarán. Los equipos contarán con silenciadores '
        'y mantenimiento al día; se evitará el funcionamiento en vacío. Los trabajos ruidosos se programarán en el '
        'horario diurno autorizado por la municipalidad.')
    D.h2("6.8 Control de polvo (PM10) con agua nebulizada")
    D.p('Se aplicará riego constante con agua nebulizada (baja presión, gota fina) sobre los elementos en demolición, '
        'el desmonte acumulado y durante el carguío, sin generar escorrentía ni lodo. Se usará camión cisterna o '
        'bomba con mangueras y boquillas nebulizadoras. La malla raschel perimetral y la cobertura de las tolvas '
        'complementan el control.')
    D.h2("6.9 Plan de respuesta ante emergencias")
    D.vinetas([
        "Brigada de emergencia capacitada en primeros auxilios, lucha contra incendios y evacuación.",
        "Botiquín de primeros auxilios, camilla rígida con inmovilizadores, extintores PQS de 6 kg y 12 kg.",
        "Directorio de emergencias visible: Bomberos (116), SAMU (106), PNP (105), centro de salud más cercano, "
        "empresa aseguradora del SCTR.",
        "Procedimiento ante sismo: suspensión inmediata de trabajos, evacuación al punto de reunión y evaluación "
        "estructural antes del reingreso.",
        "Procedimiento ante colapso parcial: aislamiento de la zona, rescate solo por personal capacitado y aviso "
        "inmediato a los Bomberos.",
    ])
    D.h2("6.10 Capacitación, inspecciones y registros")
    D.vinetas([
        "Inducción de seguridad a todo trabajador nuevo (mínimo 1 hora) y charla diaria de 5 minutos.",
        "Análisis de Trabajo Seguro (ATS) diario y permisos de trabajo de alto riesgo (altura, caliente, demolición).",
        "Inspección diaria de andamios, líneas de vida, apuntalamientos, equipos y herramientas eléctricas.",
        "Registros de accidentes, incidentes, capacitación, entrega de EPP, inspecciones y monitoreo (Ley 29783).",
        "SCTR (Seguro Complementario de Trabajo de Riesgo) vigente para todo el personal.",
    ])
    D.h2("6.11 Lista de verificación previa al inicio de la demolición")
    D.tabla(["N°", "Requisito G.050 / Municipal", "Estado"],
            [[str(i), txt, "CONFORME" if checklist.get(txt) else "PENDIENTE"]
             for i, txt in enumerate(CHECKLIST_G050, 1)], anchos=[1, 12.5, 3], alinear_num=False)
    D.salto()

    # ------------------------------ 7. MITIGACIÓN ------------------------------
    D.h1("7. PLAN DE MITIGACIÓN DE POLVO Y RUIDO")
    D.h2("7.1 Identificación de impactos")
    D.tabla(["Componente", "Impacto", "Fuente", "Medida de mitigación"], [
        ["Aire", "Emisión de PM10 y PM2.5", "Demolición, carguío y transporte",
         "Riego nebulizado, malla raschel H=4.00 m, tolvas cubiertas"],
        ["Ruido", "Incremento de nivel sonoro", "Martillos, retroexcavadora, volquetes",
         "Horario diurno, silenciadores, pantallas acústicas"],
        ["Vibraciones", "Fisuras en colindancias", "Martillo hidráulico",
         "Picado manual a 1.00 m de medianeros, monitoreo con testigos"],
        ["Suelo / vía", "Derrame de desmonte", "Transporte de RCD", "Tolva cubierta, limpieza de vía, lavado de llantas"],
        ["Social", "Molestias a vecinos", "Toda la obra", "Comunicación previa, libro de reclamos, horario"],
    ], anchos=[2.5, 3.5, 4, 6.5], tam=8, alinear_num=False)
    D.h2("7.2 Medidas para el control de polvo")
    D.numeros([
        "Humedecer con agua nebulizada los elementos antes y durante su demolición.",
        "Humedecer el desmonte acumulado y durante el carguío a los volquetes.",
        "Instalar malla raschel antipolvo en todo el perímetro colindante, elevándola con la altura de trabajo.",
        "Cubrir las tolvas de los volquetes con toldo antes de salir de la obra.",
        "Evacuar el desmonte de pisos altos únicamente por manga o chute cerrado; prohibido arrojarlo al vacío.",
        "Barrer en húmedo la vereda y la calzada frente a la obra al término de cada jornada.",
        "Suspender las actividades generadoras de polvo con vientos superiores a 25 km/h.",
    ])
    D.h2("7.3 Medidas para el control de ruido y vibraciones")
    D.numeros([
        "Horario de trabajo: lunes a viernes de 08:00 a 17:00 h y sábados de 08:00 a 13:00 h (o el que fije la "
        "ordenanza municipal). Prohibido trabajar domingos y feriados.",
        "Uso de equipos con mantenimiento al día, silenciadores y apagado cuando no se utilicen.",
        "Instalación de paneles acústicos temporales frente a los predios más sensibles.",
        "Picado manual sin percusión pesada en la franja de 1.00 m adyacente a los medianeros.",
        "Comunicar a los vecinos, con 48 h de anticipación, los días de trabajos de mayor ruido.",
    ])
    D.h2("7.4 Monitoreo ambiental")
    D.tabla(["Parámetro", "Estándar (ECA)", "Frecuencia", "Ubicación"], [
        ["Ruido diurno - zona residencial", "60 dB(A) - D.S. 085-2003-PCM", "2 mediciones", "Lindero frontal y colindancias"],
        ["Ruido diurno - zona comercial", "70 dB(A) - D.S. 085-2003-PCM", "2 mediciones", "Lindero frontal"],
        ["PM10 (24 h)", "100 µg/m³ - D.S. 003-2017-MINAM", "2 mediciones", "Barlovento y sotavento"],
        ["PM2.5 (24 h)", "50 µg/m³ - D.S. 003-2017-MINAM", "2 mediciones", "Barlovento y sotavento"],
    ], anchos=[4.5, 5, 3, 4], tam=8, alinear_num=False)
    D.h2("7.5 Relaciones comunitarias")
    D.p('Se entregará a los vecinos colindantes una comunicación escrita con el inicio y la duración estimada de los '
        'trabajos, el horario, y los datos de contacto del Ingeniero Responsable. Se habilitará un libro de '
        'reclamos en obra y toda queja será atendida en un plazo máximo de 48 horas.')

    # ------------------------------ 8. PRESUPUESTO ------------------------------
    D.h1("8. PRESUPUESTO, HORAS-MÁQUINA Y CRONOGRAMA DE EJECUCIÓN")
    D.h2("8.1 Presupuesto de demolición y eliminación")
    filas_p = []
    for s in pres["secciones"]:
        filas_p.append([s["cod"], s["titulo"], "", "", "", soles(s["subtotal"])])
        for it in s["items"]:
            filas_p.append([it["item"], cortar(it["desc"], 60), it["und"], fmt(it["metrado"]), fmt(it["cu"]),
                            fmt(it["parcial"])])
    D.tabla(["Item", "Descripción", "Und", "Metrado", "C.U. S/", "Parcial S/"], filas_p,
            anchos=[1.4, 8.2, 1.2, 1.9, 1.7, 2.3], tam=7.5, color="375623")
    D.tabla(["Concepto", "Monto"], [
        ["Costo directo", soles(pres["cd"])],
        ["Gastos generales (10 %)", soles(pres["gg"])],
        ["Utilidad (8 %)", soles(pres["ut"])],
        ["Sub total", soles(pres["subtotal"])],
        ["IGV (18 %)", soles(pres["igv"])],
        ["PRESUPUESTO TOTAL", soles(pres["total"])],
    ], anchos=[10, 6.5], color="375623", total=True)
    D.p(f"Son: {numero_a_letras(pres['total'])}", bold=True)
    D.h2("8.2 Horas-máquina de equipos")
    D.tabla(["Equipo", "H-M", "Tarifa S/ h", "Parcial S/", "Observación"],
            [[e["eq"], fmt(e["h"]), fmt(e["tarifa"]), fmt(e["parcial"]), e["nota"]] for e in res["hm"]],
            anchos=[6, 1.6, 2, 2.4, 4.5], tam=8, color="375623")
    D.h2("8.3 Cronograma de ejecución")
    D.tabla(["Fase", "Actividad", "Inicio (día)", "Fin (día)", "Duración"],
            [[a, b, str(c), str(e), f"{e - c + 1} día(s)"] for a, b, c, e in res["cron"]],
            anchos=[2.2, 8.3, 2, 2, 2], tam=8)
    D.p(f'Plazo total estimado de ejecución: {res["plazo"]} días calendario, con {res["personal"]} trabajadores '
        f'({int(fnum(d.get("cuadrillas"), 2))} cuadrillas de demolición más personal de apoyo y vigías).', bold=True)
    D.salto()

    # ------------------------------ 9. CONCLUSIONES ------------------------------
    D.h1("9. CONCLUSIONES, RECOMENDACIONES Y RESPONSABILIDAD PROFESIONAL")
    D.numeros([
        f'La demolición generará un volumen en banco de {res["vb"]:,.2f} m³ que, esponjado, representa '
        f'{res["ve"]:,.2f} m³ de RCD con un peso aproximado de {res["peso"]:,.2f} t.',
        f'La eliminación requiere {res["viajes"]} viajes de volquete de {res["cap"]} m³ con una flota de '
        f'{fl["n_vol"]} volquete(s), en {fl["dias_elim"]} días efectivos de eliminación.',
        f'El plazo total estimado es de {res["plazo"]} días calendario y el presupuesto total asciende a '
        f'{soles(pres["total"])} incluido IGV.',
        "La demolición se ejecutará en orden descendente estricto según la Norma G.050, con picado manual en la "
        "franja de 1.00 m colindante a los medianeros y apuntalamiento preventivo de los muros vecinos.",
        "No se iniciará ningún trabajo sin el corte verificado de suministros, el acta de constatación de "
        "colindancias, la licencia municipal vigente y la lista de verificación G.050 conforme.",
        "Los RCD se dispondrán exclusivamente en escombreras autorizadas, con manifiestos de manejo de residuos.",
    ])
    D.p("")
    D.p('El suscrito, Ingeniero Civil colegiado y habilitado, asume la responsabilidad técnica de la presente '
        'memoria y de la dirección de los trabajos de demolición conforme a la Ley N° 29090 y su reglamento.')
    for _ in range(4):
        D.p("")
    D.p("_______________________________________", align="c")
    D.p(f'{d.get("ingeniero", "")}\nIngeniero Civil - {d.get("cip", "")}', align="c", bold=True)
    D.p(f'{d.get("municipalidad", "")}, {hoy}', align="c")
    D.salto()

    # ------------------------------ 10. ANEXO FOTOGRÁFICO ------------------------------
    D.h1("10. ANEXO FOTOGRÁFICO DEL INMUEBLE Y COLINDANCIAS")
    D.p('Registro fotográfico del estado actual de la edificación a demoler y de las edificaciones colindantes, '
        'que forma parte del Acta de Constatación previa al inicio de los trabajos.')
    lista = list(fotos) if fotos else [(c, None) for c in CATEGORIAS_FOTO[:4]]
    for i, (cat, ruta_f) in enumerate(lista, 1):
        try:
            flujo = preparar_foto(ruta_f) if ruta_f else imagen_sin_foto(cat)
        except Exception:
            flujo = imagen_sin_foto(cat)
        D.imagen(flujo, 12.5, f"Foto N° {i:02d}. {cat}" + (f" - {os.path.basename(ruta_f)}" if ruta_f else ""))
        if i % 2 == 0 and i < len(lista):
            D.salto()

    D.doc.core_properties.title = "Memoria de Demolición y Plan de Seguridad G.050"
    D.doc.core_properties.author = d.get("ingeniero", "")
    D.doc.save(ruta)
    return ruta


# =============================================================================
# 8. EXPORTACIÓN A AUTOCAD (.DXF) - LÁMINA D-01 (ezdxf)
#    Papel A1 (841 x 594 mm) dibujado en milímetros de papel.
#    Capas: D-CUADROS, D-DETALLES, D-SEGURIDAD, D-TEXTOS
# =============================================================================
_ALINEAR = {
    "ML": TextEntityAlignment.MIDDLE_LEFT, "MC": TextEntityAlignment.MIDDLE_CENTER,
    "MR": TextEntityAlignment.MIDDLE_RIGHT, "BL": TextEntityAlignment.BOTTOM_LEFT,
    "BC": TextEntityAlignment.BOTTOM_CENTER, "TL": TextEntityAlignment.TOP_LEFT,
    "TC": TextEntityAlignment.TOP_CENTER,
}


class _Lamina:
    CAPAS = {"D-CUADROS": 5, "D-DETALLES": 3, "D-SEGURIDAD": 1, "D-TEXTOS": 7}

    def __init__(self):
        self.doc = ezdxf.new("R2010", setup=True)
        self.doc.header["$INSUNITS"] = 4          # milímetros
        self.doc.header["$LTSCALE"] = 4.0
        self.doc.header["$MEASUREMENT"] = 1
        for nombre, color in self.CAPAS.items():
            self.doc.layers.add(nombre, color=color)
        if "ARIAL_D" not in self.doc.styles:
            self.doc.styles.new("ARIAL_D", dxfattribs={"font": "arial.ttf", "width": 0.85})
        self.msp = self.doc.modelspace()

    # ---- primitivas -----------------------------------------------------------
    def _attr(self, capa, lt=None, lw=None, color=None):
        a = {"layer": capa}
        if lt:
            a["linetype"] = lt
        if lw:
            a["lineweight"] = lw
        if color is not None:
            a["color"] = color
        return a

    def linea(self, p1, p2, capa="D-DETALLES", lt=None, lw=None, color=None):
        self.msp.add_line(p1, p2, dxfattribs=self._attr(capa, lt, lw, color))

    def poli(self, pts, capa="D-DETALLES", cerrado=False, lt=None, lw=None, color=None):
        self.msp.add_lwpolyline(pts, close=cerrado, dxfattribs=self._attr(capa, lt, lw, color))

    def rect(self, x, y, w, h, capa="D-DETALLES", lt=None, lw=None, color=None):
        self.poli([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], capa, True, lt, lw, color)

    def circulo(self, c, r, capa="D-DETALLES", lt=None):
        self.msp.add_circle(c, r, dxfattribs=self._attr(capa, lt))

    def arco(self, c, r, a1, a2, capa="D-DETALLES"):
        self.msp.add_arc(c, r, a1, a2, dxfattribs=self._attr(capa))

    def achurado(self, pts, patron, escala=1.0, capa="D-DETALLES", angulo=0.0, color=None):
        col = color if color is not None else self.CAPAS.get(capa, 7)
        h = self.msp.add_hatch(color=col, dxfattribs={"layer": capa})
        if patron == "SOLID":
            h.set_solid_fill(color=col)
        else:
            h.set_pattern_fill(patron, color=col, scale=escala, angle=angulo)
        h.paths.add_polyline_path(pts, is_closed=True)
        return h

    def achurado_rect(self, x, y, w, h, patron, escala=1.0, capa="D-DETALLES", angulo=0.0, color=None):
        return self.achurado([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], patron, escala, capa, angulo, color)

    def texto(self, t, x, y, h=2.0, capa="D-TEXTOS", al="ML", rot=0.0, color=None):
        a = {"layer": capa, "style": "ARIAL_D", "height": h, "rotation": rot}
        if color is not None:
            a["color"] = color
        e = self.msp.add_text(str(t), dxfattribs=a)
        e.set_placement((x, y), align=_ALINEAR[al])
        return e

    def mtexto(self, t, x, y, h, ancho, capa="D-TEXTOS"):
        m = self.msp.add_mtext(str(t).replace("\n", "\\P"),
                               dxfattribs={"layer": capa, "style": "ARIAL_D", "char_height": h, "width": ancho})
        m.set_location((x, y), attachment_point=1)
        return m

    def flecha(self, p1, p2, capa="D-SEGURIDAD", tam=3.0, lw=None):
        self.linea(p1, p2, capa, lw=lw)
        ang = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
        c, s = math.cos(ang), math.sin(ang)
        a = (p2[0] - tam * c + tam * 0.35 * s, p2[1] - tam * s - tam * 0.35 * c)
        b = (p2[0] - tam * c - tam * 0.35 * s, p2[1] - tam * s + tam * 0.35 * c)
        self.msp.add_solid([p2, a, b], dxfattribs={"layer": capa})

    def burbuja(self, c, num, r=3.0, capa="D-SEGURIDAD"):
        self.circulo(c, r, capa)
        self.texto(str(num), c[0], c[1], r * 1.1, capa, "MC")

    def rotulo(self, p, q, txt, h=1.8, capa="D-TEXTOS"):
        """Línea guía desde el punto p del elemento hasta el texto en q."""
        self.circulo(p, 0.45, capa)
        self.linea(p, q, capa)
        lado = 1 if q[0] >= p[0] else -1
        self.linea(q, (q[0] + lado * 2.5, q[1]), capa)
        self.texto(txt, q[0] + lado * 3.3, q[1], h, capa, "ML" if lado > 0 else "MR")

    def cota_h(self, x1, x2, y, txt, h=1.8, capa="D-TEXTOS", y_ref=None):
        self.linea((x1, y), (x2, y), capa)
        for x in (x1, x2):
            self.linea((x - 0.9, y - 0.9), (x + 0.9, y + 0.9), capa)
            if y_ref is not None:
                self.linea((x, y_ref), (x, y - (1.2 if y_ref > y else -1.2)), capa)
        self.texto(txt, (x1 + x2) / 2, y + 0.9, h, capa, "BC")

    def cota_v(self, x, y1, y2, txt, h=1.8, capa="D-TEXTOS", x_ref=None):
        self.linea((x, y1), (x, y2), capa)
        for y in (y1, y2):
            self.linea((x - 0.9, y - 0.9), (x + 0.9, y + 0.9), capa)
            if x_ref is not None:
                self.linea((x_ref, y), (x + (1.2 if x_ref < x else -1.2), y), capa)
        self.texto(txt, x - 0.9, (y1 + y2) / 2, h, capa, "BC", rot=90)

    def titulo(self, t, x, y, h=3.2, ancho=None):
        self.texto(t, x, y, h, "D-TEXTOS", "BC")
        w = ancho if ancho else len(t) * h * 0.72
        self.linea((x - w / 2, y - 1.2), (x + w / 2, y - 1.2), "D-TEXTOS", lw=35)
        self.linea((x - w / 2, y - 2.0), (x + w / 2, y - 2.0), "D-TEXTOS")

    def tabla(self, x, y_top, anchos, filas, rh=6.0, h=1.8, encabezado=True, alinear=None, capa="D-CUADROS"):
        """Dibuja una tabla: filas = lista de listas. Devuelve la coordenada y inferior."""
        tot = sum(anchos)
        n = len(filas)
        y = y_top
        for i, fila in enumerate(filas):
            y2 = y - rh
            xi = x
            for j, w in enumerate(anchos):
                v = fila[j] if j < len(fila) else ""
                if v is None:
                    xi += w
                    continue
                al = (alinear[j] if alinear else "L") if not (encabezado and i == 0) else "C"
                if al == "R":
                    self.texto(v, xi + w - 1.2, (y + y2) / 2, h, "D-TEXTOS", "MR")
                elif al == "C":
                    self.texto(v, xi + w / 2, (y + y2) / 2, h, "D-TEXTOS", "MC")
                else:
                    self.texto(v, xi + 1.2, (y + y2) / 2, h, "D-TEXTOS", "ML")
                xi += w
            y = y2
        self.rect(x, y, tot, y_top - y, capa, lw=35)
        for i in range(1, n):
            yy = y_top - i * rh
            self.linea((x, yy), (x + tot, yy), capa, lw=35 if (encabezado and i == 1) else None)
        xi = x
        for w in anchos[:-1]:
            xi += w
            self.linea((xi, y_top), (xi, y), capa)
        return y

    def guardar(self, ruta):
        self.doc.set_modelspace_vport(height=620, center=(420.5, 297))
        self.doc.saveas(ruta)


def exportar_dxf(datos, res, ruta):
    L = _Lamina()
    d = datos
    fl = res["flota"]
    pres = res["presupuesto"]
    n_pisos = min(max(int(fnum(d.get("n_pisos"), 2)), 1), 8)

    # ------------------------------------------------------------------ MARCO
    L.rect(0, 0, 841, 594, "D-CUADROS")
    L.rect(20, 10, 811, 574, "D-CUADROS", lw=70)
    L.linea((20, 330), (831, 330), "D-CUADROS", lw=50)
    L.linea((400, 330), (400, 584), "D-CUADROS", lw=50)
    L.linea((620, 330), (620, 584), "D-CUADROS", lw=50)
    L.linea((200, 10), (200, 330), "D-CUADROS", lw=50)
    L.linea((562, 10), (562, 330), "D-CUADROS", lw=50)
    L.linea((20, 168), (200, 168), "D-CUADROS", lw=50)
    L.linea((562, 222), (831, 222), "D-CUADROS", lw=50)
    L.linea((562, 132), (831, 132), "D-CUADROS", lw=50)

    # ============================================================ 1. ESQUEMA
    L.titulo("ESQUEMA EN CORTE A-A: SECUENCIA DE DEMOLICIÓN DESCENDENTE (FASE 1 A FASE 6)", 210, 574, 3.0, 370)
    s = min(11.0, 282.0 / 25.4, 165.0 / (n_pisos * 2.8 + 3.8))
    g = 378.0
    xb = 26.0 + 9.9 * s
    W = 12.0
    X = lambda m: xb + m * s
    Y = lambda m: g + m * s
    hp = 2.8
    top = n_pisos * hp
    x_pl = X(-0.3)
    x_ver = x_pl - 1.8 * s
    # terreno, vereda y vía
    L.linea((24, g), (X(W) + 3.5 * s, g), "D-DETALLES", lw=50)
    for k in range(14):
        xx = 26 + k * ((X(W) + 3.5 * s - 30) / 14)
        L.linea((xx, g), (xx - 2, g - 2), "D-DETALLES")
    L.rect(x_ver, g, 1.8 * s, 0.15 * s, "D-DETALLES")
    L.texto("VEREDA", x_ver + 0.9 * s, g - 3.2, 1.6, al="MC")
    L.texto("VÍA PÚBLICA", (24 + x_ver) / 2, g - 3.2, 1.8, al="MC")
    L.linea((x_pl, g - 2 * s), (x_pl, Y(top + 1.6)), "D-TEXTOS", lt="DASHDOT")
    L.texto("L.P.", x_pl, g - 2 * s - 1.0, 1.5, al="TC")
    # cerco provisional H=2.40 y bandeja peatonal
    L.rect(x_pl - 0.06 * s, g, 0.12 * s, 2.4 * s, "D-SEGURIDAD", lw=35)
    L.achurado_rect(x_pl - 0.06 * s, g, 0.12 * s, 2.4 * s, "SOLID", capa="D-SEGURIDAD")
    L.poli([(x_ver, Y(2.6)), (x_pl, Y(2.75))], "D-SEGURIDAD", lw=35)
    L.linea((x_ver + 0.1 * s, g + 0.15 * s), (x_ver + 0.1 * s, Y(2.6)), "D-SEGURIDAD")
    L.rotulo((x_pl, Y(1.5)), (x_ver - 3, Y(4.6)), "CERCO PROVISIONAL H=2.40 m", 1.5)
    L.rotulo((x_ver + 0.9 * s, Y(2.67)), (x_ver - 3, Y(3.8)), "BANDEJA PEATONAL", 1.5)
    # volquete
    xt = x_ver - 7.8 * s
    T = lambda m: xt + m * s
    for cx in (1.2, 5.1, 6.4):
        L.circulo((T(cx), Y(0.5)), 0.5 * s, "D-DETALLES")
        L.circulo((T(cx), Y(0.5)), 0.22 * s, "D-DETALLES")
    L.rect(T(0), Y(0.75), 7.5 * s, 0.35 * s, "D-DETALLES")
    L.poli([(T(0.1), Y(1.1)), (T(0.1), Y(2.6)), (T(0.6), Y(3.05)), (T(1.9), Y(3.05)), (T(1.9), Y(1.1))],
           "D-DETALLES", cerrado=True)
    L.rect(T(0.7), Y(2.15), 1.0 * s, 0.7 * s, "D-DETALLES")
    L.poli([(T(2.1), Y(1.1)), (T(2.0), Y(2.8)), (T(7.5), Y(2.8)), (T(7.4), Y(1.1))], "D-DETALLES", cerrado=True)
    L.texto(f"VOLQUETE {res['cap']} m³", T(3.75), Y(-0.65), 1.8, al="MC")
    # edificación existente
    Yl = lambda k: Y(k * hp)
    cols = [0.0, 4.0, 8.0, W - 0.25]
    for k in range(1, n_pisos + 1):
        L.rect(X(0), Yl(k) - 0.2 * s, W * s, 0.2 * s, "D-DETALLES", lt="DASHED")
        L.achurado_rect(X(0), Yl(k) - 0.2 * s, W * s, 0.2 * s, "ANSI37", 0.22)
        L.rect(X(0), Yl(k) - 0.5 * s, W * s, 0.3 * s, "D-DETALLES", lt="DASHED")
        for c in cols:
            L.rect(X(c), Yl(k - 1), 0.25 * s, hp * s - 0.5 * s, "D-DETALLES", lt="DASHED")
            L.achurado_rect(X(c), Yl(k - 1), 0.25 * s, hp * s - 0.5 * s, "ANSI33", 0.18)
        for a, b in ((0.25, 4.0), (8.25, W - 0.25)):
            L.rect(X(a), Yl(k - 1), (b - a) * s, (hp - 0.5) * s, "D-DETALLES", lt="DASHED")
            L.achurado_rect(X(a), Yl(k - 1), (b - a) * s, (hp - 0.5) * s, "ANSI31", 0.55)
    # parapetos y cobertura liviana en azotea
    for a in (0.0, W - 0.15):
        L.rect(X(a), Yl(n_pisos), 0.15 * s, 1.0 * s, "D-DETALLES", lt="DASHED")
        L.achurado_rect(X(a), Yl(n_pisos), 0.15 * s, 1.0 * s, "ANSI31", 0.3)
    yz = Yl(n_pisos)
    L.linea((X(8.0), yz), (X(8.0), yz + 2.2 * s), "D-DETALLES", lt="DASHED")
    L.linea((X(W - 0.3), yz), (X(W - 0.3), yz + 2.2 * s), "D-DETALLES", lt="DASHED")
    L.poli([(X(7.8), yz + 2.2 * s), (X(9.9), yz + 2.9 * s), (X(W + 0.1), yz + 2.2 * s)], "D-DETALLES", lt="DASHED")
    L.poli([(X(8.0), yz + 2.2 * s), (X(W - 0.3), yz + 2.2 * s)], "D-DETALLES", lt="DASHED")
    for xm in (8.9, 9.9, 10.9):
        L.linea((X(xm), yz + 2.2 * s), (X(9.9), yz + 2.9 * s), "D-DETALLES", lt="DASHED")
    # cimentación
    for c in cols:
        L.rect(X(c), Y(-1.0), 0.25 * s, 1.0 * s, "D-DETALLES", lt="DASHED")
        L.rect(X(c + 0.125 - 0.5), Y(-1.6), 1.0 * s, 0.6 * s, "D-DETALLES", lt="DASHED")
        L.achurado_rect(X(c + 0.125 - 0.5), Y(-1.6), 1.0 * s, 0.6 * s, "DOTS", 0.35)
    for a, b in ((0.25 + 0.5, 4.0 - 0.4), (8.25 + 0.5, W - 0.25 - 0.4)):
        L.rect(X(a), Y(-0.8), (b - a) * s, 0.8 * s, "D-DETALLES", lt="DASHED")
        L.achurado_rect(X(a), Y(-0.8), (b - a) * s, 0.8 * s, "DOTS", 0.35)
    L.linea((X(0), Y(0.1)), (X(W), Y(0.1)), "D-DETALLES", lt="DASHED")
    # edificación colindante (a conservar)
    n_v = min(n_pisos + 1, 3)
    xv = X(W)
    L.rect(xv, g, 3.5 * s, n_v * hp * s, "D-DETALLES", lw=35)
    L.achurado_rect(xv, g, 0.25 * s, n_v * hp * s, "ANSI32", 0.4)
    for k in range(1, n_v):
        L.linea((xv, Y(k * hp)), (xv + 3.5 * s, Y(k * hp)), "D-DETALLES")
    L.texto("EDIFICACIÓN COLINDANTE (A CONSERVAR)", xv + 2.6 * s, g + n_v * hp * s / 2, 1.6, al="MC", rot=90)
    for k in range(0, n_pisos + 1):
        yk = Yl(k)
        L.poli([(xv + 0.5 * s, yk), (xv + 0.5 * s + 1.2, yk + 1.6), (xv + 0.5 * s - 1.2, yk + 1.6)], "D-TEXTOS", True)
        L.texto(f"N+{k * hp:.2f}", xv + 0.5 * s + 1.6, yk + 1.0, 1.3, al="ML")
    L.linea((xv - 0.15 * s, g), (xv - 0.15 * s, Y(4.0)), "D-SEGURIDAD", lt="DASHED")
    # manga / chute de evacuación
    xc1, xc2 = X(-0.75), X(-0.25)
    y_c_ini, y_c_fin = Yl(n_pisos) + 0.3 * s, Y(3.6)
    nseg = max(int((y_c_ini - y_c_fin) / (1.0 * s)), 1)
    dy = (y_c_ini - y_c_fin) / nseg
    for i in range(nseg):
        ya = y_c_ini - i * dy
        L.poli([(xc1, ya), (xc2, ya), (xc2 - 0.05 * s, ya - dy), (xc1 + 0.05 * s, ya - dy)], "D-SEGURIDAD", True)
    xbed = T(4.8)
    L.poli([(xc1 + 0.05 * s, y_c_fin), (xc2 - 0.05 * s, y_c_fin), (xbed + 0.35 * s, Y(3.15)),
            (xbed - 0.35 * s, Y(3.0))], "D-SEGURIDAD", True)
    L.flecha((xbed + 0.3 * s, Y(3.0)), (xbed, Y(2.4)), "D-SEGURIDAD", 2.0)
    yq = max(Yl(n_pisos) + 1.0 * s, Y(6.0))
    L.rotulo(((xc1 + xc2) / 2, Yl(n_pisos) - 0.5 * s), (x_ver - 3, yq), "MANGA / CHUTE DE EVACUACIÓN DE DESMONTE", 1.5)
    L.texto("(descarga hacia la tolva del volquete)", x_ver - 6.3, yq - 3.0, 1.4, al="MR")
    # burbujas de fases
    tz = Yl(n_pisos)
    L.burbuja((X(9.9), tz + 3.6 * s + 3.0), 1)
    L.linea((X(9.9), tz + 3.6 * s), (X(9.9), tz + 2.9 * s), "D-SEGURIDAD")
    y_b = tz - 0.5 * s - 0.45 * s - 2.0
    L.burbuja((X(5.0), y_b), 2)
    L.linea((X(5.0), y_b + 3.0), (X(5.0), tz - 0.1 * s), "D-SEGURIDAD")
    L.burbuja((X(7.0), y_b), 3)
    L.linea((X(7.0), y_b + 3.0), (X(7.0), tz - 0.35 * s), "D-SEGURIDAD")
    y_b2 = Yl(n_pisos - 1) + 0.9 * s
    L.burbuja((X(5.3), y_b2), 4)
    L.linea((X(5.3) - 3.0, y_b2), (X(3.0), y_b2), "D-SEGURIDAD")
    L.burbuja((X(6.9), y_b2), 5)
    L.linea((X(6.9) + 3.0, y_b2), (X(8.1), y_b2), "D-SEGURIDAD")
    L.burbuja((X(6.0), Y(-0.9)), 6)
    L.linea((X(6.0) + 3.0, Y(-0.9)), (X(7.7), Y(-1.3)), "D-SEGURIDAD")
    if n_pisos >= 2:
        L.mtexto("Fases 2 a 5 se repiten en cada nivel inferior", X(4.35), Yl(1) - 0.8 * s, 1.4, 3.3 * s,
                 "D-SEGURIDAD")
    L.texto("ESC. 1/100 (ESQUEMÁTICO)", 210, 336, 2.0, al="MC")
    # leyenda de fases (lado derecho)
    xa = 312.0
    L.flecha((xa, 566), (xa, 360), "D-SEGURIDAD", 4.0, lw=50)
    L.texto("SECUENCIA DESCENDENTE (G.050)", xa - 2.0, 463, 2.0, "D-SEGURIDAD", "BC", rot=90)
    L.texto("FASES DE DEMOLICIÓN", 356, 562, 2.2, al="MC")
    yy = 552.0
    for k in sorted(FASES):
        L.burbuja((320, yy), k, 2.6)
        L.texto(FASES[k][0].upper(), 324.5, yy + 0.8, 1.8, al="ML")
        L.mtexto(FASES[k][1] + f" ({res['dias_fase'].get(k, 0)} día/s)", 324.5, yy - 1.6, 1.6, 72)
        yy -= 30.0

    # ======================================================== 2. CERCO + RASCHEL
    L.titulo("DETALLE DE CERCO PERIMÉTRICO PROVISIONAL", 510, 574, 3.0, 200)
    L.texto("H = 2.40 m Y MALLA RASCHEL ANTIPOLVO H = 4.00 m (NORMA G.050)", 510, 566, 2.2, al="MC")
    s2 = 20.0
    x0, yb = 410.0, 478.0
    Xc = lambda m: x0 + m * s2
    L.linea((404, yb), (Xc(7.2) + 4, yb), "D-DETALLES", lw=50)
    for i in range(3):
        L.achurado_rect(Xc(i * 2.4 + 0.0375), yb, 2.4 * s2 - 1.5, 2.4 * s2, "ANSI31", 1.2, angulo=45.0)
        L.rect(Xc(i * 2.4 + 0.0375), yb, 2.4 * s2 - 1.5, 2.4 * s2, "D-DETALLES")
    for i in range(4):
        xp = Xc(i * 2.4) - 0.75
        L.rect(xp, yb, 1.5, 2.4 * s2, "D-DETALLES", lw=35)
        L.achurado_rect(xp, yb, 1.5, 2.4 * s2, "SOLID")
        L.rect(xp - 3.25, yb - 12, 8.0, 12.0, "D-DETALLES", lt="DASHED")
        L.achurado_rect(xp - 3.25, yb - 12, 8.0, 12.0, "DOTS", 0.3)
    for hm in (0.30, 1.20, 2.10):
        L.linea((Xc(0), yb + hm * s2), (Xc(7.2), yb + hm * s2), "D-DETALLES", lt="DASHED")
    L.cota_v(404.5, yb, yb + 2.4 * s2, "2.40", x_ref=Xc(0) - 1)
    L.cota_h(Xc(0), Xc(2.4), yb - 15.5, "2.40", y_ref=yb - 12.5)
    L.cota_h(Xc(2.4), Xc(4.8), yb - 15.5, "2.40", y_ref=yb - 12.5)
    L.rotulo((Xc(2.4), yb + 2.3 * s2), (Xc(2.4) + 8, 554), 'POSTE DE MADERA 3"x3" @ 2.40 m', 1.6)
    L.rotulo((Xc(3.6), yb + 1.8 * s2), (Xc(3.6) + 8, 546), "TRIPLAY 18 mm / CALAMINA GALV. PINTADA", 1.6)
    L.rotulo((Xc(5.8), yb + 2.1 * s2), (Xc(5.8) + 6, 538), 'TRAVESAÑOS 2"x3"', 1.6)
    L.rotulo((Xc(7.2) + 2, yb - 6), (Xc(7.2) + 5, yb - 17), "DADO C° 0.40x0.40x0.60", 1.5)
    L.texto("ELEVACIÓN FRONTAL DE CERCO - ESC. 1/50", Xc(3.6), 452, 2.0, al="MC")
    # Corte A-A del cerco con tornapunta
    xs = 588.0
    L.linea((566, yb), (618, yb), "D-DETALLES", lw=50)
    L.rect(xs - 0.75, yb, 1.5, 2.4 * s2, "D-DETALLES", lw=35)
    L.achurado_rect(xs - 0.75, yb, 1.5, 2.4 * s2, "SOLID")
    L.rect(xs - 4.0, yb - 12, 8.0, 12.0, "D-DETALLES", lt="DASHED")
    L.achurado_rect(xs - 4.0, yb - 12, 8.0, 12.0, "DOTS", 0.3)
    L.rect(xs - 2.5, yb, 1.75, 2.4 * s2, "D-DETALLES")
    L.linea((xs + 0.75, yb + 1.4 * s2), (xs + 0.75 + 1.4 * s2, yb), "D-DETALLES", lw=35)
    L.linea((xs + 0.75 + 1.4 * s2, yb), (xs + 0.75 + 1.4 * s2, yb - 8), "D-DETALLES", lt="DASHED")
    L.texto("CORTE A-A", xs + 6, 452, 2.0, al="MC")
    L.rotulo((xs + 0.75 + 0.4 * s2, yb + 1.0 * s2), (xs + 12, yb + 54), "TORNAPUNTA", 1.4)
    L.rotulo((xs + 0.75 + 1.4 * s2, yb - 6), (xs + 18, yb - 9), "ESTACA", 1.4)
    # Malla raschel H=4.00 m
    yr = 350.0
    L.linea((404, yr), (Xc(6.0) + 4, yr), "D-DETALLES", lw=50)
    L.achurado_rect(Xc(0), yr + 2, 6.0 * s2, 4.0 * s2 - 4, "NET", 1.2, "D-SEGURIDAD")
    L.rect(Xc(0), yr + 2, 6.0 * s2, 4.0 * s2 - 4, "D-SEGURIDAD")
    for i in range(3):
        xp = Xc(i * 3.0) - 0.6
        L.rect(xp, yr - 12, 1.2, 4.2 * s2 + 12, "D-DETALLES", lw=35)
        L.rect(xp - 2.4, yr - 12, 6.0, 12.0, "D-DETALLES", lt="DASHED")
    L.linea((Xc(0), yr + 4.0 * s2 - 2), (Xc(6.0), yr + 4.0 * s2 - 2), "D-SEGURIDAD", lw=35)
    L.linea((Xc(0), yr + 2), (Xc(6.0), yr + 2), "D-SEGURIDAD", lw=35)
    L.cota_v(404.5, yr, yr + 4.0 * s2, "4.00", x_ref=Xc(0) - 1)
    L.rotulo((Xc(5.2), yr + 4.0 * s2 - 2), (Xc(6.0) + 6, yr + 74), 'CABLE TENSOR 1/4" SUP. E INF.', 1.5)
    L.rotulo((Xc(6.0) + 0.6, yr + 3.3 * s2), (Xc(6.0) + 6, yr + 64), 'POSTE TUBO F.G. 2" @ 3.00 m', 1.5)
    L.rotulo((Xc(4.5), yr + 2.5 * s2), (Xc(6.0) + 6, yr + 50), "MALLA RASCHEL HDPE 80 % H=4.00 m", 1.5)
    L.rotulo((Xc(5.6), yr + 1.2 * s2), (Xc(6.0) + 6, yr + 36), "AMARRES ALAMBRE N°16 @ 0.30 m", 1.5)
    L.rotulo((Xc(6.0), yr - 6), (Xc(6.0) + 6, yr + 10), "DADO C° 0.30x0.30x0.60", 1.5)
    L.texto("ELEVACIÓN MALLA RASCHEL EN COLINDANCIAS - ESC. 1/50", Xc(3.6), 441, 2.0, al="MC")

    # ===================================================== 3. APUNTALAMIENTO
    L.titulo("DETALLE DE APUNTALAMIENTO PREVENTIVO", 725, 574, 3.0, 195)
    L.texto("DE MUROS COLINDANTES VECINOS - ESC. 1/50", 725, 566, 2.2, al="MC")
    s3 = 20.0
    g3 = 368.0
    xw = 642.0
    Ya = lambda m: g3 + m * s3
    L.linea((624, g3), (826, g3), "D-DETALLES", lw=50)
    for k in range(18):
        xx = 628 + k * 11
        L.linea((xx, g3), (xx - 2, g3 - 2), "D-DETALLES")
    L.rect(xw, Ya(-0.8), 0.25 * s3, 6.8 * s3, "D-DETALLES", lw=50)
    L.achurado_rect(xw, Ya(-0.8), 0.25 * s3, 6.8 * s3, "ANSI32", 0.5)
    L.rect(xw - 0.2 * s3, Ya(-1.2), 0.65 * s3, 0.4 * s3, "D-DETALLES", lt="DASHED")
    for k in (2.8, 5.6):
        L.rect(632, Ya(k) - 0.2 * s3, xw - 632, 0.2 * s3, "D-DETALLES")
    L.texto("VECINO", 636, Ya(4.2), 1.6, al="MC", rot=90)
    xs3 = xw + 0.25 * s3
    L.rect(xs3, g3, 2.0, 5.0 * s3, "D-SEGURIDAD", lw=35)
    L.linea((xs3, g3 + 1.0), (xs3 + 4.4 * s3 + 8, g3 + 1.0), "D-SEGURIDAD", lw=35)
    L.linea((xs3, g3 + 3.0), (xs3 + 4.4 * s3 + 8, g3 + 3.0), "D-SEGURIDAD", lw=35)
    for hh in (4.4, 2.2):
        p_sup = (xs3 + 2.0, Ya(hh))
        p_inf = (xs3 + 2.0 + hh * s3, g3 + 3.0)
        off = 1.41
        L.poli([(p_sup[0], p_sup[1] + off), (p_inf[0] + off, p_inf[1]), (p_inf[0] - off, p_inf[1]),
                (p_sup[0], p_sup[1] - off)], "D-SEGURIDAD", True, lw=35)
        L.rect(xs3 + 2.0, Ya(hh) - 2.5, 3.0, 5.0, "D-SEGURIDAD")
        L.poli([(p_inf[0] + off + 0.3, g3 + 3.0), (p_inf[0] + off + 3.3, g3 + 3.0), (p_inf[0] + off + 0.3, g3 + 5.5)],
               "D-SEGURIDAD", True)
        L.rect(p_inf[0] - 6.0, g3 - 8.0, 12.0, 11.0, "D-DETALLES", lw=35)
        L.achurado_rect(p_inf[0] - 6.0, g3 - 8.0, 12.0, 11.0, "DOTS", 0.3)
        L.linea((p_inf[0] + 4.0, g3 + 3.0), (p_inf[0] + 4.0, g3 - 14.0), "D-SEGURIDAD", lt="DASHED")
    foot = (xs3 + 2.0 + 4.4 * s3, g3 + 3.0)
    L.arco(foot, 14.0, 135.0, 180.0, "D-TEXTOS")
    L.texto("45°", foot[0] - 16.0, foot[1] + 5.0, 2.0, al="MC")
    L.cota_v(626.5, g3, Ya(4.4), "4.40", x_ref=xs3 + 1)
    L.cota_v(637.5, g3, Ya(2.2), "2.20")
    L.cota_h(xs3 + 2.0, foot[0], g3 - 17.0, "4.40 (45°)", y_ref=g3 - 9)
    lx = 760.0
    L.rotulo((xw + 2.5, Ya(6.3)), (lx, 520), "MURO COLINDANTE VECINO (A CONSERVAR)", 1.5)
    L.rotulo((xs3 + 1.0, Ya(4.9)), (lx, 510), 'SOLERA VERTICAL MADERA 2"x8"', 1.5)
    L.rotulo((xs3 + 4.5, Ya(4.4) + 2.5), (lx, 500), "TACO / TOPE DE MADERA", 1.5)
    L.rotulo((xs3 + 2.0 + 3.0 * s3, Ya(1.4) + 1.4), (lx, 450), 'TORNAPUNTA 4"x4" / PUNTAL METÁLICO A 45°', 1.5)
    L.rotulo((foot[0] + 3.0, g3 + 4.5), (lx, 430), "CUÑAS DE MADERA DURA", 1.5)
    L.rotulo((foot[0] + 5.0, g3 - 4.0), (lx, 410), "MUERTO DE C° f'c=140 kg/cm²", 1.5)
    L.texto("0.60 x 0.60 x 0.55 m", lx + 3.3, 406.5, 1.5, al="ML")
    L.rotulo((foot[0] + 14.0, g3 + 1.0), (lx, 390), 'SOLERA INFERIOR 2"x8" + ESTACAS Ø3/4"', 1.5)
    L.texto("TERRENO PROPIO (ZONA EN DEMOLICIÓN)", 790, g3 + 3.0, 1.5, al="MC")
    L.mtexto("ESPECIFICACIONES: tornapuntas a 45° separadas como máximo @ 2.40 m; madera tornillo seca; cuñas "
             "ajustadas diariamente; retirar el apuntalamiento solo cuando la nueva estructura confine el muro "
             "vecino. Picado manual sin percusión pesada en la franja de 1.00 m junto al medianero.",
             624, 346, 1.5, 202)

    # ========================================================== 4. LEYENDA
    L.titulo("LEYENDA OFICIAL DE SIMBOLOGÍA DE DEMOLICIÓN", 110, 320, 2.6, 170)
    yy = 307.0
    ley = [
        ("ANSI31", 0.55, "DASHED", "MURO DE ALBAÑILERÍA A DEMOLER (achurado diagonal, línea segmentada)"),
        ("ANSI37", 0.40, "DASHED", "LOSA ALIGERADA / MACIZA A DEMOLER (achurado cruzado)"),
        ("ANSI33", 0.40, "DASHED", "COLUMNA / VIGA / PLACA DE CONCRETO ARMADO A DEMOLER"),
        ("DOTS", 0.40, "DASHED", "CIMENTACIÓN Y PISOS DE CONCRETO A DEMOLER"),
        ("ANSI32", 0.50, None, "EDIFICACIÓN EXISTENTE / COLINDANTE A CONSERVAR"),
        ("NET", 1.00, None, "MALLA RASCHEL ANTIPOLVO H=4.00 m"),
    ]
    for pat, esc, lt, txt in ley:
        L.achurado_rect(26, yy - 4, 18, 8, pat, esc, "D-SEGURIDAD" if pat == "NET" else "D-DETALLES")
        L.rect(26, yy - 4, 18, 8, "D-SEGURIDAD" if pat == "NET" else "D-DETALLES", lt=lt)
        L.texto(txt, 48, yy, 1.7, al="ML")
        yy -= 12.0
    L.linea((26, yy), (44, yy), "D-DETALLES", lt="DASHED", lw=35)
    L.texto("LÍNEA DE ELEMENTO A DEMOLER (SEGMENTADA)", 48, yy, 1.7, al="ML")
    yy -= 10.0
    L.rect(26, yy - 1.5, 18, 3.0, "D-SEGURIDAD")
    L.achurado_rect(26, yy - 1.5, 18, 3.0, "SOLID", capa="D-SEGURIDAD")
    L.texto("CERCO PROVISIONAL DE OBRA H=2.40 m", 48, yy, 1.7, al="ML")
    yy -= 10.0
    L.poli([(26, yy - 3), (34, yy + 3), (44, yy + 3)], "D-SEGURIDAD", lw=35)
    L.texto("TORNAPUNTA / PUNTAL DE APUNTALAMIENTO", 48, yy, 1.7, al="ML")
    yy -= 10.0
    L.burbuja((35, yy), "n", 3.0)
    L.texto("NÚMERO DE FASE DE DEMOLICIÓN DESCENDENTE (1 a 6)", 48, yy, 1.7, al="ML")
    yy -= 10.0
    L.poli([(30, yy + 4), (34, yy + 4), (33.5, yy - 4), (30.5, yy - 4)], "D-SEGURIDAD", True)
    L.flecha((38, yy + 3), (42, yy - 3), "D-SEGURIDAD", 2.0)
    L.texto("MANGA / CHUTE DE EVACUACIÓN DE DESMONTE", 48, yy, 1.7, al="ML")

    # ========================================================== 5. NOTAS G.050
    L.titulo("NOTAS DE SEGURIDAD - NORMA G.050", 110, 160, 2.6, 170)
    notas_g050 = (
        "1. Ningún trabajo se iniciará sin licencia municipal, corte verificado de luz, agua, desagüe y gas, y "
        "acta de constatación de colindancias.\n"
        "2. Demolición en orden DESCENDENTE estricto, nivel por nivel; prohibido trabajar en niveles superpuestos "
        "y demoler por zapa o volteo.\n"
        "3. Franja de 1.00 m junto a medianeros vecinos: solo picado manual sin percusión pesada.\n"
        "4. Trabajos a más de 1.80 m con arnés de cuerpo completo, doble línea de anclaje y línea de vida "
        "certificada (2,270 kg por trabajador).\n"
        "5. EPP obligatorio: casco, lentes, protección auditiva doble, respirador P100/N95, guantes, calzado con "
        "punta de acero y chaleco reflectivo.\n"
        "6. Riego permanente con agua nebulizada para control de PM10; malla raschel H=4.00 m y tolvas cubiertas.\n"
        "7. Radio de exclusión de maquinaria = 1.5 veces el alcance del brazo; vigía permanente en vía pública.\n"
        "8. Prohibido acumular desmonte sobre losas y arrojarlo al vacío; evacuar por manga/chute.\n"
        "9. Horario: L-V 08:00-17:00 h, S 08:00-13:00 h (u ordenanza municipal).\n"
        "10. RCD a escombrera autorizada con manifiesto (D.S. 003-2013-VIVIENDA)."
    )
    L.mtexto(notas_g050, 24, 153, 2.0, 172)

    # ===================================================== 6. CUADRO METRADOS
    L.titulo("CUADRO GENERAL DE METRADOS DE DEMOLICIÓN Y ESPONJAMIENTO", 381, 321, 2.8, 330)
    L.texto("(RD N° 073-2010-VIVIENDA - RNM  /  FACTORES DE ESPONJAMIENTO CAPECO)", 381, 313, 1.8, al="MC")
    anchos = [9, 16, 148, 11, 25, 26, 14, 26, 26, 23]
    enc = ["N°", "CÓD.", "PARTIDA", "UND", "METRADO", "Vb (m³)", "FE", "Ve (m³)", "PESO (t)", "VIAJES"]
    por_cod = {}
    for r in res["filas"]:
        p = por_cod.setdefault(r["cod"], dict(m=0.0, vb=0.0, ve=0.0, peso=0.0))
        p["m"] += r["metrado"]
        p["vb"] += r["vb"]
        p["ve"] += r["ve"]
        p["peso"] += r["peso"]
    filas_t = [enc]
    for i, (cod, p) in enumerate(sorted(por_cod.items(), key=lambda kv: (CAT[kv[0]]["fase"], kv[0])), 1):
        c = CAT[cod]
        filas_t.append([str(i), cod, cortar(c["nombre"], 88), c["und"], fmt(p["m"]), fmt(p["vb"]),
                        f'{c["fe"]:.2f}', fmt(p["ve"]), fmt(p["peso"]), fmt(p["ve"] / res["cap"])])
    filas_t.append(["", "", "TOTAL GENERAL", "", "", fmt(res["vb"]), "", fmt(res["ve"]), fmt(res["peso"]),
                    str(res["viajes"])])
    rh = min(8.0, 205.0 / max(len(filas_t), 1))
    y_fin = L.tabla(206, 308, anchos, filas_t, rh, 1.7, True,
                    ["C", "C", "L", "C", "R", "R", "C", "R", "R", "R"])
    L.texto(f"N° de viajes = Ve / Capacidad ({res['cap']} m³), redondeado al entero superior = {res['viajes']} viajes.",
            206, y_fin - 4, 1.6, al="ML")
    y_r = y_fin - 9
    L.texto("RESUMEN DE PRESUPUESTO (S/)", 206, y_r - 2, 2.0, al="ML")
    filas_p = [["CONCEPTO", "MONTO (S/)"],
               ["Costo directo", fmt(pres["cd"])], ["Gastos generales (10 %)", fmt(pres["gg"])],
               ["Utilidad (8 %)", fmt(pres["ut"])], ["IGV (18 %)", fmt(pres["igv"])],
               ["PRESUPUESTO TOTAL", fmt(pres["total"])]]
    rh2 = min(6.0, max((y_r - 6 - 14) / len(filas_p), 3.0))
    L.tabla(206, y_r - 6, [110, 50], filas_p, rh2, 1.6, True, ["L", "R"])

    # ============================================ 7. CUADRO DE ÁREAS / FLOTA
    L.titulo("CUADRO GENERAL DE ÁREAS Y ELIMINACIÓN", 696, 323, 2.6, 240)
    filas_a = [
        ["Área de terreno", f'{fnum(d.get("area_terreno")):,.2f} m²'],
        ["Área techada a demoler", f'{fnum(d.get("area_techada")):,.2f} m²'],
        ["N° de pisos a demoler", f'{int(fnum(d.get("n_pisos"), 2))} + azotea'],
        ["Volumen total en banco (Vb)", f'{res["vb"]:,.2f} m³'],
        ["Volumen total esponjado (Ve)", f'{res["ve"]:,.2f} m³'],
        ["Peso total de escombros", f'{res["peso"]:,.2f} t'],
        ["Volquete / N° de viajes", f'{res["cap"]} m³ / {res["viajes"]} viajes'],
        ["Distancia a escombrera autorizada", f'{fl["dist"]:.1f} km'],
        ["Ciclo del volquete", f'{fl["ciclo"]:.1f} min'],
        ["Flota de volquetes asignada", f'{fl["n_vol"]} und'],
        ["Días efectivos de eliminación", f'{fl["dias_elim"]} días'],
        ["Plazo total de ejecución", f'{res["plazo"]} días calendario'],
    ]
    L.tabla(566, 317, [150, 111], filas_a, 7.4, 1.6, False, ["L", "R"])

    # ============================================ 8. NOTAS DE DEMOLICIÓN
    L.titulo("NOTAS DE DEMOLICIÓN", 696, 214, 2.6, 240)
    notas_dem = (
        f"1. Método: {d.get('metodo', '')}.\n"
        f"2. Sistema estructural existente: {d.get('sistema', '')}.\n"
        "3. Secuencia: desmontaje previo > techos > vigas > muros no portantes > muros portantes y columnas > "
        "cimentaciones.\n"
        "4. Apuntalar muros colindantes antes de retirar los elementos propios que los arriostran.\n"
        f"5. Escombrera autorizada: {cortar(d.get('escombrera', ''), 70)}.\n"
        "6. Metrados según RD N° 073-2010-VIVIENDA; Ve = Vb x FE; peso = Ve x peso específico suelto.\n"
        "7. Cotas en metros; el esquema de corte es referencial."
    )
    L.mtexto(notas_dem, 566, 206, 2.0, 260)

    # ============================================ 9. MEMBRETE MUNICIPAL
    L.texto(cortar(d.get("municipalidad", "MUNICIPALIDAD").upper(), 70), 696.5, 125, 2.8, al="MC")
    L.linea((562, 118), (831, 118), "D-CUADROS")
    campos = [("PROYECTO:", d.get("proyecto", "")), ("PROPIETARIO:", d.get("propietario", "")),
              ("UBICACIÓN:", d.get("ubicacion", "")), ("TRÁMITE:", d.get("tramite", ""))]
    yy = 118.0
    for k, v in campos:
        L.texto(k, 565, yy - 4, 1.6, al="ML")
        L.texto(cortar(v, 118), 590, yy - 4, 1.6, al="ML")
        yy -= 8.0
        L.linea((562, yy), (831, yy), "D-CUADROS")
    L.linea((745, 86), (745, 10), "D-CUADROS", lw=35)
    L.texto("PLANO:", 565, 81, 1.6, al="ML")
    L.texto("DEMOLICIÓN: SECUENCIA DESCENDENTE, APUNTALAMIENTOS,", 565, 75, 2.0, al="ML")
    L.texto("CERCO PROVISIONAL Y CUADRO DE METRADOS", 565, 70, 2.0, al="ML")
    L.linea((562, 64), (745, 64), "D-CUADROS")
    L.texto("ESCALA: INDICADA", 565, 59, 1.6, al="ML")
    L.texto(f"FECHA: {d.get('fecha', '')}", 650, 59, 1.6, al="ML")
    L.linea((562, 54), (745, 54), "D-CUADROS")
    L.texto("PROFESIONAL RESPONSABLE:", 565, 49, 1.6, al="ML")
    L.texto(cortar(d.get("ingeniero", ""), 60), 565, 43, 2.0, al="ML")
    L.texto(d.get("cip", ""), 565, 37, 2.0, al="ML")
    L.rect(690, 14, 52, 36, "D-CUADROS")
    L.texto("SELLO Y FIRMA", 716, 17, 1.5, al="MC")
    L.texto("LÁMINA", 788, 80, 2.2, al="MC")
    L.texto("D-01", 788, 64, 12.0, al="MC")
    L.linea((745, 50), (831, 50), "D-CUADROS")
    L.texto("N° EXPEDIENTE / LICENCIA:", 748, 45, 1.5, al="ML")
    L.texto("________________", 788, 36, 2.0, al="MC")
    L.texto("HOJA 1 DE 1", 788, 20, 2.0, al="MC")

    L.guardar(ruta)
    return ruta


# =============================================================================
# 9. INTERFAZ GRÁFICA (CustomTkinter - 1280 x 820 px)
# =============================================================================
COLOR_AZUL = ("#1F6AA5", "#1F6AA5")
COLOR_NARANJA = ("#D9730D", "#C8650A")
COLOR_VERDE = ("#2E8B57", "#257A4A")
COLOR_MORADO = ("#6A3FA0", "#5B3590")

CAMPOS_DATOS = [
    ("tramite", "Nombre del trámite", "combo", TRAMITES),
    ("proyecto", "Nombre del proyecto", "entry", None),
    ("propietario", "Propietario / Entidad", "entry", None),
    ("ubicacion", "Ubicación (dirección)", "entry", None),
    ("municipalidad", "Municipalidad / Entidad", "entry", None),
    ("n_pisos", "N° de pisos a demoler", "entry", None),
    ("area_terreno", "Área de terreno (m²)", "entry", None),
    ("area_techada", "Área techada a demoler (m²)", "entry", None),
    ("sistema", "Sistema estructural existente", "combo", SISTEMAS),
    ("metodo", "Método de demolición", "combo", METODOS),
    ("cap_volquete", "Capacidad de volquete", "combo", VOLQUETES),
    ("distancia", "Distancia a escombrera (km)", "entry", None),
    ("escombrera", "Escombrera / botadero autorizado", "entry", None),
    ("ingeniero", "Ingeniero Civil responsable", "entry", None),
    ("cip", "N° de colegiatura CIP", "entry", None),
    ("jornada", "Jornada de trabajo (h/día)", "entry", None),
    ("cuadrillas", "N° de cuadrillas de demolición", "entry", None),
    ("fecha", "Fecha del expediente", "entry", None),
]

COLUMNAS_TABLA = [
    ("piso", "Piso/Sector", 90), ("partida", "Tipo de Elemento / Partida", 230),
    ("desc", "Descripción", 250), ("n", "N° Veces", 60), ("largo", "Largo (m)", 70),
    ("ancho", "Ancho/Esp. (m)", 85), ("alto", "Alto (m)", 65), ("metrado", "Metrado Unidad", 105),
    ("vb", "Vol. Banco (m³)", 95), ("pct", "% Esponj.", 70), ("ve", "Vol. Esponjado (m³)", 115),
    ("viajes", "N° Viajes Volq.", 95),
]


class AppDemolicion(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITULO}  -  v{APP_VERSION}")
        self.geometry("1280x820")
        self.minsize(1100, 700)
        self.filas = []
        self.fotos = []
        self.res = None
        self.vars = {}
        self.check_vars = {}
        self._pendiente = None
        self._sel_index = None
        self._construir_cabecera()
        self.tabs = ctk.CTkTabview(self, corner_radius=10)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=(0, 4))
        self.tab1 = self.tabs.add("1. Datos de la Licencia / Expediente")
        self.tab2 = self.tabs.add("2. Metrado de Elementos a Demoler")
        self.tab3 = self.tabs.add("3. Dashboard de Escombros, Flota y G.050")
        self._construir_tab1()
        self._construir_tab2()
        self._construir_tab3()
        self._construir_barra_inferior()
        datos, _ = caso_ejemplo()
        for k, v in datos.items():
            if k in self.vars:
                self.vars[k].set(v)
        for k in ("proyecto", "propietario", "ubicacion"):
            self.vars[k].set("")
        for v in self.vars.values():
            v.trace_add("write", lambda *a: self.programar_recalculo())
        self.recalcular()

    # ------------------------------------------------------------------ cabecera
    def _construir_cabecera(self):
        fr = ctk.CTkFrame(self, height=56, corner_radius=0)
        fr.pack(fill="x")
        ctk.CTkLabel(fr, text="🏗  Generador Pro de Expedientes de Demolición",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=16, pady=10)
        ctk.CTkLabel(fr, text="Esponjamiento · Flota de Volquetes · Planos DXF · RNM · G.050 · Ley 29090",
                     font=ctk.CTkFont(size=12), text_color=("gray30", "gray70")).pack(side="left", padx=6)
        self.menu_tema = ctk.CTkSegmentedButton(fr, values=["Oscuro", "Claro"], command=self._cambiar_tema)
        self.menu_tema.set("Oscuro")
        self.menu_tema.pack(side="right", padx=16)
        ctk.CTkLabel(fr, text="Tema:").pack(side="right")

    def _cambiar_tema(self, valor):
        ctk.set_appearance_mode("dark" if valor == "Oscuro" else "light")
        self.after(80, self._estilo_tabla)
        self.after(120, self._dibujar_grafico)

    # ------------------------------------------------------------------ TAB 1
    def _construir_tab1(self):
        cont = ctk.CTkScrollableFrame(self.tab1, label_text="Datos generales del expediente y parámetros de eliminación")
        cont.pack(fill="both", expand=True, padx=6, pady=6)
        for c in (1, 3):
            cont.grid_columnconfigure(c, weight=1)
        for i, (clave, etiqueta, tipo, opciones) in enumerate(CAMPOS_DATOS):
            fila, col = divmod(i, 2)
            var = ctk.StringVar()
            self.vars[clave] = var
            ctk.CTkLabel(cont, text=etiqueta + ":", anchor="e", width=210).grid(
                row=fila, column=col * 2, sticky="e", padx=(8, 6), pady=7)
            if tipo == "combo":
                w = ctk.CTkComboBox(cont, values=opciones, variable=var, width=380)
            else:
                w = ctk.CTkEntry(cont, textvariable=var, width=380)
            w.grid(row=fila, column=col * 2 + 1, sticky="we", padx=(0, 14), pady=7)
        nota = ("Notas: el método de demolición define el % de picado manual (Mixto 30 %, Manual 100 %, Mecánico 0 %) y "
                "el factor del costo unitario. El ciclo del volquete considera ida cargado a 30 km/h, retorno vacío a "
                "45 km/h, descarga 4 min y maniobras 4 min; el carguío se realiza con cargador frontal o minicargador.")
        ctk.CTkLabel(cont, text=nota, wraplength=1100, justify="left", text_color=("gray25", "gray70")).grid(
            row=len(CAMPOS_DATOS) // 2 + 1, column=0, columnspan=4, sticky="w", padx=12, pady=(16, 4))

    # ------------------------------------------------------------------ TAB 2
    def _construir_tab2(self):
        barra = ctk.CTkFrame(self.tab2, fg_color="transparent")
        barra.pack(fill="x", padx=4, pady=(4, 2))
        ctk.CTkButton(barra, text="1. Importar Tabla de Elementos a Demoler desde Excel / Revit (.XLSX)",
                      fg_color=COLOR_AZUL, command=self.importar).pack(side="left", padx=3)
        ctk.CTkButton(barra, text="Cargar Caso Completo de Ejemplo (Demolición de Edificación de 2 Pisos + "
                                  "Azotea con 12 Partidas)", fg_color=COLOR_MORADO,
                      command=self.cargar_ejemplo).pack(side="left", padx=3)
        ctk.CTkButton(barra, text="Plantilla .XLSX", width=110, fg_color="gray40",
                      command=self.guardar_plantilla).pack(side="left", padx=3)

        form = ctk.CTkFrame(self.tab2)
        form.pack(fill="x", padx=4, pady=4)
        self.f_partida = ctk.StringVar(value=CAT_ETIQUETAS[0])
        self.f_campos = {k: ctk.StringVar() for k in ("piso", "desc", "n", "largo", "ancho", "alto", "md")}
        self.f_campos["n"].set("1")
        ctk.CTkLabel(form, text="Partida:").grid(row=0, column=0, padx=(8, 4), pady=4, sticky="e")
        ctk.CTkComboBox(form, values=CAT_ETIQUETAS, variable=self.f_partida, width=560,
                        command=lambda v: self._actualizar_ayuda()).grid(row=0, column=1, columnspan=5, sticky="w", pady=4)
        ctk.CTkLabel(form, text="Piso/Sector:").grid(row=0, column=6, padx=(10, 4), sticky="e")
        ctk.CTkEntry(form, textvariable=self.f_campos["piso"], width=120).grid(row=0, column=7, sticky="w")
        ctk.CTkLabel(form, text="Descripción:").grid(row=0, column=8, padx=(10, 4), sticky="e")
        ctk.CTkEntry(form, textvariable=self.f_campos["desc"], width=260).grid(row=0, column=9, columnspan=3,
                                                                               sticky="we", padx=(0, 8))
        etiquetas = [("n", "N° Veces"), ("largo", "Largo (m)"), ("ancho", "Ancho/Esp. (m)"), ("alto", "Alto (m)"),
                     ("md", "Metrado directo")]
        for j, (k, t) in enumerate(etiquetas):
            ctk.CTkLabel(form, text=t + ":").grid(row=1, column=j * 2, padx=(8, 4), pady=4, sticky="e")
            ctk.CTkEntry(form, textvariable=self.f_campos[k], width=80).grid(row=1, column=j * 2 + 1, sticky="w")
        acciones = ctk.CTkFrame(form, fg_color="transparent")
        acciones.grid(row=2, column=0, columnspan=12, sticky="w", padx=6, pady=(2, 2))
        ctk.CTkButton(acciones, text="➕ Agregar partida", width=130, fg_color=COLOR_VERDE,
                      command=self.agregar_fila).pack(side="left", padx=3)
        ctk.CTkButton(acciones, text="✎ Actualizar seleccionada", width=170,
                      command=self.actualizar_fila).pack(side="left", padx=3)
        ctk.CTkButton(acciones, text="🗑 Eliminar seleccionada", width=160, fg_color="#B03A2E",
                      command=self.eliminar_fila).pack(side="left", padx=3)
        ctk.CTkButton(acciones, text="Limpiar tabla", width=110, fg_color="gray40",
                      command=self.limpiar_tabla).pack(side="left", padx=3)
        ctk.CTkLabel(acciones, text="(Seleccione una fila para editarla en el formulario)",
                     text_color=("gray30", "gray65")).pack(side="left", padx=10)
        self.lbl_ayuda = ctk.CTkLabel(form, text="", text_color=("gray25", "gray70"), anchor="w",
                                      justify="left", wraplength=1200)
        self.lbl_ayuda.grid(row=3, column=0, columnspan=12, sticky="w", padx=10, pady=(0, 4))
        self._actualizar_ayuda()

        marco = ctk.CTkFrame(self.tab2)
        marco.pack(fill="both", expand=True, padx=4, pady=2)
        self.tree = ttk.Treeview(marco, columns=[c[0] for c in COLUMNAS_TABLA], show="headings",
                                 style="Demo.Treeview", selectmode="browse")
        for k, t, w in COLUMNAS_TABLA:
            self.tree.heading(k, text=t)
            self.tree.column(k, width=w, minwidth=50,
                             anchor="w" if k in ("piso", "partida", "desc") else "e", stretch=k in ("partida", "desc"))
        sy = ttk.Scrollbar(marco, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(marco, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        marco.grid_rowconfigure(0, weight=1)
        marco.grid_columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._seleccionar_fila)
        self.tree.bind("<Delete>", lambda e: self.eliminar_fila())
        self.lbl_totales = ctk.CTkLabel(self.tab2, text="", font=ctk.CTkFont(weight="bold"), anchor="w")
        self.lbl_totales.pack(fill="x", padx=10)

        ff = ctk.CTkFrame(self.tab2)
        ff.pack(fill="x", padx=4, pady=(2, 4))
        ctk.CTkLabel(ff, text="📷 Fotos del inmueble (Informe Técnico):",
                     font=ctk.CTkFont(weight="bold")).pack(side="left", padx=8)
        self.cat_foto = ctk.StringVar(value=CATEGORIAS_FOTO[0])
        ctk.CTkComboBox(ff, values=CATEGORIAS_FOTO, variable=self.cat_foto, width=190).pack(side="left", padx=4)
        ctk.CTkButton(ff, text="Cargar Fotos del Inmueble a Demoler", command=self.cargar_fotos).pack(side="left", padx=4)
        ctk.CTkButton(ff, text="Quitar fotos", width=100, fg_color="gray40",
                      command=self.quitar_fotos).pack(side="left", padx=4)
        self.lbl_fotos = ctk.CTkLabel(ff, text="Sin fotos cargadas (se insertarán recuadros para adjuntar).",
                                      anchor="w")
        self.lbl_fotos.pack(side="left", padx=8, fill="x", expand=True)
        self._estilo_tabla()

    def _estilo_tabla(self):
        oscuro = ctk.get_appearance_mode() == "Dark"
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        bg, fg, sel, hbg = (("#2B2B2B", "#E8E8E8", "#1F6AA5", "#1E3A5F") if oscuro
                            else ("#FFFFFF", "#1A1A1A", "#9CC3E6", "#1F4E78"))
        st.configure("Demo.Treeview", background=bg, fieldbackground=bg, foreground=fg, rowheight=25,
                     font=("Segoe UI", 10), borderwidth=0)
        st.configure("Demo.Treeview.Heading", background=hbg, foreground="white", font=("Segoe UI", 10, "bold"),
                     relief="flat")
        st.map("Demo.Treeview", background=[("selected", sel)], foreground=[("selected", "white")])
        st.map("Demo.Treeview.Heading", background=[("active", "#2E75B6")])
        if hasattr(self, "tree"):
            self.tree.tag_configure("par", background="#333333" if oscuro else "#F2F2F2")
            self.tree.tag_configure("total", background="#3A3A1E" if oscuro else "#FFF2CC",
                                    font=("Segoe UI", 10, "bold"))

    def _actualizar_ayuda(self):
        cod = self.f_partida.get().split(" | ")[0]
        c = CAT.get(cod)
        if not c:
            return
        txt = {
            "vol": "Metrado m³ = N° x Largo x Ancho x Alto.",
            "muro": f"Metrado m² = N° x Largo x Alto ; Vb = m² x espesor (Ancho/Esp.; si es 0 se usa e={c['esp']:.2f} m).",
            "planta": (f"Metrado m² = N° x Largo x Ancho ; Vb = m² x {c['esp']:.2f} m³/m²"
                       + (" (o Alto como espesor real si es > 0)." if c["usa_h"] else ".")),
            "und": f"Metrado und = N° Veces ; Vb = und x {c['esp']:.2f} m³/und.",
        }[c["modo"]]
        self.lbl_ayuda.configure(text=f"ℹ {txt}  FE = {c['fe']:.2f} ({(c['fe'] - 1) * 100:.0f} %) · "
                                      f"Peso suelto {c['peso']:.2f} t/m³ · C.U. base S/ {c['cu']:.2f} · "
                                      f"Equipo: {c['equipo']} · 'Metrado directo' > 0 reemplaza el cálculo.")

    # ------------------------------------------------------------------ TAB 3
    def _construir_tab3(self):
        tarjetas = ctk.CTkFrame(self.tab3, fg_color="transparent")
        tarjetas.pack(fill="x", padx=4, pady=4)
        defs = [("vb", "Volumen Total en Banco", "m³", "#2E75B6"),
                ("ve", "Volumen Total Esponjado", "m³", "#ED7D31"),
                ("peso", "Peso Total de Escombros", "t", "#7F6000"),
                ("viajes", "Total de Viajes de Volquete", "", "#C00000"),
                ("dias", "Días Estimados de Eliminación", "", "#548235"),
                ("total", "Presupuesto Total (inc. IGV)", "S/", "#375623")]
        self.cards = {}
        for i, (k, t, u, col) in enumerate(defs):
            fr = ctk.CTkFrame(tarjetas, corner_radius=12, border_width=2, border_color=col)
            fr.grid(row=i // 3, column=i % 3, sticky="nsew", padx=6, pady=6)
            tarjetas.grid_columnconfigure(i % 3, weight=1)
            lt = ctk.CTkLabel(fr, text=t, font=ctk.CTkFont(size=13, weight="bold"), text_color=col)
            lt.pack(pady=(10, 0))
            lv = ctk.CTkLabel(fr, text="0", font=ctk.CTkFont(size=30, weight="bold"))
            lv.pack(pady=(2, 0))
            ls = ctk.CTkLabel(fr, text=u, text_color=("gray30", "gray70"))
            ls.pack(pady=(0, 10))
            self.cards[k] = (lt, lv, ls)

        abajo = ctk.CTkFrame(self.tab3, fg_color="transparent")
        abajo.pack(fill="both", expand=True, padx=4, pady=4)
        abajo.grid_columnconfigure(0, weight=3)
        abajo.grid_columnconfigure(1, weight=2)
        abajo.grid_columnconfigure(2, weight=3)
        abajo.grid_rowconfigure(0, weight=1)
        fg = ctk.CTkFrame(abajo)
        fg.grid(row=0, column=0, sticky="nsew", padx=4)
        ctk.CTkLabel(fg, text="Volumen por tipo de material (Vb vs Ve)", font=ctk.CTkFont(weight="bold")).pack(pady=4)
        self.canvas = tk.Canvas(fg, highlightthickness=0, height=200)
        self.canvas.pack(fill="both", expand=True, padx=6, pady=6)
        self.canvas.bind("<Configure>", lambda e: self._dibujar_grafico())
        fr = ctk.CTkFrame(abajo)
        fr.grid(row=0, column=1, sticky="nsew", padx=4)
        ctk.CTkLabel(fr, text="Ciclo del volquete, flota y horas-máquina", font=ctk.CTkFont(weight="bold")).pack(pady=4)
        self.txt_resumen = ctk.CTkTextbox(fr, font=ctk.CTkFont(family="Consolas", size=12))
        self.txt_resumen.pack(fill="both", expand=True, padx=6, pady=6)
        fc = ctk.CTkScrollableFrame(abajo, label_text="Lista de verificación de Seguridad G.050")
        fc.grid(row=0, column=2, sticky="nsew", padx=4)
        for txt in CHECKLIST_G050:
            v = ctk.BooleanVar(value=False)
            self.check_vars[txt] = v
            ctk.CTkCheckBox(fc, text=txt, variable=v, command=self._actualizar_checklist).pack(anchor="w", pady=3, padx=4)
        self.lbl_check = ctk.CTkLabel(fc, text="", font=ctk.CTkFont(weight="bold"))
        self.lbl_check.pack(anchor="w", pady=6, padx=4)
        self._actualizar_checklist()

    def _actualizar_checklist(self):
        n = sum(v.get() for v in self.check_vars.values())
        tot = len(self.check_vars)
        color = "#2E8B57" if n == tot else ("#D9730D" if n else "#B03A2E")
        self.lbl_check.configure(text=f"Cumplimiento G.050: {n}/{tot} ({n / tot * 100:.0f} %)", text_color=color)

    def _dibujar_grafico(self):
        if not hasattr(self, "canvas"):
            return
        cv = self.canvas
        oscuro = ctk.get_appearance_mode() == "Dark"
        bg, fg, grid = ("#2B2B2B", "#E8E8E8", "#444444") if oscuro else ("#FFFFFF", "#1A1A1A", "#DDDDDD")
        cv.configure(bg=bg)
        cv.delete("all")
        W, H = max(cv.winfo_width(), 200), max(cv.winfo_height(), 150)
        grupos = list(self.res["grupos"].items()) if self.res else []
        if not grupos:
            cv.create_text(W / 2, H / 2, text="Sin datos: cargue o importe el metrado", fill=fg)
            return
        x0, y0, x1, y1 = 50, 15, W - 10, H - 52
        vmax = max(g["ve"] for _, g in grupos) or 1
        for k in range(5):
            y = y1 - (y1 - y0) * k / 4
            cv.create_line(x0, y, x1, y, fill=grid)
            cv.create_text(x0 - 6, y, text=f"{vmax * k / 4:.0f}", anchor="e", fill=fg, font=("Segoe UI", 8))
        ancho = (x1 - x0) / len(grupos)
        bw = min(ancho * 0.32, 40)
        for i, (nom, g) in enumerate(grupos):
            cx = x0 + ancho * (i + 0.5)
            for j, (v, col) in enumerate(((g["vb"], "#2E75B6"), (g["ve"], "#ED7D31"))):
                h = (y1 - y0) * v / vmax
                bx = cx - bw - 2 if j == 0 else cx + 2
                cv.create_rectangle(bx, y1 - h, bx + bw, y1, fill=col, outline="")
                cv.create_text(bx + bw / 2, y1 - h - 7, text=f"{v:.1f}", fill=fg, font=("Segoe UI", 7))
            cv.create_text(cx, y1 + (10 if i % 2 == 0 else 23), text=cortar(nom, 18), fill=fg, font=("Segoe UI", 8))
        cv.create_rectangle(x0, H - 12, x0 + 12, H - 2, fill="#2E75B6", outline="")
        cv.create_text(x0 + 16, H - 7, text="Vb banco", anchor="w", fill=fg, font=("Segoe UI", 8))
        cv.create_rectangle(x0 + 90, H - 12, x0 + 102, H - 2, fill="#ED7D31", outline="")
        cv.create_text(x0 + 106, H - 7, text="Ve esponjado", anchor="w", fill=fg, font=("Segoe UI", 8))

    # ------------------------------------------------------------------ barra inferior
    def _construir_barra_inferior(self):
        fr = ctk.CTkFrame(self, corner_radius=0)
        fr.pack(fill="x", side="bottom")
        fuente = ctk.CTkFont(size=12, weight="bold")
        botones = [
            ("1. Generar Memoria de Demolición + Plan de Seguridad G.050 + Mitigación en Word (.DOCX)",
             COLOR_AZUL, "docx"),
            ("2. Dibujar Lámina de Demolición, Secuencia, Apuntalamientos y Cerco en AutoCAD (.DXF)",
             COLOR_NARANJA, "dxf"),
            ("3. Exportar Planilla de Metrados de Demolición, Esponjamiento y Presupuesto en Excel (.XLSX)",
             COLOR_VERDE, "xlsx"),
        ]
        for i, (t, col, tipo) in enumerate(botones):
            b = ctk.CTkButton(fr, text=t, fg_color=col, height=48, font=fuente,
                              command=lambda x=tipo: self.exportar(x))
            b.grid(row=0, column=i, sticky="nsew", padx=6, pady=(8, 2))
            if hasattr(b, "_text_label"):           # permite texto en 2 líneas
                b._text_label.configure(wraplength=380)
            fr.grid_columnconfigure(i, weight=1, uniform="b")
        self.lbl_estado = ctk.CTkLabel(fr, text="Listo.", anchor="w", text_color=("gray25", "gray70"))
        self.lbl_estado.grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 4))

    def estado(self, txt):
        self.lbl_estado.configure(text=txt)
        self.update_idletasks()

    # ------------------------------------------------------------------ datos y cálculo
    def leer_datos(self):
        return {k: v.get() for k, v in self.vars.items()}

    def programar_recalculo(self):
        if self._pendiente:
            self.after_cancel(self._pendiente)
        self._pendiente = self.after(250, self.recalcular)

    def recalcular(self):
        self._pendiente = None
        try:
            self.res = calcular_todo(self.leer_datos(), self.filas)
        except Exception as exc:
            self.estado(f"Error de cálculo: {exc}")
            return
        self._refrescar_tabla()
        self._refrescar_dashboard()

    def _refrescar_tabla(self):
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.res["filas"]):
            vals = (r["piso"], f'{r["cod"]} {cortar(r["partida"], 60)}', r["desc"], fmt(fnum(r["n"], 1), 0),
                    fmt(fnum(r["largo"])), fmt(fnum(r["ancho"])), fmt(fnum(r["alto"])),
                    f'{fmt(r["metrado"])} {r["und"]}', fmt(r["vb"], 3), f'{r["pct"]:.0f} %', fmt(r["ve"], 3),
                    fmt(r["viajes"]))
            self.tree.insert("", "end", iid=str(i), values=vals, tags=("par",) if i % 2 else ())
        if self.res["filas"]:
            self.tree.insert("", "end", iid="total", tags=("total",), values=(
                "", "TOTAL GENERAL", f'{len(self.res["filas"])} partidas', "", "", "", "", "",
                fmt(self.res["vb"], 3), "", fmt(self.res["ve"], 3), str(self.res["viajes"])))
        self.lbl_totales.configure(
            text=f"Σ Vb = {self.res['vb']:,.2f} m³   |   Σ Ve = {self.res['ve']:,.2f} m³   |   "
                 f"Peso = {self.res['peso']:,.2f} t   |   N° viajes (volquete {self.res['cap']} m³) = "
                 f"⌈Ve / {self.res['cap']}⌉ = {self.res['viajes']}")

    def _refrescar_dashboard(self):
        r = self.res
        fl = r["flota"]
        vals = {"vb": (fmt(r["vb"]), "m³ (sólido en banco)"), "ve": (fmt(r["ve"]), "m³ (suelto a eliminar)"),
                "peso": (fmt(r["peso"]), "toneladas"),
                "viajes": (str(r["viajes"]), f"viajes de volquete de {r['cap']} m³"),
                "dias": (str(fl["dias_elim"]), f"días con {fl['n_vol']} volquete(s) · plazo total {r['plazo']} d"),
                "total": (f"{r['presupuesto']['total']:,.2f}", "Soles (S/) incluido IGV")}
        for k, (v, sub) in vals.items():
            self.cards[k][1].configure(text=v)
            self.cards[k][2].configure(text=sub)
        self.cards["viajes"][0].configure(text=f"Total de Viajes de Volquete de {r['cap']} m³")
        lineas = [
            f"Método ............ {r['metodo']['clave']} ({r['metodo']['fm'] * 100:.0f} % manual)",
            f"Equipo de carguío . {fl['cargador']}",
            f"Distancia ......... {fl['dist']:.1f} km",
            "",
            "CICLO DEL VOLQUETE (min)",
            f"  Carguío ......... {fl['t_carga']:7.1f}",
            f"  Ida (30 km/h) ... {fl['t_ida']:7.1f}",
            f"  Descarga ........ {fl['t_desc']:7.1f}",
            f"  Retorno (45 km/h) {fl['t_ret']:7.1f}",
            f"  Maniobras ....... {fl['t_man']:7.1f}",
            f"  TOTAL ........... {fl['ciclo']:7.1f}",
            "",
            f"Viajes/volquete/día {fl['viajes_vol_dia']}",
            f"Volquetes (óptimo)  {fl['n_vol']} ({fl['n_opt']})",
            f"Días eliminación .. {fl['dias_elim']}",
            f"Días demolición ... {r['dias_demol']}",
            f"Plazo total ....... {r['plazo']} días",
            "",
            "HORAS-MÁQUINA",
        ] + [f"  {cortar(e['eq'], 34):34s} {e['h']:7.1f} h" for e in r["hm"]] + [
            "",
            f"Costo directo ..... {soles(r['presupuesto']['cd'])}",
            f"PRESUPUESTO TOTAL . {soles(r['presupuesto']['total'])}",
        ]
        self.txt_resumen.configure(state="normal")
        self.txt_resumen.delete("1.0", "end")
        self.txt_resumen.insert("1.0", "\n".join(lineas))
        self.txt_resumen.configure(state="disabled")
        self._dibujar_grafico()

    # ------------------------------------------------------------------ edición de tabla
    def _fila_desde_form(self):
        cod = self.f_partida.get().split(" | ")[0].strip()
        if cod not in CAT:
            messagebox.showwarning("Partida", "Seleccione una partida válida del catálogo.")
            return None
        f = self.f_campos
        return nueva_fila(cod, f["piso"].get().strip(), f["desc"].get().strip(), fnum(f["n"].get(), 1) or 1,
                          fnum(f["largo"].get()), fnum(f["ancho"].get()), fnum(f["alto"].get()), fnum(f["md"].get()))

    def agregar_fila(self):
        fila = self._fila_desde_form()
        if fila:
            self.filas.append(fila)
            self.recalcular()
            self.estado(f"Partida agregada ({len(self.filas)} en total).")

    def actualizar_fila(self):
        if self._sel_index is None or self._sel_index >= len(self.filas):
            messagebox.showinfo("Actualizar", "Seleccione primero una fila de la tabla.")
            return
        fila = self._fila_desde_form()
        if fila:
            self.filas[self._sel_index] = fila
            self.recalcular()
            self.estado("Partida actualizada.")

    def eliminar_fila(self):
        if self._sel_index is None or self._sel_index >= len(self.filas):
            return
        del self.filas[self._sel_index]
        self._sel_index = None
        self.recalcular()
        self.estado("Partida eliminada.")

    def limpiar_tabla(self):
        if self.filas and messagebox.askyesno("Limpiar", "¿Eliminar todas las partidas de la tabla?"):
            self.filas.clear()
            self.recalcular()

    def _seleccionar_fila(self, _evento=None):
        sel = self.tree.selection()
        if not sel or sel[0] == "total":
            self._sel_index = None
            return
        i = int(sel[0])
        self._sel_index = i
        f = self.filas[i]
        c = CAT[f["cod"]]
        self.f_partida.set(f'{c["cod"]} | {c["nombre"]} [{c["und"]}]')
        for k in ("piso", "desc"):
            self.f_campos[k].set(f[k])
        for k in ("n", "largo", "ancho", "alto", "md"):
            self.f_campos[k].set(f"{fnum(f[k]):g}")
        self._actualizar_ayuda()

    # ------------------------------------------------------------------ acciones
    def cargar_ejemplo(self):
        datos, filas = caso_ejemplo()
        for k, v in datos.items():
            if k in self.vars:
                self.vars[k].set(v)
        self.filas = filas
        self.recalcular()
        self.tabs.set("2. Metrado de Elementos a Demoler")
        self.estado("Caso de ejemplo cargado: edificación de 2 pisos + azotea con 12 partidas.")

    def importar(self):
        ruta = filedialog.askopenfilename(title="Importar tabla de elementos a demoler",
                                          filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            filas, omitidas, mapa = importar_excel(ruta)
        except Exception as exc:
            messagebox.showerror("Importar", f"No se pudo leer el archivo:\n{exc}")
            return
        if not filas:
            messagebox.showwarning("Importar", "No se encontraron elementos reconocibles.\nColumnas detectadas: "
                                   f"{mapa}\nUse el botón 'Plantilla .XLSX' para ver el formato esperado.")
            return
        reemplazar = not self.filas or messagebox.askyesno(
            "Importar", f"Se leyeron {len(filas)} elementos.\n¿Reemplazar la tabla actual?\n(No = agregar al final)")
        self.filas = filas if reemplazar else self.filas + filas
        self.recalcular()
        self.estado(f"Importados {len(filas)} elementos desde {os.path.basename(ruta)}"
                    + (f" ({omitidas} filas omitidas sin partida reconocible)." if omitidas else "."))

    def guardar_plantilla(self):
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Plantilla_Elementos_a_Demoler.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if ruta:
            exportar_plantilla(ruta)
            abrir_archivo(ruta)

    def cargar_fotos(self):
        rutas = filedialog.askopenfilenames(title=f"Fotos - {self.cat_foto.get()}",
                                            filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Todos", "*.*")])
        for r in rutas:
            self.fotos.append((self.cat_foto.get(), r))
        self._refrescar_fotos()

    def quitar_fotos(self):
        self.fotos.clear()
        self._refrescar_fotos()

    def _refrescar_fotos(self):
        if not self.fotos:
            self.lbl_fotos.configure(text="Sin fotos cargadas (se insertarán recuadros para adjuntar).")
            return
        cuenta = {}
        for c, _ in self.fotos:
            cuenta[c] = cuenta.get(c, 0) + 1
        self.lbl_fotos.configure(text=f"{len(self.fotos)} foto(s): " + ", ".join(f"{k} ({v})" for k, v in cuenta.items()))

    def exportar(self, tipo):
        self.recalcular()
        if not self.filas:
            messagebox.showwarning("Exportar", "La tabla de metrados está vacía. Importe un Excel o cargue el caso "
                                               "de ejemplo.")
            return
        datos = self.leer_datos()
        base = nombre_seguro(datos.get("proyecto") or "Demolicion")
        nombres = {"docx": (f"Memoria_Demolicion_G050_{base}.docx", "Word", "*.docx"),
                   "dxf": (f"Lamina_D-01_Demolicion_{base}.dxf", "AutoCAD DXF", "*.dxf"),
                   "xlsx": (f"Planilla_Metrados_Demolicion_{base}.xlsx", "Excel", "*.xlsx")}
        nombre, desc, patron = nombres[tipo]
        ruta = filedialog.asksaveasfilename(defaultextension="." + tipo, initialfile=nombre,
                                            filetypes=[(desc, patron)])
        if not ruta:
            return
        self.estado(f"Generando {os.path.basename(ruta)} ...")
        self.configure(cursor="watch")
        self.update()
        try:
            if tipo == "docx":
                exportar_word(datos, self.res, self.fotos, {k: v.get() for k, v in self.check_vars.items()}, ruta)
            elif tipo == "dxf":
                exportar_dxf(datos, self.res, ruta)
            else:
                exportar_excel(datos, self.res, ruta)
        except PermissionError:
            messagebox.showerror("Exportar", "No se pudo guardar: el archivo está abierto en otro programa.\n"
                                             "Ciérrelo e intente nuevamente.")
            self.estado("Error: archivo en uso.")
            return
        except Exception as exc:
            messagebox.showerror("Exportar", f"Error al generar el archivo:\n{exc}")
            self.estado(f"Error: {exc}")
            return
        finally:
            self.configure(cursor="")
        self.estado(f"✔ Archivo generado: {ruta}")
        abrir_archivo(ruta)


# =============================================================================
# 10. PUNTO DE ENTRADA
# =============================================================================
def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    app = AppDemolicion()
    app.mainloop()


if __name__ == "__main__":
    main()

# =============================================================================
# INSTALACIÓN Y COMPILACIÓN
# -----------------------------------------------------------------------------
#   pip install customtkinter openpyxl pandas ezdxf python-docx Pillow
#
#   Ejecutar:
#   python app_demolicion_escombros_pro.py
#
#   Crear ejecutable único para Windows (.exe):
#   pip install pyinstaller
#   pyinstaller --noconfirm --onefile --windowed --name "GeneradorDemolicionPro" ^
#       --collect-all customtkinter --collect-all ezdxf --collect-data docx ^
#       app_demolicion_escombros_pro.py
# =============================================================================
