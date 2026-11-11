# -*- coding: utf-8 -*-
"""
===============================================================================
 GENERADOR PRO DE PLAN DE SEGURIDAD Y SALUD EN OBRA (PSST - NORMA G.050),
 MATRICES IPERC SINCRONIZADAS AL PRESUPUESTO, DESAGREGADO EPI/EPC
 Y PLANO EN AUTOCAD (.DXF)
===============================================================================
 Aplicación de escritorio 100 % offline (un solo archivo).

 Marco normativo considerado (Perú):
   - Ley N° 29783, Ley de Seguridad y Salud en el Trabajo (y Ley N° 30222).
   - D.S. N° 005-2012-TR, Reglamento de la Ley N° 29783.
   - D.S. N° 011-2019-TR, Reglamento de SST para el Sector Construcción.
   - Norma Técnica G.050 "Seguridad durante la Construcción" (RNE).
   - R.M. N° 050-2013-TR, Formatos referenciales y metodología IPERC.
   - R.M. N° 375-2008-TR, Norma Básica de Ergonomía.
   - NTP 399.010-1 Señales de seguridad / NTP 900.058 Código de colores RR.SS.

 Salidas:
   1) Libro Excel (.xlsx)  : Matriz IPERC, Presupuesto SSOMA G.050,
                             Programa de capacitaciones y Formatos ATS/PETAR.
   2) Documento Word (.docx): PSST completo con PETS y Plan de Emergencias.
   3) Plano AutoCAD (.dxf) : Lámina SS-01 (layout, cachacos, entibado,
                             andamio, señalética, RR.SS. y membrete).

 Uso:
   python app_psst_g050_iperc_pro.py               -> abre la interfaz gráfica
   python app_psst_g050_iperc_pro.py --demo CARPETA -> genera los 3 archivos de
                                                      la obra de ejemplo sin GUI
===============================================================================
"""

import math
import os
import re
import sys
import subprocess
import unicodedata
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import ezdxf
from ezdxf.enums import TextEntityAlignment

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import customtkinter as ctk
    GUI_DISPONIBLE = True
except Exception:  # permite usar los motores sin entorno gráfico
    GUI_DISPONIBLE = False


APP_TITULO = ("Generador Pro de Plan de Seguridad y Salud en Obra (PSST - Norma G.050), "
              "Matrices IPERC Sincronizadas al Presupuesto, Desagregado EPI/EPC y "
              "Plano en AutoCAD (.DXF)")
APP_CORTO = "Generador Pro PSST G.050 + IPERC + EPI/EPC + DXF"
APP_VERSION = "1.0"


# =============================================================================
# 1. UTILIDADES GENERALES
# =============================================================================
def normalizar(texto):
    """Minúsculas, sin tildes y con espacios simples (para emparejar)."""
    if texto is None:
        return ""
    t = str(texto).lower()
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = t.replace("’", "'").replace("´", "'")
    return re.sub(r"\s+", " ", t).strip()


def nivel_riesgo(valor):
    """Nivel de riesgo según R.M. N° 050-2013-TR (P x S)."""
    if valor <= 4:
        return "TRIVIAL"
    if valor <= 8:
        return "TOLERABLE"
    if valor <= 16:
        return "MODERADO"
    if valor <= 24:
        return "IMPORTANTE"
    return "INTOLERABLE"


def abrir_archivo(ruta):
    """Abre el archivo generado con la aplicación predeterminada."""
    try:
        if hasattr(os, "startfile"):
            os.startfile(ruta)  # Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta], stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
    except Exception:
        pass


def soles(valor):
    return "S/ {:,.2f}".format(valor)


def a_numero(valor):
    """Convierte textos como '1,250.50' o 1250 a float; None si no es número."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and math.isnan(valor):
            return None
        return float(valor)
    s = str(valor).strip().replace("S/", "").replace(" ", "")
    if not s:
        return None
    if s.count(",") and s.count("."):
        s = s.replace(",", "")
    elif s.count(",") == 1 and len(s.split(",")[1]) <= 2:
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


# =============================================================================
# 2. CRITERIOS IPERC (R.M. N° 050-2013-TR)
# =============================================================================
CRITERIOS_IPERC = {
    "A": ("Índice de personas expuestas",
          ["1: De 1 a 3 personas", "2: De 4 a 12 personas", "3: Más de 12 personas"]),
    "B": ("Índice de procedimientos existentes",
          ["1: Existen, son satisfactorios y suficientes",
           "2: Existen parcialmente y no son satisfactorios o suficientes",
           "3: No existen"]),
    "C": ("Índice de capacitación",
          ["1: Personal entrenado. Conoce el peligro y lo previene",
           "2: Personal parcialmente entrenado, conoce el peligro pero no toma acciones de control",
           "3: Personal no entrenado, no conoce el peligro, no toma acciones de control"]),
    "D": ("Índice de exposición al riesgo",
          ["1: Al menos una vez al año / Esporádicamente",
           "2: Al menos una vez al mes / Eventualmente",
           "3: Al menos una vez al día / Permanentemente"]),
    "S": ("Índice de severidad",
          ["1: Lesión sin incapacidad / Disconfort-incomodidad",
           "2: Lesión con incapacidad temporal / Daño a la salud reversible",
           "3: Lesión con incapacidad permanente o muerte / Daño a la salud irreversible"]),
}

NIVELES_RIESGO = [
    ("TRIVIAL", "4", "No se necesita adoptar ninguna acción.", "C6EFCE"),
    ("TOLERABLE", "5 - 8", "No se necesita mejorar la acción preventiva; considerar soluciones "
     "más rentables y comprobaciones periódicas.", "92D050"),
    ("MODERADO", "9 - 16", "Se deben hacer esfuerzos para reducir el riesgo en un periodo "
     "determinado. Implementar medidas de control.", "FFFF00"),
    ("IMPORTANTE", "17 - 24", "No debe comenzarse el trabajo hasta que se haya reducido el "
     "riesgo. Corregir de inmediato si el trabajo está en curso.", "FF9900"),
    ("INTOLERABLE", "25 - 36", "No debe comenzar ni continuar el trabajo hasta que se reduzca "
     "el riesgo. Si no es posible, debe prohibirse el trabajo.", "FF0000"),
]


# =============================================================================
# 3. DICCIONARIO EXPERTO IPERC: 18 FAMILIAS DE ACTIVIDADES DE CONSTRUCCIÓN
#    ini = (A, B, C, D, S) evaluación inicial | res = (A, B, C, D, S) residual
# =============================================================================
FAMILIAS_IPERC = [
    {
        "codigo": "F01", "nombre": "Obras provisionales, almacén, cartel de obra y movilización",
        "claves": ["obras provisionales", "obra provisional", "provisional", "almacen", "caseta",
                   "cartel", "movilizacion", "desmovilizacion", "cerco", "oficina", "guardiania",
                   "trabajos preliminares", "campamento", "flete", "transporte de materiales",
                   "limpieza de terreno", "agua para la construccion", "comedor", "vestuario"],
        "tipo": "Rutinaria",
        "peligro": "Manipulación manual de cargas (> 25 kg), tránsito de camiones en maniobra y uso de "
                   "herramientas manuales (comba, martillo, serrucho).",
        "riesgo": "Sobreesfuerzo, golpes y cortes con herramientas, atropello por vehículos de carga.",
        "consecuencia": "Lumbalgia, contusiones, heridas cortantes, fracturas, politraumatismo.",
        "ini": (2, 2, 2, 2, 3), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Planificar la descarga mecánica de materiales (camión grúa / montacargas) "
                       "eliminando el carguío manual repetitivo.",
        "sustitucion": "Uso de carretillas, buguis y estocas; materiales en presentaciones ≤ 25 kg "
                       "(R.M. N° 375-2008-TR).",
        "ingenieria": "Vías internas señalizadas y separadas de rutas peatonales; topes para camiones; "
                      "alarma de retroceso operativa.",
        "administrativo": "Vigía con paleta PARE/SIGA; capacitación en levantamiento seguro de cargas; "
                          "ATS diario; inspección de herramientas con cinta de color del mes.",
        "epp": "Casco ANSI Z89.1 con barbiquejo, lentes ANSI Z87.1, guantes de cuero/badana, botín "
               "punta de acero, uniforme con cinta reflectiva.",
        "petar": "ATS", "pets": "PETS-GEN Trabajos generales",
    },
    {
        "codigo": "F02", "nombre": "Demoliciones y desmontajes",
        "claves": ["demolicion", "demoler", "desmontaje", "desmontar", "retiro de", "picado",
                   "rotura de", "desmonte de", "remocion"],
        "tipo": "No rutinaria",
        "peligro": "Proyección de partículas, estructuras inestables, polvo con sílice cristalina, "
                   "ruido > 85 dB (martillo neumático / rotomartillo).",
        "riesgo": "Colapso de muros o losas, atrapamiento, golpes por proyección, inhalación de "
                  "polvo, exposición a ruido.",
        "consecuencia": "Atrapamiento, politraumatismo, lesiones oculares, neumoconiosis (silicosis), "
                        "hipoacusia inducida por ruido.",
        "ini": (2, 2, 2, 3, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Demolición mecánica a distancia (excavadora con martillo hidráulico); corte "
                       "previo de servicios de agua, energía y gas.",
        "sustitucion": "Corte con disco diamantado por vía húmeda en lugar de picado en seco.",
        "ingenieria": "Apuntalamiento previo; aislamiento del área con malla y cinta en un radio ≥ 1.5 H; "
                      "humedecimiento permanente; demolición de arriba hacia abajo.",
        "administrativo": "Estudio de estabilidad, PETS, ATS, vigía permanente, rotación del personal "
                          "expuesto a ruido/vibración, dosimetría de ruido.",
        "epp": "Casco, careta facial + lentes, respirador media cara con filtros P100, protección "
               "auditiva doble (tapón + orejera), guantes anticorte, botín punta de acero.",
        "petar": "ATS + Autorización de demolición", "pets": "PETS-GEN Demoliciones",
    },
    {
        "codigo": "F03", "nombre": "Trazo, nivelación y replanteo topográfico",
        "claves": ["trazo", "nivelacion", "replanteo", "topograf", "estacado", "control topografico",
                   "levantamiento topografico"],
        "tipo": "Rutinaria",
        "peligro": "Exposición prolongada a radiación UV, terreno irregular con desniveles, tránsito "
                   "vehicular en vía pública.",
        "riesgo": "Exposición a radiación solar, caídas al mismo nivel, atropello.",
        "consecuencia": "Insolación, quemaduras solares, lesiones dérmicas, esguinces, contusiones.",
        "ini": (1, 2, 2, 2, 2), "res": (1, 1, 1, 2, 1),
        "eliminacion": "Programar el trabajo de campo fuera del horario de mayor índice UV "
                       "(11:00-15:00) cuando sea posible.",
        "sustitucion": "Uso de estación total / GPS que reduce la permanencia del personal en la vía.",
        "ingenieria": "Delimitación con conos de 28\" y tranqueras; toldos o sombrillas; puntos de "
                      "hidratación.",
        "administrativo": "Vigía de tránsito, inspección previa del terreno, pausas activas e "
                          "hidratación cada hora; uso de bloqueador solar FPS 50+.",
        "epp": "Casco con cortaviento/legionario, lentes con filtro UV, chaleco reflectivo, bloqueador "
               "solar, botín punta de acero, ropa de manga larga.",
        "petar": "ATS", "pets": "PETS-GEN Trazo y replanteo",
    },
    {
        "codigo": "F04", "nombre": "Movimiento de tierras, excavación de zanjas y calzaduras (H > 1.50 m)",
        "claves": ["excavacion", "zanja", "calzadura", "corte de terreno", "movimiento de tierras",
                   "perfilado", "refine", "entibado", "banqueta", "sotano", "excavacion masiva",
                   "nivelacion de fondo"],
        "tipo": "No rutinaria",
        "peligro": "Taludes inestables (H > 1.50 m), maquinaria pesada en movimiento, redes "
                   "subterráneas (agua, gas, electricidad, telecomunicaciones).",
        "riesgo": "Derrumbe de paredes de zanja / sepultamiento, golpe o atropello por maquinaria, "
                  "contacto con redes energizadas.",
        "consecuencia": "Asfixia por sepultamiento, politraumatismo, electrocución, muerte.",
        "ini": (2, 2, 2, 3, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Evaluar métodos sin zanja en cruces críticos; ubicar redes existentes con "
                       "planos de catastro y detector de servicios antes de excavar.",
        "sustitucion": "Excavación mecánica con retroexcavadora en lugar de manual para H > 1.50 m, "
                       "con taludes estables.",
        "ingenieria": "Entibado/tablestacado o cajas de entibación para H > 1.50 m; material excavado "
                      "a ≥ H/2 (mín. 0.60 m) del borde; escalera de acceso cada 7.50 m sobresaliendo "
                      "1.00 m; cachacos con malla a 1.00 m del borde.",
        "administrativo": "PETAR de excavación, inspección diaria de la zanja por persona competente, "
                          "vigía de maquinaria, prohibición de ingreso al radio de giro, PETS-01, ATS.",
        "epp": "Casco, lentes, guantes, botín punta de acero, chaleco reflectivo, protector auditivo, "
               "respirador para polvo; arnés con línea de rescate en zanjas profundas.",
        "petar": "PETAR de Excavación y Zanjas", "pets": "PETS-01 Excavación de zanjas y entibados",
    },
    {
        "codigo": "F05", "nombre": "Eliminación de material excedente y relleno compactado",
        "claves": ["eliminacion", "material excedente", "relleno", "compactado", "compactacion",
                   "plancha compactadora", "rodillo", "afirmado", "base granular", "sub base",
                   "subbase", "conformacion", "acarreo", "carguio", "volquete",
                   "material de prestamo", "terraplen", "zarandeo"],
        "tipo": "Rutinaria",
        "peligro": "Vibración segmentaria (plancha / rodillo), ruido, polvo, retroceso de volquetes y "
                   "cargador frontal.",
        "riesgo": "Exposición a vibración y ruido, inhalación de polvo, atropello en maniobras de "
                  "retroceso.",
        "consecuencia": "Trastornos musculoesqueléticos, síndrome de vibración mano-brazo, hipoacusia, "
                        "atropello, muerte.",
        "ini": (2, 2, 2, 2, 3), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Planificar rutas de volquetes sin retroceso (circuito en una sola dirección).",
        "sustitucion": "Rodillo vibratorio autopropulsado con cabina en lugar de plancha manual en "
                       "grandes áreas.",
        "ingenieria": "Equipos con mangos antivibratorios; alarma de retroceso y cámara; riego "
                      "permanente para control de polvo.",
        "administrativo": "Vigía/cuadrador, rotación cada 2 h del operador de plancha, check-list de "
                          "maquinaria, velocidad máx. 10 km/h en obra.",
        "epp": "Casco, lentes, guantes antivibración, protector auditivo, respirador para polvo, botín "
               "punta de acero, chaleco reflectivo.",
        "petar": "ATS", "pets": "PETS-GEN Relleno y eliminación",
    },
    {
        "codigo": "F06", "nombre": "Obras de concreto simple y armado (mezclado, vibrado, vaciado)",
        "claves": ["concreto", "vaciado", "mezcla de concreto", "vibrado", "f'c", "fc=", "losa",
                   "columna", "viga", "zapata", "cimiento", "sobrecimiento", "placa", "solado",
                   "dado", "platea", "muro de concreto", "cemento", "bomba", "winche", "c:h"],
        "tipo": "Rutinaria",
        "peligro": "Contacto dérmico con cemento (alcalino pH 12-13), posturas forzadas en vibrado, "
                   "rotura / latigazo de mangueras de bomba, bordes de losa.",
        "riesgo": "Contacto con sustancia cáustica, sobreesfuerzo, golpe por manguera, caída a "
                  "distinto nivel.",
        "consecuencia": "Dermatitis de contacto / quemaduras químicas, lumbalgia, contusiones, "
                        "fracturas.",
        "ini": (2, 2, 2, 3, 2), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Concreto premezclado con bomba telescópica, eliminando el mezclado y "
                       "transporte manual.",
        "sustitucion": "Vibradores de alta frecuencia livianos; mezcladora con tolva de carga mecánica.",
        "ingenieria": "Mangueras con abrazaderas y cable de seguridad antilátigo; barandas perimetrales "
                      "en bordes de losa; plataformas de vaciado.",
        "administrativo": "PETS-04, ATS, rotación de vibradoristas, señalero de bomba, verificación de "
                          "encofrados antes del vaciado, lavado de piel al término.",
        "epp": "Casco, lentes antiempañantes, guantes de nitrilo/neopreno, botas de jebe con punta de "
               "acero, ropa impermeable, crema barrera, respirador para polvo de cemento.",
        "petar": "ATS", "pets": "PETS-04 Vaciado de concreto y vibrado",
    },
    {
        "codigo": "F07", "nombre": "Encofrado y desencofrado (madera / metálico)",
        "claves": ["encofrado", "desencofrado", "encofrar", "triplay", "puntal", "cimbra",
                   "andamiaje de encofrado"],
        "tipo": "Rutinaria",
        "peligro": "Uso de sierra circular, clavos expuestos, caída de paneles y materiales en altura.",
        "riesgo": "Contacto con disco de corte, punzonamiento, golpe por caída de objetos.",
        "consecuencia": "Cortes profundos, amputaciones, heridas punzantes plantares, contusiones.",
        "ini": (2, 2, 2, 3, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Uso de madera precortada en taller / sistemas de encofrado modular "
                       "prefabricado.",
        "sustitucion": "Encofrado metálico con conectores en lugar de clavado.",
        "ingenieria": "Sierra circular con guarda retráctil, cuchillo divisor y paro de emergencia; "
                      "desclavado inmediato de maderas; zona de corte delimitada.",
        "administrativo": "Solo personal autorizado opera la sierra; orden y limpieza al término de la "
                          "jornada; desencofrado secuencial con vigía; ATS.",
        "epp": "Casco, lentes + careta facial, protector auditivo, guantes de cuero, botín con "
               "plantilla antiperforante, mandil de cuero para el operador de sierra.",
        "petar": "ATS", "pets": "PETS-GEN Encofrado y desencofrado",
    },
    {
        "codigo": "F08", "nombre": "Habilitado y armado de acero corrugado",
        "claves": ["acero corrugado", "acero", "fierro", "corrugado", "fy=4200", "fy = 4200",
                   "habilitado", "armadura", "varilla", "estribo", "grado 60"],
        "tipo": "Rutinaria",
        "peligro": "Uso de tronzadora / amoladora, puntas de fierro expuestas sin protección, carga "
                   "manual de varillas.",
        "riesgo": "Contacto con disco abrasivo, proyección de partículas, empalamiento con puntas de "
                  "acero, sobreesfuerzo.",
        "consecuencia": "Cortes severos, lesiones oculares, empalamiento / empalomamiento, lumbalgia.",
        "ini": (2, 2, 2, 3, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Acero habilitado en planta (cortado y doblado) que elimina el corte en obra.",
        "sustitucion": "Cizalla/dobladora mecánica en lugar de amoladora para cortes rectos.",
        "ingenieria": "Uso obligatorio de CAPUCHONES PLÁSTICOS / HONGOS DE SEGURIDAD en todas las "
                      "mechas; guardas en tronzadora y amoladora; banco de habilitado fijo.",
        "administrativo": "PETS-03, ATS, transporte de varillas entre 2 personas, revisión de discos "
                          "(fecha de vencimiento y RPM), orden de material habilitado.",
        "epp": "Casco, lentes + careta facial, guantes de cuero cromo, protector auditivo, botín punta "
               "de acero, hombreras de cuero para carga.",
        "petar": "ATS", "pets": "PETS-03 Habilitado de acero y uso de amoladora/tronzadora",
    },
    {
        "codigo": "F09", "nombre": "Albañilería, muros de ladrillo, tarrajeo y derrames en altura (H > 1.80 m)",
        "claves": ["albanileria", "muro de ladrillo", "ladrillo", "tarrajeo", "tarrajeado", "derrame",
                   "revoque", "cielo raso", "solaqueo", "enlucido", "asentado", "bloqueta", "tabique",
                   "vestidura", "muros"],
        "tipo": "Rutinaria",
        "peligro": "Trabajo sobre andamios y bordes a H > 1.80 m, polvo de mortero y cemento, "
                   "manipulación de ladrillos.",
        "riesgo": "Caída a distinto nivel, caída de objetos, inhalación de polvo.",
        "consecuencia": "Fracturas, TEC, politraumatismo, muerte; irritación respiratoria.",
        "ini": (2, 2, 2, 3, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Prefabricar elementos a nivel del piso cuando sea posible.",
        "sustitucion": "Andamio multidireccional normado en lugar de andamios artesanales o "
                       "caballetes.",
        "ingenieria": "Plataforma completa, barandas a 0.90 m y rodapié de 0.15 m; línea de vida; "
                      "redes de protección; andamio arriostrado a la estructura.",
        "administrativo": "PETAR de altura, PETS-02, inspección diaria del andamio (tarjeta verde/roja), "
                          "personal con EMO para trabajos en altura, ATS.",
        "epp": "Arnés de cuerpo completo con doble línea de vida y absorbedor de impacto (ANSI Z359), "
               "casco con barbiquejo, lentes, guantes de badana, respirador para polvo.",
        "petar": "PETAR de Trabajos en Altura", "pets": "PETS-02 Trabajos en altura y uso de andamios",
    },
    {
        "codigo": "F10", "nombre": "Estructuras metálicas, soldadura y coberturas",
        "claves": ["estructura metalica", "estructuras metalicas", "soldadura", "soldar", "tijeral",
                   "cobertura", "calamina", "techo metalico", "viga metalica", "perfil", "oxicorte",
                   "aluzinc", "cubierta", "termoacustica", "policarbonato", "arco metalico"],
        "tipo": "No rutinaria",
        "peligro": "Humos metálicos de soldadura, radiación del arco eléctrico, izaje de tijerales, "
                   "trabajo en altura sobre cubiertas.",
        "riesgo": "Inhalación de humos, exposición a radiación UV/IR, contacto eléctrico, caída a "
                  "distinto nivel, incendio.",
        "consecuencia": "Quemaduras, conjuntivitis actínica, fiebre de humos metálicos, electrocución, "
                        "caída mortal.",
        "ini": (2, 3, 2, 3, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Prefabricar y soldar en taller a nivel del piso; montaje con uniones "
                       "empernadas.",
        "sustitucion": "Plataforma elevadora (manlift) en lugar de escaleras o andamios artesanales.",
        "ingenieria": "Extracción localizada de humos, biombos ignífugos, máquina de soldar con puesta a "
                      "tierra, línea de vida horizontal de cable Ø1/2\", extintor PQS en el punto.",
        "administrativo": "PETAR de trabajo en caliente + PETAR de altura, vigía contra incendio "
                          "(30 min post-trabajo), retiro de combustibles en radio de 11 m.",
        "epp": "Careta de soldar fotosensible, casaca/mandil, guantes de caña larga y escarpines de "
               "cuero, respirador para humos metálicos, arnés con doble línea de vida.",
        "petar": "PETAR Trabajo en Caliente + PETAR Altura", "pets": "PETS-GEN Soldadura y montaje",
    },
    {
        "codigo": "F11", "nombre": "Instalaciones sanitarias, redes de agua, desagüe y buzones",
        "claves": ["sanitaria", "agua potable", "desague", "alcantarillado", "buzon", "buzoneta",
                   "tuberia", "pvc", "red de agua", "conexion domiciliaria", "colector",
                   "caja de registro", "valvula", "aparatos sanitarios", "inodoro", "lavatorio",
                   "cisterna", "tanque elevado", "prueba hidraulica", "camara de inspeccion"],
        "tipo": "No rutinaria",
        "peligro": "Ingreso a buzones y cámaras profundas (espacio confinado), pegamento PVC con "
                   "solventes (COV), gases de alcantarillado (H2S, CH4).",
        "riesgo": "Atmósfera deficiente en oxígeno o tóxica, inhalación de vapores orgánicos, caída "
                  "al interior del buzón.",
        "consecuencia": "Asfixia, intoxicación aguda, pérdida de conciencia, muerte.",
        "ini": (2, 2, 2, 2, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Realizar las conexiones desde el exterior del buzón cuando sea factible.",
        "sustitucion": "Pegamentos de bajo contenido de COV; uniones flexibles con anillo de jebe.",
        "ingenieria": "Ventilación forzada; monitoreo continuo con detector multigás (O2 19.5-23.5 %, "
                      "LEL < 10 %, H2S < 10 ppm, CO < 25 ppm); trípode de rescate.",
        "administrativo": "PETAR de espacio confinado, vigía permanente en la boca del buzón, plan de "
                          "rescate, hojas MSDS en obra, ATS.",
        "epp": "Arnés con anillo dorsal y línea de rescate, respirador con cartuchos para vapores "
               "orgánicos, guantes de nitrilo, botas de jebe, casco, lentes.",
        "petar": "PETAR de Espacios Confinados", "pets": "PETS-GEN Espacios confinados",
    },
    {
        "codigo": "F12", "nombre": "Instalaciones eléctricas, tableros y puesta a tierra",
        "claves": ["electric", "tablero", "puesta a tierra", "pozo a tierra", "cableado", "cable",
                   "conductor", "alumbrado", "tomacorriente", "interruptor", "luminaria",
                   "salida de", "subestacion", "media tension", "baja tension", "iluminacion",
                   "canalizacion", "conduit", "acometida", "pararrayos"],
        "tipo": "No rutinaria",
        "peligro": "Energía eléctrica (contacto directo e indirecto), tendido de cables, equipos "
                   "energizados.",
        "riesgo": "Contacto eléctrico, arco eléctrico, caída en altura al instalar luminarias.",
        "consecuencia": "Electrocución, quemaduras de 2.° y 3.° grado, paro cardiorrespiratorio, "
                        "muerte.",
        "ini": (2, 2, 2, 2, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Trabajar siempre con circuitos desenergizados (cero energía verificada).",
        "sustitucion": "Herramientas a batería y luminarias de 24 V en zonas húmedas.",
        "ingenieria": "Tableros con interruptor diferencial 30 mA y puesta a tierra; Bloqueo y "
                      "Etiquetado (LOTO); cables con aislamiento NLT/NMT sin empalmes expuestos.",
        "administrativo": "PETAR de trabajos eléctricos, PETS-05, solo electricistas calificados, "
                          "verificación con revelador de tensión, ATS.",
        "epp": "Casco dieléctrico Clase E, guantes dieléctricos Clase 0, botín dieléctrico, lentes, "
               "ropa ignífuga (arco eléctrico), herramientas aisladas 1000 V.",
        "petar": "PETAR Trabajos Eléctricos (Bloqueo LOTO)",
        "pets": "PETS-05 Instalaciones eléctricas provisionales y bloqueo LOTO",
    },
    {
        "codigo": "F13", "nombre": "Pavimentos, asfalto en caliente, sardineles y veredas",
        "claves": ["pavimento", "asfalto", "asfaltica", "carpeta", "imprimacion", "riego de liga",
                   "sardinel", "vereda", "berma", "junta", "cortadora", "adoquin", "cuneta",
                   "rampa", "giba", "pavimentacion", "bacheo"],
        "tipo": "Rutinaria",
        "peligro": "Vapores bituminosos a alta temperatura (≈ 150 °C), corte de juntas con cortadora "
                   "de disco, tránsito vehicular, rodillos y pavimentadora.",
        "riesgo": "Contacto con material caliente, inhalación de vapores, proyección de esquirlas, "
                  "atropello.",
        "consecuencia": "Quemaduras térmicas, irritación respiratoria, lesiones oculares, "
                        "politraumatismo.",
        "ini": (2, 2, 2, 3, 3), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Desvío total del tránsito vehicular durante el asfaltado.",
        "sustitucion": "Asfalto tibio (menor temperatura de colocación) y cortadora con agua.",
        "ingenieria": "Plan de desvíos con tranqueras, conos de 28\" y señales preventivas; guardas en "
                      "cortadora; zona de exclusión alrededor de la pavimentadora.",
        "administrativo": "Vigías con paletas, control de temperatura, rotación de rastrilleros, ATS.",
        "epp": "Casco, lentes, careta facial, guantes térmicos de caña larga, botas de seguridad con "
               "suela resistente al calor, respirador para vapores, chaleco reflectivo.",
        "petar": "ATS", "pets": "PETS-GEN Pavimentación",
    },
    {
        "codigo": "F14", "nombre": "Pintura, barnizado, epóxicos y acabados",
        "claves": ["pintura", "pintado", "barniz", "barnizado", "epoxi", "epoxico", "esmalte", "latex",
                   "imprimante", "laca", "sellador", "empaste", "anticorrosivo", "tinte"],
        "tipo": "Rutinaria",
        "peligro": "Solventes orgánicos (thinner, xileno, epóxicos), uso de escaleras tijera y "
                   "andamios.",
        "riesgo": "Inhalación de vapores, contacto dérmico, incendio, caída a distinto nivel.",
        "consecuencia": "Intoxicación respiratoria, dermatitis, quemaduras, fracturas.",
        "ini": (2, 2, 2, 2, 3), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Elementos pintados de fábrica cuando sea posible.",
        "sustitucion": "Pinturas base agua (látex) de bajo COV en lugar de base solvente.",
        "ingenieria": "Ventilación natural/forzada; almacén de químicos con bandeja antiderrame; "
                      "escaleras tijera con seguro y peldaños antideslizantes.",
        "administrativo": "Hojas MSDS/SDS en el punto, prohibido fumar, inspección de escaleras, ATS.",
        "epp": "Respirador media cara con cartuchos para vapores orgánicos, lentes, guantes de nitrilo, "
               "mameluco descartable, casco.",
        "petar": "ATS", "pets": "PETS-GEN Pintura y acabados",
    },
    {
        "codigo": "F15", "nombre": "Pisos, contrapisos, enchapes cerámicos y porcelanatos",
        "claves": ["piso", "contrapiso", "ceramic", "porcelanato", "enchape", "zocalo",
                   "contrazocalo", "falso piso", "cemento pulido", "mayolica", "terrazo", "loseta",
                   "vinilico", "pulido"],
        "tipo": "Rutinaria",
        "peligro": "Posturas forzadas de rodillas, corte de cerámicos con cortadora/amoladora, "
                   "adhesivos y fragua.",
        "riesgo": "Sobreesfuerzo, cortes, proyección de partículas, inhalación de polvo de sílice.",
        "consecuencia": "Bursitis y lesiones de rodilla, cortes, lesiones oculares.",
        "ini": (2, 2, 2, 3, 2), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Piezas cortadas a medida desde fábrica en zonas repetitivas.",
        "sustitucion": "Cortadora manual de cerámico (sin disco) para cortes rectos.",
        "ingenieria": "Corte por vía húmeda o con aspiración; mesa de corte a la altura del codo.",
        "administrativo": "Pausas activas cada 2 h, rotación de tareas, ATS.",
        "epp": "Rodilleras, lentes, guantes de nitrilo, respirador para polvo, protector auditivo, "
               "casco, botín.",
        "petar": "ATS", "pets": "PETS-GEN Pisos y enchapes",
    },
    {
        "codigo": "F16", "nombre": "Carpintería de madera, metálica, cerrajería y vidrios",
        "claves": ["carpinteria", "puerta", "ventana", "vidrio", "cristal", "mampara", "closet",
                   "mueble", "cerrajeria", "bisagra", "baranda metalica", "pasamanos", "reja"],
        "tipo": "Rutinaria",
        "peligro": "Manipulación de vidrios y hojas de puertas, herramientas eléctricas portátiles, "
                   "trabajo en escaleras.",
        "riesgo": "Cortes por rotura de vidrio, golpes, contacto eléctrico, caída a distinto nivel.",
        "consecuencia": "Laceraciones, contusiones, fracturas.",
        "ini": (2, 2, 2, 2, 2), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Carpintería instalada con medidas de fábrica (sin cortes en obra).",
        "sustitucion": "Ventosas manuales para manipular vidrios.",
        "ingenieria": "Herramientas con doble aislamiento; caballetes estables; vidrios con cinta "
                      "señalizadora durante la instalación.",
        "administrativo": "Manipulación entre 2 personas, inspección de herramientas, ATS.",
        "epp": "Guantes anticorte nivel 5, mangas protectoras, lentes, casco, botín.",
        "petar": "ATS", "pets": "PETS-GEN Carpintería",
    },
    {
        "codigo": "F17", "nombre": "Izaje de cargas con grúa, camión grúa o winche",
        "claves": ["izaje", "grua", "camion grua", "montacarga", "montaje de", "elevacion de cargas",
                   "prefabricado", "izar"],
        "tipo": "No rutinaria",
        "peligro": "Cargas suspendidas, falla de accesorios de izaje (estrobos, grilletes), líneas "
                   "eléctricas aéreas, volcadura de grúa.",
        "riesgo": "Caída de carga suspendida, golpe o aplastamiento, electrocución por proximidad.",
        "consecuencia": "Aplastamiento, politraumatismo, muerte.",
        "ini": (2, 2, 2, 3, 3), "res": (1, 1, 1, 2, 3),
        "eliminacion": "Diseñar elementos más livianos o ensamblar a nivel del piso.",
        "sustitucion": "Equipo de izaje de mayor capacidad (carga ≤ 75 % de la tabla de carga).",
        "ingenieria": "Estabilizadores sobre tacos de madera; vientos de guía; distancia ≥ 3 m a "
                      "líneas eléctricas; accesorios certificados y con código de color.",
        "administrativo": "PETAR de izaje con plan de rigging, rigger certificado, área de exclusión, "
                          "operador acreditado, check-list de grúa.",
        "epp": "Casco con barbiquejo, guantes de cuero, botín, chaleco reflectivo, lentes; paleta y "
               "radio para el rigger.",
        "petar": "PETAR de Izaje de Cargas (Plan de Rigging)", "pets": "PETS-GEN Izaje de cargas",
    },
    {
        "codigo": "F18", "nombre": "Señalización vial, mitigación ambiental y limpieza final de obra",
        "claves": ["limpieza final", "limpieza permanente", "senalizacion", "seguridad y salud",
                   "mitigacion", "ambiental", "plan de manejo", "desvio de transito", "jardineria",
                   "area verde", "sembrado", "grass", "arborizacion", "entrega de obra",
                   "residuos", "pintura de trafico"],
        "tipo": "Rutinaria",
        "peligro": "Tránsito vehicular en desvíos, manipulación de residuos, herramientas manuales, "
                   "exposición a polvo.",
        "riesgo": "Atropello, cortes, contacto con residuos peligrosos.",
        "consecuencia": "Contusiones, heridas, infecciones, politraumatismo.",
        "ini": (2, 2, 2, 2, 2), "res": (1, 1, 1, 2, 2),
        "eliminacion": "Segregación de residuos en la fuente según NTP 900.058.",
        "sustitucion": "Uso de herramientas de mango largo y contenedores con ruedas.",
        "ingenieria": "Tranqueras, conos y señales en desvíos; punto de acopio techado con cilindros "
                      "por colores.",
        "administrativo": "Vigías de tránsito, manejo de residuos por EO-RS autorizada, ATS.",
        "epp": "Casco, lentes, guantes de nitrilo/cuero, respirador para polvo, chaleco reflectivo, "
               "botín.",
        "petar": "ATS", "pets": "PETS-GEN Limpieza y señalización",
    },
]

FAMILIA_GENERAL = {
    "codigo": "F00", "nombre": "Actividad general de construcción (sin palabra clave)",
    "claves": [], "tipo": "Rutinaria",
    "peligro": "Herramientas manuales, superficies irregulares, manipulación de materiales.",
    "riesgo": "Golpes, cortes, caídas al mismo nivel, sobreesfuerzo.",
    "consecuencia": "Contusiones, heridas, esguinces, lumbalgia.",
    "ini": (2, 2, 2, 2, 2), "res": (1, 1, 1, 2, 2),
    "eliminacion": "Planificar la tarea para evitar manipulación innecesaria.",
    "sustitucion": "Herramientas ergonómicas y medios mecánicos de transporte.",
    "ingenieria": "Orden y limpieza, accesos despejados, iluminación adecuada.",
    "administrativo": "ATS diario, charla de 5 minutos, inspección de herramientas.",
    "epp": "Casco, lentes, guantes, botín punta de acero, uniforme con cinta reflectiva.",
    "petar": "ATS", "pets": "PETS-GEN Trabajos generales",
}

FAMILIAS_POR_CODIGO = {f["codigo"]: f for f in FAMILIAS_IPERC + [FAMILIA_GENERAL]}

# familias que activan trabajos de alto riesgo
FAMILIA_A_RIESGO = {
    "F04": ["excavacion"], "F09": ["altura"], "F10": ["altura", "caliente", "izaje"],
    "F11": ["confinado"], "F17": ["izaje"], "F02": ["altura"],
}

PALABRAS_NO_PARTIDA = ["costo directo", "gastos generales", "utilidad", "sub total", "subtotal",
                       "igv", "total presupuesto", "presupuesto total", "valor referencial",
                       "son:", "supervision", "expediente tecnico", "total"]


def emparejar_familia(descripcion):
    """Motor de emparejamiento por palabras clave.
    Puntaje = suma de longitudes de las claves encontradas + 20 si la clave
    aparece al inicio de la partida (verbo principal)."""
    texto = normalizar(descripcion)
    mejor, mejor_puntaje = FAMILIA_GENERAL, 0
    for fam in FAMILIAS_IPERC:
        puntaje = 0
        for clave in fam["claves"]:
            pos = texto.find(clave)
            if pos >= 0:
                puntaje += len(clave)
                if pos <= 2:
                    puntaje += 20
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = fam, puntaje
    return mejor


def fila_iperc(partida, familia):
    a, b, c, d, s = familia["ini"]
    ra, rb, rc, rd, rs = familia["res"]
    ip = a + b + c + d
    rip = ra + rb + rc + rd
    return {
        "item": partida.get("item", ""), "partida": partida.get("descripcion", ""),
        "und": partida.get("und", ""), "metrado": partida.get("metrado"),
        "familia_cod": familia["codigo"], "familia": familia["nombre"], "tipo": familia["tipo"],
        "peligro": familia["peligro"], "riesgo": familia["riesgo"],
        "consecuencia": familia["consecuencia"],
        "A": a, "B": b, "C": c, "D": d, "IP": ip, "S": s, "PxS": ip * s,
        "nivel": nivel_riesgo(ip * s),
        "eliminacion": familia["eliminacion"], "sustitucion": familia["sustitucion"],
        "ingenieria": familia["ingenieria"], "administrativo": familia["administrativo"],
        "epp": familia["epp"],
        "rA": ra, "rB": rb, "rC": rc, "rD": rd, "rIP": rip, "rS": rs, "rPxS": rip * rs,
        "nivel_res": nivel_riesgo(rip * rs),
        "petar": familia["petar"], "pets": familia["pets"],
    }


def construir_iperc(partidas):
    return [fila_iperc(p, emparejar_familia(p.get("descripcion", ""))) for p in partidas]


def riesgos_detectados(filas):
    flags = set()
    for f in filas:
        for r in FAMILIA_A_RIESGO.get(f["familia_cod"], []):
            flags.add(r)
    return flags


# =============================================================================
# 4. LECTOR INTELIGENTE DE PRESUPUESTOS EN EXCEL
# =============================================================================
RE_ITEM = re.compile(r"^\d{1,3}(\.\d{1,3}){0,6}\.?$")


def _detectar_columnas(df):
    """Busca la fila de encabezados y las columnas de ítem, descripción, unidad,
    metrado y precio. Devuelve (fila_encabezado, dict_columnas)."""
    for i in range(min(40, len(df))):
        fila = [normalizar(v) for v in df.iloc[i].tolist()]
        cols = {}
        for j, v in enumerate(fila):
            if not v:
                continue
            if "descrip" in v or v in ("partida", "partidas", "actividad", "nombre de partida"):
                cols.setdefault("desc", j)
            elif v.startswith("item") or v in ("cod", "codigo", "cod."):
                cols.setdefault("item", j)
            elif v in ("und", "und.", "unid", "unid.", "unidad", "u.m.", "um", "u/m"):
                cols.setdefault("und", j)
            elif "metrado" in v or v.startswith("cant"):
                cols.setdefault("metrado", j)
            elif "precio" in v or v in ("p.u.", "pu", "p.u", "costo unitario", "c.u."):
                cols.setdefault("pu", j)
            elif "parcial" in v:
                cols.setdefault("parcial", j)
        if "desc" in cols:
            return i, cols
    # sin encabezado: deducir columnas por contenido
    cols = {}
    mejor_txt, mejor_item = (-1, 0), (-1, 0)
    for j in range(df.shape[1]):
        serie = df.iloc[:, j].dropna().astype(str)
        if serie.empty:
            continue
        largo = serie.map(len).mean()
        items = serie.map(lambda x: bool(RE_ITEM.match(x.strip()))).sum()
        if largo > mejor_txt[1]:
            mejor_txt = (j, largo)
        if items > mejor_item[1]:
            mejor_item = (j, items)
    if mejor_txt[0] >= 0:
        cols["desc"] = mejor_txt[0]
    if mejor_item[0] >= 0 and mejor_item[0] != cols.get("desc"):
        cols["item"] = mejor_item[0]
    return -1, cols


def leer_presupuesto_excel(ruta):
    """Lee cualquier presupuesto de obra (.xlsx) y devuelve la lista de partidas."""
    libros = pd.read_excel(ruta, sheet_name=None, header=None, dtype=object, engine="openpyxl")
    mejor = []
    for _, df in libros.items():
        df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
        if df.empty:
            continue
        fila_enc, cols = _detectar_columnas(df)
        if "desc" not in cols:
            continue
        partidas = []
        for i in range(fila_enc + 1, len(df)):
            fila = df.iloc[i].tolist()

            def val(clave):
                j = cols.get(clave)
                if j is None or j >= len(fila):
                    return None
                v = fila[j]
                return None if (isinstance(v, float) and math.isnan(v)) else v

            desc = val("desc")
            if desc is None or len(str(desc).strip()) < 4:
                continue
            desc = re.sub(r"\s+", " ", str(desc)).strip()
            nd = normalizar(desc)
            if any(nd.startswith(p) for p in PALABRAS_NO_PARTIDA):
                continue
            und = val("und")
            und = "" if und is None else str(und).strip()
            metrado = a_numero(val("metrado"))
            pu = a_numero(val("pu"))
            es_partida = bool(und) or (metrado is not None and metrado > 0)
            if ("und" in cols or "metrado" in cols) and not es_partida:
                continue  # títulos / subtítulos
            item = val("item")
            item = "" if item is None else str(item).strip()
            partidas.append({"item": item, "descripcion": desc, "und": und,
                             "metrado": metrado, "pu": pu})
        if len(partidas) > len(mejor):
            mejor = partidas
    if not mejor:
        raise ValueError("No se encontraron partidas en el archivo. Verifique que el Excel "
                         "tenga una columna de descripción y otra de unidad o metrado.")
    for k, p in enumerate(mejor, 1):
        if not p["item"]:
            p["item"] = "{:02d}".format(k)
    return mejor


# Obra completa de ejemplo (Colegio + Pavimentación) con 18 partidas
PRESUPUESTO_EJEMPLO = [
    ("01.01", "Cartel de obra de 3.60 x 7.20 m con estructura de madera", "und", 1, 2850.00),
    ("01.02", "Construcción de almacén, caseta de guardianía y oficina provisional", "glb", 1, 9800.00),
    ("01.03", "Movilización y desmovilización de equipos y maquinaria", "glb", 1, 12500.00),
    ("02.01", "Demolición de losa de concreto existente y veredas deterioradas", "m2", 380, 28.50),
    ("02.02", "Trazo, nivelación y replanteo durante la ejecución de obra", "m2", 4500, 1.85),
    ("03.01", "Excavación de zanjas para cimientos y redes h = 1.80 m con entibado", "m3", 620, 42.30),
    ("03.02", "Eliminación de material excedente con volquete D = 10 km", "m3", 780, 26.40),
    ("03.03", "Relleno compactado con material de préstamo (plancha compactadora)", "m3", 410, 58.70),
    ("04.01", "Concreto f'c = 210 kg/cm2 en zapatas, columnas y vigas (vaciado con bomba)", "m3", 285, 465.80),
    ("04.02", "Encofrado y desencofrado normal de columnas, vigas y losas", "m2", 1650, 48.90),
    ("04.03", "Acero corrugado fy = 4200 kg/cm2 grado 60 habilitado y colocado", "kg", 24500, 5.45),
    ("05.01", "Muros de ladrillo King Kong de soga y tarrajeo de muros exteriores h > 1.80 m", "m2", 1320, 78.60),
    ("06.01", "Estructura metálica para cobertura de patio con soldadura y tijerales", "kg", 8600, 14.20),
    ("07.01", "Red de alcantarillado: tubería PVC UF DN 200 mm y buzones de concreto", "ml", 520, 145.30),
    ("07.02", "Instalaciones eléctricas: tableros de distribución y pozo de puesta a tierra", "glb", 1, 38500.00),
    ("08.01", "Pavimento flexible: carpeta asfáltica en caliente e = 2\"", "m2", 3200, 52.40),
    ("08.02", "Sardineles y veredas de concreto con corte de juntas", "m2", 860, 96.50),
    ("09.01", "Pintura látex en muros y esmalte epóxico en estructuras metálicas", "m2", 2950, 16.80),
]


def partidas_ejemplo():
    return [{"item": i, "descripcion": d, "und": u, "metrado": float(m), "pu": float(p)}
            for i, d, u, m, p in PRESUPUESTO_EJEMPLO]


def configuracion_ejemplo():
    return {
        "proyecto": "MEJORAMIENTO DEL SERVICIO EDUCATIVO DE LA I.E. N° 38001 Y PAVIMENTACIÓN DE "
                    "VÍAS DE ACCESO",
        "entidad": "MUNICIPALIDAD DISTRITAL DE SAN JUAN BAUTISTA",
        "distrito": "San Juan Bautista", "provincia": "Huamanga", "region": "Ayacucho",
        "modalidad": "Contrata", "plazo_valor": 180, "plazo_unidad": "Días Calendario",
        "n_obr": 25, "n_ing": 4, "perimetro": 260,
        "residente": "Ing. Residente de Obra", "ssoma": "Ing. Especialista SSOMA", "cip": "000000",
        "riesgos": {"altura": True, "excavacion": True, "caliente": True, "confinado": True,
                    "izaje": True},
    }


def completar_config(cfg):
    """Calcula meses / días y valores por defecto."""
    cfg = dict(cfg)
    valor = max(1, int(round(float(cfg.get("plazo_valor", 1)))))
    if cfg.get("plazo_unidad", "Meses").lower().startswith("d"):
        cfg["dias"] = valor
        cfg["meses"] = max(1, math.ceil(valor / 30.0))
    else:
        cfg["meses"] = valor
        cfg["dias"] = valor * 30
    cfg["n_obr"] = max(1, int(cfg.get("n_obr", 1)))
    cfg["n_ing"] = max(1, int(cfg.get("n_ing", 1)))
    cfg["n_total"] = cfg["n_obr"] + cfg["n_ing"]
    cfg["perimetro"] = max(20, float(cfg.get("perimetro", 200) or 200))
    cfg.setdefault("riesgos", {})
    for k in ("altura", "excavacion", "caliente", "confinado", "izaje"):
        cfg["riesgos"].setdefault(k, False)
    cfg["fecha"] = cfg.get("fecha") or datetime.now().strftime("%d/%m/%Y")
    cfg["ubicacion"] = "Distrito de {}, Provincia de {}, Región {}".format(
        cfg.get("distrito", "-"), cfg.get("provincia", "-"), cfg.get("region", "-"))
    return cfg


# =============================================================================
# 5. MOTOR CALCULADORA DEL DESAGREGADO DE SEGURIDAD G.050
#    cada ítem: (descripción, und, base, vida_útil_meses o None, P.U. S/)
#    Factor de reposición = REDONDEAR.MAS(Plazo / Vida útil) ; Cantidad = Base x Factor
# =============================================================================
def calcular_desagregado(cfg):
    cfg = completar_config(cfg)
    no, N, M, P = cfg["n_obr"], cfg["n_total"], cfg["meses"], cfg["perimetro"]
    rg = cfg["riesgos"]
    up = math.ceil

    n_alt = max(2, up(no * 0.35)) if rg["altura"] else 2
    n_sold = max(2, up(no * 0.08))

    epi = [
        ("Casco de seguridad ANSI Z89.1 Tipo I Clase E con suspensión de 4 apoyos y barbiquejo",
         "und", N, 12, 32.00),
        ("Lentes de seguridad de policarbonato claros/oscuros antiempañantes ANSI Z87.1+",
         "und", N, 2, 9.50),
        ("Guantes de cuero cromo reforzado para trabajo pesado", "par", no, 1, 14.00),
        ("Guantes de badana para operaciones generales", "par", up(no * 0.5), 1, 9.00),
        ("Guantes de hilo con puntos de PVC / recubiertos de nitrilo", "par", no, 1, 6.50),
        ("Botín de seguridad de cuero con punta de acero/composite NTP ISO 20345 dieléctrico",
         "par", N, 6, 115.00),
        ("Botas de jebe con punta de acero (vaciado de concreto y zonas húmedas)", "par",
         up(no * 0.4), 6, 58.00),
        ("Protector auditivo tipo tapón reusable de silicona NRR 25 dB", "par", no, 1, 3.50),
        ("Protector auditivo tipo orejera acoplable al casco NRR 22 dB", "und", up(no * 0.3), 12,
         45.00),
        ("Respirador de media cara de silicona doble filtro (NIOSH)", "und", up(no * 0.5), 12, 85.00),
        ("Filtros/cartuchos para polvo P100 y vapores orgánicos (par)", "par", up(no * 0.5), 1,
         48.00),
        ("Uniforme drill (camisa + pantalón) con cinta reflectiva 2\" - 2 juegos por trabajador",
         "jgo", 2 * N, 6, 95.00),
        ("Cortaviento / legionario para casco", "und", N, 6, 12.00),
        ("Chaleco de seguridad con cinta reflectiva ANSI/ISEA 107", "und", N, 6, 18.00),
        ("Arnés de cuerpo completo ANSI Z359.11 con doble línea de enganche y absorbedor de impacto "
         "(ANSI Z359.13)", "und", n_alt, 12, 450.00),
        ("Bloqueador solar FPS 50+ (frasco de 1 L)", "und", max(1, up(N / 10)), 1, 65.00),
    ]
    if rg["caliente"]:
        epi.append(("Careta de soldar fotosensible + casaca, guantes de caña larga y escarpines de "
                    "cuero (soldador)", "jgo", n_sold, 6, 320.00))
    if rg["confinado"]:
        epi.append(("Respirador con cartuchos para gases ácidos / vapores orgánicos (espacios "
                    "confinados)", "und", 4, 12, 120.00))

    epc = [
        ("Malla de seguridad naranja HDPE H = 1.00 m (rollo de 50 m)", "rollo",
         up(P * 1.5 / 50), 6, 95.00),
        ("Cinta de señalización amarilla \"PELIGRO\" / roja \"NO PASAR\" (rollo de 500 m)", "rollo",
         up(P * 2 / 500) + 2, 1, 42.00),
        ("Cachaco de madera/PVC 2\" h = 1.20 m con base de concreto 0.30 x 0.30 m, franjas "
         "amarillo/negro", "und", up(P / 3), 12, 38.00),
        ("Baranda rígida perimetral h = 1.00 m (pasamano, intermedio y rodapié)", "ml",
         up(P * 0.4) if rg["altura"] else up(P * 0.1), None, 32.00),
        ("Línea de vida horizontal de cable de acero Ø 1/2\" con grapas Crosby, guardacabos y "
         "templador", "ml", max(30, up(P * 0.3)) if rg["altura"] else 20, None, 28.00),
        ("Capuchones plásticos de seguridad (hongos) para mechas de acero", "und",
         max(100, no * 12), 6, 0.90),
        ("Tablero eléctrico provisional con interruptor diferencial 30 mA, termomagnéticos y "
         "puesta a tierra", "und", max(1, up(no / 25)), None, 1450.00),
        ("Extensión eléctrica industrial de 25 m con enchufes blindados", "und",
         max(2, up(no / 8)), 6, 120.00),
    ]
    if rg["excavacion"]:
        epc.append(("Entibado metálico modular / cajas de entibación (alquiler mensual)", "mes",
                    1, 1, 2800.00))
        epc.append(("Escalera de aluminio para zanjas (sobresale 1.00 m, una cada 7.50 m)", "und",
                    max(2, up(P / 60)), None, 380.00))
    if rg["altura"]:
        epc.append(("Punto de anclaje certificado ANSI Z359 (≥ 2,268 kg / 5,000 lb)", "und",
                    max(6, up(no / 3)), None, 85.00))
    if rg["confinado"]:
        epc.append(("Detector multigás de 4 gases (O2, LEL, H2S, CO) con calibración", "und", 1,
                    None, 3900.00))
        epc.append(("Ventilador/extractor axial con manga para espacios confinados", "und", 1,
                    None, 1600.00))
        epc.append(("Trípode de rescate con winche (alquiler mensual)", "mes", 1, 1, 650.00))
    if rg["izaje"]:
        epc.append(("Certificación e inspección de accesorios de izaje (estrobos, grilletes, "
                    "eslingas)", "glb", 1, 6, 900.00))
    if rg["caliente"]:
        epc.append(("Biombo / manta ignífuga para trabajos de soldadura", "und", 2, None, 350.00))

    senal = [
        ("Letrero de uso obligatorio de EPP 0.60 x 0.40 m (NTP 399.010-1)", "und",
         max(6, up(N / 5)), 12, 55.00),
        ("Letrero de prohibición (prohibido fumar, ingreso no autorizado) 0.60 x 0.40 m", "und",
         6, 12, 55.00),
        ("Letrero de advertencia (caída de objetos, excavación profunda, riesgo eléctrico, "
         "maquinaria)", "und", max(8, up(P / 40)), 12, 55.00),
        ("Letrero de evacuación, salida, zona segura en caso de sismo y punto de reunión", "und",
         8, 12, 45.00),
        ("Letrero de equipos contra incendio y botiquín", "und", 4, 12, 40.00),
        ("Cono reflectivo de 28\" con dos cintas reflectivas", "und", max(10, up(P / 15)), 12, 48.00),
        ("Tranquera / barrera portátil de madera 1.20 x 1.10 m con cinta reflectiva", "und",
         max(4, up(P / 50)), None, 180.00),
        ("Paleta de señalización PARE / SIGA", "und", 4, 6, 45.00),
        ("Luz destellante / baliza a batería para desvíos nocturnos", "und", max(4, up(P / 50)),
         None, 65.00),
        ("Mapa de riesgos de la obra impreso en vinil 1.20 x 0.90 m", "und", 1, 6, 180.00),
    ]

    capac = [
        ("Charlas diarias de 5 minutos (material, registros y facilitador)", "mes", 1, 1, 150.00),
        ("Inducción de hombre nuevo (incluye material y registro, considera rotación de personal)",
         "und", up(N * 1.3), None, 25.00),
        ("Capacitaciones específicas mensuales por especialista (altura, excavaciones, caliente, "
         "eléctrico, izaje, primeros auxilios)", "und", 1, 1, 650.00),
        ("Simulacros de emergencia (sismo, incendio, rescate en zanja / altura)", "und", 1, 3,
         450.00),
        ("Reglamento Interno de SST (RISST) y cartillas de bolsillo impresas", "und", up(N * 1.3),
         None, 8.00),
    ]

    n_ext6 = 2 + up(N / 20) + (2 if rg["caliente"] else 0)
    n_ext12 = 2
    emerg = [
        ("Botiquín de primeros auxilios equipado según Norma G.050", "und", max(1, up(N / 25)),
         None, 480.00),
        ("Reposición mensual de insumos del botiquín", "mes", max(1, up(N / 25)), 1, 90.00),
        ("Camilla rígida de rescate con correas tipo araña, inmovilizador de cabeza y collarín "
         "cervical", "und", max(1, up(N / 50)), None, 520.00),
        ("Extintor PQS ABC de 6 kg con soporte y señalización", "und", n_ext6, None, 135.00),
        ("Extintor PQS ABC de 12 kg (almacén, comedor, grupo electrógeno)", "und", n_ext12, None,
         220.00),
        ("Recarga y mantenimiento anual de extintores (NTP 350.043-1)", "und", n_ext6 + n_ext12, 12,
         45.00),
        ("Estación de emergencia (panel con extintor, botiquín, camilla, directorio y linterna)",
         "und", 1, None, 750.00),
        ("Lavaojos portátil autónomo de emergencia", "und", 1, None, 650.00),
        ("Megáfono con sirena para alarma de evacuación", "und", 1, None, 280.00),
    ]

    emo = [
        ("Examen médico ocupacional preocupacional (incluye evaluación para trabajos en altura, "
         "audiometría y espirometría)", "und", up(N * 1.2), None, 190.00),
        ("Examen médico ocupacional de retiro", "und", up(N * 1.2), None, 130.00),
        ("Póliza SCTR Salud (EsSalud / EPS) - trabajador por mes", "tr-mes", N, 1, 28.00),
        ("Póliza SCTR Pensión (compañía de seguros) - trabajador por mes", "tr-mes", N, 1, 26.00),
        ("Monitoreo de agentes ocupacionales (ruido, polvo, iluminación, ergonómico)", "und", 1, 6,
         1800.00),
    ]

    secciones = [
        ("01.01", "EQUIPOS DE PROTECCIÓN INDIVIDUAL (EPI)", epi),
        ("01.02", "EQUIPOS DE PROTECCIÓN COLECTIVA (EPC)", epc),
        ("01.03", "SEÑALIZACIÓN TEMPORAL DE SEGURIDAD", senal),
        ("01.04", "CAPACITACIÓN EN SEGURIDAD Y SALUD", capac),
        ("01.05", "RECURSOS PARA RESPUESTA ANTE EMERGENCIAS EN SEGURIDAD Y SALUD", emerg),
        ("01.06", "EXÁMENES MÉDICOS OCUPACIONALES (EMO), SCTR Y MONITOREO", emo),
    ]
    resultado = []
    total = 0.0
    for cod, titulo, items in secciones:
        filas = []
        sub = 0.0
        for k, (desc, und, base, vida, pu) in enumerate(items, 1):
            base = int(math.ceil(base))
            factor = 1 if not vida else int(math.ceil(M / float(vida)))
            cant = base * factor
            parcial = round(cant * pu, 2)
            sub += parcial
            filas.append({"codigo": "{}.{:02d}".format(cod, k), "desc": desc, "und": und,
                          "base": base, "vida": vida, "factor": factor, "cant": cant, "pu": pu,
                          "parcial": parcial})
        total += sub
        resultado.append({"codigo": cod, "titulo": titulo, "items": filas,
                          "subtotal": round(sub, 2)})
    return {"secciones": resultado, "total": round(total, 2), "cfg": cfg}


# Tabla de servicios higiénicos (Norma G.050)
TABLA_SSHH = [
    ("1 a 9", 1, 1, 1, 1), ("10 a 24", 2, 2, 2, 1), ("25 a 49", 3, 3, 3, 2),
    ("50 a 100", 5, 5, 5, 4), ("Más de 100", "+1 c/30", "+1 c/30", "+1 c/30", "+1 c/30"),
]


def sshh_requeridos(n):
    if n <= 9:
        return 1, 1, 1, 1
    if n <= 24:
        return 2, 2, 2, 1
    if n <= 49:
        return 3, 3, 3, 2
    if n <= 100:
        return 5, 5, 5, 4
    extra = math.ceil((n - 100) / 30.0)
    return 5 + extra, 5 + extra, 5 + extra, 4 + extra


# =============================================================================
# 6. PROGRAMA ANUAL / CRONOGRAMA DE CAPACITACIONES, SIMULACROS E INSPECCIONES
# =============================================================================
TEMAS_POR_FAMILIA = {
    "F04": "Trabajos en excavaciones y zanjas, entibados (PETS-01)",
    "F09": "Trabajos en altura, arnés y uso de andamios (PETS-02)",
    "F08": "Habilitado de acero, uso de amoladora y tronzadora (PETS-03)",
    "F06": "Vaciado de concreto, vibrado y manejo de cemento (PETS-04)",
    "F12": "Riesgo eléctrico y bloqueo y etiquetado LOTO (PETS-05)",
    "F10": "Trabajos en caliente, soldadura y prevención de incendios",
    "F11": "Espacios confinados y monitoreo de atmósferas",
    "F17": "Izaje de cargas, rigging y señales de grúa",
    "F02": "Demoliciones seguras y control de polvo de sílice",
    "F13": "Trabajos en vía pública, desvíos y asfalto en caliente",
    "F14": "Sustancias peligrosas, hojas MSDS/SDS y pinturas",
    "F07": "Uso seguro de sierra circular, encofrado y desencofrado",
    "F05": "Maquinaria pesada, puntos ciegos y vibración",
}
TEMAS_GENERALES = [
    "Uso, cuidado y mantenimiento de EPP",
    "IPERC continuo, ATS y derecho a negarse a un trabajo inseguro",
    "Primeros auxilios básicos y RCP",
    "Manejo de extintores y lucha contra incendios",
    "Ergonomía y manipulación manual de cargas (R.M. N° 375-2008-TR)",
    "Orden, limpieza y manejo de residuos sólidos (NTP 900.058)",
    "Plan de respuesta ante emergencias y evacuación por sismo",
]


def generar_cronograma(cfg, filas_iperc):
    cfg = completar_config(cfg)
    M = cfg["meses"]
    todos = list(range(1, M + 1))
    vistos = []
    for f in filas_iperc:
        t = TEMAS_POR_FAMILIA.get(f["familia_cod"])
        if t and t not in vistos:
            vistos.append(t)
    temas = vistos + TEMAS_GENERALES
    filas = [
        {"act": "Inducción de hombre nuevo (4 h)", "tema": "Política SST, RISST, IPERC, EPP, "
         "emergencias, derechos y obligaciones", "dirigido": "Todo trabajador nuevo",
         "resp": "Ing. SSOMA", "frec": "Al ingreso", "meses": todos},
        {"act": "Charlas diarias de 5 minutos", "tema": "Riesgos de la tarea del día y "
         "lecciones aprendidas", "dirigido": "Todo el personal", "resp": "Capataz / SSOMA",
         "frec": "Diaria", "meses": todos},
    ]
    n = len(temas)
    for i, t in enumerate(temas):
        meses = [m for m in todos if (m - 1) % n == i] or [(i % M) + 1]
        filas.append({"act": "Capacitación específica", "tema": t, "dirigido": "Cuadrillas "
                      "involucradas", "resp": "Ing. SSOMA / Especialista", "frec": "Mensual",
                      "meses": meses})
    tipos = ["Sismo y evacuación", "Amago de incendio", "Rescate en zanja / altura",
             "Primeros auxilios y traslado de herido"]
    sim = [m for m in todos if m % 3 == 0] or [M]
    for k, m in enumerate(sim):
        filas.append({"act": "Simulacro de emergencia", "tema": tipos[k % len(tipos)],
                      "dirigido": "Todo el personal", "resp": "Jefe de brigadas",
                      "frec": "Trimestral", "meses": [m]})
    filas += [
        {"act": "Inspección interna de EPP", "tema": "Estado y uso correcto de EPP",
         "dirigido": "Todo el personal", "resp": "Ing. SSOMA", "frec": "Mensual", "meses": todos},
        {"act": "Inspección de andamios, arneses y líneas de vida", "tema": "Check-list pre-uso "
         "y tarjeta verde/roja", "dirigido": "Trabajos en altura", "resp": "Ing. SSOMA",
         "frec": "Semanal", "meses": todos},
        {"act": "Inspección de herramientas y equipos eléctricos", "tema": "Cinta de color del "
         "mes", "dirigido": "Todo el personal", "resp": "Capataz / SSOMA", "frec": "Mensual",
         "meses": todos},
        {"act": "Inspección de excavaciones y entibados", "tema": "Estabilidad, accesos, acopio",
         "dirigido": "Cuadrilla de excavación", "resp": "Ing. SSOMA", "frec": "Diaria",
         "meses": todos},
        {"act": "Reunión del Comité / Supervisor de SST", "tema": "Revisión de estadísticas, "
         "incidentes y acciones correctivas", "dirigido": "Miembros del Comité",
         "resp": "Presidente del Comité", "frec": "Mensual", "meses": todos},
        {"act": "Monitoreo de agentes ocupacionales", "tema": "Ruido, polvo, iluminación, "
         "ergonómico", "dirigido": "Puestos críticos", "resp": "Empresa acreditada",
         "frec": "Semestral", "meses": sorted(set([min(2, M)] + [m for m in todos if m % 6 == 2]))},
        {"act": "Auditoría interna del PSST", "tema": "Verificación del cumplimiento del plan",
         "dirigido": "Gestión de obra", "resp": "Residente / SSOMA", "frec": "Semestral",
         "meses": sorted(set([max(1, M // 2), M]))},
        {"act": "Informe mensual de SST y estadísticas", "tema": "IF, IG, IA, HHT, inspecciones",
         "dirigido": "Entidad / Supervisión", "resp": "Ing. SSOMA", "frec": "Mensual",
         "meses": todos},
    ]
    return filas


# =============================================================================
# 7. EXPORTACIÓN EXCEL (openpyxl) - 4 HOJAS PROFESIONALES CON FÓRMULAS VIVAS
# =============================================================================
XL_AZUL = "1F3864"
XL_AZUL2 = "2E75B6"
XL_CELESTE = "DDEBF7"
XL_GRIS = "F2F2F2"
XL_VERDE = "C6EFCE"
XL_AMARILLO = "FFEB9C"
XL_ROJO = "FFC7CE"
_fino = Side(style="thin", color="7F7F7F")
BORDE = Border(left=_fino, right=_fino, top=_fino, bottom=_fino)
ALIN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIN_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIN_LT = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIN_R = Alignment(horizontal="right", vertical="center")


def _fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def _caja(ws, rango, valor=None, negrita=False, tam=9, color_txt="000000", fondo=None,
          alin=ALIN_C, borde=True):
    """Combina un rango, escribe el valor y aplica bordes a todas sus celdas."""
    if ":" in rango:
        ws.merge_cells(rango)
        celdas = ws[rango]
    else:
        celdas = ((ws[rango],),)
    primera = celdas[0][0]
    if valor is not None:
        primera.value = valor
    primera.font = Font(name="Arial", bold=negrita, size=tam, color=color_txt)
    primera.alignment = alin
    for fila in celdas:
        for c in fila:
            if borde:
                c.border = BORDE
            if fondo:
                c.fill = _fill(fondo)
    return primera


def _encabezado_institucional(ws, cfg, titulo, codigo, ultima_col):
    lc = get_column_letter(ultima_col)
    _caja(ws, "A1:C4", "SISTEMA DE GESTIÓN\nDE SEGURIDAD Y\nSALUD EN EL TRABAJO", True, 9,
          "FFFFFF", XL_AZUL)
    c_ini = get_column_letter(4)
    c_fin = get_column_letter(ultima_col - 3)
    _caja(ws, "{}1:{}2".format(c_ini, c_fin), titulo, True, 14, "FFFFFF", XL_AZUL2)
    _caja(ws, "{}3:{}3".format(c_ini, c_fin), "OBRA: " + cfg["proyecto"], True, 9,
          fondo=XL_CELESTE)
    _caja(ws, "{}4:{}4".format(c_ini, c_fin),
          "ENTIDAD: {}   |   UBICACIÓN: {}   |   MODALIDAD: {}".format(
              cfg["entidad"], cfg["ubicacion"], cfg["modalidad"]), False, 8, fondo=XL_CELESTE)
    c3 = get_column_letter(ultima_col - 2)
    _caja(ws, "{}1:{}1".format(c3, lc), "Código: " + codigo, True, 8, alin=ALIN_L)
    _caja(ws, "{}2:{}2".format(c3, lc), "Versión: 01", False, 8, alin=ALIN_L)
    _caja(ws, "{}3:{}3".format(c3, lc), "Fecha: " + cfg["fecha"], False, 8, alin=ALIN_L)
    _caja(ws, "{}4:{}4".format(c3, lc), "Página: 1 de 1", False, 8, alin=ALIN_L)
    for r in range(1, 5):
        ws.row_dimensions[r].height = 20


def _config_impresion(ws, horizontal=True, area=None, filas_rep=None):
    ws.page_setup.orientation = "landscape" if horizontal else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    if area:
        ws.print_area = area
    if filas_rep:
        ws.print_title_rows = filas_rep


def _hoja_iperc(wb, cfg, filas):
    ws = wb.active
    ws.title = "1_Matriz_IPERC_Obra"
    ws.sheet_view.showGridLines = False
    NC = 32
    _encabezado_institucional(ws, cfg, "MATRIZ DE IDENTIFICACIÓN DE PELIGROS, EVALUACIÓN DE "
                              "RIESGOS Y MEDIDAS DE CONTROL (IPERC) - NORMA G.050 / R.M. 050-2013-TR",
                              "SST-IPERC-01", NC)
    _caja(ws, "A5:H5", "Residente de obra: {}".format(cfg["residente"]), False, 8, alin=ALIN_L)
    _caja(ws, "I5:Q5", "Especialista SSOMA: {}  CIP N° {}".format(cfg["ssoma"], cfg["cip"]),
          False, 8, alin=ALIN_L)
    _caja(ws, "R5:AF5", "N° obreros: {}   |   Staff técnico: {}   |   Plazo: {} días ({} meses)"
          .format(cfg["n_obr"], cfg["n_ing"], cfg["dias"], cfg["meses"]), False, 8, alin=ALIN_L)

    r1, r2 = 7, 8
    grupos = [("A", "H", "IDENTIFICACIÓN DE PELIGROS Y RIESGOS", XL_AZUL),
              ("I", "Q", "EVALUACIÓN DE RIESGO INICIAL", "C00000"),
              ("R", "V", "MEDIDAS DE CONTROL (JERARQUÍA - ART. 21 LEY 29783)", XL_AZUL2),
              ("W", "AD", "EVALUACIÓN DE RIESGO RESIDUAL", "548235"),
              ("AE", "AF", "CONTROL OPERACIONAL", XL_AZUL)]
    for a, b, t, col in grupos:
        _caja(ws, "{}{}:{}{}".format(a, r1, b, r1), t, True, 9, "FFFFFF", col)
    cab = ["N°", "Ítem", "Partida / Actividad", "Familia de riesgo", "Tipo (R / NR)", "Peligro",
           "Riesgo (evento peligroso)", "Consecuencia / Daño",
           "A\nPersonas expuestas", "B\nProcedi-\nmientos", "C\nCapaci-\ntación",
           "D\nExposición", "Índice de Probabilidad\nIP = A+B+C+D", "S\nSeveridad",
           "Riesgo\nP x S", "Nivel de riesgo", "Riesgo signifi-\ncativo",
           "Eliminación", "Sustitución", "Controles de ingeniería", "Controles administrativos",
           "EPP específico",
           "A", "B", "C", "D", "IP", "S", "P x S", "Nivel residual",
           "PETAR / ATS requerido", "PETS aplicable"]
    for j, t in enumerate(cab, 1):
        c = ws.cell(row=r2, column=j, value=t)
        c.font = Font(name="Arial", bold=True, size=8, color="FFFFFF")
        c.fill = _fill("404040")
        c.alignment = ALIN_C
        c.border = BORDE
    ws.row_dimensions[r2].height = 54
    anchos = [4, 7, 30, 20, 9, 30, 28, 26, 7, 7, 7, 7, 10, 7, 7, 13, 8, 28, 26, 32, 32, 30,
              5, 5, 5, 5, 5, 5, 6, 12, 18, 22]
    for j, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    nivel_f = ('IF({0}<=4,"TRIVIAL",IF({0}<=8,"TOLERABLE",IF({0}<=16,"MODERADO",'
               'IF({0}<=24,"IMPORTANTE","INTOLERABLE"))))')
    colores_nivel = {"TRIVIAL": "C6EFCE", "TOLERABLE": "92D050", "MODERADO": "FFFF00",
                     "IMPORTANTE": "FF9900", "INTOLERABLE": "FF0000"}
    r = r2 + 1
    for k, f in enumerate(filas, 1):
        valores = [k, f["item"], f["partida"], f["familia"], f["tipo"], f["peligro"], f["riesgo"],
                   f["consecuencia"], f["A"], f["B"], f["C"], f["D"],
                   "=SUM(I{0}:L{0})".format(r), f["S"], "=M{0}*N{0}".format(r),
                   "=" + nivel_f.format("O{}".format(r)),
                   '=IF(O{0}>16,"SÍ","NO")'.format(r),
                   f["eliminacion"], f["sustitucion"], f["ingenieria"], f["administrativo"],
                   f["epp"], f["rA"], f["rB"], f["rC"], f["rD"],
                   "=SUM(W{0}:Z{0})".format(r), f["rS"], "=AA{0}*AB{0}".format(r),
                   "=" + nivel_f.format("AC{}".format(r)), f["petar"], f["pets"]]
        for j, v in enumerate(valores, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = Font(name="Arial", size=8, bold=(j in (13, 15, 16, 27, 29, 30)))
            c.alignment = ALIN_C if (j in (1, 2, 5) or 9 <= j <= 17 or 23 <= j <= 30) else ALIN_L
            c.border = BORDE
        # colores estáticos (visibles aun sin recalcular) + formato condicional
        ws.cell(row=r, column=16).fill = _fill(colores_nivel[f["nivel"]])
        ws.cell(row=r, column=30).fill = _fill(colores_nivel[f["nivel_res"]])
        ws.cell(row=r, column=15).fill = _fill(XL_ROJO if f["PxS"] > 16 else XL_AMARILLO)
        ws.cell(row=r, column=29).fill = _fill(XL_VERDE if f["rPxS"] <= 8 else XL_AMARILLO)
        ws.row_dimensions[r].height = 96
        r += 1
    ult = r - 1
    rojo = _fill("FF0000")
    naranja = _fill("FF9900")
    amarillo = _fill("FFFF00")
    verde = _fill("92D050")
    for col in ("P", "AD"):
        rng = "{0}{1}:{0}{2}".format(col, r2 + 1, ult)
        p = "{}{}".format(col, r2 + 1)
        ws.conditional_formatting.add(rng, FormulaRule(formula=['{}="INTOLERABLE"'.format(p)],
                                                       fill=rojo, font=Font(color="FFFFFF",
                                                                            bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=['{}="IMPORTANTE"'.format(p)],
                                                       fill=naranja))
        ws.conditional_formatting.add(rng, FormulaRule(formula=['{}="MODERADO"'.format(p)],
                                                       fill=amarillo))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=['OR({0}="TOLERABLE",{0}="TRIVIAL")'.format(p)], fill=verde))
    ws.conditional_formatting.add("O{}:O{}".format(r2 + 1, ult), FormulaRule(
        formula=["O{}>16".format(r2 + 1)], fill=_fill(XL_ROJO)))
    ws.conditional_formatting.add("AC{}:AC{}".format(r2 + 1, ult), FormulaRule(
        formula=["AC{}<=8".format(r2 + 1)], fill=_fill(XL_VERDE)))
    ws.conditional_formatting.add("AC{}:AC{}".format(r2 + 1, ult), FormulaRule(
        formula=["AND(AC{0}>8,AC{0}<=16)".format(r2 + 1)], fill=_fill(XL_AMARILLO)))

    # Leyendas de valoración
    r = ult + 2
    _caja(ws, "A{0}:H{0}".format(r), "CRITERIOS DE EVALUACIÓN (R.M. N° 050-2013-TR)", True, 9,
          "FFFFFF", XL_AZUL)
    _caja(ws, "R{0}:V{0}".format(r), "NIVELES DE RIESGO E INTERPRETACIÓN", True, 9, "FFFFFF",
          XL_AZUL)
    rr = r + 1
    for clave in ("A", "B", "C", "D", "S"):
        titulo, opciones = CRITERIOS_IPERC[clave]
        _caja(ws, "A{0}:B{0}".format(rr), clave, True, 9, fondo=XL_CELESTE)
        _caja(ws, "C{0}:H{0}".format(rr), "{}:  {}".format(titulo, "  |  ".join(opciones)),
              False, 8, alin=ALIN_L)
        ws.row_dimensions[rr].height = 30
        rr += 1
    rr = r + 1
    for nombre, rango, accion, color in NIVELES_RIESGO:
        _caja(ws, "R{0}".format(rr), nombre, True, 8, fondo=color)
        _caja(ws, "S{0}".format(rr), rango, True, 8)
        _caja(ws, "T{0}:V{0}".format(rr), accion, False, 8, alin=ALIN_L)
        rr += 1
    rf = max(rr, r + 6) + 2
    for a, b, t in (("A", "F", "Elaborado por: Especialista SSOMA"),
                    ("I", "Q", "Revisado por: Residente de Obra"),
                    ("R", "V", "Aprobado por: Supervisión / Entidad")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, rf, b), "\n\n\n______________________________\n" + t,
              False, 8)
        ws.row_dimensions[rf].height = 60
    ws.freeze_panes = "D9"
    _config_impresion(ws, True, "A1:AF{}".format(rf), "7:8")


def _hoja_presupuesto(wb, cfg, desag):
    ws = wb.create_sheet("2_Presupuesto_Seguridad_G050")
    ws.sheet_view.showGridLines = False
    for j, w in enumerate([11, 70, 8, 11, 11, 11, 11, 12, 15], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.merge_cells("A1:I1")
    ws["A1"] = "PRESUPUESTO DESAGREGADO DE SEGURIDAD Y SALUD EN EL TRABAJO - NORMA G.050"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = _fill(XL_AZUL)
    ws["A1"].alignment = ALIN_C
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:I2")
    ws["A2"] = "OBRA: {}  |  ENTIDAD: {}  |  {}".format(cfg["proyecto"], cfg["entidad"],
                                                         cfg["ubicacion"])
    ws["A2"].font = Font(name="Arial", size=9, bold=True)
    ws["A2"].alignment = ALIN_C
    ws["A2"].fill = _fill(XL_CELESTE)
    ws.row_dimensions[2].height = 30
    params = [("N° de obreros (Nobr)", cfg["n_obr"]), ("N° de ingenieros / staff (Ning)",
                                                        cfg["n_ing"]),
              ("Plazo de ejecución (meses) M", cfg["meses"]),
              ("Total personal (Nobr + Ning)", "=C4+C5")]
    for k, (t, v) in enumerate(params):
        rr = 4 + k
        _caja(ws, "A{0}:B{0}".format(rr), t, True, 9, fondo=XL_GRIS, alin=ALIN_L)
        c = _caja(ws, "C{0}".format(rr), v, True, 10, "C00000")
    _caja(ws, "E4:I7", "Factor de reposición = REDONDEAR.MAS(M / Vida útil)\n"
          "Cantidad = Base x Factor de reposición\nParcial = Cantidad x Precio unitario\n"
          "(Modifique M o los precios y el presupuesto se recalcula)", False, 8, fondo=XL_AMARILLO,
          alin=ALIN_L)
    cab = ["Ítem", "Descripción", "Und.", "Base\n(pers./cant.)", "Vida útil\n(meses)",
           "Factor\nreposición", "Cantidad", "P.U.\n(S/)", "Parcial\n(S/)"]
    rh = 9
    for j, t in enumerate(cab, 1):
        c = ws.cell(row=rh, column=j, value=t)
        c.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        c.fill = _fill("404040")
        c.alignment = ALIN_C
        c.border = BORDE
    ws.row_dimensions[rh].height = 32
    r = rh + 1
    _caja(ws, "A{0}".format(r), "01", True, 10, "FFFFFF", XL_AZUL)
    _caja(ws, "B{0}:I{0}".format(r), "SEGURIDAD Y SALUD EN EL TRABAJO (NORMA G.050 / "
          "D.S. N° 011-2019-TR)", True, 10, "FFFFFF", XL_AZUL, ALIN_L)
    r += 1
    celdas_sub = []
    resumen = []
    for sec in desag["secciones"]:
        r_sec = r
        _caja(ws, "A{}".format(r), sec["codigo"], True, 9, fondo=XL_CELESTE)
        _caja(ws, "B{0}:H{0}".format(r), sec["titulo"], True, 9, fondo=XL_CELESTE, alin=ALIN_L)
        r += 1
        r_ini = r
        for it in sec["items"]:
            vals = [it["codigo"], it["desc"], it["und"], it["base"],
                    it["vida"] if it["vida"] else "",
                    '=IF(E{0}="",1,ROUNDUP($C$6/E{0},0))'.format(r), "=D{0}*F{0}".format(r),
                    it["pu"], "=G{0}*H{0}".format(r)]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.font = Font(name="Arial", size=8)
                c.border = BORDE
                c.alignment = ALIN_L if j == 2 else ALIN_C
                if j in (8, 9):
                    c.number_format = '#,##0.00'
                    c.alignment = ALIN_R
            ws.row_dimensions[r].height = 24
            r += 1
        sub = _caja(ws, "I{}".format(r_sec), "=SUM(I{}:I{})".format(r_ini, r - 1), True, 9,
                    fondo=XL_CELESTE, alin=ALIN_R)
        sub.number_format = '#,##0.00'
        celdas_sub.append("I{}".format(r_sec))
        resumen.append((sec["codigo"], sec["titulo"], "I{}".format(r_sec)))
    _caja(ws, "A{0}:H{0}".format(r), "COSTO DIRECTO TOTAL DE SEGURIDAD Y SALUD EN EL TRABAJO (S/)",
          True, 11, "FFFFFF", "C00000", ALIN_R)
    tot = _caja(ws, "I{}".format(r), "=" + "+".join(celdas_sub), True, 11, "FFFFFF", "C00000",
                ALIN_R)
    tot.number_format = '#,##0.00'
    celda_total = "I{}".format(r)
    ws.row_dimensions[r].height = 26
    r += 2
    _caja(ws, "A{0}:I{0}".format(r), "RESUMEN DEL PRESUPUESTO DE SEGURIDAD", True, 10, "FFFFFF",
          XL_AZUL)
    r += 1
    for t, rng in (("Ítem", "A"), ("Subpartida G.050", "B{0}:G{0}"), ("Monto (S/)", "H"),
                   ("% del total", "I")):
        _caja(ws, (rng.format(r) if "{" in rng else "{}{}".format(rng, r)), t, True, 9, "FFFFFF",
              "404040")
    r += 1
    for cod, tit, ref in resumen:
        _caja(ws, "A{}".format(r), cod, False, 9)
        _caja(ws, "B{0}:G{0}".format(r), tit, False, 9, alin=ALIN_L)
        c = _caja(ws, "H{}".format(r), "={}".format(ref), True, 9, alin=ALIN_R)
        c.number_format = '#,##0.00'
        c = _caja(ws, "I{}".format(r), "=IF({1}=0,0,{0}/{1})".format(ref, celda_total), False, 9,
                  alin=ALIN_R)
        c.number_format = '0.00%'
        r += 1
    _caja(ws, "A{0}:G{0}".format(r), "TOTAL", True, 10, fondo=XL_CELESTE, alin=ALIN_R)
    c = _caja(ws, "H{}".format(r), "={}".format(celda_total), True, 10, fondo=XL_CELESTE,
              alin=ALIN_R)
    c.number_format = '#,##0.00'
    c = _caja(ws, "I{}".format(r), "=SUM(I{}:I{})".format(r - len(resumen), r - 1), True, 10,
              fondo=XL_CELESTE, alin=ALIN_R)
    c.number_format = '0.00%'
    r += 2
    ws.merge_cells("A{0}:I{0}".format(r))
    ws["A{}".format(r)] = ("Nota: Precios referenciales del mercado peruano (S/) sin IGV. El costo "
                           "de seguridad G.050 se incluye como partida de costo directo del "
                           "presupuesto de obra. Personal SSOMA se considera en gastos generales.")
    ws["A{}".format(r)].font = Font(name="Arial", size=8, italic=True)
    ws["A{}".format(r)].alignment = ALIN_L
    ws.row_dimensions[r].height = 28
    ws.freeze_panes = "A10"
    _config_impresion(ws, False, "A1:I{}".format(r), "9:9")


def _hoja_cronograma(wb, cfg, crono):
    ws = wb.create_sheet("3_Programa_Capacitaciones")
    ws.sheet_view.showGridLines = False
    M = cfg["meses"]
    NC = 6 + M + 1
    lc = get_column_letter(NC)
    for j, w in enumerate([5, 30, 42, 20, 20, 12], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for m in range(M):
        ws.column_dimensions[get_column_letter(7 + m)].width = 5.5
    ws.column_dimensions[lc].width = 8
    ws.merge_cells("A1:{}1".format(lc))
    ws["A1"] = "PROGRAMA DE CAPACITACIÓN, SENSIBILIZACIÓN, SIMULACROS E INSPECCIONES INTERNAS DE SST"
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    ws["A1"].fill = _fill(XL_AZUL)
    ws["A1"].alignment = ALIN_C
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:{}2".format(lc))
    ws["A2"] = "OBRA: {}  |  Plazo: {} días calendario ({} meses)  |  Personal: {} obreros + {} " \
               "staff".format(cfg["proyecto"], cfg["dias"], M, cfg["n_obr"], cfg["n_ing"])
    ws["A2"].font = Font(name="Arial", size=9, bold=True)
    ws["A2"].fill = _fill(XL_CELESTE)
    ws["A2"].alignment = ALIN_C
    ws.row_dimensions[2].height = 30
    rh = 4
    cab = ["N°", "Actividad", "Tema / Descripción", "Dirigido a", "Responsable", "Frecuencia"]
    cab += ["M{}".format(m + 1) for m in range(M)] + ["Total"]
    for j, t in enumerate(cab, 1):
        c = ws.cell(row=rh, column=j, value=t)
        c.font = Font(name="Arial", bold=True, size=9, color="FFFFFF")
        c.fill = _fill("404040")
        c.alignment = ALIN_C
        c.border = BORDE
    ws.row_dimensions[rh].height = 24
    colores = {"Inducción": "BDD7EE", "Charlas": "BDD7EE", "Capacitación": "C6EFCE",
               "Simulacro": "F8CBAD", "Inspección": "FFE699", "Reunión": "D9D2E9",
               "Monitoreo": "FCE4D6", "Auditoría": "D9D2E9", "Informe": "EDEDED"}
    r = rh + 1
    for k, f in enumerate(crono, 1):
        vals = [k, f["act"], f["tema"], f["dirigido"], f["resp"], f["frec"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = Font(name="Arial", size=8)
            c.alignment = ALIN_L if j in (2, 3, 4, 5) else ALIN_C
            c.border = BORDE
        color = next((v for kk, v in colores.items() if f["act"].startswith(kk)), "C6EFCE")
        for m in range(1, M + 1):
            c = ws.cell(row=r, column=6 + m, value="X" if m in f["meses"] else "")
            c.border = BORDE
            c.alignment = ALIN_C
            c.font = Font(name="Arial", size=8, bold=True)
            if m in f["meses"]:
                c.fill = _fill(color)
        c = ws.cell(row=r, column=NC, value='=COUNTIF({}{}:{}{},"X")'.format(
            get_column_letter(7), r, get_column_letter(6 + M), r))
        c.border = BORDE
        c.alignment = ALIN_C
        c.font = Font(name="Arial", size=8, bold=True)
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1
    ws.merge_cells("A{0}:{1}{0}".format(r, lc))
    ws["A{}".format(r)] = ("Leyenda: X = actividad programada en el mes. Registro según R.M. "
                           "N° 050-2013-TR (Registro de inducción, capacitación, entrenamiento y "
                           "simulacros de emergencia; Registro de inspecciones internas).")
    ws["A{}".format(r)].font = Font(name="Arial", size=8, italic=True)
    ws["A{}".format(r)].alignment = ALIN_L
    ws.row_dimensions[r].height = 28
    ws.freeze_panes = "G5"
    _config_impresion(ws, True, "A1:{}{}".format(lc, r), "4:4")


def _hoja_formatos(wb, cfg):
    ws = wb.create_sheet("4_Formatos_ATS_y_PETAR")
    ws.sheet_view.showGridLines = False
    for j in range(1, 13):
        ws.column_dimensions[get_column_letter(j)].width = 11.5
    ob = cfg["proyecto"]

    def titulo(r, texto, codigo):
        _caja(ws, "A{0}:B{1}".format(r, r + 1), "SG-SST\nOBRA", True, 9, "FFFFFF", XL_AZUL)
        _caja(ws, "C{0}:J{1}".format(r, r + 1), texto, True, 12, "FFFFFF", XL_AZUL2)
        _caja(ws, "K{0}:L{0}".format(r), "Código: " + codigo, True, 8, alin=ALIN_L)
        _caja(ws, "K{0}:L{0}".format(r + 1), "Versión: 01", False, 8, alin=ALIN_L)
        ws.row_dimensions[r].height = 20
        ws.row_dimensions[r + 1].height = 20
        _caja(ws, "A{0}:L{0}".format(r + 2), "Obra: " + ob, True, 8, fondo=XL_CELESTE,
              alin=ALIN_L)
        return r + 3

    def banda(r, texto, color="404040"):
        _caja(ws, "A{0}:L{0}".format(r), texto, True, 9, "FFFFFF", color, ALIN_L)
        return r + 1

    # ---------------- ATS ----------------
    r = titulo(1, "ANÁLISIS DE TRABAJO SEGURO (ATS)", "SST-ATS-01")
    for a, b, t in (("A", "D", "Fecha:"), ("E", "H", "Hora inicio:      Hora fin:"),
                    ("I", "L", "Frente / Sector:")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, False, 8, alin=ALIN_L)
    r += 1
    _caja(ws, "A{0}:H{0}".format(r), "Trabajo a realizar:", False, 8, alin=ALIN_L)
    _caja(ws, "I{0}:L{0}".format(r), "Supervisor / Capataz:", False, 8, alin=ALIN_L)
    r += 1
    _caja(ws, "A{0}:L{0}".format(r), "EPP requerido:  ☐ Casco  ☐ Lentes  ☐ Guantes  ☐ Botín  "
          "☐ Arnés  ☐ Respirador  ☐ Protector auditivo  ☐ Careta  ☐ Botas de jebe  ☐ Chaleco",
          False, 8, alin=ALIN_L)
    r += 1
    r = banda(r, "SECUENCIA DE PASOS, PELIGROS, RIESGOS Y MEDIDAS DE CONTROL")
    for a, b, t in (("A", "A", "N°"), ("B", "D", "Pasos de la tarea"),
                    ("E", "G", "Peligros / Riesgos"), ("H", "K", "Medidas de control"),
                    ("L", "L", "Responsable")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, True, 8, fondo=XL_GRIS)
    r += 1
    for k in range(1, 9):
        for a, b, t in (("A", "A", str(k)), ("B", "D", ""), ("E", "G", ""), ("H", "K", ""),
                        ("L", "L", "")):
            _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, False, 8)
        ws.row_dimensions[r].height = 24
        r += 1
    r = banda(r, "PERSONAL EJECUTANTE (declaro haber participado en la elaboración del ATS)")
    for a, b, t in (("A", "A", "N°"), ("B", "E", "Apellidos y nombres"), ("F", "G", "DNI"),
                    ("H", "I", "Ocupación"), ("J", "L", "Firma")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, True, 8, fondo=XL_GRIS)
    r += 1
    for k in range(1, 7):
        for a, b, t in (("A", "A", str(k)), ("B", "E", ""), ("F", "G", ""), ("H", "I", ""),
                        ("J", "L", "")):
            _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, False, 8)
        ws.row_dimensions[r].height = 20
        r += 1
    for a, b, t in (("A", "D", "Capataz / Supervisor"), ("E", "H", "Ing. SSOMA"),
                    ("I", "L", "Residente de obra")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), "\n\n_______________________\n" + t, False, 8)
    ws.row_dimensions[r].height = 50
    fin_ats = r
    ws.row_breaks.append(_salto(fin_ats))

    # ---------------- PETAR ----------------
    r = titulo(fin_ats + 2, "PERMISO ESCRITO PARA TRABAJO DE ALTO RIESGO (PETAR)", "SST-PETAR-01")
    _caja(ws, "A{0}:L{0}".format(r), "Tipo de trabajo:  ☐ Altura > 1.80 m   ☐ Excavación > 1.50 m"
          "   ☐ Caliente   ☐ Espacio confinado   ☐ Izaje de cargas   ☐ Eléctrico (LOTO)",
          True, 8, alin=ALIN_L)
    r += 1
    for a, b, t in (("A", "F", "Descripción del trabajo:"), ("G", "L", "Ubicación / Área:")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, False, 8, alin=ALIN_LT)
    ws.row_dimensions[r].height = 30
    r += 1
    for a, b, t in (("A", "D", "Fecha:"), ("E", "H", "Hora inicio:"), ("I", "L", "Hora fin "
                                                                               "(máx. 1 turno):")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, False, 8, alin=ALIN_L)
    r += 1
    r = banda(r, "VERIFICACIONES PREVIAS AL INICIO DEL TRABAJO")
    for a, b, t in (("A", "I", "Ítem de verificación"), ("J", "J", "SÍ"), ("K", "K", "NO"),
                    ("L", "L", "N/A")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, True, 8, fondo=XL_GRIS)
    r += 1
    verif = [
        "ATS elaborado y firmado por todo el personal ejecutante",
        "Personal capacitado, con EMO vigente y SCTR activo",
        "Área delimitada y señalizada (cachacos, malla, cinta, letreros)",
        "Arnés, línea de vida y punto de anclaje inspeccionados (altura)",
        "Andamio con tarjeta verde, plataforma completa, baranda 0.90 m y rodapié 0.15 m",
        "Entibado instalado, acopio ≥ H/2 (mín. 0.60 m) y escalera cada 7.50 m (excavación)",
        "Redes subterráneas identificadas / interferencias liberadas",
        "Extintor PQS operativo en el punto y vigía contra incendio (caliente)",
        "Combustibles retirados en un radio de 11 m (caliente)",
        "Medición de atmósfera: O2 19.5-23.5 %, LEL < 10 %, H2S < 10 ppm, CO < 25 ppm (confinado)",
        "Ventilación forzada y vigía en la boca del espacio confinado",
        "Plan de rigging, grúa certificada y accesorios inspeccionados (izaje)",
        "Bloqueo y etiquetado LOTO aplicado y cero energía verificada (eléctrico)",
        "Plan de rescate y equipo de emergencia disponible",
    ]
    for t in verif:
        _caja(ws, "A{0}:I{0}".format(r), t, False, 8, alin=ALIN_L)
        for col in "JKL":
            _caja(ws, "{}{}".format(col, r), "☐", False, 9)
        r += 1
    r = banda(r, "EPP Y EQUIPOS ESPECIALES REQUERIDOS")
    _caja(ws, "A{0}:L{0}".format(r), "☐ Arnés doble línea  ☐ Línea de vida  ☐ Detector multigás  "
          "☐ Trípode de rescate  ☐ Careta de soldar  ☐ Extintor PQS  ☐ Guantes dieléctricos  "
          "☐ Candados/tarjetas LOTO  ☐ Respirador  ☐ Otros: ________", False, 8, alin=ALIN_L)
    ws.row_dimensions[r].height = 30
    r += 1
    r = banda(r, "PERSONAL AUTORIZADO")
    for a, b, t in (("A", "A", "N°"), ("B", "E", "Apellidos y nombres"), ("F", "G", "DNI"),
                    ("H", "I", "Ocupación"), ("J", "L", "Firma")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, True, 8, fondo=XL_GRIS)
    r += 1
    for k in range(1, 6):
        for a, b, t in (("A", "A", str(k)), ("B", "E", ""), ("F", "G", ""), ("H", "I", ""),
                        ("J", "L", "")):
            _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, False, 8)
        ws.row_dimensions[r].height = 20
        r += 1
    r = banda(r, "AUTORIZACIONES (sin todas las firmas el trabajo NO puede iniciar)", "C00000")
    for a, b, t in (("A", "D", "Supervisor del trabajo"), ("E", "H", "Ing. SSOMA"),
                    ("I", "L", "Residente de obra")):
        _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), "\n\n_______________________\n" + t, False, 8)
    ws.row_dimensions[r].height = 50
    r += 1
    _caja(ws, "A{0}:L{0}".format(r), "CIERRE DEL PERMISO:  Hora: ______   Área ordenada y segura "
          "☐   Equipos retirados ☐   Firma del supervisor: ____________________", False, 8,
          alin=ALIN_L)
    fin_petar = r
    ws.row_breaks.append(_salto(fin_petar))

    # ---------------- CHECK-LISTS ----------------
    r = fin_petar + 2
    listas = [
        ("CHECK-LIST PRE-USO DE ARNÉS DE CUERPO COMPLETO Y LÍNEA DE VIDA", "SST-CHK-01", [
            "Etiqueta de certificación ANSI Z359 legible", "Correas sin cortes, quemaduras ni "
            "deshilachados", "Costuras completas y sin desgaste", "Hebillas y anillos D sin "
            "deformación ni corrosión", "Ganchos con doble seguro operativos",
            "Absorbedor de impacto sin activar", "Línea de vida sin nudos ni cortes",
            "Fecha de inspección vigente (cinta de color del mes)"]),
        ("CHECK-LIST DE INSPECCIÓN DE ANDAMIO MULTIDIRECCIONAL", "SST-CHK-02", [
            "Placas base y husillos sobre terreno firme / tablones de reparto",
            "Verticales aplomados y crucetas/diagonales completas",
            "Plataforma completa sin aberturas, con seguros", "Baranda superior a 0.90 m e "
            "intermedia a 0.45 m", "Rodapié de 0.15 m en todo el perímetro",
            "Andamio arriostrado a la estructura cada 4 m de altura",
            "Escalera interna de acceso con trampilla", "Tarjeta de inspección VERDE colocada"]),
        ("CHECK-LIST DE INSPECCIÓN DE EXCAVACIONES Y ZANJAS", "SST-CHK-03", [
            "Profundidad > 1.50 m con entibado o talud adecuado", "Material excavado a ≥ H/2 "
            "(mín. 0.60 m) del borde", "Escaleras de acceso cada 7.50 m sobresaliendo 1.00 m",
            "Cachacos, malla naranja y cinta a 1.00 m del borde", "Sin filtraciones ni grietas "
            "en las paredes", "Maquinaria a distancia segura del borde",
            "Redes subterráneas identificadas y señalizadas", "Pasarelas con baranda para "
            "cruce de zanjas"]),
    ]
    for tit, cod, items in listas:
        r = titulo(r, tit, cod)
        for a, b, t in (("A", "A", "N°"), ("B", "H", "Ítem a verificar"), ("I", "I", "SÍ"),
                        ("J", "J", "NO"), ("K", "L", "Observación")):
            _caja(ws, "{0}{1}:{2}{1}".format(a, r, b), t, True, 8, fondo=XL_GRIS)
        r += 1
        for k, t in enumerate(items, 1):
            _caja(ws, "A{}".format(r), k, False, 8)
            _caja(ws, "B{0}:H{0}".format(r), t, False, 8, alin=ALIN_L)
            _caja(ws, "I{}".format(r), "☐", False, 9)
            _caja(ws, "J{}".format(r), "☐", False, 9)
            _caja(ws, "K{0}:L{0}".format(r), "", False, 8)
            r += 1
        _caja(ws, "A{0}:F{0}".format(r), "Inspeccionado por:", False, 8, alin=ALIN_L)
        _caja(ws, "G{0}:L{0}".format(r), "Firma / Fecha:", False, 8, alin=ALIN_L)
        r += 2
    _config_impresion(ws, False, "A1:L{}".format(r))


def _salto(fila):
    from openpyxl.worksheet.pagebreak import Break
    return Break(id=fila)


def exportar_excel(ruta, cfg, filas_iperc, desag, crono):
    cfg = completar_config(cfg)
    wb = Workbook()
    _hoja_iperc(wb, cfg, filas_iperc)
    _hoja_presupuesto(wb, cfg, desag)
    _hoja_cronograma(wb, cfg, crono)
    _hoja_formatos(wb, cfg)
    wb.properties.title = "Matriz IPERC y Presupuesto SSOMA G.050"
    wb.properties.creator = APP_CORTO
    wb.save(ruta)
    return ruta


# =============================================================================
# 8. PROCEDIMIENTOS ESCRITOS DE TRABAJO SEGURO (PETS)
# =============================================================================
PETS_LISTA = [
    {
        "codigo": "PETS-01", "titulo": "EXCAVACIÓN DE ZANJAS Y ENTIBADOS",
        "objetivo": "Establecer la secuencia segura para la excavación manual y mecánica de zanjas "
                    "y su entibado, previniendo derrumbes, sepultamientos, atropellos por "
                    "maquinaria y contactos con redes subterráneas.",
        "alcance": "Aplica a toda excavación de zanjas, cimientos, buzones y calzaduras, en "
                   "especial a las de profundidad mayor a 1.50 m que requieren PETAR.",
        "responsables": [
            "Residente de obra: aprueba el PETS, garantiza recursos (entibado, escaleras, malla).",
            "Ing. SSOMA: verifica el PETAR, inspecciona diariamente la zanja y capacita.",
            "Capataz: dirige la tarea, elabora el ATS con la cuadrilla y controla el acopio.",
            "Operador de retroexcavadora: opera con check-list diario y respeta al vigía.",
            "Vigía: controla la maquinaria y el ingreso de personas al área de exclusión.",
        ],
        "epp": ["Casco con barbiquejo", "Lentes de seguridad", "Guantes de cuero",
                "Botín punta de acero", "Chaleco reflectivo", "Protector auditivo",
                "Respirador para polvo", "Arnés con línea de rescate (H > 2.00 m)"],
        "equipos": ["Retroexcavadora / excavadora", "Detector de servicios subterráneos",
                    "Entibado metálico modular o tablones 2\"x8\", largueros 4\"x4\" y puntales",
                    "Escaleras de aluminio", "Cachacos, malla naranja y cinta de peligro",
                    "Picos, lampas y barretas en buen estado", "Pasarelas con barandas"],
        "pasos": [
            ("Revisar planos de interferencias, solicitar catastro de redes y obtener PETAR de "
             "excavación.", "Contacto con redes eléctricas, gas o agua", "Detector de servicios, "
             "calicatas manuales de ubicación, PETAR firmado."),
            ("Elaborar el ATS con toda la cuadrilla y verificar EPP.", "Desconocimiento de "
             "peligros", "Charla de 5 minutos, ATS firmado, verificación de EMO y SCTR."),
            ("Delimitar el área con cachacos, malla naranja y cinta a 1.00 m del borde.",
             "Caída de personas a la zanja", "Barrera continua, letreros de excavación profunda, "
             "iluminación nocturna."),
            ("Trazar el eje de la zanja y definir la zona de acopio a ≥ H/2 (mín. 0.60 m).",
             "Sobrecarga del borde", "Marcar con yeso la franja de acopio; prohibido acopiar "
             "dentro de ella."),
            ("Excavar con retroexcavadora con vigía; nadie dentro del radio de giro.",
             "Golpe / atropello por maquinaria", "Vigía con paleta, alarma de retroceso, "
             "distancia mínima de 3 m."),
            ("Al superar 1.50 m de profundidad, instalar el entibado de arriba hacia abajo antes "
             "del ingreso del personal.", "Derrumbe y sepultamiento", "Entibado continuo, "
             "largueros y puntales cada 1.00 m, cajas de entibación."),
            ("Colocar escaleras de acceso cada 7.50 m sobresaliendo 1.00 m del borde.",
             "Atrapamiento sin vía de escape", "Escaleras amarradas en la parte superior."),
            ("Realizar el perfilado manual del fondo con un máximo de 2 trabajadores por tramo, "
             "separados 3 m.", "Golpes con herramientas", "Herramientas inspeccionadas; "
             "separación mínima entre trabajadores."),
            ("Instalar pasarelas con barandas para cruzar zanjas.", "Caída a distinto nivel",
             "Pasarela de 0.60 m de ancho con barandas a 0.90 m."),
            ("Inspeccionar la zanja al inicio de cada jornada, después de lluvias o sismos.",
             "Fallas de talud no detectadas", "Check-list SST-CHK-03 firmado por persona "
             "competente."),
            ("Rellenar y retirar el entibado de abajo hacia arriba por capas compactadas.",
             "Derrumbe durante el desentibado", "Retiro progresivo; personal fuera de la zanja "
             "al retirar puntales superiores."),
            ("Orden y limpieza del área; cierre del PETAR.", "Tropiezos", "Retiro de "
             "excedentes, herramientas guardadas, firma de cierre."),
        ],
        "restricciones": ["Ningún trabajador ingresa a una zanja de más de 1.50 m sin entibado.",
                          "Prohibido acopiar material o estacionar maquinaria cerca del borde.",
                          "Suspender el trabajo ante lluvias intensas, filtraciones o grietas.",
                          "Nunca trabajar solo dentro de una zanja."],
    },
    {
        "codigo": "PETS-02", "titulo": "TRABAJOS EN ALTURA Y USO DE ANDAMIOS",
        "objetivo": "Establecer los controles para todo trabajo realizado a más de 1.80 m de "
                    "altura sobre el nivel del piso, incluido el armado, uso y desmontaje de "
                    "andamios multidireccionales.",
        "alcance": "Tarrajeo exterior, albañilería, encofrado de losas, montaje de estructuras, "
                   "coberturas, pintura e instalaciones en altura.",
        "responsables": [
            "Residente de obra: provee andamios normados y sistemas de detención de caídas.",
            "Ing. SSOMA: autoriza el PETAR de altura e inspecciona andamios y arneses.",
            "Capataz: verifica la tarjeta del andamio y el uso permanente del arnés.",
            "Andamiero competente: arma, modifica y desmonta el andamio según el fabricante.",
            "Trabajador: inspecciona su arnés (pre-uso) y permanece enganchado al 100 %.",
        ],
        "epp": ["Arnés de cuerpo completo ANSI Z359.11", "Línea de vida doble con absorbedor",
                "Casco con barbiquejo", "Lentes", "Guantes de badana", "Botín punta de acero"],
        "equipos": ["Andamio multidireccional certificado", "Línea de vida horizontal de cable "
                    "Ø1/2\" / vertical con freno", "Puntos de anclaje ≥ 2,268 kg (5,000 lb)",
                    "Tarjetas de inspección verde/roja", "Sogas para izaje de herramientas"],
        "pasos": [
            ("Obtener el PETAR de altura y elaborar el ATS.", "Caída a distinto nivel", "PETAR "
             "firmado, personal con EMO apto para altura."),
            ("Inspeccionar arnés y líneas (check-list SST-CHK-01).", "Falla del sistema de "
             "detención", "Retirar de servicio todo equipo defectuoso."),
            ("Nivelar el terreno y colocar placas base y husillos sobre tablones de reparto.",
             "Hundimiento / volcadura del andamio", "Terreno compactado, husillos ≤ 0.30 m de "
             "extensión."),
            ("Armar el andamio por cuerpos, instalando crucetas y diagonales en cada nivel.",
             "Colapso estructural", "Armado solo por andamiero competente según manual del "
             "fabricante."),
            ("Instalar plataforma completa, baranda superior a 0.90 m, intermedia a 0.45 m y "
             "rodapié de 0.15 m.", "Caída de personas y objetos", "Plataformas con seguros "
             "anti-levantamiento."),
            ("Arriostrar el andamio a la estructura cada 4 m de altura (relación altura/base "
             "≤ 3).", "Volcadura", "Amarres rígidos a elementos estructurales."),
            ("Colocar la tarjeta VERDE tras la inspección del Ing. SSOMA.", "Uso de andamio "
             "inseguro", "Tarjeta ROJA = prohibido su uso."),
            ("Acceder por la escalera interna, siempre con tres puntos de apoyo.", "Caída "
             "durante el ascenso", "Prohibido trepar por crucetas o marcos."),
            ("Engancharse a la línea de vida independiente antes de exponerse al borde.",
             "Caída libre", "Anclaje por encima del anillo dorsal; 100 % de enganche."),
            ("Izar herramientas y materiales con soga y balde, nunca lanzarlos.", "Caída de "
             "objetos", "Área inferior delimitada con malla y cinta."),
            ("Al término, retirar materiales de la plataforma y desmontar en orden inverso.",
             "Caída durante el desmontaje", "Desmontaje con arnés enganchado."),
        ],
        "restricciones": ["Prohibido el uso de andamios artesanales, cilindros o escaleras como "
                          "plataforma.", "Suspender trabajos con vientos fuertes o lluvia.",
                          "Prohibido trabajar en altura sin vigía y sin plan de rescate.",
                          "No exceder la carga máxima de la plataforma."],
    },
    {
        "codigo": "PETS-03", "titulo": "HABILITADO DE ACERO Y USO DE AMOLADORA / TRONZADORA",
        "objetivo": "Prevenir cortes, proyección de partículas, empalamientos y sobreesfuerzos "
                    "durante el habilitado, corte, doblado y colocación del acero de refuerzo.",
        "alcance": "Patio de habilitado de acero y frentes de armado de cimentaciones, columnas, "
                   "vigas, losas y muros.",
        "responsables": [
            "Residente de obra: asegura herramientas con guardas y capuchones de seguridad.",
            "Ing. SSOMA: verifica herramientas eléctricas y capacita en su uso.",
            "Capataz fierrero: dirige la tarea y controla la colocación de capuchones.",
            "Operario: usa la herramienta según el fabricante e inspecciona discos.",
        ],
        "epp": ["Casco", "Lentes + careta facial", "Guantes de cuero cromo", "Protector auditivo",
                "Botín punta de acero", "Hombreras de cuero (carga de varillas)"],
        "equipos": ["Tronzadora con guarda", "Amoladora con guarda y mango lateral",
                    "Dobladora / cizalla manual", "Banco de habilitado fijo",
                    "Capuchones plásticos (hongos) para mechas", "Extensiones con enchufe blindado"],
        "pasos": [
            ("Elaborar el ATS e inspeccionar herramientas (cinta de color del mes).",
             "Herramientas defectuosas", "Retirar equipos sin guarda o con cables dañados."),
            ("Verificar discos: diámetro, RPM ≥ RPM del equipo y fecha de vencimiento.",
             "Rotura del disco", "Disco adecuado, guarda instalada."),
            ("Trasladar varillas entre dos personas sobre el hombro protegido.",
             "Sobreesfuerzo", "Carga ≤ 25 kg por persona; R.M. 375-2008-TR."),
            ("Asegurar la varilla en la prensa del banco antes de cortar.", "Proyección / "
             "latigazo", "Nunca sujetar la pieza con el pie o la mano."),
            ("Cortar con la tronzadora / amoladora con ambas manos y posición estable.",
             "Contacto con disco, chispas", "Biombo de protección; área libre de combustibles."),
            ("Doblar con dobladora manual sobre banco fijo.", "Golpes / atrapamiento de manos",
             "Uso de tubo de palanca adecuado."),
            ("Colocar CAPUCHONES PLÁSTICOS / HONGOS en todas las mechas verticales expuestas.",
             "Empalamiento / empalomamiento", "Inspección diaria del 100 % de mechas."),
            ("Armar la estructura con apoyo de andamio o plataforma cuando supere 1.80 m.",
             "Caída a distinto nivel", "Aplicar PETS-02 y arnés."),
            ("Ordenar el acero habilitado sobre durmientes, clasificado por diámetro.",
             "Tropiezos y cortes", "Pasillos libres de 0.80 m."),
            ("Desconectar herramientas y limpiar el área al finalizar.", "Arranque "
             "involuntario", "Desenchufar al cambiar discos o al terminar."),
        ],
        "restricciones": ["Prohibido retirar la guarda de la amoladora o tronzadora.",
                          "Prohibido dejar mechas sin capuchones.",
                          "Prohibido usar discos de corte para desbaste o viceversa."],
    },
    {
        "codigo": "PETS-04", "titulo": "VACIADO DE CONCRETO Y VIBRADO",
        "objetivo": "Establecer la secuencia segura para el vaciado de concreto premezclado o "
                    "preparado en obra (mezcladora, bomba o winche) y su vibrado.",
        "alcance": "Vaciado de cimientos, zapatas, columnas, vigas, losas, muros, veredas y "
                   "sardineles.",
        "responsables": [
            "Residente / Ing. de producción: verifica encofrados y apuntalamiento antes del vaciado.",
            "Ing. SSOMA: verifica bordes protegidos, mangueras aseguradas y EPP.",
            "Capataz: coordina con el operador de bomba y el señalero.",
            "Operador de bomba / mezcladora: opera con check-list y respeta al señalero.",
            "Vibradorista: rota cada 2 h y usa EPP anti-vibración.",
        ],
        "epp": ["Casco", "Lentes antiempañantes", "Guantes de nitrilo / neopreno",
                "Botas de jebe con punta de acero", "Ropa impermeable", "Protector auditivo",
                "Respirador para polvo de cemento"],
        "equipos": ["Bomba de concreto / mezcladora / winche", "Vibrador de inmersión",
                    "Plataformas y buguis", "Mangueras con cable antilátigo",
                    "Iluminación (vaciados nocturnos)"],
        "pasos": [
            ("Verificar encofrados, puntales y acero (protocolo de pre-vaciado).", "Colapso del "
             "encofrado", "Liberación del encofrado por el Ing. de campo."),
            ("Elaborar el ATS y asignar funciones (señalero, vibradorista, regleador).",
             "Desorganización", "Charla específica del vaciado."),
            ("Ubicar la bomba en terreno firme con estabilizadores sobre tacos.", "Volcadura",
             "Distancia ≥ 3 m a redes eléctricas y bordes de zanja."),
            ("Asegurar las uniones de mangueras con abrazaderas y cable de seguridad.",
             "Latigazo por rotura", "Nadie se ubica frente a la salida de la manguera."),
            ("Delimitar el área de trabajo y proteger bordes de losa con barandas.", "Caída a "
             "distinto nivel", "Baranda a 0.90 m y rodapié; arnés en bordes."),
            ("Vaciar el concreto en capas, controlando la altura de caída (≤ 1.50 m).",
             "Sobrecarga del encofrado", "Vigía del encofrado durante el vaciado."),
            ("Vibrar sin apoyar el vibrador en el acero ni en el encofrado.", "Vibración "
             "mano-brazo, contacto eléctrico", "Vibrador con puesta a tierra y rotación."),
            ("Evitar el contacto de la piel con el concreto; lavar de inmediato salpicaduras.",
             "Dermatitis / quemadura química", "Agua limpia y lavaojos disponible."),
            ("Limpiar la bomba y mangueras con el sistema despresurizado.", "Proyección de "
             "concreto", "Prohibido abrir uniones con presión."),
            ("Orden y limpieza; curado del concreto.", "Resbalones", "Retiro de lechada y "
             "restos de concreto."),
        ],
        "restricciones": ["Prohibido desatorar mangueras con presión.",
                          "Prohibido vaciar sin la liberación del encofrado.",
                          "Suspender vaciados nocturnos sin iluminación ≥ 200 lux."],
    },
    {
        "codigo": "PETS-05", "titulo": "INSTALACIONES ELÉCTRICAS PROVISIONALES Y BLOQUEO LOTO",
        "objetivo": "Prevenir contactos eléctricos directos e indirectos mediante instalaciones "
                    "provisionales seguras y la aplicación del Bloqueo y Etiquetado (LOTO).",
        "alcance": "Tableros provisionales, extensiones, grupos electrógenos, instalaciones "
                   "eléctricas definitivas y mantenimiento de equipos energizados.",
        "responsables": [
            "Residente de obra: designa al electricista calificado responsable.",
            "Ing. SSOMA: verifica tableros, diferenciales, puesta a tierra y kits LOTO.",
            "Electricista calificado: ejecuta el bloqueo, verifica cero energía y libera.",
            "Trabajador: coloca su candado y tarjeta personal; nunca retira el de otro.",
        ],
        "epp": ["Casco dieléctrico Clase E", "Guantes dieléctricos Clase 0", "Botín dieléctrico",
                "Lentes", "Ropa ignífuga (arco eléctrico)"],
        "equipos": ["Tablero con diferencial 30 mA y termomagnéticos", "Pozo de puesta a tierra "
                    "≤ 25 Ω", "Revelador / detector de tensión", "Candados, pinzas multibloqueo "
                    "y tarjetas \"PELIGRO NO OPERAR\"", "Herramientas aisladas 1000 V",
                    "Pértiga y alfombra dieléctrica"],
        "pasos": [
            ("Obtener el PETAR eléctrico e identificar todas las fuentes de energía.",
             "Energía no identificada", "Diagrama unifilar actualizado."),
            ("Comunicar el bloqueo al personal afectado.", "Reenergización accidental",
             "Comunicación verbal y escrita."),
            ("Apagar el equipo y abrir el interruptor / seccionador.", "Arco eléctrico",
             "Maniobra con EPP dieléctrico y de pie a un costado."),
            ("Colocar candado personal y tarjeta de bloqueo en cada fuente.", "Operación por "
             "terceros", "Un candado por trabajador (pinza multibloqueo)."),
            ("Liberar energías residuales (condensadores, puesta a tierra temporal).",
             "Descarga residual", "Puesta a tierra y cortocircuito temporal."),
            ("Verificar CERO ENERGÍA con revelador probado antes y después.", "Contacto directo",
             "Prueba de arranque del equipo."),
            ("Ejecutar el trabajo con herramientas aisladas.", "Contacto eléctrico",
             "Distancias de seguridad; alfombra dieléctrica."),
            ("Al finalizar, retirar herramientas y verificar que el personal esté fuera.",
             "Energización con personal expuesto", "Conteo de personal."),
            ("Retirar candados (cada dueño el suyo) y reenergizar.", "Retiro indebido",
             "Retiro de candado ajeno solo con procedimiento formal."),
            ("Inspeccionar mensualmente tableros, diferenciales (botón de prueba) y puesta a "
             "tierra.", "Falla de protección", "Registro de inspección y medición de resistencia."),
        ],
        "restricciones": ["Prohibido trabajar en circuitos energizados.",
                          "Prohibido usar extensiones con empalmes, cables mellizos o enchufes "
                          "artesanales.", "Prohibido el ingreso de personal no autorizado a "
                          "tableros."],
    },
]

BOTIQUIN_G050 = [
    ("Guantes quirúrgicos descartables", "1 caja (50 pares)"),
    ("Yodopovidona solución antiséptica 120 ml", "1 frasco"),
    ("Agua oxigenada 120 ml", "1 frasco"),
    ("Alcohol medicinal 70° 250 ml", "1 frasco"),
    ("Clorhexidina / jabón antiséptico", "1 frasco"),
    ("Gasas esterilizadas 10 x 10 cm", "20 paquetes"),
    ("Apósitos estériles", "10 unidades"),
    ("Esparadrapo 2.5 cm x 5 m", "2 rollos"),
    ("Vendas elásticas de 2\", 4\" y 6\"", "2 de cada una"),
    ("Venda triangular", "2 unidades"),
    ("Algodón 100 g", "1 paquete"),
    ("Curitas / tiritas adhesivas", "1 caja"),
    ("Paletas bajalengua (para entablillado)", "10 unidades"),
    ("Férulas inmovilizadoras (brazo y pierna)", "1 juego"),
    ("Collarín cervical regulable", "1 unidad"),
    ("Suero fisiológico 1 L", "2 frascos"),
    ("Tijera punta roma y pinza", "1 de cada una"),
    ("Manta térmica de emergencia", "2 unidades"),
    ("Linterna con pilas", "1 unidad"),
    ("Mascarilla para RCP (barrera de protección)", "1 unidad"),
    ("Camilla rígida con correas e inmovilizador de cabeza", "1 unidad"),
]

DIRECTORIO_EMERGENCIA = [
    ("Central de Emergencias - Bomberos Voluntarios del Perú", "116"),
    ("Policía Nacional del Perú", "105"),
    ("SAMU - Atención Médica de Urgencia (MINSA)", "106"),
    ("EsSalud - Emergencias", "117"),
    ("Defensa Civil (INDECI)", "115"),
    ("Centro de salud / hospital más cercano", "Completar en obra"),
    ("Comisaría del sector", "Completar en obra"),
    ("Empresa eléctrica (emergencias)", "Completar en obra"),
    ("Empresa de agua y saneamiento (emergencias)", "Completar en obra"),
    ("Residente de obra", "Completar en obra"),
    ("Ingeniero SSOMA", "Completar en obra"),
    ("Aseguradora SCTR (central 24 h)", "Completar en obra"),
]

REQUISITOS_LEGALES = [
    ("Ley N° 29783 y modificatoria Ley N° 30222", "Ley de Seguridad y Salud en el Trabajo",
     "Principios, SG-SST, IPERC, comité, registros y derechos/obligaciones."),
    ("D.S. N° 005-2012-TR y modificatorias", "Reglamento de la Ley N° 29783",
     "Desarrollo del SG-SST, notificación de accidentes y registros."),
    ("D.S. N° 011-2019-TR", "Reglamento de SST para el Sector Construcción",
     "Plan de SST de obra, responsabilidades, trabajos de alto riesgo."),
    ("Norma Técnica G.050 (RNE)", "Seguridad durante la Construcción",
     "Requisitos técnicos de seguridad, PSST, instalaciones provisionales, EPP y EPC."),
    ("R.M. N° 050-2013-TR", "Formatos referenciales de registros obligatorios",
     "Metodología IPERC y formatos de registros del SG-SST."),
    ("R.M. N° 375-2008-TR", "Norma Básica de Ergonomía",
     "Manipulación manual de cargas (máx. 25 kg varones / 15 kg mujeres), posturas."),
    ("Ley N° 26790 y D.S. N° 003-98-SA", "Seguro Complementario de Trabajo de Riesgo",
     "SCTR Salud y Pensión obligatorio para actividades de construcción."),
    ("R.M. N° 312-2011/MINSA", "Protocolos de exámenes médico ocupacionales",
     "EMO de ingreso, periódicos y de retiro."),
    ("D.S. N° 015-2005-SA", "Valores límite permisibles para agentes químicos",
     "Control de exposición a polvo, humos y vapores."),
    ("R.M. N° 111-2013-MEM/DM", "Reglamento de SST con Electricidad (RESESATE)",
     "Trabajos eléctricos, bloqueo y distancias de seguridad."),
    ("NTP 399.010-1", "Señales de seguridad: colores, símbolos y formas",
     "Señalización temporal de obra y mapa de riesgos."),
    ("NTP 350.043-1", "Extintores portátiles: selección, distribución e inspección",
     "Dotación y mantenimiento de extintores."),
    ("NTP 900.058:2019", "Código de colores para el almacenamiento de residuos",
     "Segregación de residuos sólidos en obra."),
    ("D.L. N° 1278 y D.S. N° 014-2017-MINAM", "Ley de Gestión Integral de Residuos Sólidos",
     "Manejo de residuos peligrosos y no peligrosos."),
    ("D.S. N° 002-2022-VIVIENDA", "Reglamento de gestión de residuos de construcción y "
     "demolición", "Manejo y disposición de RCD."),
    ("ANSI Z89.1 / Z87.1 / Z359", "Normas técnicas de EPP", "Cascos, lentes y sistemas de "
     "detención de caídas."),
]


# =============================================================================
# 9. GENERACIÓN DEL PSST EN WORD (python-docx)
# =============================================================================
W_AZUL = RGBColor(0x1F, 0x38, 0x64)
W_AZUL2 = RGBColor(0x2E, 0x75, 0xB6)


def _sombrear(celda, color_hex):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def _campo(parrafo, instruccion, texto="Clic derecho > Actualizar campo"):
    run = parrafo.add_run()
    for tipo, contenido in (("begin", None), ("instr", instruccion), ("separate", None),
                            ("text", texto), ("end", None)):
        if tipo == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = contenido
        elif tipo == "text":
            el = OxmlElement("w:t")
            el.text = contenido
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tipo)
        run._r.append(el)
    return run


class DocPSST:
    """Ayudante para redactar el documento Word con estilo uniforme."""

    def __init__(self):
        self.doc = Document()
        st = self.doc.styles["Normal"]
        st.font.name = "Arial"
        st.font.size = Pt(10.5)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        st.paragraph_format.space_after = Pt(4)
        st.paragraph_format.line_spacing = 1.15
        for nombre, tam, color in (("Heading 1", 14, W_AZUL), ("Heading 2", 12, W_AZUL2),
                                   ("Heading 3", 11, W_AZUL)):
            s = self.doc.styles[nombre]
            s.font.name = "Arial"
            s.font.size = Pt(tam)
            s.font.bold = True
            s.font.color.rgb = color
            s.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            s.paragraph_format.space_before = Pt(12 if nombre == "Heading 1" else 8)
            s.paragraph_format.space_after = Pt(6)
        sec = self.doc.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
        sec.left_margin = sec.right_margin = Cm(2.2)
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(1.8)

    def h1(self, t):
        return self.doc.add_heading(t, level=1)

    def h2(self, t):
        return self.doc.add_heading(t, level=2)

    def h3(self, t):
        return self.doc.add_heading(t, level=3)

    def p(self, t, negrita=False, centrado=False, tam=None, cursiva=False, color=None):
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER if centrado else WD_ALIGN_PARAGRAPH.JUSTIFY
        run = par.add_run(t)
        run.bold = negrita
        run.italic = cursiva
        if tam:
            run.font.size = Pt(tam)
        if color:
            run.font.color.rgb = color
        return par

    def pk(self, clave, valor):
        """Párrafo con etiqueta en negrita."""
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        par.add_run(clave).bold = True
        par.add_run(valor)
        return par

    def vinetas(self, items, estilo="List Bullet"):
        for t in items:
            par = self.doc.add_paragraph(style=estilo)
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if isinstance(t, tuple):
                par.add_run(t[0]).bold = True
                par.add_run(t[1])
            else:
                par.add_run(t)

    def numeros(self, items):
        self.vinetas(items, "List Number")

    def tabla(self, encabezado, filas, anchos=None, tam=8.5, color_cab="1F3864",
              resaltar=None):
        t = self.doc.add_table(rows=1, cols=len(encabezado))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, txt in enumerate(encabezado):
            c = t.rows[0].cells[j]
            c.text = ""
            run = c.paragraphs[0].add_run(str(txt))
            run.bold = True
            run.font.size = Pt(tam)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            _sombrear(c, color_cab)
        # repetir encabezado en cada página
        trPr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trPr.append(th)
        for i, fila in enumerate(filas):
            cells = t.add_row().cells
            for j, v in enumerate(fila):
                cells[j].text = ""
                run = cells[j].paragraphs[0].add_run("" if v is None else str(v))
                run.font.size = Pt(tam)
                if resaltar and resaltar(i, j, v):
                    run.bold = True
                    _sombrear(cells[j], resaltar(i, j, v))
        if anchos:
            t.autofit = False
            for j, w in enumerate(anchos):
                t.columns[j].width = Cm(w)
            for fila in t.rows:
                for j, w in enumerate(anchos):
                    fila.cells[j].width = Cm(w)
        self.doc.add_paragraph()
        return t

    def salto(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def guardar(self, ruta):
        self.doc.save(ruta)


def _color_nivel_hex(nivel):
    return {"TRIVIAL": "C6EFCE", "TOLERABLE": "92D050", "MODERADO": "FFFF00",
            "IMPORTANTE": "FF9900", "INTOLERABLE": "FF4D4D"}.get(nivel)


def exportar_word(ruta, cfg, filas_iperc, desag, crono):
    cfg = completar_config(cfg)
    D = DocPSST()
    doc = D.doc
    N, M = cfg["n_total"], cfg["meses"]
    rg = cfg["riesgos"]
    total = desag["total"]
    comite = N >= 20
    alto_riesgo = [t for k, t in (("altura", "Trabajos en altura > 1.80 m"),
                                  ("excavacion", "Excavaciones > 1.50 m"),
                                  ("caliente", "Trabajos en caliente"),
                                  ("confinado", "Espacios confinados"),
                                  ("izaje", "Izaje de cargas")) if rg.get(k)]

    # ---------- encabezado y pie ----------
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    hp = sec.header.paragraphs[0]
    hp.text = "PLAN DE SEGURIDAD Y SALUD EN EL TRABAJO (PSST) - NORMA G.050  |  " + \
              cfg["proyecto"][:70]
    hp.runs[0].font.size = Pt(7.5)
    hp.runs[0].font.color.rgb = W_AZUL2
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = fp.add_run("SST-PSST-01  |  Versión 01  |  Página ")
    r0.font.size = Pt(8)
    _campo(fp, "PAGE", "1").font.size = Pt(8)
    fp.add_run(" de ").font.size = Pt(8)
    _campo(fp, "NUMPAGES", "1").font.size = Pt(8)

    # ---------- portada ----------
    for _ in range(3):
        doc.add_paragraph()
    D.p(cfg["entidad"].upper(), True, True, 13, color=W_AZUL)
    doc.add_paragraph()
    D.p("PLAN DE SEGURIDAD Y SALUD EN EL TRABAJO", True, True, 24, color=W_AZUL)
    D.p("(PSST - NORMA TÉCNICA G.050 / D.S. N° 011-2019-TR)", True, True, 13, color=W_AZUL2)
    doc.add_paragraph()
    D.p("OBRA:", True, True, 11)
    D.p(cfg["proyecto"], True, True, 14)
    doc.add_paragraph()
    D.p("UBICACIÓN: " + cfg["ubicacion"], False, True, 11)
    D.p("MODALIDAD DE EJECUCIÓN: " + cfg["modalidad"].upper(), False, True, 11)
    D.p("PLAZO DE EJECUCIÓN: {} DÍAS CALENDARIO ({} MESES)".format(cfg["dias"], M), False, True,
        11)
    for _ in range(4):
        doc.add_paragraph()
    D.tabla(["Versión", "Fecha", "Elaborado por", "Revisado por", "Aprobado por"],
            [["01", cfg["fecha"], "{}\nCIP N° {}".format(cfg["ssoma"], cfg["cip"]),
              cfg["residente"], "Supervisión / Entidad"]], [2, 2.5, 4.5, 4, 4], 9)
    D.salto()

    # ---------- índice ----------
    D.p("ÍNDICE", True, True, 14, color=W_AZUL)
    _campo(doc.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u',
           "Presione F9 o clic derecho > Actualizar campo para generar el índice.")
    D.salto()

    # ================= 1 =================
    D.h1("1. OBJETIVO, ALCANCE Y DESCRIPCIÓN DEL PROYECTO")
    D.h2("1.1 Objetivo general")
    D.p("Establecer las medidas preventivas, los recursos, las responsabilidades y los "
        "procedimientos necesarios para garantizar la seguridad y salud de todos los "
        "trabajadores, visitantes y terceros durante la ejecución de la obra \"{}\", en "
        "cumplimiento de la Ley N° 29783, su Reglamento D.S. N° 005-2012-TR, el D.S. "
        "N° 011-2019-TR y la Norma Técnica G.050 \"Seguridad durante la Construcción\" del "
        "Reglamento Nacional de Edificaciones.".format(cfg["proyecto"]))
    D.h2("1.2 Objetivos específicos")
    D.vinetas([
        "Identificar los peligros, evaluar los riesgos y establecer medidas de control para "
        "cada partida del presupuesto de obra mediante la matriz IPERC.",
        "Alcanzar la meta de CERO accidentes incapacitantes y CERO enfermedades ocupacionales.",
        "Cumplir el 100 % del programa de capacitación, inspecciones y simulacros.",
        "Dotar oportunamente de equipos de protección individual (EPI) y colectiva (EPC) "
        "certificados a todo el personal.",
        "Responder de manera eficaz ante emergencias (sismos, incendios, derrumbes, "
        "electrocución) minimizando sus consecuencias.",
        "Controlar los aspectos ambientales significativos: polvo, ruido y residuos sólidos.",
    ])
    D.h2("1.3 Alcance")
    D.p("El presente plan es de cumplimiento obligatorio para el contratista principal, "
        "subcontratistas, proveedores, personal de supervisión y visitantes, durante todas las "
        "etapas de la obra: movilización, obras provisionales, ejecución de partidas, pruebas, "
        "desmovilización y entrega. Comprende todas las actividades del presupuesto de obra "
        "analizadas en la matriz IPERC y todos los frentes de trabajo y áreas auxiliares "
        "(almacenes, comedor, patio de maniobras y zonas de acopio).")
    D.h2("1.4 Descripción del proyecto")
    D.tabla(["Dato", "Descripción"], [
        ["Nombre de la obra", cfg["proyecto"]],
        ["Entidad / Cliente", cfg["entidad"]],
        ["Ubicación", cfg["ubicacion"]],
        ["Modalidad de ejecución", cfg["modalidad"]],
        ["Plazo de ejecución", "{} días calendario ({} meses)".format(cfg["dias"], M)],
        ["N° promedio de obreros", str(cfg["n_obr"])],
        ["N° de staff técnico", str(cfg["n_ing"])],
        ["Ingeniero residente", cfg["residente"]],
        ["Especialista SSOMA", "{} - CIP N° {}".format(cfg["ssoma"], cfg["cip"])],
        ["Trabajos de alto riesgo", ", ".join(alto_riesgo) if alto_riesgo else "No se prevén"],
        ["Costo directo de seguridad (G.050)", soles(total)],
    ], [5, 11.5], 9)
    D.h2("1.5 Partidas del presupuesto comprendidas en el plan")
    D.p("Las siguientes partidas fueron leídas del presupuesto de obra y emparejadas "
        "automáticamente con su familia de riesgo para la elaboración de la matriz IPERC:")
    D.tabla(["Ítem", "Partida", "Und.", "Metrado", "Familia de riesgo"],
            [[f["item"], f["partida"], f["und"],
              "" if f["metrado"] is None else "{:,.2f}".format(f["metrado"]),
              f["familia"]] for f in filas_iperc], [1.5, 6.5, 1.2, 2, 5.3], 8)

    # ================= 2 =================
    D.h1("2. POLÍTICA DEL SISTEMA DE GESTIÓN DE SEGURIDAD Y SALUD EN EL TRABAJO")
    D.p("La empresa ejecutora de la obra \"{}\" reconoce que su principal activo son sus "
        "trabajadores y que la seguridad y la salud en el trabajo forman parte integral de la "
        "gestión del proyecto. En ese sentido, en concordancia con el artículo 22 de la Ley "
        "N° 29783, asume los siguientes compromisos:".format(cfg["proyecto"]))
    D.numeros([
        "Proteger la seguridad y salud de todos los miembros de la organización mediante la "
        "prevención de lesiones, dolencias, enfermedades e incidentes relacionados con el "
        "trabajo.",
        "Cumplir la normativa legal vigente en materia de seguridad y salud en el trabajo, la "
        "Norma G.050, el D.S. N° 011-2019-TR y los compromisos voluntarios suscritos.",
        "Garantizar que los trabajadores y sus representantes sean consultados y participen "
        "activamente en todos los elementos del Sistema de Gestión de SST.",
        "Identificar los peligros, evaluar y controlar los riesgos aplicando la jerarquía de "
        "controles: eliminación, sustitución, controles de ingeniería, controles "
        "administrativos y equipos de protección personal.",
        "Proveer los recursos necesarios (EPI, EPC, capacitación, exámenes médicos y seguros) "
        "para la implementación del presente plan.",
        "Mejorar continuamente el desempeño del Sistema de Gestión de SST e integrarlo con "
        "la gestión ambiental de la obra.",
        "Respetar el derecho de todo trabajador a paralizar una tarea cuando identifique un "
        "peligro inminente que ponga en riesgo su vida o la de sus compañeros.",
    ])
    D.p("La presente política es difundida a todo el personal, exhibida en lugares visibles de "
        "la obra y revisada periódicamente para asegurar su vigencia.")
    doc.add_paragraph()
    D.p("_________________________________\n{}\nResidente de Obra".format(cfg["residente"]),
        False, True)

    # ================= 3 =================
    D.h1("3. ORGANIZACIÓN Y RESPONSABILIDADES")
    D.h2("3.1 Organigrama de seguridad y salud en obra")
    D.tabla(["Nivel", "Cargo", "Relación funcional"], [
        ["1", "Representante legal de la empresa / Gerente de proyecto", "Aprueba la política "
         "y asigna recursos"],
        ["2", "Ingeniero Residente de Obra", "Máxima autoridad en obra, responsable del PSST"],
        ["3", "Ingeniero Especialista SSOMA (Prevencionista CIP)", "Asesoría técnica, "
         "implementación y control del PSST"],
        ["3", "Comité de SST" if comite else "Supervisor de SST", "Órgano paritario de "
         "consulta y vigilancia" if comite else "Vigilancia y representación de los "
         "trabajadores"],
        ["4", "Ingenieros de campo / producción", "Aplicación del PSST en sus frentes"],
        ["5", "Maestro de obra / Capataces", "Supervisión directa de cuadrillas"],
        ["6", "Trabajadores (operarios, oficiales, peones) y subcontratistas", "Cumplimiento "
         "de estándares y procedimientos"],
        ["-", "Brigadas de emergencia", "Primeros auxilios, contra incendio y evacuación"],
    ], [1.3, 7.5, 7.7], 9)
    D.h2("3.2 Responsabilidades")
    D.h3("3.2.1 Ingeniero Residente de Obra")
    D.vinetas([
        "Liderar la implementación del PSST y asegurar los recursos económicos y logísticos.",
        "Aprobar la matriz IPERC, los PETS y los PETAR de trabajos de alto riesgo.",
        "Paralizar cualquier trabajo que represente riesgo inminente para los trabajadores.",
        "Participar en las investigaciones de accidentes e implementar las acciones "
        "correctivas.",
        "Asegurar que todo trabajador cuente con EMO, SCTR, inducción y EPP antes de iniciar "
        "labores.",
    ])
    D.h3("3.2.2 Ingeniero Especialista SSOMA / Prevencionista de Riesgos")
    D.vinetas([
        "Elaborar, actualizar y difundir el PSST, la matriz IPERC y los procedimientos.",
        "Ejecutar el programa de capacitación, inspecciones, observaciones y simulacros.",
        "Verificar la correcta emisión y cierre de los PETAR y ATS diarios.",
        "Controlar la entrega, uso y reposición de EPP (Registro de equipos de seguridad).",
        "Elaborar las estadísticas mensuales (IF, IG, IA) y reportarlas a la entidad.",
        "Investigar incidentes y accidentes, y notificar a la autoridad en los plazos de ley.",
        "Asesorar al Comité / Supervisor de SST y gestionar las no conformidades.",
    ])
    D.h3("3.2.3 {}".format("Comité de Seguridad y Salud en el Trabajo" if comite else
                           "Supervisor de Seguridad y Salud en el Trabajo"))
    if comite:
        D.p("Al contar la obra con {} trabajadores (20 o más), se conforma un Comité de SST "
            "paritario, integrado por representantes del empleador y de los trabajadores (estos "
            "últimos elegidos por votación secreta), conforme a los artículos 29 al 31 de la "
            "Ley N° 29783 y el D.S. N° 005-2012-TR.".format(N))
        D.vinetas([
            "Aprobar el Programa Anual de SST y el Reglamento Interno de SST.",
            "Reunirse mensualmente en forma ordinaria y en forma extraordinaria ante accidentes "
            "graves; registrar los acuerdos en el Libro de Actas.",
            "Realizar inspecciones periódicas a las áreas de trabajo.",
            "Vigilar el cumplimiento de la legislación, los reglamentos y el PSST.",
            "Analizar las causas y estadísticas de los incidentes y accidentes.",
        ])
    else:
        D.p("Al contar la obra con {} trabajadores (menos de 20), los trabajadores eligen un "
            "Supervisor de SST conforme al artículo 30 de la Ley N° 29783, quien ejerce las "
            "funciones de vigilancia del Comité.".format(N))
        D.vinetas([
            "Vigilar el cumplimiento del PSST y de la normativa de SST.",
            "Participar en las inspecciones e investigaciones de accidentes.",
            "Canalizar las sugerencias y reportes de actos y condiciones inseguras.",
        ])
    D.h3("3.2.4 Maestro de obra y capataces")
    D.vinetas([
        "Elaborar el ATS diario con su cuadrilla y dictar la charla de 5 minutos.",
        "Verificar el uso correcto del EPP y el estado de las herramientas.",
        "Mantener el orden y la limpieza de su frente de trabajo.",
        "Reportar inmediatamente todo incidente, acto o condición insegura.",
        "Detener el trabajo ante un peligro inminente y comunicarlo al Ing. SSOMA.",
    ])
    D.h3("3.2.5 Trabajadores")
    D.vinetas([
        "Cumplir los estándares, PETS, PETAR y las indicaciones del presente plan.",
        "Usar adecuadamente el EPP y cuidar los equipos de protección colectiva.",
        "Participar en las capacitaciones, simulacros y elaboración del ATS.",
        "No operar equipos para los que no estén autorizados ni capacitados.",
        "Someterse a los exámenes médicos ocupacionales obligatorios.",
        "Ejercer su derecho a negarse a realizar un trabajo inseguro (art. 63 Ley N° 29783).",
    ])
    D.h3("3.2.6 Subcontratistas y proveedores")
    D.p("Todo subcontratista debe cumplir el presente PSST, presentar su propia matriz IPERC y "
        "PETS de sus actividades, acreditar SCTR, EMO y capacitación de su personal, y asignar "
        "un responsable de seguridad cuando su personal supere los 10 trabajadores.")

    # ================= 4 =================
    D.h1("4. ELEMENTOS DEL PLAN")
    D.h2("4.1 Identificación de requisitos legales y contractuales")
    D.tabla(["Norma", "Título", "Aplicación en la obra"],
            [list(x) for x in REQUISITOS_LEGALES], [4, 5.5, 7], 8)
    D.h2("4.2 Objetivos y metas de SST")
    D.tabla(["Objetivo", "Meta", "Indicador", "Responsable"], [
        ["Prevenir accidentes", "0 accidentes incapacitantes", "Índice de frecuencia = 0",
         "Residente / SSOMA"],
        ["Capacitar al personal", "100 % del programa", "Capacitaciones ejecutadas / "
         "programadas", "Ing. SSOMA"],
        ["Inspeccionar condiciones", "≥ 4 inspecciones por mes", "Inspecciones ejecutadas / "
         "programadas", "Ing. SSOMA"],
        ["Dotar de EPP", "100 % del personal", "Trabajadores con EPP / total", "Almacén / "
         "SSOMA"],
        ["Controlar trabajos de alto riesgo", "100 % con PETAR", "PETAR emitidos / trabajos "
         "de alto riesgo", "Ing. SSOMA"],
        ["Preparar ante emergencias", "1 simulacro trimestral", "Simulacros ejecutados",
         "Jefe de brigadas"],
        ["Salud ocupacional", "100 % con EMO", "Trabajadores con EMO / total", "Residente"],
    ], [4, 3.5, 5, 4], 8.5)
    D.h2("4.3 Estándares de seguridad en obra")
    D.h3("4.3.1 Equipos de protección individual (EPI)")
    D.p("El EPP básico obligatorio para todo el personal es: casco con barbiquejo (ANSI Z89.1), "
        "lentes de seguridad (ANSI Z87.1), guantes según la tarea, botín de seguridad con punta "
        "de acero (NTP ISO 20345) y uniforme con cinta reflectiva. El EPP específico (arnés, "
        "respirador, protector auditivo, careta de soldar, botas de jebe, guantes dieléctricos) "
        "se asigna según la matriz IPERC. Su entrega se registra en el Registro de equipos de "
        "seguridad o emergencia y su reposición se realiza por desgaste o deterioro.")
    D.h3("4.3.2 Equipos de protección colectiva (EPC)")
    D.p("Se instalarán barandas rígidas en bordes de losa y aberturas, líneas de vida "
        "horizontales de cable de acero Ø1/2\" con grapas Crosby, malla naranja de seguridad "
        "H = 1.00 m con cachacos en bordes de excavación, capuchones plásticos en mechas de "
        "acero, entibados en zanjas de más de 1.50 m y tableros eléctricos con interruptor "
        "diferencial y puesta a tierra.")
    D.h3("4.3.3 Orden y limpieza")
    D.p("Las áreas de trabajo, pasillos y accesos se mantendrán libres de materiales, "
        "desperdicios y clavos. Los materiales se apilarán de forma estable a no más de 1.80 m "
        "de altura. Al final de cada jornada se efectuará la limpieza del frente de trabajo y "
        "los residuos serán segregados según el código de colores NTP 900.058.")
    D.h3("4.3.4 Herramientas y equipos")
    D.p("Las herramientas manuales y eléctricas se inspeccionarán mensualmente y se marcarán "
        "con la cinta de color del mes. Las herramientas eléctricas deben contar con doble "
        "aislamiento o puesta a tierra, guardas de protección y enchufes industriales. Los "
        "equipos pesados contarán con check-list diario de pre-uso, alarma de retroceso y "
        "operador acreditado.")
    D.h3("4.3.5 Instalaciones provisionales y servicios higiénicos")
    nsh = sshh_requeridos(N)
    D.p("Se implementarán caseta de guardianía, oficinas técnicas, almacén general, almacén de "
        "productos químicos con bandeja antiderrame y hojas MSDS, comedor, vestuarios y "
        "servicios higiénicos (baños químicos) conforme a la Norma G.050. Para {} trabajadores "
        "se requiere como mínimo: {} inodoros, {} lavatorios, {} duchas y {} urinarios."
        .format(N, *nsh))
    D.tabla(["N° de trabajadores", "Inodoros", "Lavatorios", "Duchas", "Urinarios"],
            [list(x) for x in TABLA_SSHH], [4, 3, 3, 3, 3], 9)
    D.h3("4.3.6 Señalización temporal")
    D.p("La señalización se realizará conforme a la NTP 399.010-1: señales de obligación "
        "(círculo azul), prohibición (círculo rojo con banda diagonal), advertencia (triángulo "
        "amarillo con borde negro), evacuación y emergencia (rectángulo verde) y equipos contra "
        "incendio (cuadrado rojo). El detalle se presenta en la lámina SS-01.")
    D.h3("4.3.7 Almacenamiento de materiales peligrosos")
    D.p("Los combustibles, solventes, pinturas, aditivos y gases comprimidos se almacenarán en "
        "un ambiente ventilado, señalizado, con bandeja antiderrame, kit antiderrame, extintor "
        "PQS y hojas de datos de seguridad (MSDS/SDS) en idioma español. Los cilindros de gas "
        "se mantendrán en posición vertical, encadenados y con capuchón.")

    # ================= 5 =================
    D.h1("5. METODOLOGÍA DE LA MATRIZ IPERC Y MAPA DE RIESGOS")
    D.h2("5.1 Metodología de evaluación (R.M. N° 050-2013-TR)")
    D.p("La identificación de peligros y evaluación de riesgos se realizó para cada partida del "
        "presupuesto de obra. La probabilidad se determina como la suma de cuatro índices: "
        "IP = A + B + C + D, y el nivel de riesgo como el producto de la probabilidad por la "
        "severidad: Riesgo = IP x S.")
    D.tabla(["Índice", "Criterio", "Valores"],
            [[k, CRITERIOS_IPERC[k][0], "\n".join(CRITERIOS_IPERC[k][1])]
             for k in ("A", "B", "C", "D", "S")], [1.5, 4.5, 10.5], 8.5)
    D.tabla(["Nivel de riesgo", "Puntaje", "Interpretación / acción requerida"],
            [[n, r_, a] for n, r_, a, _ in NIVELES_RIESGO], [3, 2, 11.5], 8.5,
            resaltar=lambda i, j, v: (NIVELES_RIESGO[i][3] if j == 0 else None))
    D.h2("5.2 Jerarquía de controles")
    D.p("Conforme al artículo 21 de la Ley N° 29783, las medidas de control se aplican en el "
        "siguiente orden de prioridad:")
    D.numeros([("Eliminación: ", "suprimir el peligro desde el diseño o la planificación."),
               ("Sustitución: ", "reemplazar por un proceso, material o equipo menos peligroso."),
               ("Controles de ingeniería: ", "aislar el peligro con barreras, guardas, "
                "entibados, barandas, ventilación."),
               ("Controles administrativos: ", "procedimientos, PETS, PETAR, ATS, capacitación, "
                "señalización, rotación de personal."),
               ("Equipos de protección personal: ", "última barrera entre el peligro y el "
                "trabajador.")])
    D.h2("5.3 Resumen de la matriz IPERC por partida")
    D.p("La matriz completa con fórmulas activas se entrega en el libro Excel "
        "(hoja 1_Matriz_IPERC_Obra). A continuación se presenta el resumen:")
    D.tabla(["Ítem", "Partida", "Peligro principal", "IP x S", "Nivel inicial",
             "Control clave (ingeniería)", "Residual"],
            [[f["item"], f["partida"], f["peligro"], "{}x{}={}".format(f["IP"], f["S"], f["PxS"]),
              f["nivel"], f["ingenieria"], "{} ({})".format(f["nivel_res"], f["rPxS"])]
             for f in filas_iperc], [1.2, 3.2, 3.6, 1.5, 1.9, 3.6, 1.8], 7,
            resaltar=lambda i, j, v: (_color_nivel_hex(v) if j == 4 else
                                      (_color_nivel_hex(str(v).split(" ")[0]) if j == 6
                                       else None)))
    n_alto = sum(1 for f in filas_iperc if f["PxS"] > 16)
    D.p("De las {} partidas evaluadas, {} presentan un riesgo inicial IMPORTANTE o INTOLERABLE, "
        "las cuales requieren obligatoriamente la implementación de las medidas de control "
        "antes de su ejecución. Luego de aplicar los controles, el 100 % de las partidas "
        "alcanza un riesgo residual TOLERABLE o MODERADO.".format(len(filas_iperc), n_alto))
    D.h2("5.4 Mapa de riesgos de la obra")
    D.p("El mapa de riesgos se elabora sobre el plano de distribución de instalaciones "
        "provisionales (lámina SS-01) y se exhibe en la caseta de guardianía, comedor y "
        "oficina técnica. Representa mediante la simbología de la NTP 399.010-1 los peligros "
        "presentes en cada zona y los equipos de emergencia:")
    D.tabla(["Zona de la obra", "Peligros representados", "Señalización / equipo"], [
        ["Ingreso y patio de maniobras", "Tránsito de vehículos y maquinaria", "Advertencia "
         "de maquinaria, velocidad máx. 10 km/h, conos"],
        ["Frente de excavación", "Derrumbe, caída a desnivel", "Excavación profunda, cachacos "
         "con malla, uso obligatorio de casco"],
        ["Bloque en construcción", "Caída de objetos y personas, altura", "Uso obligatorio de "
         "arnés, caída de objetos, barandas"],
        ["Patio de habilitado de acero", "Cortes, proyección de partículas, ruido", "Uso "
         "obligatorio de lentes y protector auditivo"],
        ["Tablero eléctrico provisional", "Riesgo eléctrico", "Riesgo eléctrico, extintor CO2/"
         "PQS"],
        ["Almacén de productos químicos", "Incendio, intoxicación, derrames", "Prohibido fumar, "
         "MSDS, extintor PQS, kit antiderrame"],
        ["Comedor / oficinas", "Incendio, sismo", "Ruta de evacuación, zona segura, extintor, "
         "botiquín"],
        ["Punto de reunión", "-", "Punto de reunión en caso de sismo"],
    ], [4.5, 5.5, 6.5], 8.5)
    D.h2("5.5 IPERC continuo y Análisis de Trabajo Seguro (ATS)")
    D.p("Antes de iniciar cada tarea, la cuadrilla elabora el ATS (formato SST-ATS-01) "
        "identificando los pasos, peligros y controles específicos del día. Ante cambios en las "
        "condiciones de trabajo (clima, nuevos equipos, interferencias) el IPERC se actualiza. "
        "Para los trabajos de alto riesgo se emite adicionalmente el PETAR (formato "
        "SST-PETAR-01), válido por un turno.")

    D.h2("5.6 Fichas IPERC detalladas por partida")
    D.p("Para cada partida del presupuesto se detallan el peligro, el riesgo, la valoración y "
        "las medidas de control en el orden de la jerarquía de controles. Estas fichas se "
        "difunden a las cuadrillas antes del inicio de cada partida.")
    for f in filas_iperc:
        D.h3("{} - {}".format(f["item"], f["partida"]))
        D.tabla(["Elemento", "Descripción"], [
            ["Familia de riesgo", "{} ({})".format(f["familia"], f["tipo"])],
            ["Peligro", f["peligro"]], ["Riesgo", f["riesgo"]],
            ["Consecuencia", f["consecuencia"]],
            ["Evaluación inicial", "A={A}  B={B}  C={C}  D={D}  ->  IP={IP} ;  S={S} ;  "
             "P x S = {PxS}  ->  {nivel}".format(**f)],
            ["1. Eliminación", f["eliminacion"]], ["2. Sustitución", f["sustitucion"]],
            ["3. Controles de ingeniería", f["ingenieria"]],
            ["4. Controles administrativos", f["administrativo"]],
            ["5. EPP específico", f["epp"]],
            ["Evaluación residual", "A={rA}  B={rB}  C={rC}  D={rD}  ->  IP={rIP} ;  S={rS} ;  "
             "P x S = {rPxS}  ->  {nivel_res}".format(**f)],
            ["Permiso / procedimiento", "{}  |  {}".format(f["petar"], f["pets"])],
        ], [4.2, 12.3], 8, color_cab="2E75B6",
            resaltar=lambda i, j, v, f=f: (_color_nivel_hex(f["nivel"]) if (i == 4 and j == 1)
                                           else _color_nivel_hex(f["nivel_res"])
                                           if (i == 10 and j == 1) else None))

    # ================= 6 =================
    D.h1("6. PROCEDIMIENTOS ESCRITOS DE TRABAJO SEGURO (PETS)")
    D.p("Los PETS describen paso a paso cómo realizar de forma segura las tareas críticas "
        "identificadas en la matriz IPERC. Son difundidos a las cuadrillas antes del inicio de "
        "la actividad y su cumplimiento es verificado por el capataz y el Ing. SSOMA.")
    for k, pets in enumerate(PETS_LISTA, 1):
        D.h2("6.{} {} {}".format(k, pets["codigo"], pets["titulo"]))
        D.pk("Objetivo: ", pets["objetivo"])
        D.pk("Alcance: ", pets["alcance"])
        D.pk("Responsables:", "")
        D.vinetas(pets["responsables"])
        D.pk("EPP requerido: ", ", ".join(pets["epp"]) + ".")
        D.pk("Equipos, herramientas y materiales: ", "; ".join(pets["equipos"]) + ".")
        D.pk("Procedimiento:", "")
        D.tabla(["N°", "Paso / Secuencia de la tarea", "Peligro / Riesgo", "Medida de control"],
                [[str(i), a, b, c] for i, (a, b, c) in enumerate(pets["pasos"], 1)],
                [1, 7, 3.8, 4.7], 8)
        D.pk("Restricciones (PARE el trabajo si):", "")
        D.vinetas(pets["restricciones"])
        D.pk("Registros: ", "ATS, PETAR (cuando aplique), check-list de equipos, registro de "
             "capacitación del PETS.")

    # ================= 7 =================
    D.h1("7. PROGRAMA DE CAPACITACIÓN, SENSIBILIZACIÓN E INDUCCIÓN")
    D.h2("7.1 Inducción de hombre nuevo")
    D.p("Todo trabajador que ingrese a la obra recibirá una inducción de 4 horas como mínimo, "
        "antes de iniciar sus labores, con los siguientes contenidos:")
    D.vinetas(["Política de SST y Reglamento Interno de SST.", "Peligros y riesgos de la obra "
               "(matriz IPERC) y medidas de control.", "Uso, cuidado y mantenimiento del EPP.",
               "Procedimientos de trabajo seguro y trabajos de alto riesgo (PETAR).",
               "Plan de respuesta ante emergencias, rutas de evacuación y punto de reunión.",
               "Reporte de incidentes, actos y condiciones inseguras.", "Derechos y "
               "obligaciones de los trabajadores (Ley N° 29783)."])
    D.h2("7.2 Charlas diarias de 5 minutos")
    D.p("Se dictarán al inicio de cada jornada por el capataz o el Ing. SSOMA, sobre los "
        "riesgos de las tareas del día, lecciones aprendidas e incidentes ocurridos. Se "
        "registrarán con la firma de los asistentes.")
    D.h2("7.3 Capacitaciones específicas y cronograma")
    D.p("Las capacitaciones específicas se programan en función de las partidas del presupuesto "
        "y de su secuencia constructiva. Se presenta el cronograma para los {} meses de obra "
        "(el detalle mensual con marcas se encuentra en el libro Excel, hoja "
        "3_Programa_Capacitaciones):".format(M))
    D.tabla(["N°", "Actividad", "Tema", "Responsable", "Frecuencia", "Meses programados"],
            [[str(i), f["act"], f["tema"], f["resp"], f["frec"],
              ", ".join("M{}".format(m) for m in f["meses"])]
             for i, f in enumerate(crono, 1)], [0.8, 3.5, 5, 2.7, 1.8, 2.7], 7.5)
    D.h2("7.4 Sensibilización")
    D.p("Se desarrollarán campañas mensuales de sensibilización (afiches, periódico mural, "
        "reconocimiento al trabajador seguro del mes) orientadas a fortalecer la cultura de "
        "prevención y el autocuidado.")

    # ================= 8 =================
    D.h1("8. GESTIÓN DE NO CONFORMIDADES, INVESTIGACIÓN DE ACCIDENTES Y ESTADÍSTICAS")
    D.h2("8.1 Inspecciones y observaciones")
    D.p("Se realizarán inspecciones planeadas (EPP, andamios, excavaciones, herramientas, "
        "instalaciones eléctricas, extintores y botiquines) e inspecciones no planeadas. Los "
        "hallazgos se registran en el Registro de inspecciones internas de SST, indicando la "
        "acción correctiva, el responsable y el plazo.")
    D.h2("8.2 No conformidades y acciones correctivas")
    D.numeros(["Detectar la no conformidad (inspección, auditoría, reporte del trabajador).",
               "Registrar y comunicar al responsable del área.",
               "Analizar la causa raíz (método de los 5 porqués / árbol de causas).",
               "Definir la acción correctiva / preventiva, responsable y plazo.",
               "Verificar la eficacia de la acción y cerrar la no conformidad."])
    D.h2("8.3 Notificación de accidentes e incidentes")
    D.tabla(["Evento", "Plazo de notificación", "Medio"], [
        ["Accidente mortal", "Dentro de las 24 horas de ocurrido", "Sistema de notificación "
         "del MTPE"],
        ["Incidente peligroso", "Dentro de las 24 horas de ocurrido", "Sistema de "
         "notificación del MTPE"],
        ["Accidente de trabajo no mortal", "Hasta el último día hábil del mes siguiente",
         "Centro médico asistencial / MTPE"],
        ["Enfermedad ocupacional", "Dentro de los 5 días hábiles de conocido el diagnóstico",
         "Centro médico asistencial / MTPE"],
    ], [5, 6, 5.5], 8.5)
    D.h2("8.4 Investigación de accidentes")
    D.p("Todo accidente e incidente peligroso será investigado por el Residente, el Ing. "
        "SSOMA y el Comité / Supervisor de SST, con el objetivo de identificar las causas "
        "inmediatas (actos y condiciones subestándar) y básicas (factores personales y del "
        "trabajo), y definir las medidas correctivas. Se registrará en el Registro de "
        "accidentes de trabajo, enfermedades ocupacionales, incidentes peligrosos e incidentes.")
    D.h2("8.5 Estadísticas de seguridad")
    hht = int(round(N * 48 * 4.33))
    D.p("Las estadísticas se calculan mensualmente y de forma acumulada con las horas hombre "
        "trabajadas (HHT). Para {} trabajadores con jornada de 48 horas semanales, las HHT "
        "mensuales estimadas son {:,} y las HHT del proyecto ({} meses) son {:,}."
        .format(N, hht, M, hht * M))
    D.tabla(["Indicador", "Fórmula", "Interpretación"], [
        ["Índice de Frecuencia (IF)", "IF = N° accidentes con tiempo perdido x 1 000 000 / HHT",
         "Accidentes por cada millón de horas trabajadas"],
        ["Índice de Gravedad (IG)", "IG = N° días perdidos x 1 000 000 / HHT",
         "Días perdidos por cada millón de horas trabajadas"],
        ["Índice de Accidentabilidad (IA)", "IA = IF x IG / 1 000",
         "Indicador combinado de frecuencia y gravedad"],
    ], [4, 7, 5.5], 8.5)
    D.h2("8.6 Registros obligatorios del SG-SST (R.M. N° 050-2013-TR)")
    D.numeros(["Registro de accidentes de trabajo, enfermedades ocupacionales, incidentes "
               "peligrosos e incidentes.", "Registro de exámenes médicos ocupacionales.",
               "Registro del monitoreo de agentes físicos, químicos, biológicos, "
               "psicosociales y disergonómicos.", "Registro de inspecciones internas de SST.",
               "Registro de estadísticas de SST.", "Registro de equipos de seguridad o "
               "emergencia.", "Registro de inducción, capacitación, entrenamiento y simulacros "
               "de emergencia.", "Registro de auditorías."])

    # ================= 9 =================
    D.h1("9. PLAN DE RESPUESTA ANTE EMERGENCIAS")
    D.h2("9.1 Objetivo")
    D.p("Establecer la organización, los recursos y los procedimientos para responder de forma "
        "rápida y eficaz ante situaciones de emergencia, protegiendo la vida de los "
        "trabajadores y terceros, y minimizando los daños a la propiedad y al ambiente.")
    D.h2("9.2 Organización de brigadas")
    nb = max(2, math.ceil(N / 10))
    D.tabla(["Brigada", "Integrantes", "Funciones principales"], [
        ["Jefe de emergencia", "Residente de obra", "Dirige la respuesta, comunica a entidades "
         "externas y dispone la evacuación"],
        ["Brigada de primeros auxilios", "{} trabajadores capacitados".format(nb), "Atención "
         "inmediata, RCP, inmovilización y traslado del herido"],
        ["Brigada contra incendios", "{} trabajadores capacitados".format(nb), "Uso de "
         "extintores, corte de energía, control de amagos"],
        ["Brigada de evacuación y rescate", "{} trabajadores capacitados".format(nb),
         "Conducción al punto de reunión, conteo de personal, rescate en zanjas / altura"],
    ], [4.5, 4, 8], 8.5)
    D.h2("9.3 Procedimientos de respuesta")
    procedimientos = [
        ("9.3.1 Sismo", [
            "Durante el sismo: suspender los trabajos, apagar equipos y alejarse de taludes, "
            "zanjas, muros y andamios; ubicarse en zonas seguras.",
            "Los trabajadores en altura deben permanecer enganchados y descender una vez "
            "terminado el movimiento.", "Después del sismo: evacuar por las rutas señalizadas "
            "hacia el punto de reunión; el brigadista realiza el conteo de personal.",
            "Inspeccionar excavaciones, encofrados y andamios antes de reanudar los trabajos."]),
        ("9.3.2 Incendio", [
            "Dar la alarma (megáfono / sirena) y comunicar al jefe de emergencia.",
            "Cortar la energía eléctrica del sector afectado.", "Atacar el amago con el "
            "extintor PQS adecuado (técnica PASO: Presionar, Apuntar, Sostener, Oscilar).",
            "Si el fuego no se controla en 30 segundos, evacuar y llamar a Bomberos (116)."]),
        ("9.3.3 Derrumbe de zanja", [
            "Detener toda maquinaria cercana y evitar vibraciones adicionales.",
            "No ingresar a la zanja sin entibado de emergencia; llamar a Bomberos (116).",
            "Ubicar a la víctima y despejar primero la cabeza y el tórax para permitir la "
            "respiración, con herramientas manuales.", "Brindar primeros auxilios y trasladar "
            "al centro de salud más cercano."]),
        ("9.3.4 Electrocución", [
            "NO tocar a la víctima mientras esté en contacto con la fuente eléctrica.",
            "Cortar la energía desde el tablero o separar a la víctima con un elemento "
            "aislante seco (pértiga, madera).", "Verificar signos vitales; iniciar RCP si no "
            "respira.", "Tratar quemaduras con apósitos estériles y trasladar de inmediato."]),
        ("9.3.5 Caída de altura / rescate", [
            "Activar el plan de rescate: el trabajador suspendido debe ser rescatado en menos "
            "de 15 minutos (síndrome del arnés).", "No mover a la víctima si se sospecha "
            "lesión de columna; inmovilizar con collarín y camilla rígida.",
            "Llamar a SAMU (106) / EsSalud (117) y trasladar al centro de salud."]),
        ("9.3.6 Primeros auxilios", [
            "Evaluar la escena y protegerse (guantes).", "Verificar conciencia, respiración y "
            "pulso (ABC).", "Controlar hemorragias con presión directa.", "Inmovilizar "
            "fracturas con férulas.", "Mantener a la víctima abrigada y trasladar con la "
            "unidad de emergencia de la obra o la ambulancia."]),
    ]
    for tit, pasos in procedimientos:
        D.h3(tit)
        D.vinetas(pasos)
    D.h2("9.4 Directorio telefónico de emergencia")
    D.tabla(["Institución / Responsable", "Teléfono"], [list(x) for x in DIRECTORIO_EMERGENCIA],
            [11, 5.5], 9)
    D.h2("9.5 Contenido obligatorio del botiquín (Norma G.050)")
    D.tabla(["Elemento", "Cantidad mínima"], [list(x) for x in BOTIQUIN_G050], [11, 5.5], 8.5)
    D.h2("9.6 Equipos de emergencia en obra")
    D.vinetas(["Extintores PQS ABC de 6 kg y 12 kg, señalizados, con tarjeta de inspección "
               "mensual.", "Botiquín(es) de primeros auxilios y camilla rígida con collarín "
               "cervical.", "Estación de emergencia con directorio, linterna y megáfono.",
               "Unidad vehicular disponible para traslado de heridos durante toda la jornada.",
               "Lavaojos portátil en zonas de manipulación de químicos y concreto."])
    D.h2("9.7 Simulacros")
    D.p("Se ejecutará como mínimo un simulacro trimestral (sismo, incendio, rescate en zanja o "
        "altura y primeros auxilios), evaluando tiempos de respuesta y oportunidades de "
        "mejora. Los resultados se registran en el Registro de inducción, capacitación, "
        "entrenamiento y simulacros de emergencia.")

    # ================= 10 =================
    D.h1("10. PRESUPUESTO DE IMPLEMENTACIÓN DEL PSST")
    D.p("El presupuesto de seguridad se ha calculado en función de {} obreros, {} profesionales "
        "del staff técnico y un plazo de {} meses, aplicando factores de reposición por "
        "desgaste según la vida útil de cada equipo. Precios referenciales en soles sin IGV."
        .format(cfg["n_obr"], cfg["n_ing"], M))
    D.h2("10.1 Resumen por subpartidas (Norma G.050)")
    filas_res = [[s["codigo"], s["titulo"], soles(s["subtotal"]),
                  "{:.2f} %".format(100.0 * s["subtotal"] / total if total else 0)]
                 for s in desag["secciones"]]
    filas_res.append(["01", "COSTO DIRECTO TOTAL DE SEGURIDAD Y SALUD", soles(total),
                      "100.00 %"])
    D.tabla(["Ítem", "Subpartida", "Monto", "%"], filas_res, [1.5, 9.5, 3.3, 2.2], 9,
            resaltar=lambda i, j, v: ("DDEBF7" if i == len(filas_res) - 1 else None))
    D.h2("10.2 Presupuesto desagregado")
    filas_det = []
    for s in desag["secciones"]:
        filas_det.append([s["codigo"], s["titulo"], "", "", "", soles(s["subtotal"])])
        for it in s["items"]:
            filas_det.append([it["codigo"], it["desc"], it["und"], str(it["cant"]),
                              "{:,.2f}".format(it["pu"]), "{:,.2f}".format(it["parcial"])])
    filas_det.append(["", "COSTO DIRECTO TOTAL DE SEGURIDAD Y SALUD", "", "", "", soles(total)])
    titulos = set([s["codigo"] for s in desag["secciones"]] + [""])
    D.tabla(["Ítem", "Descripción", "Und.", "Cant.", "P.U. (S/)", "Parcial (S/)"], filas_det,
            [1.7, 8.6, 1.3, 1.3, 1.6, 2.0], 7.5,
            resaltar=lambda i, j, v: ("DDEBF7" if filas_det[i][0] in titulos else None))

    # ================= 11 =================
    D.h1("11. ANEXOS Y FIRMAS")
    D.vinetas([
        "Anexo 01: Matriz IPERC completa por partidas (Excel - hoja 1_Matriz_IPERC_Obra).",
        "Anexo 02: Presupuesto desagregado de seguridad G.050 con fórmulas (Excel - hoja "
        "2_Presupuesto_Seguridad_G050).", "Anexo 03: Programa de capacitaciones, simulacros e "
        "inspecciones (Excel - hoja 3_Programa_Capacitaciones).", "Anexo 04: Formatos ATS, "
        "PETAR y check-lists (Excel - hoja 4_Formatos_ATS_y_PETAR).", "Anexo 05: Lámina SS-01 "
        "- Instalaciones provisionales, señalización temporal y detalles G.050 (AutoCAD .DXF)."])
    doc.add_paragraph()
    t = doc.add_table(rows=1, cols=3)
    for j, (nombre, cargo) in enumerate(((cfg["ssoma"], "Especialista SSOMA - CIP N° " +
                                          cfg["cip"]), (cfg["residente"], "Residente de Obra"),
                                         ("", "Supervisión / Entidad"))):
        c = t.rows[0].cells[j]
        c.text = ""
        par = c.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.add_run("\n\n\n__________________________\n{}\n{}".format(nombre, cargo)).font.size = \
            Pt(9)

    # actualizar campos (índice) al abrir
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    doc.settings.element.append(upd)
    D.guardar(ruta)
    return ruta


# =============================================================================
# 10. LÁMINA SS-01 EN AUTOCAD (.DXF) CON EZDXF
#     Hoja A1 (841 x 594 mm) dibujada en milímetros de papel; cada detalle se
#     dibuja en metros reales y se escala (k = mm de papel por metro real).
# =============================================================================
ALINEACIONES = {
    "MC": TextEntityAlignment.MIDDLE_CENTER, "ML": TextEntityAlignment.MIDDLE_LEFT,
    "MR": TextEntityAlignment.MIDDLE_RIGHT, "BL": TextEntityAlignment.BOTTOM_LEFT,
    "BC": TextEntityAlignment.BOTTOM_CENTER, "TL": TextEntityAlignment.TOP_LEFT,
    "TC": TextEntityAlignment.TOP_CENTER,
}
CAPAS_DXF = [
    ("SS-DETALLES", 7, "CONTINUOUS"), ("SS-SENALETICA", 1, "CONTINUOUS"),
    ("SS-TEXTOS", 7, "CONTINUOUS"), ("SS-COTAS", 3, "CONTINUOUS"),
    ("SS-MARCO", 8, "CONTINUOUS"), ("SS-ACHURADO", 9, "CONTINUOUS"),
    ("SS-CERCO", 30, "DASHED"), ("SS-EVACUACION", 3, "CONTINUOUS"),
]


class Dib:
    """Dibujante con origen (ox, oy) en mm de papel y escala k (mm por metro)."""

    def __init__(self, msp, ox=0.0, oy=0.0, k=1.0):
        self.msp, self.ox, self.oy, self.k = msp, ox, oy, k

    def P(self, x, y):
        return (self.ox + x * self.k, self.oy + y * self.k)

    @staticmethod
    def _at(capa, color=None, lt=None):
        d = {"layer": capa}
        if color is not None:
            d["color"] = color
        if lt:
            d["linetype"] = lt
        return d

    def linea(self, x1, y1, x2, y2, capa="SS-DETALLES", color=None, lt=None, grosor=None):
        at = self._at(capa, color, lt)
        if grosor:
            at["lineweight"] = grosor
        return self.msp.add_line(self.P(x1, y1), self.P(x2, y2), dxfattribs=at)

    def poli(self, pts, cerrada=False, capa="SS-DETALLES", color=None, lt=None, grosor=None):
        at = self._at(capa, color, lt)
        if grosor:
            at["lineweight"] = grosor
        return self.msp.add_lwpolyline([self.P(x, y) for x, y in pts], close=cerrada,
                                       dxfattribs=at)

    def rect(self, x1, y1, x2, y2, capa="SS-DETALLES", color=None, lt=None, grosor=None):
        return self.poli([(x1, y1), (x2, y1), (x2, y2), (x1, y2)], True, capa, color, lt, grosor)

    def circulo(self, x, y, r, capa="SS-DETALLES", color=None):
        return self.msp.add_circle(self.P(x, y), r * self.k, dxfattribs=self._at(capa, color))

    def texto(self, s, x, y, h=2.0, capa="SS-TEXTOS", alin="MC", color=None, rot=0.0):
        at = self._at(capa, color)
        at["style"] = "OpenSans"
        if rot:
            at["rotation"] = rot
        t = self.msp.add_text(s, height=h, dxfattribs=at)
        t.set_placement(self.P(x, y), align=ALINEACIONES[alin])
        return t

    def mtexto(self, s, x, y, h=1.8, ancho=60.0, capa="SS-TEXTOS", color=None):
        at = self._at(capa, color)
        at.update({"char_height": h, "width": ancho, "attachment_point": 1,
                   "insert": self.P(x, y), "style": "OpenSans"})
        return self.msp.add_mtext(s, dxfattribs=at)

    def achurado(self, pts, color=7, patron=None, escala=1.0, capa="SS-ACHURADO", angulo=0.0):
        h = self.msp.add_hatch(color=color, dxfattribs={"layer": capa})
        if patron:
            h.set_pattern_fill(patron, scale=escala, angle=angulo)
        h.paths.add_polyline_path([self.P(x, y) for x, y in pts], is_closed=True)
        return h

    def solido_rect(self, x1, y1, x2, y2, color, capa="SS-ACHURADO"):
        return self.achurado([(x1, y1), (x2, y1), (x2, y2), (x1, y2)], color, None, 1, capa)

    def cota_h(self, x1, x2, y, y_cota, txt, h=1.8):
        d = self.msp.add_linear_dim(
            base=self.P(x1, y_cota), p1=self.P(x1, y), p2=self.P(x2, y), angle=0,
            dimstyle="EZDXF", text=txt,
            override={"dimtxt": h, "dimasz": 1.4, "dimexe": 0.8, "dimexo": 0.6, "dimgap": 0.5,
                      "dimtad": 1, "dimclrd": 3, "dimclre": 3, "dimclrt": 3},
            dxfattribs={"layer": "SS-COTAS"})
        d.render()
        return d

    def cota_v(self, y1, y2, x, x_cota, txt, h=1.8):
        d = self.msp.add_linear_dim(
            base=self.P(x_cota, y1), p1=self.P(x, y1), p2=self.P(x, y2), angle=90,
            dimstyle="EZDXF", text=txt,
            override={"dimtxt": h, "dimasz": 1.4, "dimexe": 0.8, "dimexo": 0.6, "dimgap": 0.5,
                      "dimtad": 1, "dimclrd": 3, "dimclre": 3, "dimclrt": 3},
            dxfattribs={"layer": "SS-COTAS"})
        d.render()
        return d

    def llamada(self, x1, y1, x2, y2, s, h=1.6, alin="ML"):
        """Línea de referencia desde un punto del dibujo hasta el texto (coords reales)."""
        self.linea(x1, y1, x2, y2, "SS-TEXTOS", 7)
        self.circulo(x1, y1, 0.35 / self.k, "SS-TEXTOS", 7)
        dx = 1.0 / self.k if alin == "ML" else -1.0 / self.k
        at = {"layer": "SS-TEXTOS", "char_height": h, "insert": self.P(x2 + dx, y2),
              "attachment_point": 4 if alin == "ML" else 6, "style": "OpenSans"}
        m = self.msp.add_mtext(s, dxfattribs=at)
        m.set_bg_color("canvas", scale=1.2)  # máscara de fondo: legible sobre achurados
        return m

    def flecha(self, x1, y1, x2, y2, capa="SS-EVACUACION", color=3, tam=2.0):
        """Flecha (coords reales); tam en mm de papel."""
        self.linea(x1, y1, x2, y2, capa, color, grosor=35)
        ang = math.atan2(y2 - y1, x2 - x1)
        L = tam / self.k
        a1 = (x2 - L * math.cos(ang - 0.4), y2 - L * math.sin(ang - 0.4))
        a2 = (x2 - L * math.cos(ang + 0.4), y2 - L * math.sin(ang + 0.4))
        self.achurado([(x2, y2), a1, a2], color, None, 1, capa)


def _titulo_detalle(dp, x, y, numero, titulo, escala, ancho=90):
    """Título normalizado de detalle (dp = dibujante en mm de papel)."""
    dp.circulo(x, y, 5.0, "SS-TEXTOS", 7)
    dp.texto(numero, x, y, 4.0, alin="MC")
    dp.texto(titulo, x + 8, y + 1.2, 3.2, alin="BL")
    dp.linea(x + 8, y, x + 8 + ancho, y, "SS-TEXTOS", 7, grosor=50)
    dp.texto("ESCALA " + escala, x + 8, y - 1.2, 2.2, alin="TL")


def _senal_obligacion(dp, cx, cy, r, texto):
    dp.achurado([(cx + r * math.cos(a * math.pi / 18), cy + r * math.sin(a * math.pi / 18))
                 for a in range(36)], 5, None, 1, "SS-SENALETICA")
    dp.circulo(cx, cy, r, "SS-SENALETICA", 5)
    dp.texto(texto, cx, cy, r * 0.32, "SS-SENALETICA", "MC", 7)


def _senal_prohibicion(dp, cx, cy, r, texto):
    dp.circulo(cx, cy, r, "SS-SENALETICA", 1)
    dp.circulo(cx, cy, r * 0.8, "SS-SENALETICA", 1)
    anillo = [(cx + r * math.cos(a * math.pi / 18), cy + r * math.sin(a * math.pi / 18))
              for a in range(36)]
    h = dp.msp.add_hatch(color=1, dxfattribs={"layer": "SS-SENALETICA"})
    h.paths.add_polyline_path(anillo, is_closed=True, flags=1)
    h.paths.add_polyline_path([(cx + 0.8 * r * math.cos(a * math.pi / 18),
                                cy + 0.8 * r * math.sin(a * math.pi / 18)) for a in range(36)],
                              is_closed=True, flags=16)
    d = 0.8 * r * 0.7071
    dp.linea(cx - d, cy + d, cx + d, cy - d, "SS-SENALETICA", 1, grosor=100)
    dp.texto(texto, cx, cy - r - 2.5, 1.6, "SS-SENALETICA", "MC", 7)


def _senal_advertencia(dp, cx, cy, lado, texto):
    hgt = lado * math.sqrt(3) / 2
    pts = [(cx - lado / 2, cy - hgt / 3), (cx + lado / 2, cy - hgt / 3), (cx, cy + 2 * hgt / 3)]
    dp.achurado(pts, 2, None, 1, "SS-SENALETICA")
    dp.poli(pts, True, "SS-SENALETICA", 250, grosor=70)
    dp.texto("!", cx, cy + 0.5, lado * 0.35, "SS-SENALETICA", "MC", 250)
    dp.texto(texto, cx, cy - hgt / 3 - 2.5, 1.6, "SS-SENALETICA", "MC", 7)


def _senal_rect(dp, cx, cy, w, h, color, texto, txt_color=7):
    dp.solido_rect(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2, color, "SS-SENALETICA")
    dp.rect(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2, "SS-SENALETICA", color)
    dp.texto(texto, cx, cy, min(h * 0.3, 2.6), "SS-SENALETICA", "MC", txt_color)


def _cachaco(d, x, y0, alto=1.20, ancho=0.05, base=0.30, h_base=0.20):
    """Poste de señalización tipo 'cachaco' (en metros reales) apoyado en y0."""
    d.rect(x - base / 2, y0, x + base / 2, y0 + h_base)
    d.achurado([(x - base / 2, y0), (x + base / 2, y0), (x + base / 2, y0 + h_base),
                (x - base / 2, y0 + h_base)], 8, "AR-CONC", 0.02 * d.k)
    yy = y0 + h_base
    franja = 0.20
    n = 0
    while yy < y0 + alto - 1e-6:
        y2 = min(yy + franja, y0 + alto)
        d.solido_rect(x - ancho / 2, yy, x + ancho / 2, y2, 2 if n % 2 == 0 else 250)
        yy = y2
        n += 1
    d.rect(x - ancho / 2, y0 + h_base, x + ancho / 2, y0 + alto)


def exportar_dxf(ruta, cfg, desag=None):
    cfg = completar_config(cfg)
    doc = ezdxf.new("R2010", setup=True)
    doc.units = ezdxf.units.MM
    doc.header["$LTSCALE"] = 2.0
    for nombre, color, lt in CAPAS_DXF:
        if nombre not in doc.layers:
            doc.layers.add(nombre, color=color, linetype=lt)
    msp = doc.modelspace()
    dp = Dib(msp)  # papel (mm)

    # ---------------- marco A1 ----------------
    dp.rect(0, 0, 841, 594, "SS-MARCO")
    dp.rect(20, 10, 831, 584, "SS-MARCO", grosor=70)
    for x in (365, 600):
        dp.linea(x, 10, x, 584, "SS-MARCO")
    dp.linea(20, 172, 365, 172, "SS-MARCO")
    dp.linea(20, 325, 365, 325, "SS-MARCO")
    dp.linea(365, 395, 831, 395, "SS-MARCO")
    dp.linea(365, 185, 831, 185, "SS-MARCO")

    # ======================================================================
    # (A) LAYOUT DE INSTALACIONES PROVISIONALES  - 1:250 (k = 4)
    # ======================================================================
    _titulo_detalle(dp, 32, 572, "A", "LAYOUT TIPO DE INSTALACIONES PROVISIONALES DE OBRA",
                    "1:250", 150)
    L = Dib(msp, 30, 352, 4.0)
    W, H = 80.0, 50.0
    # vía pública
    L.linea(-2, -1.5, W + 2, -1.5, "SS-DETALLES", 8)
    L.linea(-2, -6.0, W + 2, -6.0, "SS-DETALLES", 8)
    L.linea(-2, -3.75, W + 2, -3.75, "SS-DETALLES", 2, "DASHED")
    L.texto("VÍA PÚBLICA (CALLE / AV.) - DESVÍO SEÑALIZADO", 18, -4.9, 1.6)
    for xc in range(1, 80, 6):
        if 32 <= xc <= 47:
            continue
        L.achurado([(xc - 0.35, -1.2), (xc + 0.35, -1.2), (xc, 0.0 - 0.1)], 30, None, 1,
                   "SS-SENALETICA")
    L.texto("CONOS 28\"", 60, -2.6, 1.4, color=30)
    # cerco perimétrico con accesos
    tramos = [((0, 0), (34, 0)), ((40, 0), (43, 0)), ((45, 0), (W, 0)), ((W, 0), (W, H)),
              ((W, H), (0, H)), ((0, H), (0, 0))]
    for (x1, y1), (x2, y2) in tramos:
        L.linea(x1, y1, x2, y2, "SS-CERCO", 30, grosor=50)
        L.linea(x1 + (0.4 if y1 == y2 else 0), y1 + (0.4 if x1 == x2 else 0),
                x2 + (0.4 if y1 == y2 else 0), y2 + (0.4 if x1 == x2 else 0), "SS-CERCO", 30)
    for (x1, y1), (x2, y2) in tramos:
        n = int(max(abs(x2 - x1), abs(y2 - y1)) // 3)
        for i in range(n + 1):
            t = i / float(max(n, 1))
            px, py = x1 + (x2 - x1) * t, y1 + (y2 - y1) * t
            L.rect(px - 0.2, py - 0.2, px + 0.2, py + 0.2, "SS-CERCO", 30)
    L.texto("CERCO PERIMÉTRICO H = 2.40 m (MALLA RASCHEL / TRIPLAY) CON POSTES @ 3.00 m", 40,
            H + 1.1, 1.6, color=30)
    L.texto("INGRESO VEHICULAR", 37, 4.2, 1.5)
    L.texto("PEATONAL", 44, -0.9, 1.3)
    L.flecha(37, -3.0, 37, 3.0, "SS-EVACUACION", 7, 2.0)

    def ambiente(x1, y1, x2, y2, nombre, sub=None, color=None, h=1.6):
        L.rect(x1, y1, x2, y2, "SS-DETALLES", color, grosor=35)
        cy = (y1 + y2) / 2
        L.texto(nombre, (x1 + x2) / 2, cy + (0.6 if sub else 0), h)
        if sub:
            L.texto(sub, (x1 + x2) / 2, cy - 1.0, 1.2)

    ambiente(46, 1, 49.5, 4.5, "GUARD.", None, None, 1.3)
    ambiente(51, 1, 63, 5.5, "OFICINAS TÉCNICAS", "RESIDENCIA | SSOMA | SUPERVISIÓN")
    ambiente(64, 1, 68, 5.5, "TÓPICO", "1.er AUXILIO", None, 1.4)
    ambiente(69, 1, 79, 7, "ALMACÉN GENERAL")
    # almacén de químicos con bandeja antiderrame
    L.rect(71, 10, 79, 14, "SS-DETALLES", 1, grosor=50)
    L.rect(71.4, 10.4, 78.6, 13.6, "SS-DETALLES", 1, "DASHED")
    L.texto("ALM. QUÍMICOS", 75, 12.6, 1.4, color=1)
    L.texto("BANDEJA ANTIDERRAME + MSDS", 75, 11.2, 0.95, color=1)
    # SS.HH. según G.050
    n_ino = sshh_requeridos(cfg["n_total"])[0]
    n_cab = max(n_ino + 1, 3)
    for i in range(n_cab):
        L.rect(2 + i * 1.5, 1, 3.3 + i * 1.5, 2.4, "SS-DETALLES", 4)
    L.texto("BAÑOS QUÍMICOS ({} und - G.050)".format(n_cab), 2 + n_cab * 0.75, 3.2, 1.3)
    ambiente(2, 5, 10, 9, "VESTUARIOS", "Y DUCHAS")
    ambiente(12, 1, 21, 6.5, "COMEDOR", "CAP. {} PERS.".format(cfg["n_total"]))
    # zonas de acopio
    ambiente(2, 40, 16, 46, "ACOPIO DE ACERO", "SOBRE DURMIENTES")
    ambiente(18, 40, 27, 46, "HABILITADO", "DE ACERO")
    ambiente(29, 40, 39, 46, "ACOPIO DE MADERA", "/ ENCOFRADOS")
    ambiente(41, 40, 49, 46, "AGREGADOS", "(CUBIERTOS)")
    ambiente(51, 40, 57, 46, "MEZCLA", "CONCRETO")
    # punto de acopio de RR.SS.
    L.rect(60, 39, 79, 47, "SS-DETALLES", 3, grosor=35)
    L.texto("PUNTO DE ACOPIO DE RR.SS. (NTP 900.058)", 69.5, 45.8, 1.3)
    colores_rs = [5, 7, 2, 9, 32, 250, 1]
    for i, c in enumerate(colores_rs):
        cx = 61.8 + i * 2.5
        L.achurado([(cx + 0.8 * math.cos(a * math.pi / 10), 42.2 + 0.8 * math.sin(a * math.pi / 10))
                    for a in range(20)], c, None, 1, "SS-SENALETICA")
        L.circulo(cx, 42.2, 0.8, "SS-SENALETICA", 7)
        L.texto(("AZ", "BL", "AM", "PL", "MR", "NG", "RJ")[i], cx, 40.4, 1.0)
    # área de construcción
    L.rect(28, 15, 62, 35, "SS-DETALLES", 7, grosor=70)
    L.achurado([(28, 15), (62, 15), (62, 35), (28, 35)], 8, "ANSI31", 1.5)
    L.rect(26, 13, 64, 37, "SS-CERCO", 30, "DASHED")
    L.texto("ÁREA DE CONSTRUCCIÓN", 45, 26.5, 2.2)
    L.texto("(BLOQUE PRINCIPAL / FRENTE DE TRABAJO)", 45, 23.5, 1.5)
    L.texto("ZONA DE EXCLUSIÓN 2.00 m: MALLA NARANJA + CACHACOS", 45, 38.0, 1.3, color=30)
    # punto de reunión
    L.rect(6, 17, 16, 27, "SS-EVACUACION", 3, grosor=50)
    for ang in (0, 90, 180, 270):
        a = math.radians(ang)
        L.flecha(11 + 4.2 * math.cos(a), 22 + 4.2 * math.sin(a), 11 + 1.6 * math.cos(a),
                 22 + 1.6 * math.sin(a), "SS-EVACUACION", 3, 2.2)
    L.texto("PR", 11, 22, 2.2, color=3)
    L.texto("PUNTO DE REUNIÓN EN CASO DE SISMO", 11, 28.2, 1.3, color=3)
    # rutas de evacuación
    L.poli([(28, 25), (22, 25), (17, 22)], False, "SS-EVACUACION", 3, grosor=35)
    L.flecha(18.5, 22.8, 16.4, 22.0, "SS-EVACUACION", 3, 2.0)
    L.poli([(45, 13), (45, 10), (11, 10)], False, "SS-EVACUACION", 3, grosor=35)
    L.flecha(11, 10, 11, 16.8, "SS-EVACUACION", 3, 2.0)
    L.texto("RUTA DE EVACUACIÓN", 30, 10.8, 1.2, color=3)
    # tablero eléctrico, extintores y botiquines
    L.rect(65, 19, 67, 21, "SS-SENALETICA", 2)
    L.texto("TG", 66, 20, 1.2, color=2)
    L.texto("TABLERO PROV. + POZO TIERRA", 66, 17.8, 1.0)
    for x, y in ((50, 6.3), (67.2, 6.3), (70.2, 8), (72, 12), (21.8, 4), (64, 22.5), (27, 15),
                 (63, 35), (78.5, 38)):
        L.achurado([(x + 0.55 * math.cos(a * math.pi / 8), y + 0.55 * math.sin(a * math.pi / 8))
                    for a in range(16)], 1, None, 1, "SS-SENALETICA")
        L.texto("E", x, y, 1.0, "SS-SENALETICA", "MC", 7)
    for x, y in ((64.9, 6.3), (27, 36)):
        L.solido_rect(x - 0.6, y - 0.6, x + 0.6, y + 0.6, 3, "SS-SENALETICA")
        L.texto("+", x, y, 1.4, "SS-SENALETICA", "MC", 7)
    L.texto("PATIO DE MANIOBRAS", 37, 8.0, 1.5)
    L.texto("VEL. MÁX. 10 km/h", 37, 6.4, 1.3, color=1)
    # norte
    L.poli([(76, 30), (77.5, 34), (79, 30), (77.5, 31.2)], True, "SS-DETALLES", 7)
    L.achurado([(76, 30), (77.5, 34), (77.5, 31.2)], 7, None, 1, "SS-DETALLES")
    L.texto("N", 77.5, 35.3, 2.0)
    # cotas generales
    L.cota_h(0, W, H, H + 3.0, "{:.2f}".format(W))
    L.cota_v(0, H, 0, -3.5, "{:.2f}".format(H))

    # ======================================================================
    # (B) LEYENDA DEL LAYOUT + TABLA DE SS.HH.
    # ======================================================================
    dp.texto("LEYENDA DEL LAYOUT", 30, 317, 3.0, alin="BL")
    ley = [("E", 1, "Extintor PQS ABC 6/12 kg (NTP 350.043-1)"),
           ("+", 3, "Botiquín de primeros auxilios (G.050)"),
           ("PR", 3, "Punto de reunión / zona segura en caso de sismo"),
           ("->", 3, "Ruta de evacuación"), ("TG", 2, "Tablero eléctrico provisional "
                                                       "(diferencial 30 mA + puesta a tierra)"),
           ("---", 30, "Cerco perimétrico / zona de exclusión (malla naranja)"),
           ("RS", 3, "Punto de acopio de residuos sólidos por colores"),
           ("Q", 1, "Almacén de productos químicos con bandeja antiderrame y MSDS")]
    for i, (sim, col, txt) in enumerate(ley):
        y = 309 - i * 6.5
        dp.rect(32, y - 2.2, 40, y + 2.2, "SS-SENALETICA", col)
        dp.texto(sim, 36, y, 1.8, "SS-SENALETICA", "MC", col)
        dp.texto(txt, 43, y, 2.0, alin="ML")
    # tabla SS.HH.
    tx, ty = 200, 312
    dp.texto("SERVICIOS HIGIÉNICOS - NORMA G.050", tx, ty + 3, 2.4, alin="BL")
    cols = [tx, tx + 34, tx + 66, tx + 98, tx + 130, tx + 155]
    filas = [("N° TRABAJ.", "INODOROS", "LAVATORIOS", "DUCHAS", "URINARIOS")] + \
            [tuple(str(v) for v in f) for f in TABLA_SSHH]
    nt = cfg["n_total"]
    fila_act = 1 if nt <= 9 else 2 if nt <= 24 else 3 if nt <= 49 else 4 if nt <= 100 else 5
    for i, f in enumerate(filas):
        y1, y2 = ty - i * 7, ty - (i + 1) * 7
        if i == fila_act:
            dp.solido_rect(cols[0], y2, cols[-1], y1, 150, "SS-ACHURADO")
        for j in range(5):
            dp.rect(cols[j], y2, cols[j + 1], y1, "SS-DETALLES", 7)
            dp.texto(f[j], (cols[j] + cols[j + 1]) / 2, (y1 + y2) / 2, 2.0 if i else 1.8)
    dp.texto("Obra con {} trabajadores: se resalta la fila aplicable.".format(nt), tx,
             ty - 46, 1.8, alin="BL")
    dp.mtexto("ALMACÉN DE MATERIALES PELIGROSOS: ventilado, techado, piso impermeable, "
              "bandeja antiderrame (110 % del mayor envase), kit antiderrame, extintor PQS, "
              "hojas MSDS en español y señal \"PROHIBIDO FUMAR\".", tx, ty - 51, 1.8, 150)

    # ======================================================================
    # (C) NOTAS GENERALES
    # ======================================================================
    dp.texto("NOTAS GENERALES", 30, 163, 3.0, alin="BL")
    notas = ("1. La distribución es referencial y se adecuará al terreno disponible de la obra.\\P"
             "2. Toda excavación > 1.50 m de profundidad se entiba y se señaliza con cachacos, "
             "malla naranja H = 1.00 m y cinta \"PELIGRO\" a 1.00 m del borde.\\P"
             "3. Trabajos en altura > 1.80 m: andamio normado, plataforma completa, baranda a "
             "0.90 m, rodapié de 0.15 m y arnés con doble línea de vida (ANSI Z359).\\P"
             "4. Las señales cumplen la NTP 399.010-1 (formas, colores y pictogramas).\\P"
             "5. Extintores PQS a máx. 1.50 m de altura (parte superior), señalizados y con "
             "tarjeta de inspección mensual.\\P"
             "6. Los residuos se segregan según NTP 900.058:2019 y se disponen mediante EO-RS "
             "autorizada (D.S. N° 002-2022-VIVIENDA).\\P"
             "7. La velocidad máxima de vehículos en obra es 10 km/h; todo retroceso requiere "
             "vigía.\\P"
             "8. Personal promedio: {} obreros + {} staff. Plazo: {} días calendario.\\P"
             "9. Costo directo de seguridad G.050 estimado: {}."
             .format(cfg["n_obr"], cfg["n_ing"], cfg["dias"],
                     soles(desag["total"]) if desag else "ver presupuesto"))
    dp.mtexto(notas, 30, 156, 2.1, 325)

    # ======================================================================
    # DETALLE 01: CACHACO  - 1:20 (k = 50)
    # ======================================================================
    _titulo_detalle(dp, 377, 576, "1", "POSTE DE SEÑALIZACIÓN (CACHACO) Y MALLA", "1:20", 110)
    C = Dib(msp, 392, 452, 50.0)
    # terreno y borde de excavación
    C.linea(-0.25, 0, 3.30, 0, grosor=50)
    C.linea(3.30, 0, 3.30, -0.45, grosor=50)
    C.achurado([(-0.25, 0), (3.30, 0), (3.30, -0.45), (-0.25, -0.45)], 8, "EARTH", 0.02 * 50)
    C.linea(3.30, -0.45, 3.65, -0.45, lt="DASHED")
    C.texto("EXCAVACIÓN", 3.55, -0.2, 1.5, rot=90)
    _cachaco(C, 0.30, 0.0)
    _cachaco(C, 2.30, 0.0)
    # malla naranja H = 1.00 m entre postes
    yb, yt = 0.15, 1.15
    xs = 0.30 + 0.025
    xe = 2.30 - 0.025
    C.rect(xs, yb, xe, yt, "SS-SENALETICA", 30)
    yy = yb
    while yy < yt - 1e-6:
        C.linea(xs, yy, xe, yy, "SS-SENALETICA", 30)
        yy += 0.10
    xx = xs
    while xx < xe - 1e-6:
        C.linea(xx, yb, xx, yt, "SS-SENALETICA", 30)
        xx += 0.10
    # cinta de peligro
    C.linea(0.30, 1.18, 2.30, 1.18, "SS-SENALETICA", 1, grosor=70)
    C.linea(0.30, 1.16, 2.30, 1.16, "SS-SENALETICA", 2, grosor=70)
    # cotas
    C.cota_v(0, 1.20, 0.30, -0.15, "1.20")
    C.cota_h(0.15, 0.45, 0, -0.32, "0.30")
    C.cota_h(0.30, 2.30, 1.20, 1.42, "2.00 (máx. 2.50)")
    C.cota_h(2.30, 3.30, 0, -0.32, "1.00 MÍN.")
    C.cota_v(yb, yt, 2.30, 2.55, "1.00")
    # llamadas
    C.llamada(0.30, 0.9, -0.05, 2.05, "POSTE DE MADERA 2\"x2\" / TUBO PVC Ø2\" PINTADO "
              "EN FRANJAS AMARILLO/NEGRO @ 0.20 m")
    C.llamada(1.30, 0.75, 0.95, 1.85, "MALLA DE SEGURIDAD NARANJA HDPE H = 1.00 m")
    C.llamada(1.80, 1.17, 2.45, 1.62, "CINTA \"PELIGRO\" AMARILLA / ROJA")
    C.llamada(0.40, 0.10, 0.70, -0.62, "BASE DE CONCRETO CICLÓPEO 0.30 x 0.30 x 0.20 m")
    C.texto("BORDE DE EXCAVACIÓN", 2.80, 0.08, 1.4, alin="BC")

    # ======================================================================
    # DETALLE 02: ENTIBADO DE ZANJA  - 1:40 (k = 25)
    # ======================================================================
    _titulo_detalle(dp, 612, 576, "2", "ENTIBADO TÍPICO DE ZANJA PROFUNDA (H > 1.50 m)", "1:40",
                    150)
    Z = Dib(msp, 700, 498, 25.0)
    Hz, B = 2.50, 1.20
    d_acopio = max(Hz / 2.0, 0.60)
    tierra = [(-1.3, 0), (0, 0), (0, -Hz), (B, -Hz), (B, 0), (B + 1.3, 0), (B + 1.3, -Hz - 0.5),
              (-1.3, -Hz - 0.5)]
    Z.achurado(tierra, 8, "EARTH", 0.05 * 25)
    Z.poli([(-3.2, 0), (0, 0), (0, -Hz), (B, -Hz), (B, 0), (B + 3.0, 0)], False, grosor=50)
    # tablones verticales (entibado continuo)
    for x1, x2 in ((0.0, 0.05), (B - 0.05, B)):
        Z.rect(x1, -Hz, x2, 0.20, "SS-DETALLES", 40)
        Z.solido_rect(x1, -Hz, x2, 0.20, 40, "SS-ACHURADO")
    # largueros y puntales
    for yl in (-0.40, -1.40, -2.20):
        Z.rect(0.05, yl - 0.10, 0.15, yl + 0.10, "SS-DETALLES", 42)
        Z.rect(B - 0.15, yl - 0.10, B - 0.05, yl + 0.10, "SS-DETALLES", 42)
        Z.rect(0.15, yl - 0.05, B - 0.15, yl + 0.05, "SS-DETALLES", 7)
        Z.solido_rect(0.15, yl - 0.05, B - 0.15, yl + 0.05, 9, "SS-ACHURADO")
    # escalera de escape que sobresale 1.00 m
    xb, xt = B - 0.95, B - 0.10
    for dx in (0.0, 0.12):
        Z.linea(xb + dx, -Hz, xt + dx, 1.0, "SS-DETALLES", 6, grosor=35)
    n_pel = int((Hz + 1.0) / 0.30)
    for i in range(1, n_pel + 1):
        t = i / float(n_pel + 1)
        x = xb + (xt - xb) * t
        y = -Hz + (Hz + 1.0) * t
        Z.linea(x, y, x + 0.12, y, "SS-DETALLES", 6)
    # acopio de material excavado
    Z.poli([(-d_acopio, 0), (-d_acopio - 1.0, 0.85), (-d_acopio - 2.0, 0)], True, grosor=35)
    Z.achurado([(-d_acopio, 0), (-d_acopio - 1.0, 0.85), (-d_acopio - 2.0, 0)], 33, "DOTS",
               0.02 * 25)
    Z.texto("MATERIAL EXCAVADO", -d_acopio - 1.0, 0.25, 1.4, alin="BC")
    # cachacos con malla en el lado opuesto
    for xc in (B + 1.0, B + 2.6):
        _cachaco(Z, xc, 0.0, 1.20, 0.06, 0.30, 0.20)
    Z.linea(B + 1.0, 0.15, B + 2.6, 0.15, "SS-SENALETICA", 30)
    Z.linea(B + 1.0, 1.15, B + 2.6, 1.15, "SS-SENALETICA", 30)
    for xm in (B + 1.3, B + 1.6, B + 1.9, B + 2.2):
        Z.linea(xm, 0.15, xm, 1.15, "SS-SENALETICA", 30)
    # cotas
    Z.cota_v(-Hz, 0, 0, -0.45, "H = {:.2f}".format(Hz))
    Z.cota_h(0, B, -Hz, -Hz - 0.35, "B = {:.2f}".format(B))
    Z.cota_h(-d_acopio, 0, 0, 1.25, "≥ H/2 = {:.2f} (mín. 0.60)".format(d_acopio))
    Z.cota_v(0, 1.0, B, B + 0.35, "1.00")
    Z.cota_h(B, B + 1.0, 0, -0.30, "1.00")
    # llamadas
    Z.llamada(0.03, -1.0, -1.6, -0.9, "TABLONES 2\"x8\" (ENTIBADO CONTINUO)", alin="MR")
    Z.llamada(0.10, -1.40, -1.6, -1.6, "LARGUEROS 4\"x4\"", alin="MR")
    Z.llamada(0.60, -2.20, -1.6, -2.3, "PUNTALES Ø4\" / GATAS METÁLICAS @ 1.00 m", alin="MR")
    Z.llamada(xt + 0.06, 0.8, B + 0.7, 1.75, "ESCALERA DE ESCAPE CADA 7.50 m (SOBRESALE 1.00 m)")
    Z.llamada(B + 1.8, 0.6, B + 1.0, 2.2, "CACHACOS + MALLA NARANJA A 1.00 m DEL BORDE")
    Z.texto("MAQUINARIA PESADA A UNA DISTANCIA ≥ H DEL BORDE", 0.6, -Hz - 0.9, 1.5)

    # ======================================================================
    # DETALLE 03: ANDAMIO NORMATIVO  - 1:50 (k = 20)
    # ======================================================================
    _titulo_detalle(dp, 377, 382, "3", "ANDAMIO METÁLICO MULTIDIRECCIONAL NORMATIVO", "1:50",
                    110)
    A = Dib(msp, 432, 212, 20.0)
    Lb = 1.80
    A.linea(-2.2, 0, 4.3, 0, grosor=50)
    A.achurado([(-2.2, 0), (4.3, 0), (4.3, -0.25), (-2.2, -0.25)], 8, "EARTH", 0.02 * 20)
    # muro / estructura existente
    A.rect(2.15, 0, 2.55, 6.6, "SS-DETALLES", 7)
    A.achurado([(2.15, 0), (2.55, 0), (2.55, 6.6), (2.15, 6.6)], 8, "ANSI31", 0.6)
    A.texto("ESTRUCTURA", 2.35, 3.3, 1.5, rot=90)
    # tablones de reparto, placas base y husillos
    for x in (0.0, Lb):
        A.rect(x - 0.20, 0, x + 0.20, 0.05, "SS-DETALLES", 40)
        A.rect(x - 0.08, 0.05, x + 0.08, 0.08)
        A.linea(x, 0.08, x, 0.35, "SS-DETALLES", 7, grosor=50)
    # verticales
    for x in (0.0, Lb):
        A.rect(x - 0.025, 0.35, x + 0.025, 4.0 + 1.0)
    # horizontales y diagonales
    niveles = [0.50, 2.0, 4.0]
    for y in niveles:
        A.linea(0, y, Lb, y, grosor=35)
    A.linea(0, 0.50, Lb, 2.0)
    A.linea(Lb, 2.0, 0, 4.0)
    # plataformas completas
    for y in (2.0, 4.0):
        A.rect(-0.02, y, Lb + 0.02, y + 0.06, "SS-DETALLES", 40)
        A.solido_rect(-0.02, y, Lb + 0.02, y + 0.06, 40, "SS-ACHURADO")
    # barandas y rodapié en la plataforma de trabajo
    yp = 4.06
    A.linea(0, yp + 0.90, Lb, yp + 0.90, "SS-DETALLES", 1, grosor=50)
    A.linea(0, yp + 0.45, Lb, yp + 0.45, "SS-DETALLES", 1, grosor=35)
    A.rect(0, yp, Lb, yp + 0.15, "SS-DETALLES", 2)
    A.solido_rect(0, yp, Lb, yp + 0.15, 2, "SS-ACHURADO")
    # amarres a la estructura
    for y in (2.0, 4.0):
        A.linea(Lb, y + 0.2, 2.15, y + 0.2, "SS-DETALLES", 7, grosor=35)
        A.circulo(2.15, y + 0.2, 0.04)
    # línea de vida vertical independiente
    A.linea(-0.70, 0.30, -0.70, 6.30, "SS-SENALETICA", 6, grosor=35)
    A.poli([(-0.85, 6.30), (-0.55, 6.30), (-0.70, 6.55)], True, "SS-SENALETICA", 6)
    A.linea(-0.70, 6.55, 2.15, 6.55, "SS-SENALETICA", 6, "DASHED")
    A.rect(-0.80, 0.10, -0.60, 0.30, "SS-SENALETICA", 6)
    A.rect(-0.76, 4.55, -0.64, 4.75, "SS-SENALETICA", 6)
    # trabajador con arnés
    A.circulo(0.75, 5.75, 0.12)
    A.linea(0.75, 5.63, 0.75, 4.95)
    A.linea(0.75, 4.95, 0.60, 4.10)
    A.linea(0.75, 4.95, 0.90, 4.10)
    A.linea(0.75, 5.45, 1.05, 5.15)
    A.linea(0.75, 5.45, 0.50, 5.20)
    A.linea(0.70, 5.40, -0.64, 4.70, "SS-SENALETICA", 6, "DASHED")
    # cotas
    A.cota_h(0, Lb, 0, -0.55, "{:.2f}".format(Lb))
    A.cota_v(0, 2.0, 0, -1.25, "2.00")
    A.cota_v(2.0, 4.0, 0, -1.25, "2.00")
    A.cota_v(yp, yp + 0.90, Lb, Lb + 0.20, "0.90")
    A.cota_v(yp, yp + 0.45, 0, -0.30, "0.45")
    # llamadas
    A.llamada(-0.70, 6.0, -2.0, 7.1, "LÍNEA DE VIDA VERTICAL INDEPENDIENTE CON ARRESTADOR, "
              "ANCLAJE ≥ 2,268 kg", alin="ML")
    A.llamada(1.30, yp + 0.90, 3.0, 5.9, "BARANDA SUPERIOR A 0.90 m + INTERMEDIA A 0.45 m")
    A.llamada(1.40, yp + 0.08, 3.0, 4.6, "RODAPIÉ 0.15 m")
    A.llamada(0.90, 4.03, 3.0, 3.9, "PLATAFORMA COMPLETA CON SEGUROS")
    A.llamada(1.0, 1.25, 3.0, 2.6, "DIAGONALES / CRUCETAS EN CADA NIVEL")
    A.llamada(2.0, 2.2, 3.0, 1.9, "AMARRE A ESTRUCTURA @ 4.00 m")
    A.llamada(Lb, 0.25, 3.0, 0.9, "HUSILLO REGULABLE + PLACA BASE")
    A.llamada(Lb + 0.15, 0.03, 3.0, 0.35, "TABLÓN DE REPARTO")
    A.texto("ARNÉS ANSI Z359.11 CON DOBLE LÍNEA", 0.9, 6.2, 1.4, alin="BC")

    # ======================================================================
    # (D) CUADRO DE SEÑALÉTICA NTP 399.010-1
    # ======================================================================
    dp.texto("CUADRO DE SEÑALÉTICA DE SEGURIDAD EN OBRA (NTP 399.010-1)", 608, 383, 3.2,
             alin="BL")
    x0, x1, x2, x3, x4 = 606, 646, 680, 770, 826
    yt = 378
    dp.rect(x0, yt - 8, x4, yt)
    for x, t in ((x0, "SEÑAL"), (x1, "TIPO"), (x2, "SIGNIFICADO / EJEMPLOS"), (x3, "COLOR")):
        dp.texto(t, x + 2, yt - 4, 2.0, alin="ML")
    for x in (x1, x2, x3):
        dp.linea(x, yt - 8 - 6 * 28, x, yt)
    filas_s = [
        ("OBLIGACIÓN", "Uso obligatorio de casco, lentes, guantes, botines, arnés, protector "
         "auditivo y respirador.", "Azul / blanco", "obl"),
        ("PROHIBICIÓN", "Prohibido fumar, prohibido el ingreso a personal no autorizado, "
         "prohibido hacer fuego.", "Rojo / blanco", "pro"),
        ("ADVERTENCIA", "Caída de objetos, excavación profunda, riesgo eléctrico, maquinaria "
         "pesada en movimiento.", "Amarillo / negro", "adv"),
        ("EVACUACIÓN", "Salida de emergencia, ruta de evacuación, primeros auxilios, ducha / "
         "lavaojos.", "Verde / blanco", "eva"),
        ("ZONA SEGURA", "Zona segura en caso de sismo y punto de reunión.", "Verde / blanco",
         "seg"),
        ("CONTRA INCENDIO", "Extintor, gabinete contra incendio, alarma.", "Rojo / blanco", "inc"),
    ]
    for i, (tipo, sig, color, cod) in enumerate(filas_s):
        y2, y1 = yt - 8 - i * 28, yt - 8 - (i + 1) * 28
        dp.rect(x0, y1, x4, y2)
        cx, cy = (x0 + x1) / 2, (y1 + y2) / 2
        if cod == "obl":
            _senal_obligacion(dp, cx, cy, 10, "CASCO")
        elif cod == "pro":
            _senal_prohibicion(dp, cx, cy + 2, 9, "NO FUMAR")
        elif cod == "adv":
            _senal_advertencia(dp, cx, cy + 1, 20, "PELIGRO")
        elif cod == "eva":
            _senal_rect(dp, cx, cy, 32, 16, 3, "SALIDA ->")
        elif cod == "seg":
            _senal_rect(dp, cx, cy, 32, 18, 3, "ZONA SEGURA")
        else:
            _senal_rect(dp, cx, cy, 22, 18, 1, "EXTINTOR")
        dp.texto(tipo, x1 + 2, cy, 2.2, alin="ML")
        dp.mtexto(sig, x2 + 2, cy + 6, 1.9, 86)
        dp.texto(color, x3 + 2, cy, 2.0, alin="ML")
    dp.texto("Dimensión mínima de letreros: 0.60 x 0.40 m, en vinil reflectivo sobre plancha "
             "galvanizada o PVC.", 606, 191, 1.8, alin="BL")

    # ======================================================================
    # (E) CÓDIGO DE COLORES RR.SS. + DETALLE 04 ALMACÉN DE MATERIALES PELIGROSOS
    # ======================================================================
    dp.texto("CÓDIGO DE COLORES PARA RESIDUOS SÓLIDOS (NTP 900.058:2019)", 370, 176, 2.6,
             alin="BL")
    rrss = [(5, "AZUL", "Papel y cartón"), (7, "BLANCO", "Plástico"), (2, "AMARILLO", "Metales"),
            (9, "PLOMO", "Vidrio"), (32, "MARRÓN", "Residuos orgánicos"),
            (250, "NEGRO", "Residuos no aprovechables"), (1, "ROJO", "Residuos peligrosos")]
    for i, (col, nom, desc) in enumerate(rrss):
        y = 168 - i * 9.5
        dp.solido_rect(372, y - 3.5, 384, y + 3.5, col, "SS-SENALETICA")
        dp.rect(372, y - 3.5, 384, y + 3.5, "SS-SENALETICA", 7)
        dp.texto(nom, 388, y, 2.0, alin="ML")
        dp.texto(desc, 418, y, 2.0, alin="ML")
    dp.texto("RCD: acopio separado y disposición en escombrera autorizada.", 372, 98, 1.7,
             alin="BL")
    _titulo_detalle(dp, 377, 88, "4", "ALMACÉN DE MATERIALES PELIGROSOS", "1:50", 100)
    Q = Dib(msp, 395, 24, 20.0)
    Q.poli([(2.9, 2.6), (5.0, 2.6), (5.0, 0), (0, 0), (0, 2.6), (2.0, 2.6)], grosor=70)
    Q.rect(0.2, 0.2, 3.4, 1.0, "SS-DETALLES", 1)
    Q.rect(0.3, 0.3, 3.3, 0.9, "SS-DETALLES", 1, "DASHED")
    Q.texto("BANDEJA ANTIDERRAME (110 %)", 1.8, 0.6, 1.4, color=1)
    for i in range(4):
        Q.circulo(0.6 + i * 0.8, 1.7, 0.28)
    Q.texto("CILINDROS / GALONES", 1.8, 2.25, 1.3)
    Q.rect(3.8, 0.2, 4.8, 2.4, "SS-DETALLES", 7)
    Q.texto("ESTANTE", 4.3, 1.3, 1.2, rot=90)
    Q.texto("ACCESO", 2.45, 2.85, 1.2)
    Q.achurado([(5.25 + 0.25 * math.cos(a * math.pi / 8), 2.2 + 0.25 * math.sin(a * math.pi / 8))
                for a in range(16)], 1, None, 1, "SS-SENALETICA")
    Q.texto("E", 5.25, 2.2, 1.0, "SS-SENALETICA", "MC", 7)
    Q.texto("EXTINTOR PQS", 5.6, 2.2, 1.3, alin="ML")
    Q.texto("MSDS / KIT ANTIDERRAME", 5.6, 1.4, 1.3, alin="ML")
    Q.texto("NO FUMAR", 5.6, 0.6, 1.3, alin="ML", color=1)
    Q.cota_h(0, 5.0, 0, -0.45, "5.00")
    Q.cota_v(0, 2.6, 0, -0.45, "2.60")

    # ======================================================================
    # (F) MEMBRETE TÉCNICO
    # ======================================================================
    mx0, mx1 = 606, 826
    dp.rect(mx0, 15, mx1, 180, grosor=50)
    dp.solido_rect(mx0, 166, mx1, 180, 5, "SS-ACHURADO")
    dp.texto("PLAN DE SEGURIDAD Y SALUD EN EL TRABAJO - NORMA G.050", 716, 173, 3.0, color=7)
    filas_m = [("ENTIDAD:", cfg["entidad"]),
               ("OBRA:", cfg["proyecto"]),
               ("UBICACIÓN:", "{} - {} - {}".format(cfg["distrito"], cfg["provincia"],
                                                    cfg["region"])),
               ("LÁMINA:", "INSTALACIONES PROVISIONALES, SEÑALIZACIÓN TEMPORAL, CERCO, "
                           "CACHACOS Y DETALLES G.050"),
               ("PROFESIONAL SSOMA:", "{}  -  CIP N° {}".format(cfg["ssoma"], cfg["cip"])),
               ("RESIDENTE DE OBRA:", cfg["residente"]),
               ("MODALIDAD:", cfg["modalidad"])]
    y = 160
    for et, val in filas_m:
        dp.texto(et, mx0 + 3, y, 2.0, alin="ML")
        dp.mtexto(val, mx0 + 42, y + 1.4, 2.0, 175)
        y -= 12 if et in ("OBRA:", "LÁMINA:") else 9
        dp.linea(mx0, y + 4.5, mx1, y + 4.5, "SS-MARCO")
    rs = cfg["riesgos"]
    alto = ", ".join(t for k, t in (("altura", "Altura"), ("excavacion", "Excavación"),
                                    ("caliente", "Caliente"), ("confinado", "Esp. confinado"),
                                    ("izaje", "Izaje")) if rs.get(k)) or "Ninguno"
    datos = [("PERSONAL:", "{} obreros + {} staff técnico".format(cfg["n_obr"], cfg["n_ing"])),
             ("PLAZO:", "{} días calendario ({} meses)".format(cfg["dias"], cfg["meses"])),
             ("ALTO RIESGO:", alto),
             ("COSTO SST G.050:", soles(desag["total"]) if desag else "-")]
    for et, val in datos:
        dp.texto(et, mx0 + 3, y, 2.0, alin="ML")
        dp.texto(val, mx0 + 42, y, 2.0, alin="ML")
        y -= 9
        dp.linea(mx0, y + 4.5, mx1, y + 4.5, "SS-MARCO")
    dp.linea(mx0, 50, mx1, 50, "SS-MARCO")
    for x in (666, 726, 776):
        dp.linea(x, 15, x, 50, "SS-MARCO")
    dp.texto("ESCALA:", 610, 44, 1.8, alin="ML")
    dp.texto("INDICADA", 636, 30, 3.0)
    dp.texto("FECHA:", 670, 44, 1.8, alin="ML")
    dp.texto(cfg["fecha"], 696, 30, 3.0)
    dp.texto("FIRMA Y SELLO:", 730, 44, 1.8, alin="ML")
    dp.texto("LÁMINA:", 780, 44, 1.8, alin="ML")
    dp.texto("SS-01", 801, 28, 8.0, color=1)

    doc.set_modelspace_vport(height=620, center=(420.5, 297))
    doc.saveas(ruta)
    return ruta


# =============================================================================
# 11. INTERFAZ GRÁFICA (CustomTkinter - 1280 x 820 px)
# =============================================================================
COLOR_NIVEL_GUI = {"TRIVIAL": ("#27ae60", "white"), "TOLERABLE": ("#2ecc71", "black"),
                   "MODERADO": ("#f1c40f", "black"), "IMPORTANTE": ("#e67e22", "black"),
                   "INTOLERABLE": ("#c0392b", "white")}

if GUI_DISPONIBLE:

    class AppPSST(ctk.CTk):
        def __init__(self):
            super().__init__()
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")
            self.title(APP_CORTO + "  v" + APP_VERSION)
            self.geometry("1280x820")
            self.minsize(1100, 700)
            self.partidas = []
            self.filas_iperc = []
            self.desag = None
            self.vars = {}
            self._construir_cabecera()
            self._construir_tabs()
            self._construir_barra_inferior()
            self._estilo_tablas()

        # ------------------------------------------------------------------
        def _construir_cabecera(self):
            cab = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1F3864", "#14213d"))
            cab.pack(fill="x")
            ctk.CTkLabel(cab, text="GENERADOR PRO PSST - NORMA G.050  |  IPERC  |  EPI/EPC  |  "
                         "AutoCAD DXF", font=ctk.CTkFont(size=20, weight="bold"),
                         text_color="white").pack(side="left", padx=16, pady=(10, 0))
            self.sw_modo = ctk.CTkSwitch(cab, text="Modo oscuro", command=self._cambiar_modo,
                                         text_color="white")
            self.sw_modo.select()
            self.sw_modo.pack(side="right", padx=16)
            ctk.CTkLabel(cab, text="Ley N° 29783 · D.S. 005-2012-TR · D.S. 011-2019-TR · "
                         "R.M. 050-2013-TR · G.050  —  100 % offline",
                         font=ctk.CTkFont(size=12), text_color="#BDD7EE").pack(
                side="left", padx=10, pady=(10, 0))

        def _cambiar_modo(self):
            ctk.set_appearance_mode("dark" if self.sw_modo.get() else "light")
            self._estilo_tablas()

        def _estilo_tablas(self):
            oscuro = ctk.get_appearance_mode() == "Dark"
            st = ttk.Style(self)
            st.theme_use("clam")
            bg, fg, campo = ("#2b2b2b", "#eaeaea", "#2b2b2b") if oscuro else \
                ("#ffffff", "#111111", "#ffffff")
            st.configure("Treeview", background=bg, foreground=fg, fieldbackground=campo,
                         rowheight=26, font=("Segoe UI", 9), borderwidth=0)
            st.configure("Treeview.Heading", background="#1F3864", foreground="white",
                         font=("Segoe UI", 9, "bold"), relief="flat")
            st.map("Treeview", background=[("selected", "#2E75B6")],
                   foreground=[("selected", "white")])
            for arbol in (getattr(self, "tree_iperc", None), getattr(self, "tree_desag", None)):
                if arbol is not None:
                    arbol.tag_configure("seccion", background="#DDEBF7" if not oscuro
                                        else "#1F3864", foreground=fg)

        # ------------------------------------------------------------------
        def _construir_tabs(self):
            self.tabs = ctk.CTkTabview(self, command=self._al_cambiar_tab)
            self.tabs.pack(fill="both", expand=True, padx=10, pady=(6, 0))
            self.t1 = self.tabs.add("1. Configuración de la Obra y Presupuesto")
            self.t2 = self.tabs.add("2. Matriz IPERC por Partidas")
            self.t3 = self.tabs.add("3. Presupuesto Desagregado G.050 (EPI/EPC/EMO)")
            self._tab_config()
            self._tab_iperc()
            self._tab_desag()

        def _entrada(self, padre, fila, col, etiqueta, clave, valor="", ancho=360):
            ctk.CTkLabel(padre, text=etiqueta, anchor="w").grid(row=fila, column=col,
                                                                sticky="w", padx=(10, 6),
                                                                pady=4)
            var = tk.StringVar(value=str(valor))
            ctk.CTkEntry(padre, textvariable=var, width=ancho).grid(row=fila, column=col + 1,
                                                                    sticky="we", padx=(0, 14),
                                                                    pady=4)
            self.vars[clave] = var
            return var

        def _tab_config(self):
            marco = ctk.CTkScrollableFrame(self.t1)
            marco.pack(fill="both", expand=True)
            marco.grid_columnconfigure(1, weight=1)
            marco.grid_columnconfigure(3, weight=1)
            ctk.CTkLabel(marco, text="DATOS GENERALES DE LA OBRA",
                         font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(6, 4))
            self._entrada(marco, 1, 0, "Nombre del Proyecto / Obra:", "proyecto", "", 420)
            self._entrada(marco, 1, 2, "Entidad / Cliente:", "entidad", "", 300)
            self._entrada(marco, 2, 0, "Distrito:", "distrito", "")
            self._entrada(marco, 2, 2, "Provincia:", "provincia", "")
            self._entrada(marco, 3, 0, "Región:", "region", "")
            ctk.CTkLabel(marco, text="Modalidad de Ejecución:", anchor="w").grid(
                row=3, column=2, sticky="w", padx=(10, 6))
            self.var_modalidad = tk.StringVar(value="Contrata")
            ctk.CTkSegmentedButton(marco, values=["Contrata", "Administración Directa"],
                                   variable=self.var_modalidad).grid(row=3, column=3,
                                                                     sticky="w", pady=4)
            ctk.CTkLabel(marco, text="Plazo de Ejecución:", anchor="w").grid(
                row=4, column=0, sticky="w", padx=(10, 6))
            fpl = ctk.CTkFrame(marco, fg_color="transparent")
            fpl.grid(row=4, column=1, sticky="w")
            self.vars["plazo_valor"] = tk.StringVar(value="180")
            ctk.CTkEntry(fpl, textvariable=self.vars["plazo_valor"], width=90).pack(side="left")
            self.var_unidad = tk.StringVar(value="Días Calendario")
            ctk.CTkOptionMenu(fpl, values=["Días Calendario", "Meses"],
                              variable=self.var_unidad, width=160).pack(side="left", padx=8)
            self._entrada(marco, 4, 2, "Perímetro de obra / frentes (m):", "perimetro", 260,
                          120)
            self._entrada(marco, 5, 0, "N° promedio de obreros en obra:", "n_obr", 25, 120)
            self._entrada(marco, 5, 2, "N° de staff técnico (ingenieros):", "n_ing", 4, 120)
            self._entrada(marco, 6, 0, "Ingeniero Residente:", "residente", "")
            self._entrada(marco, 6, 2, "Especialista SSOMA:", "ssoma", "")
            self._entrada(marco, 7, 2, "N° CIP SSOMA:", "cip", "", 160)

            ctk.CTkLabel(marco, text="PRESENCIA DE TRABAJOS DE ALTO RIESGO",
                         font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=8, column=0, columnspan=4, sticky="w", padx=10, pady=(14, 4))
            fr = ctk.CTkFrame(marco)
            fr.grid(row=9, column=0, columnspan=4, sticky="we", padx=10)
            self.chk = {}
            for k, (clave, txt) in enumerate((("altura", "Altura > 1.80 m"),
                                              ("excavacion", "Excavaciones > 1.50 m"),
                                              ("caliente", "Trabajos en Caliente"),
                                              ("confinado", "Espacios Confinados"),
                                              ("izaje", "Izaje de Cargas"))):
                v = tk.BooleanVar(value=False)
                ctk.CTkCheckBox(fr, text=txt, variable=v).grid(row=0, column=k, padx=14,
                                                                pady=10)
                self.chk[clave] = v

            ctk.CTkLabel(marco, text="CARGA DEL PRESUPUESTO DE OBRA",
                         font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=10, column=0, columnspan=4, sticky="w", padx=10, pady=(14, 4))
            fb = ctk.CTkFrame(marco, fg_color="transparent")
            fb.grid(row=11, column=0, columnspan=4, sticky="we", padx=10)
            ctk.CTkButton(fb, text="1. Cargar Excel de Presupuesto de Obra (.XLSX)", height=42,
                          font=ctk.CTkFont(size=13, weight="bold"),
                          command=self.cargar_excel).pack(side="left", padx=(0, 10))
            ctk.CTkButton(fb, text="Cargar Obra Completa de Ejemplo (Colegio / Pavimentación "
                          "con 18 Partidas)", height=42, fg_color="#7030A0",
                          hover_color="#5a2580", font=ctk.CTkFont(size=13, weight="bold"),
                          command=self.cargar_ejemplo).pack(side="left")
            self.txt_log = ctk.CTkTextbox(marco, height=210, font=ctk.CTkFont(family="Consolas",
                                                                              size=12))
            self.txt_log.grid(row=12, column=0, columnspan=4, sticky="we", padx=10, pady=10)
            self._log("Bienvenido. Complete los datos de la obra y cargue su presupuesto en "
                      "Excel (.xlsx) o pruebe con la obra de ejemplo.\nEl programa detecta la "
                      "columna de descripción, unidad y metrado, y asigna a cada partida su "
                      "familia de riesgo IPERC (18 familias G.050).")

        def _log(self, texto):
            self.txt_log.insert("end", texto + "\n")
            self.txt_log.see("end")

        def _arbol(self, padre, columnas, anchos):
            fr = tk.Frame(padre)
            fr.pack(fill="both", expand=True, padx=4, pady=4)
            arbol = ttk.Treeview(fr, columns=columnas, show="headings", selectmode="browse")
            for c, w in zip(columnas, anchos):
                arbol.heading(c, text=c)
                arbol.column(c, width=w, minwidth=40, anchor="w" if w > 120 else "center",
                             stretch=w > 120)
            ys = ttk.Scrollbar(fr, orient="vertical", command=arbol.yview)
            xs = ttk.Scrollbar(fr, orient="horizontal", command=arbol.xview)
            arbol.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
            arbol.grid(row=0, column=0, sticky="nsew")
            ys.grid(row=0, column=1, sticky="ns")
            xs.grid(row=1, column=0, sticky="ew")
            fr.grid_rowconfigure(0, weight=1)
            fr.grid_columnconfigure(0, weight=1)
            return arbol

        def _tab_iperc(self):
            sup = ctk.CTkFrame(self.t2, fg_color="transparent")
            sup.pack(fill="x")
            self.lbl_iperc = ctk.CTkLabel(sup, text="Cargue un presupuesto para generar la "
                                          "matriz IPERC.", font=ctk.CTkFont(size=13,
                                                                           weight="bold"))
            self.lbl_iperc.pack(side="left", padx=8, pady=4)
            ctk.CTkLabel(sup, text="Doble clic en una partida: ver detalle y reasignar familia",
                         text_color="gray").pack(side="right", padx=8)
            cols = ("N°", "Ítem", "Partida", "Familia de riesgo", "Peligro", "Riesgo", "IP", "S",
                    "P x S", "Nivel inicial", "Medidas de control (ingeniería / adm.)",
                    "P x S res.", "Nivel residual", "PETAR / ATS")
            anchos = (40, 60, 280, 200, 280, 240, 40, 40, 55, 100, 360, 70, 100, 200)
            self.tree_iperc = self._arbol(self.t2, cols, anchos)
            for niv, (bg, fg) in COLOR_NIVEL_GUI.items():
                self.tree_iperc.tag_configure(niv, background=bg, foreground=fg)
            self.tree_iperc.bind("<Double-1>", self._detalle_partida)

        def _tab_desag(self):
            sup = ctk.CTkFrame(self.t3, fg_color="transparent")
            sup.pack(fill="x")
            self.lbl_total = ctk.CTkLabel(sup, text="COSTO DIRECTO TOTAL DE SEGURIDAD: S/ 0.00",
                                          font=ctk.CTkFont(size=18, weight="bold"),
                                          text_color="#e74c3c")
            self.lbl_total.pack(side="left", padx=8, pady=6)
            ctk.CTkButton(sup, text="Recalcular con los datos actuales", width=240,
                          command=self.recalcular).pack(side="right", padx=8)
            cols = ("Ítem", "Descripción", "Und.", "Base", "Vida útil (meses)", "Factor repos.",
                    "Cantidad", "P.U. (S/)", "Parcial (S/)")
            anchos = (80, 560, 60, 70, 110, 100, 80, 90, 120)
            self.tree_desag = self._arbol(self.t3, cols, anchos)

        def _construir_barra_inferior(self):
            barra = ctk.CTkFrame(self, corner_radius=0)
            barra.pack(fill="x", side="bottom")
            fuente = ctk.CTkFont(size=12, weight="bold")
            ctk.CTkButton(barra, text="1. Exportar Libro Excel Completo: Matriz IPERC + "
                          "Presupuesto EPI/EPC +\nFormatos ATS/PETAR (.XLSX)", height=54,
                          fg_color="#1e8449", hover_color="#145a32", font=fuente,
                          command=self.exportar_excel).pack(side="left", expand=True,
                                                           fill="x", padx=6, pady=8)
            ctk.CTkButton(barra, text="2. Generar Plan de Seguridad y Salud en Obra (PSST - "
                          "Norma G.050)\n+ PETS en Word (.DOCX)", height=54,
                          fg_color="#1f618d", hover_color="#154360", font=fuente,
                          command=self.exportar_word).pack(side="left", expand=True, fill="x",
                                                          padx=6, pady=8)
            ctk.CTkButton(barra, text="3. Dibujar Lámina de Señalización Temporal, Cerco, "
                          "Cachacos\ny Detalles G.050 en AutoCAD (.DXF)", height=54,
                          fg_color="#ca6f1e", hover_color="#935116", font=fuente,
                          command=self.exportar_dxf).pack(side="left", expand=True, fill="x",
                                                         padx=6, pady=8)
            self.lbl_estado = ctk.CTkLabel(self, text="Listo.", anchor="w")
            self.lbl_estado.pack(fill="x", side="bottom", padx=12)

        # ------------------------------------------------------------------
        def leer_config(self):
            def entero(clave, nombre):
                try:
                    v = int(float(self.vars[clave].get().replace(",", ".")))
                    if v <= 0:
                        raise ValueError
                    return v
                except ValueError:
                    raise ValueError("Ingrese un número válido mayor que cero en: " + nombre)

            cfg = {k: self.vars[k].get().strip() for k in ("proyecto", "entidad", "distrito",
                                                            "provincia", "region", "residente",
                                                            "ssoma", "cip")}
            cfg["proyecto"] = cfg["proyecto"] or "OBRA SIN NOMBRE"
            cfg["entidad"] = cfg["entidad"] or "ENTIDAD / CLIENTE"
            cfg["residente"] = cfg["residente"] or "Ing. Residente de Obra"
            cfg["ssoma"] = cfg["ssoma"] or "Ing. Especialista SSOMA"
            cfg["cip"] = cfg["cip"] or "------"
            cfg["modalidad"] = self.var_modalidad.get()
            cfg["plazo_valor"] = entero("plazo_valor", "Plazo de ejecución")
            cfg["plazo_unidad"] = self.var_unidad.get()
            cfg["n_obr"] = entero("n_obr", "N° de obreros")
            cfg["n_ing"] = entero("n_ing", "N° de staff técnico")
            cfg["perimetro"] = entero("perimetro", "Perímetro de obra")
            cfg["riesgos"] = {k: bool(v.get()) for k, v in self.chk.items()}
            return completar_config(cfg)

        def _poner_config(self, cfg):
            for k in ("proyecto", "entidad", "distrito", "provincia", "region", "residente",
                      "ssoma", "cip", "n_obr", "n_ing", "plazo_valor", "perimetro"):
                self.vars[k].set(str(cfg[k]))
            self.var_modalidad.set(cfg["modalidad"])
            self.var_unidad.set(cfg["plazo_unidad"])

        def _marcar_riesgos(self):
            for r in riesgos_detectados(self.filas_iperc):
                self.chk[r].set(True)

        def cargar_excel(self):
            ruta = filedialog.askopenfilename(title="Seleccione el presupuesto de obra",
                                              filetypes=[("Excel", "*.xlsx *.xlsm")])
            if not ruta:
                return
            try:
                self.partidas = leer_presupuesto_excel(ruta)
            except Exception as e:
                messagebox.showerror("Error al leer el presupuesto", str(e))
                return
            self._log("\n>> Presupuesto cargado: {}  ({} partidas)".format(
                os.path.basename(ruta), len(self.partidas)))
            self._procesar_partidas()

        def cargar_ejemplo(self):
            self._poner_config(configuracion_ejemplo())
            self.partidas = partidas_ejemplo()
            self._log("\n>> Obra de ejemplo cargada (Colegio + Pavimentación, 18 partidas).")
            self._procesar_partidas()

        def _procesar_partidas(self):
            self.filas_iperc = construir_iperc(self.partidas)
            self._marcar_riesgos()
            for f in self.filas_iperc:
                self._log("   {:<7} {:<55.55} -> {} [{} → {}]".format(
                    f["item"], f["partida"], f["familia_cod"], f["nivel"], f["nivel_res"]))
            self._log("   Trabajos de alto riesgo detectados: {}".format(
                ", ".join(sorted(riesgos_detectados(self.filas_iperc))) or "ninguno"))
            self._llenar_iperc()
            self.recalcular()
            self.tabs.set("2. Matriz IPERC por Partidas")
            self._estado("Matriz IPERC generada con {} partidas.".format(len(self.filas_iperc)))

        def _llenar_iperc(self):
            t = self.tree_iperc
            t.delete(*t.get_children())
            for k, f in enumerate(self.filas_iperc, 1):
                t.insert("", "end", iid=str(k - 1), tags=(f["nivel"],), values=(
                    k, f["item"], f["partida"], f["familia"], f["peligro"], f["riesgo"], f["IP"],
                    f["S"], f["PxS"], f["nivel"], f["ingenieria"] + " | " + f["administrativo"],
                    f["rPxS"], f["nivel_res"], f["petar"]))
            altos = sum(1 for f in self.filas_iperc if f["PxS"] > 16)
            self.lbl_iperc.configure(text="{} partidas evaluadas  |  {} con riesgo inicial "
                                     "IMPORTANTE / INTOLERABLE  |  100 % con riesgo residual "
                                     "TOLERABLE / MODERADO".format(len(self.filas_iperc), altos))

        def _detalle_partida(self, _evento=None):
            sel = self.tree_iperc.selection()
            if not sel:
                return
            idx = int(sel[0])
            f = self.filas_iperc[idx]
            win = ctk.CTkToplevel(self)
            win.title("Detalle IPERC - " + f["item"])
            win.geometry("860x640")
            win.transient(self)
            txt = ctk.CTkTextbox(win, wrap="word", font=ctk.CTkFont(size=12))
            txt.pack(fill="both", expand=True, padx=10, pady=10)
            txt.insert("end", (
                "PARTIDA: {item}  {partida}\nFAMILIA: {familia}\n\nPELIGRO: {peligro}\n"
                "RIESGO: {riesgo}\nCONSECUENCIA: {consecuencia}\n\nEVALUACIÓN INICIAL: "
                "A={A} B={B} C={C} D={D}  IP={IP}  S={S}  ->  P x S = {PxS} ({nivel})\n\n"
                "MEDIDAS DE CONTROL\n  1. Eliminación: {eliminacion}\n  2. Sustitución: "
                "{sustitucion}\n  3. Ingeniería: {ingenieria}\n  4. Administrativo: "
                "{administrativo}\n  5. EPP: {epp}\n\nEVALUACIÓN RESIDUAL: A={rA} B={rB} "
                "C={rC} D={rD}  IP={rIP}  S={rS}  ->  P x S = {rPxS} ({nivel_res})\n\n"
                "PETAR / ATS: {petar}\nPETS: {pets}").format(**f))
            txt.configure(state="disabled")
            fr = ctk.CTkFrame(win)
            fr.pack(fill="x", padx=10, pady=(0, 10))
            ctk.CTkLabel(fr, text="Reasignar familia de riesgo:").pack(side="left", padx=8)
            nombres = ["{} - {}".format(x["codigo"], x["nombre"]) for x in FAMILIAS_IPERC]
            var = tk.StringVar(value="{} - {}".format(f["familia_cod"], f["familia"]))
            ctk.CTkOptionMenu(fr, values=nombres, variable=var, width=520).pack(side="left",
                                                                                padx=6, pady=8)

            def aplicar():
                cod = var.get().split(" - ")[0]
                self.filas_iperc[idx] = fila_iperc(self.partidas[idx], FAMILIAS_POR_CODIGO[cod])
                self._marcar_riesgos()
                self._llenar_iperc()
                self.recalcular()
                win.destroy()

            ctk.CTkButton(fr, text="Aplicar", width=100, command=aplicar).pack(side="left",
                                                                                padx=6)

        def recalcular(self):
            try:
                cfg = self.leer_config()
            except ValueError as e:
                messagebox.showwarning("Datos incompletos", str(e))
                return
            self.desag = calcular_desagregado(cfg)
            t = self.tree_desag
            t.delete(*t.get_children())
            for s in self.desag["secciones"]:
                t.insert("", "end", tags=("seccion",), values=(
                    s["codigo"], s["titulo"], "", "", "", "", "", "", "{:,.2f}".format(
                        s["subtotal"])))
                for it in s["items"]:
                    t.insert("", "end", values=(
                        it["codigo"], it["desc"], it["und"], it["base"], it["vida"] or "-",
                        it["factor"], it["cant"], "{:,.2f}".format(it["pu"]),
                        "{:,.2f}".format(it["parcial"])))
            t.insert("", "end", tags=("seccion",), values=(
                "01", "COSTO DIRECTO TOTAL DE SEGURIDAD Y SALUD EN EL TRABAJO", "", "", "", "",
                "", "", "{:,.2f}".format(self.desag["total"])))
            self._estilo_tablas()
            self.lbl_total.configure(text="COSTO DIRECTO TOTAL DE SEGURIDAD: {}   ({} obreros + "
                                     "{} staff, {} meses)".format(
                                         soles(self.desag["total"]), cfg["n_obr"], cfg["n_ing"],
                                         cfg["meses"]))

        def _al_cambiar_tab(self):
            if self.tabs.get().startswith("3"):
                self.recalcular()

        def _estado(self, texto):
            self.lbl_estado.configure(text=texto)
            self.update_idletasks()

        # ------------------------------------------------------------------
        def _preparar(self):
            if not self.filas_iperc:
                messagebox.showinfo("Sin presupuesto", "Primero cargue un presupuesto de obra "
                                    "o la obra de ejemplo.")
                return None
            try:
                cfg = self.leer_config()
            except ValueError as e:
                messagebox.showwarning("Datos incompletos", str(e))
                return None
            desag = calcular_desagregado(cfg)
            crono = generar_cronograma(cfg, self.filas_iperc)
            return cfg, desag, crono

        def _nombre_base(self, cfg):
            base = re.sub(r"[^A-Za-z0-9]+", "_", normalizar(cfg["proyecto"]).upper())[:40]
            return base.strip("_") or "OBRA"

        def _ejecutar(self, tipo):
            datos = self._preparar()
            if not datos:
                return
            cfg, desag, crono = datos
            ext = {"xlsx": ".xlsx", "docx": ".docx", "dxf": ".dxf"}[tipo]
            pref = {"xlsx": "IPERC_PRESUPUESTO_SSOMA_", "docx": "PSST_G050_",
                    "dxf": "LAMINA_SS-01_"}[tipo]
            ruta = filedialog.asksaveasfilename(defaultextension=ext,
                                                initialfile=pref + self._nombre_base(cfg) + ext,
                                                filetypes=[(ext.upper()[1:], "*" + ext)])
            if not ruta:
                return
            self._estado("Generando {} ...".format(os.path.basename(ruta)))
            try:
                if tipo == "xlsx":
                    exportar_excel(ruta, cfg, self.filas_iperc, desag, crono)
                elif tipo == "docx":
                    exportar_word(ruta, cfg, self.filas_iperc, desag, crono)
                else:
                    exportar_dxf(ruta, cfg, desag)
            except PermissionError:
                messagebox.showerror("Archivo en uso", "No se pudo guardar. Cierre el archivo "
                                     "si está abierto en otro programa e intente de nuevo.")
                self._estado("Error: archivo en uso.")
                return
            except Exception as e:
                messagebox.showerror("Error al exportar", str(e))
                self._estado("Error al exportar.")
                return
            self._estado("Archivo generado: " + ruta)
            self._log(">> Generado: " + ruta)
            abrir_archivo(ruta)

        def exportar_excel(self):
            self._ejecutar("xlsx")

        def exportar_word(self):
            self._ejecutar("docx")

        def exportar_dxf(self):
            self._ejecutar("dxf")


# =============================================================================
# 12. PUNTO DE ENTRADA
# =============================================================================
def generar_demo(carpeta):
    """Genera los 3 entregables de la obra de ejemplo sin interfaz gráfica."""
    os.makedirs(carpeta, exist_ok=True)
    cfg = completar_config(configuracion_ejemplo())
    filas = construir_iperc(partidas_ejemplo())
    desag = calcular_desagregado(cfg)
    crono = generar_cronograma(cfg, filas)
    rutas = [
        exportar_excel(os.path.join(carpeta, "IPERC_PRESUPUESTO_SSOMA_EJEMPLO.xlsx"), cfg, filas,
                       desag, crono),
        exportar_word(os.path.join(carpeta, "PSST_G050_EJEMPLO.docx"), cfg, filas, desag, crono),
        exportar_dxf(os.path.join(carpeta, "LAMINA_SS-01_EJEMPLO.dxf"), cfg, desag),
    ]
    print("Costo directo total de seguridad: " + soles(desag["total"]))
    for r in rutas:
        print("Generado: " + r)
    return rutas


def main():
    if len(sys.argv) >= 2 and sys.argv[1] == "--demo":
        generar_demo(sys.argv[2] if len(sys.argv) > 2 else "salida_demo")
        return
    if not GUI_DISPONIBLE:
        print("No se encontró tkinter/customtkinter. Instale: pip install customtkinter\n"
              "o use:  python app_psst_g050_iperc_pro.py --demo CARPETA")
        sys.exit(1)
    app = AppPSST()
    app.mainloop()


if __name__ == "__main__":
    main()

# =============================================================================
# INSTALACIÓN Y EMPAQUETADO
# -----------------------------------------------------------------------------
#   pip install customtkinter openpyxl pandas python-docx ezdxf
#
#   pyinstaller --noconfirm --onefile --windowed --name "Generador_PSST_G050" ^
#       --collect-all customtkinter app_psst_g050_iperc_pro.py
# =============================================================================
