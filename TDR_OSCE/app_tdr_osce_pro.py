# -*- coding: utf-8 -*-
"""
===============================================================================
 GENERADOR PRO DE TÉRMINOS DE REFERENCIA (TDR) Y REQUISITOS TÉCNICOS MÍNIMOS
 PARA LICITACIONES (OSCE / OECE - SEACE) - PERÚ
===============================================================================
 Aplicación de escritorio 100 % offline.

 Librerías:  customtkinter, python-docx, openpyxl (+ json, os, math, datetime)

 Uso:
     python app_tdr_osce_pro.py            -> abre la interfaz gráfica
     python app_tdr_osce_pro.py --prueba   -> genera un TDR y un Excel de
                                              ejemplo sin abrir la interfaz

 IMPORTANTE (normativa):
   * Desde el 22-04-2025 rige la Ley N° 32069, Ley General de Contrataciones
     Públicas, y su Reglamento (D.S. N° 009-2025-EF). El OSCE pasó a ser el
     OECE. Para procesos convocados antes se mantiene el TUO de la Ley
     N° 30225 y el D.S. N° 344-2018-EF. Ambos marcos están disponibles.
   * Los umbrales en soles, la UIT y los porcentajes de adelantos son
     PARÁMETROS EDITABLES (pestaña 3). Verifíquelos cada año con la Ley de
     Presupuesto del Sector Público y las Bases Estándar vigentes.
===============================================================================
"""

import json
import math
import os
import subprocess
import sys
from datetime import datetime

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

APP_TITULO = ("Generador Pro de Términos de Referencia (TDR) y Requisitos "
              "Técnicos Mínimos para Licitaciones (OSCE / SEACE)")
APP_VERSION = "1.0.0"

# =============================================================================
# 1. BASE DE DATOS NORMATIVA INTERNA
# =============================================================================

OBJETOS = [
    "Consultoría de Expediente Técnico",
    "Ejecución de Obra",
    "Supervisión de Obra",
    "Servicio General",
    "Consultoría de Obra",
]

ESPECIALIDADES = ["Vial", "Edificaciones", "Saneamiento", "Saneamiento/Riego"]

SISTEMAS = ["Suma Alzada", "Precios Unitarios", "Mixto"]

# Tipo interno de cada objeto de contratación
TIPO_OBJETO = {
    "Consultoría de Expediente Técnico": "expediente",
    "Ejecución de Obra": "obra",
    "Supervisión de Obra": "supervision",
    "Servicio General": "servicio",
    "Consultoría de Obra": "expediente",
}

# Categoría para umbrales y fórmula de penalidad
CATEGORIA_UMBRAL = {
    "expediente": "consultoria",
    "supervision": "consultoria",
    "obra": "obra",
    "servicio": "servicio",
}

MARCOS_LEGALES = {
    "Ley N° 32069 y D.S. N° 009-2025-EF (vigente)": {
        "organismo": "OECE",
        "organismo_largo": ("Organismo Especializado para las Contrataciones "
                            "Públicas Eficientes (OECE)"),
        "ley": "Ley N° 32069, Ley General de Contrataciones Públicas",
        "reglamento": ("Reglamento de la Ley N° 32069, aprobado por Decreto "
                       "Supremo N° 009-2025-EF"),
        "corto": "la Ley N° 32069 y su Reglamento",
    },
    "TUO Ley N° 30225 y D.S. N° 344-2018-EF (procesos anteriores)": {
        "organismo": "OSCE",
        "organismo_largo": ("Organismo Supervisor de las Contrataciones del "
                            "Estado (OSCE)"),
        "ley": ("Texto Único Ordenado de la Ley N° 30225, Ley de Contrataciones "
                "del Estado, aprobado por D.S. N° 082-2019-EF"),
        "reglamento": ("Reglamento de la Ley N° 30225, aprobado por Decreto "
                       "Supremo N° 344-2018-EF y sus modificatorias"),
        "corto": "la Ley N° 30225 y su Reglamento",
    },
}
MARCO_DEFAULT = list(MARCOS_LEGALES.keys())[0]

# --- Parámetros editables (valores referenciales, verificar cada año) ---------
UIT_DEFAULT = 5500.00
PARAMETROS_DEFAULT = {
    "uit": UIT_DEFAULT,
    "menor_cuantia_uit": 8.0,              # Contrataciones <= 8 UIT
    "limite_as_obra": 2_300_000.00,        # Hasta este monto: AS (obras)
    "limite_as_consultoria": 400_000.00,   # Hasta este monto: AS (consultorías)
    "limite_as_servicio": 400_000.00,      # Hasta este monto: AS (servicios)
    "garantia_fiel_pct": 10.0,
    "adelanto_directo_obra_pct": 10.0,
    "adelanto_materiales_obra_pct": 20.0,
    "adelanto_directo_otros_pct": 30.0,
    "revision_entidad_dias": 10,
    "subsanacion_dias": 5,
    "factor_exp_postor": 1.0,              # Veces el valor referencial
    "subcontratacion_max_pct": 40.0,
}

CAT_MENOR = "Menor Cuantía"
CAT_AS = "Adjudicación Simplificada"
CAT_MAYOR = "Concurso Público / Licitación Pública"
CATEGORIAS = [CAT_MENOR, CAT_AS, CAT_MAYOR]

# Criterios de personal clave por categoría (líder / especialistas)
CRITERIOS_PERSONAL = {
    CAT_MENOR: {"lider": {"gen": 2, "esp": 1, "cap": 40},
                "especialista": {"gen": 2, "esp": 1, "cap": 40}},
    CAT_AS: {"lider": {"gen": 3, "esp": 2, "cap": 60},
             "especialista": {"gen": 2, "esp": 1, "cap": 40}},
    CAT_MAYOR: {"lider": {"gen": 5, "esp": 3, "cap": 80},
                "especialista": {"gen": 4, "esp": 2, "cap": 60}},
}

# Definición de "obras / servicios similares" por especialidad
SIMILARES = {
    "Vial": ("construcción, creación, mejoramiento, rehabilitación, "
             "ampliación y/o reconstrucción de carreteras, caminos vecinales, "
             "trochas carrozables, vías urbanas, pistas y veredas, y/o puentes "
             "vehiculares"),
    "Edificaciones": ("construcción, creación, mejoramiento, ampliación, "
                      "rehabilitación y/o reconstrucción de edificaciones en "
                      "general (instituciones educativas, establecimientos de "
                      "salud, locales institucionales, edificios "
                      "administrativos, mercados, complejos deportivos y/o "
                      "edificaciones multifamiliares)"),
    "Saneamiento": ("construcción, creación, mejoramiento, ampliación, "
                    "rehabilitación y/o renovación de sistemas de agua "
                    "potable, alcantarillado sanitario, plantas de "
                    "tratamiento de agua potable o de aguas residuales, "
                    "reservorios, líneas de conducción, aducción e "
                    "impulsión y/o unidades básicas de saneamiento"),
    "Saneamiento/Riego": ("construcción, creación, mejoramiento, ampliación "
                          "y/o rehabilitación de sistemas de riego, canales "
                          "de riego, bocatomas, reservorios, presas, "
                          "sistemas de riego tecnificado, obras de drenaje, "
                          "defensas ribereñas y/o sistemas de agua potable y "
                          "alcantarillado"),
}

# Especialista propio de cada especialidad
ESPECIALISTA_ESPECIALIDAD = {
    "Vial": ("Especialista en Suelos y Pavimentos",
             "Ingeniero Civil", "CIP",
             "Mecánica de suelos, geotecnia y diseño/control de pavimentos"),
    "Edificaciones": ("Especialista en Arquitectura",
                      "Arquitecto", "CAP",
                      "Diseño arquitectónico, accesibilidad y normativa A.010 / A.120"),
    "Saneamiento": ("Especialista en Hidráulica y Sanitaria",
                    "Ingeniero Sanitario o Ingeniero Civil", "CIP",
                    "Diseño hidráulico de redes de agua y alcantarillado (WaterCAD/SewerCAD)"),
    "Saneamiento/Riego": ("Especialista en Hidráulica y Obras de Riego",
                          "Ingeniero Agrícola o Ingeniero Civil", "CIP",
                          "Hidrología, hidráulica y diseño de obras de riego"),
}

FORMACION_LIDER = {
    "Vial": ("Ingeniero Civil", "CIP"),
    "Edificaciones": ("Ingeniero Civil o Arquitecto", "CIP / CAP"),
    "Saneamiento": ("Ingeniero Civil o Ingeniero Sanitario", "CIP"),
    "Saneamiento/Riego": ("Ingeniero Civil o Ingeniero Agrícola", "CIP"),
}

# --- Normas técnicas aplicables ---------------------------------------------
NORMAS_GENERALES = [
    ("D.S. N° 011-2006-VIVIENDA", "Reglamento Nacional de Edificaciones (RNE) y sus modificatorias"),
    ("Norma G.050", "Seguridad durante la Construcción (RNE)"),
    ("Ley N° 29783", "Ley de Seguridad y Salud en el Trabajo y su Reglamento D.S. N° 005-2012-TR"),
    ("D.S. N° 011-2019-TR", "Reglamento de Seguridad y Salud en el Trabajo para el Sector Construcción"),
    ("Ley N° 27446", "Ley del Sistema Nacional de Evaluación del Impacto Ambiental (SEIA) y su Reglamento D.S. N° 019-2009-MINAM"),
    ("D. Leg. N° 1278", "Ley de Gestión Integral de Residuos Sólidos y su Reglamento D.S. N° 014-2017-MINAM"),
    ("D.S. N° 003-2013-VIVIENDA", "Reglamento para la Gestión y Manejo de los Residuos de las Actividades de la Construcción y Demolición"),
    ("D. Leg. N° 1252", "Sistema Nacional de Programación Multianual y Gestión de Inversiones (Invierte.pe) y su Reglamento"),
    ("Ley N° 28296", "Ley General del Patrimonio Cultural de la Nación (Certificado de Inexistencia de Restos Arqueológicos - CIRA)"),
    ("Ley N° 29664", "Ley del Sistema Nacional de Gestión del Riesgo de Desastres (SINAGERD)"),
    ("NTP vigentes", "Normas Técnicas Peruanas aplicables a materiales y ensayos (INACAL)"),
]

NORMAS_ESPECIALIDAD = {
    "Vial": [
        ("Manual DG-2018", "Manual de Carreteras: Diseño Geométrico (MTC)"),
        ("Manual EG-2013", "Manual de Carreteras: Especificaciones Técnicas Generales para Construcción (MTC)"),
        ("Manual SGGP-2014", "Manual de Carreteras: Suelos, Geología, Geotecnia y Pavimentos (MTC)"),
        ("Manual HHD-2011", "Manual de Hidrología, Hidráulica y Drenaje (MTC)"),
        ("Manual DCT-2016", "Manual de Dispositivos de Control del Tránsito Automotor para Calles y Carreteras (MTC)"),
        ("Manual de Puentes 2018", "Manual de Puentes (MTC)"),
        ("Manual de Ensayo de Materiales", "Manual de Ensayo de Materiales (MTC, 2016)"),
        ("Manual de Mantenimiento 2018", "Manual de Carreteras: Mantenimiento o Conservación Vial (MTC)"),
        ("Norma CE.010", "Pavimentos Urbanos (RNE)"),
        ("Norma GH.020", "Componentes de Diseño Urbano (RNE)"),
    ],
    "Edificaciones": [
        ("Norma A.010", "Condiciones Generales de Diseño (RNE)"),
        ("Norma A.120", "Accesibilidad Universal en Edificaciones (RNE)"),
        ("Norma A.130", "Requisitos de Seguridad (RNE)"),
        ("Norma E.020", "Cargas (RNE)"),
        ("Norma E.030", "Diseño Sismorresistente (RNE)"),
        ("Norma E.050", "Suelos y Cimentaciones (RNE)"),
        ("Norma E.060", "Concreto Armado (RNE)"),
        ("Norma E.070", "Albañilería (RNE)"),
        ("Norma E.090", "Estructuras Metálicas (RNE)"),
        ("Norma IS.010", "Instalaciones Sanitarias para Edificaciones (RNE)"),
        ("Norma EM.010", "Instalaciones Eléctricas Interiores (RNE) y Código Nacional de Electricidad"),
        ("Ley N° 29090", "Ley de Regulación de Habilitaciones Urbanas y de Edificaciones"),
    ],
    "Saneamiento": [
        ("Norma OS.010", "Captación y Conducción de Agua para Consumo Humano (RNE)"),
        ("Norma OS.020", "Plantas de Tratamiento de Agua para Consumo Humano (RNE)"),
        ("Norma OS.030", "Almacenamiento de Agua para Consumo Humano (RNE)"),
        ("Norma OS.050", "Redes de Distribución de Agua para Consumo Humano (RNE)"),
        ("Norma OS.070", "Redes de Aguas Residuales (RNE)"),
        ("Norma OS.090", "Plantas de Tratamiento de Aguas Residuales (RNE)"),
        ("Norma OS.100", "Consideraciones Básicas de Diseño de Infraestructura Sanitaria (RNE)"),
        ("R.M. N° 192-2018-VIVIENDA", "Norma Técnica de Diseño: Opciones Tecnológicas para Sistemas de Saneamiento en el Ámbito Rural"),
        ("NTP ISO 1452", "Sistemas de tuberías de PVC-U para abastecimiento de agua"),
        ("NTP ISO 4435", "Tuberías de PVC-U para sistemas de drenaje y alcantarillado"),
        ("D.S. N° 031-2010-SA", "Reglamento de la Calidad del Agua para Consumo Humano"),
    ],
    "Saneamiento/Riego": [
        ("Ley N° 29338", "Ley de Recursos Hídricos y su Reglamento D.S. N° 001-2010-AG"),
        ("Manual ANA 2010", "Manual: Criterios de Diseños de Obras Hidráulicas para la Formulación de Proyectos Hidráulicos Multisectoriales y de Afianzamiento Hídrico (ANA)"),
        ("Manual HHD-2011", "Manual de Hidrología, Hidráulica y Drenaje (MTC)"),
        ("Norma E.060", "Concreto Armado (RNE)"),
        ("Norma E.050", "Suelos y Cimentaciones (RNE)"),
        ("Norma OS.010", "Captación y Conducción de Agua para Consumo Humano (RNE)"),
        ("Norma OS.030", "Almacenamiento de Agua para Consumo Humano (RNE)"),
        ("R.J. ANA vigente", "Procedimientos de acreditación de disponibilidad hídrica y autorización de ejecución de obras de aprovechamiento hídrico"),
        ("NTP ISO 1452", "Sistemas de tuberías de PVC-U para abastecimiento de agua"),
        ("Guías MIDAGRI", "Guías y lineamientos técnicos del Ministerio de Desarrollo Agrario y Riego para proyectos de riego"),
    ],
}

SOFTWARE_ESPECIALIDAD = {
    "Vial": "AutoCAD Civil 3D, S10 Costos y Presupuestos, MS Project, HEC-RAS, HEC-HMS, CSiBridge/SAP2000",
    "Edificaciones": "AutoCAD/Revit (BIM), ETABS, SAP2000, SAFE, S10 Costos y Presupuestos, MS Project, DIALux",
    "Saneamiento": "AutoCAD Civil 3D, WaterCAD/WaterGEMS, SewerCAD/SewerGEMS, EPANET, S10 Costos y Presupuestos, MS Project",
    "Saneamiento/Riego": "AutoCAD Civil 3D, HEC-RAS, HEC-HMS, CROPWAT, SAP2000, S10 Costos y Presupuestos, MS Project, ArcGIS/QGIS",
}

# --- Equipamiento estratégico -------------------------------------------------
EQUIPAMIENTO = {
    "obra": {
        "Vial": [("Tractor sobre orugas 190-240 HP", 1), ("Motoniveladora 125 HP", 1),
                 ("Rodillo liso vibratorio autopropulsado 10-12 t", 1), ("Cargador frontal 125 HP", 1),
                 ("Excavadora sobre orugas 115-165 HP", 1), ("Camión volquete 15 m³", 4),
                 ("Camión cisterna 2000 gal", 1), ("Estación total con accesorios", 1),
                 ("Nivel topográfico con accesorios", 1)],
        "Edificaciones": [("Mezcladora de concreto 9-11 p³", 2), ("Vibrador de concreto 4 HP", 2),
                          ("Compactadora tipo plancha 5.8 HP", 1), ("Retroexcavadora sobre llantas 62 HP", 1),
                          ("Winche eléctrico / elevador de carga", 1), ("Andamios metálicos (cuerpos)", 40),
                          ("Estación total con accesorios", 1), ("Nivel topográfico con accesorios", 1)],
        "Saneamiento": [("Retroexcavadora sobre llantas 62 HP", 1), ("Compactadora tipo canguro (apisonador)", 2),
                        ("Equipo para prueba hidráulica de tuberías", 1), ("Mezcladora de concreto 9-11 p³", 1),
                        ("Vibrador de concreto 4 HP", 1), ("Motobomba 4\" para achique", 1),
                        ("Camión volquete 15 m³", 2), ("Estación total con accesorios", 1)],
        "Saneamiento/Riego": [("Excavadora sobre orugas 115-165 HP", 1), ("Retroexcavadora sobre llantas 62 HP", 1),
                              ("Cargador frontal 125 HP", 1), ("Camión volquete 15 m³", 3),
                              ("Mezcladora de concreto 9-11 p³", 2), ("Vibrador de concreto 4 HP", 2),
                              ("Motobomba 4\" para achique", 1), ("Estación total con accesorios", 1)],
    },
    "expediente": [("Estación total con accesorios (certificado de calibración vigente)", 1),
                   ("GPS diferencial geodésico de doble frecuencia", 1),
                   ("Nivel topográfico automático", 1), ("Drone para fotogrametría (opcional)", 1),
                   ("Computadoras con software licenciado de diseño", 4),
                   ("Camioneta 4x4 doble cabina", 1)],
    "supervision": [("Camioneta 4x4 doble cabina", 1), ("Estación total con accesorios", 1),
                    ("Nivel topográfico automático", 1),
                    ("Equipo de control de densidad de campo (cono de arena o densímetro)", 1),
                    ("Esclerómetro y equipo para ensayo de concreto fresco (cono de Abrams)", 1),
                    ("Computadoras con software licenciado", 3)],
    "servicio": [("Camioneta 4x4 doble cabina", 1), ("Computadoras con software licenciado", 2),
                 ("Equipos de protección personal (EPP) completos", 10)],
}

# --- Actividades técnicas por especialidad -----------------------------------
ACT_DISENO = {
    "Vial": [
        ("Estudio de tráfico", "Conteos vehiculares clasificados de 7 días consecutivos, encuestas origen-destino, determinación del IMDA, proyección del tráfico normal, generado y derivado, y cálculo del número de repeticiones de ejes equivalentes (ESAL) para el periodo de diseño."),
        ("Estudio de suelos, canteras y fuentes de agua", "Ejecución de calicatas cada 500 m como mínimo o según el Manual de Suelos, Geología, Geotecnia y Pavimentos, ensayos estándar y especiales (granulometría, límites de Atterberg, Proctor modificado, CBR), caracterización de canteras, fuentes de agua y botaderos (DME)."),
        ("Estudio geológico y geotécnico", "Identificación de zonas inestables, taludes críticos, deslizamientos y zonas de falla; diseño de estabilización de taludes y propuesta de tratamiento de puntos críticos."),
        ("Diseño geométrico", "Diseño en planta, perfil longitudinal y secciones transversales conforme al Manual DG-2018, definición de la velocidad de diseño, radios mínimos, peraltes, sobreanchos, distancias de visibilidad y secciones típicas."),
        ("Diseño de pavimentos", "Diseño estructural del pavimento (flexible, rígido o afirmado) por los métodos AASHTO 93 y/o mecanicista-empírico, con alternativas comparativas técnico-económicas."),
        ("Estudio hidrológico y de drenaje", "Análisis de precipitaciones máximas, caudales de diseño, dimensionamiento de cunetas, alcantarillas, badenes, subdrenes y obras de arte conforme al Manual HHD-2011."),
        ("Diseño de estructuras y obras de arte", "Diseño estructural de muros de contención, pontones, puentes y obras complementarias, con memorias de cálculo y planos de detalle."),
        ("Señalización y seguridad vial", "Diseño de señalización vertical, horizontal, elementos de contención (guardavías) y dispositivos de seguridad conforme al Manual DCT-2016, incluyendo auditoría de seguridad vial."),
    ],
    "Edificaciones": [
        ("Arquitectura", "Desarrollo del programa arquitectónico, anteproyecto y proyecto definitivo: plantas, cortes, elevaciones, detalles, cuadro de acabados, carpintería, accesibilidad universal (A.120) y evacuación y señalización (A.130)."),
        ("Estructuras", "Análisis y diseño estructural sismorresistente conforme a la E.030, modelamiento tridimensional, diseño de cimentaciones según el EMS, diseño de elementos de concreto armado, albañilería y estructuras metálicas."),
        ("Instalaciones sanitarias", "Diseño de redes de agua fría, agua caliente, desagüe, ventilación, drenaje pluvial y sistema contra incendios conforme a la IS.010 y la A.130; cálculo de dotaciones, cisternas y tanques elevados."),
        ("Instalaciones eléctricas", "Diseño de redes eléctricas, alumbrado, tomacorrientes, tableros, puesta a tierra, cálculo de máxima demanda, coordinación de factibilidad de suministro con la concesionaria y sistema de pararrayos."),
        ("Instalaciones mecánicas y comunicaciones", "Diseño de sistemas de ventilación, aire acondicionado, ascensores, cableado estructurado, detección y alarma contra incendios, CCTV y sonido, según corresponda."),
        ("Estudio de mecánica de suelos con fines de cimentación", "Exploración de campo con calicatas y/o perforaciones según la E.050, ensayos de laboratorio, perfil estratigráfico, capacidad portante, asentamientos, agresividad química y parámetros sísmicos."),
        ("Eficiencia energética y confort", "Análisis de asoleamiento, ventilación natural, iluminación (DIALux) y propuestas de eficiencia energética e hídrica."),
        ("Equipamiento y mobiliario", "Determinación del equipamiento y mobiliario necesario, especificaciones técnicas y presupuesto de equipamiento, según la naturaleza del proyecto."),
    ],
    "Saneamiento": [
        ("Diagnóstico de los servicios existentes", "Evaluación del estado operativo de captaciones, líneas de conducción, plantas de tratamiento, reservorios, redes y conexiones; catastro técnico y comercial de usuarios."),
        ("Estudio de fuentes de agua", "Aforos, evaluación de la oferta hídrica, análisis físico-químicos y bacteriológicos, y gestión de la acreditación de disponibilidad hídrica ante la ANA."),
        ("Estudio de población y demanda", "Determinación de la población de diseño, dotaciones, coeficientes de variación, caudales de diseño y balance oferta-demanda para el horizonte del proyecto."),
        ("Diseño hidráulico del sistema de agua potable", "Modelamiento hidráulico (WaterCAD/EPANET) de líneas de conducción, aducción, redes de distribución y sectorización, verificación de presiones y velocidades conforme a la OS.050."),
        ("Diseño del sistema de alcantarillado", "Diseño hidráulico de colectores, buzones, emisores y conexiones domiciliarias conforme a la OS.070, verificación de tensión tractiva y pendientes mínimas."),
        ("Diseño de plantas de tratamiento", "Dimensionamiento de procesos de la PTAP y/o PTAR conforme a la OS.020 y OS.090, cumplimiento de los LMP y ECA vigentes, y diseño de la disposición final de lodos."),
        ("Diseño estructural de componentes", "Diseño estructural de reservorios, cámaras, captaciones, cisternas y estructuras especiales, considerando la E.030, E.060 y el ACI 350."),
        ("Gestión de factibilidades y servidumbres", "Obtención de factibilidad de servicios ante la EPS u operador, autorización de vertimiento, libre disponibilidad de terrenos y servidumbres."),
    ],
    "Saneamiento/Riego": [
        ("Estudio hidrológico", "Análisis de series históricas de caudales y precipitaciones, generación de caudales medios y máximos, determinación de la oferta hídrica al 75 % de persistencia y caudales de avenidas para diseño."),
        ("Estudio agrológico y de demanda de agua", "Clasificación de tierras, cédula de cultivo, cálculo de la evapotranspiración (CROPWAT), demanda de riego, eficiencia de riego y balance hídrico."),
        ("Diseño hidráulico de obras de captación y conducción", "Diseño de bocatomas, desarenadores, canales, sifones, acueductos, caídas y obras de arte conforme a los criterios de la ANA."),
        ("Diseño de almacenamiento", "Diseño de reservorios, presas o qochas, incluyendo estabilidad, impermeabilización (geomembrana), aliviaderos y obras de toma."),
        ("Diseño de riego tecnificado", "Diseño agronómico e hidráulico de sistemas de riego por aspersión o goteo, cabezal de filtrado, red de tuberías, hidrantes y parcelas demostrativas."),
        ("Estudio geológico-geotécnico", "Caracterización geológica del eje de las estructuras, sondeos, ensayos de permeabilidad y estabilidad de taludes de canales y presas."),
        ("Diseño estructural", "Diseño estructural de bocatomas, muros, canales, reservorios y estructuras hidráulicas conforme a la E.060 y el ACI 350."),
        ("Gestión de recursos hídricos y organización de usuarios", "Acreditación de disponibilidad hídrica, autorización de ejecución de obras ante la ANA, plan de operación y mantenimiento y fortalecimiento de la organización de usuarios."),
    ],
}

ACT_EJECUCION = {
    "Vial": [
        ("Obras preliminares", "Movilización y desmovilización de equipos, campamento, cartel de obra, trazo, nivelación y replanteo topográfico, mantenimiento de tránsito y seguridad vial durante la ejecución."),
        ("Movimiento de tierras", "Desbroce y limpieza, cortes en material suelto y roca, terraplenes con material propio y de préstamo, perfilado y compactación de subrasante, y eliminación de material excedente a DME autorizados."),
        ("Pavimentos", "Conformación de subbase y base granular, imprimación asfáltica, carpeta asfáltica en caliente o tratamiento superficial, o losas de concreto hidráulico, según el expediente técnico, con control de espesores y densidades."),
        ("Obras de arte y drenaje", "Construcción de cunetas revestidas, alcantarillas TMC o de concreto, badenes, subdrenes, muros de contención y pontones, con sus respectivos ensayos de calidad."),
        ("Señalización y seguridad vial", "Instalación de señales verticales preventivas, reglamentarias e informativas, marcas en el pavimento, postes delineadores, hitos kilométricos y guardavías."),
        ("Control de calidad vial", "Ensayos de densidad de campo, CBR, granulometría, Marshall, lavado asfáltico, extracción de diamantinas y medición de rugosidad (IRI) y deflexiones (Viga Benkelman)."),
        ("Mitigación ambiental", "Restauración de canteras, DME, campamentos y patios de máquinas; riego para el control de polvo y revegetación de áreas afectadas."),
    ],
    "Edificaciones": [
        ("Obras provisionales y preliminares", "Almacén, oficinas, servicios higiénicos, cerco perimétrico provisional, demoliciones, limpieza de terreno, trazo, nivelación y replanteo."),
        ("Estructuras", "Movimiento de tierras, cimentaciones, concreto armado en zapatas, columnas, vigas, losas y placas; muros de albañilería; estructuras metálicas y coberturas, conforme a planos y especificaciones."),
        ("Arquitectura y acabados", "Muros y tabiques, tarrajeos, pisos, contrapisos, zócalos, cielorrasos, carpintería de madera, metálica y de aluminio, vidrios, pintura y aparatos sanitarios."),
        ("Instalaciones sanitarias", "Redes de agua fría y caliente, desagüe, ventilación, drenaje pluvial, cisternas, tanques, equipos de bombeo y sistema contra incendios, con pruebas hidráulicas."),
        ("Instalaciones eléctricas y mecánicas", "Acometida, tableros, alimentadores, circuitos de alumbrado y tomacorrientes, pozos a tierra, comunicaciones, ascensores y sistemas especiales, con protocolos de pruebas."),
        ("Control de calidad en edificaciones", "Ensayos de resistencia a la compresión del concreto, slump, ensayos de unidades de albañilería, densidad de campo, pruebas de instalaciones y protocolos de calidad por partida."),
        ("Equipamiento y obras exteriores", "Veredas, rampas, áreas verdes, cerco perimétrico definitivo, mobiliario y equipamiento, según el expediente técnico."),
    ],
    "Saneamiento": [
        ("Obras preliminares", "Movilización, cartel de obra, trazo y replanteo de redes, desvío de tránsito, señalización y cercos de seguridad en zanjas."),
        ("Captación, conducción y tratamiento", "Construcción de captaciones, cámaras rompe presión, líneas de conducción, plantas de tratamiento y casetas de cloración, con sus pruebas hidráulicas y de estanqueidad."),
        ("Almacenamiento", "Construcción de reservorios apoyados o elevados, con pruebas de estanqueidad, desinfección y equipamiento hidráulico."),
        ("Redes de agua potable y conexiones", "Excavación de zanjas, refine y nivelación, cama de apoyo, instalación de tuberías y accesorios, válvulas, grifos contra incendio, conexiones domiciliarias, relleno compactado y pruebas hidráulicas."),
        ("Redes de alcantarillado y conexiones", "Instalación de colectores, buzones, emisores y conexiones domiciliarias, pruebas de nivelación, de escorrentía e hidráulicas, e inspección televisiva cuando corresponda."),
        ("Tratamiento de aguas residuales", "Construcción de los componentes de la PTAR, equipamiento electromecánico y puesta en marcha, con verificación de la eficiencia de remoción."),
        ("Reposición de pavimentos y veredas", "Reposición de pavimentos, veredas y áreas afectadas a su estado original o mejorado, con control de calidad."),
    ],
    "Saneamiento/Riego": [
        ("Obras preliminares", "Movilización, campamento, cartel de obra, trazo, nivelación y replanteo, caminos de acceso y desvío temporal del cauce cuando corresponda."),
        ("Obras de captación", "Construcción de bocatomas, barrajes, compuertas, desarenadores y obras de protección, con control de niveles y cotas hidráulicas."),
        ("Canales y obras de arte", "Excavación, perfilado, revestimiento de canales de concreto, juntas, sifones, acueductos, caídas, tomas laterales y partidores."),
        ("Almacenamiento", "Construcción de reservorios, presas o qochas, impermeabilización con geomembrana, aliviaderos, obras de toma y cerco perimétrico."),
        ("Riego tecnificado", "Instalación del cabezal de filtrado, tuberías de conducción y distribución, hidrantes, válvulas, aspersores o goteros y pruebas del sistema."),
        ("Control de calidad hidráulico", "Ensayos de concreto, densidad de campo, pruebas de estanqueidad y de funcionamiento hidráulico, y verificación de caudales de diseño."),
        ("Capacitación y operación", "Capacitación a la organización de usuarios en operación y mantenimiento, manual de O&M y entrega formal de la infraestructura."),
    ],
}

ACT_COMUNES = {
    "expediente": [
        ("Plan de trabajo", "Presentación del plan de trabajo, cronograma detallado de actividades, organigrama del equipo profesional y metodología específica, dentro de los primeros días del plazo contractual."),
        ("Diagnóstico y recopilación de información", "Revisión del estudio de preinversión o ficha técnica aprobada en Invierte.pe, verificación de la concepción técnica, visita de campo y recopilación de información existente."),
        ("Levantamiento topográfico", "Levantamiento topográfico georreferenciado al sistema WGS-84 con BMs enlazados a puntos geodésicos del IGN, curvas de nivel a escala adecuada y monumentación de puntos de control."),
        ("Saneamiento físico-legal", "Verificación de la libre disponibilidad de terrenos, actas de libre disponibilidad, servidumbres, afectaciones prediales y gestión del CIRA ante el Ministerio de Cultura."),
        ("Evaluación de riesgos de desastres", "Identificación de peligros, análisis de vulnerabilidad y estimación del riesgo, con medidas estructurales y no estructurales de reducción del riesgo."),
        ("Estudio de impacto ambiental", "Elaboración del instrumento de gestión ambiental correspondiente (FTA, DIA o el que determine la autoridad sectorial) y gestión de su aprobación o conformidad."),
        ("Metrados, costos y presupuestos", "Metrados por partidas sustentados con planillas, análisis de precios unitarios con insumos cotizados en la zona, presupuesto de obra, desagregado de gastos generales, fórmulas polinómicas y relación de insumos."),
        ("Cronogramas", "Cronograma de ejecución de obra (PERT-CPM / Gantt) con ruta crítica, cronograma valorizado, cronograma de adquisición de materiales y calendario de utilización de equipos."),
        ("Especificaciones técnicas", "Especificaciones técnicas por partida: descripción, materiales, equipos, método de construcción, control de calidad, método de medición y bases de pago."),
        ("Planos de ejecución", "Juego completo de planos por especialidad, a escala adecuada, firmados por los profesionales responsables, en formato físico y digital editable."),
        ("Plan de seguridad y salud en el trabajo", "Elaboración del Plan de SST conforme a la Norma G.050, con su presupuesto específico, matriz IPERC y partidas de seguridad."),
        ("Plan de mantenimiento", "Elaboración del plan de operación y mantenimiento de la infraestructura proyectada, con estimación de costos anuales."),
        ("Gestión de riesgos en la ejecución", "Identificación, análisis y asignación de riesgos previsibles de la ejecución de obra, con el formato de asignación de riesgos requerido por la normativa."),
        ("Levantamiento de observaciones", "Absolución de las observaciones formuladas por la Entidad dentro de los plazos establecidos, sin costo adicional."),
    ],
    "supervision": [
        ("Revisión del expediente técnico", "Revisión integral del expediente técnico de obra y emisión del informe de compatibilidad con el terreno, identificando deficiencias, omisiones o incongruencias."),
        ("Entrega de terreno e inicio de obra", "Participación en el acto de entrega de terreno, verificación de condiciones para el inicio del plazo y apertura del cuaderno de obra digital."),
        ("Control técnico", "Verificación de que los trabajos se ejecuten conforme al expediente técnico, especificaciones y normas aplicables, aprobando o rechazando los trabajos ejecutados."),
        ("Control de calidad", "Verificación y aprobación de los ensayos de laboratorio y campo, certificados de calidad de materiales y protocolos por partida, y ejecución de ensayos de contraste."),
        ("Control de plazos", "Seguimiento del avance físico mediante el cronograma valorizado y curva S, alertas tempranas de atrasos y exigencia de cronogramas acelerados cuando corresponda."),
        ("Control económico y valorizaciones", "Elaboración y aprobación de metrados y valorizaciones mensuales, reajustes por fórmulas polinómicas, amortización de adelantos y control del presupuesto."),
        ("Prestaciones adicionales, deductivos y ampliaciones de plazo", "Evaluación y emisión de pronunciamiento técnico sobre consultas, prestaciones adicionales, reducciones, ampliaciones de plazo y mayores gastos generales, dentro de los plazos legales."),
        ("Supervisión de seguridad, salud y medio ambiente", "Verificación del cumplimiento del Plan de SST, del instrumento de gestión ambiental y del manejo de residuos de construcción."),
        ("Recepción de obra", "Verificación de la culminación de la obra, participación en el comité de recepción y verificación del levantamiento de observaciones."),
        ("Liquidación del contrato de obra", "Revisión o elaboración de la liquidación del contrato de obra, planos de replanteo (post construcción) y memoria descriptiva valorizada."),
        ("Informes", "Presentación de informes mensuales, informes especiales y el informe final, conforme al contenido mínimo establecido en los presentes TDR."),
    ],
    "obra": [
        ("Gestión de inicio", "Presentación del calendario de avance de obra valorizado actualizado, calendario de adquisición de materiales, programación CPM y apertura del cuaderno de obra."),
        ("Ejecución conforme al expediente técnico", "Ejecución de todas las partidas del presupuesto de obra conforme a los planos, especificaciones técnicas y normas aplicables, bajo la dirección del residente de obra."),
        ("Control de calidad del contratista", "Implementación del Plan de Aseguramiento y Control de Calidad, ejecución de ensayos en laboratorios acreditados y elaboración de protocolos por partida."),
        ("Seguridad y salud en el trabajo", "Implementación del Plan de SST específico, dotación de EPP, señalización, SCTR del personal, capacitaciones y registros obligatorios."),
        ("Gestión ambiental y social", "Cumplimiento de las medidas del instrumento de gestión ambiental, manejo de residuos de construcción, relaciones comunitarias y contratación de mano de obra local cuando corresponda."),
        ("Valorizaciones y documentación", "Elaboración de metrados y valorizaciones mensuales, registro de ocurrencias en el cuaderno de obra y presentación de informes mensuales de avance."),
        ("Planos de replanteo", "Elaboración de los planos de replanteo (post construcción / as built) y memoria descriptiva valorizada al culminar la obra."),
        ("Recepción y liquidación", "Solicitud de recepción, levantamiento de observaciones, entrega de manuales, garantías de equipos y elaboración de la liquidación del contrato de obra."),
    ],
    "servicio": [
        ("Plan de trabajo", "Presentación del plan de trabajo con cronograma, organización del personal y metodología de ejecución del servicio."),
        ("Ejecución de las actividades del servicio", "Desarrollo de las actividades técnicas del servicio con el personal, equipos e insumos ofertados, conforme a los presentes términos de referencia."),
        ("Control de calidad del servicio", "Aplicación de protocolos y registros de control que acrediten la calidad del servicio prestado."),
        ("Seguridad y salud", "Cumplimiento de la normativa de SST, dotación de EPP y SCTR del personal destacado."),
        ("Coordinación con la Entidad", "Reunión de inicio, reuniones periódicas de seguimiento con el coordinador designado y registro de acuerdos en actas."),
        ("Gestión documental", "Organización, codificación y archivo de toda la documentación generada, en formato físico y digital, entregada a la Entidad al cierre."),
        ("Gestión ambiental y de residuos", "Manejo adecuado de los residuos generados, cumplimiento de la normativa ambiental aplicable y limpieza de las áreas intervenidas."),
        ("Capacitación y transferencia", "Capacitación al personal de la Entidad sobre los resultados del servicio y transferencia de información, manuales y archivos."),
        ("Informes", "Presentación de los informes de avance y del informe final con evidencias fotográficas y registros de las actividades ejecutadas."),
        ("Cierre del servicio", "Verificación de la culminación de todas las actividades, levantamiento de observaciones y suscripción del acta de conformidad."),
    ],
}

# Factores de evaluación referenciales (editables)
FACTORES_DEFAULT = {
    "consultoria": [
        ("Precio", "Oferta económica. Pi = (Om x PMP) / Oi", 20),
        ("Experiencia del postor en la especialidad", "Monto facturado acumulado en servicios de consultoría similares (hasta 3 veces el valor referencial)", 50),
        ("Metodología propuesta", "Sustento de la metodología, plan de trabajo, identificación de riesgos y control de calidad", 25),
        ("Sostenibilidad ambiental y social / Integridad", "Certificaciones ISO 14001, ISO 45001 o ISO 37001 vigentes", 5),
    ],
    "obra": [
        ("Precio", "Oferta económica. Pi = (Om x PMP) / Oi", 80),
        ("Experiencia del postor en la especialidad", "Monto facturado acumulado en obras similares (hasta 3 veces el valor referencial)", 10),
        ("Metodología propuesta", "Plan de ejecución, procedimiento constructivo y gestión de riesgos", 5),
        ("Sostenibilidad ambiental y social / Integridad", "Certificaciones ISO 14001, ISO 45001 o ISO 37001 vigentes", 5),
    ],
    "servicio": [
        ("Precio", "Oferta económica. Pi = (Om x PMP) / Oi", 70),
        ("Experiencia del postor en la especialidad", "Monto facturado acumulado en servicios similares", 20),
        ("Metodología propuesta", "Plan de trabajo y metodología de ejecución del servicio", 10),
    ],
}

COLUMNAS_PERSONAL = [
    ("cargo", "Cargo / Personal Clave", 190),
    ("formacion", "Formación Académica", 200),
    ("colegiatura", "Colegiatura", 85),
    ("exp_gen", "Exp. Gral. (años)", 75),
    ("exp_esp", "Exp. Espec. (años)", 75),
    ("cap_h", "Capacit. (h)", 70),
    ("cap_tema", "Tema de Capacitación", 210),
    ("exp_desc", "Experiencia específica como", 230),
]

COLUMNAS_ENTREGABLES = [
    ("descripcion", "Entregable / Contenido", 430),
    ("plazo", "Plazo (días)", 90),
    ("porcentaje", "% de Pago", 90),
]

COLUMNAS_FACTORES = [
    ("factor", "Factor de Evaluación", 300),
    ("criterio", "Criterio / Metodología de Asignación", 470),
    ("puntaje", "Puntaje Máx.", 90),
]


# =============================================================================
# 2. MOTOR DE CÁLCULO NORMATIVO
# =============================================================================

def a_float(valor, defecto=0.0):
    """Convierte texto con comas o 'S/' a float."""
    try:
        if isinstance(valor, (int, float)):
            return float(valor)
        txt = str(valor).replace("S/", "").replace(",", "").replace(" ", "").strip()
        return float(txt) if txt else defecto
    except (ValueError, TypeError):
        return defecto


def a_int(valor, defecto=0):
    return int(round(a_float(valor, defecto)))


def fmt_num(valor):
    """Número sin notación científica ni ceros sobrantes (2300000, 10, 0.5)."""
    if isinstance(valor, (int, float)):
        txt = f"{float(valor):.4f}".rstrip("0").rstrip(".")
        return txt or "0"
    return str(valor)


def soles(monto):
    return f"S/ {monto:,.2f}"


def clasificar_procedimiento(monto, objeto, params):
    """Determina la categoría del procedimiento según el monto referencial."""
    tipo = TIPO_OBJETO.get(objeto, "servicio")
    cat = CATEGORIA_UMBRAL[tipo]
    uit = a_float(params.get("uit", UIT_DEFAULT), UIT_DEFAULT)
    limite_menor = a_float(params.get("menor_cuantia_uit", 8)) * uit
    limite_as = a_float(params.get(f"limite_as_{cat}",
                                   PARAMETROS_DEFAULT[f"limite_as_{cat}"]))
    if monto <= limite_menor:
        return CAT_MENOR
    if monto <= limite_as:
        return CAT_AS
    return CAT_MAYOR


def nombre_procedimiento(categoria, objeto):
    tipo = TIPO_OBJETO.get(objeto, "servicio")
    if categoria == CAT_MENOR:
        return "Contratación por monto igual o inferior a 8 UIT (menor cuantía)"
    if categoria == CAT_AS:
        return "Adjudicación Simplificada"
    return "Licitación Pública" if tipo == "obra" else "Concurso Público"


def generar_personal_clave(objeto, especialidad, categoria):
    """Genera los perfiles de personal clave según el OSCE/OECE."""
    tipo = TIPO_OBJETO.get(objeto, "servicio")
    crit = CRITERIOS_PERSONAL[categoria]
    lid, esp = crit["lider"], crit["especialista"]
    similares = "obras similares" if tipo in ("obra", "supervision") else "consultorías de obras similares"
    form_lider, col_lider = FORMACION_LIDER[especialidad]
    e_cargo, e_form, e_col, e_tema = ESPECIALISTA_ESPECIALIDAD[especialidad]

    def fila(cargo, formacion, colegiatura, nivel, tema, desc):
        return {"cargo": cargo, "formacion": formacion, "colegiatura": colegiatura,
                "exp_gen": str(nivel["gen"]), "exp_esp": str(nivel["esp"]),
                "cap_h": str(nivel["cap"]), "cap_tema": tema, "exp_desc": desc}

    filas = []
    if tipo == "obra":
        filas.append(fila("Residente de Obra", form_lider, col_lider, lid,
                          "Gestión de obras, residencia y supervisión de obras / gestión de contratos del Estado",
                          f"Residente, supervisor, jefe de supervisión o inspector en {similares}"))
        filas.append(fila(e_cargo, e_form, e_col, esp, e_tema,
                          f"Especialista o responsable de la especialidad en {similares}"))
        filas.append(fila("Especialista en Estructuras", "Ingeniero Civil", "CIP", esp,
                          "Diseño sismorresistente y construcción de estructuras de concreto armado",
                          f"Especialista o ingeniero de estructuras en {similares}"))
        filas.append(fila("Especialista en Costos, Metrados y Valorizaciones",
                          "Ingeniero Civil", "CIP", esp,
                          "Costos y presupuestos, valorizaciones y liquidación de obras (S10)",
                          f"Especialista o ingeniero de costos, metrados y/o valorizaciones en {similares}"))
    elif tipo == "supervision":
        filas.append(fila("Jefe de Supervisión", form_lider, col_lider, lid,
                          "Supervisión de obras y gestión de contratos del Estado",
                          f"Jefe de supervisión, supervisor, inspector o residente en {similares}"))
        filas.append(fila(e_cargo, e_form, e_col, esp, e_tema,
                          f"Especialista o supervisor de la especialidad en {similares}"))
        filas.append(fila("Especialista en Estructuras", "Ingeniero Civil", "CIP", esp,
                          "Diseño y construcción de estructuras de concreto armado",
                          f"Especialista o supervisor de estructuras en {similares}"))
        filas.append(fila("Especialista en Costos y Presupuestos", "Ingeniero Civil", "CIP", esp,
                          "Costos y presupuestos, valorizaciones, adicionales y liquidación de obras",
                          f"Especialista en costos, metrados, valorizaciones o liquidaciones en {similares}"))
    elif tipo == "expediente":
        filas.append(fila("Jefe de Proyecto", form_lider, col_lider, lid,
                          "Formulación y elaboración de expedientes técnicos / gestión de proyectos",
                          f"Jefe de proyecto, coordinador o proyectista en la elaboración de expedientes técnicos de {SIMILARES[especialidad].split(',')[0]} y similares"))
        filas.append(fila(e_cargo, e_form, e_col, esp, e_tema,
                          f"Especialista de la especialidad en la elaboración de expedientes técnicos de {similares}"))
        filas.append(fila("Especialista en Estructuras", "Ingeniero Civil", "CIP", esp,
                          "Análisis y diseño estructural (ETABS / SAP2000)",
                          f"Especialista en estructuras en la elaboración de expedientes técnicos de {similares}"))
        filas.append(fila("Especialista en Costos y Presupuestos", "Ingeniero Civil", "CIP", esp,
                          "Costos y presupuestos (S10) y programación de obras (MS Project)",
                          f"Especialista en costos y presupuestos en la elaboración de expedientes técnicos de {similares}"))
        if especialidad in ("Vial", "Saneamiento/Riego", "Saneamiento"):
            filas.append(fila("Especialista en Topografía y Geodesia", "Ingeniero Civil o Ingeniero Geógrafo",
                              "CIP", esp, "Topografía, geodesia y sistemas de información geográfica",
                              f"Especialista o responsable de topografía en {similares}"))
    else:  # servicio
        filas.append(fila("Coordinador / Jefe del Servicio", form_lider, col_lider, lid,
                          "Gestión de proyectos y contratos de servicios",
                          f"Coordinador, jefe o responsable de servicios relacionados con {SIMILARES[especialidad].split(',')[0]}"))
        filas.append(fila(e_cargo, e_form, e_col, esp, e_tema,
                          "Especialista o responsable técnico en servicios similares"))

    filas.append(fila("Especialista en Seguridad y Salud en el Trabajo (SSOMA)",
                      "Ingeniero de Seguridad, Ingeniero Industrial, Ingeniero Civil o Ingeniero Ambiental",
                      "CIP", esp,
                      "Seguridad y salud ocupacional en construcción (Ley N° 29783 / Norma G.050)",
                      f"Especialista, supervisor, jefe o responsable de seguridad y salud en {similares}"))
    if tipo != "servicio":
        filas.append(fila("Especialista Ambiental", "Ingeniero Ambiental, Ingeniero Forestal, Ingeniero Geógrafo o Biólogo",
                          "CIP / CBP", esp,
                          "Gestión ambiental, evaluación de impacto ambiental y SEIA",
                          f"Especialista, supervisor o responsable ambiental en {similares}"))
    return filas


def _repartir_porcentajes(n, total=100.0):
    if n <= 0:
        return []
    base = math.floor(total / n * 100) / 100
    lista = [base] * n
    lista[-1] = round(total - base * (n - 1), 2)
    return lista


def generar_entregables(objeto, plazo):
    """Configura entregables, plazos y porcentajes según el objeto."""
    tipo = TIPO_OBJETO.get(objeto, "servicio")
    plazo = max(1, int(plazo))
    if tipo == "expediente":
        d1, d2 = max(1, round(plazo * 0.30)), max(2, round(plazo * 0.65))
        return [
            {"descripcion": "Primer Entregable: Plan de trabajo, diagnóstico, levantamiento topográfico, estudio de mecánica de suelos, estudios básicos y anteproyecto",
             "plazo": str(d1), "porcentaje": "30"},
            {"descripcion": "Segundo Entregable: Ingeniería de detalle por especialidades (memorias de cálculo, planos de desarrollo, metrados preliminares e instrumento de gestión ambiental)",
             "plazo": str(d2), "porcentaje": "40"},
            {"descripcion": "Informe Final: Expediente técnico completo (memoria descriptiva, especificaciones técnicas, metrados, presupuesto, APU, fórmulas polinómicas, cronogramas, planos finales, plan de SST y plan de mantenimiento)",
             "plazo": str(plazo), "porcentaje": "30"},
        ]
    if tipo in ("supervision", "obra"):
        meses = max(1, math.ceil(plazo / 30))
        filas = []
        if tipo == "supervision":
            pct = _repartir_porcentajes(meses, 90.0)
            for i in range(meses):
                filas.append({"descripcion": f"Informe mensual de supervisión N° {i + 1:02d} (valorización de la supervisión y de la obra)",
                              "plazo": str(min(plazo, 30 * (i + 1))), "porcentaje": f"{pct[i]:g}"})
            filas.append({"descripcion": "Informe final, recepción de obra y revisión/elaboración de la liquidación del contrato de obra",
                          "plazo": str(plazo + 60), "porcentaje": "10"})
        else:
            pct = _repartir_porcentajes(meses, 100.0)
            for i in range(meses):
                filas.append({"descripcion": f"Valorización mensual N° {i + 1:02d} por avance de obra ejecutado (referencial según calendario valorizado)",
                              "plazo": str(min(plazo, 30 * (i + 1))), "porcentaje": f"{pct[i]:g}"})
        return filas
    d1, d2 = max(1, round(plazo * 0.33)), max(2, round(plazo * 0.66))
    return [
        {"descripcion": "Primer Entregable: Plan de trabajo e informe de avance N° 01", "plazo": str(d1), "porcentaje": "30"},
        {"descripcion": "Segundo Entregable: Informe de avance N° 02", "plazo": str(d2), "porcentaje": "40"},
        {"descripcion": "Informe Final del servicio", "plazo": str(plazo), "porcentaje": "30"},
    ]


def factor_f(objeto, plazo):
    """Factor F del Reglamento: 0.40 si plazo <= 60 d; si > 60 d: 0.25
    (bienes, servicios y consultorías) o 0.15 (obras)."""
    tipo = TIPO_OBJETO.get(objeto, "servicio")
    if plazo <= 60:
        return 0.40
    return 0.15 if tipo == "obra" else 0.25


def calcular_penalidad(monto, plazo, objeto):
    plazo = max(1, int(plazo))
    f = factor_f(objeto, plazo)
    diaria = 0.10 * monto / (f * plazo)
    tope = 0.10 * monto
    dias_tope = math.ceil(tope / diaria) if diaria > 0 else 0
    return {"F": f, "diaria": diaria, "tope": tope, "dias_tope": dias_tope}


def calcular_garantias(monto, objeto, params):
    tipo = TIPO_OBJETO.get(objeto, "servicio")
    gfc_pct = a_float(params.get("garantia_fiel_pct", 10))
    if tipo == "obra":
        ad_pct = a_float(params.get("adelanto_directo_obra_pct", 10))
        mat_pct = a_float(params.get("adelanto_materiales_obra_pct", 20))
    else:
        ad_pct = a_float(params.get("adelanto_directo_otros_pct", 30))
        mat_pct = 0.0
    return {
        "gfc_pct": gfc_pct, "gfc": monto * gfc_pct / 100,
        "ad_pct": ad_pct, "ad": monto * ad_pct / 100,
        "mat_pct": mat_pct, "mat": monto * mat_pct / 100,
    }


# --- Número a letras (montos en soles) ---------------------------------------
_UNIDADES = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE",
             "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE",
             "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUNO", "VEINTIDÓS", "VEINTITRÉS",
             "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"]
_DECENAS = ["", "", "", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
_CENTENAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS",
             "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"]


def _cientos(n):
    if n == 0:
        return ""
    if n == 100:
        return "CIEN"
    c, r = divmod(n, 100)
    partes = [_CENTENAS[c]] if c else []
    if r < 30:
        if r:
            partes.append(_UNIDADES[r])
    else:
        d, u = divmod(r, 10)
        partes.append(_DECENAS[d] + (f" Y {_UNIDADES[u]}" if u else ""))
    return " ".join(partes)


def numero_a_letras(numero):
    entero = int(numero)
    centimos = int(round((numero - entero) * 100))
    if centimos == 100:
        entero, centimos = entero + 1, 0
    if entero == 0:
        texto = "CERO"
    else:
        millones, resto = divmod(entero, 1_000_000)
        miles, cientos = divmod(resto, 1000)
        partes = []
        if millones:
            partes.append("UN MILLÓN" if millones == 1 else f"{_cientos(millones) if millones < 1000 else numero_a_letras_simple(millones)} MILLONES")
        if miles:
            partes.append("MIL" if miles == 1 else f"{_cientos(miles)} MIL")
        if cientos:
            partes.append(_cientos(cientos))
        texto = " ".join(partes)
    texto = texto.replace("VEINTIUNO MIL", "VEINTIÚN MIL").replace("UNO MIL", "UN MIL").replace("UNO MILLONES", "UN MILLONES")
    return f"{texto} CON {centimos:02d}/100 SOLES"


def numero_a_letras_simple(n):
    miles, cientos = divmod(n, 1000)
    partes = []
    if miles:
        partes.append("MIL" if miles == 1 else f"{_cientos(miles)} MIL")
    if cientos:
        partes.append(_cientos(cientos))
    return " ".join(partes)


# =============================================================================
# 3. GENERADOR DEL TDR EN WORD (.DOCX)
# =============================================================================

AZUL_OSCURO = RGBColor(0x1F, 0x38, 0x64)
AZUL_HEX = "1F3864"
CELESTE_HEX = "D9E2F3"


def _sombrear(celda, color_hex):
    tc_pr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def _repetir_encabezado(fila):
    tr_pr = fila._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def _campo(parrafo, instruccion):
    """Inserta un campo de Word (PAGE, NUMPAGES...)."""
    run = parrafo.add_run()
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instruccion
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = "1"
    f3 = OxmlElement("w:fldChar")
    f3.set(qn("w:fldCharType"), "end")
    for el in (f1, it, f2, t, f3):
        run._r.append(el)
    return run


def _borde_inferior(parrafo, color="1F3864", grosor="12"):
    p_pr = parrafo._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), grosor)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    bdr.append(bottom)
    p_pr.append(bdr)


class GeneradorTDR:
    """Construye el documento oficial de Términos de Referencia."""

    def __init__(self, datos, personal, entregables, factores, params):
        self.d = datos
        self.personal = personal
        self.entregables = entregables
        self.factores = factores
        self.params = params
        self.tipo = TIPO_OBJETO.get(datos["objeto"], "servicio")
        self.marco = MARCOS_LEGALES.get(datos.get("marco", MARCO_DEFAULT),
                                        MARCOS_LEGALES[MARCO_DEFAULT])
        self.monto = a_float(datos["monto"])
        self.plazo = a_int(datos["plazo"])
        self.categoria = clasificar_procedimiento(self.monto, datos["objeto"], params)
        self.pen = calcular_penalidad(self.monto, self.plazo, datos["objeto"])
        self.gar = calcular_garantias(self.monto, datos["objeto"], params)
        self.doc = Document()
        self.esp = datos["especialidad"]
        self.n_tabla = 0

    # ---------- utilidades de formato ----------
    def _estilos(self):
        st = self.doc.styles
        normal = st["Normal"]
        normal.font.name = "Arial"
        normal.font.size = Pt(10.5)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        pf = normal.paragraph_format
        pf.space_after = Pt(6)
        pf.line_spacing = 1.15
        for nombre, tam in (("Heading 1", 13), ("Heading 2", 11.5), ("Heading 3", 10.5)):
            s = st[nombre]
            s.font.name = "Arial"
            s.font.size = Pt(tam)
            s.font.bold = True
            s.font.color.rgb = AZUL_OSCURO
            s.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
            s.paragraph_format.space_before = Pt(12 if nombre == "Heading 1" else 8)
            s.paragraph_format.space_after = Pt(6)
            s.paragraph_format.keep_with_next = True
        for nombre in ("List Bullet", "List Number"):
            st[nombre].font.name = "Arial"
            st[nombre].font.size = Pt(10.5)

    def _margenes(self):
        for s in self.doc.sections:
            s.page_height, s.page_width = Cm(29.7), Cm(21.0)
            s.top_margin, s.bottom_margin = Cm(3.0), Cm(2.5)
            s.left_margin, s.right_margin = Cm(3.0), Cm(2.5)
            s.header_distance, s.footer_distance = Cm(1.2), Cm(1.2)

    def _membrete(self):
        sec = self.doc.sections[0]
        sec.different_first_page_header_footer = True
        hdr = sec.header
        p = hdr.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(self.d["entidad"].upper())
        r.bold, r.font.size, r.font.color.rgb = True, Pt(9), AZUL_OSCURO
        p2 = hdr.add_paragraph()
        r2 = p2.add_run(f"Términos de Referencia - {self.d['objeto']} | "
                        f"{self._corto(self.d['proyecto'], 85)}")
        r2.italic, r2.font.size = True, Pt(8)
        _borde_inferior(p2)
        ftr = sec.footer
        pf = ftr.paragraphs[0]
        pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rr = pf.add_run(f"{self.marco['organismo']} / SEACE  -  Página ")
        rr.font.size = Pt(8)
        _campo(pf, "PAGE").font.size = Pt(8)
        r3 = pf.add_run(" de ")
        r3.font.size = Pt(8)
        _campo(pf, "NUMPAGES").font.size = Pt(8)

    @staticmethod
    def _corto(texto, n):
        return texto if len(texto) <= n else texto[: n - 3] + "..."

    def h1(self, texto):
        self.doc.add_heading(texto, level=1)

    def h2(self, texto):
        self.doc.add_heading(texto, level=2)

    def h3(self, texto):
        self.doc.add_heading(texto, level=3)

    def p(self, texto, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=None):
        par = self.doc.add_paragraph()
        par.alignment = align
        run = par.add_run(texto)
        run.bold, run.italic = bold, italic
        if size:
            run.font.size = Pt(size)
        return par

    def p_mixto(self, partes, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        """partes: lista de (texto, negrita)."""
        par = self.doc.add_paragraph()
        par.alignment = align
        for texto, neg in partes:
            par.add_run(texto).bold = neg
        return par

    def viñeta(self, texto, negrita_inicial=None):
        par = self.doc.add_paragraph(style="List Bullet")
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if negrita_inicial:
            par.add_run(negrita_inicial + ": ").bold = True
        par.add_run(texto)
        return par

    def literal(self, letra, titulo, texto):
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        par.paragraph_format.left_indent = Cm(0.8)
        par.paragraph_format.first_line_indent = Cm(-0.6)
        par.add_run(f"{letra}) ").bold = True
        par.add_run(f"{titulo}: ").bold = True
        par.add_run(texto)
        return par

    def tabla(self, encabezados, filas, anchos=None, titulo=None, tam=9, negrita_ultima=False):
        if titulo:
            self.n_tabla += 1
            cap = self.p(f"Cuadro N° {self.n_tabla:02d}: {titulo}", bold=True,
                         align=WD_ALIGN_PARAGRAPH.CENTER, size=9.5)
            cap.paragraph_format.keep_with_next = True
        t = self.doc.add_table(rows=1, cols=len(encabezados))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        fila_h = t.rows[0]
        _repetir_encabezado(fila_h)
        for i, h in enumerate(encabezados):
            c = fila_h.cells[i]
            c.text = ""
            run = c.paragraphs[0].add_run(h)
            run.bold, run.font.size = True, Pt(tam)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _sombrear(c, AZUL_HEX)
        for k, f in enumerate(filas):
            cells = t.add_row().cells
            for i, val in enumerate(f):
                cells[i].text = ""
                par = cells[i].paragraphs[0]
                run = par.add_run(str(val))
                run.font.size = Pt(tam)
                if negrita_ultima and k == len(filas) - 1:
                    run.bold = True
                    _sombrear(cells[i], CELESTE_HEX)
                cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                par.paragraph_format.space_after = Pt(2)
        if anchos:
            t.autofit = False
            for i, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
                if i < len(anchos):
                    gc.set(qn("w:w"), str(int(Cm(anchos[i]).twips)))
            for fila in t.rows:
                for i, w in enumerate(anchos):
                    fila.cells[i].width = Cm(w)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return t

    def salto(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---------- textos auxiliares ----------
    def _sustantivo(self):
        return {"obra": "la obra", "supervision": "el servicio de supervisión",
                "expediente": "la consultoría", "servicio": "el servicio"}[self.tipo]

    def _contratista(self):
        return {"obra": "EL CONTRATISTA", "supervision": "EL SUPERVISOR",
                "expediente": "EL CONSULTOR", "servicio": "EL CONTRATISTA"}[self.tipo]

    # ---------- secciones ----------
    def caratula(self):
        for _ in range(3):
            self.doc.add_paragraph()
        self.p(self.d["entidad"].upper(), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=16)
        self.p("Área Usuaria / Unidad Orgánica Requirente", italic=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, size=11)
        for _ in range(3):
            self.doc.add_paragraph()
        par = self.p("TÉRMINOS DE REFERENCIA", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=26)
        par.runs[0].font.color.rgb = AZUL_OSCURO
        self.p(f"CONTRATACIÓN DE: {self.d['objeto'].upper()}", bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, size=13)
        self.doc.add_paragraph()
        self.p(f"“{self.d['proyecto'].upper()}”", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=13)
        for _ in range(3):
            self.doc.add_paragraph()
        filas = [
            ("Especialidad", self.esp),
            ("Ubicación", self.d.get("ubicacion") or "-"),
            ("Código Único de Inversión (CUI)", self.d.get("cui") or "-"),
            ("Valor referencial / estimado", f"{soles(self.monto)} (incluye impuestos de ley)"),
            ("Plazo de ejecución", f"{self.plazo} días calendario"),
            ("Sistema de contratación", self.d["sistema"]),
            ("Procedimiento de selección", nombre_procedimiento(self.categoria, self.d["objeto"])),
            ("Marco normativo", self.marco["corto"]),
        ]
        self.tabla(["Dato", "Descripción"], filas, anchos=[6, 9.5], tam=10)
        self.doc.add_paragraph()
        fecha = datetime.now()
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                 "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        self.p(f"{(self.d.get('ubicacion') or 'Perú').split(',')[0].strip()}, "
               f"{meses[fecha.month - 1]} de {fecha.year}", align=WD_ALIGN_PARAGRAPH.CENTER, size=11)
        self.salto()

    def indice(self):
        self.p("CONTENIDO", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=13)
        items = [
            "1.0 GENERALIDADES", "    1.1 Denominación de la contratación", "    1.2 Entidad convocante",
            "    1.3 Finalidad pública", "    1.4 Antecedentes", "    1.5 Base legal",
            "    1.6 Valor referencial y procedimiento de selección",
            "2.0 OBJETIVO DE LA CONTRATACIÓN", "3.0 ALCANCE Y DESCRIPCIÓN",
            "    3.1 Alcance general", "    3.2 Actividades técnicas específicas", "    3.3 Actividades comunes y de gestión",
            "    3.4 Obligaciones del contratista", "    3.5 Obligaciones de la Entidad",
            "    3.6 Equipamiento estratégico mínimo", "    3.7 Personal de apoyo (no clave)",
            "    3.8 Identificación y asignación de riesgos",
            "4.0 METODOLOGÍA DE TRABAJO Y NORMAS TÉCNICAS APLICABLES",
            "5.0 ENTREGABLES Y FORMA DE PAGO",
            "6.0 REQUISITOS DE CALIFICACIÓN DEL PROVEEDOR Y DEL PERSONAL CLAVE",
            "7.0 OTRAS CONDICIONES", "8.0 SISTEMA DE CONTRATACIÓN Y FÓRMULA DE PENALIDADES",
            "ANEXO 01: Cuadro resumen de requisitos del personal clave",
            "ANEXO 02: Cronograma de entregables y pagos", "ANEXO 03: Glosario de términos",
            "ANEXO 04: Modelo de carta de compromiso del personal clave",
        ]
        for it in items:
            par = self.doc.add_paragraph()
            par.paragraph_format.space_after = Pt(3)
            r = par.add_run(it)
            r.bold = not it.startswith("    ")
        self.salto()

    def cap1(self):
        d = self.d
        self.h1("1.0 GENERALIDADES")
        self.h2("1.1 Denominación de la contratación")
        self.p_mixto([("Contratación de ", False), (d["objeto"].upper(), True), (" para el proyecto: ", False),
                      (f"“{d['proyecto']}”", True),
                      (f", ubicado en {d.get('ubicacion') or 'el ámbito de la Entidad'}"
                       f"{', con Código Único de Inversión N° ' + d['cui'] if d.get('cui') else ''}.", False)])
        self.h2("1.2 Entidad convocante")
        self.tabla(["Concepto", "Detalle"], [
            ("Nombre de la Entidad", d["entidad"]),
            ("Área usuaria", "Gerencia / Subgerencia de Infraestructura, Obras o Desarrollo Urbano (o la que haga sus veces)"),
            ("Órgano encargado de las contrataciones", "Unidad / Subgerencia de Logística y Abastecimiento"),
            ("Ubicación del proyecto", d.get("ubicacion") or "-"),
            ("Fuente de financiamiento", "Recursos determinados / ordinarios / transferencias (según certificación presupuestal)"),
        ], anchos=[6, 9.5])
        self.h2("1.3 Finalidad pública")
        finalidad = {
            "Vial": ("mejorar las condiciones de transitabilidad vehicular y peatonal, reducir los tiempos y costos de "
                     "viaje y de operación vehicular, y facilitar el acceso de la población a los servicios básicos de "
                     "salud, educación y a los mercados, contribuyendo a la integración y al desarrollo socioeconómico "
                     "de la zona de influencia"),
            "Edificaciones": ("brindar a la población usuaria una infraestructura adecuada, segura, accesible y "
                              "funcional, que cumpla con la normativa técnica vigente y permita la prestación de "
                              "servicios públicos de calidad en condiciones óptimas"),
            "Saneamiento": ("garantizar a la población beneficiaria el acceso a servicios de agua potable y "
                            "saneamiento de calidad, en cantidad, continuidad y presión adecuadas, reduciendo la "
                            "incidencia de enfermedades de origen hídrico y mejorando la calidad de vida"),
            "Saneamiento/Riego": ("incrementar la disponibilidad y eficiencia en el uso del recurso hídrico, ampliar la "
                                  "frontera agrícola bajo riego, mejorar la producción y productividad agropecuaria y "
                                  "asegurar el abastecimiento de agua para la población beneficiaria"),
        }[self.esp]
        self.p(f"La presente contratación tiene como finalidad pública {finalidad}. La contratación se enmarca en "
               f"los objetivos estratégicos y en el Plan Operativo Institucional de {d['entidad']}, así como en las "
               f"políticas nacionales de cierre de brechas de infraestructura y acceso a servicios públicos.")
        self.p("Asimismo, la contratación responde a la necesidad de la Entidad de contar con un proveedor "
               "especializado que garantice la calidad técnica, el cumplimiento de los plazos y el uso eficiente de "
               "los recursos públicos, en concordancia con los principios de eficacia y eficiencia, transparencia, "
               "integridad, competencia, valor por dinero y sostenibilidad que rigen las contrataciones públicas.")
        self.h2("1.4 Antecedentes")
        self.p(f"{d['entidad']}, en el marco de sus competencias y funciones, ha identificado la necesidad de "
               f"intervenir mediante el proyecto “{d['proyecto']}”, el cual cuenta con la viabilidad o aprobación "
               f"correspondiente en el marco del Sistema Nacional de Programación Multianual y Gestión de "
               f"Inversiones (Invierte.pe){' con CUI N° ' + d['cui'] if d.get('cui') else ''}.")
        ant = {
            "expediente": ("Para continuar con el ciclo de inversión, resulta necesario elaborar el expediente técnico "
                           "a nivel de ejecución de obra, el cual debe desarrollar a detalle la concepción técnica, "
                           "el dimensionamiento y el costo aprobados en la fase de formulación y evaluación, "
                           "debiendo mantener la consistencia con dicha concepción técnica."),
            "obra": ("El expediente técnico de obra ha sido elaborado y aprobado mediante el acto resolutivo "
                     "correspondiente, contando con la disponibilidad presupuestal y la libre disponibilidad del "
                     "terreno, por lo que corresponde contratar la ejecución de la obra."),
            "supervision": ("La Entidad viene convocando la contratación de la ejecución de la obra, por lo que, de "
                            "conformidad con la normativa de contrataciones públicas, cuando el valor de la obra "
                            "supera el monto establecido en la Ley de Presupuesto del Sector Público, resulta "
                            "obligatorio contratar la supervisión de obra, la cual se encargará del control técnico, "
                            "económico y administrativo de la ejecución."),
            "servicio": ("Para el cumplimiento de sus objetivos institucionales, el área usuaria requiere la "
                         "contratación de un servicio especializado que no puede ser atendido con personal propio, "
                         "por lo que se ha formulado el presente requerimiento."),
        }[self.tipo]
        self.p(ant)
        self.p("El área usuaria ha elaborado los presentes Términos de Referencia considerando la indagación de "
               "mercado, las Bases Estándar aprobadas por el organismo técnico especializado y la normativa técnica "
               "vigente, a fin de determinar las condiciones mínimas que debe cumplir el proveedor.")
        self.h2("1.5 Base legal")
        base = [
            self.marco["ley"] + ".",
            self.marco["reglamento"] + ".",
            "Directivas, pronunciamientos y Bases Estándar emitidas por el " + self.marco["organismo_largo"] + ".",
            "Ley de Presupuesto del Sector Público para el Año Fiscal vigente.",
            "Decreto Legislativo N° 1252, que crea el Sistema Nacional de Programación Multianual y Gestión de Inversiones, y su Reglamento.",
            "Ley N° 27444, Ley del Procedimiento Administrativo General (TUO aprobado por D.S. N° 004-2019-JUS).",
            "Código Civil, de aplicación supletoria.",
            "Reglamento Nacional de Edificaciones (RNE) y normas técnicas sectoriales aplicables.",
            "Ley N° 29783, Ley de Seguridad y Salud en el Trabajo, y su Reglamento.",
            "Ley N° 27446, Ley del Sistema Nacional de Evaluación del Impacto Ambiental, y su Reglamento.",
        ]
        for b in base:
            self.viñeta(b)
        self.p("Las referidas normas incluyen sus respectivas modificaciones, ampliaciones y normas conexas, de ser el caso.",
               italic=True)
        self.h2("1.6 Valor referencial y procedimiento de selección")
        self.p_mixto([("El valor referencial / estimado asciende a ", False), (soles(self.monto), True),
                      (f" ({numero_a_letras(self.monto)}), ", False),
                      ("incluidos los impuestos de ley y cualquier otro concepto que incida en el costo total "
                       f"de {self._sustantivo()}. El monto ha sido determinado sobre la base de la indagación de "
                       "mercado y/o el presupuesto del expediente técnico, según corresponda.", False)])
        sub = self.monto / 1.18
        self.tabla(["Concepto", "Monto (S/)"], [
            ("Subtotal (sin IGV)", f"{sub:,.2f}"),
            ("IGV (18 %)", f"{self.monto - sub:,.2f}"),
            ("VALOR REFERENCIAL TOTAL", f"{self.monto:,.2f}"),
        ], anchos=[10, 5.5], titulo="Composición del valor referencial", negrita_ultima=True)
        self.p_mixto([("De acuerdo con el monto y el objeto de la contratación, corresponde convocar el "
                       "procedimiento de selección de ", False),
                      (nombre_procedimiento(self.categoria, d["objeto"]), True),
                      (f" (categoría de exigencia: {self.categoria}), conforme a los umbrales establecidos en la "
                       "Ley de Presupuesto del Sector Público vigente.", False)])

    def cap2(self):
        d = self.d
        self.h1("2.0 OBJETIVO DE LA CONTRATACIÓN")
        self.h2("2.1 Objetivo general")
        obj = {
            "expediente": f"Contratar a una persona natural o jurídica que elabore el expediente técnico del proyecto “{d['proyecto']}”, a nivel de ejecución de obra, que permita a la Entidad contar con un documento técnico completo, consistente y conforme a la normativa vigente para la posterior contratación de la ejecución de la obra.",
            "obra": f"Contratar a una persona natural o jurídica para la ejecución de la obra “{d['proyecto']}”, conforme al expediente técnico aprobado, con la calidad requerida y dentro del plazo y costo establecidos.",
            "supervision": f"Contratar a una persona natural o jurídica para la supervisión de la ejecución de la obra “{d['proyecto']}”, que controle técnica, económica y administrativamente su ejecución, velando directa y permanentemente por la correcta ejecución y el cumplimiento del contrato de obra.",
            "servicio": f"Contratar a una persona natural o jurídica para la prestación del servicio vinculado al proyecto “{d['proyecto']}”, conforme a las condiciones establecidas en los presentes términos de referencia.",
        }[self.tipo]
        self.p(obj)
        self.h2("2.2 Objetivos específicos")
        especificos = {
            "expediente": [
                "Elaborar los estudios básicos de ingeniería (topografía, mecánica de suelos, hidrología, geología y otros) necesarios para el diseño.",
                "Desarrollar el diseño de ingeniería de detalle de todas las especialidades involucradas, conforme a la normativa técnica vigente.",
                "Determinar el presupuesto de obra con metrados sustentados y precios unitarios actualizados a la zona del proyecto.",
                "Elaborar las especificaciones técnicas, cronogramas, fórmulas polinómicas y planos de ejecución.",
                "Obtener las autorizaciones, certificaciones y aprobaciones requeridas (CIRA, instrumento de gestión ambiental, factibilidades, disponibilidad hídrica, según corresponda).",
                "Mantener la consistencia del expediente técnico con la concepción técnica y el dimensionamiento aprobados en Invierte.pe.",
            ],
            "obra": [
                "Ejecutar todas las partidas del expediente técnico, en las cantidades y calidades especificadas.",
                "Cumplir el calendario de avance de obra valorizado y el plazo de ejecución contractual.",
                "Garantizar la calidad de los materiales, equipos y procesos constructivos mediante el plan de control de calidad.",
                "Implementar las medidas de seguridad y salud en el trabajo y de gestión ambiental durante toda la ejecución.",
                "Entregar la obra culminada, operativa y con la documentación de cierre (planos de replanteo, manuales, garantías y liquidación).",
            ],
            "supervision": [
                "Controlar que la obra se ejecute conforme al expediente técnico, las especificaciones técnicas y la normativa aplicable.",
                "Controlar el cumplimiento del plazo de ejecución, del calendario de avance y del presupuesto contratado.",
                "Revisar y aprobar las valorizaciones, y emitir opinión técnica sobre adicionales, deductivos y ampliaciones de plazo.",
                "Verificar el cumplimiento de las obligaciones de seguridad, salud ocupacional y medio ambiente del contratista.",
                "Participar en la recepción de la obra y revisar o elaborar la liquidación del contrato de obra.",
            ],
            "servicio": [
                "Ejecutar las actividades del servicio conforme a las condiciones técnicas exigidas.",
                "Cumplir con los plazos y entregables establecidos.",
                "Garantizar la calidad del servicio y el cumplimiento de la normativa de seguridad y salud.",
                "Documentar las actividades ejecutadas con registros verificables y evidencias fotográficas.",
                "Transferir a la Entidad la información, archivos y conocimientos generados durante el servicio.",
            ],
        }[self.tipo]
        for e in especificos:
            self.viñeta(e)

    def cap3(self):
        self.h1("3.0 ALCANCE Y DESCRIPCIÓN DEL SERVICIO / OBRA")
        self.h2("3.1 Alcance general")
        self.p(f"{self._contratista()} será responsable de la ejecución integral de {self._sustantivo()} de la "
               f"especialidad de {self.esp}, comprendiendo todas las actividades, recursos, personal, equipos, "
               f"materiales, insumos, ensayos, gestiones y documentación necesarios para cumplir con el objeto de la "
               f"contratación, aun cuando no estén expresamente descritos en los presentes términos de referencia, "
               f"siempre que resulten indispensables para la correcta ejecución de la prestación.")
        self.p(f"Se consideran obras o prestaciones similares a las de {SIMILARES[self.esp]}.")
        self.h2("3.2 Actividades técnicas específicas de la especialidad")
        if self.tipo == "expediente":
            self.p(f"{self._contratista()} desarrollará, como mínimo, los siguientes estudios y diseños de la especialidad de {self.esp}:")
            actividades = ACT_DISENO[self.esp]
        elif self.tipo == "supervision":
            self.p(f"{self._contratista()} controlará, verificará y aprobará que el ejecutor de obra realice correctamente, "
                   f"como mínimo, los siguientes trabajos de la especialidad de {self.esp}, exigiendo los ensayos y "
                   f"protocolos de calidad correspondientes:")
            actividades = ACT_EJECUCION[self.esp]
        elif self.tipo == "obra":
            self.p(f"{self._contratista()} ejecutará, como mínimo, los siguientes trabajos de la especialidad de {self.esp}, "
                   f"conforme a las partidas del expediente técnico aprobado:")
            actividades = ACT_EJECUCION[self.esp]
        else:
            self.p(f"{self._contratista()} desarrollará las siguientes actividades vinculadas a la especialidad de {self.esp}:")
            actividades = ACT_EJECUCION[self.esp]
        for i, (tit, desc) in enumerate(actividades, start=1):
            self.h3(f"3.2.{i} {tit}")
            self.p(desc)
            self.p(self._complemento_actividad(tit))
        self.h2("3.3 Actividades comunes y de gestión")
        for i, (tit, desc) in enumerate(ACT_COMUNES[self.tipo]):
            self.literal(chr(97 + i) if i < 26 else str(i), tit, desc)
        self.h2("3.4 Obligaciones del contratista")
        oblig = [
            "Ejecutar la prestación con el personal clave ofertado, el cual no podrá ser sustituido sin la autorización previa de la Entidad, salvo los supuestos previstos en la normativa, debiendo el reemplazante reunir iguales o superiores calificaciones.",
            "Cumplir con la normativa técnica, de seguridad y salud en el trabajo, ambiental y laboral vigente, siendo responsable de los daños que se ocasionen a terceros.",
            "Contratar el Seguro Complementario de Trabajo de Riesgo (SCTR) de salud y pensión para todo el personal que participe en campo.",
            "Asumir los costos de movilización, alimentación, alojamiento, transporte, equipos, materiales, ensayos y demás gastos necesarios para la ejecución de la prestación.",
            "Mantener comunicación permanente con la Entidad y asistir a las reuniones de coordinación que sean convocadas.",
            "Absolver las consultas y observaciones formuladas por la Entidad en los plazos establecidos, sin costo adicional.",
            "Guardar reserva de la información a la que tenga acceso con ocasión de la ejecución contractual.",
            "Responder por la calidad ofrecida y por los vicios ocultos de la prestación durante el plazo establecido en el contrato.",
        ]
        for o in oblig:
            self.viñeta(o)
        self.h2("3.5 Obligaciones de la Entidad")
        for o in [
            "Proporcionar la información técnica disponible (estudios previos, expediente técnico, documentos de gestión) necesaria para la ejecución de la prestación.",
            "Designar un coordinador o inspector como contraparte técnica para el seguimiento del contrato.",
            "Revisar los entregables y otorgar la conformidad dentro de los plazos establecidos en los presentes términos de referencia.",
            "Efectuar los pagos conforme a la forma de pago pactada, previa conformidad del área usuaria.",
            "Brindar las facilidades administrativas para las gestiones ante otras entidades públicas.",
        ]:
            self.viñeta(o)
        self.h2("3.6 Equipamiento estratégico mínimo")
        equipos = EQUIPAMIENTO["obra"][self.esp] if self.tipo == "obra" else EQUIPAMIENTO[self.tipo]
        self.tabla(["N°", "Descripción del equipo", "Cantidad mínima"],
                   [(i + 1, e, c) for i, (e, c) in enumerate(equipos)],
                   anchos=[1.2, 11.3, 3], titulo="Equipamiento estratégico mínimo")
        self.p("El equipamiento se acreditará con copia de documentos que sustenten la propiedad, la posesión, el "
               "compromiso de compra venta o alquiler u otro documento que acredite la disponibilidad del equipamiento "
               "estratégico requerido, para la suscripción del contrato.", italic=True)
        self.h2("3.7 Personal de apoyo (no clave)")
        self.p("Adicionalmente al personal clave, el contratista dispondrá del siguiente personal de apoyo, cuya "
               "calificación no es materia de acreditación para la suscripción del contrato, pero cuya participación "
               "será verificada por la Entidad durante la ejecución:")
        apoyo = {
            "obra": [("Ingeniero asistente de residencia", 1, "Ingeniero Civil con experiencia en obras similares"),
                     ("Maestro de obra", 1, "Técnico en construcción civil con experiencia en obras similares"),
                     ("Topógrafo", 1, "Técnico en topografía"), ("Técnico laboratorista", 1, "Técnico en ensayos de suelos y concreto"),
                     ("Administrador de obra", 1, "Profesional o técnico en administración o contabilidad"),
                     ("Almacenero", 1, "Técnico con experiencia en control de almacenes de obra")],
            "supervision": [("Ingeniero asistente de supervisión", 1, "Ingeniero Civil con experiencia en supervisión de obras"),
                            ("Topógrafo", 1, "Técnico en topografía"), ("Técnico laboratorista", 1, "Técnico en ensayos de suelos y concreto"),
                            ("Asistente administrativo", 1, "Técnico en administración o secretariado")],
            "expediente": [("Asistente de ingeniería / CAD", 2, "Bachiller o técnico en ingeniería con dominio de AutoCAD/Civil 3D"),
                           ("Topógrafo y ayudantes", 1, "Técnico en topografía con equipo propio"),
                           ("Técnico laboratorista", 1, "Técnico en ensayos de mecánica de suelos"),
                           ("Asistente administrativo", 1, "Técnico en administración")],
            "servicio": [("Técnico de campo", 2, "Técnico con experiencia en servicios similares"),
                         ("Operarios / ayudantes", 4, "Personal con experiencia en las actividades del servicio"),
                         ("Asistente administrativo", 1, "Técnico en administración")],
        }[self.tipo]
        self.tabla(["N°", "Cargo", "Cantidad", "Perfil referencial"],
                   [(i + 1, c, n, pf) for i, (c, n, pf) in enumerate(apoyo)],
                   anchos=[0.8, 5.0, 1.8, 7.9], titulo="Personal de apoyo referencial")
        self.h2("3.8 Identificación y asignación de riesgos")
        self.p("De conformidad con la normativa de contrataciones públicas, se identifican los riesgos previsibles de "
               "la ejecución contractual, su valoración cualitativa y la respuesta prevista. El contratista deberá "
               "actualizar esta matriz en su plan de trabajo y reportar su seguimiento en cada informe.")
        self.tabla(["N°", "Riesgo identificado", "Prob.", "Impacto", "Respuesta / Mitigación", "Asignado a"],
                   [(i + 1, *r) for i, r in enumerate(self._riesgos())],
                   anchos=[0.8, 4.2, 1.4, 1.5, 5.8, 2.0], titulo="Matriz de riesgos de la contratación", tam=8.5)

    def _riesgos(self):
        comunes = [
            ("Lluvias intensas o eventos climáticos extraordinarios que paralicen los trabajos de campo", "Media", "Alto",
             "Programar actividades críticas fuera de la temporada de lluvias, prever holguras y protección de frentes de trabajo", "Compartido"),
            ("Conflictos sociales o falta de aceptación de la población beneficiaria", "Media", "Alto",
             "Plan de relaciones comunitarias, reuniones informativas y actas de compromiso con autoridades locales", "Entidad"),
            ("Ausencia o rotación del personal clave ofertado", "Baja", "Alto",
             "Exigir permanencia, aplicar penalidades y aprobar reemplazos con calificación igual o superior", "Contratista"),
            ("Accidentes de trabajo y enfermedades ocupacionales", "Media", "Alto",
             "Implementar el Plan de SST, matriz IPERC, charlas diarias, EPP y SCTR del personal", "Contratista"),
            ("Retraso en la revisión de entregables o en la absolución de consultas por la Entidad", "Media", "Medio",
             "Designar coordinador exclusivo, cumplir plazos de revisión y reuniones quincenales de seguimiento", "Entidad"),
        ]
        especificos = {
            "expediente": [
                ("Información de base insuficiente o desactualizada (estudio de preinversión, catastro, topografía)", "Alta", "Medio",
                 "Validar la información en el diagnóstico inicial y complementar con trabajos de campo propios", "Contratista"),
                ("Falta de libre disponibilidad de terrenos o demoras en la obtención del CIRA y autorizaciones", "Media", "Alto",
                 "Iniciar gestiones desde el primer entregable y coordinar con la Entidad las actas de libre disponibilidad", "Entidad"),
                ("Variaciones sustanciales respecto a la concepción técnica aprobada en Invierte.pe", "Media", "Alto",
                 "Informar oportunamente a la Unidad Formuladora y registrar las modificaciones en el Formato correspondiente", "Compartido"),
            ],
            "obra": [
                ("Deficiencias o incompatibilidades del expediente técnico", "Media", "Alto",
                 "Revisión del expediente al inicio, consultas oportunas en el cuaderno de obra y absolución por el proyectista", "Entidad"),
                ("Desabastecimiento o incremento de precios de materiales e insumos", "Media", "Medio",
                 "Calendario de adquisición de materiales, adelanto para materiales y reajuste por fórmulas polinómicas", "Compartido"),
                ("Interferencias con redes de servicios existentes no previstas", "Media", "Medio",
                 "Catastro de interferencias previo, coordinación con las empresas prestadoras y protocolos de reubicación", "Entidad"),
            ],
            "supervision": [
                ("Atrasos injustificados del ejecutor de obra", "Media", "Alto",
                 "Control semanal del avance, alertas tempranas, exigencia de calendario acelerado e informe a la Entidad", "Contratista"),
                ("Ampliaciones de plazo de la obra que extiendan el servicio de supervisión", "Media", "Medio",
                 "Pronunciamiento oportuno y sustentado; gestión de la prestación adicional de supervisión conforme a ley", "Entidad"),
                ("Deficiencias de calidad no detectadas oportunamente", "Baja", "Alto",
                 "Ensayos de contraste propios, protocolos por partida y presencia permanente del personal clave en obra", "Contratista"),
            ],
            "servicio": [
                ("Indefinición de alcances o cambios en las necesidades del área usuaria", "Media", "Medio",
                 "Validar el plan de trabajo al inicio y formalizar cualquier variación conforme a la normativa", "Entidad"),
                ("Incumplimiento de estándares de calidad del servicio", "Baja", "Medio",
                 "Protocolos de control, registros de conformidad parcial y penalidades", "Contratista"),
            ],
        }[self.tipo]
        return especificos + comunes

    def _complemento_actividad(self, titulo):
        if self.tipo == "expediente":
            return (f"Los resultados de la actividad “{titulo}” se presentarán en un informe técnico con memoria "
                    "descriptiva, memoria de cálculo, panel fotográfico, registros de campo, certificados de "
                    "laboratorio y/o calibración de equipos, y planos a escala adecuada, firmados por el especialista "
                    "responsable y visados por el jefe de proyecto. Las metodologías empleadas deberán sustentarse en la "
                    "normativa técnica vigente indicada en el capítulo 4.0.")
        if self.tipo == "supervision":
            return (f"Para la actividad “{titulo}”, la supervisión verificará los protocolos de calidad, la frecuencia de "
                    "ensayos, los certificados de materiales y los metrados ejecutados, registrando en el cuaderno de obra "
                    "sus aprobaciones u observaciones, e informando a la Entidad en el informe mensual sobre su estado.")
        if self.tipo == "obra":
            return (f"Para la actividad “{titulo}”, el contratista elaborará los protocolos de control de calidad por partida, "
                    "ejecutará los ensayos con la frecuencia indicada en las especificaciones técnicas y solicitará la "
                    "aprobación del supervisor o inspector antes de continuar con las partidas subsecuentes.")
        return f"La actividad “{titulo}” será registrada con evidencias y reportada en los informes del servicio."

    def cap4(self):
        self.h1("4.0 METODOLOGÍA DE TRABAJO Y NORMAS TÉCNICAS APLICABLES")
        self.h2("4.1 Metodología de trabajo")
        fases = {
            "expediente": [
                ("Fase I - Planificación", "Revisión del estudio de preinversión, reunión de inicio con la Entidad, presentación del plan de trabajo y programación detallada de actividades."),
                ("Fase II - Trabajos de campo", "Levantamiento topográfico, exploración geotécnica, aforos, inventarios, estudio de tráfico, encuestas y coordinación con la población y autoridades."),
                ("Fase III - Trabajos de gabinete", "Procesamiento de información, ensayos de laboratorio, análisis, modelamientos y diseño de ingeniería de detalle por especialidad."),
                ("Fase IV - Integración y compatibilización", "Compatibilización interdisciplinaria de planos (de preferencia en entorno BIM), elaboración de metrados, costos, presupuestos y cronogramas."),
                ("Fase V - Revisión y aprobación", "Presentación de entregables, levantamiento de observaciones y obtención de la conformidad del expediente técnico."),
            ],
            "supervision": [
                ("Fase I - Actividades previas", "Revisión del expediente técnico, informe de compatibilidad, verificación de la entrega de terreno y de los requisitos para el inicio del plazo de obra."),
                ("Fase II - Supervisión de la ejecución", "Control técnico, de calidad, de plazos, económico, de seguridad y ambiental en forma directa y permanente durante toda la ejecución de la obra."),
                ("Fase III - Recepción", "Verificación de la culminación de la obra, participación en el comité de recepción y control del levantamiento de observaciones."),
                ("Fase IV - Liquidación", "Revisión o elaboración de la liquidación del contrato de obra y presentación del informe final."),
            ],
            "obra": [
                ("Fase I - Movilización e inicio", "Movilización de personal y equipos, instalaciones provisionales, replanteo, actualización de la programación y apertura del cuaderno de obra."),
                ("Fase II - Ejecución", "Ejecución de las partidas conforme a la programación CPM, aplicando el Plan de Calidad, el Plan de SST y el instrumento de gestión ambiental."),
                ("Fase III - Pruebas y puesta en marcha", "Pruebas de funcionamiento, protocolos finales, capacitación y entrega de manuales de operación y mantenimiento."),
                ("Fase IV - Recepción y liquidación", "Recepción de la obra, levantamiento de observaciones, planos de replanteo y liquidación del contrato."),
            ],
            "servicio": [
                ("Fase I - Planificación", "Plan de trabajo, organización del personal y coordinación con la Entidad."),
                ("Fase II - Ejecución", "Desarrollo de las actividades del servicio con control de calidad permanente."),
                ("Fase III - Cierre", "Informe final y conformidad del servicio."),
            ],
        }[self.tipo]
        for i, (f, desc) in enumerate(fases):
            self.literal(chr(97 + i), f, desc)
        self.p("La metodología específica propuesta por el postor deberá desarrollar, como mínimo, estas fases, "
               "identificando los riesgos de la prestación, las medidas de mitigación y los mecanismos de control de "
               "calidad, y podrá ser objeto de evaluación como factor de evaluación cuando así se establezca en las bases.")
        self.h2("4.2 Normas técnicas aplicables")
        self.p("La prestación se desarrollará en estricto cumplimiento de las siguientes normas técnicas, sin carácter "
               "limitativo, en sus versiones vigentes:")
        normas = NORMAS_ESPECIALIDAD[self.esp] + NORMAS_GENERALES
        self.tabla(["N°", "Norma / Código", "Denominación"],
                   [(i + 1, c, n) for i, (c, n) in enumerate(normas)],
                   anchos=[1.0, 4.5, 10.0], titulo=f"Normas técnicas aplicables - Especialidad {self.esp}")
        self.p("En caso de discrepancia entre normas, prevalecerá la de mayor exigencia técnica, previa coordinación con "
               "la Entidad. Cualquier norma que sea modificada o sustituida durante la ejecución deberá aplicarse en lo "
               "que corresponda, conforme a su régimen de aplicación en el tiempo.")
        self.h2("4.3 Herramientas informáticas")
        self.p(f"Se utilizarán herramientas informáticas con licencia vigente, tales como: {SOFTWARE_ESPECIALIDAD[self.esp]}. "
               "Los archivos se entregarán en formato nativo editable y en PDF.")
        self.h2("4.4 Aseguramiento y control de calidad")
        self.p("El contratista implementará un sistema de aseguramiento y control de calidad que comprenda: "
               "procedimientos de trabajo, listas de verificación, control de documentos, calibración de equipos, "
               "trazabilidad de ensayos y registro de no conformidades con sus acciones correctivas. Todo ensayo deberá "
               "realizarse en laboratorios que cuenten con equipos calibrados y, de preferencia, acreditados por INACAL.")
        self.h2("4.5 Coordinación y comunicaciones")
        self.p("Se realizará una reunión de inicio dentro de los cinco (5) primeros días del plazo contractual, con la "
               "participación del personal clave y del coordinador de la Entidad, en la que se validará el plan de "
               "trabajo, los canales de comunicación y el cronograma de entregables. Durante la ejecución se realizarán "
               "reuniones de seguimiento, como mínimo, quincenales, cuyas actas formarán parte del siguiente entregable.")
        self.p("Toda comunicación formal se efectuará por escrito, a través de la mesa de partes de la Entidad o del "
               "correo electrónico consignado en el contrato. Las coordinaciones técnicas no generan, por sí mismas, "
               "modificación alguna de las condiciones contractuales.")

    def cap5(self):
        self.h1("5.0 ENTREGABLES Y FORMA DE PAGO")
        self.h2("5.1 Plazo de ejecución")
        inicio = {"obra": "a partir del día siguiente de cumplidas las condiciones de inicio del plazo de ejecución de obra previstas en la normativa",
                  "supervision": "en concordancia con el plazo de ejecución de la obra, desde el inicio de esta hasta su recepción, más el periodo de liquidación del contrato de obra",
                  "expediente": "a partir del día siguiente de la suscripción del contrato o de la entrega de la información de base, lo que ocurra último",
                  "servicio": "a partir del día siguiente de la suscripción del contrato"}[self.tipo]
        self.p_mixto([("El plazo de ejecución es de ", False), (f"{self.plazo} días calendario", True),
                      (f", computado {inicio}. Dicho plazo no incluye los periodos de revisión de la Entidad ni de "
                       "levantamiento de observaciones, salvo indicación expresa.", False)])
        self.h2("5.2 Entregables")
        filas, total_pct, total_monto = [], 0.0, 0.0
        for i, e in enumerate(self.entregables, start=1):
            pct = a_float(e["porcentaje"])
            monto = self.monto * pct / 100
            total_pct += pct
            total_monto += monto
            filas.append((i, e["descripcion"], f"Hasta los {a_int(e['plazo'])} d.c.", f"{pct:g} %", f"{monto:,.2f}"))
        filas.append(("", "TOTAL", "", f"{total_pct:g} %", f"{total_monto:,.2f}"))
        self.tabla(["N°", "Entregable", "Plazo", "% Pago", "Monto (S/)"], filas,
                   anchos=[0.9, 8.3, 2.6, 1.6, 2.4], titulo="Entregables, plazos y porcentajes de pago",
                   negrita_ultima=True)
        if abs(total_pct - 100) > 0.01:
            self.p(f"Nota: la suma de porcentajes es {total_pct:g} %. Verificar antes de la convocatoria.", italic=True)
        self.h2("5.3 Contenido mínimo de los entregables")
        contenido = [
            "Carta de presentación dirigida a la Entidad, indicando el número de entregable y el contrato.",
            "Índice general, memoria descriptiva del avance y resumen ejecutivo.",
            "Desarrollo técnico de las actividades ejecutadas en el periodo, con memorias de cálculo, registros de campo y ensayos.",
            "Planos y anexos, firmados y sellados por los profesionales responsables.",
            "Panel fotográfico fechado y georreferenciado de las actividades de campo.",
            "Archivos digitales en formato nativo editable (DWG, XLSX, DOCX, S10, MPP, RVT u otros) y en PDF, en medio digital.",
        ]
        for c in contenido:
            self.viñeta(c)
        if self.tipo in ("obra", "supervision"):
            self.h3("5.3.1 Contenido mínimo del informe mensual")
            for i, (tit, desc) in enumerate([
                ("Resumen ejecutivo", "Avance físico y financiero programado versus ejecutado, situación del plazo y principales ocurrencias del periodo."),
                ("Valorización del periodo", "Metrados ejecutados, valorización mensual y acumulada, reajustes, amortización de adelantos y deducciones."),
                ("Control de calidad", "Relación de ensayos ejecutados, resultados, certificados de calidad de materiales y no conformidades con sus acciones correctivas."),
                ("Control de plazos", "Diagrama de Gantt actualizado, curva S programada versus real, análisis de la ruta crítica y causas de atraso, de existir."),
                ("Seguridad, salud y medio ambiente", "Indicadores de SST (horas-hombre, incidentes, accidentes), inspecciones, capacitaciones y cumplimiento de las medidas ambientales."),
                ("Personal y equipos", "Relación del personal clave y del equipo en obra, con la acreditación de su permanencia."),
                ("Documentación", "Copia de los asientos del cuaderno de obra del periodo, consultas y absoluciones, y panel fotográfico."),
                ("Conclusiones y recomendaciones", "Problemas identificados, riesgos emergentes y recomendaciones para la Entidad."),
            ]):
                self.literal(chr(97 + i), tit, desc)
        self.p("Los entregables se presentarán por mesa de partes física o virtual de la Entidad, en un (01) original "
               "y una (01) copia impresa, foliados y visados en todas sus páginas, salvo que la Entidad disponga su "
               "presentación únicamente digital con firma electrónica.")
        self.h2("5.4 Revisión, observaciones y conformidad")
        rev = a_int(self.params.get("revision_entidad_dias", 10))
        sub = a_int(self.params.get("subsanacion_dias", 5))
        self.p_mixto([("La Entidad revisará cada entregable en un plazo máximo de ", False), (f"{rev} días calendario", True),
                      (" contados desde su recepción. De existir observaciones, se comunicarán por escrito al "
                       "contratista, indicando claramente su sentido, otorgándole un plazo para subsanar no menor de ", False),
                      (f"{sub} días calendario", True),
                      (" ni mayor al que establezca el Reglamento, en función de la complejidad. Si pese al plazo "
                       "otorgado el contratista no cumpliese a cabalidad con la subsanación, la Entidad podrá resolver el "
                       "contrato, sin perjuicio de aplicar las penalidades que correspondan.", False)])
        self.p("La conformidad será otorgada por el área usuaria, previo informe del coordinador o inspector "
               "designado. La conformidad no enerva el derecho de la Entidad a reclamar posteriormente por defectos o "
               "vicios ocultos.")
        self.h2("5.5 Forma de pago")
        if self.tipo == "obra":
            self.p("El pago se realizará mediante valorizaciones mensuales, que tienen el carácter de pagos a cuenta, "
                   "elaboradas el último día de cada mes por el inspector o supervisor y el contratista, sobre la base de "
                   "los metrados realmente ejecutados y los precios unitarios ofertados (o el porcentaje de avance en "
                   "suma alzada), agregando los gastos generales, la utilidad y el IGV, con los reajustes por fórmulas "
                   "polinómicas y deduciendo las amortizaciones de adelantos y penalidades que correspondan.")
        else:
            self.p("El pago se efectuará en soles, en armadas que corresponden a los entregables señalados en el "
                   "numeral 5.2, previa conformidad de cada entregable, deduciendo la amortización proporcional de los "
                   "adelantos otorgados y las penalidades que correspondan.")
        self.p("Para efectos del pago, el contratista presentará: (i) comprobante de pago, (ii) informe de conformidad "
               "del área usuaria, (iii) entregable aprobado y (iv) constancia de pago del SCTR del periodo, de "
               "corresponder. La Entidad realizará el pago dentro de los plazos que establece la normativa, bajo "
               "responsabilidad del funcionario competente; en caso de retraso injustificado, el contratista tendrá "
               "derecho al pago de intereses legales.")

    def cap6(self):
        self.h1("6.0 REQUISITOS DE CALIFICACIÓN DEL PROVEEDOR Y DEL PERSONAL CLAVE")
        self.h2("6.1 Capacidad legal - Habilitación")
        rnp = {"obra": "Ejecutor de Obras, con capacidad máxima de contratación igual o mayor al valor referencial",
               "expediente": "Consultor de Obras, en la especialidad y categoría que corresponda",
               "supervision": "Consultor de Obras, en la especialidad y categoría que corresponda",
               "servicio": "Proveedor de Servicios"}[self.tipo]
        self.p_mixto([("El postor deberá contar con inscripción vigente en el Registro Nacional de Proveedores (RNP) "
                       "como ", False), (rnp, True),
                      (", y no encontrarse impedido, suspendido ni inhabilitado para contratar con el Estado.", False)])
        self.h2("6.2 Experiencia del postor en la especialidad")
        factor = a_float(self.params.get("factor_exp_postor", 1.0))
        if self.categoria == CAT_MENOR:
            factor = min(factor, 0.5)
        self.p_mixto([("El postor debe acreditar un monto facturado acumulado equivalente a ", False),
                      (f"{factor:g} vez(veces) el valor referencial ({soles(self.monto * factor)})", True),
                      (", por la contratación de servicios / obras iguales o similares al objeto de la convocatoria, "
                       "durante los diez (10) años anteriores a la fecha de presentación de ofertas que se computarán "
                       "desde la fecha de la conformidad o emisión del comprobante de pago, según corresponda.", False)])
        self.p(f"Se consideran similares a: {SIMILARES[self.esp]}.")
        self.p("Acreditación: copia simple de (i) contratos u órdenes de servicio y su respectiva conformidad o "
               "constancia de prestación; o (ii) comprobantes de pago cuya cancelación se acredite documental y "
               "fehacientemente, con voucher de depósito, nota de abono, reporte de estado de cuenta o cualquier otro "
               "documento emitido por entidad del sistema financiero que acredite el abono.", italic=True)
        self.h2("6.3 Calificaciones y experiencia del personal clave")
        self.p("El personal clave requerido para la ejecución de la prestación es el siguiente. Su calificación se "
               "acreditará para la suscripción del contrato, de conformidad con las Bases Estándar vigentes:")
        self.tabla(["Cargo", "Formación", "Exp. general", "Exp. específica", "Capacitación"],
                   [(r["cargo"], f"{r['formacion']} ({r['colegiatura']})", f"{r['exp_gen']} años",
                     f"{r['exp_esp']} años", f"{r['cap_h']} h") for r in self.personal],
                   anchos=[4.0, 5.0, 2.0, 2.2, 2.3], titulo="Resumen del personal clave requerido")
        funciones = {
            "Residente": "Dirigir técnicamente la obra en forma permanente y directa, representar al contratista ante la Entidad, anotar en el cuaderno de obra y suscribir las valorizaciones.",
            "Jefe de Supervisión": "Dirigir el equipo de supervisión, representar a la supervisión ante la Entidad y el contratista, y suscribir los informes y pronunciamientos técnicos.",
            "Jefe de Proyecto": "Coordinar, dirigir y compatibilizar el trabajo de los especialistas, responder por la calidad integral del expediente técnico y representar al consultor ante la Entidad.",
            "Coordinador": "Coordinar la ejecución del servicio, supervisar al personal y representar al contratista ante la Entidad.",
            "Seguridad": "Implementar, controlar y registrar el cumplimiento del Plan de Seguridad y Salud en el Trabajo, la matriz IPERC y los programas de capacitación y respuesta a emergencias.",
            "Ambiental": "Implementar y controlar las medidas del instrumento de gestión ambiental, el manejo de residuos y elaborar los informes de monitoreo ambiental.",
            "Costos": "Elaborar y controlar metrados, presupuestos, análisis de precios unitarios, valorizaciones, reajustes y liquidaciones.",
            "Estructuras": "Responsable del análisis, diseño y/o control de las estructuras, verificando su conformidad con la normativa sismorresistente y de concreto armado.",
            "Topografía": "Responsable de los levantamientos topográficos, georreferenciación, control geodésico y replanteos.",
        }
        for i, r in enumerate(self.personal, start=1):
            self.h3(f"6.3.{i} {r['cargo']}")
            self.literal("a", "Formación académica",
                         f"Título profesional de {r['formacion']}. Requisito de habilitación: colegiatura y habilitación "
                         f"vigente ({r['colegiatura']}), que se verificará al inicio de su participación efectiva en la ejecución.")
            self.literal("b", "Experiencia general",
                         f"Mínimo {r['exp_gen']} año(s) de experiencia en el ejercicio de su profesión, computada desde "
                         "la colegiatura.")
            self.literal("c", "Experiencia específica",
                         f"Mínimo {r['exp_esp']} año(s) como {r['exp_desc']}. Se consideran similares a: "
                         f"{self._corto(SIMILARES[self.esp], 400)}.")
            self.literal("d", "Capacitación",
                         f"Mínimo {r['cap_h']} horas lectivas en {r['cap_tema']}, acreditadas con constancias, "
                         "certificados o diplomas.")
            fun = next((v for k, v in funciones.items() if k.lower() in r["cargo"].lower()),
                       "Desarrollar las actividades de su especialidad, elaborar los informes técnicos respectivos y suscribir los documentos de su competencia.")
            self.literal("e", "Funciones principales", fun)
        self.h2("6.4 Acreditación del personal clave")
        self.p("La experiencia del personal clave se acreditará con cualquiera de los siguientes documentos: (i) copia "
               "simple de contratos y su respectiva conformidad, (ii) constancias, (iii) certificados o (iv) cualquier "
               "otra documentación que, de manera fehaciente, demuestre la experiencia del personal propuesto. En caso "
               "de periodos superpuestos, se considerará solo una vez el periodo traslapado. La formación académica se "
               "verificará en el Registro Nacional de Grados Académicos y Títulos Profesionales de la SUNEDU.")
        self.p("Los documentos que acrediten la experiencia deberán incluir los nombres del cargo o puesto, el objeto "
               "de la contratación, el periodo de ejecución y la entidad o empresa que emite el documento.")

    def cap7(self):
        self.h1("7.0 OTRAS CONDICIONES")
        self.h2("7.1 Confidencialidad")
        self.p(f"{self._contratista()} se obliga a mantener absoluta reserva y confidencialidad sobre la información "
               "que le sea proporcionada o a la que tenga acceso con ocasión de la ejecución del contrato, quedando "
               "prohibido revelarla, reproducirla o utilizarla para fines distintos, durante la vigencia del contrato y "
               "después de su culminación, salvo autorización expresa de la Entidad o mandato legal.")
        self.h2("7.2 Propiedad intelectual")
        self.p("La Entidad tendrá todos los derechos de propiedad intelectual, incluidos los derechos patrimoniales de "
               "autor, sobre los estudios, informes, planos, diseños, programas informáticos y demás documentos "
               "producidos como consecuencia de la ejecución contractual, pudiendo utilizarlos, modificarlos, "
               "reproducirlos y difundirlos libremente, sin pago adicional alguno.")
        self.h2("7.3 Subcontratación")
        self.p(f"{self._contratista()} podrá subcontratar parte de las prestaciones, previa autorización de la Entidad, "
               f"hasta un máximo del {a_float(self.params.get('subcontratacion_max_pct', 40)):g} % del monto del contrato "
               "vigente, salvo prohibición expresa en las bases. No podrán subcontratarse las prestaciones esenciales "
               "vinculadas a las condiciones que han sido objeto de evaluación. El contratista mantiene la "
               "responsabilidad por la ejecución total del contrato.")
        self.h2("7.4 Ítems, sublotes y paquetes")
        self.p("La presente contratación se convoca como un único ítem/paquete, por tratarse de prestaciones "
               "vinculadas a un mismo proyecto, cuya ejecución integral permite optimizar la coordinación técnica y la "
               "responsabilidad única del proveedor. No procede la contratación por sublotes.")
        self.h2("7.5 Garantías")
        g = self.gar
        self.tabla(["Garantía", "Porcentaje", "Monto (S/)", "Vigencia"], [
            ("Fiel cumplimiento", f"{g['gfc_pct']:g} % del monto del contrato", f"{g['gfc']:,.2f}",
             "Hasta la conformidad de la recepción de la prestación / consentimiento de la liquidación (obras y consultorías de obra)"),
            ("Por adelanto directo", "100 % del monto del adelanto", f"{g['ad']:,.2f}" if g["ad"] else "-",
             "Hasta la amortización total del adelanto"),
            ("Por adelanto de materiales", "100 % del monto del adelanto", f"{g['mat']:,.2f}" if g["mat"] else "-",
             "Hasta la utilización de los materiales"),
        ], anchos=[3.5, 3.5, 2.8, 5.7], titulo="Garantías del contrato")
        self.p("Las garantías serán incondicionales, solidarias, irrevocables y de realización automática en el país, "
               "al solo requerimiento de la Entidad, emitidas por empresas bajo supervisión de la SBS o mediante las "
               "modalidades alternativas que permita la normativa vigente (retención, póliza de caución, entre otras). "
               "Las MYPE podrán optar por la retención del 10 % del monto del contrato, cuando corresponda.")
        self.h2("7.6 Adelantos")
        if g["ad_pct"] > 0:
            self.p_mixto([("La Entidad podrá otorgar un adelanto directo de hasta el ", False), (f"{g['ad_pct']:g} %", True),
                          (f" del monto del contrato original ({soles(g['ad'])}), previa solicitud del contratista y "
                           "presentación de la garantía por adelanto y el comprobante de pago, dentro del plazo "
                           "establecido en la normativa.", False)])
        if g["mat_pct"] > 0:
            self.p_mixto([("Asimismo, se podrán otorgar adelantos para materiales e insumos por un monto total de hasta el ", False),
                          (f"{g['mat_pct']:g} %", True),
                          (f" del monto del contrato original ({soles(g['mat'])}), conforme al calendario de "
                           "adquisición de materiales presentado por el contratista.", False)])
        self.p("La amortización de los adelantos se realizará mediante descuentos proporcionales en cada uno de los "
               "pagos parciales o valorizaciones, de acuerdo con la normativa vigente.")
        self.h2("7.7 Responsabilidad por vicios ocultos")
        plazo_vo = {"obra": "siete (7) años", "expediente": "tres (3) años", "supervision": "tres (3) años",
                    "servicio": "un (1) año"}[self.tipo]
        self.p(f"La conformidad de la prestación no enerva el derecho de la Entidad a reclamar por defectos o vicios "
               f"ocultos. El plazo máximo de responsabilidad del contratista es de {plazo_vo}, contado a partir de la "
               f"conformidad otorgada por la Entidad o de la recepción, según corresponda.")
        self.h2("7.8 Seguros")
        self.p("El contratista deberá contar con el Seguro Complementario de Trabajo de Riesgo (SCTR) para todo su "
               "personal y, de corresponder, con pólizas de responsabilidad civil frente a terceros y de todo riesgo "
               "contratista (CAR) durante el plazo de ejecución.")
        self.h2("7.9 Cláusula anticorrupción y de integridad")
        self.p("El contratista declara y garantiza no haber ofrecido, negociado o efectuado, ni ofrecerá, negociará o "
               "efectuará, cualquier pago, objeto de valor o dádiva a funcionarios o servidores de la Entidad, "
               "comprometiéndose a conducirse con honestidad, probidad, veracidad e integridad, y a comunicar a las "
               "autoridades competentes cualquier acto ilícito o de corrupción del que tuviera conocimiento.")
        self.h2("7.10 Solución de controversias")
        self.p("Las controversias que surjan entre las partes durante la ejecución contractual se resolverán mediante "
               "conciliación, junta de prevención y resolución de disputas (cuando corresponda) y/o arbitraje, conforme "
               "a lo dispuesto en " + self.marco["corto"] + ".")

    def cap8(self):
        self.h1("8.0 SISTEMA DE CONTRATACIÓN Y FÓRMULA DE PENALIDADES")
        self.h2("8.1 Sistema de contratación")
        textos = {
            "Suma Alzada": ("La presente contratación se rige por el sistema de SUMA ALZADA, aplicable cuando las cantidades, "
                            "magnitudes y calidades de la prestación están totalmente definidas en los términos de referencia "
                            "o en el expediente técnico. El postor formula su oferta por un monto fijo integral y por un "
                            "determinado plazo de ejecución, asumiendo el riesgo de las variaciones de cantidades."),
            "Precios Unitarios": ("La presente contratación se rige por el sistema de PRECIOS UNITARIOS, aplicable cuando no "
                                  "puede conocerse con exactitud o precisión las cantidades o magnitudes requeridas. El postor "
                                  "formula su oferta proponiendo precios en función de las partidas o cantidades referenciales, "
                                  "y el pago se realiza en función de las cantidades realmente ejecutadas."),
            "Mixto": ("La presente contratación se rige por el sistema MIXTO (suma alzada y precios unitarios), aplicable "
                      "cuando la prestación contiene componentes cuyas cantidades están totalmente definidas (suma alzada) y "
                      "otros cuyas cantidades no pueden determinarse con precisión (precios unitarios). Las bases detallarán "
                      "los componentes sujetos a cada sistema."),
        }
        self.p(textos.get(self.d["sistema"], textos["Suma Alzada"]))
        self.h2("8.2 Penalidad por mora en la ejecución de la prestación")
        self.p("En caso de retraso injustificado del contratista en la ejecución de las prestaciones objeto del "
               "contrato, la Entidad le aplicará automáticamente una penalidad por mora por cada día de atraso, de "
               "acuerdo con la siguiente fórmula:")
        par = self.p("Penalidad diaria  =  (0.10  ×  Monto vigente)  /  (F  ×  Plazo vigente en días)", bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER, size=12)
        par.paragraph_format.space_before = Pt(6)
        par.paragraph_format.space_after = Pt(10)
        self.p("Donde F tiene los siguientes valores:")
        self.viñeta("Para plazos menores o iguales a sesenta (60) días, para bienes, servicios en general, consultorías y ejecución de obras: F = 0.40.")
        self.viñeta("Para plazos mayores a sesenta (60) días, para bienes, servicios en general y consultorías: F = 0.25.")
        self.viñeta("Para plazos mayores a sesenta (60) días, para obras: F = 0.15.")
        self.p("Tanto el monto como el plazo se refieren, según corresponda, al monto vigente del contrato o ítem que "
               "debió ejecutarse o, en caso de que estos involucren obligaciones de ejecución periódica o entregas "
               "parciales, a la prestación individual que fue materia de retraso.")
        pen = self.pen
        self.tabla(["Parámetro", "Valor"], [
            ("Monto vigente del contrato (referencial)", soles(self.monto)),
            ("Plazo vigente (días calendario)", f"{self.plazo}"),
            ("Factor F aplicable", f"{pen['F']:.2f}"),
            ("Penalidad diaria", soles(pen["diaria"])),
            ("Monto máximo de penalidad por mora (10 %)", soles(pen["tope"])),
            ("Días de atraso para alcanzar el tope", f"{pen['dias_tope']} días"),
        ], anchos=[9.5, 6], titulo="Cálculo de la penalidad por mora (con el valor referencial)")
        self.p("La penalidad por mora y las otras penalidades pueden alcanzar cada una un monto máximo equivalente al "
               "diez por ciento (10 %) del monto del contrato vigente o, de ser el caso, del ítem que debió ejecutarse. "
               "Cuando se alcance el monto máximo de penalidad, la Entidad podrá resolver el contrato por incumplimiento.")
        self.h2("8.3 Otras penalidades")
        uit = a_float(self.params.get("uit", UIT_DEFAULT))
        otras = [
            ("Ausencia injustificada del personal clave en la ejecución de la prestación", "0.5 UIT por cada día de ausencia del personal", f"{0.5 * uit:,.2f}", "Informe del coordinador / cuaderno de obra"),
            ("Sustitución del personal clave sin autorización de la Entidad", "1.0 UIT por cada ocurrencia", f"{1.0 * uit:,.2f}", "Informe del área usuaria"),
            ("Personal de campo sin EPP o sin SCTR vigente", "0.5 UIT por cada ocurrencia", f"{0.5 * uit:,.2f}", "Informe del especialista de seguridad / inspector"),
            ("No subsanar las observaciones dentro del plazo otorgado", "0.3 UIT por cada día de retraso", f"{0.3 * uit:,.2f}", "Informe del área usuaria"),
            ("Incumplimiento de las medidas del instrumento de gestión ambiental", "0.5 UIT por cada ocurrencia", f"{0.5 * uit:,.2f}", "Informe del especialista ambiental"),
        ]
        if self.tipo in ("obra", "supervision"):
            otras.append(("No mantener el cuaderno de obra al día o no anotar las ocurrencias relevantes", "0.3 UIT por cada ocurrencia", f"{0.3 * uit:,.2f}", "Informe del inspector / supervisor"))
        self.tabla(["N°", "Supuesto de aplicación", "Forma de cálculo", "Monto (S/)", "Procedimiento"],
                   [(i + 1, *o) for i, o in enumerate(otras)],
                   anchos=[0.8, 5.6, 3.5, 2.0, 3.6], titulo=f"Otras penalidades (UIT = {soles(uit)})")
        self.p("Estas penalidades son objetivas, razonables, congruentes y proporcionales con el objeto de la "
               "contratación, y se deducirán de los pagos a cuenta, del pago final o de la liquidación final; o, si "
               "fuese necesario, del monto resultante de la ejecución de la garantía de fiel cumplimiento.")

    def firmas(self):
        self.doc.add_paragraph()
        self.p("Elaborado por el área usuaria:", bold=True)
        t = self.doc.add_table(rows=1, cols=2)
        for i, cargo in enumerate(["Responsable del Área Usuaria", "Jefe de la Unidad Orgánica"]):
            c = t.rows[0].cells[i]
            c.text = ""
            par = c.paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            par.add_run("\n\n\n_______________________________\n")
            par.add_run(cargo).bold = True
            par.add_run(f"\n{self.d['entidad']}")

    def anexos(self):
        self.salto()
        self.h1("ANEXO 01: CUADRO RESUMEN DE REQUISITOS DEL PERSONAL CLAVE")
        self.tabla(["N°", "Cargo", "Formación / Colegiatura", "Exp. Gral.", "Exp. Espec.", "Capacitación", "Experiencia específica como"],
                   [(i + 1, r["cargo"], f"{r['formacion']} / {r['colegiatura']}", f"{r['exp_gen']} años",
                     f"{r['exp_esp']} años", f"{r['cap_h']} h - {r['cap_tema']}", r["exp_desc"])
                    for i, r in enumerate(self.personal)],
                   anchos=[0.7, 2.6, 2.8, 1.3, 1.3, 3.2, 3.6], tam=8)
        self.h1("ANEXO 02: CRONOGRAMA DE ENTREGABLES Y PAGOS")
        acum = 0.0
        filas = []
        for i, e in enumerate(self.entregables, start=1):
            pct = a_float(e["porcentaje"])
            acum += pct
            filas.append((i, self._corto(e["descripcion"], 90), a_int(e["plazo"]), f"{pct:g} %",
                          f"{acum:g} %", f"{self.monto * pct / 100:,.2f}", f"{self.monto * acum / 100:,.2f}"))
        self.tabla(["N°", "Entregable", "Día", "% parcial", "% acumulado", "Pago (S/)", "Acumulado (S/)"], filas,
                   anchos=[0.7, 6.2, 1.1, 1.6, 1.8, 2.1, 2.2], tam=8)
        self.h1("ANEXO 03: GLOSARIO DE TÉRMINOS")
        glosario = [
            ("Área usuaria", "Dependencia cuyas necesidades pretenden ser atendidas con la contratación o que, dada su especialidad y funciones, canaliza los requerimientos."),
            ("Bases Estándar", "Documentos del procedimiento de selección aprobados por el organismo técnico especializado, de uso obligatorio por las entidades."),
            ("Consultoría de obra", "Servicio profesional altamente calificado consistente en la elaboración del expediente técnico de obra o en la supervisión de obra."),
            ("Contratista", "Proveedor que celebra un contrato con una entidad de conformidad con la normativa de contrataciones públicas."),
            ("Cuaderno de obra", "Documento, de preferencia digital, en el que el residente y el inspector o supervisor anotan las ocurrencias, órdenes, consultas y respuestas durante la ejecución de la obra."),
            ("Expediente técnico de obra", "Conjunto de documentos que comprende memoria descriptiva, especificaciones técnicas, planos, metrados, presupuesto, valor referencial, análisis de precios, cronogramas, fórmulas polinómicas y estudios básicos."),
            ("Penalidad por mora", "Sanción económica aplicada automáticamente por cada día de atraso injustificado en la ejecución de las prestaciones."),
            ("Personal clave", "Profesional o técnico cuya participación es esencial para la ejecución de la prestación y cuya calificación se exige como requisito."),
            ("RNP", "Registro Nacional de Proveedores, en el que deben inscribirse los proveedores para participar en procedimientos de selección y contratar con el Estado."),
            ("SEACE", "Sistema Electrónico de Contrataciones del Estado, plataforma de registro y difusión de los procedimientos de contratación pública."),
            ("Valorización", "Cuantificación económica de un avance físico en la ejecución de la obra, realizada en un periodo determinado."),
            ("UIT", "Unidad Impositiva Tributaria, valor de referencia fijado anualmente por el Ministerio de Economía y Finanzas."),
        ]
        self.tabla(["Término", "Definición"], glosario, anchos=[3.5, 12])
        self.salto()
        self.h1("ANEXO 04: MODELO DE CARTA DE COMPROMISO DEL PERSONAL CLAVE")
        self.p("Señores:", bold=True)
        self.p(f"COMITÉ DE SELECCIÓN / ÓRGANO ENCARGADO DE LAS CONTRATACIONES\n{self.d['entidad']}\nPresente.-",
               align=WD_ALIGN_PARAGRAPH.LEFT)
        self.p(f"De mi consideración:")
        self.p("Yo, ………………………………………………………, identificado(a) con DNI N° ……………………, de profesión "
               "………………………………………, con registro de colegiatura N° ……………………, domiciliado(a) en "
               "……………………………………………………, declaro bajo juramento:")
        for t in [
            f"Que me comprometo a prestar mis servicios en el cargo de …………………………………… para ejecutar {self._sustantivo()} "
            f"“{self.d['proyecto']}”, en caso de que el postor ……………………………………… resulte favorecido con la buena pro "
            "y suscriba el contrato correspondiente.",
            f"Que mi participación será por el plazo de {self.plazo} días calendario o el que corresponda según la "
            "participación prevista, en forma permanente cuando así se exija en los términos de referencia.",
            "Que cuento con la formación académica, la experiencia y las capacitaciones exigidas para el cargo, cuya "
            "documentación sustentatoria adjunto a la presente.",
            "Que no me encuentro comprometido(a) con otro contrato que me impida participar en la ejecución de la "
            "prestación durante el plazo señalado.",
        ]:
            self.viñeta(t)
        self.p("Asimismo, declaro conocer que, en caso de incumplimiento de lo declarado, la Entidad podrá aplicar las "
               "penalidades previstas en los términos de referencia y comunicar los hechos al Tribunal competente.")
        self.tabla(["N°", "Experiencia (entidad / empresa)", "Cargo desempeñado", "Fecha inicio", "Fecha fin", "Tiempo"],
                   [(i, "", "", "", "", "") for i in range(1, 7)],
                   anchos=[0.8, 5.0, 3.6, 2.0, 2.0, 2.1], titulo="Resumen de la experiencia del profesional propuesto")
        self.p("…………………………, ……… de ………………………… de 20……", align=WD_ALIGN_PARAGRAPH.RIGHT)
        self.p("\n\n\n………………………………………………\nFirma, nombres y apellidos del profesional\nDNI N° …………………",
               align=WD_ALIGN_PARAGRAPH.CENTER)

    def generar(self, ruta):
        self._estilos()
        self._margenes()
        self._membrete()
        self.caratula()
        self.indice()
        self.cap1()
        self.cap2()
        self.cap3()
        self.cap4()
        self.cap5()
        self.cap6()
        self.cap7()
        self.cap8()
        self.firmas()
        self.anexos()
        self.doc.core_properties.title = f"TDR - {self.d['proyecto']}"
        self.doc.core_properties.author = self.d["entidad"]
        self.doc.save(ruta)
        return ruta


# =============================================================================
# 4. EXPORTACIÓN DEL CUADRO DE EVALUACIÓN EN EXCEL (.XLSX)
# =============================================================================

def exportar_excel(datos, personal, entregables, factores, params, ruta, n_postores=5):
    monto = a_float(datos["monto"])
    plazo = a_int(datos["plazo"])
    objeto = datos["objeto"]
    tipo = TIPO_OBJETO.get(objeto, "servicio")
    categoria = clasificar_procedimiento(monto, objeto, params)
    pen = calcular_penalidad(monto, plazo, objeto)
    gar = calcular_garantias(monto, objeto, params)

    azul = PatternFill("solid", fgColor=AZUL_HEX)
    celeste = PatternFill("solid", fgColor=CELESTE_HEX)
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    verde = PatternFill("solid", fgColor="E2EFDA")
    blanco_b = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    negrita = Font(bold=True, name="Arial", size=10)
    normal = Font(name="Arial", size=10)
    fino = Side(style="thin", color="808080")
    borde = Border(left=fino, right=fino, top=fino, bottom=fino)
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq = Alignment(horizontal="left", vertical="center", wrap_text=True)
    soles_fmt = '"S/ "#,##0.00'

    def titulo(ws, texto, ncols, fila=1):
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncols)
        c = ws.cell(fila, 1, texto)
        c.font = Font(bold=True, size=13, color="FFFFFF", name="Arial")
        c.fill = azul
        c.alignment = centro
        ws.row_dimensions[fila].height = 28

    def encabezado(ws, fila, textos):
        for i, t in enumerate(textos, start=1):
            c = ws.cell(fila, i, t)
            c.font, c.fill, c.alignment, c.border = blanco_b, azul, centro, borde
        ws.row_dimensions[fila].height = 32

    def celda(ws, fila, col, valor, fmt=None, font=normal, fill=None, align=izq):
        c = ws.cell(fila, col, valor)
        c.font, c.border, c.alignment = font, borde, align
        if fmt:
            c.number_format = fmt
        if fill:
            c.fill = fill
        return c

    wb = Workbook()

    # ---- Hoja 1: Datos del proceso ----
    ws = wb.active
    ws.title = "Datos del Proceso"
    titulo(ws, "DATOS GENERALES DEL PROCEDIMIENTO DE SELECCIÓN (SEACE)", 2)
    filas = [
        ("Entidad", datos["entidad"]), ("Proyecto / Obra", datos["proyecto"]),
        ("Objeto de contratación", objeto), ("Especialidad", datos["especialidad"]),
        ("Ubicación", datos.get("ubicacion", "")), ("CUI", datos.get("cui", "")),
        ("Sistema de contratación", datos["sistema"]), ("Marco normativo", datos.get("marco", MARCO_DEFAULT)),
        ("Valor referencial (S/)", monto), ("Plazo de ejecución (d.c.)", plazo),
        ("Categoría de exigencia", categoria),
        ("Procedimiento de selección", nombre_procedimiento(categoria, objeto)),
        ("UIT considerada (S/)", a_float(params.get("uit", UIT_DEFAULT))),
        ("Fecha de elaboración", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
    for i, (k, v) in enumerate(filas, start=3):
        celda(ws, i, 1, k, font=negrita, fill=celeste)
        c = celda(ws, i, 2, v)
        if k.endswith("(S/)"):
            c.number_format = soles_fmt
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 90

    # ---- Hoja 2: Factores de evaluación ----
    ws2 = wb.create_sheet("Evaluación de Ofertas")
    ncols = 3 + n_postores
    titulo(ws2, "CUADRO DE FACTORES DE EVALUACIÓN DE OFERTAS - COMITÉ DE SELECCIÓN", ncols)
    ws2.cell(2, 1, f"Valor referencial: {soles(monto)}  |  {nombre_procedimiento(categoria, objeto)}  |  "
                   "Celdas amarillas: ingreso del comité").font = Font(italic=True, name="Arial", size=9)
    # Bloque de ofertas económicas
    encabezado(ws2, 4, ["N°", "Concepto", "Puntaje Máx."] + [f"Postor {i}" for i in range(1, n_postores + 1)])
    celda(ws2, 5, 1, "", align=centro)
    celda(ws2, 5, 2, "Oferta económica (S/)", font=negrita)
    celda(ws2, 5, 3, "-", align=centro)
    for j in range(n_postores):
        celda(ws2, 5, 4 + j, None, fmt=soles_fmt, fill=amarillo)
    rango_ofertas = f"$D$5:${get_column_letter(3 + n_postores)}$5"
    celda(ws2, 6, 2, "Dentro de límites (90 % - 110 % del VR)" if tipo == "obra" else "Oferta ≤ valor estimado",
          font=negrita)
    celda(ws2, 6, 1, "")
    celda(ws2, 6, 3, monto, fmt=soles_fmt, align=centro)
    for j in range(n_postores):
        col = get_column_letter(4 + j)
        if tipo == "obra":
            f = f'=IF({col}5="","",IF(AND({col}5>=0.9*$C$6,{col}5<=1.1*$C$6),"SÍ","NO"))'
        else:
            f = f'=IF({col}5="","",IF({col}5<=$C$6,"SÍ","REVISAR"))'
        celda(ws2, 6, 4 + j, f, align=centro)
    fila = 8
    encabezado(ws2, fila, ["N°", "Factor de evaluación / Criterio", "Puntaje Máx."] +
               [f"Postor {i}" for i in range(1, n_postores + 1)])
    primera = fila + 1
    for i, fct in enumerate(factores, start=1):
        r = fila + i
        nombre = fct["factor"]
        celda(ws2, r, 1, i, align=centro)
        celda(ws2, r, 2, f"{nombre}\n{fct['criterio']}")
        celda(ws2, r, 3, a_float(fct["puntaje"]), fmt="0.00", align=centro, font=negrita)
        for j in range(n_postores):
            col = get_column_letter(4 + j)
            if nombre.strip().lower().startswith("precio"):
                f = (f'=IF(OR({col}$5="",{col}$5<=0),0,'
                     f'ROUND(SMALL({rango_ofertas},COUNTIF({rango_ofertas},"<=0")+1)*$C{r}/{col}$5,2))')
                celda(ws2, r, 4 + j, f, fmt="0.00", align=centro, fill=verde)
            else:
                celda(ws2, r, 4 + j, None, fmt="0.00", align=centro, fill=amarillo)
        ws2.row_dimensions[r].height = 42
    ultima = fila + len(factores)
    rt = ultima + 1
    celda(ws2, rt, 1, "")
    celda(ws2, rt, 2, "PUNTAJE TOTAL", font=negrita, fill=celeste)
    celda(ws2, rt, 3, f"=SUM(C{primera}:C{ultima})", fmt="0.00", align=centro, font=negrita, fill=celeste)
    for j in range(n_postores):
        col = get_column_letter(4 + j)
        celda(ws2, rt, 4 + j, f"=SUM({col}{primera}:{col}{ultima})", fmt="0.00", align=centro,
              font=negrita, fill=celeste)
    rr = rt + 1
    celda(ws2, rr, 1, "")
    celda(ws2, rr, 2, "ORDEN DE PRELACIÓN", font=negrita, fill=celeste)
    celda(ws2, rr, 3, "", fill=celeste)
    rango_tot = f"$D${rt}:${get_column_letter(3 + n_postores)}${rt}"
    for j in range(n_postores):
        col = get_column_letter(4 + j)
        celda(ws2, rr, 4 + j, f'=IF({col}$5="","",RANK({col}{rt},{rango_tot},0))', align=centro,
              font=negrita, fill=celeste)
    nota = rr + 2
    for k, t in enumerate([
        "Puntaje de precio: Pi = (Om x PMP) / Oi, donde Om = oferta de monto más bajo, PMP = puntaje máximo del precio, Oi = oferta evaluada.",
        "Los factores distintos al precio se asignan según la metodología de las Bases Estándar aplicables (celdas amarillas).",
        "Verificar que la suma de puntajes máximos sea 100 y que los factores coincidan con las bases integradas publicadas en el SEACE.",
    ]):
        ws2.cell(nota + k, 1, t).font = Font(italic=True, size=9, name="Arial")
    ws2.column_dimensions["A"].width = 6
    ws2.column_dimensions["B"].width = 62
    ws2.column_dimensions["C"].width = 13
    for j in range(n_postores):
        ws2.column_dimensions[get_column_letter(4 + j)].width = 15
    ws2.conditional_formatting.add(f"D6:{get_column_letter(3 + n_postores)}6",
                                   CellIsRule(operator="equal", formula=['"NO"'],
                                              fill=PatternFill("solid", fgColor="F8CBAD")))
    ws2.freeze_panes = "D5"

    # ---- Hoja 3: Requisitos mínimos ----
    ws3 = wb.create_sheet("Requisitos Mínimos")
    cols3 = ["N°", "Requisito de calificación", "Detalle mínimo exigido"] + [f"Postor {i}" for i in range(1, n_postores + 1)]
    titulo(ws3, "VERIFICACIÓN DE REQUISITOS DE CALIFICACIÓN (CUMPLE / NO CUMPLE)", len(cols3))
    encabezado(ws3, 3, cols3)
    dv = DataValidation(type="list", formula1='"CUMPLE,NO CUMPLE,N/A"', allow_blank=True)
    ws3.add_data_validation(dv)
    reqs = [("Capacidad legal - Habilitación (RNP vigente, no impedido)",
             "Inscripción vigente en el RNP en el capítulo correspondiente")]
    factor = a_float(params.get("factor_exp_postor", 1.0))
    if categoria == CAT_MENOR:
        factor = min(factor, 0.5)
    reqs.append(("Experiencia del postor en la especialidad",
                 f"Facturación ≥ {factor:g} x VR = {soles(monto * factor)} en prestaciones similares (10 años)"))
    reqs.append(("Equipamiento estratégico", "Según cuadro de equipamiento mínimo del TDR"))
    for r in personal:
        reqs.append((f"Personal clave: {r['cargo']}",
                     f"{r['formacion']} ({r['colegiatura']}); Exp. general {r['exp_gen']} años; "
                     f"Exp. específica {r['exp_esp']} años como {r['exp_desc']}; Capacitación {r['cap_h']} h en {r['cap_tema']}"))
    for i, (req, det) in enumerate(reqs, start=1):
        fr = 3 + i
        celda(ws3, fr, 1, i, align=centro)
        celda(ws3, fr, 2, req, font=negrita)
        celda(ws3, fr, 3, det)
        for j in range(n_postores):
            c = celda(ws3, fr, 4 + j, None, align=centro, fill=amarillo)
            dv.add(c)
        ws3.row_dimensions[fr].height = 48
    fin = 3 + len(reqs)
    rng = f"D4:{get_column_letter(3 + n_postores)}{fin}"
    ws3.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"NO CUMPLE"'],
                                                   fill=PatternFill("solid", fgColor="F8CBAD")))
    ws3.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"CUMPLE"'],
                                                   fill=PatternFill("solid", fgColor="C6EFCE")))
    rf = fin + 1
    celda(ws3, rf, 2, "RESULTADO: CALIFICA", font=negrita, fill=celeste)
    celda(ws3, rf, 1, "", fill=celeste)
    celda(ws3, rf, 3, "Califica si no tiene ningún 'NO CUMPLE'", fill=celeste)
    for j in range(n_postores):
        col = get_column_letter(4 + j)
        celda(ws3, rf, 4 + j,
              f'=IF(COUNTA({col}4:{col}{fin})=0,"",IF(COUNTIF({col}4:{col}{fin},"NO CUMPLE")>0,"NO CALIFICA","CALIFICA"))',
              align=centro, font=negrita, fill=celeste)
    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 40
    ws3.column_dimensions["C"].width = 70
    for j in range(n_postores):
        ws3.column_dimensions[get_column_letter(4 + j)].width = 14
    ws3.freeze_panes = "D4"

    # ---- Hoja 4: Personal clave ----
    ws4 = wb.create_sheet("Personal Clave")
    titulo(ws4, "PERFILES DEL PERSONAL CLAVE (CRITERIOS OSCE / OECE)", len(COLUMNAS_PERSONAL) + 1)
    encabezado(ws4, 3, ["N°"] + [c[1] for c in COLUMNAS_PERSONAL])
    for i, r in enumerate(personal, start=1):
        celda(ws4, 3 + i, 1, i, align=centro)
        for j, (k, _, _) in enumerate(COLUMNAS_PERSONAL, start=2):
            v = r.get(k, "")
            if k in ("exp_gen", "exp_esp", "cap_h"):
                v = a_float(v)
            celda(ws4, 3 + i, j, v, align=centro if k in ("exp_gen", "exp_esp", "cap_h", "colegiatura") else izq)
        ws4.row_dimensions[3 + i].height = 45
    for j, w in enumerate([6, 32, 34, 12, 12, 12, 12, 40, 48], start=1):
        ws4.column_dimensions[get_column_letter(j)].width = w

    # ---- Hoja 5: Entregables, penalidades y garantías ----
    ws5 = wb.create_sheet("Pagos y Penalidades")
    titulo(ws5, "ENTREGABLES, FORMA DE PAGO, PENALIDADES Y GARANTÍAS", 5)
    celda(ws5, 3, 1, "Valor referencial (S/)", font=negrita, fill=celeste)
    celda(ws5, 3, 2, monto, fmt=soles_fmt, fill=amarillo)
    celda(ws5, 4, 1, "Plazo (días)", font=negrita, fill=celeste)
    celda(ws5, 4, 2, plazo, fill=amarillo)
    encabezado(ws5, 6, ["N°", "Entregable", "Plazo (días)", "% Pago", "Monto (S/)"])
    for i, e in enumerate(entregables, start=1):
        r = 6 + i
        celda(ws5, r, 1, i, align=centro)
        celda(ws5, r, 2, e["descripcion"])
        celda(ws5, r, 3, a_int(e["plazo"]), align=centro)
        celda(ws5, r, 4, a_float(e["porcentaje"]) / 100, fmt="0.00%", align=centro, fill=amarillo)
        celda(ws5, r, 5, f"=$B$3*D{r}", fmt=soles_fmt)
        ws5.row_dimensions[r].height = 36
    ue = 6 + len(entregables)
    celda(ws5, ue + 1, 2, "TOTAL", font=negrita, fill=celeste)
    celda(ws5, ue + 1, 4, f"=SUM(D7:D{ue})", fmt="0.00%", align=centro, font=negrita, fill=celeste)
    celda(ws5, ue + 1, 5, f"=SUM(E7:E{ue})", fmt=soles_fmt, font=negrita, fill=celeste)
    b = ue + 3
    encabezado(ws5, b, ["", "Penalidad por mora", "Valor", "", ""])
    celda(ws5, b + 1, 2, "Factor F (0.40 si plazo ≤ 60 d; 0.25 bienes/servicios/consultorías o 0.15 obras si > 60 d)")
    f_formula = f'=IF($B$4<=60,0.4,{0.15 if tipo == "obra" else 0.25})'
    celda(ws5, b + 1, 3, f_formula, fmt="0.00", align=centro)
    celda(ws5, b + 2, 2, "Penalidad diaria = 0.10 x Monto / (F x Plazo)")
    celda(ws5, b + 2, 3, f"=0.1*$B$3/(C{b + 1}*$B$4)", fmt=soles_fmt)
    celda(ws5, b + 3, 2, "Tope de penalidad (10 % del monto)")
    celda(ws5, b + 3, 3, "=0.1*$B$3", fmt=soles_fmt)
    celda(ws5, b + 4, 2, "Días de atraso para alcanzar el tope")
    celda(ws5, b + 4, 3, f"=ROUNDUP(C{b + 3}/C{b + 2},0)", align=centro)
    g = b + 6
    encabezado(ws5, g, ["", "Garantías y adelantos", "%", "Monto (S/)", ""])
    for k, (nom, pct) in enumerate([("Garantía de fiel cumplimiento", gar["gfc_pct"]),
                                    ("Adelanto directo", gar["ad_pct"]),
                                    ("Adelanto para materiales", gar["mat_pct"])], start=1):
        celda(ws5, g + k, 2, nom)
        celda(ws5, g + k, 3, pct / 100, fmt="0.00%", align=centro, fill=amarillo)
        celda(ws5, g + k, 4, f"=$B$3*C{g + k}", fmt=soles_fmt)
    ws5.column_dimensions["A"].width = 24
    ws5.column_dimensions["B"].width = 70
    ws5.column_dimensions["C"].width = 16
    ws5.column_dimensions["D"].width = 16
    ws5.column_dimensions["E"].width = 18

    for hoja in wb.worksheets:
        hoja.sheet_view.showGridLines = False
        hoja.page_setup.orientation = "landscape"
        hoja.page_setup.fitToWidth = 1
        hoja.sheet_properties.pageSetUpPr.fitToPage = True
        hoja.page_setup.fitToHeight = 0
    wb.save(ruta)
    return ruta


# =============================================================================
# 5. UTILIDADES
# =============================================================================

def abrir_archivo(ruta):
    """Abre el archivo con la aplicación predeterminada del sistema."""
    try:
        if hasattr(os, "startfile"):
            os.startfile(ruta)  # Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta])
    except Exception:
        pass


def nombre_archivo_seguro(texto, maximo=60):
    permitido = "".join(ch if ch.isalnum() or ch in " -_" else "_" for ch in texto)
    return "_".join(permitido.split())[:maximo] or "TDR"


def datos_ejemplo():
    datos = {
        "entidad": "Municipalidad Distrital de Ejemplo",
        "proyecto": "Mejoramiento del servicio de transitabilidad vehicular y peatonal en las principales calles del distrito",
        "objeto": "Consultoría de Expediente Técnico",
        "especialidad": "Vial",
        "monto": "350000",
        "plazo": "90",
        "sistema": "Suma Alzada",
        "ubicacion": "Distrito de Ejemplo, Provincia de Ejemplo, Región Ejemplo",
        "cui": "2512345",
        "marco": MARCO_DEFAULT,
    }
    params = dict(PARAMETROS_DEFAULT)
    cat = clasificar_procedimiento(a_float(datos["monto"]), datos["objeto"], params)
    personal = generar_personal_clave(datos["objeto"], datos["especialidad"], cat)
    entregables = generar_entregables(datos["objeto"], a_int(datos["plazo"]))
    factores = [{"factor": f, "criterio": c, "puntaje": str(p)}
                for f, c, p in FACTORES_DEFAULT[CATEGORIA_UMBRAL[TIPO_OBJETO[datos["objeto"]]]]]
    return datos, personal, entregables, factores, params


def factores_por_objeto(objeto):
    cat = CATEGORIA_UMBRAL[TIPO_OBJETO.get(objeto, "servicio")]
    return [{"factor": f, "criterio": c, "puntaje": str(p)} for f, c, p in FACTORES_DEFAULT[cat]]


# =============================================================================
# 6. INTERFAZ GRÁFICA (CUSTOMTKINTER)
# =============================================================================

def lanzar_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import customtkinter as ctk

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    class TablaEditable(ctk.CTkFrame):
        """Tabla editable con encabezados fijos, filas de CTkEntry y botón eliminar."""

        def __init__(self, master, columnas, al_cambiar=None, **kw):
            super().__init__(master, **kw)
            self.columnas = columnas
            self.al_cambiar = al_cambiar
            self.filas = []
            enc = ctk.CTkFrame(self, fg_color=("#1F3864", "#1F3864"), corner_radius=6)
            enc.pack(fill="x", padx=2, pady=(2, 0))
            ctk.CTkLabel(enc, text="N°", width=34, text_color="white",
                         font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=2, pady=4)
            for _, titulo, ancho in columnas:
                ctk.CTkLabel(enc, text=titulo, width=ancho, text_color="white", wraplength=ancho - 6,
                             font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=2, pady=4)
            ctk.CTkLabel(enc, text="", width=34).pack(side="left", padx=2)
            self.cuerpo = ctk.CTkScrollableFrame(self, corner_radius=0)
            self.cuerpo.pack(fill="both", expand=True, padx=2, pady=2)

        def _notificar(self, *_):
            if self.al_cambiar:
                self.al_cambiar()

        def agregar(self, datos=None):
            datos = datos or {}
            marco = ctk.CTkFrame(self.cuerpo, fg_color="transparent")
            marco.pack(fill="x", pady=1)
            num = ctk.CTkLabel(marco, text=str(len(self.filas) + 1), width=34)
            num.pack(side="left", padx=2)
            entradas = {}
            for clave, _, ancho in self.columnas:
                e = ctk.CTkEntry(marco, width=ancho)
                e.insert(0, str(datos.get(clave, "")))
                e.pack(side="left", padx=2)
                e.bind("<KeyRelease>", self._notificar)
                entradas[clave] = e
            fila = {"marco": marco, "num": num, "entradas": entradas}
            btn = ctk.CTkButton(marco, text="✕", width=34, fg_color="#9B2C2C", hover_color="#742A2A",
                                command=lambda f=fila: self.eliminar(f))
            btn.pack(side="left", padx=2)
            self.filas.append(fila)
            self._notificar()

        def eliminar(self, fila):
            fila["marco"].destroy()
            self.filas.remove(fila)
            for i, f in enumerate(self.filas, start=1):
                f["num"].configure(text=str(i))
            self._notificar()

        def cargar(self, lista):
            for f in list(self.filas):
                f["marco"].destroy()
            self.filas.clear()
            for d in lista:
                self.agregar(d)

        def obtener(self):
            return [{k: e.get().strip() for k, e in f["entradas"].items()} for f in self.filas]

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title(f"{APP_TITULO}  v{APP_VERSION}")
            self.geometry("1280x800")
            self.minsize(1100, 700)
            self.params_vars = {}
            self._construir_superior()
            self._construir_pestanas()
            self._construir_inferior()
            self.aplicar_criterios(confirmar=False)

        # ---------------- Panel superior ----------------
        def _construir_superior(self):
            sup = ctk.CTkFrame(self, corner_radius=10)
            sup.pack(fill="x", padx=10, pady=(10, 5))
            cab = ctk.CTkFrame(sup, fg_color="transparent")
            cab.pack(fill="x", padx=10, pady=(6, 2))
            ctk.CTkLabel(cab, text="📋  Datos Generales del Proceso SEACE",
                         font=ctk.CTkFont(size=17, weight="bold")).pack(side="left")
            self.modo = ctk.CTkSegmentedButton(cab, values=["Oscuro", "Claro"], command=self._cambiar_modo)
            self.modo.set("Oscuro")
            self.modo.pack(side="right")
            ctk.CTkLabel(cab, text="Tema:").pack(side="right", padx=6)

            g = ctk.CTkFrame(sup, fg_color="transparent")
            g.pack(fill="x", padx=10, pady=(0, 6))
            for c in range(8):
                g.grid_columnconfigure(c, weight=1 if c % 2 else 0)

            def lbl(texto, r, c):
                ctk.CTkLabel(g, text=texto, anchor="e").grid(row=r, column=c, sticky="e", padx=(6, 4), pady=3)

            lbl("Entidad / Municipalidad:", 0, 0)
            self.e_entidad = ctk.CTkEntry(g, placeholder_text="Ej. Municipalidad Provincial de ...")
            self.e_entidad.grid(row=0, column=1, columnspan=3, sticky="ew", pady=3)
            lbl("Proyecto / Obra:", 0, 4)
            self.e_proyecto = ctk.CTkEntry(g, placeholder_text="Nombre del proyecto según Invierte.pe")
            self.e_proyecto.grid(row=0, column=5, columnspan=3, sticky="ew", pady=3)

            lbl("Objeto:", 1, 0)
            self.o_objeto = ctk.CTkOptionMenu(g, values=OBJETOS, command=lambda _: self.aplicar_criterios(False))
            self.o_objeto.grid(row=1, column=1, sticky="ew", pady=3)
            lbl("Especialidad:", 1, 2)
            self.o_especialidad = ctk.CTkOptionMenu(g, values=ESPECIALIDADES,
                                                    command=lambda _: self.aplicar_criterios(False))
            self.o_especialidad.grid(row=1, column=3, sticky="ew", pady=3)
            lbl("Sistema:", 1, 4)
            self.o_sistema = ctk.CTkOptionMenu(g, values=SISTEMAS)
            self.o_sistema.grid(row=1, column=5, sticky="ew", pady=3)
            lbl("Marco legal:", 1, 6)
            self.o_marco = ctk.CTkOptionMenu(g, values=list(MARCOS_LEGALES.keys()), dynamic_resizing=False)
            self.o_marco.grid(row=1, column=7, sticky="ew", pady=3)

            lbl("Monto Referencial (S/):", 2, 0)
            self.e_monto = ctk.CTkEntry(g, placeholder_text="0.00")
            self.e_monto.grid(row=2, column=1, sticky="ew", pady=3)
            self.e_monto.bind("<KeyRelease>", lambda e: self.actualizar_resumen())
            lbl("Plazo (días calendario):", 2, 2)
            self.e_plazo = ctk.CTkEntry(g, placeholder_text="90")
            self.e_plazo.grid(row=2, column=3, sticky="ew", pady=3)
            self.e_plazo.bind("<KeyRelease>", lambda e: self.actualizar_resumen())
            lbl("Ubicación:", 2, 4)
            self.e_ubicacion = ctk.CTkEntry(g, placeholder_text="Distrito, Provincia, Región")
            self.e_ubicacion.grid(row=2, column=5, sticky="ew", pady=3)
            lbl("CUI:", 2, 6)
            self.e_cui = ctk.CTkEntry(g, placeholder_text="Código Único de Inversión")
            self.e_cui.grid(row=2, column=7, sticky="ew", pady=3)

            fila_res = ctk.CTkFrame(sup, fg_color="transparent")
            fila_res.pack(fill="x", padx=10, pady=(0, 8))
            self.lbl_resumen = ctk.CTkLabel(fila_res, text="", anchor="w", justify="left",
                                            font=ctk.CTkFont(size=13, weight="bold"),
                                            text_color=("#1F3864", "#8AB4F8"))
            self.lbl_resumen.pack(side="left", fill="x", expand=True)
            ctk.CTkButton(fila_res, text="⚙ Aplicar criterios OSCE automáticos", width=260,
                          command=lambda: self.aplicar_criterios(True)).pack(side="right")

            self.e_monto.insert(0, "350000")
            self.e_plazo.insert(0, "90")

        def _cambiar_modo(self, valor):
            ctk.set_appearance_mode("dark" if valor == "Oscuro" else "light")

        # ---------------- Pestañas ----------------
        def _construir_pestanas(self):
            self.tabs = ctk.CTkTabview(self, corner_radius=10)
            self.tabs.pack(fill="both", expand=True, padx=10, pady=5)
            t1 = self.tabs.add("1. Personal Clave y Proveedor")
            t2 = self.tabs.add("2. Entregables, Pago y Penalidades")
            t3 = self.tabs.add("3. Factores de Evaluación y Parámetros")

            # --- Pestaña 1 ---
            barra1 = ctk.CTkFrame(t1, fg_color="transparent")
            barra1.pack(fill="x", pady=(0, 4))
            ctk.CTkLabel(barra1, text="Requisitos de calificación del personal clave (editable). "
                                      "Los valores se asignan según la categoría del procedimiento.",
                         anchor="w").pack(side="left")
            ctk.CTkButton(barra1, text="+ Agregar profesional", width=170,
                          command=lambda: self.tabla_personal.agregar()).pack(side="right")
            self.tabla_personal = TablaEditable(t1, COLUMNAS_PERSONAL)
            self.tabla_personal.pack(fill="both", expand=True)

            # --- Pestaña 2 ---
            cont = ctk.CTkFrame(t2, fg_color="transparent")
            cont.pack(fill="both", expand=True)
            izq = ctk.CTkFrame(cont, fg_color="transparent")
            izq.pack(side="left", fill="both", expand=True)
            barra2 = ctk.CTkFrame(izq, fg_color="transparent")
            barra2.pack(fill="x", pady=(0, 4))
            ctk.CTkLabel(barra2, text="Entregables porcentuales y plazos", anchor="w").pack(side="left")
            ctk.CTkButton(barra2, text="↻ Recalcular", width=110,
                          command=self.recalcular_entregables).pack(side="right", padx=4)
            ctk.CTkButton(barra2, text="+ Agregar", width=100,
                          command=lambda: self.tabla_entregables.agregar()).pack(side="right")
            self.tabla_entregables = TablaEditable(izq, COLUMNAS_ENTREGABLES, al_cambiar=self.actualizar_resumen)
            self.tabla_entregables.pack(fill="both", expand=True)

            der = ctk.CTkFrame(cont, width=420)
            der.pack(side="right", fill="y", padx=(8, 0))
            der.pack_propagate(False)
            ctk.CTkLabel(der, text="Condiciones de conformidad y adelantos",
                         font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(8, 4))
            frm = ctk.CTkFrame(der, fg_color="transparent")
            frm.pack(fill="x", padx=8)
            for i, (clave, texto) in enumerate([
                ("revision_entidad_dias", "Plazo de revisión de la Entidad (días)"),
                ("subsanacion_dias", "Plazo de subsanación (días)"),
                ("garantia_fiel_pct", "Garantía de fiel cumplimiento (%)"),
                ("adelanto_directo_obra_pct", "Adelanto directo - obras (%)"),
                ("adelanto_materiales_obra_pct", "Adelanto de materiales - obras (%)"),
                ("adelanto_directo_otros_pct", "Adelanto directo - consultorías/servicios (%)"),
                ("subcontratacion_max_pct", "Subcontratación máxima (%)"),
            ]):
                self._campo_param(frm, i, clave, texto)
            self.txt_calculos = ctk.CTkTextbox(der, font=ctk.CTkFont(family="Consolas", size=12), wrap="word")
            self.txt_calculos.pack(fill="both", expand=True, padx=8, pady=8)

            # --- Pestaña 3 ---
            cont3 = ctk.CTkFrame(t3, fg_color="transparent")
            cont3.pack(fill="both", expand=True)
            izq3 = ctk.CTkFrame(cont3, fg_color="transparent")
            izq3.pack(side="left", fill="both", expand=True)
            barra3 = ctk.CTkFrame(izq3, fg_color="transparent")
            barra3.pack(fill="x", pady=(0, 4))
            self.lbl_factores = ctk.CTkLabel(barra3, text="Factores de evaluación (la suma debe ser 100)", anchor="w")
            self.lbl_factores.pack(side="left")
            ctk.CTkButton(barra3, text="+ Agregar factor", width=140,
                          command=lambda: self.tabla_factores.agregar()).pack(side="right")
            self.tabla_factores = TablaEditable(izq3, COLUMNAS_FACTORES, al_cambiar=self._suma_factores)
            self.tabla_factores.pack(fill="both", expand=True)
            der3 = ctk.CTkFrame(cont3, width=430)
            der3.pack(side="right", fill="y", padx=(8, 0))
            der3.pack_propagate(False)
            ctk.CTkLabel(der3, text="Parámetros normativos (editables)",
                         font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(8, 4))
            frm3 = ctk.CTkFrame(der3, fg_color="transparent")
            frm3.pack(fill="x", padx=8)
            for i, (clave, texto) in enumerate([
                ("uit", "Valor de la UIT (S/)"),
                ("menor_cuantia_uit", "Menor cuantía (N° UIT)"),
                ("limite_as_obra", "Límite AS - Obras (S/)"),
                ("limite_as_consultoria", "Límite AS - Consultoría (S/)"),
                ("limite_as_servicio", "Límite AS - Servicios (S/)"),
                ("factor_exp_postor", "Exp. postor (veces el VR)"),
            ]):
                self._campo_param(frm3, i, clave, texto)
            ctk.CTkLabel(der3, justify="left", anchor="w", text_color="gray",
                         text=("⚠ Valores referenciales. Verifique la UIT y los\n"
                               "umbrales con la Ley de Presupuesto del año fiscal\n"
                               "y las Bases Estándar vigentes antes de convocar.")
                         ).pack(fill="x", padx=8, pady=10)

        def _campo_param(self, frame, fila, clave, texto):
            ctk.CTkLabel(frame, text=texto, anchor="w").grid(row=fila, column=0, sticky="w", pady=3)
            e = ctk.CTkEntry(frame, width=110, justify="right")
            e.insert(0, fmt_num(PARAMETROS_DEFAULT[clave]))
            e.grid(row=fila, column=1, sticky="e", padx=(6, 0), pady=3)
            e.bind("<KeyRelease>", lambda ev: self.actualizar_resumen())
            frame.grid_columnconfigure(0, weight=1)
            self.params_vars[clave] = e

        # ---------------- Barra inferior ----------------
        def _construir_inferior(self):
            inf = ctk.CTkFrame(self, corner_radius=10)
            inf.pack(fill="x", padx=10, pady=(5, 10))
            ctk.CTkButton(inf, text="💾 Guardar", width=100, fg_color="gray40", hover_color="gray30",
                          command=self.guardar_proyecto).pack(side="left", padx=(10, 4), pady=10)
            ctk.CTkButton(inf, text="📂 Abrir", width=90, fg_color="gray40", hover_color="gray30",
                          command=self.abrir_proyecto).pack(side="left", padx=4, pady=10)
            ctk.CTkButton(inf, text="2. Exportar Cuadro de Evaluación y Requisitos Mínimos en Excel (.XLSX)",
                          height=42, fg_color="#1E7B34", hover_color="#155D27",
                          font=ctk.CTkFont(size=13, weight="bold"),
                          command=self.exportar_xlsx).pack(side="right", padx=(4, 10), pady=10)
            ctk.CTkButton(inf, text="1. Generar Términos de Referencia (TDR) Oficiales en Word (.DOCX)",
                          height=42, fg_color="#1F5FBF", hover_color="#174A94",
                          font=ctk.CTkFont(size=13, weight="bold"),
                          command=self.exportar_docx).pack(side="right", padx=4, pady=10)

        # ---------------- Lógica ----------------
        def leer_params(self):
            p = dict(PARAMETROS_DEFAULT)
            for k, e in self.params_vars.items():
                p[k] = a_float(e.get(), PARAMETROS_DEFAULT[k])
            return p

        def leer_datos(self):
            return {
                "entidad": self.e_entidad.get().strip(),
                "proyecto": self.e_proyecto.get().strip(),
                "objeto": self.o_objeto.get(),
                "especialidad": self.o_especialidad.get(),
                "sistema": self.o_sistema.get(),
                "marco": self.o_marco.get(),
                "monto": self.e_monto.get().strip(),
                "plazo": self.e_plazo.get().strip(),
                "ubicacion": self.e_ubicacion.get().strip(),
                "cui": self.e_cui.get().strip(),
            }

        def aplicar_criterios(self, confirmar=True):
            if confirmar and not messagebox.askyesno(
                    "Aplicar criterios", "Se reemplazarán el personal clave, los entregables y los factores "
                                         "con los valores automáticos. ¿Continuar?"):
                return
            d = self.leer_datos()
            monto, plazo = a_float(d["monto"]), max(1, a_int(d["plazo"], 1))
            cat = clasificar_procedimiento(monto, d["objeto"], self.leer_params())
            self.tabla_personal.cargar(generar_personal_clave(d["objeto"], d["especialidad"], cat))
            self.tabla_entregables.cargar(generar_entregables(d["objeto"], plazo))
            self.tabla_factores.cargar(factores_por_objeto(d["objeto"]))
            self.actualizar_resumen()

        def recalcular_entregables(self):
            d = self.leer_datos()
            self.tabla_entregables.cargar(generar_entregables(d["objeto"], max(1, a_int(d["plazo"], 1))))

        def _suma_factores(self):
            total = sum(a_float(f["puntaje"]) for f in self.tabla_factores.obtener())
            color = ("#1E7B34", "#7CD992") if abs(total - 100) < 0.01 else ("#B00020", "#FF8A80")
            self.lbl_factores.configure(text=f"Factores de evaluación - Suma de puntajes: {total:g} / 100",
                                        text_color=color)

        def actualizar_resumen(self):
            if not hasattr(self, "txt_calculos"):
                return
            d = self.leer_datos()
            params = self.leer_params()
            monto, plazo = a_float(d["monto"]), max(1, a_int(d["plazo"], 1))
            cat = clasificar_procedimiento(monto, d["objeto"], params)
            pen = calcular_penalidad(monto, plazo, d["objeto"])
            gar = calcular_garantias(monto, d["objeto"], params)
            self.lbl_resumen.configure(
                text=f"Categoría: {cat}  →  {nombre_procedimiento(cat, d['objeto'])}   |   "
                     f"Penalidad diaria: {soles(pen['diaria'])} (F = {pen['F']:.2f})   |   "
                     f"Garantía F.C.: {soles(gar['gfc'])}")
            ents = self.tabla_entregables.obtener()
            suma = sum(a_float(e["porcentaje"]) for e in ents)
            lineas = [
                "PENALIDAD POR MORA",
                "  P = 0.10 x Monto / (F x Plazo)",
                f"  Monto  = {soles(monto)}",
                f"  Plazo  = {plazo} días   F = {pen['F']:.2f}",
                f"  Diaria = {soles(pen['diaria'])}",
                f"  Tope 10% = {soles(pen['tope'])}",
                f"  Tope en  = {pen['dias_tope']} días de atraso",
                "",
                "GARANTÍAS Y ADELANTOS",
                f"  Fiel cumplimiento ({gar['gfc_pct']:g}%): {soles(gar['gfc'])}",
                f"  Adelanto directo ({gar['ad_pct']:g}%):  {soles(gar['ad'])}",
                f"  Adelanto materiales ({gar['mat_pct']:g}%): {soles(gar['mat'])}",
                "",
                "ENTREGABLES",
            ]
            for i, e in enumerate(ents, start=1):
                pct = a_float(e["porcentaje"])
                lineas.append(f"  {i:02d}. día {a_int(e['plazo']):>4}  {pct:>6.2f}%  {soles(monto * pct / 100)}")
            lineas.append(f"  SUMA: {suma:g} %  {'✔ OK' if abs(suma - 100) < 0.01 else '✖ DEBE SUMAR 100 %'}")
            self.txt_calculos.configure(state="normal")
            self.txt_calculos.delete("1.0", "end")
            self.txt_calculos.insert("1.0", "\n".join(lineas))
            self.txt_calculos.configure(state="disabled")
            self._suma_factores()

        def validar(self):
            d = self.leer_datos()
            errores = []
            if not d["entidad"]:
                errores.append("• Ingrese el nombre de la Entidad.")
            if not d["proyecto"]:
                errores.append("• Ingrese el nombre del proyecto.")
            if a_float(d["monto"]) <= 0:
                errores.append("• El monto referencial debe ser mayor a cero.")
            if a_int(d["plazo"]) <= 0:
                errores.append("• El plazo debe ser mayor a cero.")
            if not self.tabla_personal.obtener():
                errores.append("• Debe existir al menos un profesional en el personal clave.")
            if not self.tabla_entregables.obtener():
                errores.append("• Debe existir al menos un entregable.")
            if errores:
                messagebox.showerror("Datos incompletos", "\n".join(errores))
                return None
            suma = sum(a_float(e["porcentaje"]) for e in self.tabla_entregables.obtener())
            if abs(suma - 100) > 0.01 and not messagebox.askyesno(
                    "Advertencia", f"Los porcentajes de pago suman {suma:g} %. ¿Desea continuar?"):
                return None
            return d

        def _ruta_guardado(self, ext, prefijo, tipos):
            d = self.leer_datos()
            nombre = f"{prefijo}_{nombre_archivo_seguro(d['proyecto'] or 'proyecto')}.{ext}"
            return filedialog.asksaveasfilename(defaultextension=f".{ext}", initialfile=nombre, filetypes=tipos)

        def exportar_docx(self):
            d = self.validar()
            if not d:
                return
            ruta = self._ruta_guardado("docx", "TDR", [("Documento Word", "*.docx")])
            if not ruta:
                return
            try:
                GeneradorTDR(d, self.tabla_personal.obtener(), self.tabla_entregables.obtener(),
                             self.tabla_factores.obtener(), self.leer_params()).generar(ruta)
            except PermissionError:
                messagebox.showerror("Error", "No se pudo guardar. Cierre el archivo si está abierto en Word.")
                return
            except Exception as ex:
                messagebox.showerror("Error", f"No se pudo generar el TDR:\n{ex}")
                return
            messagebox.showinfo("TDR generado", f"Términos de Referencia generados:\n{ruta}")
            abrir_archivo(ruta)

        def exportar_xlsx(self):
            d = self.validar()
            if not d:
                return
            ruta = self._ruta_guardado("xlsx", "Evaluacion", [("Libro de Excel", "*.xlsx")])
            if not ruta:
                return
            try:
                exportar_excel(d, self.tabla_personal.obtener(), self.tabla_entregables.obtener(),
                               self.tabla_factores.obtener(), self.leer_params(), ruta)
            except PermissionError:
                messagebox.showerror("Error", "No se pudo guardar. Cierre el archivo si está abierto en Excel.")
                return
            except Exception as ex:
                messagebox.showerror("Error", f"No se pudo generar el Excel:\n{ex}")
                return
            messagebox.showinfo("Excel generado", f"Cuadro de evaluación generado:\n{ruta}")
            abrir_archivo(ruta)

        def guardar_proyecto(self):
            ruta = self._ruta_guardado("json", "Proyecto_TDR", [("Proyecto TDR", "*.json")])
            if not ruta:
                return
            data = {"version": APP_VERSION, "datos": self.leer_datos(), "params": self.leer_params(),
                    "personal": self.tabla_personal.obtener(), "entregables": self.tabla_entregables.obtener(),
                    "factores": self.tabla_factores.obtener()}
            with open(ruta, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
            messagebox.showinfo("Guardado", f"Proyecto guardado en:\n{ruta}")

        def abrir_proyecto(self):
            ruta = filedialog.askopenfilename(filetypes=[("Proyecto TDR", "*.json")])
            if not ruta:
                return
            try:
                with open(ruta, encoding="utf-8") as fh:
                    data = json.load(fh)
            except Exception as ex:
                messagebox.showerror("Error", f"No se pudo leer el archivo:\n{ex}")
                return
            d = data.get("datos", {})
            for entry, clave in ((self.e_entidad, "entidad"), (self.e_proyecto, "proyecto"),
                                 (self.e_monto, "monto"), (self.e_plazo, "plazo"),
                                 (self.e_ubicacion, "ubicacion"), (self.e_cui, "cui")):
                entry.delete(0, "end")
                entry.insert(0, d.get(clave, ""))
            for menu, clave, opciones in ((self.o_objeto, "objeto", OBJETOS),
                                          (self.o_especialidad, "especialidad", ESPECIALIDADES),
                                          (self.o_sistema, "sistema", SISTEMAS),
                                          (self.o_marco, "marco", list(MARCOS_LEGALES))):
                if d.get(clave) in opciones:
                    menu.set(d[clave])
            for k, v in data.get("params", {}).items():
                if k in self.params_vars:
                    self.params_vars[k].delete(0, "end")
                    self.params_vars[k].insert(0, fmt_num(v))
            self.tabla_personal.cargar(data.get("personal", []))
            self.tabla_entregables.cargar(data.get("entregables", []))
            self.tabla_factores.cargar(data.get("factores", []))
            self.actualizar_resumen()

    app = App()
    app.mainloop()


# =============================================================================
# 7. PUNTO DE ENTRADA
# =============================================================================

def modo_prueba(carpeta="."):
    datos, personal, entregables, factores, params = datos_ejemplo()
    os.makedirs(carpeta, exist_ok=True)
    r1 = GeneradorTDR(datos, personal, entregables, factores, params).generar(
        os.path.join(carpeta, "TDR_ejemplo.docx"))
    r2 = exportar_excel(datos, personal, entregables, factores, params,
                        os.path.join(carpeta, "Evaluacion_ejemplo.xlsx"))
    print("Generados:", r1, r2)
    return r1, r2


if __name__ == "__main__":
    if "--prueba" in sys.argv:
        idx = sys.argv.index("--prueba")
        destino = sys.argv[idx + 1] if len(sys.argv) > idx + 1 else "."
        modo_prueba(destino)
    else:
        lanzar_gui()

# =============================================================================
# INSTALACIÓN Y EMPAQUETADO
# -----------------------------------------------------------------------------
#   pip install customtkinter python-docx openpyxl
#   pyinstaller --onefile --windowed --name "Generador_TDR_OSCE" --collect-all customtkinter app_tdr_osce_pro.py
# =============================================================================
