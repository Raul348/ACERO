# -*- coding: utf-8 -*-
"""
===============================================================================
 Diseñador Pro de Cimentaciones para Equipos Vibratorios (Grupos Electrógenos /
 Subestaciones) y Macizos para Postes/Torres (Sulzberger) a Word y AutoCAD (.DXF)
===============================================================================
 Módulo A : Bancada / dado antivibratorio de equipos (ACI 351.3R-04, NTE E.050,
            NTE E.060, ACI 318-19 Cap. 17 para pernos de anclaje).
 Módulo B : Macizo de cimentación de postes, reflectores, mástiles y torres por
            el Método de Sulzberger con cargas de viento NTE E.020 / CNE.

 Exporta:  Memoria de cálculo (.docx) · Lámina estructural CE-01 (.dxf)
           Hojas de cálculo y metrados (.xlsx) · Proceso por lotes desde Excel.

 Aplicación 100 % offline. Unidades base: t, m, kg/cm² (indicadas en cada dato).
===============================================================================
"""

import math
import os
import sys
import re
import datetime
import subprocess
import traceback

# Librerías de exportación (se importan aquí para detectar faltantes al inicio)
try:
    import ezdxf
    from ezdxf.enums import TextEntityAlignment
except ImportError:  # pragma: no cover
    ezdxf = None
try:
    import docx
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:  # pragma: no cover
    docx = None
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import ScatterChart, Reference, Series
except ImportError:  # pragma: no cover
    openpyxl = None
try:
    import pandas as pd
except ImportError:  # pragma: no cover
    pd = None

APP_TITULO = ("Diseñador Pro de Cimentaciones para Equipos Vibratorios (Grupos Electrógenos / "
              "Subestaciones) y Macizos para Postes/Torres (Sulzberger) a Word y AutoCAD (.DXF)")
APP_VERSION = "1.0"

# =============================================================================
# 1. CONSTANTES Y TABLAS NORMATIVAS
# =============================================================================
G_ACC = 9.81            # m/s²
GAMMA_C = 2.40          # t/m³ concreto armado
FY = 4200.0             # kg/cm² acero ASTM A615 Gr.60
KGF_N = 9.80665         # N por kgf

# Barras corrugadas: diámetro (cm), área (cm²), peso (kg/m)
BARRAS = {
    '3/8"': (0.953, 0.71, 0.560),
    '1/2"': (1.270, 1.29, 0.994),
    '5/8"': (1.588, 1.99, 1.552),
    '3/4"': (1.905, 2.84, 2.235),
    '1"':   (2.540, 5.10, 3.973),
}

# Pernos de anclaje: diámetro nominal da (cm), área efectiva roscada Ase (cm²)
PERNOS = {
    '5/8"':   {"da": 1.588, "Ase": 1.465},
    '3/4"':   {"da": 1.905, "Ase": 2.155},
    '7/8"':   {"da": 2.223, "Ase": 2.981},
    '1"':     {"da": 2.540, "Ase": 3.910},
    '1 1/4"': {"da": 3.175, "Ase": 6.252},
}

# Grados de pernos: fya, futa (kg/cm²)
GRADOS_PERNO = {
    "ASTM F1554 Grado 36": (2530.0, 4080.0),
    "ASTM F1554 Grado 55": (3870.0, 5270.0),
    "ASTM A307 Grado A":   (2530.0, 4220.0),
    "ASTM F1554 Grado 105": (7380.0, 8790.0),
}

TIPOS_EQUIPO = {
    "Grupo Electrógeno Diésel Encapsulado": {
        "clase": "rotativa", "ratio_min": 2.5, "ratio_rec": 3.0, "pct_rotor": 30.0,
        "G_bal": 6.3, "desc": "máquina rotativa (motor diésel + alternador sobre skid)"},
    "Compresor Industrial": {
        "clase": "reciprocante", "ratio_min": 4.0, "ratio_rec": 5.0, "pct_rotor": 25.0,
        "G_bal": 16.0, "desc": "máquina reciprocante / alternativa (pistones)"},
    "Bomba Contra Incendios UL/FM": {
        "clase": "rotativa", "ratio_min": 2.5, "ratio_rec": 3.0, "pct_rotor": 20.0,
        "G_bal": 6.3, "desc": "máquina rotativa (bomba centrífuga + motor)"},
    "Transformador Seco/Aceite": {
        "clase": "electromagnetica", "ratio_min": 1.5, "ratio_rec": 2.0, "pct_rotor": 1.0,
        "G_bal": 0.0, "desc": "equipo estático con excitación magnetostrictiva a 2·f_red"},
}

VELOCIDADES = ["1800 RPM - 60Hz", "1500 RPM - 50Hz", "3600 RPM - 60Hz",
               "1200 RPM - 60Hz", "900 RPM - 60Hz", "3000 RPM - 50Hz"]

AISLAMIENTOS = [
    'Junta Perimetral de Poliestireno Elastificado / Neopreno e=1" + Amortiguadores Spring/Elastómero',
    'Junta Perimetral de Neopreno e=1" + Amortiguadores de Elastómero (Neopreno)',
    'Junta Perimetral e=1" + Resortes (Spring Isolators) deflexión estática 25 mm',
    'Junta Perimetral e=1" sin amortiguadores (anclaje rígido al bloque)',
]

TIPOS_POSTE = [
    "Poste de Concreto CAC 13m/400daN",
    "Poste Metálico Octogonal con Corona de 6 Reflectores LED",
    "Mástil de Pararrayos PDC 18m",
    "Torre de Comunicaciones",
]

# Factor s de Sulzberger en función de M1/M2 (tabla clásica)
TABLA_SULZBERGER = [(0.1, 1.500), (0.2, 1.383), (0.3, 1.317), (0.4, 1.260), (0.5, 1.208),
                    (0.6, 1.150), (0.7, 1.115), (0.8, 1.075), (0.9, 1.040), (1.0, 1.000)]

# Precios unitarios referenciales (S/.) para el presupuesto (editables en Excel)
PRECIOS = {
    "excavacion": 45.00, "eliminacion": 38.00, "relleno": 42.00, "solado": 285.00,
    "concreto_210": 520.00, "concreto_280": 585.00, "encofrado": 68.00, "acero": 6.80,
    "perno_eq": 145.00, "perno_poste": 185.00, "grout": 9500.00, "junta": 38.00,
    "sello": 30.00, "amortiguador": 680.00, "conduit": 28.00, "pvc_sap": 18.00,
}


# =============================================================================
# 2. UTILIDADES
# =============================================================================
def num(x, default=0.0):
    """Convierte texto a float tolerando comas y espacios."""
    try:
        if x is None:
            return default
        if isinstance(x, (int, float)):
            if isinstance(x, float) and math.isnan(x):
                return default
            return float(x)
        s = str(x).replace(",", ".").strip()
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group(0)) if m else default
    except Exception:
        return default


def f(x, d=2):
    """Formato numérico con separador de miles."""
    try:
        return f"{x:,.{d}f}"
    except Exception:
        return str(x)


def abrir_archivo(ruta):
    """Abre el archivo generado con la aplicación predeterminada del sistema."""
    try:
        if hasattr(os, "startfile"):
            os.startfile(ruta)  # Windows
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


def interp_tabla(tabla, x):
    if x <= tabla[0][0]:
        return tabla[0][1]
    if x >= tabla[-1][0]:
        return tabla[-1][1]
    for (x1, y1), (x2, y2) in zip(tabla[:-1], tabla[1:]):
        if x1 <= x <= x2:
            return y1 + (y2 - y1) * (x - x1) / (x2 - x1)
    return tabla[-1][1]


def espaciamiento(area_req_cm2_m, barra, s_max=0.30, s_min=0.10):
    """Espaciamiento redondeado hacia abajo a múltiplos de 0.05 m."""
    ab = BARRAS[barra][1]
    if area_req_cm2_m <= 0:
        return s_max
    s = ab / area_req_cm2_m  # m
    s = math.floor(s / 0.05 + 1e-9) * 0.05
    return max(s_min, min(s, s_max))


# =============================================================================
# 3. DATOS DE EJEMPLO
# =============================================================================
def datos_proyecto_ejemplo():
    return {
        "nombre": "Centro de Salud - Sala de Máquinas y Losa Deportiva",
        "ubicacion": "Lima - Lima - Perú",
        "q_adm": "1.50",
        "gamma_s": "1.80",
        "Ct": "5.00",
        "fc": "280",
        "ingeniero": "Ing. Civil Responsable",
        "cip": "000000",
    }


def datos_equipo_ejemplo():
    return {
        "tipo_eq": "Grupo Electrógeno Diésel Encapsulado",
        "potencia": "250 kVA / 200 kW",
        "W_eq": "2450",
        "rpm": "1800 RPM - 60Hz",
        "L_eq": "3.20",
        "B_eq": "1.10",
        "h_cg": "0.90",
        "e_cg": "0.05",
        "L": "3.80",
        "B": "1.70",
        "H": "0.75",
        "h_sobre": "0.15",
        "aislamiento": AISLAMIENTOS[0],
        "defl_mm": "25",
        "Vs": "300",
        "nu": "0.33",
        "pct_rotor": "30",
        "G_bal": "6.3",
        "Sf": "2.0",
        "A_lim": "0.05",
        "Ch": "0.50",
        "n_pernos": "6",
        "d_perno": '3/4"',
        "grado": "ASTM F1554 Grado 36",
        "h_ef": "0.45",
        "e_grout": "3.0",
        "barra": '5/8"',
    }


def datos_poste_ejemplo():
    return {
        "tipo_poste": "Poste Metálico Octogonal con Corona de 6 Reflectores LED",
        "Hp": "15.00",
        "H_libre": "15.00",
        "d_top": "0.12",
        "d_base": "0.30",
        "W_poste": "650",
        "V": "90",
        "C_f": "0.70",
        "A_eq": "1.80",
        "C_eq": "1.50",
        "W_eqp": "200",
        "tiro": "0",
        "a": "1.00",
        "b": "1.00",
        "t": "1.80",
        "a_ped": "0.60",
        "h_ped": "0.50",
        "n_pp": "8",
        "d_pp": '1"',
        "Dc": "0.40",
        "Cb_fac": "1.20",
        "tan_adm": "0.01",
        "barra_p": '5/8"',
    }


def datos_poste_cac():
    return {
        "tipo_poste": "Poste de Concreto CAC 13m/400daN", "Hp": "13.00", "H_libre": "11.70",
        "d_top": "0.18", "d_base": "0.375", "W_poste": "1650", "V": "90", "C_f": "0.70",
        "A_eq": "0.30", "C_eq": "1.50", "W_eqp": "80", "tiro": "408", "a": "1.10", "b": "1.10",
        "t": "1.80", "a_ped": "0.70", "h_ped": "0.20", "n_pp": "0", "d_pp": '1"', "Dc": "0.40",
        "Cb_fac": "1.20", "tan_adm": "0.01", "barra_p": '1/2"'}


# =============================================================================
# 4. MÓDULO A — CIMENTACIÓN DE EQUIPOS VIBRATORIOS (ACI 351.3R / E.050 / E.060)
# =============================================================================
def calcular_equipo(P, E):
    R = {"ok": True, "errores": []}
    fc = num(P.get("fc"), 210)
    qadm = num(P.get("q_adm"), 1.0)          # kg/cm²
    gs = num(P.get("gamma_s"), 1.8)          # t/m³
    tipo = E.get("tipo_eq", list(TIPOS_EQUIPO)[0])
    info = TIPOS_EQUIPO.get(tipo, TIPOS_EQUIPO["Grupo Electrógeno Diésel Encapsulado"])
    R["tipo"], R["clase"], R["info"] = tipo, info["clase"], info

    W_kg = num(E.get("W_eq"), 1000)
    Weq = W_kg / 1000.0
    L, B, H = num(E.get("L"), 3), num(E.get("B"), 1.5), num(E.get("H"), 0.6)
    Leq, Beq = num(E.get("L_eq"), 2.5), num(E.get("B_eq"), 1.0)
    h_cg = num(E.get("h_cg"), 0.9)
    e_cg = num(E.get("e_cg"), 0.0)
    h_sobre = num(E.get("h_sobre"), 0.15)
    if min(L, B, H, Weq) <= 0:
        R["ok"] = False
        R["errores"].append("Dimensiones o peso no válidos.")
        return R
    R.update(dict(fc=fc, qadm=qadm, gs=gs, Weq=Weq, W_kg=W_kg, L=L, B=B, H=H, Leq=Leq, Beq=Beq,
                  h_cg=h_cg, e_cg=e_cg, h_sobre=h_sobre))

    # --- 4.1 Regla de masa -------------------------------------------------------------
    Wc = L * B * H * GAMMA_C
    ratio = Wc / Weq
    R.update(Wc=Wc, ratio=ratio, ratio_min=info["ratio_min"], ratio_rec=info["ratio_rec"],
             ok_masa=ratio >= info["ratio_min"])
    R["Wc_min"] = info["ratio_min"] * Weq
    R["H_masa"] = R["Wc_min"] / (L * B * GAMMA_C)

    # --- 4.2 Geometría ------------------------------------------------------------------
    borde_L = (L - Leq) / 2.0
    borde_B = (B - Beq) / 2.0
    R.update(borde_L=borde_L, borde_B=borde_B, ok_borde=min(borde_L, borde_B) >= 0.25 - 1e-6)
    H_min = max(0.60, 0.60 + L / 30.0)
    R.update(H_min=H_min, ok_H=H >= H_min - 1e-6)

    # --- 4.3 Parámetros dinámicos del suelo --------------------------------------------
    Vs, nu = num(E.get("Vs"), 250), num(E.get("nu"), 0.33)
    rho = gs / G_ACC                     # t·s²/m⁴
    G = rho * Vs ** 2                    # t/m²
    A = L * B
    r0 = math.sqrt(A / math.pi)
    D_emb = max(0.0, H - h_sobre)
    Kz0 = 4.0 * G * r0 / (1.0 - nu)
    eta_z = 1.0 + 0.6 * (1.0 - nu) * (D_emb / r0)
    Kz = Kz0 * eta_z
    Cz = Kz / A                          # t/m³
    Wt = Wc + Weq
    m = Wt / G_ACC                       # t·s²/m
    Bz = (1.0 - nu) / 4.0 * m / (rho * r0 ** 3)
    zeta_z0 = 0.425 / math.sqrt(Bz)
    zeta_z_emb = zeta_z0 * (1.0 + 1.9 * (1.0 - nu) * D_emb / r0) / math.sqrt(eta_z)
    zeta_z = min(zeta_z_emb, 0.50)
    R.update(Vs=Vs, nu=nu, rho=rho, G=G, A=A, r0=r0, D_emb=D_emb, Kz0=Kz0, eta_z=eta_z, Kz=Kz,
             Cz=Cz, Cz_kgcm3=Cz / 1000.0, Wt=Wt, m=m, Bz=Bz, zeta_z0=zeta_z0,
             zeta_z_emb=zeta_z_emb, zeta_z=zeta_z)

    # --- 4.4 Frecuencias ----------------------------------------------------------------
    fn = (1.0 / (2.0 * math.pi)) * math.sqrt(Kz * G_ACC / Wt)
    rpm = num(E.get("rpm"), 1800)
    f_red = 50.0 if "50" in str(E.get("rpm", "")).split("-")[-1] else 60.0
    if info["clase"] == "electromagnetica":
        fop = 2.0 * f_red
        R["fop_desc"] = f"2·f_red = 2×{f(f_red, 0)} Hz (magnetostricción del núcleo)"
    else:
        fop = rpm / 60.0
        R["fop_desc"] = f"RPM/60 = {f(rpm, 0)}/60"
    omega = 2.0 * math.pi * fop
    r = fop / fn
    resonancia = 0.70 <= r <= 1.40
    M = 1.0 / math.sqrt((1 - r ** 2) ** 2 + (2 * zeta_z * r) ** 2)
    R.update(fn=fn, rpm=rpm, f_red=f_red, fop=fop, omega=omega, r=r, resonancia=resonancia,
             ok_res=not resonancia, M=M)

    # Modo horizontal (deslizamiento) — informativo
    Kx0 = 32.0 * (1 - nu) * G * r0 / (7 - 8 * nu)
    eta_x = 1.0 + 0.55 * (2 - nu) * (D_emb / r0)
    Kx = Kx0 * eta_x
    Bx = (7 - 8 * nu) / (32 * (1 - nu)) * m / (rho * r0 ** 3)
    zeta_x = min(0.288 / math.sqrt(Bx), 0.50)
    fnx = math.sqrt(Kx / m) / (2 * math.pi)
    rx = fop / fnx
    # Modo de cabeceo alrededor del eje longitudinal (giro en dirección B)
    r_psi = (L * B ** 3 / (3 * math.pi)) ** 0.25
    Kpsi0 = 8.0 * G * r_psi ** 3 / (3.0 * (1 - nu))
    eta_psi = 1.0 + 1.2 * (1 - nu) * (D_emb / r_psi) + 0.2 * (2 - nu) * (D_emb / r_psi) ** 3
    Kpsi = Kpsi0 * eta_psi
    mc, meq = Wc / G_ACC, Weq / G_ACC
    I_psi = mc * (B ** 2 + H ** 2) / 12.0 + mc * (H / 2.0) ** 2 + \
        meq * (Beq ** 2) / 12.0 + meq * (H + h_cg) ** 2
    B_psi = 3 * (1 - nu) / 8.0 * I_psi / (rho * r_psi ** 5)
    zeta_psi = min(0.15 / ((1 + B_psi) * math.sqrt(B_psi)), 0.50)
    fn_psi = math.sqrt(Kpsi / I_psi) / (2 * math.pi)
    r_psi_rel = fop / fn_psi
    R.update(Kx0=Kx0, eta_x=eta_x, Kx=Kx, Bx=Bx, zeta_x=zeta_x, fnx=fnx, rx=rx,
             ok_rx=not (0.70 <= rx <= 1.40), r_psi=r_psi, Kpsi0=Kpsi0, eta_psi=eta_psi,
             Kpsi=Kpsi, I_psi=I_psi, B_psi=B_psi, zeta_psi=zeta_psi, fn_psi=fn_psi,
             r_rel_psi=r_psi_rel, ok_rpsi=not (0.70 <= r_psi_rel <= 1.40))

    # --- 4.5 Fuerza dinámica y amplitud -------------------------------------------------
    pct_rotor = num(E.get("pct_rotor"), info["pct_rotor"])
    G_bal = num(E.get("G_bal"), info["G_bal"])
    Sf = num(E.get("Sf"), 2.0)
    A_lim = num(E.get("A_lim"), 0.05)
    m_r = pct_rotor / 100.0 * W_kg                    # kg
    if info["clase"] == "electromagnetica":
        F0_N = pct_rotor / 100.0 * W_kg * KGF_N * Sf   # % del peso como fuerza armónica
        R["F0_desc"] = "F₀ = %W·W·S_f (fuerza magnetostrictiva equivalente)"
    else:
        F0_N = Sf * m_r * (G_bal / 1000.0) * omega       # N  (ISO 1940: e·ω = G)
        R["F0_desc"] = "F₀ = S_f·m_r·e·ω² = S_f·m_r·G·ω (ISO 1940)"
    F0 = F0_N / (KGF_N * 1000.0)                        # t
    A_est = F0 / Kz                                     # m
    A_din = A_est * M
    A_mm = A_din * 1000.0
    vel = A_din * omega * 1000.0                        # mm/s
    e_bal = (G_bal / omega * 1000.0) if omega > 0 else 0  # μm
    R.update(pct_rotor=pct_rotor, G_bal=G_bal, Sf=Sf, A_lim=A_lim, m_r=m_r, F0_N=F0_N, F0=F0,
             A_est=A_est, A_mm=A_mm, vel=vel, e_bal=e_bal, ok_amp=A_mm <= A_lim)

    # Aislamiento (transmisibilidad)
    ais = E.get("aislamiento", AISLAMIENTOS[0])
    con_ais = "sin amortiguadores" not in ais
    defl = num(E.get("defl_mm"), 25.0)
    if con_ais and defl > 0:
        f_iso = 15.76 / math.sqrt(defl)
        ri = fop / f_iso
        T = 1.0 / abs(ri ** 2 - 1.0) if abs(ri ** 2 - 1.0) > 1e-6 else 99.0
        T = min(T, 99.0)
    else:
        f_iso, ri, T = 0.0, 0.0, 1.0
    efic = max(0.0, (1.0 - T) * 100.0)
    R.update(aislamiento=ais, con_ais=con_ais, defl=defl, f_iso=f_iso, ri=ri, T=T, efic=efic,
             F_trans=F0 * T, A_iso_mm=A_mm * T, ok_iso=(not con_ais) or ri > math.sqrt(2) * 1.5)

    # --- 4.6 Presiones y excentricidad -------------------------------------------------
    e = Weq * e_cg / Wt
    q_est = Wt / A                                      # t/m²
    q_max = (Wt + F0) / A * (1.0 + 6.0 * e / L)
    q_min = (Wt - F0) / A * (1.0 - 6.0 * e / L)
    R.update(e=e, e_lim=L / 6.0, e_aci=0.05 * L, ok_e=(e < L / 6.0) and (e <= 0.05 * L),
             q_est=q_est, q_max=q_max, q_min=q_min, q_lim=0.50 * qadm * 10.0,
             ok_q=q_max <= 0.50 * qadm * 10.0)

    # --- 4.7 Acero de refuerzo (E.060 9.7 / ACI 351.3R) --------------------------------
    barra = E.get("barra", '5/8"')
    db, ab, wb = BARRAS.get(barra, BARRAS['5/8"'])
    rec = 0.075
    d = H - rec - db / 100.0
    As_tot = 0.0018 * 100.0 * H * 100.0                 # cm²/m (total en el espesor)
    As_cara = As_tot / 2.0
    s = espaciamiento(As_cara, barra, s_max=0.30)
    As_prov = ab / s
    v = max(borde_L, borde_B) + 0.10
    Mu = 1.5 * q_max * v ** 2 / 2.0                     # t·m/m
    As_flex = Mu * 1e5 / (0.9 * FY * 0.9 * d * 100.0)   # cm²/m
    R.update(barra=barra, db=db, ab=ab, wb=wb, rec=rec, d=d, As_tot=As_tot, As_cara=As_cara,
             s=s, As_prov=As_prov, v_vol=v, Mu=Mu, As_flex=As_flex,
             ok_acero=As_prov >= max(As_cara, As_flex) - 1e-6,
             rho_prov=2 * As_prov / (100.0 * H * 100.0))

    # --- 4.8 Pernos de anclaje (ACI 318-19 Cap. 17) -----------------------------------
    n = max(2, int(num(E.get("n_pernos"), 4)))
    if n % 2:
        n += 1
    dp = E.get("d_perno", '3/4"')
    pinfo = PERNOS.get(dp, PERNOS['3/4"'])
    da, Ase = pinfo["da"], pinfo["Ase"]
    grado = E.get("grado", "ASTM F1554 Grado 36")
    fya, futa = GRADOS_PERNO.get(grado, GRADOS_PERNO["ASTM F1554 Grado 36"])
    hef = num(E.get("h_ef"), 0.45)
    Ch = num(E.get("Ch"), 0.50)
    e_grout = num(E.get("e_grout"), 3.0)
    V_sis = Ch * Weq
    Vu = V_sis + F0                                      # t
    Vu_b = Vu / n * 1000.0                               # kg
    brazo = max(0.30, Beq - 0.10)
    Mvu = Vu * (h_cg + 0.35)
    T_tot = max(0.0, Mvu / brazo - 0.9 * Weq / 2.0)
    Nu_b = T_tot / (n / 2.0) * 1000.0                    # kg
    # Resistencias
    phiNsa = 0.75 * Ase * futa
    fc_MPa = fc * 0.0980665
    hef_mm = hef * 1000.0
    if hef_mm <= 280:
        Nb_N = 10.0 * math.sqrt(fc_MPa) * hef_mm ** 1.5
    else:
        Nb_N = 3.9 * math.sqrt(fc_MPa) * hef_mm ** (5.0 / 3.0)
    ca = max(0.05, min(borde_B, borde_L) + 0.0375)
    hef = max(hef, 0.05)
    ca_mm = ca * 1000.0
    ANco = 9.0 * hef_mm ** 2
    ANc = (min(ca_mm, 1.5 * hef_mm) + 1.5 * hef_mm) * (3.0 * hef_mm)
    psi_ed = 1.0 if ca_mm >= 1.5 * hef_mm else 0.7 + 0.3 * ca_mm / (1.5 * hef_mm)
    psi_cN = 1.25
    Ncb_N = min(ANc / ANco, 1.0) * psi_ed * psi_cN * Nb_N
    Ncb = Ncb_N / KGF_N
    phiNcb = 0.70 * Ncb
    eh = 4.5 * da
    Np = 0.9 * fc * eh * da
    phiNpn = 0.70 * 1.4 * Np
    phiNn = min(phiNsa, phiNcb, phiNpn)
    modo_N = ["acero", "arrancamiento del concreto", "extracción (pullout)"][
        [phiNsa, phiNcb, phiNpn].index(phiNn)]
    phiVsa = 0.65 * 0.6 * Ase * futa * 0.80
    le = min(hef_mm, 8 * da * 10.0)
    da_mm = da * 10.0
    Vb_N = min(0.6 * (le / da_mm) ** 0.2 * math.sqrt(da_mm) * math.sqrt(fc_MPa) * ca_mm ** 1.5,
               3.7 * math.sqrt(fc_MPa) * ca_mm ** 1.5)
    ha_mm = H * 1000.0
    psi_hV = max(1.0, math.sqrt(1.5 * ca_mm / ha_mm)) if ha_mm < 1.5 * ca_mm else 1.0
    Vcb = Vb_N * 1.4 * psi_hV / KGF_N
    phiVcb = 0.70 * Vcb
    kcp = 2.0 if hef_mm >= 65 else 1.0
    phiVcp = 0.70 * kcp * Ncb
    phiVn = min(phiVsa, phiVcb, phiVcp)
    modo_V = ["acero", "desprendimiento del borde", "desprendimiento por palanca (pryout)"][
        [phiVsa, phiVcb, phiVcp].index(phiVn)]
    rN = Nu_b / phiNn if phiNn > 0 else 9
    rV = Vu_b / phiVn if phiVn > 0 else 9
    if rN <= 0.2:
        inter = rV
        ok_int = rV <= 1.0
    elif rV <= 0.2:
        inter = rN
        ok_int = rN <= 1.0
    else:
        inter = rN + rV
        ok_int = inter <= 1.2
    hef_min = 12.0 * da / 100.0
    Fp = 0.50 * fya * Ase                               # pretensado (kg)
    torque = 0.20 * (da / 100.0) * Fp                   # kg·m
    R.update(n=n, dp=dp, da=da, Ase=Ase, grado=grado, fya=fya, futa=futa, hef=hef, Ch=Ch,
             e_grout=e_grout, V_sis=V_sis, Vu=Vu, Vu_b=Vu_b, brazo=brazo, Mvu=Mvu, T_tot=T_tot,
             Nu_b=Nu_b, phiNsa=phiNsa, fc_MPa=fc_MPa, Nb_N=Nb_N, ca=ca, ANc=ANc, ANco=ANco,
             psi_ed=psi_ed, psi_cN=psi_cN, Ncb=Ncb, phiNcb=phiNcb, eh=eh, Np=Np, phiNpn=phiNpn,
             phiNn=phiNn, modo_N=modo_N, phiVsa=phiVsa, le=le, Vb_N=Vb_N, psi_hV=psi_hV,
             Vcb=Vcb, phiVcb=phiVcb, kcp=kcp, phiVcp=phiVcp, phiVn=phiVn, modo_V=modo_V,
             rN=rN, rV=rV, inter=inter, ok_int=ok_int, hef_min=hef_min,
             ok_hef=(hef >= hef_min - 1e-6) and (hef <= H - 0.10 + 1e-6),
             ok_grout=2.5 <= e_grout <= 5.0, Fp=Fp, torque=torque)

    # Posiciones de pernos (coordenadas locales respecto al centro del bloque)
    npl = n // 2
    if npl == 1:
        xs = [0.0]
    else:
        x0 = -(Leq / 2.0 - 0.25)
        xs = [x0 + i * (2 * abs(x0)) / (npl - 1) for i in range(npl)]
    yrail = Beq / 2.0 - 0.0375
    R["pernos_xy"] = [(x, s_) for s_ in (-yrail, yrail) for x in xs]
    R["yrail"] = yrail

    checks = [R["ok_masa"], R["ok_borde"], R["ok_H"], R["ok_res"], R["ok_amp"], R["ok_e"],
              R["ok_q"], R["ok_acero"], R["ok_int"], R["ok_hef"], R["ok_grout"]]
    R["cumple"] = all(checks)
    R["n_ok"] = sum(1 for c in checks if c)
    R["n_chk"] = len(checks)
    return R


def curva_amplificacion(zeta, rmax=3.0, n=31):
    filas = []
    for i in range(n):
        r = rmax * i / (n - 1)
        M = 1.0 / math.sqrt((1 - r ** 2) ** 2 + (2 * zeta * r) ** 2)
        filas.append((r, M))
    return filas


# =============================================================================
# 5. MÓDULO B — MACIZO PARA POSTES / TORRES (SULZBERGER + VIENTO E.020)
# =============================================================================
def sulz_M1(b, t, Ct, tan):
    return b * t ** 3 / 36.0 * Ct * tan


def sulz_M2(P, a, b, Cb, tan):
    """Momento resistente de fondo. Devuelve (M2, a0, contacto_total)."""
    if tan <= 0:
        return 0.0, 0.0, False
    a0 = math.sqrt(2.0 * P / (b * Cb * tan))
    if a0 <= a:
        M2 = P * a * (0.5 - (2.0 / 3.0) * math.sqrt(P / (2.0 * b * a ** 2 * Cb * tan)))
        return M2, a0, False
    return b * a ** 3 / 12.0 * Cb * tan, a, True


def calcular_poste(P, S):
    R = {"ok": True, "errores": []}
    qadm = num(P.get("q_adm"), 1.0)
    C2 = num(P.get("Ct"), 5.0)                       # kg/cm³ a 2 m
    fc = num(P.get("fc"), 210)
    tipo = S.get("tipo_poste", TIPOS_POSTE[1])
    Hp, Hl = num(S.get("Hp"), 15), num(S.get("H_libre"), 15)
    d_top, d_base = num(S.get("d_top"), 0.12), num(S.get("d_base"), 0.30)
    Wp, Weq = num(S.get("W_poste"), 600), num(S.get("W_eqp"), 0)
    V = num(S.get("V"), 90)
    Cf, Aeq, Ceq = num(S.get("C_f"), 0.7), num(S.get("A_eq"), 0), num(S.get("C_eq"), 1.5)
    tiro = num(S.get("tiro"), 0)
    a, b, t = num(S.get("a"), 1.0), num(S.get("b"), 1.0), num(S.get("t"), 1.8)
    a_ped, h_ped = num(S.get("a_ped"), 0.6), num(S.get("h_ped"), 0.5)
    Cb_fac = num(S.get("Cb_fac"), 1.2)
    tan_adm = num(S.get("tan_adm"), 0.01)
    if min(Hl, a, b, t, V) <= 0:
        R["ok"] = False
        R["errores"].append("Datos del poste no válidos.")
        return R
    empotrado = "Concreto" in tipo
    R.update(tipo=tipo, Hp=Hp, Hl=Hl, d_top=d_top, d_base=d_base, Wp=Wp, Weq=Weq, V=V, Cf=Cf,
             Aeq=Aeq, Ceq=Ceq, tiro=tiro, a=a, b=b, t=t, a_ped=a_ped, h_ped=h_ped,
             Cb_fac=Cb_fac, tan_adm=tan_adm, C2=C2, qadm=qadm, fc=fc, empotrado=empotrado)

    # --- 5.1 Viento NTE E.020 sobre el fuste (integración por tramos) ------------------
    nseg = max(10, int(math.ceil(Hl / 0.50)))
    dh = Hl / nseg
    tramos = []
    F_f, M_f = 0.0, 0.0
    for i in range(nseg):
        z = (i + 0.5) * dh                        # sobre la base del poste
        zg = z + h_ped                            # sobre el terreno
        dz = d_base + (d_top - d_base) * z / Hl
        Vh = V * (zg / 10.0) ** 0.22
        Ph = 0.005 * Cf * Vh ** 2
        Fi = Ph * dz * dh
        F_f += Fi
        M_f += Fi * z
        tramos.append(dict(i=i + 1, z=z, zg=zg, d=dz, Vh=Vh, Ph=Ph, A=dz * dh, F=Fi, M=Fi * z))
    zg_top = Hl + h_ped
    Vh_top = V * (zg_top / 10.0) ** 0.22
    P_eq = 0.005 * Ceq * Vh_top ** 2
    F_eq = P_eq * Aeq
    M_eq = F_eq * Hl
    h_tiro = max(0.0, Hl - 0.15) if tiro > 0 else Hl
    M_tiro = tiro * h_tiro
    # pedestal (C = 1.5 superficie plana)
    Vh_ped = V * (max(h_ped, 0.01) / 2.0 / 10.0) ** 0.22
    F_ped = 0.005 * 1.5 * Vh_ped ** 2 * a_ped * h_ped
    Ht = (F_f + F_eq + tiro + F_ped) / 1000.0                 # t
    M_base = (M_f + M_eq + M_tiro) / 1000.0                   # t·m (sobre pedestal)
    M_suelo = M_base + (F_f + F_eq + tiro) / 1000.0 * h_ped + F_ped / 1000.0 * h_ped / 2.0
    Mv = M_suelo + Ht * (2.0 * t / 3.0)                        # eje de giro a 2t/3
    R.update(tramos=tramos, F_f=F_f, M_f=M_f, Vh_top=Vh_top, P_eq=P_eq, F_eq=F_eq, M_eq=M_eq,
             h_tiro=h_tiro, M_tiro=M_tiro, F_ped=F_ped, Ht=Ht, M_base=M_base, M_suelo=M_suelo,
             Mv=Mv, zg_top=zg_top)

    # --- 5.2 Cargas verticales -----------------------------------------------------------
    W_maz = a * b * t * GAMMA_C
    W_ped = a_ped * a_ped * h_ped * GAMMA_C
    Pv = (Wp + Weq) / 1000.0 + W_maz + W_ped
    R.update(W_maz=W_maz, W_ped=W_ped, Pv=Pv)

    # --- 5.3 Sulzberger -------------------------------------------------------------------
    Ct = C2 * 1000.0 * t / 2.0                     # t/m³ (crece linealmente con la profundidad)
    Cb = Cb_fac * Ct
    M1 = sulz_M1(b, t, Ct, tan_adm)
    M2, a0, total = sulz_M2(Pv, a, b, Cb, tan_adm)
    FS = (M1 + M2) / Mv if Mv > 0 else 99.0
    rel = M1 / M2 if M2 > 0 else 10.0
    s_sulz = interp_tabla(TABLA_SULZBERGER, rel)
    # Giro real bajo Mv (bisección)
    def cap(x):
        return sulz_M1(b, t, Ct, x) + sulz_M2(Pv, a, b, Cb, x)[0]
    lo, hi = 1e-7, 0.10
    if cap(hi) < Mv:
        tan_req = hi
        conv = False
    else:
        conv = True
        for _ in range(80):
            mid = 0.5 * (lo + hi)
            if cap(mid) >= Mv:
                hi = mid
            else:
                lo = mid
        tan_req = hi
    M2r, a0r, total_r = sulz_M2(Pv, a, b, Cb, tan_req)
    sig_lat = Ct * t * tan_req / 9.0                # t/m² (máx. zona superior, a t/3)
    sig_lat_inf = Ct * t * tan_req / 3.0            # t/m² (máx. zona inferior, en el fondo)
    sig_lat_adm = Ct * t * tan_adm / 9.0
    sig_fondo = Cb * tan_req * a0r if not total_r else 2 * Pv / (a * b)
    sig_med = Pv / (a * b)
    R.update(Ct=Ct, Cb=Cb, M1=M1, M2=M2, a0=a0, contacto_total=total, FS=FS, rel=rel,
             s_sulz=s_sulz, ok_FS=FS >= 1.50, ok_s=(M1 + M2) >= s_sulz * Mv,
             tan_req=tan_req, conv=conv, ok_tan=conv and tan_req <= tan_adm + 1e-12,
             M1r=sulz_M1(b, t, Ct, tan_req), M2r=M2r, a0r=a0r, sig_lat=sig_lat, sig_lat_inf=sig_lat_inf,
             sig_lat_adm=sig_lat_adm, sig_fondo=sig_fondo, sig_med=sig_med,
             ok_q=sig_med <= qadm * 10.0, ok_qf=sig_fondo <= 1.33 * qadm * 10.0 * 1.5)

    # --- 5.4 Pedestal y armadura -------------------------------------------------------
    barra = S.get("barra_p", '5/8"')
    db, ab, wb = BARRAS.get(barra, BARRAS['5/8"'])
    d_ped = a_ped - 0.05 - 0.0095 - db / 200.0
    Mu_ped = 1.25 * (M_base + (F_f + F_eq + tiro) / 1000.0 * h_ped)
    As_ped = Mu_ped * 1e5 / (0.9 * FY * 0.9 * d_ped * 100.0) if d_ped > 0 else 0
    As_min_ped = 0.005 * (a_ped * 100) ** 2
    n_cara = max(3, int(math.ceil(max(As_ped, As_min_ped / 4.0 + 1e-9) / ab)))
    n_ped = 4 * (n_cara - 1)
    while n_ped * ab < As_min_ped:
        n_cara += 1
        n_ped = 4 * (n_cara - 1)
    As_maz = 0.0018 * 100.0 * 40.0 / 2.0
    s_maz = espaciamiento(As_maz, barra, s_max=0.25)
    R.update(barra=barra, db=db, ab=ab, wb=wb, d_ped=d_ped, Mu_ped=Mu_ped, As_ped=As_ped,
             As_min_ped=As_min_ped, n_cara=n_cara, n_ped=n_ped, As_maz=As_maz, s_maz=s_maz,
             ok_ped=n_cara * ab >= As_ped)

    # --- 5.5 Pernos de la canastilla / empotramiento ----------------------------------
    n_pp = int(num(S.get("n_pp"), 0))
    d_pp = S.get("d_pp", '1"')
    Dc = num(S.get("Dc"), 0.40)
    pinfo = PERNOS.get(d_pp, PERNOS['1"'])
    fya, futa = GRADOS_PERNO["ASTM F1554 Grado 36"]
    R.update(n_pp=n_pp, d_pp=d_pp, Dc=Dc, da_p=pinfo["da"], Ase_p=pinfo["Ase"])
    if n_pp > 0 and not empotrado:
        Tb = (1.25 * 4.0 * M_base / (n_pp * Dc)) * 1000.0 - 0.9 * (Wp + Weq) / n_pp
        Vb = 1.25 * Ht * 1000.0 / n_pp
        phiN = 0.75 * pinfo["Ase"] * futa
        phiV = 0.65 * 0.6 * pinfo["Ase"] * futa * 0.80
        inter = max(Tb, 0) / phiN + Vb / phiV
        L_anc = max(0.60, 25 * pinfo["da"] / 100.0)
        R.update(Tb=max(Tb, 0), Vb_p=Vb, phiN_p=phiN, phiV_p=phiV, inter_p=inter,
                 ok_pp=inter <= 1.2, L_anc=L_anc, ok_Lanc=L_anc <= h_ped + t - 0.15,
                 e_emp=0.0, e_min=0.0, ok_emp=True)
    else:
        e_emp = Hp - Hl
        e_min = Hp / 10.0
        R.update(Tb=0, Vb_p=0, phiN_p=0, phiV_p=0, inter_p=0, ok_pp=True, L_anc=0,
                 ok_Lanc=True, e_emp=e_emp, e_min=e_min,
                 ok_emp=(e_emp >= e_min - 1e-6) and (e_emp <= t + h_ped - 0.20))

    checks = [R["ok_FS"], R["ok_tan"], R["ok_q"], R["ok_ped"], R["ok_pp"], R["ok_emp"],
              R["ok_Lanc"]]
    R["cumple"] = all(checks)
    R["n_ok"] = sum(1 for c in checks if c)
    R["n_chk"] = len(checks)
    return R


# =============================================================================
# 6. METRADOS Y PRESUPUESTO
# =============================================================================
def calcular_metrados(P, RE, RP):
    items = []
    fc = int(num(P.get("fc"), 210))
    pu_c = PRECIOS["concreto_280"] if fc >= 280 else PRECIOS["concreto_210"]
    if RE.get("ok"):
        L, B, H = RE["L"], RE["B"], RE["H"]
        Demb = RE["D_emb"]
        exc = (L + 0.60) * (B + 0.60) * (Demb + 0.10 + 0.05)
        vol = L * B * H
        sol = (L + 0.20) * (B + 0.20) * 0.10
        rec, s = RE["rec"], RE["s"]
        pata = 0.20
        nL = int((B - 2 * rec) / s) + 1
        nB = int((L - 2 * rec) / s) + 1
        lgL = L - 2 * rec + 2 * pata
        lgB = B - 2 * rec + 2 * pata
        n_lat = max(1, int((H - 2 * rec) / s) - 1)
        lg_lat = 2 * (L + B) - 8 * rec + 0.60
        kg = (2 * (nL * lgL + nB * lgB) + n_lat * lg_lat) * RE["wb"]
        items += [
            ("A", "BANCADA ANTIVIBRATORIA DEL EQUIPO", "", 0, 0),
            ("A.01", "Excavación manual/máquina para bancada", "m³", exc, PRECIOS["excavacion"]),
            ("A.02", "Eliminación de material excedente (esponj. 25 %)", "m³", exc * 1.25 - 0 * vol,
             PRECIOS["eliminacion"]),
            ("A.03", "Solado de concreto C:H 1:12 e=0.10 m", "m²", (L + 0.2) * (B + 0.2),
             PRECIOS["solado"] * 0.10),
            ("A.04", f"Concreto f'c={fc} kg/cm² en bloque de cimentación", "m³", vol, pu_c),
            ("A.05", "Encofrado y desencofrado de bloque", "m²", 2 * (L + B) * H, PRECIOS["encofrado"]),
            ("A.06", f"Acero de refuerzo fy=4200 (Ø{RE['barra']} @ {f(s)} doble parrilla)", "kg",
             kg, PRECIOS["acero"]),
            ("A.07", f"Perno de anclaje tipo J Ø{RE['dp']} {RE['grado']} L={f(RE['hef'] + 0.20)} m",
             "und", RE["n"], PRECIOS["perno_eq"]),
            ("A.08", f"Grout no contráctil e={f(RE['e_grout'], 1)} cm bajo placas", "m³",
             RE["n"] * 0.25 * 0.25 * RE["e_grout"] / 100.0, PRECIOS["grout"]),
            ("A.09", 'Junta perimetral poliestireno elastificado/neopreno e=1"', "m",
             2 * (L + B) + 0.20, PRECIOS["junta"]),
            ("A.10", "Sello elastomérico de poliuretano en junta", "m", 2 * (L + B) + 0.20,
             PRECIOS["sello"]),
            ("A.11", "Amortiguador antivibratorio (spring/elastómero)", "und",
             RE["n"] if RE["con_ais"] else 0, PRECIOS["amortiguador"]),
            ("A.12", 'Tubo conduit PVC-SAP Ø4" para salida de cables de fuerza', "m", 3.00,
             PRECIOS["conduit"]),
        ]
        RE["kg_acero"] = kg
        RE["vol_conc"] = vol
        RE["vol_sol"] = sol
        RE["vol_exc"] = exc
    if RP.get("ok"):
        a, b, t, ap, hp = RP["a"], RP["b"], RP["t"], RP["a_ped"], RP["h_ped"]
        exc = (a + 0.40) * (b + 0.40) * (t + 0.10)
        v_maz = a * b * t
        v_ped = ap * ap * hp
        s = RP["s_maz"]
        rec = 0.075
        n1 = int((b - 2 * rec) / s) + 1
        n2 = int((a - 2 * rec) / s) + 1
        kg_maz = 2 * (n1 * (a - 2 * rec + 0.30) + n2 * (b - 2 * rec + 0.30)) * RP["wb"]
        lg_v = hp + t - 0.10 + 0.25
        kg_ped = RP["n_ped"] * lg_v * RP["wb"]
        n_est = int((hp + t - 0.15) / 0.15) + 1
        kg_est = n_est * (4 * (ap - 0.08) + 0.25) * BARRAS['3/8"'][2]
        kg = kg_maz + kg_ped + kg_est
        items += [
            ("B", "MACIZO DE CIMENTACIÓN DE POSTE / TORRE", "", 0, 0),
            ("B.01", "Excavación de macizo", "m³", exc, PRECIOS["excavacion"]),
            ("B.02", "Relleno compactado con material propio", "m³", max(exc - v_maz, 0),
             PRECIOS["relleno"]),
            ("B.03", "Solado C:H 1:12 e=0.10 m", "m²", (a + 0.2) * (b + 0.2), PRECIOS["solado"] * 0.10),
            ("B.04", f"Concreto f'c={fc} kg/cm² en macizo", "m³", v_maz, pu_c),
            ("B.05", f"Concreto f'c={fc} kg/cm² en pedestal con chaflán", "m³", v_ped, pu_c),
            ("B.06", "Encofrado y desencofrado de pedestal", "m²", 4 * ap * (hp + 0.20),
             PRECIOS["encofrado"]),
            ("B.07", "Acero de refuerzo fy=4200 (macizo + pedestal + estribos)", "kg", kg,
             PRECIOS["acero"]),
            ("B.08", f"Pernos de canastilla Ø{RP['d_pp']} F1554 Gr.36 (juego)", "und",
             RP["n_pp"] if not RP["empotrado"] else 0, PRECIOS["perno_poste"]),
            ("B.09", "Grout no contráctil bajo plancha base e=3 cm", "m³",
             0 if RP["empotrado"] else (RP["Dc"] + 0.15) ** 2 * 0.03, PRECIOS["grout"]),
            ("B.10", 'Tubo PVC-SAP Ø1" bajada para puesta a tierra', "m", t + hp + 1.50,
             PRECIOS["pvc_sap"]),
        ]
        RP["kg_acero"] = kg
        RP["vol_conc"] = v_maz + v_ped
        RP["vol_exc"] = exc
    return items


# =============================================================================
# 7. EXPORTACIÓN A WORD (.DOCX) — MEMORIA DE CÁLCULO
# =============================================================================
class Memoria:
    def __init__(self):
        self.doc = docx.Document()
        try:  # actualizar índice y campos al abrir en Word
            uf = OxmlElement("w:updateFields")
            uf.set(qn("w:val"), "true")
            self.doc.settings.element.append(uf)
        except Exception:
            pass
        sec = self.doc.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
        sec.left_margin = sec.right_margin = Cm(2.5)
        sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)
        st = self.doc.styles["Normal"]
        st.font.name = "Calibri"
        st.font.size = Pt(10.5)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        for nivel, tam in ((1, 15), (2, 12.5), (3, 11)):
            hs = self.doc.styles[f"Heading {nivel}"]
            hs.font.name = "Calibri"
            hs.font.size = Pt(tam)
            hs.font.color.rgb = RGBColor(0x0B, 0x3D, 0x6B)

    # --- utilidades de formato ---------------------------------------------------------
    @staticmethod
    def rich(par, texto, bold=False, italic=False, size=None, color=None, font=None):
        """Texto con sub/superíndices: X_{sub} y X^{sup}."""
        partes = re.split(r"(_\{[^}]*\}|\^\{[^}]*\})", texto)
        for p in partes:
            if not p:
                continue
            if p.startswith("_{"):
                run = par.add_run(p[2:-1])
                run.font.subscript = True
            elif p.startswith("^{"):
                run = par.add_run(p[2:-1])
                run.font.superscript = True
            else:
                run = par.add_run(p)
            run.bold, run.italic = bold, italic
            if size:
                run.font.size = Pt(size)
            if color:
                run.font.color.rgb = RGBColor(*color)
            if font:
                run.font.name = font
        return par

    def h(self, txt, nivel=1):
        return self.doc.add_heading(txt, level=nivel)

    def p(self, txt, bold=False, italic=False, align=None, size=None):
        par = self.doc.add_paragraph()
        self.rich(par, txt, bold=bold, italic=italic, size=size)
        par.paragraph_format.space_after = Pt(4)
        if align == "j":
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif align == "c":
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return par

    def eq(self, txt):
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_after = Pt(3)
        self.rich(par, txt, italic=True, font="Cambria Math", size=11)
        return par

    def bullet(self, txt):
        par = self.doc.add_paragraph(style="List Bullet")
        self.rich(par, txt)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return par

    def check(self, txt, ok):
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.rich(par, txt + "   →   ", bold=True)
        self.rich(par, "CUMPLE ✔" if ok else "NO CUMPLE ✘", bold=True,
                  color=(0x1E, 0x84, 0x49) if ok else (0xC0, 0x39, 0x2B))
        return par

    @staticmethod
    def sombrear(celda, color):
        tcPr = celda._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    def tabla(self, filas, encab=None, anchos=None, size=9):
        nf = len(filas) + (1 if encab else 0)
        nc = len(encab) if encab else len(filas[0])
        tb = self.doc.add_table(rows=nf, cols=nc)
        tb.style = "Table Grid"
        tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        i0 = 0
        if encab:
            for j, e in enumerate(encab):
                c = tb.cell(0, j)
                c.text = ""
                self.rich(c.paragraphs[0], str(e), bold=True, size=size, color=(255, 255, 255))
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                self.sombrear(c, "0B3D6B")
            i0 = 1
        for i, fila in enumerate(filas):
            for j, v in enumerate(fila):
                c = tb.cell(i + i0, j)
                c.text = ""
                txt = str(v)
                col = None
                if txt.startswith("CUMPLE") or txt == "OK":
                    col = (0x1E, 0x84, 0x49)
                elif txt.startswith("NO CUMPLE") or txt.startswith("EN RESONANCIA"):
                    col = (0xC0, 0x39, 0x2B)
                self.rich(c.paragraphs[0], txt, size=size, color=col, bold=col is not None)
                if j > 0:
                    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i % 2 == 1:
                for j in range(nc):
                    self.sombrear(tb.cell(i + i0, j), "EAF1F8")
        if anchos:
            tb.autofit = False
            grid = tb._tbl.tblGrid
            for j, w in enumerate(anchos):
                if j < len(grid.gridCol_lst):
                    grid.gridCol_lst[j].w = Cm(w)
                tb.columns[j].width = Cm(w)
                for row in tb.rows:
                    row.cells[j].width = Cm(w)
        self.doc.add_paragraph()
        return tb

    def salto(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def campo(self, par, instr):
        run = par.add_run()
        f1 = OxmlElement("w:fldChar")
        f1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        f2 = OxmlElement("w:fldChar")
        f2.set(qn("w:fldCharType"), "separate")
        t = OxmlElement("w:t")
        t.text = "Actualice el campo (F9)" if "TOC" in instr else "1"
        f3 = OxmlElement("w:fldChar")
        f3.set(qn("w:fldCharType"), "end")
        run._r.append(f1)
        run._r.append(it)
        run._r.append(f2)
        run._r.append(t)
        run._r.append(f3)


def ok_txt(c):
    return "CUMPLE" if c else "NO CUMPLE"


def exportar_word(ruta, P, RE, RP):
    if docx is None:
        raise RuntimeError("Falta python-docx (pip install python-docx)")
    m = Memoria()
    d = m.doc
    hoy = datetime.date.today().strftime("%d/%m/%Y")
    items = calcular_metrados(P, RE, RP)

    # Encabezado y pie
    sec = d.sections[0]
    hp = sec.header.paragraphs[0]
    hp.text = f"MEMORIA DE CÁLCULO – CIMENTACIONES DE EQUIPOS Y POSTES · {P.get('nombre', '')}"
    hp.runs[0].font.size = Pt(8)
    hp.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ = fp.add_run(f"{P.get('ingeniero', '')} – CIP {P.get('cip', '')}      Página ")
    r_.font.size = Pt(8)
    m.campo(fp, "PAGE")

    # ---------------- PORTADA ----------------
    for _ in range(4):
        d.add_paragraph()
    m.p("MEMORIA DE CÁLCULO ESTRUCTURAL, GEOTÉCNICA Y DINÁMICA", bold=True, align="c", size=20)
    m.p("CIMENTACIÓN ANTIVIBRATORIA DE EQUIPOS (ACI 351.3R-04) Y MACIZOS DE FUNDACIÓN PARA "
        "POSTES / TORRES (MÉTODO DE SULZBERGER)", bold=True, align="c", size=13)
    d.add_paragraph()
    m.p(f"PROYECTO: {P.get('nombre', '')}", bold=True, align="c", size=13)
    m.p(f"UBICACIÓN: {P.get('ubicacion', '')}", align="c", size=12)
    d.add_paragraph()
    if RE.get("ok"):
        m.p(f"Equipo: {RE['tipo']} – {P.get('_potencia', '')} – W = {f(RE['W_kg'], 0)} kg – "
            f"{f(RE['rpm'], 0)} RPM", align="c")
    if RP.get("ok"):
        m.p(f"Estructura: {RP['tipo']} – H = {f(RP['Hp'])} m – V = {f(RP['V'], 0)} km/h", align="c")
    for _ in range(6):
        d.add_paragraph()
    m.tabla([["Ingeniero responsable", P.get("ingeniero", "")], ["Registro CIP", P.get("cip", "")],
             ["Fecha", hoy], ["Lámina asociada", "CE-01 (Planta, cortes y detalles)"],
             ["Software", f"Diseñador Pro de Cimentaciones v{APP_VERSION}"]], anchos=[6, 9], size=10)
    m.salto()

    # ---------------- ÍNDICE ----------------
    m.h("ÍNDICE", 1)
    m.campo(d.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
    m.p("(En Word: clic derecho sobre el índice → Actualizar campo → Actualizar toda la tabla.)",
        italic=True, size=9)
    m.salto()

    # ---------------- 1. GENERALIDADES ----------------
    m.h("1. GENERALIDADES", 1)
    m.h("1.1 Objeto", 2)
    m.p("La presente memoria desarrolla el análisis y diseño de: (a) la bancada o dado de cimentación "
        "antivibratorio que soporta un equipo con partes en movimiento (grupo electrógeno, compresor, "
        "bomba o transformador), y (b) el macizo de cimentación de un poste, mástil o torre sometido "
        "a cargas de viento. Para el primero se aplica la metodología de ACI 351.3R-04 *Foundations "
        "for Dynamic Equipment* con la teoría del semiespacio elástico (Richart–Lysmer–Barkan); para el "
        "segundo se aplica el Método de Sulzberger, de uso tradicional en las fundaciones de postes de "
        "líneas eléctricas y alumbrado, con cargas de viento según la NTE E.020.", align="j")
    m.h("1.2 Alcance", 2)
    for t in ["Pre-dimensionamiento por regla de masa y borde libre alrededor del chasis (skid).",
              "Verificación de presiones estáticas + dinámicas transmitidas al suelo y excentricidad.",
              "Determinación de la rigidez dinámica del suelo, del amortiguamiento geométrico y de las "
              "frecuencias naturales vertical, horizontal y de cabeceo del sistema bloque–suelo.",
              "Verificación de ausencia de resonancia, factor de amplificación dinámica y amplitud "
              "máxima de vibración frente al límite admisible.",
              "Diseño del acero de refuerzo del bloque y de los pernos de anclaje (ACI 318-19 Cap. 17).",
              "Cálculo de la presión de viento sobre el fuste y los equipos en punta, momento de vuelco "
              "y diseño del macizo por Sulzberger con giro admisible tan α ≤ 0.01.",
              "Metrados, presupuesto referencial, conclusiones y recomendaciones constructivas."]:
        m.bullet(t)
    m.h("1.3 Descripción del proyecto", 2)
    m.p(f"El proyecto «{P.get('nombre', '')}» se ubica en {P.get('ubicacion', '')}. La sala de máquinas "
        f"alberga el equipo descrito en el capítulo 4, que se apoya sobre un bloque de concreto armado "
        f"independizado del piso terminado mediante una junta perimetral de 1\" (2.5 cm). En el área "
        f"exterior se ubica la estructura descrita en el capítulo 5, cimentada en un macizo de concreto "
        f"armado con pedestal.", align="j")

    # ---------------- 2. MARCO NORMATIVO ----------------
    m.h("2. MARCO NORMATIVO", 1)
    normas = [
        ["NTE E.020 Cargas", "Presión dinámica del viento P_{h} = 0.005·C·V_{h}^{2} y variación con la altura."],
        ["NTE E.030 Diseño Sismorresistente", "Fuerzas sísmicas sobre equipos y elementos no estructurales (Cap. 6)."],
        ["NTE E.050 Suelos y Cimentaciones", "Capacidad portante, asentamientos y parámetros geotécnicos."],
        ["NTE E.060 Concreto Armado", "Cuantías mínimas, recubrimientos, diseño por resistencia."],
        ["ACI 351.3R-04", "Foundations for Dynamic Equipment: criterios de masa, rigidez, frecuencias, amplitudes."],
        ["ACI 318-19 Cap. 17", "Anclajes al concreto: tracción, corte e interacción de pernos."],
        ["ISO 1940-1 / ISO 10816", "Grado de balanceo G de rotores y severidad de vibración."],
        ["CNE Suministro 2011", "Cimentación de postes y estructuras de soporte (Sulzberger)."],
        ["ASTM F1554 / A307 / F436 / A563", "Pernos de anclaje, arandelas y tuercas."],
        ["ASTM C1107", "Grout cementicio no contráctil para nivelación de bases."],
    ]
    m.tabla(normas, ["Norma", "Aplicación en el presente cálculo"], anchos=[5, 11])
    m.p("Cuando la normativa peruana no regula un aspecto específico de la dinámica de máquinas, se "
        "adoptan las recomendaciones de ACI 351.3R y de la literatura técnica (Richart, Hall & Woods, "
        "*Vibrations of Soils and Foundations*; Barkan, *Dynamics of Bases and Foundations*; Arya, "
        "O'Neill & Pincus, *Design of Structures and Foundations for Vibrating Machines*).", align="j")

    # ---------------- 3. PARÁMETROS ----------------
    m.h("3. PARÁMETROS GEOTÉCNICOS Y MATERIALES", 1)
    fc = num(P.get("fc"), 210)
    Ec = 15000 * math.sqrt(fc)
    m.tabla([
        ["Capacidad portante admisible q_{adm}", f"{f(num(P.get('q_adm')))} kg/cm²", "EMS"],
        ["Peso específico del suelo γ_{s}", f"{f(num(P.get('gamma_s')))} t/m³", "EMS"],
        ["Coeficiente de balasto lateral a 2 m C_{2}", f"{f(num(P.get('Ct')))} kg/cm³", "EMS / Sulzberger"],
        ["Velocidad de onda de corte V_{s}", f"{f(RE.get('Vs', 0), 0)} m/s", "EMS / ensayo MASW"],
        ["Coeficiente de Poisson ν", f"{f(RE.get('nu', 0))}", "EMS"],
        ["Concreto f'c", f"{f(fc, 0)} kg/cm²", "E.060"],
        ["Módulo de elasticidad E_{c} = 15000√f'c", f"{f(Ec, 0)} kg/cm²", "E.060 8.5"],
        ["Peso específico del concreto armado γ_{c}", f"{f(GAMMA_C)} t/m³", "E.020"],
        ["Acero de refuerzo f_{y}", f"{f(FY, 0)} kg/cm²", "ASTM A615 Gr. 60"],
        ["Recubrimiento (concreto contra el terreno)", "7.5 cm", "E.060 7.7"],
    ], ["Parámetro", "Valor", "Fuente"], anchos=[8, 4, 4])
    m.p("El módulo de corte dinámico del suelo se obtiene a partir de la velocidad de onda de corte "
        "(pequeñas deformaciones), lo que es consistente con el nivel de deformación angular que "
        "producen las máquinas en operación normal (γ < 10^{-5}).", align="j")
    m.salto()

    # ---------------- 4. MÓDULO A ----------------
    if RE.get("ok"):
        R = RE
        m.h("4. DISEÑO DE LA BANCADA ANTIVIBRATORIA DEL EQUIPO (ACI 351.3R-04)", 1)
        m.h("4.1 Datos del equipo y del bloque", 2)
        m.tabla([
            ["Tipo de equipo", R["tipo"]], ["Clase dinámica", R["info"]["desc"]],
            ["Potencia", P.get("_potencia", "")],
            ["Peso del equipo con fluidos W_{eq}", f"{f(R['W_kg'], 0)} kg = {f(R['Weq'], 3)} t"],
            ["Velocidad de operación", f"{f(R['rpm'], 0)} RPM → f_{{op}} = {f(R['fop'])} Hz"],
            ["Chasis (skid) L_{eq} × B_{eq}", f"{f(R['Leq'])} × {f(R['Beq'])} m"],
            ["Altura del centro de gravedad h_{cg}", f"{f(R['h_cg'])} m"],
            ["Bloque L × B × H", f"{f(R['L'])} × {f(R['B'])} × {f(R['H'])} m"],
            ["Altura sobre NPT / empotramiento D", f"{f(R['h_sobre'])} m / {f(R['D_emb'])} m"],
            ["Sistema de aislamiento", R["aislamiento"]],
        ], ["Dato", "Valor"], anchos=[6.5, 9.5])

        m.h("4.2 Criterio de masa (relación Wc/Weq)", 2)
        m.p("ACI 351.3R recomienda que el bloque tenga un peso varias veces superior al de la máquina para "
            "reducir las amplitudes y bajar el centro de gravedad del conjunto. Se adoptan los valores "
            "mínimos: W_{c}/W_{eq} ≥ 2.5 a 3.0 para máquinas rotativas y ≥ 4.0 a 5.0 para máquinas "
            "reciprocantes.", align="j")
        m.eq("W_{c} = L · B · H · γ_{c}")
        m.eq(f"W_{{c}} = {f(R['L'])} × {f(R['B'])} × {f(R['H'])} × {f(GAMMA_C)} = {f(R['Wc'], 3)} t")
        m.eq(f"W_{{c}}/W_{{eq}} = {f(R['Wc'], 3)} / {f(R['Weq'], 3)} = {f(R['ratio'])}  ≥  "
             f"{f(R['ratio_min'], 1)} (recomendado {f(R['ratio_rec'], 1)})")
        m.check(f"Relación de masas = {f(R['ratio'])}", R["ok_masa"])
        m.p(f"Espesor mínimo por masa: H_{{masa}} = {f(R['ratio_min'], 1)}·W_{{eq}}/(L·B·γ_{{c}}) = "
            f"{f(R['H_masa'], 3)} m.", align="j")

        m.h("4.3 Geometría, borde libre y espesor mínimo", 2)
        m.eq(f"b_{{L}} = (L − L_{{eq}})/2 = ({f(R['L'])} − {f(R['Leq'])})/2 = {f(R['borde_L'], 3)} m")
        m.eq(f"b_{{B}} = (B − B_{{eq}})/2 = ({f(R['B'])} − {f(R['Beq'])})/2 = {f(R['borde_B'], 3)} m")
        m.check("Borde libre ≥ 0.25 – 0.30 m", R["ok_borde"])
        m.eq(f"H_{{min}} = 0.60 + L/30 = 0.60 + {f(R['L'])}/30 = {f(R['H_min'], 3)} m  ≤  H = {f(R['H'])} m")
        m.check("Espesor del bloque (regla práctica ACI 351.3R)", R["ok_H"])

        m.h("4.4 Presión transmitida al suelo y excentricidad", 2)
        m.p("Para evitar asentamientos acumulativos inducidos por la vibración, la presión total "
            "(estática + dinámica) se limita al 50 % de la capacidad portante admisible estática. La "
            "excentricidad del centro de gravedad combinado se limita al 5 % de la dimensión en la "
            "dirección considerada (ACI 351.3R) y, en todo caso, al núcleo central (L/6).", align="j")
        m.eq(f"W_{{t}} = W_{{c}} + W_{{eq}} = {f(R['Wc'], 3)} + {f(R['Weq'], 3)} = {f(R['Wt'], 3)} t")
        m.eq(f"e = W_{{eq}}·e_{{cg}}/W_{{t}} = {f(R['Weq'], 3)} × {f(R['e_cg'], 3)} / {f(R['Wt'], 3)} = "
             f"{f(R['e'], 4)} m")
        m.eq(f"e = {f(R['e'], 4)} m  <  0.05·L = {f(R['e_aci'], 3)} m  y  <  L/6 = {f(R['e_lim'], 3)} m")
        m.check("Excentricidad", R["ok_e"])
        m.eq(f"q_{{est}} = W_{{t}}/(L·B) = {f(R['Wt'], 3)}/{f(R['A'], 3)} = {f(R['q_est'], 3)} t/m²")
        m.eq("q_{máx} = (W_{t} + F_{0})/(L·B) · (1 + 6e/L)")
        m.eq(f"q_{{máx}} = ({f(R['Wt'], 3)} + {f(R['F0'], 4)})/{f(R['A'], 3)} × (1 + 6×{f(R['e'], 4)}/"
             f"{f(R['L'])}) = {f(R['q_max'], 3)} t/m² = {f(R['q_max'] / 10, 3)} kg/cm²")
        m.eq(f"q_{{máx}} = {f(R['q_max'] / 10, 3)} kg/cm²  ≤  0.50·q_{{adm}} = 0.50 × {f(R['qadm'])} = "
             f"{f(R['q_lim'] / 10, 3)} kg/cm²")
        m.check("Presión estática + dinámica ≤ 50 % q_{adm}", R["ok_q"])

        m.h("4.5 Parámetros dinámicos del suelo y rigidez vertical", 2)
        m.p("Se modela el suelo como un semiespacio elástico, homogéneo e isótropo, y el bloque como un "
            "cuerpo rígido con una zapata circular equivalente de radio r_{0} (Lysmer, 1965; Richart, 1970).",
            align="j")
        m.eq("ρ = γ_{s}/g ;   G = ρ · V_{s}^{2}")
        m.eq(f"ρ = {f(R['gs'])}/{f(G_ACC)} = {f(R['rho'], 4)} t·s²/m⁴ ;   G = {f(R['rho'], 4)} × "
             f"{f(R['Vs'], 0)}² = {f(R['G'], 0)} t/m² = {f(R['G'] / 10, 1)} kg/cm²")
        m.eq("r_{0} = √(L·B/π)")
        m.eq(f"r_{{0}} = √({f(R['L'])} × {f(R['B'])}/π) = {f(R['r0'], 4)} m")
        m.eq("K_{z} = 4·G·r_{0}/(1 − ν)   (Lysmer – Richart)")
        m.eq(f"K_{{z0}} = 4 × {f(R['G'], 0)} × {f(R['r0'], 4)}/(1 − {f(R['nu'])}) = {f(R['Kz0'], 0)} t/m")
        m.p("Corrección por empotramiento del bloque (ACI 351.3R, Tabla 3.4):", align="j")
        m.eq("η_{z} = 1 + 0.6·(1 − ν)·(D/r_{0})")
        m.eq(f"η_{{z}} = 1 + 0.6 × (1 − {f(R['nu'])}) × ({f(R['D_emb'], 3)}/{f(R['r0'], 4)}) = "
             f"{f(R['eta_z'], 4)}")
        m.eq(f"K_{{z}} = η_{{z}}·K_{{z0}} = {f(R['eta_z'], 4)} × {f(R['Kz0'], 0)} = {f(R['Kz'], 0)} t/m")
        m.p("Coeficiente de balasto dinámico vertical equivalente (Barkan):", align="j")
        m.eq(f"C_{{z}} = K_{{z}}/(L·B) = {f(R['Kz'], 0)}/{f(R['A'], 3)} = {f(R['Cz'], 0)} t/m³ = "
             f"{f(R['Cz_kgcm3'], 2)} kg/cm³")

        m.h("4.6 Amortiguamiento geométrico (analogía de Lysmer)", 2)
        m.eq("B_{z} = (1 − ν)/4 · m/(ρ·r_{0}^{3}) ;   ζ_{z} = 0.425/√B_{z}")
        m.eq(f"m = W_{{t}}/g = {f(R['Wt'], 3)}/{f(G_ACC)} = {f(R['m'], 4)} t·s²/m")
        m.eq(f"B_{{z}} = (1 − {f(R['nu'])})/4 × {f(R['m'], 4)}/({f(R['rho'], 4)} × {f(R['r0'], 4)}³) = "
             f"{f(R['Bz'], 4)}")
        m.eq(f"ζ_{{z0}} = 0.425/√{f(R['Bz'], 4)} = {f(R['zeta_z0'], 4)}")
        m.eq("ζ_{z} = ζ_{z0}·[1 + 1.9·(1 − ν)·D/r_{0}]/√η_{z}")
        m.eq(f"ζ_{{z}} = {f(R['zeta_z0'], 4)} × [1 + 1.9 × {f(1 - R['nu'], 2)} × {f(R['D_emb'] / R['r0'], 4)}]"
             f"/√{f(R['eta_z'], 4)} = {f(R['zeta_z_emb'], 4)}")
        m.p(f"Por criterio conservador se limita la razón de amortiguamiento a 0.50: ζ_{{z}} de diseño = "
            f"{f(R['zeta_z'], 3)}.", align="j")

        m.h("4.7 Frecuencias naturales del sistema bloque–suelo", 2)
        m.eq("f_{n} = (1/2π) · √(K_{z}·g/W_{total})")
        m.eq(f"f_{{n}} = (1/2π) × √({f(R['Kz'], 0)} × {f(G_ACC)}/{f(R['Wt'], 3)}) = {f(R['fn'])} Hz")
        m.eq(f"f_{{op}} = {R['fop_desc']} = {f(R['fop'])} Hz ;   ω = 2π·f_{{op}} = {f(R['omega'])} rad/s")
        m.p("Modos complementarios (informativos, Richart–Hall–Woods):", bold=True)
        m.eq("K_{x} = 32·(1 − ν)·G·r_{0}/(7 − 8ν) · η_{x} ;  η_{x} = 1 + 0.55·(2 − ν)·D/r_{0}")
        m.eq(f"K_{{x}} = {f(R['Kx0'], 0)} × {f(R['eta_x'], 4)} = {f(R['Kx'], 0)} t/m ;   f_{{nx}} = "
             f"(1/2π)·√(K_{{x}}/m) = {f(R['fnx'])} Hz")
        m.eq("r_{ψ} = (L·B^{3}/3π)^{1/4} ;   K_{ψ} = 8·G·r_{ψ}^{3}/[3·(1 − ν)] · η_{ψ}")
        m.eq(f"r_{{ψ}} = {f(R['r_psi'], 4)} m ;   K_{{ψ}} = {f(R['Kpsi0'], 0)} × {f(R['eta_psi'], 4)} = "
             f"{f(R['Kpsi'], 0)} t·m/rad")
        m.eq(f"I_{{ψ}} = {f(R['I_psi'], 4)} t·s²·m ;   f_{{nψ}} = (1/2π)·√(K_{{ψ}}/I_{{ψ}}) = {f(R['fn_psi'])} Hz")

        m.h("4.8 Verificación de no resonancia", 2)
        m.p("ACI 351.3R recomienda que la frecuencia de operación se aleje al menos 20 – 30 % de cualquier "
            "frecuencia natural del sistema. Se adopta la zona crítica 0.70 ≤ r ≤ 1.40.", align="j")
        m.eq(f"r = f_{{op}}/f_{{n}} = {f(R['fop'])}/{f(R['fn'])} = {f(R['r'], 3)}")
        m.check(f"Modo vertical r = {f(R['r'], 3)} " + ("(FUERA DE RESONANCIA)" if R["ok_res"] else
                                                       "(EN ZONA DE RESONANCIA)"), R["ok_res"])
        m.tabla([
            ["Vertical", f(R["fn"]), f(R["r"], 3), ok_txt(R["ok_res"])],
            ["Horizontal (deslizamiento)", f(R["fnx"]), f(R["rx"], 3), ok_txt(R["ok_rx"])],
            ["Cabeceo (giro en B)", f(R["fn_psi"]), f(R["r_rel_psi"], 3), ok_txt(R["ok_rpsi"])],
        ], ["Modo", "f_{n} (Hz)", "r = f_{op}/f_{n}", "Fuera de 0.70–1.40"], anchos=[5, 3.5, 3.5, 4])

        m.h("4.9 Fuerza dinámica, amplificación y amplitud de vibración", 2)
        m.p(f"La fuerza desbalanceada se estima con el grado de balanceo del rotor (ISO 1940-1). {R['F0_desc']}, "
            f"con un factor de servicio S_{{f}} = {f(R['Sf'], 1)} (ACI 351.3R 3.2).", align="j")
        if R["clase"] != "electromagnetica":
            m.eq(f"m_{{r}} = {f(R['pct_rotor'], 0)} % × {f(R['W_kg'], 0)} = {f(R['m_r'], 1)} kg ;   "
                 f"G = {f(R['G_bal'], 1)} mm/s ;   e = G/ω = {f(R['e_bal'], 2)} μm")
            m.eq(f"F_{{0}} = {f(R['Sf'], 1)} × {f(R['m_r'], 1)} × {f(R['G_bal'] / 1000, 4)} × {f(R['omega'])} = "
                 f"{f(R['F0_N'], 1)} N = {f(R['F0'] * 1000, 2)} kg")
        else:
            m.eq(f"F_{{0}} = {f(R['pct_rotor'], 1)} % × {f(R['W_kg'], 0)} kg × {f(R['Sf'], 1)} = "
                 f"{f(R['F0'] * 1000, 2)} kg")
        m.eq("M = 1/√[(1 − r^{2})^{2} + (2·ζ·r)^{2}]")
        m.eq(f"M = 1/√[(1 − {f(R['r'], 3)}²)² + (2 × {f(R['zeta_z'], 3)} × {f(R['r'], 3)})²] = {f(R['M'], 4)}")
        m.eq("A_{máx} = (F_{0}/K_{z}) · M")
        m.eq(f"A_{{máx}} = ({f(R['F0'], 5)}/{f(R['Kz'], 0)}) × {f(R['M'], 4)} = {f(R['A_mm'], 5)} mm")
        m.eq(f"v_{{máx}} = A·ω = {f(R['vel'], 4)} mm/s")
        m.check(f"A_{{máx}} = {f(R['A_mm'], 5)} mm < A_{{adm}} = {f(R['A_lim'], 3)} mm (Richart/Novak)",
                R["ok_amp"])
        m.p("Curva del factor de amplificación dinámica del sistema (para el amortiguamiento de diseño):",
            align="j")
        curva = curva_amplificacion(R["zeta_z"], 3.0, 16)
        m.tabla([[f(rr, 2), f(MM, 4), "← operación" if abs(rr - R["r"]) < 0.1 else ""]
                 for rr, MM in curva], ["r = f/f_{n}", "M", "Observación"], anchos=[4, 4, 4], size=8.5)

        m.h("4.10 Aislamiento de vibraciones (amortiguadores)", 2)
        if R["con_ais"]:
            m.eq("f_{i} = (1/2π)·√(g/δ_{st}) = 15.76/√δ_{st}(mm)")
            m.eq(f"f_{{i}} = 15.76/√{f(R['defl'], 1)} = {f(R['f_iso'], 3)} Hz ;   r_{{i}} = {f(R['fop'])}/"
                 f"{f(R['f_iso'], 3)} = {f(R['ri'], 3)}")
            m.eq("T = 1/|r_{i}^{2} − 1|")
            m.eq(f"T = {f(R['T'], 4)}  →  eficiencia de aislamiento = {f(R['efic'], 1)} %")
            m.p(f"Fuerza transmitida al bloque: F_{{t}} = T·F_{{0}} = {f(R['F_trans'] * 1000, 3)} kg; "
                f"amplitud con aisladores: {f(R['A_iso_mm'], 6)} mm (la verificación del ítem 4.9 se "
                f"realizó conservadoramente con montaje rígido).", align="j")
            m.check("Relación r_{i} > 1.5·√2 (zona de aislamiento eficaz)", R["ok_iso"])
        else:
            m.p("El equipo se ancla rígidamente al bloque (T = 1.0); toda la fuerza dinámica es transmitida.")

        m.h("4.11 Acero de refuerzo del bloque (NTE E.060 / ACI 351.3R)", 2)
        m.p("El bloque se refuerza con doble parrilla (superior e inferior) y acero perimetral lateral para "
            "controlar la fisuración por contracción y temperatura y para absorber esfuerzos de flexión "
            "inducidos por la reacción del suelo en los volados.", align="j")
        m.eq(f"A_{{s,mín}} = 0.0018·b·H = 0.0018 × 100 × {f(R['H'] * 100, 0)} = {f(R['As_tot'], 2)} cm²/m")
        m.eq(f"A_{{s,cara}} = {f(R['As_tot'], 2)}/2 = {f(R['As_cara'], 2)} cm²/m")
        m.eq(f"M_{{u}} = 1.5·q_{{máx}}·v²/2 = 1.5 × {f(R['q_max'], 3)} × {f(R['v_vol'], 3)}²/2 = "
             f"{f(R['Mu'], 3)} t·m/m")
        m.eq(f"A_{{s,flex}} = M_{{u}}/(φ·f_{{y}}·0.9d) = {f(R['Mu'] * 1e5, 0)}/(0.9 × 4200 × 0.9 × "
             f"{f(R['d'] * 100, 1)}) = {f(R['As_flex'], 3)} cm²/m")
        m.eq(f"s = A_{{b}}/A_{{s}} = {f(R['ab'])}/{f(max(R['As_cara'], R['As_flex']), 2)} → Ø{R['barra']} @ "
             f"{f(R['s'])} m  (A_{{s,prov}} = {f(R['As_prov'], 2)} cm²/m)")
        m.check(f"Doble parrilla Ø{R['barra']} @ {f(R['s'])} m + acero lateral", R["ok_acero"])

        m.h("4.12 Pernos de anclaje (ACI 318-19, Capítulo 17)", 2)
        m.tabla([["Cantidad / diámetro", f"{R['n']} pernos Ø{R['dp']} (d_{{a}} = {f(R['da'], 3)} cm)"],
                 ["Material", f"{R['grado']}: f_{{ya}} = {f(R['fya'], 0)}, f_{{uta}} = {f(R['futa'], 0)} kg/cm²"],
                 ["Área efectiva A_{se,N}", f"{f(R['Ase'], 3)} cm²"],
                 ["Empotramiento h_{ef}", f"{f(R['hef'] * 100, 1)} cm (gancho J, e_{{h}} = 4.5d_{{a}})"],
                 ["Distancia al borde c_{a,min}", f"{f(R['ca'] * 100, 1)} cm"],
                 ["Coef. sísmico horizontal C_{h}", f"{f(R['Ch'])} (E.030 Cap. 6)"]],
                ["Dato", "Valor"], anchos=[6, 10])
        m.p("Solicitaciones:", bold=True)
        m.eq(f"V_{{u}} = C_{{h}}·W_{{eq}} + F_{{0}} = {f(R['Ch'])} × {f(R['Weq'], 3)} + {f(R['F0'], 4)} = "
             f"{f(R['Vu'], 3)} t ;   V_{{u,perno}} = {f(R['Vu_b'], 1)} kg")
        m.eq(f"M_{{vu}} = V_{{u}}·(h_{{cg}} + 0.35) = {f(R['Mvu'], 3)} t·m ;   brazo = {f(R['brazo'], 3)} m")
        m.eq(f"N_{{u,perno}} = [M_{{vu}}/brazo − 0.9·W_{{eq}}/2]/(n/2) = {f(R['Nu_b'], 1)} kg")
        m.p("Resistencias de diseño a tracción:", bold=True)
        m.eq(f"φN_{{sa}} = 0.75·A_{{se}}·f_{{uta}} = 0.75 × {f(R['Ase'], 3)} × {f(R['futa'], 0)} = "
             f"{f(R['phiNsa'], 0)} kg")
        m.eq(f"N_{{b}} = k_{{c}}·√f'c·h_{{ef}}^{{1.5}} = 10 × √{f(R['fc_MPa'], 2)} × {f(R['hef'] * 1000, 0)}^{{1.5}} = "
             f"{f(R['Nb_N'] / 1000, 2)} kN")
        m.eq(f"N_{{cb}} = (A_{{Nc}}/A_{{Nco}})·ψ_{{ed,N}}·ψ_{{c,N}}·N_{{b}} = {f(min(R['ANc'] / R['ANco'], 1), 3)} × "
             f"{f(R['psi_ed'], 3)} × {f(R['psi_cN'], 2)} × N_{{b}} = {f(R['Ncb'], 0)} kg ;   φN_{{cb}} = "
             f"{f(R['phiNcb'], 0)} kg")
        m.eq(f"N_{{p}} = 0.9·f'c·e_{{h}}·d_{{a}} = 0.9 × {f(R['fc'], 0)} × {f(R['eh'], 2)} × {f(R['da'], 3)} = "
             f"{f(R['Np'], 0)} kg ;   φψ_{{c,P}}N_{{p}} = {f(R['phiNpn'], 0)} kg")
        m.eq(f"φN_{{n}} = {f(R['phiNn'], 0)} kg  (gobierna: {R['modo_N']})")
        m.p("Resistencias de diseño a corte:", bold=True)
        m.eq(f"φV_{{sa}} = 0.65 × 0.6·A_{{se}}·f_{{uta}} × 0.80 (grout) = {f(R['phiVsa'], 0)} kg")
        m.eq(f"V_{{b}} = 0.6·(l_{{e}}/d_{{a}})^{{0.2}}·√d_{{a}}·√f'c·c_{{a1}}^{{1.5}} → φV_{{cb}} = {f(R['phiVcb'], 0)} kg")
        m.eq(f"φV_{{cp}} = 0.70·k_{{cp}}·N_{{cb}} = 0.70 × {f(R['kcp'], 0)} × {f(R['Ncb'], 0)} = {f(R['phiVcp'], 0)} kg")
        m.eq(f"φV_{{n}} = {f(R['phiVn'], 0)} kg  (gobierna: {R['modo_V']})")
        m.p("Interacción tracción – corte (ACI 318-19 17.8):", bold=True)
        m.eq(f"N_{{ua}}/φN_{{n}} + V_{{ua}}/φV_{{n}} = {f(R['rN'], 3)} + {f(R['rV'], 3)} = "
             f"{f(R['rN'] + R['rV'], 3)} ≤ 1.20")
        m.check("Interacción de pernos", R["ok_int"])
        m.eq(f"h_{{ef}} = {f(R['hef'] * 100, 1)} cm ≥ 12·d_{{a}} = {f(R['hef_min'] * 100, 1)} cm  y  ≤ H − 10 cm "
             f"= {f((R['H'] - 0.10) * 100, 1)} cm")
        m.check("Longitud de empotramiento", R["ok_hef"])
        m.eq(f"Pretensado F_{{p}} = 0.5·f_{{ya}}·A_{{se}} = {f(R['Fp'], 0)} kg ;   Torque T = K·d·F_{{p}} = "
             f"0.20 × {f(R['da'] / 100, 4)} × {f(R['Fp'], 0)} = {f(R['torque'], 1)} kg·m")

        m.h("4.13 Mortero de nivelación (Grout)", 2)
        m.p(f"Se especifica grout cementicio no contráctil ASTM C1107 (o grout epóxico en ambientes con "
            f"vibración severa) de espesor e = {f(R['e_grout'], 1)} cm, dentro del rango recomendado de "
            f"2.5 cm a 5.0 cm, con resistencia mínima a la compresión de 350 kg/cm² a 28 días.", align="j")
        m.check("Espesor de grout 2.5 cm ≤ e ≤ 5.0 cm", R["ok_grout"])

        m.h("4.14 Resumen de verificaciones – bancada", 2)
        m.tabla([
            ["Relación de masas W_{c}/W_{eq}", f(R["ratio"]), f"≥ {f(R['ratio_min'], 1)}", ok_txt(R["ok_masa"])],
            ["Borde libre mínimo (m)", f(min(R["borde_L"], R["borde_B"]), 3), "≥ 0.25", ok_txt(R["ok_borde"])],
            ["Espesor H (m)", f(R["H"]), f"≥ {f(R['H_min'], 3)}", ok_txt(R["ok_H"])],
            ["Excentricidad e (m)", f(R["e"], 4), f"≤ {f(R['e_aci'], 3)}", ok_txt(R["ok_e"])],
            ["Presión q_{máx} (kg/cm²)", f(R["q_max"] / 10, 3), f"≤ {f(R['q_lim'] / 10, 3)}", ok_txt(R["ok_q"])],
            ["Frecuencia natural f_{n} (Hz)", f(R["fn"]), "—", "—"],
            ["Relación r = f_{op}/f_{n}", f(R["r"], 3), "∉ [0.70 ; 1.40]",
             "CUMPLE" if R["ok_res"] else "EN RESONANCIA"],
            ["Amplitud A_{máx} (mm)", f(R["A_mm"], 5), f"< {f(R['A_lim'], 3)}", ok_txt(R["ok_amp"])],
            [f"Acero Ø{R['barra']} @ {f(R['s'])} (cm²/m)", f(R["As_prov"], 2),
             f"≥ {f(max(R['As_cara'], R['As_flex']), 2)}", ok_txt(R["ok_acero"])],
            ["Interacción pernos", f(R["inter"], 3), "≤ 1.20", ok_txt(R["ok_int"])],
            ["Empotramiento h_{ef} (cm)", f(R["hef"] * 100, 1), f"≥ {f(R['hef_min'] * 100, 1)}", ok_txt(R["ok_hef"])],
            ["Grout e (cm)", f(R["e_grout"], 1), "2.5 – 5.0", ok_txt(R["ok_grout"])],
        ], ["Verificación", "Valor", "Límite", "Estado"], anchos=[6.5, 3, 3, 3.5])
        m.salto()

    # ---------------- 5. MÓDULO B ----------------
    if RP.get("ok"):
        R = RP
        m.h("5. DISEÑO DEL MACIZO PARA POSTE / TORRE (MÉTODO DE SULZBERGER)", 1)
        m.h("5.1 Datos de la estructura", 2)
        m.tabla([
            ["Tipo", R["tipo"]], ["Altura total H_{p} / libre H_{libre}", f"{f(R['Hp'])} m / {f(R['Hl'])} m"],
            ["Diámetro en punta / base", f"{f(R['d_top'], 3)} m / {f(R['d_base'], 3)} m"],
            ["Peso del poste / equipos", f"{f(R['Wp'], 0)} kg / {f(R['Weq'], 0)} kg"],
            ["Área de luminarias/equipos en punta", f"{f(R['Aeq'])} m² (C = {f(R['Ceq'])})"],
            ["Tiro en la punta", f"{f(R['tiro'], 0)} kg"],
            ["Velocidad de viento de zona V", f"{f(R['V'], 0)} km/h (Mapa eólico E.020)"],
            ["Macizo a × b × t", f"{f(R['a'])} × {f(R['b'])} × {f(R['t'])} m"],
            ["Pedestal a_{ped} × a_{ped} × h_{ped}", f"{f(R['a_ped'])} × {f(R['a_ped'])} × {f(R['h_ped'])} m"],
        ], ["Dato", "Valor"], anchos=[6.5, 9.5])

        m.h("5.2 Presión de viento (NTE E.020)", 2)
        m.p("La velocidad de diseño varía con la altura sobre el terreno y la presión dinámica depende del "
            "factor de forma C (E.020, Tabla 4). El fuste troncocónico se discretiza en tramos de 0.50 m.",
            align="j")
        m.eq("V_{h} = V·(h/10)^{0.22}      P_{h} = 0.005·C·V_{h}^{2}  (kg/m²)")
        m.eq(f"En la punta: V_{{h}} = {f(R['V'], 0)} × ({f(R['zg_top'])}/10)^{{0.22}} = {f(R['Vh_top'])} km/h ;  "
             f"P_{{h}} = 0.005 × {f(R['Ceq'])} × {f(R['Vh_top'])}² = {f(R['P_eq'])} kg/m²")
        filas = [[str(tr["i"]), f(tr["zg"]), f(tr["d"], 3), f(tr["Vh"]), f(tr["Ph"]), f(tr["F"]), f(tr["M"])]
                 for tr in R["tramos"]]
        m.tabla(filas, ["Tramo", "h (m)", "d (m)", "V_{h} (km/h)", "P_{h} (kg/m²)", "F (kg)", "M (kg·m)"],
                anchos=[1.5, 2, 2, 2.5, 2.5, 2.5, 2.8], size=8)

        m.h("5.3 Fuerzas y momento de volcamiento", 2)
        m.eq(f"F_{{fuste}} = ΣP_{{h}}·d·Δh = {f(R['F_f'])} kg ;   M_{{fuste}} = {f(R['M_f'])} kg·m")
        m.eq(f"F_{{equipos}} = P_{{h}}·A = {f(R['P_eq'])} × {f(R['Aeq'])} = {f(R['F_eq'])} kg ;   "
             f"M = {f(R['M_eq'])} kg·m")
        m.eq(f"Tiro: T = {f(R['tiro'], 0)} kg a {f(R['h_tiro'])} m → M = {f(R['M_tiro'])} kg·m ;  "
             f"F_{{pedestal}} = {f(R['F_ped'])} kg")
        m.eq(f"H_{{t}} = {f(R['Ht'], 4)} t ;   M_{{base}} = {f(R['M_base'], 3)} t·m ;   "
             f"M_{{terreno}} = {f(R['M_suelo'], 3)} t·m")
        m.p("Según Sulzberger, el macizo gira alrededor de un eje situado a 2t/3 bajo la superficie:", align="j")
        m.eq("M_{v} = M_{terreno} + H_{t}·(2t/3)")
        m.eq(f"M_{{v}} = {f(R['M_suelo'], 3)} + {f(R['Ht'], 4)} × (2 × {f(R['t'])}/3) = {f(R['Mv'], 3)} t·m")

        m.h("5.4 Cargas verticales", 2)
        m.eq(f"P = W_{{poste}} + W_{{eq}} + W_{{macizo}} + W_{{pedestal}} = {f((R['Wp'] + R['Weq']) / 1000, 3)} + "
             f"{f(R['W_maz'], 3)} + {f(R['W_ped'], 3)} = {f(R['Pv'], 3)} t")

        m.h("5.5 Momentos resistentes de Sulzberger", 2)
        m.p("El índice de compresibilidad del terreno crece linealmente con la profundidad. Con C_{2} "
            "referido a 2 m de profundidad:", align="j")
        m.eq(f"C_{{t}} = C_{{2}}·t/2 = {f(R['C2'])} × {f(R['t'])}/2 = {f(R['Ct'] / 1000, 3)} kg/cm³ = "
             f"{f(R['Ct'], 0)} t/m³")
        m.eq(f"C_{{b}} = {f(R['Cb_fac'])}·C_{{t}} = {f(R['Cb'], 0)} t/m³")
        m.p(f"Giro máximo admisible: tan α = {f(R['tan_adm'], 3)}.", align="j")
        m.eq("M_{1} = (b·t^{3}/36)·C_{t}·tan α")
        m.eq(f"M_{{1}} = ({f(R['b'])} × {f(R['t'])}³/36) × {f(R['Ct'], 0)} × {f(R['tan_adm'], 3)} = "
             f"{f(R['M1'], 3)} t·m")
        m.eq("M_{2} = P·a·[0.5 − (2/3)·√(P/(2·b·a^{2}·C_{b}·tan α))]")
        if not R["contacto_total"]:
            m.eq(f"M_{{2}} = {f(R['Pv'], 3)} × {f(R['a'])} × [0.5 − (2/3)·√({f(R['Pv'], 3)}/(2 × {f(R['b'])} × "
                 f"{f(R['a'])}² × {f(R['Cb'], 0)} × {f(R['tan_adm'], 3)}))] = {f(R['M2'], 3)} t·m")
        else:
            m.p("Como el ancho de contacto a_{0} resulta mayor que a, el fondo trabaja en contacto total y "
                "se usa M_{2} = b·a^{3}·C_{b}·tan α/12:", align="j")
            m.eq(f"M_{{2}} = {f(R['M2'], 3)} t·m")
        m.eq(f"Ancho de contacto en el fondo a_{{0}} = √(2P/(b·C_{{b}}·tan α)) = {f(R['a0'], 3)} m")
        m.eq("FS = (M_{1} + M_{2})/M_{v} ≥ 1.50")
        m.eq(f"FS = ({f(R['M1'], 3)} + {f(R['M2'], 3)})/{f(R['Mv'], 3)} = {f(R['FS'], 3)}")
        m.check(f"Factor de seguridad de Sulzberger FS = {f(R['FS'], 2)} ≥ 1.50", R["ok_FS"])
        m.p(f"Comprobación complementaria con el coeficiente s de Sulzberger: M_{{1}}/M_{{2}} = {f(R['rel'], 3)} "
            f"→ s = {f(R['s_sulz'], 3)};  M_{{1}} + M_{{2}} = {f(R['M1'] + R['M2'], 3)} ≥ s·M_{{v}} = "
            f"{f(R['s_sulz'] * R['Mv'], 3)} t·m → {ok_txt(R['ok_s'])}.", align="j")
        m.h("5.6 Giro real y presiones sobre el terreno", 2)
        m.p("Se obtiene por bisección el giro tan α para el cual M_{1}(α) + M_{2}(α) = M_{v}:", align="j")
        m.eq(f"tan α_{{req}} = {f(R['tan_req'], 5)}  ≤  tan α_{{adm}} = {f(R['tan_adm'], 3)}")
        m.check("Giro del macizo", R["ok_tan"])
        m.eq(f"σ_{{lat,máx}} = C_{{t}}·t·tan α/9 = {f(R['sig_lat'], 3)} t/m² = {f(R['sig_lat'] / 10, 3)} kg/cm²  "
             f"(zona superior, a la profundidad t/3)")
        m.eq(f"σ_{{lat,inf}} = C_{{t}}·t·tan α/3 = {f(R['sig_lat_inf'], 3)} t/m² = {f(R['sig_lat_inf'] / 10, 3)} kg/cm²  "
             f"(zona inferior, en el fondo)")
        m.eq(f"σ_{{fondo,máx}} = C_{{b}}·tan α·a_{{0}} = {f(R['sig_fondo'], 3)} t/m² ;   "
             f"σ_{{medio}} = P/(a·b) = {f(R['sig_med'], 3)} t/m²")
        m.check(f"σ_{{medio}} = {f(R['sig_med'] / 10, 3)} kg/cm² ≤ q_{{adm}} = {f(R['qadm'])} kg/cm²", R["ok_q"])

        m.h("5.7 Pedestal, armadura y anclaje del poste", 2)
        m.eq(f"M_{{u}} = 1.25·M_{{terreno}} = {f(R['Mu_ped'], 3)} t·m ;   d = {f(R['d_ped'] * 100, 1)} cm")
        m.eq(f"A_{{s}} = M_{{u}}/(φ·f_{{y}}·0.9d) = {f(R['As_ped'], 2)} cm² por cara ;  A_{{s,mín}} = "
             f"0.005·A_{{g}} = {f(R['As_min_ped'], 2)} cm²")
        m.eq(f"Se colocan {R['n_ped']} Ø{R['barra']} ({R['n_cara']} por cara) + estribos Ø3/8\" @ 0.15 m")
        m.check("Acero del pedestal", R["ok_ped"])
        m.eq(f"Malla inferior y superior del macizo: Ø{R['barra']} @ {f(R['s_maz'])} m en ambas direcciones")
        if R["n_pp"] > 0 and not R["empotrado"]:
            m.eq(f"T_{{perno}} = 1.25·4·M_{{base}}/(n·D_{{c}}) − 0.9·W/n = {f(R['Tb'], 1)} kg ;   "
                 f"V_{{perno}} = {f(R['Vb_p'], 1)} kg")
            m.eq(f"φN_{{sa}} = {f(R['phiN_p'], 0)} kg ;  φV_{{sa}} = {f(R['phiV_p'], 0)} kg ;  "
                 f"interacción = {f(R['inter_p'], 3)} ≤ 1.20")
            m.check(f"Canastilla {R['n_pp']} pernos Ø{R['d_pp']} sobre círculo D_{{c}} = {f(R['Dc'])} m",
                    R["ok_pp"])
            m.check(f"Longitud de anclaje L = {f(R['L_anc'])} m ≤ h_{{ped}} + t − 0.15", R["ok_Lanc"])
        else:
            m.eq(f"Empotramiento del poste e = H_{{p}} − H_{{libre}} = {f(R['e_emp'])} m ≥ H_{{p}}/10 = "
                 f"{f(R['e_min'])} m")
            m.check("Empotramiento del poste en el macizo", R["ok_emp"])

        m.h("5.8 Resumen de verificaciones – macizo", 2)
        m.tabla([
            ["Momento de vuelco M_{v} (t·m)", f(R["Mv"], 3), "—", "—"],
            ["Momento lateral M_{1} (t·m)", f(R["M1"], 3), "—", "—"],
            ["Momento de fondo M_{2} (t·m)", f(R["M2"], 3), "—", "—"],
            ["FS = (M_{1}+M_{2})/M_{v}", f(R["FS"], 3), "≥ 1.50", ok_txt(R["ok_FS"])],
            ["Giro tan α", f(R["tan_req"], 5), f"≤ {f(R['tan_adm'], 3)}", ok_txt(R["ok_tan"])],
            ["σ medio en el fondo (kg/cm²)", f(R["sig_med"] / 10, 3), f"≤ {f(R['qadm'])}", ok_txt(R["ok_q"])],
            ["Acero pedestal", f"{R['n_ped']} Ø{R['barra']}", f"A_{{s}} ≥ {f(R['As_ped'], 2)}", ok_txt(R["ok_ped"])],
            ["Pernos / empotramiento", f(R["inter_p"], 3) if R["n_pp"] else f(R["e_emp"]), "—",
             ok_txt(R["ok_pp"] and R["ok_emp"] and R["ok_Lanc"])],
        ], ["Verificación", "Valor", "Límite", "Estado"], anchos=[6.5, 3, 3, 3.5])
        m.salto()

    # ---------------- 6. METRADOS ----------------
    m.h("6. METRADOS Y PRESUPUESTO REFERENCIAL", 1)
    filas, total = [], 0.0
    for it in items:
        if it[2] == "":
            filas.append([it[0], it[1], "", "", "", ""])
        else:
            parc = it[3] * it[4]
            total += parc
            filas.append([it[0], it[1], it[2], f(it[3], 2), f(it[4], 2), f(parc, 2)])
    filas.append(["", "COSTO DIRECTO (S/.)", "", "", "", f(total, 2)])
    m.tabla(filas, ["Ítem", "Descripción", "Und", "Metrado", "P.U. (S/.)", "Parcial (S/.)"],
            anchos=[1.3, 7.2, 1.2, 2.0, 2.0, 2.5], size=8.5)
    m.p("Los precios unitarios son referenciales y deben actualizarse según el mercado local "
        "(la hoja Excel exportada permite editarlos).", italic=True, size=9)

    # ---------------- 7. CONCLUSIONES ----------------
    m.h("7. CONCLUSIONES Y RECOMENDACIONES CONSTRUCTIVAS", 1)
    m.h("7.1 Conclusiones", 2)
    if RE.get("ok"):
        m.bullet(f"El bloque de {f(RE['L'])} × {f(RE['B'])} × {f(RE['H'])} m pesa {f(RE['Wc'], 2)} t, "
                 f"{f(RE['ratio'])} veces el peso del equipo ({ok_txt(RE['ok_masa'])} el criterio de masa).")
        m.bullet(f"La frecuencia natural vertical es {f(RE['fn'])} Hz frente a una frecuencia de operación de "
                 f"{f(RE['fop'])} Hz (r = {f(RE['r'], 3)}): el sistema se encuentra "
                 f"{'FUERA de' if RE['ok_res'] else 'DENTRO de'} la zona de resonancia.")
        m.bullet(f"La amplitud máxima de vibración es {f(RE['A_mm'], 5)} mm, "
                 f"{'menor' if RE['ok_amp'] else 'mayor'} que el límite de {f(RE['A_lim'], 3)} mm.")
        m.bullet(f"La presión máxima sobre el suelo es {f(RE['q_max'] / 10, 3)} kg/cm² "
                 f"({f(RE['q_max'] / 10 / RE['qadm'] * 100, 1)} % de q_{{adm}}).")
        m.bullet(f"Refuerzo: doble parrilla Ø{RE['barra']} @ {f(RE['s'])} m y {RE['n']} pernos tipo J "
                 f"Ø{RE['dp']} {RE['grado']} con h_{{ef}} = {f(RE['hef'] * 100, 0)} cm.")
    if RP.get("ok"):
        m.bullet(f"El macizo de {f(RP['a'])} × {f(RP['b'])} × {f(RP['t'])} m resiste el momento de vuelco "
                 f"M_{{v}} = {f(RP['Mv'], 3)} t·m con FS = {f(RP['FS'], 2)} ({ok_txt(RP['ok_FS'])}) y un giro "
                 f"tan α = {f(RP['tan_req'], 5)} ({ok_txt(RP['ok_tan'])}).")
    m.h("7.2 Recomendaciones constructivas", 2)
    for t in [
        "Usar grout cementicio no contráctil (ASTM C1107) o epóxico de espesor 2.5 – 5.0 cm bajo las placas "
        "base. Humedecer el concreto 24 h antes (grout cementicio) o mantenerlo seco (grout epóxico).",
        "Picar y limpiar la superficie superior del bloque (rugosidad ≥ 6 mm) antes del vaciado del grout.",
        "Colocar la junta perimetral de poliestireno elastificado o neopreno de 1\" en toda la altura de "
        "contacto con el piso, y sellar superiormente con sellador elastomérico de poliuretano; no debe "
        "existir ningún puente rígido (mortero, cerámica, contrapiso) entre el bloque y el piso.",
        "Vaciar el bloque de forma monolítica, sin juntas frías; curar con agua durante al menos 7 días.",
        "Posicionar los pernos con plantilla metálica rígida; tolerancia de ± 3 mm en planta y ± 5 mm en nivel. "
        "Proteger las roscas con cinta y grasa durante el vaciado.",
        f"Aplicar el torque de ajuste recomendado ({f(RE.get('torque', 0), 1)} kg·m para el perno del equipo) "
        "con torquímetro calibrado, en secuencia cruzada, y colocar contratuerca y arandela de presión.",
        "Re-torquear los pernos luego de 30 días de operación y verificar el nivel del equipo.",
        "Realizar una medición de vibraciones (velocidad RMS) en la puesta en marcha y comparar con ISO 10816.",
        "Los aisladores (resortes/elastómero) deben ser suministrados por el fabricante del equipo y "
        "seleccionados para la deflexión estática indicada.",
        "En el macizo del poste: compactar el relleno lateral en capas de 0.20 m al 95 % del Proctor "
        "Modificado, pues la resistencia lateral de Sulzberger depende del confinamiento.",
        "El pedestal debe tener chaflán y pendiente superior para escurrimiento del agua de lluvia; "
        "dejar el tubo PVC-SAP para el conductor de puesta a tierra antes del vaciado.",
        "Verificar en obra que el terreno corresponde al indicado en el EMS; si se encuentra relleno no "
        "controlado, profundizar la cimentación hasta el terreno natural.",
    ]:
        m.bullet(t)

    # ---------------- 8. ESPECIFICACIONES ----------------
    m.h("8. ESPECIFICACIONES TÉCNICAS", 1)
    m.tabla([
        ["Concreto", f"f'c = {int(fc)} kg/cm², cemento tipo I (tipo V si hay sulfatos), a/c ≤ 0.50"],
        ["Solado", "Concreto C:H 1:12, e = 0.10 m"],
        ["Acero", "ASTM A615 Gr. 60, f_{y} = 4200 kg/cm²"],
        ["Recubrimientos", "7.5 cm (contra el terreno), 5.0 cm (expuesto)"],
        ["Pernos de anclaje", f"{RE.get('grado', 'ASTM F1554 Gr. 36')}, galvanizados en caliente ASTM A153"],
        ["Tuercas / arandelas", "ASTM A563 pesadas / ASTM F436 planas + arandela de presión"],
        ["Grout", "No contráctil ASTM C1107, f'c ≥ 350 kg/cm² a 28 días"],
        ["Junta perimetral", "Poliestireno elastificado o neopreno e = 1\" (2.5 cm)"],
        ["Sello de junta", "Sellador elastomérico de poliuretano, ancho 2.5 cm, prof. 1.5 cm"],
        ["Relleno", "Material propio seleccionado, compactado al 95 % del Proctor Modificado"],
    ], ["Material", "Especificación"], anchos=[4, 12])
    m.h("8.1 Procedimiento constructivo de la bancada", 2)
    for i, t in enumerate([
        "Trazo y replanteo de ejes del equipo según los planos del fabricante.",
        "Excavación hasta el nivel de fondo del solado; perfilado y compactación del fondo.",
        "Vaciado del solado e = 0.10 m.",
        "Armado de la parrilla inferior, acero lateral y parrilla superior con dados de concreto.",
        "Colocación de pernos con plantilla, camisas de PVC y tubo conduit de salida de cables.",
        "Encofrado, colocación de la junta perimetral y vaciado monolítico del bloque con vibrado.",
        "Curado durante 7 días; montaje de aisladores y equipo; nivelación con lainas.",
        "Vaciado del grout, curado y torque final de pernos; sellado de la junta.",
    ], 1):
        m.p(f"{i}. {t}", align="j")
    m.h("8.2 Procedimiento constructivo del macizo", 2)
    for i, t in enumerate([
        "Excavación del macizo con las dimensiones de diseño, sin sobre-excavar las paredes laterales.",
        "Vaciado de solado, armado de malla inferior, canastilla de pernos o poste aplomado.",
        "Vaciado del macizo y del pedestal con chaflán; colocación del tubo PVC de puesta a tierra.",
        "Relleno compactado lateral por capas; izaje y aplomado del poste; grout y torque de pernos.",
    ], 1):
        m.p(f"{i}. {t}", align="j")
    # ---------------- 9. ANEXOS ----------------
    m.salto()
    m.h("9. ANEXOS", 1)
    m.h("Anexo A – Parámetros dinámicos referenciales del suelo", 2)
    m.tabla([["Arcilla blanda / limo suelto", "80 – 150", "120 – 400", "0.40 – 0.45"],
             ["Arcilla media / arena suelta", "150 – 250", "400 – 1,150", "0.35 – 0.40"],
             ["Arena densa / arcilla dura", "250 – 400", "1,150 – 3,000", "0.30 – 0.35"],
             ["Grava arenosa (conglomerado)", "400 – 700", "3,000 – 9,500", "0.25 – 0.33"],
             ["Roca meteorizada", "700 – 1,200", "9,500 – 30,000", "0.20 – 0.30"]],
            ["Tipo de suelo", "V_{s} (m/s)", "G (kg/cm²)", "ν"], anchos=[6, 3, 3.5, 3])
    m.p("Valores orientativos; el diseño debe basarse en el Estudio de Mecánica de Suelos (ensayos "
        "MASW, cross-hole, down-hole o correlaciones con N_{SPT}).", italic=True, size=9)
    m.h("Anexo B – Índice de compresibilidad C para Sulzberger (a 2 m de profundidad)", 2)
    m.tabla([["Terreno pantanoso, turba", "0.5 – 1.0"], ["Arcilla blanda, arena fina suelta", "1.0 – 2.0"],
             ["Arcilla semidura, arena media", "3.0 – 5.0"], ["Arcilla dura, arena gruesa compacta", "6.0 – 8.0"],
             ["Grava arenosa compacta", "8.0 – 12.0"], ["Conglomerado / roca fisurada", "12.0 – 20.0"]],
            ["Tipo de terreno", "C (kg/cm³)"], anchos=[10, 5])
    m.h("Anexo C – Coeficiente s de Sulzberger", 2)
    m.tabla([[f(x, 1), f(y, 3)] for x, y in TABLA_SULZBERGER], ["M_{1}/M_{2}", "s"], anchos=[5, 5])
    m.h("Anexo D – Criterios de amplitud y severidad de vibración", 2)
    m.tabla([["< 0.0025 mm (a 30 Hz)", "No perceptible por las personas"],
             ["0.0025 – 0.010 mm", "Apenas perceptible"],
             ["0.010 – 0.025 mm", "Perceptible – molesto para las personas"],
             ["0.025 – 0.050 mm", "Límite usual para máquinas (Richart / Novak)"],
             ["> 0.050 mm", "Problemas para máquinas – riesgo de daño"]],
            ["Amplitud de desplazamiento", "Clasificación (carta de Richart, 1962)"], anchos=[6, 9])
    m.tabla([["≤ 2.8 mm/s", "Zona A/B – Buena (máquinas grupo 2, ISO 10816-3)"],
             ["2.8 – 4.5 mm/s", "Zona B/C – Aceptable para operación continua"],
             ["4.5 – 7.1 mm/s", "Zona C – Operación limitada"],
             ["> 7.1 mm/s", "Zona D – Daño probable"]],
            ["Velocidad RMS", "Severidad (ISO 10816-3)"], anchos=[6, 9])
    if RE.get("ok"):
        m.h("Anexo E – Datos de entrada de la bancada", 2)
        claves = [("Tipo", RE["tipo"]), ("Potencia", P.get("_potencia", "")), ("W_{eq} (kg)", f(RE["W_kg"], 0)),
                  ("RPM", f(RE["rpm"], 0)), ("L_{eq} × B_{eq} (m)", f"{f(RE['Leq'])} × {f(RE['Beq'])}"),
                  ("L × B × H (m)", f"{f(RE['L'])} × {f(RE['B'])} × {f(RE['H'])}"),
                  ("V_{s} (m/s) / ν", f"{f(RE['Vs'], 0)} / {f(RE['nu'])}"),
                  ("Masa rotante (%) / G (mm/s) / S_{f}", f"{f(RE['pct_rotor'], 0)} / {f(RE['G_bal'], 1)} / {f(RE['Sf'], 1)}"),
                  ("Aislamiento", RE["aislamiento"]), ("Deflexión del aislador (mm)", f(RE["defl"], 0)),
                  ("Pernos", f"{RE['n']} Ø{RE['dp']} {RE['grado']} h_{{ef}} = {f(RE['hef'])} m"),
                  ("Grout (cm) / C_{h}", f"{f(RE['e_grout'], 1)} / {f(RE['Ch'])}")]
        m.tabla([[k_, v_] for k_, v_ in claves], ["Dato", "Valor"], anchos=[6, 10])
    if RP.get("ok"):
        m.h("Anexo F – Datos de entrada del macizo", 2)
        claves = [("Tipo", RP["tipo"]), ("H_{p} / H_{libre} (m)", f"{f(RP['Hp'])} / {f(RP['Hl'])}"),
                  ("d punta / base (m)", f"{f(RP['d_top'], 3)} / {f(RP['d_base'], 3)}"),
                  ("V (km/h) / C fuste / C equipos", f"{f(RP['V'], 0)} / {f(RP['Cf'])} / {f(RP['Ceq'])}"),
                  ("Área equipos (m²) / tiro (kg)", f"{f(RP['Aeq'])} / {f(RP['tiro'], 0)}"),
                  ("a × b × t (m)", f"{f(RP['a'])} × {f(RP['b'])} × {f(RP['t'])}"),
                  ("Pedestal (m)", f"{f(RP['a_ped'])} × {f(RP['a_ped'])} × {f(RP['h_ped'])}"),
                  ("C_{2} (kg/cm³) / C_{b}/C_{t} / tan α", f"{f(RP['C2'])} / {f(RP['Cb_fac'])} / {f(RP['tan_adm'], 3)}"),
                  ("Pernos de canastilla", f"{RP['n_pp']} Ø{RP['d_pp']} D_{{c}} = {f(RP['Dc'])} m" if RP["n_pp"] else "Poste empotrado")]
        m.tabla([[k_, v_] for k_, v_ in claves], ["Dato", "Valor"], anchos=[6, 10])

    d.add_paragraph()
    m.p("_______________________________", align="c")
    m.p(P.get("ingeniero", ""), bold=True, align="c")
    m.p(f"Ingeniero Civil – CIP {P.get('cip', '')}", align="c")

    d.save(ruta)
    return ruta


# =============================================================================
# 8. EXPORTACIÓN A AUTOCAD (.DXF) — LÁMINA CE-01
# =============================================================================
class Lamina:
    def __init__(self):
        self.doc = ezdxf.new("R2010", setup=True)
        self.doc.units = 6  # metros
        self.doc.header["$INSUNITS"] = 6
        self.doc.header["$MEASUREMENT"] = 1
        self.doc.header["$LTSCALE"] = 0.04
        self.msp = self.doc.modelspace()
        capas = [("E-CONCRETO", 7, "Continuous", 35), ("E-ACERO", 1, "Continuous", 30),
                 ("E-EQUIPOS", 4, "Continuous", 25), ("E-COTAS", 3, "Continuous", 13),
                 ("E-TEXTOS", 7, "Continuous", 18), ("E-HATCH", 8, "Continuous", 9),
                 ("E-TERRENO", 33, "Continuous", 13), ("E-EJES", 6, "CENTER", 13),
                 ("E-MARCO", 7, "Continuous", 50)]
        for n, c, lt, lw in capas:
            ly = self.doc.layers.add(n, color=c, linetype=lt)
            ly.dxf.lineweight = lw
        if "ARIAL" not in self.doc.styles:
            self.doc.styles.add("ARIAL", font="arial.ttf")
        ds = self.doc.dimstyles.new("CE-COTAS")
        ds.dxf.dimtxt = 0.065
        ds.dxf.dimasz = 0.05
        ds.dxf.dimexe = 0.04
        ds.dxf.dimexo = 0.03
        ds.dxf.dimgap = 0.02
        ds.dxf.dimdec = 2
        ds.dxf.dimzin = 0
        ds.dxf.dimdsep = 46
        ds.dxf.dimtad = 1
        ds.dxf.dimtih = 0
        ds.dxf.dimtoh = 0
        ds.dxf.dimtxsty = "ARIAL"
        ds.dxf.dimclrd = 3
        ds.dxf.dimclre = 3
        ds.dxf.dimclrt = 7
        try:
            ds.set_arrows(blk=ezdxf.ARROWS.architectural_tick)
        except Exception:
            pass

    # --- primitivas ---------------------------------------------------------------------
    def L(self, p1, p2, layer="E-CONCRETO", ltype=None, color=None):
        at = {"layer": layer}
        if ltype:
            at["linetype"] = ltype
        if color is not None:
            at["color"] = color
        return self.msp.add_line(p1, p2, dxfattribs=at)

    def PL(self, pts, layer="E-CONCRETO", close=False, ltype=None, color=None, width=0):
        at = {"layer": layer}
        if ltype:
            at["linetype"] = ltype
        if color is not None:
            at["color"] = color
        e = self.msp.add_lwpolyline(pts, close=close, dxfattribs=at)
        if width:
            e.dxf.const_width = width
        return e

    def R(self, x, y, w, h, layer="E-CONCRETO", ltype=None, color=None):
        return self.PL([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], layer, True, ltype, color)

    def C(self, c, r, layer="E-ACERO", ltype=None, color=None):
        at = {"layer": layer}
        if ltype:
            at["linetype"] = ltype
        if color is not None:
            at["color"] = color
        return self.msp.add_circle(c, r, dxfattribs=at)

    def A(self, c, r, a1, a2, layer="E-ACERO"):
        return self.msp.add_arc(c, r, a1, a2, dxfattribs={"layer": layer})

    def T(self, x, y, s, h=0.07, layer="E-TEXTOS", align="L", rot=0.0, color=None):
        at = {"layer": layer, "style": "ARIAL", "height": h, "rotation": rot}
        if color is not None:
            at["color"] = color
        e = self.msp.add_text(s, dxfattribs=at)
        al = {"L": TextEntityAlignment.LEFT, "C": TextEntityAlignment.CENTER,
              "R": TextEntityAlignment.RIGHT, "MC": TextEntityAlignment.MIDDLE_CENTER,
              "ML": TextEntityAlignment.MIDDLE_LEFT, "MR": TextEntityAlignment.MIDDLE_RIGHT}[align]
        e.set_placement((x, y), align=al)
        return e

    def HATCH(self, pts, pattern="AR-CONC", scale=0.02, layer="E-HATCH", color=8, angle=0.0):
        h = self.msp.add_hatch(color=color, dxfattribs={"layer": layer})
        if pattern == "SOLID":
            h.set_solid_fill(color=color)
        else:
            try:
                h.set_pattern_fill(pattern, color=color, scale=scale, angle=angle)
            except Exception:
                h.set_pattern_fill("ANSI31", color=color, scale=scale, angle=angle)
        h.paths.add_polyline_path(pts, is_closed=True)
        return h

    def HR(self, x, y, w, hh, pattern="AR-CONC", scale=0.02, color=8, angle=0.0):
        return self.HATCH([(x, y), (x + w, y), (x + w, y + hh), (x, y + hh)], pattern, scale,
                          color=color, angle=angle)

    def PUNTO(self, c, r, layer="E-ACERO", color=1):
        h = self.msp.add_hatch(color=color, dxfattribs={"layer": layer})
        h.set_solid_fill(color=color)
        ep = h.paths.add_edge_path()
        ep.add_arc(c, r, 0, 360)
        self.C(c, r, layer)

    def DIM(self, p1, p2, base, angle=0, text=None, lfac=None):
        ov = {}
        if lfac:
            ov["dimlfac"] = lfac
        try:
            kw = dict(base=base, p1=p1, p2=p2, angle=angle, dimstyle="CE-COTAS",
                      dxfattribs={"layer": "E-COTAS"}, override=ov or None)
            if text is not None:
                kw["text"] = text
            dim = self.msp.add_linear_dim(**kw)
            dim.render()
        except Exception:
            pass

    def iniciar_etiquetas(self):
        self._buf = []

    def volcar_etiquetas(self, sep=0.08):
        """Dibuja las etiquetas acumuladas evitando superposiciones (por columna)."""
        buf, self._buf = self._buf, None
        columnas = {}
        for it in buf:
            columnas.setdefault(round(it[1][0], 1), []).append(it)
        for col in columnas.values():
            col.sort(key=lambda it: -it[1][1])
            y_lim = None
            for target, pos, texto, h, lado in col:
                y = pos[1]
                if y_lim is not None:
                    y = min(y, y_lim - h - sep)
                self.FLECHA(target, (pos[0], y), texto, h, lado)
                y_lim = y - (texto.count("\n")) * h * 1.45

    def FLECHA(self, target, pos, texto, h=0.06, lado="R"):
        """Línea de referencia con punta de flecha y texto."""
        if getattr(self, "_buf", None) is not None:
            self._buf.append((target, pos, texto, h, lado))
            return
        tx, ty = target
        px, py = pos
        self.L((tx, ty), (px, py), "E-TEXTOS")
        ext = 0.12 if lado == "R" else -0.12
        self.L((px, py), (px + ext, py), "E-TEXTOS")
        ang = math.atan2(py - ty, px - tx)
        s = 0.045
        p1 = (tx + s * math.cos(ang + 0.30), ty + s * math.sin(ang + 0.30))
        p2 = (tx + s * math.cos(ang - 0.30), ty + s * math.sin(ang - 0.30))
        self.msp.add_solid([(tx, ty), p1, p2], dxfattribs={"layer": "E-TEXTOS"})
        lineas = texto.split("\n")
        for i, ln in enumerate(lineas):
            if lado == "R":
                self.T(px + ext + 0.04, py + 0.01 - i * h * 1.45, ln, h, align="L")
            else:
                self.T(px + ext - 0.04, py + 0.01 - i * h * 1.45, ln, h, align="R")

    def TITULO(self, xc, y, titulo, escala):
        y -= 0.04
        self.T(xc, y, titulo, 0.13, align="C")
        w = len(titulo) * 0.13 * 0.62
        self.PL([(xc - w / 2, y - 0.04), (xc + w / 2, y - 0.04)], "E-TEXTOS", width=0.012)
        self.L((xc - w / 2, y - 0.07), (xc + w / 2, y - 0.07), "E-TEXTOS")
        self.T(xc, y - 0.20, escala, 0.07, align="C")

    def CORTE(self, p1, p2, letra):
        """Línea de corte con letras y flechas de dirección (hacia +x o +y)."""
        self.L(p1, p2, "E-EJES", ltype="PHANTOM", color=6)
        for p in (p1, p2):
            self.C(p, 0.10, "E-TEXTOS")
            self.T(p[0], p[1], letra, 0.10, align="MC")
        horiz = abs(p2[0] - p1[0]) > abs(p2[1] - p1[1])
        for p in (p1, p2):
            if horiz:
                pts = [(p[0] - 0.05, p[1] + 0.14), (p[0] + 0.05, p[1] + 0.14), (p[0], p[1] + 0.26)]
            else:
                pts = [(p[0] + 0.14, p[1] - 0.05), (p[0] + 0.14, p[1] + 0.05), (p[0] + 0.26, p[1])]
            self.msp.add_solid(pts, dxfattribs={"layer": "E-TEXTOS"})

    def RESORTE(self, xc, y0, y1, w):
        n = 6
        pts = [(xc, y0)]
        for i in range(1, n * 2):
            yy = y0 + (y1 - y0) * i / (n * 2)
            xx = xc + (w / 2 if i % 2 else -w / 2)
            pts.append((xx, yy))
        pts.append((xc, y1))
        self.PL(pts, "E-EQUIPOS")

    def CORTE_ROTURA(self, x1, x2, y):
        xm = (x1 + x2) / 2
        self.PL([(x1, y), (xm - 0.05, y), (xm - 0.02, y + 0.08), (xm + 0.02, y - 0.08),
                 (xm + 0.05, y), (x2, y)], "E-CONCRETO")

    def TABLA(self, x, ytop, anchos, filas, h=0.065, titulo=None, nhead=1):
        rh = h * 2.1
        y = ytop
        W = sum(anchos)
        if titulo:
            self.R(x, y - rh * 1.2, W, rh * 1.2, "E-MARCO")
            self.HR(x, y - rh * 1.2, W, rh * 1.2, "SOLID", color=252)
            self.T(x + W / 2, y - rh * 0.6, titulo, h * 1.15, align="MC")
            y -= rh * 1.2
        for k, fila in enumerate(filas):
            self.R(x, y - rh, W, rh, "E-MARCO")
            if k < nhead:
                self.HR(x, y - rh, W, rh, "SOLID", color=254)
            xx = x
            for j, w in enumerate(anchos):
                if j > 0:
                    self.L((xx, y), (xx, y - rh), "E-MARCO")
                txt = str(fila[j]) if j < len(fila) else ""
                maxc = max(3, int(w / (h * 0.62)))
                if len(txt) > maxc:
                    txt = txt[:maxc - 1] + "…"
                if j == 0 and k >= nhead:
                    self.T(xx + 0.04, y - rh / 2, txt, h, align="ML")
                else:
                    self.T(xx + w / 2, y - rh / 2, txt, h, align="MC")
                xx += w
            y -= rh
        return ytop - y


def _tamanos_lamina(RE, RP):
    k = 5.0
    if RE.get("ok"):
        L, B, H, hef, da = RE["L"], RE["B"], RE["H"], RE["hef"], RE["da"] / 100.0
        wP1, hP1 = L + 1.8, B + 2.3
        hS1 = 0.9 + 0.45 + 0.10 + H + 0.03 + 0.13 + 0.20 + 1.40 + 0.35
        wS1 = B + 1.8 + 3.2
        hD = 0.9 + (hef + 0.12) * k + (0.03 + 0.0127 + 0.012 + 2.2 * da + 0.03) * k + 0.4
        wD = 0.40 * k + 3.4
    else:
        wP1 = hP1 = wS1 = hS1 = wD = hD = 0
    if RP.get("ok"):
        a, b, t, hp = RP["a"], RP["b"], RP["t"], RP["h_ped"]
        wP2, hP2 = max(a, b) + 1.9, max(a, b) + 2.3
        wS2 = a + 2 * (0.55 + 1.0) + 3.0
        hS2 = 0.9 + 0.8 + 0.10 + t + hp + 1.9 + 0.3
    else:
        wP2 = hP2 = wS2 = hS2 = 0
    return dict(k=k, wP1=wP1, hP1=hP1, wS1=wS1, hS1=hS1, wD=wD, hD=hD,
                wP2=wP2, hP2=hP2, wS2=wS2, hS2=hS2)


def dxf_planta_dado(D, RE, ox, oy, w):
    L, B = RE["L"], RE["B"]
    Leq, Beq = RE["Leq"], RE["Beq"]
    bx, by = ox + 0.9, oy + 1.25
    cx, cy = bx + L / 2, by + B / 2
    D.TITULO(ox + w / 2, oy + 0.42, "PLANTA - DADO DE CIMENTACIÓN ANTIVIBRATORIO", "ESC. 1:25")
    # Piso terminado
    D.R(bx - 0.55, by - 0.45, L + 1.10, B + 0.90, "E-CONCRETO", ltype="DASHED", color=8)
    D.T(bx - 0.50, by + B + 0.33, "PISO TERMINADO SALA DE MÁQUINAS  NPT ±0.00", 0.05)
    # Bloque y junta
    D.R(bx, by, L, B, "E-CONCRETO")
    D.R(bx - 0.025, by - 0.025, L + 0.05, B + 0.05, "E-CONCRETO", color=30)
    D.HATCH([(bx - 0.025, by - 0.025), (bx + L + 0.025, by - 0.025), (bx + L + 0.025, by + B + 0.025),
             (bx - 0.025, by + B + 0.025), (bx - 0.025, by - 0.025), (bx, by), (bx, by + B),
             (bx + L, by + B), (bx + L, by), (bx, by)], "SOLID", color=30)
    # Malla de acero (parrilla superior, línea discontinua)
    rec, s = RE["rec"], RE["s"]
    y = by + rec
    while y <= by + B - rec + 1e-6:
        D.L((bx + rec, y), (bx + L - rec, y), "E-ACERO", ltype="DASHED")
        y += s
    x = bx + rec
    while x <= bx + L - rec + 1e-6:
        D.L((x, by + rec), (x, by + B - rec), "E-ACERO", ltype="DASHED")
        x += s
    # Chasis y rieles
    ex, ey = cx - Leq / 2, cy - Beq / 2
    D.R(ex, ey, Leq, Beq, "E-EQUIPOS", ltype="DASHDOT")
    for sgn in (-1, 1):
        yr = cy + sgn * RE["yrail"]
        D.R(ex, yr - 0.0375, Leq, 0.075, "E-EQUIPOS")
    D.T(cx, cy + 0.10, "GRUPO ELECTRÓGENO" if "Grupo" in RE["tipo"] else RE["tipo"].upper(), 0.08,
        "E-EQUIPOS", align="MC")
    D.T(cx, cy - 0.06, f"W = {f(RE['W_kg'], 0)} kg   {f(RE['rpm'], 0)} RPM", 0.06, "E-EQUIPOS", align="MC")
    # Amortiguadores y pernos
    for (px, py) in RE["pernos_xy"]:
        D.R(cx + px - 0.10, cy + py - 0.08, 0.20, 0.16, "E-EQUIPOS")
        D.C((cx + px, cy + py), RE["da"] / 200.0 + 0.004, "E-ACERO")
        D.L((cx + px - 0.05, cy + py), (cx + px + 0.05, cy + py), "E-ACERO")
        D.L((cx + px, cy + py - 0.05), (cx + px, cy + py + 0.05), "E-ACERO")
    # Conduit
    xc_c = cx + Leq / 2 - 0.45
    D.C((xc_c, cy), 0.057, "E-EQUIPOS")
    D.C((xc_c, cy), 0.050, "E-EQUIPOS")
    # Ejes
    D.L((cx, by - 0.35), (cx, by + B + 0.35), "E-EJES", ltype="CENTER")
    D.L((bx - 0.35, cy), (bx + L + 0.35, cy), "E-EJES", ltype="CENTER")
    # Corte A-A (transversal, a través de un perno)
    xa = cx + RE["pernos_xy"][0][0]
    D.CORTE((xa, by - 0.62), (xa, by + B + 0.62), "A")
    # Cotas
    D.DIM((bx, by), (bx + L, by), (0, by - 0.30))
    D.DIM((bx, by), (bx, by + B), (bx - 0.30, 0), angle=90)
    D.DIM((bx, by + B), (ex, by + B), (0, by + B + 0.17))
    D.DIM((ex, by + B), (ex + Leq, by + B), (0, by + B + 0.17))
    D.DIM((ex + Leq, by + B), (bx + L, by + B), (0, by + B + 0.17))
    D.DIM((bx + L, ey), (bx + L, ey + Beq), (bx + L + 0.22, 0), angle=90)
    # Referencias
    D.FLECHA((bx + L + 0.012, by + 0.15), (bx + L + 0.30, by - 0.20), 'JUNTA PERIMETRAL e=1"', 0.05)
    D.FLECHA((xc_c, cy - 0.05), (xc_c + 0.15, by + 0.12), 'CONDUIT PVC-SAP Ø4"', 0.05)
    p0 = RE["pernos_xy"][-1]
    D.FLECHA((cx + p0[0] + 0.10, cy + p0[1]), (bx + L - 0.05, by + B + 0.33),
             f"AMORTIG. + PERNO J Ø{RE['dp']}", 0.05)
    D.T(bx + 0.05, by - 0.12, f"PARRILLA Ø{RE['barra']} @ {f(RE['s'])} SUP. E INF. (AMBAS DIR.)", 0.045,
        "E-ACERO", color=1)


def dxf_corte_dado(D, RE, ox, oy):
    B, H, hs = RE["B"], RE["H"], RE["h_sobre"]
    Beq = RE["Beq"]
    cx = ox + 0.9 + B / 2
    y0 = oy + 0.9 + 0.45 + 0.10 + (H - hs)        # NPT ±0.00
    yt = y0 + hs                                     # cara superior del bloque
    yb = yt - H                                      # fondo del bloque
    D.TITULO(cx + 0.8, oy + 0.42, "CORTE A-A  (TRANSVERSAL)", "ESC. 1:25")
    D.iniciar_etiquetas()
    xl, xr = cx - B / 2, cx + B / 2
    # Terreno / relleno
    ybot = yb - 0.10 - 0.40
    for (x1, x2) in ((xl - 0.85, xl - 0.025), (xr + 0.025, xr + 0.85)):
        D.HATCH([(x1, ybot), (x2, ybot), (x2, y0 - 0.10), (x1, y0 - 0.10)], "EARTH", 0.03,
                color=33, layer="E-TERRENO")
    D.HATCH([(xl - 0.025, ybot), (xr + 0.025, ybot), (xr + 0.10, yb - 0.10), (xl - 0.10, yb - 0.10)],
            "EARTH", 0.03, color=33, layer="E-TERRENO")
    D.L((xl - 0.85, ybot), (xr + 0.85, ybot), "E-TERRENO")
    # Piso terminado (losa e=0.10)
    for (x1, x2) in ((xl - 0.85, xl - 0.025), (xr + 0.025, xr + 0.85)):
        D.R(x1, y0 - 0.10, x2 - x1, 0.10, "E-CONCRETO")
        D.HR(x1, y0 - 0.10, x2 - x1, 0.10, "AR-SAND", 0.004, color=8)
    D.T(xl - 0.83, y0 + 0.03, "NPT ±0.00", 0.05)
    # Solado
    D.R(xl - 0.10, yb - 0.10, B + 0.20, 0.10, "E-CONCRETO")
    D.HR(xl - 0.10, yb - 0.10, B + 0.20, 0.10, "AR-SAND", 0.004, color=8)
    # Bloque
    D.R(xl, yb, B, H, "E-CONCRETO")
    D.HR(xl, yb, B, H, "AR-CONC", 0.012, color=8)
    # Junta perimetral y sello
    for xj in (xl - 0.025, xr):
        D.R(xj, yb + 0.05, 0.025, (y0 - 0.015) - (yb + 0.05), "E-CONCRETO", color=30)
        D.HR(xj, yb + 0.05, 0.025, (y0 - 0.015) - (yb + 0.05), "ANSI37", 0.006, color=30)
        D.R(xj, y0 - 0.015, 0.025, 0.015, "E-CONCRETO", color=1)
        D.HR(xj, y0 - 0.015, 0.025, 0.015, "SOLID", color=1)
    # Acero: doble parrilla
    rec = RE["rec"]
    rb = RE["db"] / 200.0
    pata = min(0.20, H / 2 - rec)
    ysup, yinf = yt - rec, yb + rec
    D.PL([(xl + rec, ysup - pata), (xl + rec, ysup), (xr - rec, ysup), (xr - rec, ysup - pata)], "E-ACERO",
         width=2 * rb)
    D.PL([(xl + rec, yinf + pata), (xl + rec, yinf), (xr - rec, yinf), (xr - rec, yinf + pata)], "E-ACERO",
         width=2 * rb)
    s = RE["s"]
    n = int((B - 2 * rec - 2 * rb) / s) + 1
    esp = (B - 2 * rec - 4 * rb) / max(1, n - 1)
    for i in range(n):
        xx = xl + rec + 2 * rb + i * esp
        D.PUNTO((xx, ysup - 2 * rb), rb)
        D.PUNTO((xx, yinf + 2 * rb), rb)
    nl = max(1, int((H - 2 * rec) / s) - 1)
    for i in range(1, nl + 1):
        yy = yinf + (ysup - yinf) * i / (nl + 1)
        D.PUNTO((xl + rec + 2 * rb, yy), rb)
        D.PUNTO((xr - rec - 2 * rb, yy), rb)
    # Pernos J, camisas, grout, amortiguadores, rieles
    hef = RE["hef"]
    da = RE["da"] / 100.0
    eg = RE["e_grout"] / 100.0
    for sgn in (-1, 1):
        xp = cx + sgn * RE["yrail"]
        # camisa
        D.R(xp - 0.0375, yt - 0.20, 0.075, 0.20, "E-ACERO", color=6)
        # perno
        ybot_p = yt - hef
        D.L((xp, ybot_p + 0.04), (xp, yt + eg + 0.13 + 0.06), "E-ACERO", color=1)
        D.A((xp + sgn * 0.04, ybot_p + 0.04), 0.04, 180 if sgn > 0 else 180, 360)
        D.L((xp + sgn * 0.08, ybot_p + 0.04), (xp + sgn * 0.08, ybot_p + 0.12), "E-ACERO", color=1)
        # grout
        D.R(xp - 0.125, yt, 0.25, eg, "E-CONCRETO")
        D.HR(xp - 0.125, yt, 0.25, eg, "AR-SAND", 0.003, color=9)
        # amortiguador
        y_a0 = yt + eg
        D.R(xp - 0.09, y_a0, 0.18, 0.012, "E-EQUIPOS")
        D.R(xp - 0.08, y_a0 + 0.012, 0.16, 0.105, "E-EQUIPOS")
        D.RESORTE(xp, y_a0 + 0.02, y_a0 + 0.11, 0.08)
        D.R(xp - 0.09, y_a0 + 0.117, 0.18, 0.013, "E-EQUIPOS")
        D.R(xp - 0.03, y_a0 + 0.13, 0.06, 0.03, "E-ACERO", color=1)   # tuercas
        # riel del skid (perfil C)
        yr0 = y_a0 + 0.16
        xi, xo = xp - sgn * 0.0375, xp + sgn * 0.0375
        D.PL([(xo, yr0), (xi, yr0), (xi, yr0 + 0.20), (xo, yr0 + 0.20)], "E-EQUIPOS")
    y_eq = yt + eg + 0.16 + 0.20
    D.R(cx - Beq / 2, y_eq, Beq, 1.40, "E-EQUIPOS", ltype="DASHDOT")
    D.T(cx, y_eq + 0.80, "GRUPO ELECTRÓGENO" if "Grupo" in RE["tipo"] else "EQUIPO", 0.07, "E-EQUIPOS",
        align="MC")
    D.T(cx, y_eq + 0.65, "ENCAPSULADO (CHASIS / SKID)", 0.05, "E-EQUIPOS", align="MC")
    # Conduit (salida de cables)
    xc1 = cx + 0.10
    ycurve = yb + 0.25
    for off in (-0.05, 0.05):
        D.PL([(xc1 + off, yt + 0.08), (xc1 + off, ycurve + off), (xr + 0.85, ycurve + off)], "E-EQUIPOS",
             color=5)
    # Cotas
    D.DIM((xl, yb - 0.10), (xr, yb - 0.10), (0, ybot - 0.12))
    D.DIM((xr + 0.025, yb), (xr + 0.025, yt), (xr + 0.55, 0), angle=90)
    D.DIM((xr + 0.025, y0), (xr + 0.025, yt), (xr + 0.35, 0), angle=90)
    D.DIM((xr + 0.10, yb - 0.10), (xr + 0.10, yb), (xr + 0.35, 0), angle=90)
    D.DIM((cx - RE["yrail"], yt), (cx - RE["yrail"], yt - hef), (xl - 0.40, 0), angle=90)
    # Referencias (al lado derecho)
    xt = xr + 1.05
    D.FLECHA((xr + 0.012, y0 - 0.25), (xt - 0.20, y0 - 0.05), 'JUNTA PERIMETRAL POLIESTIRENO\nELASTIFICADO/NEOPRENO e=1"', 0.05)
    D.FLECHA((xr + 0.012, y0 - 0.008), (xt - 0.20, y0 + 0.25), "SELLO ELASTOMÉRICO\nPOLIURETANO", 0.05)
    D.FLECHA((xr - 0.30, yb + H * 0.55), (xt - 0.20, yb + H * 0.55),
             f"CONCRETO f'c={f(RE['fc'], 0)} kg/cm²\nPARRILLA Ø{RE['barra']} @ {f(RE['s'])}", 0.05)
    D.FLECHA((xr + 0.05, yb - 0.05), (xt - 0.20, yb - 0.12), "SOLADO C:H 1:12 e=0.10", 0.05)
    D.FLECHA((xr + 0.6, ybot + 0.15), (xt - 0.20, ybot + 0.02), "RELLENO COMPACTADO / TERRENO", 0.05)
    D.FLECHA((xr + 0.6, ycurve), (xt - 0.20, ycurve + 0.15), 'CONDUIT PVC-SAP Ø4"\nCABLES DE FUERZA', 0.05)
    xp = cx + RE["yrail"]
    D.FLECHA((xp + 0.09, yt + eg + 0.06), (xt - 0.20, yt + 0.45), "AMORTIGUADOR SPRING /\nELASTÓMERO", 0.05)
    D.FLECHA((xp + 0.125, yt + eg / 2), (xt - 0.20, yt + 0.15), f"GROUT NO CONTRÁCTIL e={f(RE['e_grout'], 1)} cm", 0.05)
    D.FLECHA((xp + 0.04, yt + eg + 0.26), (xt - 0.20, yt + 0.75), "PERFIL DEL CHASIS (SKID)", 0.05)
    D.volcar_etiquetas()


def dxf_detalle_perno(D, RE, ox, oy, k):
    hef = RE["hef"]
    da = RE["da"] / 100.0
    eg = RE["e_grout"] / 100.0
    x = ox + 1.10
    yt = oy + 0.9 + (hef + 0.12) * k
    D.TITULO(ox + 1.9, oy + 0.42, f"DETALLE PERNO TIPO J Ø{RE['dp']}", "ESC. 1:5")
    D.iniciar_etiquetas()
    # Concreto
    wc = 0.40 * k
    D.PL([(x - wc / 2, yt), (x + wc / 2, yt)], "E-CONCRETO")
    D.L((x - wc / 2, yt), (x - wc / 2, yt - (hef + 0.12) * k), "E-CONCRETO")
    D.L((x + wc / 2, yt), (x + wc / 2, yt - (hef + 0.12) * k), "E-CONCRETO")
    D.CORTE_ROTURA(x - wc / 2, x + wc / 2, yt - (hef + 0.12) * k)
    D.HR(x - wc / 2, yt - (hef + 0.12) * k, wc, (hef + 0.12) * k, "AR-CONC", 0.018, color=8)
    # Camisa PVC Ø3" L=0.20 con grout
    rs = 0.0381 * k
    D.R(x - rs, yt - 0.20 * k, 2 * rs, 0.20 * k, "E-ACERO", color=6)
    D.HR(x - rs, yt - 0.20 * k, 2 * rs, 0.20 * k, "AR-SAND", 0.01, color=9)
    # Perno: fuste
    r = da / 2 * k
    R = 3 * da * k                     # radio de doblez
    ybend = yt - hef * k + R
    y_top = yt + (eg + 0.0127 + 0.004 + 0.006 + 2 * 0.9 * da + 0.02) * k
    D.L((x - r, ybend), (x - r, y_top), "E-ACERO", color=1)
    D.L((x + r, ybend), (x + r, y_top), "E-ACERO", color=1)
    D.L((x - r, y_top), (x + r, y_top), "E-ACERO", color=1)
    # Gancho J
    cxj = x + R
    D.A((cxj, ybend), R + r, 180, 360)
    D.A((cxj, ybend), R - r, 180, 360)
    hj = 1.5 * da * k * 1.5
    D.L((cxj + R - r, ybend), (cxj + R - r, ybend + hj), "E-ACERO", color=1)
    D.L((cxj + R + r, ybend), (cxj + R + r, ybend + hj), "E-ACERO", color=1)
    D.L((cxj + R - r, ybend + hj), (cxj + R + r, ybend + hj), "E-ACERO", color=1)
    # Rosca (líneas finas)
    for i in range(10):
        yy = y_top - 0.02 * k - i * 0.004 * k * 2
        D.L((x - r, yy), (x + r, yy - 0.002 * k), "E-ACERO", color=1)
    # Grout
    D.R(x - 0.10 * k, yt, 0.20 * k, eg * k, "E-CONCRETO")
    D.HR(x - 0.10 * k, yt, 0.20 * k, eg * k, "AR-SAND", 0.01, color=9)
    # Plancha base
    yp = yt + eg * k
    D.R(x - 0.09 * k, yp, 0.18 * k, 0.0127 * k, "E-EQUIPOS")
    D.HR(x - 0.09 * k, yp, 0.18 * k, 0.0127 * k, "ANSI31", 0.01, color=4)
    # Arandela plana, grover, tuerca, contratuerca
    y1 = yp + 0.0127 * k
    D.R(x - 1.1 * da * k, y1, 2.2 * da * k, 0.004 * k, "E-ACERO", color=2)
    y2 = y1 + 0.004 * k
    D.R(x - 0.9 * da * k, y2, 1.8 * da * k, 0.006 * k, "E-ACERO", color=2)
    y3 = y2 + 0.006 * k
    hn = 0.9 * da * k
    D.R(x - 0.85 * da * k, y3, 1.7 * da * k, hn, "E-ACERO", color=2)
    D.R(x - 0.85 * da * k, y3 + hn, 1.7 * da * k, hn, "E-ACERO", color=2)
    for yy in (y3, y3 + hn):
        D.L((x - 0.45 * da * k, yy), (x - 0.45 * da * k, yy + hn), "E-ACERO", color=2)
        D.L((x + 0.45 * da * k, yy), (x + 0.45 * da * k, yy + hn), "E-ACERO", color=2)
    # Cotas (dimlfac = 1/k)
    D.DIM((x - wc / 2, yt), (x - wc / 2, yt - hef * k), (x - wc / 2 - 0.30, 0), angle=90, lfac=1 / k)
    D.DIM((x - 0.10 * k, yt), (x - 0.10 * k, yt + eg * k), (x - wc / 2 - 0.30, 0), angle=90, lfac=1 / k)
    D.DIM((x - rs, yt - 0.20 * k), (x - rs, yt), (x - wc / 2 - 0.10, 0), angle=90, lfac=1 / k)
    # Referencias
    xt = x + wc / 2 + 0.55
    D.FLECHA((x + 0.85 * da * k, y3 + 1.5 * hn), (xt, y_top + 0.20), "CONTRATUERCA HEX. PESADA A563", 0.055)
    D.FLECHA((x + 0.85 * da * k, y3 + 0.5 * hn), (xt, y_top - 0.05), "TUERCA HEX. PESADA A563", 0.055)
    D.FLECHA((x + 0.9 * da * k, y2 + 0.003 * k), (xt, y_top - 0.30), "ARANDELA DE PRESIÓN (GROVER)", 0.055)
    D.FLECHA((x + 1.1 * da * k, y1 + 0.002 * k), (xt, y_top - 0.55), "ARANDELA PLANA ASTM F436", 0.055)
    D.FLECHA((x + 0.09 * k, yp + 0.006 * k), (xt, y_top - 0.80), "PLANCHA BASE AMORTIGUADOR/CHASIS", 0.055)
    D.FLECHA((x + 0.10 * k, yt + eg * k / 2), (xt, y_top - 1.05),
             f"GROUT EPÓXICO NO CONTRÁCTIL e={f(RE['e_grout'], 1)} cm", 0.055)
    D.FLECHA((x + rs, yt - 0.10 * k), (xt, yt - 0.35), 'CAMISA PVC Ø3" L=0.20 m\nRELLENA CON GROUT', 0.055)
    D.FLECHA((x + r, yt - hef * k * 0.6), (xt, yt - hef * k * 0.6),
             f"PERNO J Ø{RE['dp']} {RE['grado'].replace('ASTM ', '')}\nGALVANIZADO EN CALIENTE", 0.055)
    D.FLECHA((cxj + R + r, ybend), (xt, ybend - 0.10), "GANCHO J: R=3d , L=4.5d", 0.055)
    D.FLECHA((x + wc / 2 - 0.05, yt - (hef + 0.08) * k), (xt, yt - (hef + 0.12) * k + 0.05),
             f"CONCRETO DEL BLOQUE f'c={f(RE['fc'], 0)}", 0.055)
    D.volcar_etiquetas()


def dxf_planta_macizo(D, RP, ox, oy, w):
    a, b, ap = RP["a"], RP["b"], RP["a_ped"]
    cx, cy = ox + 0.95 + a / 2, oy + 1.25 + b / 2
    D.TITULO(ox + w / 2, oy + 0.42, "PLANTA - MACIZO DE POSTE", "ESC. 1:25")
    D.R(cx - a / 2, cy - b / 2, a, b, "E-CONCRETO", ltype="DASHED")
    D.R(cx - ap / 2, cy - ap / 2, ap, ap, "E-CONCRETO")
    D.R(cx - ap / 2 + 0.05, cy - ap / 2 + 0.05, ap - 0.10, ap - 0.10, "E-CONCRETO", color=8)
    for (x1, y1), (x2, y2) in (((-1, -1), (1, 1)), ((-1, 1), (1, -1))):
        D.L((cx + x1 * ap / 2, cy + y1 * ap / 2), (cx + x1 * (ap / 2 - 0.05), cy + y1 * (ap / 2 - 0.05)),
            "E-CONCRETO", color=8)
        D.L((cx + x2 * ap / 2, cy + y2 * ap / 2), (cx + x2 * (ap / 2 - 0.05), cy + y2 * (ap / 2 - 0.05)),
            "E-CONCRETO", color=8)
    # Malla inferior
    s, rec = RP["s_maz"], 0.075
    yy = cy - b / 2 + rec
    while yy <= cy + b / 2 - rec + 1e-6:
        D.L((cx - a / 2 + rec, yy), (cx + a / 2 - rec, yy), "E-ACERO", ltype="DASHED")
        yy += s
    xx = cx - a / 2 + rec
    while xx <= cx + a / 2 - rec + 1e-6:
        D.L((xx, cy - b / 2 + rec), (xx, cy + b / 2 - rec), "E-ACERO", ltype="DASHED")
        xx += s
    # Poste / plancha y pernos
    if RP["empotrado"]:
        D.C((cx, cy), RP["d_base"] / 2, "E-EQUIPOS")
        D.C((cx, cy), RP["d_base"] / 2 - 0.06, "E-EQUIPOS")
    else:
        pl = RP["Dc"] + 0.12
        n = 8
        pts = [(cx + pl / 2 / math.cos(math.pi / n) * math.cos(math.pi / n + 2 * math.pi * i / n),
                cy + pl / 2 / math.cos(math.pi / n) * math.sin(math.pi / n + 2 * math.pi * i / n))
               for i in range(n)]
        D.PL(pts, "E-EQUIPOS", close=True)
        d0 = RP["d_base"]
        pts2 = [(cx + d0 / 2 * math.cos(math.pi / 8 + 2 * math.pi * i / 8),
                 cy + d0 / 2 * math.sin(math.pi / 8 + 2 * math.pi * i / 8)) for i in range(8)]
        D.PL(pts2, "E-EQUIPOS", close=True)
        D.C((cx, cy), RP["Dc"] / 2, "E-EJES", ltype="CENTER")
        for i in range(RP["n_pp"]):
            ang = 2 * math.pi * i / RP["n_pp"] + math.pi / RP["n_pp"]
            D.C((cx + RP["Dc"] / 2 * math.cos(ang), cy + RP["Dc"] / 2 * math.sin(ang)),
                RP["da_p"] / 200.0 + 0.005, "E-ACERO")
    # PVC puesta a tierra
    D.C((cx + RP["d_base"] / 2 * 0.45, cy - RP["d_base"] / 2 * 0.35), 0.017, "E-EQUIPOS", color=5)
    D.L((cx, cy - b / 2 - 0.30), (cx, cy + b / 2 + 0.30), "E-EJES", ltype="CENTER")
    D.CORTE((cx - a / 2 - 0.55, cy), (cx + a / 2 + 0.55, cy), "B")
    D.DIM((cx - a / 2, cy - b / 2), (cx + a / 2, cy - b / 2), (0, cy - b / 2 - 0.28))
    D.DIM((cx + a / 2, cy - b / 2), (cx + a / 2, cy + b / 2), (cx + a / 2 + 0.28, 0), angle=90)
    D.DIM((cx - ap / 2, cy + ap / 2), (cx + ap / 2, cy + ap / 2), (0, cy + b / 2 + 0.15))
    if not RP["empotrado"]:
        D.T(cx, cy + b / 2 + 0.42, f"{RP['n_pp']} PERNOS Ø{RP['d_pp']} EN CÍRCULO Dc={f(RP['Dc'])}", 0.05,
            align="C")
    D.T(cx - a / 2, cy - b / 2 - 0.47, f"MALLA Ø{RP['barra']} @ {f(RP['s_maz'])} INF. Y SUP.", 0.045,
        "E-ACERO", color=1)


def dxf_corte_macizo(D, RP, ox, oy):
    a, t, ap, hp = RP["a"], RP["t"], RP["a_ped"], RP["h_ped"]
    cx = ox + 1.0 + 0.55 + a / 2
    y0 = oy + 0.9 + 0.8 + 0.10 + t              # NTN
    D.TITULO(cx + 1.0, oy + 0.42, "CORTE B-B  MACIZO DE POSTE (SULZBERGER)", "ESC. 1:25")
    D.iniciar_etiquetas()
    xl, xr = cx - a / 2, cx + a / 2
    yb = y0 - t
    # Terreno
    for (x1, x2) in ((xl - 0.55, xl), (xr, xr + 0.55)):
        D.HATCH([(x1, yb - 0.30), (x2, yb - 0.30), (x2, y0), (x1, y0)], "EARTH", 0.03, color=33,
                layer="E-TERRENO")
    D.HATCH([(xl, yb - 0.30), (xr, yb - 0.30), (xr, yb - 0.10), (xl, yb - 0.10)], "EARTH", 0.03, color=33,
            layer="E-TERRENO")
    D.L((xl - 0.55, y0), (xl, y0), "E-TERRENO")
    D.L((xr, y0), (xr + 0.55, y0), "E-TERRENO")
    D.T(xl - 0.53, y0 + 0.03, "NTN ±0.00", 0.05)
    # Solado, macizo
    D.R(xl - 0.10, yb - 0.10, a + 0.20, 0.10, "E-CONCRETO")
    D.HR(xl - 0.10, yb - 0.10, a + 0.20, 0.10, "AR-SAND", 0.004, color=8)
    D.R(xl, yb, a, t, "E-CONCRETO")
    D.HR(xl, yb, a, t, "AR-CONC", 0.012, color=8)
    # Pedestal con chaflán y pendiente superior
    ch = 0.05
    yt = y0 + hp
    ped = [(cx - ap / 2, y0), (cx + ap / 2, y0), (cx + ap / 2, yt - ch), (cx + ap / 2 - ch, yt - 0.01),
           (cx, yt), (cx - ap / 2 + ch, yt - 0.01), (cx - ap / 2, yt - ch)]
    D.PL(ped, "E-CONCRETO", close=True)
    D.HATCH(ped, "AR-CONC", 0.012, color=8)
    # Acero del macizo: malla inferior y superior
    rec = 0.075
    rb = RP["db"] / 200.0
    D.PL([(xl + rec, yb + rec + 0.15), (xl + rec, yb + rec), (xr - rec, yb + rec), (xr - rec, yb + rec + 0.15)],
         "E-ACERO", width=2 * rb)
    D.PL([(xl + rec, y0 - rec - 0.15), (xl + rec, y0 - rec), (xr - rec, y0 - rec), (xr - rec, y0 - rec - 0.15)],
         "E-ACERO", width=2 * rb)
    n = int((a - 2 * rec) / RP["s_maz"]) + 1
    for i in range(n):
        xx = xl + rec + 2 * rb + i * (a - 2 * rec - 4 * rb) / max(1, n - 1)
        D.PUNTO((xx, yb + rec + 2 * rb), rb)
        D.PUNTO((xx, y0 - rec - 2 * rb), rb)
    # Verticales del pedestal con estribos
    for sgn in (-1, 1):
        xv = cx + sgn * (ap / 2 - 0.05 - rb)
        D.PL([(xv + sgn * 0.25, yb + rec + 2 * rb), (xv, yb + rec + 2 * rb), (xv, yt - 0.05)], "E-ACERO",
             width=2 * rb)
    ye = yb + rec + 0.10
    while ye < yt - 0.06:
        D.L((cx - ap / 2 + 0.04, ye), (cx + ap / 2 - 0.04, ye), "E-ACERO", color=1)
        ye += 0.15
    # Poste
    db_ = RP["d_base"]
    hstub = 1.70
    dsup = db_ + (RP["d_top"] - db_) * hstub / RP["Hl"]
    if RP["empotrado"]:
        e = RP["e_emp"]
        yp0 = yt - e
        D.PL([(cx - db_ / 2, yp0), (cx - dsup / 2, yt + hstub)], "E-EQUIPOS")
        D.PL([(cx + db_ / 2, yp0), (cx + dsup / 2, yt + hstub)], "E-EQUIPOS")
        D.L((cx - db_ / 2, yp0), (cx + db_ / 2, yp0), "E-EQUIPOS")
        D.T(cx, yp0 - 0.10, f"EMPOTRAMIENTO e={f(e)} m", 0.05, align="C")
    else:
        pl = RP["Dc"] + 0.12
        D.R(cx - pl / 2, yt + 0.03, pl, 0.025, "E-EQUIPOS")
        D.HR(cx - pl / 2, yt + 0.03, pl, 0.025, "ANSI31", 0.01, color=4)
        D.R(cx - pl / 2 + 0.02, yt, pl - 0.04, 0.03, "E-CONCRETO")
        D.HR(cx - pl / 2 + 0.02, yt, pl - 0.04, 0.03, "AR-SAND", 0.003, color=9)
        D.L((cx - db_ / 2, yt + 0.055), (cx - dsup / 2, yt + hstub), "E-EQUIPOS")
        D.L((cx + db_ / 2, yt + 0.055), (cx + dsup / 2, yt + hstub), "E-EQUIPOS")
        for sgn in (-1, 1):
            xb = cx + sgn * RP["Dc"] / 2
            ybb = yt - RP["L_anc"]
            D.L((xb, ybb), (xb, yt + 0.12), "E-ACERO", color=1)
            D.L((xb, ybb), (xb - sgn * 0.10, ybb), "E-ACERO", color=1)
            D.R(xb - 0.025, yt + 0.055, 0.05, 0.04, "E-ACERO", color=2)
        D.L((cx - RP["Dc"] / 2, yt - RP["L_anc"] + 0.10), (cx + RP["Dc"] / 2, yt - RP["L_anc"] + 0.10),
            "E-ACERO", color=1)
    D.CORTE_ROTURA(cx - dsup / 2 - 0.05, cx + dsup / 2 + 0.05, yt + hstub)
    tit = RP["tipo"].upper()
    if len(tit) > 30 and " CON " in tit:
        l1, l2 = tit.split(" CON ", 1)
        D.T(cx, yt + hstub + 0.33, l1, 0.055, "E-EQUIPOS", align="C")
        D.T(cx, yt + hstub + 0.24, "CON " + l2, 0.055, "E-EQUIPOS", align="C")
    else:
        D.T(cx, yt + hstub + 0.24, tit, 0.055, "E-EQUIPOS", align="C")
    D.T(cx, yt + hstub + 0.12, f"H = {f(RP['Hp'])} m", 0.055, "E-EQUIPOS", align="C")
    # Fuerza de viento
    D.L((cx - 0.75, yt + hstub - 0.20), (cx - dsup / 2 - 0.03, yt + hstub - 0.20), "E-COTAS", color=1)
    D.msp.add_solid([(cx - dsup / 2 - 0.03, yt + hstub - 0.20), (cx - dsup / 2 - 0.12, yt + hstub - 0.17),
                     (cx - dsup / 2 - 0.12, yt + hstub - 0.23)], dxfattribs={"layer": "E-COTAS", "color": 1})
    D.T(cx - 0.75, yt + hstub - 0.15, f"Ht={f(RP['Ht'] * 1000, 0)} kg", 0.05, color=1)
    D.T(cx - 0.75, yt + hstub - 0.32, f"Mv={f(RP['Mv'], 2)} t·m", 0.05, color=1)
    # Tubo PVC puesta a tierra
    xg = cx + db_ * 0.25
    for off in (-0.017, 0.017):
        D.PL([(xg + off, yt), (xg + off, y0 - 0.45 + off), (xr + 0.55, y0 - 0.45 + off)], "E-EQUIPOS", color=5)
    # Eje de giro a 2t/3
    yg = y0 - 2 * t / 3
    D.L((xl - 0.10, yg), (xr + 0.10, yg), "E-EJES", ltype="CENTER")
    D.C((cx, yg), 0.03, "E-COTAS", color=1)
    D.T(xl + 0.05, yg + 0.04, "EJE DE GIRO 2t/3", 0.045, color=6)
    # Diagramas de presiones de Sulzberger (parabólicos laterales + triangular de fondo)
    sig_max = max(RP["sig_lat"], RP["sig_lat_inf"], 1e-6)
    esc = 0.75 / sig_max
    npts = 24
    # Lado derecho: de 0 a 2t/3 (empuje pasivo superior)
    xd = xr + 0.60
    D.L((xd, y0), (xd, y0 - t), "E-COTAS")
    pts = []
    for i in range(npts + 1):
        yy = 2 * t / 3 * i / npts
        sg = RP["Ct"] * (yy / t) * (2 * t / 3 - yy) * RP["tan_req"]
        pts.append((xd + sg * esc, y0 - yy))
    D.PL(pts, "E-COTAS", color=1)
    D.HATCH([(xd, y0)] + pts + [(xd, y0 - 2 * t / 3)], "ANSI31", 0.008, color=1, layer="E-COTAS")
    # Lado izquierdo: de 2t/3 a t
    xi = xl - 0.60
    D.L((xi, y0), (xi, y0 - t), "E-COTAS")
    pts2 = []
    for i in range(npts + 1):
        yy = 2 * t / 3 + (t / 3) * i / npts
        sg = RP["Ct"] * (yy / t) * (yy - 2 * t / 3) * RP["tan_req"]
        pts2.append((xi - sg * esc, y0 - yy))
    D.PL(pts2, "E-COTAS", color=1)
    D.HATCH([(xi, y0 - 2 * t / 3)] + pts2 + [(xi, y0 - t)], "ANSI31", 0.008, color=1, layer="E-COTAS")
    D.T(xd + 0.05, y0 - t / 3 + 0.10, f"σmáx={f(RP['sig_lat'] / 10, 3)}", 0.045, color=1)
    D.T(xd + 0.05, y0 - t / 3 - 0.02, "kg/cm²", 0.045, color=1)
    D.T(xi - 0.05, y0 - t + 0.10, f"σmáx={f(RP['sig_lat_inf'] / 10, 3)}", 0.045, align="R", color=1)
    D.T(xi - 0.05, y0 - t - 0.02, "kg/cm²", 0.045, align="R", color=1)
    # Fondo
    yd = yb - 0.42
    D.L((xl, yd), (xr, yd), "E-COTAS")
    a0 = min(RP["a0r"], a)
    sb = max(RP["sig_fondo"], 1e-6)
    hb = 0.30
    D.PL([(xr - a0, yd), (xr, yd - hb), (xr, yd)], "E-COTAS", close=True, color=1)
    D.HATCH([(xr - a0, yd), (xr, yd - hb), (xr, yd)], "ANSI31", 0.008, color=1, layer="E-COTAS")
    D.T(xr + 0.05, yd - hb, f"σb={f(sb / 10, 3)} kg/cm²", 0.045, color=1)
    D.T(xr - a0 / 2, yd + 0.04, f"a0={f(a0)}", 0.045, align="C", color=1)
    # Cotas
    D.DIM((xl, yb - 0.10), (xr, yb - 0.10), (0, yb - 0.80))
    D.DIM((xl, yb), (xl, y0), (xl - 0.30, 0), angle=90)
    D.DIM((cx + ap / 2, y0), (cx + ap / 2, yt), (cx + ap / 2 + 0.25, 0), angle=90)
    D.DIM((cx - ap / 2, yt), (cx + ap / 2, yt), (0, yt + 0.22))
    # Referencias
    xt = xr + 1.55
    D.FLECHA((cx + ap / 2 - 0.02, yt - 0.03), (xt - 0.15, yt + 0.40), "PEDESTAL CON CHAFLÁN 5 cm\nY PENDIENTE (ESCURRIMIENTO)", 0.05)
    if not RP["empotrado"]:
        D.FLECHA((cx + RP["Dc"] / 2, yt - 0.20), (xt - 0.15, yt + 0.05),
                 f"CANASTILLA {RP['n_pp']} PERNOS Ø{RP['d_pp']}\nL={f(RP['L_anc'])} m + GROUT 3 cm", 0.05)
    D.FLECHA((xg + 0.017, y0 - 0.20), (xt - 0.15, y0 - 0.25), 'TUBO PVC-SAP Ø1" PUESTA A TIERRA', 0.05)
    D.FLECHA((xr - 0.15, yb + t * 0.45), (xt - 0.15, yb + t * 0.45),
             f"MACIZO C°A° f'c={f(RP['fc'], 0)}\nMALLA Ø{RP['barra']} @ {f(RP['s_maz'])}", 0.05)
    D.FLECHA((cx + ap / 2 - 0.05, y0 + hp * 0.3), (xt - 0.15, y0 + 0.02 - 0.0),
             f"{RP['n_ped']} Ø{RP['barra']} + EST. Ø3/8\" @ 0.15", 0.05)
    D.FLECHA((xr + 0.05, yb - 0.05), (xt - 0.15, yb - 0.15), "SOLADO C:H 1:12 e=0.10", 0.05)
    D.volcar_etiquetas()


def exportar_dxf(ruta, P, RE, RP):
    if ezdxf is None:
        raise RuntimeError("Falta ezdxf (pip install ezdxf)")
    D = Lamina()
    S = _tamanos_lamina(RE, RP)
    items = calcular_metrados(P, RE, RP)
    gap = 0.6
    h = 0.06
    fc = num(P.get("fc"), 210)
    # ---- Contenido de los cuadros ----
    espec = [["ESPECIFICACIÓN", "VALOR"],
             ["Normas", "E.020 · E.030 · E.050 · E.060 · ACI 351.3R-04 · ACI 318-19"],
             ["Concreto bloque/macizo", f"f'c = {int(fc)} kg/cm²"],
             ["Solado", "C:H 1:12  e=0.10 m"],
             ["Acero", "ASTM A615 Gr.60 fy=4200 kg/cm²"],
             ["Recubrimiento", "7.5 cm contra terreno / 5.0 cm expuesto"],
             ["Suelo q adm", f"{f(num(P.get('q_adm')))} kg/cm²  (q din ≤ 50 %)"],
             ["Grout", "No contráctil ASTM C1107 f'c≥350, e=2.5–5 cm"],
             ["Junta perimetral", 'Poliestireno elastif./neopreno e=1" + sello PU'],
             ["Pernos", "ASTM F1554 Gr.36 galv. A153 + tuerca A563"],
             ["Relleno", "Compactado al 95 % Proctor Modificado"]]
    if RE.get("ok"):
        espec.append(["Bancada", f"fn={f(RE['fn'])} Hz  r={f(RE['r'], 2)}  A={f(RE['A_mm'], 4)} mm"])
    if RP.get("ok"):
        espec.append(["Macizo", f"Mv={f(RP['Mv'], 2)} t·m  FS={f(RP['FS'], 2)}  tanα={f(RP['tan_req'], 4)}"])
    pernos = [["ELEM.", "CANT.", "Ø", "TIPO / GRADO", "hef / L"]]
    if RE.get("ok"):
        pernos.append(["Equipo", str(RE["n"]), RE["dp"], "J " + RE["grado"].replace("ASTM ", ""),
                       f"{f(RE['hef'])} m"])
    if RP.get("ok") and not RP["empotrado"] and RP["n_pp"] > 0:
        pernos.append(["Poste", str(RP["n_pp"]), RP["d_pp"], "L F1554 Grado 36", f"{f(RP['L_anc'])} m"])
    if RE.get("ok"):
        pernos.append(["Torque", "", "", f"{f(RE['torque'], 1)} kg·m (equipo)", ""])
    met = [["ÍTEM", "DESCRIPCIÓN", "UND", "METRADO"]]
    for it in items:
        if it[2]:
            met.append([it[0], it[1], it[2], f(it[3], 2)])
    anch_esp, anch_per, anch_met = [2.3, 4.7], [0.9, 0.7, 0.6, 2.6, 1.0], [0.7, 5.6, 0.6, 1.1]
    rh = h * 2.1
    h_esp = rh * 1.2 + rh * len(espec)
    h_per = rh * 1.2 + rh * len(pernos)
    h_met = 0.052 * 2.1 * (1.2 + len(met))
    w_tab = max(sum(anch_esp), sum(anch_per))
    w_met = sum(anch_met)
    W_MEM, H_MEM = 6.5, 3.4

    # ---- Distribución por filas (de abajo hacia arriba) ----
    h1 = max(S["hP1"], S["hS1"], S["hD"]) if RE.get("ok") else 0
    h2 = max(S["hP2"], S["hS2"], h_esp + gap + h_per) if RP.get("ok") else h_esp + gap + h_per
    h3 = max(h_met, H_MEM)
    y3 = 0.0
    y2 = y3 + h3 + gap
    y1 = y2 + h2 + gap
    # Fila 1: bancada
    x = 0.0
    if RE.get("ok"):
        dxf_planta_dado(D, RE, x, y1 + (h1 - S["hP1"]), S["wP1"])
        x += S["wP1"] + gap
        dxf_corte_dado(D, RE, x, y1)
        x += S["wS1"] + gap
        dxf_detalle_perno(D, RE, x, y1, S["k"])
        x += S["wD"]
    W1 = x
    # Fila 2: macizo + cuadros de especificaciones y pernos
    x = 0.0
    if RP.get("ok"):
        dxf_planta_macizo(D, RP, x, y2 + (h2 - S["hP2"]), S["wP2"])
        x += S["wP2"] + gap
        dxf_corte_macizo(D, RP, x, y2)
        x += S["wS2"] + gap
    W2a = x
    # Fila 3: metrados (izquierda) y membrete (derecha)
    W = max(W1, W2a + w_tab, w_met + gap + W_MEM)
    xt = W - w_tab
    D.TABLA(xt, y2 + h2, anch_esp, espec, h, "CUADRO DE ESPECIFICACIONES TÉCNICAS")
    D.TABLA(xt, y2 + h2 - h_esp - gap, anch_per, pernos, h, "CUADRO DE PERNOS DE ANCLAJE")
    D.TABLA(0.0, y3 + h3, anch_met, met, 0.052, "PLANILLA DE METRADOS")
    # Membrete estructural
    xm, ym = W - W_MEM, y3
    D.R(xm, ym, W_MEM, H_MEM, "E-MARCO")
    filas_m = [("PROYECTO:", P.get("nombre", "")[:58]), ("UBICACIÓN:", P.get("ubicacion", "")),
               ("CONTENIDO:", "CIMENTACIÓN ANTIVIBRATORIA DE EQUIPO"), ("", "Y MACIZO DE POSTE (SULZBERGER)"),
               ("ING. RESPONSABLE:", P.get("ingeniero", "")), ("CIP:", P.get("cip", "")),
               ("FECHA:", datetime.date.today().strftime("%d/%m/%Y")), ("ESCALA:", "INDICADA")]
    yy = ym + H_MEM
    for k_, v_ in filas_m:
        yy -= 0.30
        D.L((xm, yy), (xm + 4.6, yy), "E-MARCO")
        D.T(xm + 0.06, yy + 0.10, k_, 0.06)
        D.T(xm + 1.35, yy + 0.10, v_, 0.065)
    D.L((xm + 4.6, ym), (xm + 4.6, ym + H_MEM), "E-MARCO")
    D.T(xm + 5.55, ym + 2.6, "LÁMINA", 0.10, align="C")
    D.T(xm + 5.55, ym + 1.6, "CE-01", 0.32, align="C")
    D.T(xm + 5.55, ym + 0.9, "ESTRUCTURAS", 0.08, align="C")
    D.T(xm + 2.3, ym + 0.35, "DISEÑADOR PRO DE CIMENTACIONES – ACI 351.3R / SULZBERGER", 0.055, align="C")
    # Marco
    Htot = y1 + h1 if RE.get("ok") else y2 + h2
    D.R(-0.4, -0.4, W + 0.8, Htot + 0.8, "E-MARCO")
    D.R(-0.55, -0.55, W + 1.1, Htot + 1.1, "E-MARCO", color=8)
    D.doc.header["$EXTMIN"] = (-0.6, -0.6, 0)
    D.doc.header["$EXTMAX"] = (W + 0.6, Htot + 0.6, 0)
    try:
        D.doc.set_modelspace_vport(height=Htot + 2, center=(W / 2, Htot / 2))
    except Exception:
        pass
    D.doc.saveas(ruta)
    return ruta


# =============================================================================
# 9. EXPORTACIÓN A EXCEL (.XLSX)
# =============================================================================
_AZUL = "0B3D6B"


def _xl_estilos():
    thin = Side(style="thin", color="9AA5B1")
    return dict(
        borde=Border(left=thin, right=thin, top=thin, bottom=thin),
        head=PatternFill("solid", fgColor=_AZUL),
        sec=PatternFill("solid", fgColor="D6E4F0"),
        ok=PatternFill("solid", fgColor="C6EFCE"),
        bad=PatternFill("solid", fgColor="FFC7CE"),
        fh=Font(bold=True, color="FFFFFF"),
        fb=Font(bold=True),
    )


def _xl_tabla(ws, fila, filas, est, encab=None, anchos=None):
    if encab:
        for j, e in enumerate(encab, 1):
            c = ws.cell(row=fila, column=j, value=e)
            c.fill, c.font, c.border = est["head"], est["fh"], est["borde"]
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        fila += 1
    for fl in filas:
        es_sec = len(fl) >= 1 and isinstance(fl[0], str) and fl[0].startswith("§")
        for j, v in enumerate(fl, 1):
            val = v[1:] if (j == 1 and es_sec) else v
            c = ws.cell(row=fila, column=j, value=val)
            c.border = est["borde"]
            if es_sec:
                c.fill, c.font = est["sec"], est["fb"]
            if isinstance(val, str) and val in ("CUMPLE", "OK", "FUERA DE RESONANCIA - CUMPLE"):
                c.fill = est["ok"]
            elif isinstance(val, str) and (val.startswith("NO CUMPLE") or val.startswith("EN RESONANCIA")):
                c.fill = est["bad"]
            if isinstance(val, float):
                c.number_format = "#,##0.000"
        fila += 1
    if anchos:
        for j, w in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return fila


def exportar_excel(ruta, P, RE, RP):
    if openpyxl is None:
        raise RuntimeError("Falta openpyxl (pip install openpyxl)")
    est = _xl_estilos()
    wb = openpyxl.Workbook()
    # ---------------- Hoja 1 ----------------
    ws = wb.active
    ws.title = "1_Calc_Dinamico_Grupo_Electrog"
    ws["A1"] = "CÁLCULO DINÁMICO DE BANCADA DE EQUIPO – ACI 351.3R-04 / NTE E.050 / E.060 / ACI 318-19"
    ws["A1"].font = Font(bold=True, size=13, color=_AZUL)
    ws["A2"] = f"Proyecto: {P.get('nombre', '')}   ·   Ubicación: {P.get('ubicacion', '')}   ·   " \
               f"{P.get('ingeniero', '')} CIP {P.get('cip', '')}"
    R = RE
    if R.get("ok"):
        filas = [
            ["§1. DATOS", "", "", "", ""],
            ["Tipo de equipo", R["tipo"], "", "", ""],
            ["Peso del equipo W_eq", R["Weq"], "t", "", ""],
            ["Velocidad", R["rpm"], "RPM", "f_op = RPM/60", ""],
            ["Bloque L", R["L"], "m", "", ""], ["Bloque B", R["B"], "m", "", ""],
            ["Bloque H", R["H"], "m", "", ""],
            ["Vs", R["Vs"], "m/s", "", ""], ["ν", R["nu"], "", "", ""],
            ["§2. CRITERIO DE MASA", "", "", "", ""],
            ["W_c = L·B·H·γc", R["Wc"], "t", "", ""],
            ["W_c / W_eq", R["ratio"], "", f"≥ {R['ratio_min']}", ok_txt(R["ok_masa"])],
            ["Borde libre mínimo", min(R["borde_L"], R["borde_B"]), "m", "≥ 0.25", ok_txt(R["ok_borde"])],
            ["H mínimo = 0.60 + L/30", R["H_min"], "m", f"H = {R['H']}", ok_txt(R["ok_H"])],
            ["§3. PRESIONES", "", "", "", ""],
            ["W_total", R["Wt"], "t", "", ""],
            ["Excentricidad e", R["e"], "m", f"≤ 0.05L = {R['e_aci']:.3f}", ok_txt(R["ok_e"])],
            ["q estática", R["q_est"] / 10, "kg/cm²", "", ""],
            ["q máx (est + din)", R["q_max"] / 10, "kg/cm²", f"≤ 0.5 q_adm = {R['q_lim'] / 10:.3f}",
             ok_txt(R["ok_q"])],
            ["§4. RIGIDEZ DINÁMICA", "", "", "", ""],
            ["ρ = γs/g", R["rho"], "t·s²/m⁴", "", ""],
            ["G = ρ·Vs²", R["G"], "t/m²", "", ""],
            ["r0 = √(LB/π)", R["r0"], "m", "", ""],
            ["Kz0 = 4·G·r0/(1−ν)", R["Kz0"], "t/m", "Lysmer", ""],
            ["ηz = 1 + 0.6(1−ν)D/r0", R["eta_z"], "", "", ""],
            ["Kz", R["Kz"], "t/m", "", ""],
            ["Cz = Kz/A", R["Cz_kgcm3"], "kg/cm³", "Barkan", ""],
            ["Bz = (1−ν)/4·m/(ρ r0³)", R["Bz"], "", "", ""],
            ["ζz (diseño)", R["zeta_z"], "", "≤ 0.50", ""],
            ["§5. FRECUENCIAS Y RESONANCIA", "", "", "", ""],
            ["fn = (1/2π)√(Kz·g/W)", R["fn"], "Hz", "", ""],
            ["f_op", R["fop"], "Hz", "", ""],
            ["r = f_op/fn", R["r"], "", "∉ [0.70 ; 1.40]",
             "FUERA DE RESONANCIA - CUMPLE" if R["ok_res"] else "EN RESONANCIA - NO CUMPLE"],
            ["fn horizontal", R["fnx"], "Hz", f"r = {R['rx']:.3f}", ok_txt(R["ok_rx"])],
            ["fn cabeceo", R["fn_psi"], "Hz", f"r = {R['r_rel_psi']:.3f}", ok_txt(R["ok_rpsi"])],
            ["M = 1/√((1−r²)²+(2ζr)²)", R["M"], "", "", ""],
            ["§6. AMPLITUD", "", "", "", ""],
            ["F0 dinámica", R["F0"] * 1000, "kg", "", ""],
            ["A máx = F0/Kz·M", R["A_mm"], "mm", f"< {R['A_lim']}", ok_txt(R["ok_amp"])],
            ["Velocidad pico", R["vel"], "mm/s", "", ""],
            ["Transmisibilidad aisladores T", R["T"], "", f"Eficiencia {R['efic']:.1f} %", ""],
            ["§7. ACERO", "", "", "", ""],
            ["As mín (0.0018·b·H)", R["As_tot"], "cm²/m", "", ""],
            ["As por cara", max(R["As_cara"], R["As_flex"]), "cm²/m", "", ""],
            [f"Ø{R['barra']} @ {R['s']:.2f}", R["As_prov"], "cm²/m", "", ok_txt(R["ok_acero"])],
            ["§8. PERNOS DE ANCLAJE", "", "", "", ""],
            [f"{R['n']} pernos Ø{R['dp']} {R['grado']}", R["hef"], "m (hef)", f"≥ {R['hef_min']:.3f}",
             ok_txt(R["ok_hef"])],
            ["Nu por perno", R["Nu_b"], "kg", "", ""], ["Vu por perno", R["Vu_b"], "kg", "", ""],
            ["φNsa", R["phiNsa"], "kg", "", ""], ["φNcb", R["phiNcb"], "kg", "", ""],
            ["φNpn", R["phiNpn"], "kg", "", ""], ["φVsa", R["phiVsa"], "kg", "", ""],
            ["φVcb", R["phiVcb"], "kg", "", ""], ["φVcp", R["phiVcp"], "kg", "", ""],
            ["Interacción N/φN + V/φV", R["inter"], "", "≤ 1.20", ok_txt(R["ok_int"])],
            ["Grout e", R["e_grout"], "cm", "2.5 – 5.0", ok_txt(R["ok_grout"])],
            ["Torque recomendado", R["torque"], "kg·m", "", ""],
        ]
        fin = _xl_tabla(ws, 4, filas, est, ["Parámetro / Ecuación", "Valor", "Unidad", "Límite / Nota",
                                             "Estado"], [38, 18, 12, 30, 30])
        # Curva M vs r con gráfico
        c0 = 7
        ws.cell(row=4, column=c0, value="r").font = est["fb"]
        ws.cell(row=4, column=c0 + 1, value="M").font = est["fb"]
        curva = curva_amplificacion(R["zeta_z"], 3.0, 31)
        for i, (rr, MM) in enumerate(curva):
            ws.cell(row=5 + i, column=c0, value=round(rr, 3))
            ws.cell(row=5 + i, column=c0 + 1, value=round(MM, 4))
        ch = ScatterChart()
        ch.title = "Factor de amplificación dinámica M vs r"
        ch.x_axis.title = "r = f_op / f_n"
        ch.y_axis.title = "M"
        ch.height, ch.width = 8, 15
        xs = Reference(ws, min_col=c0, min_row=5, max_row=5 + len(curva) - 1)
        ys = Reference(ws, min_col=c0 + 1, min_row=4, max_row=5 + len(curva) - 1)
        se = Series(ys, xs, title_from_data=True)
        ch.series.append(se)
        ws.add_chart(ch, "J4")
    ws.freeze_panes = "A5"

    # ---------------- Hoja 2 ----------------
    ws2 = wb.create_sheet("2_Calc_Sulzberger_Postes_Torres")
    ws2["A1"] = "MACIZO DE CIMENTACIÓN PARA POSTE / TORRE – MÉTODO DE SULZBERGER + VIENTO NTE E.020"
    ws2["A1"].font = Font(bold=True, size=13, color=_AZUL)
    R = RP
    if R.get("ok"):
        filas = [
            ["§1. DATOS", "", "", "", ""],
            ["Tipo", R["tipo"], "", "", ""], ["Altura total Hp", R["Hp"], "m", "", ""],
            ["Altura libre", R["Hl"], "m", "", ""], ["Velocidad V", R["V"], "km/h", "E.020", ""],
            ["Macizo a", R["a"], "m", "", ""], ["Macizo b", R["b"], "m", "", ""],
            ["Profundidad t", R["t"], "m", "", ""],
            ["§2. VIENTO", "", "", "", ""],
            ["F fuste", R["F_f"], "kg", "Σ P·d·Δh", ""], ["F equipos", R["F_eq"], "kg", "P·A", ""],
            ["Tiro", R["tiro"], "kg", "", ""], ["Ht", R["Ht"], "t", "", ""],
            ["M base", R["M_base"], "t·m", "", ""], ["M terreno", R["M_suelo"], "t·m", "", ""],
            ["Mv = M + Ht·2t/3", R["Mv"], "t·m", "", ""],
            ["§3. SULZBERGER", "", "", "", ""],
            ["P vertical", R["Pv"], "t", "", ""],
            ["Ct = C2·t/2", R["Ct"], "t/m³", "", ""], ["Cb", R["Cb"], "t/m³", "", ""],
            ["tan α adm", R["tan_adm"], "", "", ""],
            ["M1 = b t³/36·Ct·tanα", R["M1"], "t·m", "", ""],
            ["M2 = P a (0.5 − 2/3 √(P/(2 b a² Cb tanα)))", R["M2"], "t·m", "", ""],
            ["FS = (M1+M2)/Mv", R["FS"], "", "≥ 1.50", ok_txt(R["ok_FS"])],
            ["Coef. s (M1/M2)", R["s_sulz"], "", f"M1/M2 = {R['rel']:.3f}", ok_txt(R["ok_s"])],
            ["tan α requerido", R["tan_req"], "", f"≤ {R['tan_adm']}", ok_txt(R["ok_tan"])],
            ["σ lateral máx", R["sig_lat"] / 10, "kg/cm²", "", ""],
            ["σ fondo máx", R["sig_fondo"] / 10, "kg/cm²", "", ""],
            ["σ medio", R["sig_med"] / 10, "kg/cm²", f"≤ {R['qadm']}", ok_txt(R["ok_q"])],
            ["§4. PEDESTAL Y ANCLAJE", "", "", "", ""],
            ["Mu pedestal", R["Mu_ped"], "t·m", "", ""],
            [f"{R['n_ped']} Ø{R['barra']}", R["n_cara"] * R["ab"], "cm²/cara", f"≥ {R['As_ped']:.2f}",
             ok_txt(R["ok_ped"])],
            ["Pernos canastilla (interacción)", R["inter_p"], "", "≤ 1.20", ok_txt(R["ok_pp"])],
        ]
        fin = _xl_tabla(ws2, 3, filas, est, ["Parámetro / Ecuación", "Valor", "Unidad", "Límite / Nota",
                                              "Estado"], [42, 18, 12, 26, 18])
        ws2.cell(row=3, column=7, value="TRAMOS DE VIENTO SOBRE EL FUSTE").font = est["fb"]
        filas_t = [[tr["i"], round(tr["zg"], 3), round(tr["d"], 3), round(tr["Vh"], 2), round(tr["Ph"], 2),
                    round(tr["F"], 3), round(tr["M"], 3)] for tr in R["tramos"]]
        r0 = 4
        for j, e in enumerate(["Tramo", "h (m)", "d (m)", "Vh (km/h)", "Ph (kg/m²)", "F (kg)", "M (kg·m)"]):
            c = ws2.cell(row=r0, column=7 + j, value=e)
            c.fill, c.font, c.border = est["head"], est["fh"], est["borde"]
        for i, fl in enumerate(filas_t):
            for j, v in enumerate(fl):
                c = ws2.cell(row=r0 + 1 + i, column=7 + j, value=v)
                c.border = est["borde"]
        n = len(filas_t)
        ws2.cell(row=r0 + 1 + n, column=7, value="TOTAL").font = est["fb"]
        ws2.cell(row=r0 + 1 + n, column=12, value=f"=SUM(L{r0 + 1}:L{r0 + n})").font = est["fb"]
        ws2.cell(row=r0 + 1 + n, column=13, value=f"=SUM(M{r0 + 1}:M{r0 + n})").font = est["fb"]
        for j in range(7, 14):
            ws2.column_dimensions[get_column_letter(j)].width = 12

    # ---------------- Hoja 3 ----------------
    ws3 = wb.create_sheet("3_Metrados_y_Presupuesto")
    ws3["A1"] = "PLANILLA DE METRADOS Y PRESUPUESTO REFERENCIAL (S/.)"
    ws3["A1"].font = Font(bold=True, size=13, color=_AZUL)
    ws3["A2"] = "Edite los precios unitarios (columna E): los parciales y el total se recalculan."
    items = calcular_metrados(P, RE, RP)
    fila = 4
    for j, e in enumerate(["Ítem", "Descripción", "Und", "Metrado", "P.U. (S/.)", "Parcial (S/.)"], 1):
        c = ws3.cell(row=fila, column=j, value=e)
        c.fill, c.font, c.border = est["head"], est["fh"], est["borde"]
    fila += 1
    f_ini = fila
    for it in items:
        if it[2] == "":
            for j, v in enumerate([it[0], it[1]], 1):
                c = ws3.cell(row=fila, column=j, value=v)
                c.fill, c.font = est["sec"], est["fb"]
            for j in range(1, 7):
                ws3.cell(row=fila, column=j).border = est["borde"]
        else:
            vals = [it[0], it[1], it[2], round(it[3], 3), round(it[4], 2), f"=D{fila}*E{fila}"]
            for j, v in enumerate(vals, 1):
                c = ws3.cell(row=fila, column=j, value=v)
                c.border = est["borde"]
                if j >= 4:
                    c.number_format = "#,##0.00"
        fila += 1
    ws3.cell(row=fila, column=2, value="COSTO DIRECTO").font = est["fb"]
    c = ws3.cell(row=fila, column=6, value=f"=SUM(F{f_ini}:F{fila - 1})")
    c.font, c.number_format = est["fb"], "#,##0.00"
    ws3.cell(row=fila + 1, column=2, value="Gastos generales (10 %)")
    ws3.cell(row=fila + 1, column=6, value=f"=F{fila}*0.10").number_format = "#,##0.00"
    ws3.cell(row=fila + 2, column=2, value="Utilidad (8 %)")
    ws3.cell(row=fila + 2, column=6, value=f"=F{fila}*0.08").number_format = "#,##0.00"
    ws3.cell(row=fila + 3, column=2, value="Sub total")
    ws3.cell(row=fila + 3, column=6, value=f"=SUM(F{fila}:F{fila + 2})").number_format = "#,##0.00"
    ws3.cell(row=fila + 4, column=2, value="IGV (18 %)")
    ws3.cell(row=fila + 4, column=6, value=f"=F{fila + 3}*0.18").number_format = "#,##0.00"
    c = ws3.cell(row=fila + 5, column=2, value="PRESUPUESTO TOTAL (S/.)")
    c.font = est["fb"]
    c = ws3.cell(row=fila + 5, column=6, value=f"=F{fila + 3}+F{fila + 4}")
    c.font, c.number_format = est["fb"], "#,##0.00"
    for j, w in enumerate([8, 62, 7, 12, 12, 15], 1):
        ws3.column_dimensions[get_column_letter(j)].width = w
    wb.save(ruta)
    return ruta


# =============================================================================
# 10. PROCESO POR LOTES DESDE EXCEL (pandas)
# =============================================================================
def crear_plantilla_lote(ruta):
    if pd is None:
        raise RuntimeError("Falta pandas (pip install pandas)")
    e1 = datos_equipo_ejemplo()
    e2 = dict(e1, tipo_eq="Compresor Industrial", potencia="75 HP", W_eq="1850", L_eq="2.10", B_eq="0.90",
              L="3.00", B="1.60", H="0.90", rpm="1200 RPM - 60Hz", n_pernos="4")
    p1 = datos_poste_ejemplo()
    p2 = datos_poste_cac()
    with pd.ExcelWriter(ruta, engine="openpyxl") as xw:
        pd.DataFrame([dict(id="GE-01", **e1), dict(id="CP-01", **e2)]).to_excel(xw, sheet_name="EQUIPOS", index=False)
        pd.DataFrame([dict(id="PR-01", **p1), dict(id="PC-01", **p2)]).to_excel(xw, sheet_name="POSTES", index=False)
    return ruta


def procesar_lote(ruta_in, ruta_out, P):
    if pd is None:
        raise RuntimeError("Falta pandas (pip install pandas)")
    hojas = pd.read_excel(ruta_in, sheet_name=None, dtype=str)
    res_eq, res_po = [], []
    for nombre, df in hojas.items():
        nom = nombre.upper()
        for _, fila in df.iterrows():
            d = {k: v for k, v in fila.to_dict().items() if isinstance(v, str) and v.strip() != ""}
            if nom.startswith("EQUIP"):
                E = datos_equipo_ejemplo()
                E.update(d)
                R = calcular_equipo(P, E)
                if R.get("ok"):
                    res_eq.append({"ID": d.get("id", ""), "Tipo": R["tipo"], "W_eq (kg)": R["W_kg"],
                                   "LxBxH (m)": f"{R['L']}x{R['B']}x{R['H']}", "Wc/Weq": round(R["ratio"], 2),
                                   "fn (Hz)": round(R["fn"], 2), "fop (Hz)": round(R["fop"], 2),
                                   "r": round(R["r"], 3), "A (mm)": round(R["A_mm"], 5),
                                   "q máx (kg/cm²)": round(R["q_max"] / 10, 3),
                                   "Acero": f"Ø{R['barra']}@{R['s']:.2f}",
                                   "Estado": "CUMPLE" if R["cumple"] else "NO CUMPLE"})
            elif nom.startswith("POST"):
                S = datos_poste_ejemplo()
                S.update(d)
                R = calcular_poste(P, S)
                if R.get("ok"):
                    res_po.append({"ID": d.get("id", ""), "Tipo": R["tipo"], "H (m)": R["Hp"],
                                   "axbxt (m)": f"{R['a']}x{R['b']}x{R['t']}", "Mv (t·m)": round(R["Mv"], 3),
                                   "M1 (t·m)": round(R["M1"], 3), "M2 (t·m)": round(R["M2"], 3),
                                   "FS": round(R["FS"], 3), "tan α": round(R["tan_req"], 5),
                                   "Estado": "CUMPLE" if R["cumple"] else "NO CUMPLE"})
    with pd.ExcelWriter(ruta_out, engine="openpyxl") as xw:
        pd.DataFrame(res_eq).to_excel(xw, sheet_name="Resultados_Equipos", index=False)
        pd.DataFrame(res_po).to_excel(xw, sheet_name="Resultados_Postes", index=False)
    try:
        wb = openpyxl.load_workbook(ruta_out)
        est = _xl_estilos()
        for ws in wb.worksheets:
            for c in ws[1]:
                c.fill, c.font = est["head"], est["fh"]
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 16
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    if c.value == "CUMPLE":
                        c.fill = est["ok"]
                    elif c.value == "NO CUMPLE":
                        c.fill = est["bad"]
        wb.save(ruta_out)
    except Exception:
        pass
    return ruta_out, len(res_eq), len(res_po)


# =============================================================================
# 11. INTERFAZ GRÁFICA (CustomTkinter)
# =============================================================================
CAMPOS_PROY = [
    ("nombre", "Proyecto / Edificación", "entry", None, 300),
    ("ubicacion", "Ubicación", "entry", None, 170),
    ("q_adm", "q adm (kg/cm²)", "entry", None, 70),
    ("gamma_s", "γs (t/m³)", "entry", None, 70),
    ("Ct", "Ct a 2 m (kg/cm³)", "entry", None, 70),
    ("fc", "f'c (kg/cm²)", "combo", ["210", "280"], 80),
    ("ingeniero", "Ingeniero Civil Responsable", "entry", None, 220),
    ("cip", "CIP", "entry", None, 90),
]

CAMPOS_EQ = [
    ("DATOS DEL EQUIPO", [
        ("tipo_eq", "Tipo de equipo", "combo", list(TIPOS_EQUIPO)),
        ("potencia", "Potencia (kW / kVA)", "entry", None),
        ("W_eq", "Peso total con fluidos W_eq (kg)", "entry", None),
        ("rpm", "Velocidad de operación", "combo", VELOCIDADES),
        ("L_eq", "Largo del chasis L_eq (m)", "entry", None),
        ("B_eq", "Ancho del chasis W_eq (m)", "entry", None),
        ("h_cg", "Altura CG del equipo (m)", "entry", None),
        ("e_cg", "Excentricidad CG equipo (m)", "entry", None),
    ]),
    ("BLOQUE DE CIMENTACIÓN Y AISLAMIENTO", [
        ("L", "Largo del bloque L (m)", "entry", None),
        ("B", "Ancho del bloque B (m)", "entry", None),
        ("H", "Altura del bloque H (m)", "entry", None),
        ("h_sobre", "Altura sobre NPT (m)", "entry", None),
        ("aislamiento", "Tipo de aislamiento", "combo", AISLAMIENTOS),
        ("defl_mm", "Deflexión estática aislador (mm)", "entry", None),
        ("barra", "Barra de la parrilla", "combo", list(BARRAS)),
    ]),
    ("SUELO DINÁMICO Y EXCITACIÓN", [
        ("Vs", "Velocidad de onda de corte Vs (m/s)", "entry", None),
        ("nu", "Coeficiente de Poisson ν", "entry", None),
        ("pct_rotor", "Masa rotante (% de W_eq)", "entry", None),
        ("G_bal", "Grado de balanceo ISO 1940 G (mm/s)", "entry", None),
        ("Sf", "Factor de servicio S_f", "entry", None),
        ("A_lim", "Amplitud admisible (mm)", "entry", None),
    ]),
    ("PERNOS DE ANCLAJE Y GROUT", [
        ("n_pernos", "Número de pernos", "entry", None),
        ("d_perno", "Diámetro del perno", "combo", list(PERNOS)),
        ("grado", "Grado del perno", "combo", list(GRADOS_PERNO)),
        ("h_ef", "Empotramiento h_ef (m)", "entry", None),
        ("e_grout", "Espesor de grout (cm)", "entry", None),
        ("Ch", "Coef. sísmico horizontal C_h", "entry", None),
    ]),
]

CAMPOS_POSTE = [
    ("DATOS DEL POSTE / ESTRUCTURA", [
        ("tipo_poste", "Tipo de estructura", "combo", TIPOS_POSTE),
        ("Hp", "Altura total H_p (m)", "entry", None),
        ("H_libre", "Altura libre H_libre (m)", "entry", None),
        ("d_top", "Diámetro / ancho en punta (m)", "entry", None),
        ("d_base", "Diámetro / ancho en base (m)", "entry", None),
        ("W_poste", "Peso del poste (kg)", "entry", None),
        ("C_f", "Factor de forma fuste C (E.020)", "entry", None),
    ]),
    ("VIENTO Y CARGAS EN PUNTA", [
        ("V", "Velocidad de viento de zona V (km/h)", "entry", None),
        ("A_eq", "Área luminarias / equipos en punta (m²)", "entry", None),
        ("C_eq", "Factor de forma equipos C", "entry", None),
        ("W_eqp", "Peso de equipos en punta (kg)", "entry", None),
        ("tiro", "Tiro en la punta (kg)", "entry", None),
    ]),
    ("MACIZO Y PEDESTAL", [
        ("a", "Macizo a (dirección del viento, m)", "entry", None),
        ("b", "Macizo b (m)", "entry", None),
        ("t", "Profundidad t (m)", "entry", None),
        ("a_ped", "Lado del pedestal (m)", "entry", None),
        ("h_ped", "Altura del pedestal sobre NTN (m)", "entry", None),
        ("barra_p", "Barra del macizo/pedestal", "combo", list(BARRAS)),
    ]),
    ("SULZBERGER Y ANCLAJE", [
        ("Cb_fac", "Relación Cb / Ct", "entry", None),
        ("tan_adm", "Giro admisible tan α", "entry", None),
        ("n_pp", "N° pernos canastilla (0 = empotrado)", "entry", None),
        ("d_pp", "Diámetro pernos canastilla", "combo", list(PERNOS)),
        ("Dc", "Diámetro círculo de pernos Dc (m)", "entry", None),
    ]),
]


def iniciar_gui():
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title(APP_TITULO)
            self.geometry("1280x820")
            self.minsize(1100, 700)
            self.vars = {}
            self._job = None
            self.RE, self.RP = {}, {}
            self._construir()
            self._cargar(datos_proyecto_ejemplo())
            self._cargar(datos_equipo_ejemplo())
            self._cargar(datos_poste_ejemplo())
            self.recalcular()

        # ------------------------------------------------------------------
        def _var(self, key):
            v = ctk.StringVar()
            v.trace_add("write", lambda *a: self._programar())
            self.vars[key] = v
            return v

        def _construir(self):
            # Cabecera
            top = ctk.CTkFrame(self, corner_radius=8)
            top.pack(fill="x", padx=8, pady=(8, 4))
            ctk.CTkLabel(top, text="DISEÑADOR PRO DE CIMENTACIONES · ACI 351.3R + SULZBERGER",
                         font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=6,
                                                                        sticky="w", padx=8, pady=(6, 2))
            self.modo = ctk.CTkSegmentedButton(top, values=["Oscuro", "Claro"],
                                               command=lambda v: ctk.set_appearance_mode(
                                                   "dark" if v == "Oscuro" else "light"))
            self.modo.set("Oscuro")
            self.modo.grid(row=0, column=6, columnspan=2, sticky="e", padx=8)
            for i, (k, lab, tp, ops, w) in enumerate(CAMPOS_PROY):
                r, c = 1 + i // 4, (i % 4) * 2
                ctk.CTkLabel(top, text=lab, anchor="e").grid(row=r, column=c, sticky="e", padx=(8, 4), pady=3)
                v = self._var(k)
                if tp == "combo":
                    wdg = ctk.CTkComboBox(top, values=ops, variable=v, width=w)
                else:
                    wdg = ctk.CTkEntry(top, textvariable=v, width=w)
                wdg.grid(row=r, column=c + 1, sticky="w", padx=(0, 8), pady=3)

            # Pestañas
            self.tabs = ctk.CTkTabview(self, corner_radius=8)
            self.tabs.pack(fill="both", expand=True, padx=8, pady=4)
            t1 = self.tabs.add("1. Equipo Vibratorio (ACI 351.3R)")
            t2 = self.tabs.add("2. Macizo Poste/Torre (Sulzberger)")
            t3 = self.tabs.add("3. Resultados y Semáforos")
            self._pestana_campos(t1, CAMPOS_EQ, "Cargar Ejemplo Grupo Electrógeno (250 kVA - 2,450 kg - 1800 RPM)",
                                 lambda: self._cargar(datos_equipo_ejemplo()), "eq")
            self._pestana_campos(t2, CAMPOS_POSTE,
                                 "Cargar Ejemplo Poste Reflector de Estadio / Losa Deportiva (15m de Altura + 6 Reflectores)",
                                 lambda: self._cargar(datos_poste_ejemplo()), "po")
            self._pestana_resultados(t3)

            # Barra inferior
            bot = ctk.CTkFrame(self, corner_radius=8)
            bot.pack(fill="x", padx=8, pady=(4, 8))
            bot.grid_columnconfigure((0, 1, 2), weight=1)
            ctk.CTkButton(bot, height=46, fg_color="#1f6aa5", hover_color="#144870",
                          text="1. Generar Memoria de Cálculo Dinámica (ACI 351.3R) y de Postes\n(Sulzberger) en Word (.DOCX)",
                          command=self.exp_word).grid(row=0, column=0, sticky="ew", padx=6, pady=6)
            ctk.CTkButton(bot, height=46, fg_color="#d9731a", hover_color="#a35512",
                          text="2. Dibujar Lámina Estructural: Dado Antivibratorio + Pernos J +\nMacizo de Poste en AutoCAD (.DXF)",
                          command=self.exp_dxf).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
            ctk.CTkButton(bot, height=46, fg_color="#2e8b57", hover_color="#1f5f3b",
                          text="3. Exportar Hojas de Cálculo Dinámico, Sulzberger y\nPlanilla de Metrados en Excel (.XLSX)",
                          command=self.exp_excel).grid(row=0, column=2, sticky="ew", padx=6, pady=6)

        def _pestana_campos(self, tab, grupos, txt_btn, cmd, sufijo):
            barra = ctk.CTkFrame(tab, fg_color="transparent")
            barra.pack(fill="x", pady=(2, 4))
            ctk.CTkButton(barra, text=txt_btn, command=cmd, width=520).pack(side="left", padx=4)
            if sufijo == "po":
                ctk.CTkButton(barra, text="Ejemplo Poste CAC 13m/400daN", width=220,
                              command=lambda: self._cargar(datos_poste_cac())).pack(side="left", padx=4)
            lbl = ctk.CTkLabel(barra, text="", font=ctk.CTkFont(size=14, weight="bold"))
            lbl.pack(side="right", padx=10)
            setattr(self, "estado_" + sufijo, lbl)
            sc = ctk.CTkScrollableFrame(tab)
            sc.pack(fill="both", expand=True)
            sc.grid_columnconfigure((0, 1), weight=1)
            for gi, (titulo, campos) in enumerate(grupos):
                fr = ctk.CTkFrame(sc)
                fr.grid(row=gi // 2, column=gi % 2, sticky="nsew", padx=6, pady=6)
                ctk.CTkLabel(fr, text=titulo, font=ctk.CTkFont(size=13, weight="bold")).grid(
                    row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(6, 4))
                for i, (k, lab, tp, ops) in enumerate(campos, 1):
                    ctk.CTkLabel(fr, text=lab, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=2)
                    v = self._var(k)
                    if tp == "combo":
                        w = 380 if k == "aislamiento" else 260
                        wdg = ctk.CTkComboBox(fr, values=ops, variable=v, width=w)
                    else:
                        wdg = ctk.CTkEntry(fr, textvariable=v, width=160)
                    wdg.grid(row=i, column=1, sticky="w", padx=8, pady=2)

        def _pestana_resultados(self, tab):
            barra = ctk.CTkFrame(tab, fg_color="transparent")
            barra.pack(fill="x", pady=(2, 4))
            ctk.CTkButton(barra, text="Recalcular", width=120, command=self.recalcular).pack(side="left", padx=4)
            ctk.CTkButton(barra, text="Crear plantilla de lote (Excel)", width=220,
                          command=self.plantilla_lote).pack(side="left", padx=4)
            ctk.CTkButton(barra, text="Importar lote de equipos/postes desde Excel", width=300,
                          command=self.importar_lote).pack(side="left", padx=4)
            cont = ctk.CTkFrame(tab, fg_color="transparent")
            cont.pack(fill="both", expand=True)
            cont.grid_columnconfigure((0, 1), weight=1)
            cont.grid_rowconfigure(0, weight=1)
            fuente = ctk.CTkFont(family="Consolas", size=13)
            self.txtA = ctk.CTkTextbox(cont, font=fuente, wrap="none")
            self.txtA.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
            self.txtB = ctk.CTkTextbox(cont, font=fuente, wrap="none")
            self.txtB.grid(row=0, column=1, sticky="nsew", padx=4, pady=4)
            for tx in (self.txtA, self.txtB):
                tx.tag_config("ok", foreground="#2ecc71")
                tx.tag_config("bad", foreground="#ff5c5c")
                tx.tag_config("tit", foreground="#4fa3ff")
                tx.tag_config("info", foreground="#f1c40f")

        # ------------------------------------------------------------------
        def _cargar(self, dic):
            for k, v in dic.items():
                if k in self.vars:
                    self.vars[k].set(v)

        def _datos(self, claves):
            return {k: self.vars[k].get() for k in claves if k in self.vars}

        def datos(self):
            P = self._datos([c[0] for c in CAMPOS_PROY])
            E = self._datos([c[0] for g in CAMPOS_EQ for c in g[1]])
            S = self._datos([c[0] for g in CAMPOS_POSTE for c in g[1]])
            P["_potencia"] = E.get("potencia", "")
            return P, E, S

        def _programar(self):
            if self._job:
                self.after_cancel(self._job)
            self._job = self.after(350, self.recalcular)

        def recalcular(self):
            self._job = None
            P, E, S = self.datos()
            try:
                self.RE = calcular_equipo(P, E)
            except Exception as ex:
                self.RE = {"ok": False, "errores": [str(ex)]}
            try:
                self.RP = calcular_poste(P, S)
            except Exception as ex:
                self.RP = {"ok": False, "errores": [str(ex)]}
            self._mostrar()

        def _linea(self, tx, nombre, valor, estado=None):
            tx.insert("end", f"  {nombre:<34}{valor:>16}  ")
            if estado is None:
                tx.insert("end", "\n")
            else:
                tx.insert("end", ("● CUMPLE" if estado else "● NO CUMPLE") + "\n", "ok" if estado else "bad")

        def _mostrar(self):
            A, B = self.txtA, self.txtB
            for tx in (A, B):
                tx.configure(state="normal")
                tx.delete("1.0", "end")
            R = self.RE
            A.insert("end", " MÓDULO A · BANCADA ANTIVIBRATORIA (ACI 351.3R)\n", "tit")
            A.insert("end", " " + "─" * 66 + "\n", "tit")
            if R.get("ok"):
                self._linea(A, "Peso bloque Wc (t)", f(R["Wc"], 3))
                self._linea(A, "Relación de masas Wc/Weq", f"{f(R['ratio'])} ≥ {f(R['ratio_min'], 1)}", R["ok_masa"])
                self._linea(A, "Borde libre mínimo (m)", f(min(R["borde_L"], R["borde_B"]), 3), R["ok_borde"])
                self._linea(A, "Espesor H ≥ 0.60+L/30 (m)", f"{f(R['H'])}≥{f(R['H_min'])}", R["ok_H"])
                self._linea(A, "Excentricidad e (m)", f(R["e"], 4), R["ok_e"])
                self._linea(A, "q máx / 0.5·q adm (kg/cm²)", f"{f(R['q_max'] / 10, 3)}/{f(R['q_lim'] / 10, 3)}",
                            R["ok_q"])
                A.insert("end", "\n  DINÁMICA DEL SISTEMA BLOQUE–SUELO\n", "info")
                self._linea(A, "G dinámico (kg/cm²)", f(R["G"] / 10, 1))
                self._linea(A, "Rigidez Kz (t/m)", f(R["Kz"], 0))
                self._linea(A, "Balasto dinámico Cz (kg/cm³)", f(R["Cz_kgcm3"], 2))
                self._linea(A, "Amortiguamiento ζz", f(R["zeta_z"], 3))
                self._linea(A, "Frecuencia natural fn (Hz)", f(R["fn"]))
                self._linea(A, "Frecuencia operación fop (Hz)", f(R["fop"]))
                A.insert("end", f"  {'Relación r = fop/fn':<34}{f(R['r'], 3):>16}  ")
                A.insert("end", "● FUERA DE RESONANCIA - CUMPLE\n" if R["ok_res"] else "● EN RESONANCIA - NO CUMPLE\n",
                         "ok" if R["ok_res"] else "bad")
                self._linea(A, "f horizontal / cabeceo (Hz)", f"{f(R['fnx'], 1)}/{f(R['fn_psi'], 1)}",
                            R["ok_rx"] and R["ok_rpsi"])
                self._linea(A, "Factor de amplificación M", f(R["M"], 4))
                self._linea(A, "Fuerza dinámica F0 (kg)", f(R["F0"] * 1000, 2))
                self._linea(A, f"Amplitud A (mm) < {f(R['A_lim'], 3)}", f(R["A_mm"], 5), R["ok_amp"])
                self._linea(A, "Eficiencia de aislamiento (%)", f(R["efic"], 1))
                A.insert("end", "\n  REFUERZO Y ANCLAJES\n", "info")
                self._linea(A, f"Parrilla Ø{R['barra']} @ {f(R['s'])}", f"{f(R['As_prov'], 2)} cm²/m", R["ok_acero"])
                self._linea(A, f"{R['n']} pernos J Ø{R['dp']} – interacción", f(R["inter"], 3), R["ok_int"])
                self._linea(A, "Empotramiento hef (cm)", f"{f(R['hef'] * 100, 0)}≥{f(R['hef_min'] * 100, 0)}",
                            R["ok_hef"])
                self._linea(A, "Grout (cm)", f(R["e_grout"], 1), R["ok_grout"])
                self._linea(A, "Torque de pernos (kg·m)", f(R["torque"], 1))
                A.insert("end", "\n  RESULTADO GLOBAL: ", "tit")
                A.insert("end", ("CUMPLE ✔" if R["cumple"] else f"REVISAR ✘ ({R['n_ok']}/{R['n_chk']})") + "\n",
                         "ok" if R["cumple"] else "bad")
                self.estado_eq.configure(text=("● CUMPLE" if R["cumple"] else "● REVISAR") +
                                         f"   fn={f(R['fn'], 1)} Hz  r={f(R['r'], 2)}",
                                         text_color="#2ecc71" if R["cumple"] else "#ff5c5c")
            else:
                A.insert("end", "  Datos incompletos: " + "; ".join(R.get("errores", [])) + "\n", "bad")
                self.estado_eq.configure(text="● DATOS INVÁLIDOS", text_color="#ff5c5c")
            R = self.RP
            B.insert("end", " MÓDULO B · MACIZO DE POSTE / TORRE (SULZBERGER)\n", "tit")
            B.insert("end", " " + "─" * 66 + "\n", "tit")
            if R.get("ok"):
                self._linea(B, "Velocidad en la punta Vh (km/h)", f(R["Vh_top"]))
                self._linea(B, "Presión en equipos Ph (kg/m²)", f(R["P_eq"]))
                self._linea(B, "Fuerza en fuste (kg)", f(R["F_f"], 1))
                self._linea(B, "Fuerza en equipos (kg)", f(R["F_eq"], 1))
                self._linea(B, "Cortante total Ht (kg)", f(R["Ht"] * 1000, 1))
                self._linea(B, "Momento en la base (t·m)", f(R["M_base"], 3))
                self._linea(B, "Momento de vuelco Mv (t·m)", f(R["Mv"], 3))
                B.insert("end", "\n  SULZBERGER (tan α = %s)\n" % f(R["tan_adm"], 3), "info")
                self._linea(B, "Peso vertical P (t)", f(R["Pv"], 3))
                self._linea(B, "Ct / Cb (kg/cm³)", f"{f(R['Ct'] / 1000, 2)}/{f(R['Cb'] / 1000, 2)}")
                self._linea(B, "M1 lateral (t·m)", f(R["M1"], 3))
                self._linea(B, "M2 fondo (t·m)", f(R["M2"], 3))
                self._linea(B, "FS = (M1+M2)/Mv ≥ 1.50", f(R["FS"], 3), R["ok_FS"])
                self._linea(B, "Coeficiente s de Sulzberger", f(R["s_sulz"], 3), R["ok_s"])
                self._linea(B, "Giro tan α real", f(R["tan_req"], 5), R["ok_tan"])
                self._linea(B, "σ lateral máx (kg/cm²)", f(R["sig_lat"] / 10, 3))
                self._linea(B, "σ medio fondo (kg/cm²)", f(R["sig_med"] / 10, 3), R["ok_q"])
                B.insert("end", "\n  PEDESTAL Y ANCLAJE\n", "info")
                self._linea(B, f"Pedestal {R['n_ped']} Ø{R['barra']}", f"As={f(R['As_ped'], 2)}", R["ok_ped"])
                self._linea(B, f"Malla macizo Ø{R['barra']} @", f(R["s_maz"]))
                if R["n_pp"] > 0 and not R["empotrado"]:
                    self._linea(B, f"{R['n_pp']} pernos Ø{R['d_pp']} – interacción", f(R["inter_p"], 3), R["ok_pp"])
                else:
                    self._linea(B, "Empotramiento del poste (m)", f"{f(R['e_emp'])}≥{f(R['e_min'])}", R["ok_emp"])
                B.insert("end", "\n  RESULTADO GLOBAL: ", "tit")
                B.insert("end", ("CUMPLE ✔" if R["cumple"] else f"REVISAR ✘ ({R['n_ok']}/{R['n_chk']})") + "\n",
                         "ok" if R["cumple"] else "bad")
                self.estado_po.configure(text=("● CUMPLE" if R["cumple"] else "● REVISAR") +
                                         f"   FS={f(R['FS'], 2)}  Mv={f(R['Mv'], 2)} t·m",
                                         text_color="#2ecc71" if R["cumple"] else "#ff5c5c")
            else:
                B.insert("end", "  Datos incompletos: " + "; ".join(R.get("errores", [])) + "\n", "bad")
                self.estado_po.configure(text="● DATOS INVÁLIDOS", text_color="#ff5c5c")
            for tx in (A, B):
                tx.configure(state="disabled")

        # ------------------------------------------------------------------
        def _ruta(self, ext, base):
            nombre = re.sub(r"[^\w\-]+", "_", self.vars["nombre"].get())[:40] or "Proyecto"
            return filedialog.asksaveasfilename(
                defaultextension=ext, initialfile=f"{base}_{nombre}{ext}",
                filetypes=[(ext.upper()[1:], "*" + ext)])

        def _exportar(self, fn, ext, base):
            self.recalcular()
            if not (self.RE.get("ok") or self.RP.get("ok")):
                messagebox.showerror("Datos", "Revise los datos de entrada.")
                return
            ruta = self._ruta(ext, base)
            if not ruta:
                return
            try:
                P, _, _ = self.datos()
                fn(ruta, P, self.RE, self.RP)
                abrir_archivo(ruta)
            except PermissionError:
                messagebox.showerror("Error", "El archivo está abierto en otro programa. Ciérrelo e intente de nuevo.")
            except Exception as ex:
                messagebox.showerror("Error", f"{ex}\n\n{traceback.format_exc()[-800:]}")

        def exp_word(self):
            self._exportar(exportar_word, ".docx", "Memoria_Calculo")

        def exp_dxf(self):
            self._exportar(exportar_dxf, ".dxf", "CE-01_Lamina")

        def exp_excel(self):
            self._exportar(exportar_excel, ".xlsx", "Hojas_Calculo")

        def plantilla_lote(self):
            ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Plantilla_Lote.xlsx",
                                                filetypes=[("Excel", "*.xlsx")])
            if ruta:
                try:
                    crear_plantilla_lote(ruta)
                    abrir_archivo(ruta)
                except Exception as ex:
                    messagebox.showerror("Error", str(ex))

        def importar_lote(self):
            ruta_in = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
            if not ruta_in:
                return
            ruta_out = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    initialfile="Resultados_Lote.xlsx",
                                                    filetypes=[("Excel", "*.xlsx")])
            if not ruta_out:
                return
            try:
                P, _, _ = self.datos()
                _, ne, npst = procesar_lote(ruta_in, ruta_out, P)
                messagebox.showinfo("Lote", f"Procesados {ne} equipos y {npst} postes.")
                abrir_archivo(ruta_out)
            except Exception as ex:
                messagebox.showerror("Error", str(ex))

    App().mainloop()


# =============================================================================
# 12. PUNTO DE ENTRADA
# =============================================================================
def demo_sin_gui(carpeta="."):
    """Genera los 3 entregables con los ejemplos (útil para pruebas: --demo)."""
    P = datos_proyecto_ejemplo()
    E = datos_equipo_ejemplo()
    P["_potencia"] = E["potencia"]
    RE = calcular_equipo(P, E)
    RP = calcular_poste(P, datos_poste_ejemplo())
    os.makedirs(carpeta, exist_ok=True)
    rutas = [exportar_word(os.path.join(carpeta, "Memoria_Calculo_Ejemplo.docx"), P, RE, RP),
             exportar_dxf(os.path.join(carpeta, "CE-01_Lamina_Ejemplo.dxf"), P, RE, RP),
             exportar_excel(os.path.join(carpeta, "Hojas_Calculo_Ejemplo.xlsx"), P, RE, RP)]
    return RE, RP, rutas


if __name__ == "__main__":
    if "--demo" in sys.argv:
        i = sys.argv.index("--demo")
        carpeta = sys.argv[i + 1] if len(sys.argv) > i + 1 else "salida_demo"
        RE, RP, rutas = demo_sin_gui(carpeta)
        print(f"Bancada: fn={RE['fn']:.2f} Hz r={RE['r']:.3f} A={RE['A_mm']:.5f} mm -> "
              f"{'CUMPLE' if RE['cumple'] else 'REVISAR'}")
        print(f"Macizo : Mv={RP['Mv']:.3f} t·m FS={RP['FS']:.3f} tanα={RP['tan_req']:.5f} -> "
              f"{'CUMPLE' if RP['cumple'] else 'REVISAR'}")
        for r in rutas:
            print("Generado:", r)
    else:
        iniciar_gui()

# =============================================================================
# INSTALACIÓN Y EMPAQUETADO
# -----------------------------------------------------------------------------
#   pip install customtkinter openpyxl pandas ezdxf python-docx
#   python app_cimentacion_equipos_postes_pro.py
#   python app_cimentacion_equipos_postes_pro.py --demo salida_demo   (sin interfaz)
#   pyinstaller --noconfirm --onefile --windowed --collect-all customtkinter ^
#       --name "DisenadorProCimentaciones" app_cimentacion_equipos_postes_pro.py
# =============================================================================
