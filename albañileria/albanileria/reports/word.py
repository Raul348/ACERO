"""Memoria de cálculo en Word (python-docx) con estilo de cálculo manual y ecuaciones OMML."""
from __future__ import annotations

import datetime as _dt
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from ..e070.materiales import pasos_material
from ..models import (ESTADO_CUMPLE, Check, Edificio, ResultadoEdificio, Step, fmt_kN, fmt_kNm, fmt_MPa,
                      kgf_cm2, texto_valor)
from .omml import omml

VERDE = RGBColor(0x1B, 0x7F, 0x3B)
ROJO = RGBColor(0xB0, 0x1C, 0x1C)
AZUL = RGBColor(0x1F, 0x4E, 0x96)
GRIS = RGBColor(0x55, 0x55, 0x55)


class _Memoria:
    def __init__(self, ed: Edificio, res: ResultadoEdificio):
        self.ed, self.res = ed, res
        self.doc = Document()
        self.ecuaciones_ok = 0
        self.ecuaciones_texto = 0
        self._estilos()

    # ------------------------------------------------------------ formato base
    def _estilos(self):
        st = self.doc.styles["Normal"]
        st.font.name = "Calibri"
        st.font.size = Pt(10)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        for s in self.doc.sections:
            s.left_margin = s.right_margin = Cm(2.0)
            s.top_margin = s.bottom_margin = Cm(2.0)
            s.orientation = WD_ORIENT.PORTRAIT

    def encabezado(self):
        for s in self.doc.sections:
            p = s.header.paragraphs[0]
            p.text = f"{self.ed.memoria.proyecto} — Memoria de cálculo de albañilería confinada (NTE E.070)"
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for r in p.runs:
                r.font.size = Pt(8)
                r.font.color.rgb = GRIS
            f = s.footer.paragraphs[0]
            f.text = "Programa de apoyo ALBAÑILERÍA E.070 — resultados a revisar por el ingeniero responsable"
            f.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in f.runs:
                r.font.size = Pt(7)
                r.font.color.rgb = GRIS

    def h(self, texto: str, nivel: int = 1):
        return self.doc.add_heading(texto, level=nivel)

    def p(self, texto: str = "", negrita: bool = False, color=None, tam=None, cursiva=False):
        par = self.doc.add_paragraph()
        run = par.add_run(texto)
        run.bold = negrita
        run.italic = cursiva
        if color is not None:
            run.font.color.rgb = color
        if tam:
            run.font.size = Pt(tam)
        return par

    def ecuacion_en(self, par, formula: str) -> bool:
        x = omml(formula)
        if x is None:
            par.add_run(formula)
            self.ecuaciones_texto += 1
            return False
        par._p.append(parse_xml(x))
        self.ecuaciones_ok += 1
        return True

    def tabla(self, encabezados: list[str], filas: list[list[str]], anchos: list[float] | None = None):
        t = self.doc.add_table(rows=1, cols=len(encabezados))
        t.style = "Light Grid Accent 1"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, e in enumerate(encabezados):
            c = t.rows[0].cells[i]
            c.text = ""
            r = c.paragraphs[0].add_run(e)
            r.bold = True
            r.font.size = Pt(8.5)
        for f in filas:
            cells = t.add_row().cells
            for i, v in enumerate(f):
                cells[i].text = ""
                r = cells[i].paragraphs[0].add_run(str(v))
                r.font.size = Pt(8.5)
        if anchos:
            for row in t.rows:
                for i, w in enumerate(anchos):
                    row.cells[i].width = Cm(w)
        return t

    def tabla_pasos(self, pasos: list[Step]):
        if not pasos:
            return
        t = self.doc.add_table(rows=1, cols=4)
        t.style = "Light List Accent 1"
        for i, e in enumerate(("Símbolo", "Ecuación", "Sustitución", "Valor")):
            c = t.rows[0].cells[i]
            c.text = ""
            r = c.paragraphs[0].add_run(e)
            r.bold = True
            r.font.size = Pt(8.5)
        for s in pasos:
            cells = t.add_row().cells
            cells[0].text = ""
            r = cells[0].paragraphs[0].add_run(s.simbolo)
            r.font.size = Pt(8.5)
            r.bold = True
            par = cells[1].paragraphs[0]
            if s.formula:
                self.ecuacion_en(par, s.formula)
            if s.descripcion:
                d = cells[1].add_paragraph().add_run(s.descripcion)
                d.font.size = Pt(7.5)
                d.italic = True
                d.font.color.rgb = GRIS
            cells[2].text = ""
            r = cells[2].paragraphs[0].add_run(s.sustitucion)
            r.font.size = Pt(8)
            cells[3].text = ""
            r = cells[3].paragraphs[0].add_run(s.valor)
            r.font.size = Pt(8.5)
            r.bold = True
        for row in t.rows:
            for i, w in enumerate((2.0, 6.0, 6.0, 3.0)):
                row.cells[i].width = Cm(w)

    def bloque_check(self, c: Check, nivel: int = 4):
        self.h(f"{c.nombre} — {c.referencia}", nivel)
        self.tabla_pasos(c.pasos)
        for n in c.notas:
            self.p("• " + n, cursiva=True, tam=8.5, color=ROJO if "FUERA DE ALCANCE" in n else None)
        if not c.aplica:
            if any("FUERA DE ALCANCE" in n for n in c.notas):
                self.p("FUERA DE ALCANCE", negrita=True, color=ROJO)
            else:
                self.p("Verificación informativa (no gobierna el estado).", cursiva=True, color=GRIS, tam=8.5)
            return
        par = self.doc.add_paragraph()
        r = par.add_run(f"Demanda = {texto_valor(c.demanda, c.unidad)};  Capacidad = {texto_valor(c.capacidad, c.unidad)};  ")
        r.font.size = Pt(9)
        dcr = "∞" if c.dcr == float("inf") else f"{c.dcr:.3f}"
        r2 = par.add_run(f"DCR = {dcr} → {'CUMPLE' if c.cumple else 'NO CUMPLE'}")
        r2.bold = True
        r2.font.color.rgb = VERDE if c.cumple else ROJO

    # ------------------------------------------------------------ contenido
    def portada(self):
        m = self.ed.memoria
        for _ in range(6):
            self.doc.add_paragraph()
        t = self.p("MEMORIA DE CÁLCULO", negrita=True, tam=24, color=AZUL)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t = self.p("Verificación y diseño de albañilería confinada", negrita=True, tam=16)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t = self.p("NTE E.070 Albañilería — NTE E.030 Diseño Sismorresistente (RNE, Perú)", tam=12)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.doc.add_paragraph()
        t = self.p(m.proyecto, negrita=True, tam=14)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for et, v in (("Ubicación", m.ubicacion), ("Ingeniero", m.ingeniero),
                      ("Fecha", m.fecha or _dt.date.today().strftime("%d/%m/%Y"))):
            if v:
                t = self.p(f"{et}: {v}", tam=11)
                t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.doc.add_paragraph()
        est = self.res.estado
        t = self.p(f"Estado del edificio: {est}", negrita=True, tam=14, color=VERDE if est == ESTADO_CUMPLE else ROJO)
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.doc.add_page_break()

    def s1_datos(self):
        ed, s = self.ed, self.res.sismo["X"]
        self.h("1. Datos del proyecto y del edificio", 1)
        filas = [
            ["Proyecto", ed.memoria.proyecto], ["Ubicación", ed.memoria.ubicacion or "—"],
            ["Origen de los datos", "Modelo ETABS" if ed.origen == "etabs" else "Ingreso manual"],
            ["Número de pisos N", str(ed.N)], ["Altura de entrepiso", f"{ed.h_entrepiso / 1000:.2f} m"],
            ["Área en planta típica Ap", f"{ed.Ap / 1e6:.2f} m²"],
            ["Zona sísmica / Z", f"{ed.zona} / {s.Z:.2f}"], ["Categoría / U", f"{ed.categoria} / {s.U:.2f}"],
            ["Suelo / S / Tp / Tl", f"{ed.suelo} / {s.S:.2f} / {s.Tp:.2f} s / {s.Tl:.2f} s"],
            ["Sistema", f"Albañilería confinada, R0 = 3, Ia = {ed.Ia}, Ip = {ed.Ip}"],
            ["Losa / acabados / tabiquería", f"{ed.cargas.peso_losa * 1000:.2f} / {ed.cargas.acabados * 1000:.2f} / "
                                             f"{ed.cargas.tabiqueria * 1000:.2f} kN/m²"],
            ["Sobrecarga / azotea", f"{ed.cargas.sobrecarga * 1000:.2f} / {ed.cargas.sobrecarga_azotea * 1000:.2f} kN/m²"],
        ]
        self.tabla(["Dato", "Valor"], filas, [6, 10])
        if ed.memoria.observaciones:
            self.p("Observaciones: " + ed.memoria.observaciones, cursiva=True)

    def s2_materiales(self):
        self.h("2. Materiales (E.070 Cap. 5, Tabla 9)", 1)
        self.tabla_pasos(pasos_material(self.ed.material))
        for c in self.res.checks_globales:
            self.bloque_check(c, 3)

    def s3_requisitos(self):
        self.h("3. Requisitos estructurales mínimos (E.070 Art. 19)", 1)
        self.h("3.1 Densidad de muros por dirección", 2)
        for pd in self.res.pisos:
            if pd.piso != 1:
                continue
            for c in pd.checks:
                if c.nombre.startswith("Densidad"):
                    self.bloque_check(c, 3)
        filas = [[str(pd.piso), pd.direccion, f"{pd.densidad_existente:.4f}", f"{pd.densidad_requerida:.4f}",
                  f"{pd.dcr_densidad:.3f}", "CUMPLE" if pd.dcr_densidad <= 1 else "NO CUMPLE"] for pd in self.res.pisos]
        self.tabla(["Piso", "Dir.", "ΣL·t/Ap", "ZUSN/56", "DCR", "Estado"], filas)
        self.h("3.2 Espesor efectivo y esfuerzo axial", 2)
        filas = []
        for rm in self.res.muros:
            if rm.fuera_de_alcance:
                continue
            ce = next((c for c in rm.checks if c.nombre.startswith("Espesor")), None)
            ca = next((c for c in rm.checks if c.nombre.startswith("Esfuerzo axial")), None)
            filas.append([rm.nombre, str(rm.piso), rm.direccion,
                          f"{ce.capacidad:.0f}", f"{ce.demanda:.0f}", f"{ca.demanda:.3f} ({kgf_cm2(ca.demanda):.2f})",
                          f"{ca.capacidad:.3f} ({kgf_cm2(ca.capacidad):.2f})", f"{max(ce.dcr, ca.dcr):.3f}"])
        self.tabla(["Muro", "Piso", "Dir.", "t (mm)", "h/20 (mm)", "σm MPa (kgf/cm²)", "Fa MPa (kgf/cm²)", "DCR"], filas)

    def s4_sismo(self):
        self.h("4. Análisis sísmico (E.030 2018 y E.070)", 1)
        for d in ("X", "Y"):
            s = self.res.sismo[d]
            self.h(f"4.{1 if d == 'X' else 2} Dirección {d} — {s.origen}", 2)
            self.tabla_pasos(s.pasos)
            filas = [[str(i + 1), f"{s.alturas[i] / 1000:.2f}", f"{s.pesos[i] / 1000:.1f}", f"{s.F[i] / 1000:.2f}",
                      f"{s.VE[i] / 1000:.2f}", f"{s.Ve[i] / 1000:.2f}"] for i in range(len(s.VE))]
            self.tabla(["Nivel", "hi (m)", "Pi (kN)", "Fi (kN)", "VE severo (kN)", "Ve moderado (kN)"], filas)
        self.p("Sismo moderado = 0.5 × sismo severo (E.070 Art. 2). Ve y Me de cada muro corresponden al sismo "
               "moderado; VE es el cortante de entrepiso del sismo severo.", cursiva=True, tam=9)
        if any(c.nombre.startswith("Deriva") for pd in self.res.pisos for c in pd.checks):
            self.h("4.3 Derivas de entrepiso", 2)
            for pd in self.res.pisos:
                for c in pd.checks:
                    if c.nombre.startswith("Deriva"):
                        self.bloque_check(c, 3)

    def s5_muros(self):
        self.h("5. Muros por piso: fisuración ante sismo moderado (E.070 Art. 26.2 y 26.3)", 1)
        filas = []
        for rm in self.res.muros:
            v = rm.valores
            if rm.fuera_de_alcance:
                filas.append([rm.nombre, str(rm.piso), rm.direccion] + ["—"] * 6 + ["FUERA DE ALCANCE"])
                continue
            filas.append([rm.nombre, str(rm.piso), rm.direccion, f"{v['Pg'] / 1000:.1f}", f"{v['Ve'] / 1000:.1f}",
                          f"{v['Me'] / 1e6:.1f}", f"{v['alfa']:.3f}", f"{v['Vm'] / 1000:.1f}", f"{v['Vadm'] / 1000:.1f}",
                          f"{v['dcr_fis']:.3f} " + ("✔" if v["dcr_fis"] <= 1 else "✘")])
        self.tabla(["Muro", "Piso", "Dir.", "Pg kN", "Ve kN", "Me kN·m", "α", "Vm kN", "0.55Vm kN", "DCR fis."], filas)
        self.h("5.1 Cálculo detallado por muro", 2)
        for rm in self.res.muros:
            self.h(f"Muro {rm.nombre} — piso {rm.piso}, dirección {rm.direccion}: {rm.estado}", 3)
            for c in rm.checks:
                if c.nombre.startswith(("Columna", "Solera")):
                    continue
                self.bloque_check(c, 4)

    def s6_global(self):
        self.h("6. Resistencia global al corte: ΣVm vs VE (E.070 Art. 26.4)", 1)
        filas = [[str(pd.piso), pd.direccion, f"{pd.suma_Vm / 1000:.1f}", f"{pd.VE / 1000:.1f}", f"{pd.relacion:.2f}",
                  "Elástico (refuerzo mínimo)" if pd.elastico else ("Cumple" if pd.relacion >= 1 else "NO CUMPLE")]
                 for pd in self.res.pisos]
        self.tabla(["Piso", "Dir.", "ΣVm (kN)", "VE (kN)", "ΣVm/VE", "Comportamiento"], filas)
        for pd in self.res.pisos:
            for c in pd.checks:
                if c.nombre.startswith("Resistencia al corte"):
                    self.bloque_check(c, 3)

    def s7_severo(self):
        self.h("7. Diseño por sismo severo (E.070 Art. 27)", 1)
        self.p("Vu = Ve·(Vm1/Ve1) y Mu = Me·(Vm1/Ve1), con 2 ≤ Vm1/Ve1 ≤ 3 tomado del muro en el primer piso. "
               "Los muros del primer piso y los de pisos superiores con Vu ≥ Vm se consideran agrietados y se diseñan "
               "por capacidad (Tabla 11).", tam=9)
        filas = []
        for rm in self.res.muros:
            v = rm.valores
            if rm.fuera_de_alcance:
                continue
            filas.append([rm.nombre, str(rm.piso), rm.direccion, f"{v['factor']:.2f}", f"{v['Vu'] / 1000:.1f}",
                          f"{v['Mu'] / 1e6:.1f}", f"{v['Vm'] / 1000:.1f}", "Sí" if v["agrietado"] else "No",
                          v["tipo_diseno"], "Sí" if v.get("refuerzo_horizontal") else "No"])
        self.tabla(["Muro", "Piso", "Dir.", "Vm1/Ve1", "Vu kN", "Mu kN·m", "Vm kN", "Agrietado", "Diseño",
                    "Ref. horiz."], filas)

    def s8_confinamiento(self):
        self.h("8. Columnas de confinamiento y vigas soleras (E.070 Art. 27, Tabla 11)", 1)
        filas = []
        for rm in self.res.muros:
            rc = rm.confinamiento
            if rc is None:
                continue
            filas.append([f"{rm.nombre} (P{rm.piso})", rc.tipo, f"{rc.Vc_ext / 1000:.1f}", f"{rc.T_ext / 1000:.1f}",
                          f"{rc.C_ext / 1000:.1f}", f"{rc.Acf / 100:.0f}", rc.seccion_ext, rc.refuerzo_ext,
                          f"{rc.Ts / 1000:.1f}", rc.refuerzo_solera])
        self.tabla(["Muro", "Diseño", "Vc kN", "T kN", "C kN", "Acf cm²", "Columna", "Refuerzo", "Ts kN", "Solera"], filas)
        self.h("8.1 Cálculo detallado", 2)
        for rm in self.res.muros:
            rc = rm.confinamiento
            if rc is None:
                continue
            self.h(f"Muro {rm.nombre} — piso {rm.piso} ({rc.tipo})", 3)
            self.p(f"Columnas extremas: {rc.seccion_ext}, {rc.refuerzo_ext}; estribos {rc.estribos}."
                   + (f" Columnas interiores: {rc.seccion_int}, {rc.refuerzo_int}." if rc.seccion_int else "")
                   + f" Solera {rc.seccion_solera}: {rc.refuerzo_solera}; {rc.estribos_solera}.", tam=9)
            for c in rc.checks:
                self.bloque_check(c, 4)
            for n in rc.notas:
                self.p("• " + n, cursiva=True, tam=8.5)

    def s9_no_portantes(self):
        self.h("9. Muros no portantes, tabiques, parapetos y cercos (E.070 Art. 29)", 1)
        if not self.res.no_portantes:
            self.p("No se ingresaron elementos no portantes.")
        else:
            filas = [[n.nombre, n.tipo, str(n.valores.get("caso")), f"{n.valores['C1']:.2f}",
                      f"{n.valores['w'] * 1000:.3f}", f"{n.valores['m']:.4f}", f"{n.valores['Ms'] / 1000:.3f}",
                      f"{n.valores['fm']:.3f}", f"{n.valores['ft']:.3f}", n.estado,
                      f"{n.valores['sep_arriostre'] / 1000:.2f} m" if "sep_arriostre" in n.valores else "—"]
                     for n in self.res.no_portantes]
            self.tabla(["Elemento", "Tipo", "Caso", "C1", "w kN/m²", "m", "Ms kN·m/m", "fm MPa", "f't MPa", "Estado",
                        "Arriostres @"], filas)
            for n in self.res.no_portantes:
                for c in n.checks:
                    self.bloque_check(c, 3)
        perp = [c for rm in self.res.muros for c in rm.checks if c.nombre.startswith("Carga perpendicular")]
        if perp:
            self.h("9.1 Muros portantes ante carga perpendicular", 2)
            for c in perp:
                self.bloque_check(c, 3)

    def s10_observaciones(self):
        self.h("10. Observaciones y advertencias", 1)
        avs = list(dict.fromkeys(self.res.advertencias))
        if not avs:
            self.p("Sin advertencias.")
        for a in avs:
            self.p("• " + a, tam=9)
        self.h("Alcances no implementados (FUERA DE ALCANCE)", 2)
        for f in self.res.fuera_de_alcance:
            self.p("• FUERA DE ALCANCE: " + f, tam=9, color=ROJO)

    def s11_resumen(self):
        self.h("11. Resumen final", 1)
        r = self.res
        nc = [f"{m.nombre} (P{m.piso})" for m in r.muros if m.estado == "NO CUMPLE"]
        filas = [[str(pd.piso), pd.direccion, f"{pd.dcr_densidad:.3f}", f"{pd.relacion:.2f}",
                  ", ".join(pd.muros_no_cumplen) or "—", ", ".join(pd.muros_agrietados) or "—"] for pd in r.pisos]
        self.tabla(["Piso", "Dir.", "DCR densidad", "ΣVm/VE", "Muros que no cumplen", "Muros agrietados (sismo severo)"],
                   filas)
        npnc = [n.nombre for n in r.no_portantes if n.estado == "NO CUMPLE"]
        self.p(f"Muros que no cumplen: {', '.join(nc) if nc else 'ninguno'}.", tam=10)
        self.p(f"Elementos no portantes que requieren arriostre o cambio: {', '.join(npnc) if npnc else 'ninguno'}.", tam=10)
        dcr = "∞" if r.dcr_max == float("inf") else f"{r.dcr_max:.3f}"
        self.p(f"DCR máximo del edificio = {dcr} → {r.estado}", negrita=True, tam=13,
               color=VERDE if r.estado == ESTADO_CUMPLE else ROJO)
        self.p("Este documento fue generado por un programa de apoyo. El ingeniero responsable debe revisar los datos "
               "y resultados antes de usarlos en planos o expedientes.", cursiva=True, tam=8.5, color=GRIS)

    def generar(self, ruta: Path) -> Path:
        self.encabezado()
        self.portada()
        self.s1_datos()
        self.s2_materiales()
        self.s3_requisitos()
        self.s4_sismo()
        self.s5_muros()
        self.s6_global()
        self.s7_severo()
        self.s8_confinamiento()
        self.s9_no_portantes()
        self.s10_observaciones()
        self.s11_resumen()
        ruta = Path(ruta)
        ruta.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(ruta)
        return ruta


def generar_memoria(ed: Edificio, res: ResultadoEdificio, ruta: str | Path) -> Path:
    """Genera la memoria de cálculo en Word y devuelve la ruta."""
    m = _Memoria(ed, res)
    out = m.generar(Path(ruta))
    generar_memoria.ultimas_estadisticas = {"omml": m.ecuaciones_ok, "texto": m.ecuaciones_texto}
    return out


__all__ = ["generar_memoria", "fmt_kN", "fmt_kNm", "fmt_MPa"]
