"""Análisis sísmico estático E.030 (2018): peso, cortante basal y distribución en altura."""
from __future__ import annotations

from ..models import Edificio, ResultadoSismo, Step, fmt_kN
from . import norma


def pesos_por_nivel(ed: Edificio) -> list[float]:
    """Peso sísmico por nivel P = D + %L (E.030 Art. 26): losa + mitad de los muros
    del piso inferior y del superior."""
    if ed.pesos_override:
        return list(ed.pesos_override)
    pct = norma.porcentaje_carga_viva(ed.categoria)
    g = ed.material.gamma
    pesos = []
    for i in range(1, ed.N + 1):
        azotea = i == ed.N
        wd = ed.cargas.muerta if not azotea else ed.cargas.peso_losa + ed.cargas.acabados
        wl = ed.cargas.sobrecarga_azotea if azotea else ed.cargas.sobrecarga
        pct_i = 0.25 if azotea else pct
        p = ed.Ap * (wd + pct_i * wl)
        muros_inf = sum(g * m.t * m.L * m.h for m in ed.muros_piso(i))
        muros_sup = sum(g * m.t * m.L * m.h for m in ed.muros_piso(i + 1)) if not azotea else 0.0
        p += 0.5 * muros_inf + 0.5 * muros_sup
        pesos.append(p)
    return pesos


def analisis_estatico(ed: Edificio, direccion: str = "X") -> ResultadoSismo:
    r = ResultadoSismo(direccion=direccion)
    r.Z = norma.factor_Z(ed.zona)
    r.U = ed.U if ed.U else norma.factor_U(ed.categoria)
    r.S = norma.factor_S(ed.zona, ed.suelo)
    tp, tl = norma.periodos_suelo(ed.suelo)
    r.Tp = ed.Tp or tp
    r.Tl = ed.Tl or tl
    alturas = []
    acc = 0.0
    for h in ed.alturas_entrepiso:
        acc += h
        alturas.append(acc)
    r.alturas = alturas
    hn = alturas[-1] / 1000.0
    r.T = hn / norma.CT_ALBANILERIA
    r.C = norma.factor_C(r.T, r.Tp, r.Tl)
    r.R = norma.R0_ALBANILERIA * ed.Ia * ed.Ip
    cr = max(r.C / r.R, norma.C_SOBRE_R_MIN)
    r.k = norma.exponente_k(r.T)
    r.pesos = pesos_por_nivel(ed)
    r.P = sum(r.pesos)
    r.V = r.Z * r.U * cr * r.S * r.P
    s = sum(p * h ** r.k for p, h in zip(r.pesos, alturas))
    r.F = [r.V * p * h ** r.k / s for p, h in zip(r.pesos, alturas)]
    r.VE = [sum(r.F[i:]) for i in range(ed.N)]
    if ed.VE_override and ed.VE_override.get(direccion):
        r.VE = list(ed.VE_override[direccion])
        r.origen = "ETABS (cortante por piso del modelo)"
    r.Ve = [v * ed.opciones.factor_moderado for v in r.VE]

    P = r.pasos
    P.append(Step("Z", f"{r.Z:.2f}", f"Factor de zona (zona {ed.zona})", "Z", "E.030 Tabla N°1"))
    P.append(Step("U", f"{r.U:.2f}", f"Factor de uso (categoría {ed.categoria})", "U", "E.030 Tabla N°5"))
    P.append(Step("S", f"{r.S:.2f}", f"Factor de suelo ({ed.suelo})", "S", "E.030 Tabla N°3"))
    P.append(Step("Tp", f"{r.Tp:.2f} s", "Periodo que define la plataforma", "Tp", "E.030 Tabla N°4"))
    P.append(Step("Tl", f"{r.Tl:.2f} s", "Periodo que define el inicio de desplazamiento constante", "Tl", "E.030 Tabla N°4"))
    P.append(Step("T", f"{r.T:.3f} s", "Periodo fundamental estimado (CT = 60, albañilería)",
                  "T = hn/CT", f"{hn:.2f}/{norma.CT_ALBANILERIA:.0f}"))
    if r.T < r.Tp:
        P.append(Step("C", f"{r.C:.3f}", "T < Tp", "C = 2.5", "2.5"))
    elif r.T < r.Tl:
        P.append(Step("C", f"{r.C:.3f}", "Tp ≤ T < Tl", "C = 2.5·Tp/T", f"2.5·{r.Tp:.2f}/{r.T:.3f}"))
    else:
        P.append(Step("C", f"{r.C:.3f}", "T ≥ Tl", "C = 2.5·Tp·Tl/T^2", f"2.5·{r.Tp:.2f}·{r.Tl:.2f}/{r.T:.3f}^2"))
    P.append(Step("R", f"{r.R:.2f}", "Coeficiente de reducción (R0 = 3, albañilería confinada)",
                  "R = R0·Ia·Ip", f"3·{ed.Ia:.2f}·{ed.Ip:.2f}"))
    P.append(Step("C/R", f"{r.C / r.R:.3f}", "Debe ser ≥ 0.11" + (" → se usa 0.11" if r.C / r.R < 0.11 else ""),
                  "C/R ≥ 0.11", f"{r.C:.3f}/{r.R:.2f}"))
    P.append(Step("P", fmt_kN(r.P), "Peso sísmico total (D + %L, E.030 Art. 26)", "P = ΣPi",
                  " + ".join(f"{p / 1000:.1f}" for p in r.pesos)))
    P.append(Step("V", fmt_kN(r.V), "Cortante basal, sismo severo (E.030 Art. 28.2)",
                  "V = Z·U·C·S·P/R", f"{r.Z:.2f}·{r.U:.2f}·{cr * r.R:.3f}·{r.S:.2f}·{r.P / 1000:.1f}/{r.R:.2f}"))
    P.append(Step("k", f"{r.k:.2f}", "Exponente de distribución en altura", "k = 1 (T ≤ 0.5 s)" if r.T <= 0.5 else "k = 0.75 + 0.5·T", ""))
    for i, (p, h, f) in enumerate(zip(r.pesos, alturas, r.F), start=1):
        P.append(Step(f"F{i}", fmt_kN(f), f"Fuerza en el nivel {i} (h = {h / 1000:.2f} m)",
                      "Fi = αi·V,  αi = Pi·hi^k/Σ(Pj·hj^k)",
                      f"α{i} = {p * h ** r.k / s:.4f};  F{i} = {p * h ** r.k / s:.4f}·{r.V / 1000:.2f}"))
    for i, (ve, vm) in enumerate(zip(r.VE, r.Ve), start=1):
        P.append(Step(f"VE{i}", fmt_kN(ve), f"Cortante del piso {i}, sismo severo ({r.origen})", "VEi = ΣFj (j ≥ i)", ""))
        P.append(Step(f"Ve{i}", fmt_kN(vm), f"Cortante del piso {i}, sismo moderado (E.070 Art. 2)",
                      "Vei = 0.5·VEi", f"0.5·{ve / 1000:.2f}"))
    return r
