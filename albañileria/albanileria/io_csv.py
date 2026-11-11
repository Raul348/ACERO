"""Importar / exportar muros y elementos no portantes en CSV (unidades de presentación)."""
from __future__ import annotations

import csv
from pathlib import Path

from .models import ElementoNoPortante, Muro

# (encabezado, atributo, tipo, factor presentación → interno, opcional)
COLUMNAS_MURO = [
    ("Nombre", "nombre", str, 1, False),
    ("Piso", "piso", int, 1, False),
    ("Dir", "direccion", str, 1, False),
    ("L (m)", "L", float, 1000, False),
    ("t (cm)", "t", float, 10, False),
    ("h (m)", "h", float, 1000, False),
    ("Portante", "portante", bool, 1, False),
    ("x (m)", "x", float, 1000, False),
    ("y (m)", "y", float, 1000, False),
    ("Pg (kN)", "Pg", float, 1000, True),
    ("Pm (kN)", "Pm", float, 1000, True),
    ("Área trib (m²)", "area_trib", float, 1e6, True),
    ("Lm (m)", "Lm", float, 1000, True),
    ("Nc", "Nc", int, 1, False),
    ("Tabique", "tabique", bool, 1, False),
    ("Grupo", "grupo", str, 1, False),
]

COLUMNAS_NP = [
    ("Nombre", "nombre", str, 1, False),
    ("Tipo", "tipo", str, 1, False),
    ("Caso", "caso", int, 1, False),
    ("a (m)", "a", float, 1000, False),
    ("b (m)", "b", float, 1000, False),
    ("t (cm)", "t", float, 10, False),
    ("C1", "C1", float, 1, True),
    ("Piso", "piso", int, 1, False),
]


def a_texto(obj, col) -> str:
    _, attr, tipo, f, _ = col
    v = getattr(obj, attr)
    if v is None:
        return ""
    if tipo is bool:
        return "sí" if v else "no"
    if tipo is float:
        return f"{v / f:.4g}" if abs(v / f) < 1e5 else f"{v / f:.1f}"
    return str(v)


def de_texto(txt: str, col):
    _, attr, tipo, f, opcional = col
    txt = (txt or "").strip()
    if txt == "":
        if opcional:
            return None
        if tipo is str:
            return ""
        raise ValueError(f"Falta el valor de '{col[0]}'")
    if tipo is bool:
        return txt.lower() in ("sí", "si", "s", "1", "true", "x", "yes")
    if tipo is int:
        return int(float(txt.replace(",", ".")))
    if tipo is float:
        return float(txt.replace(",", ".")) * f
    if attr == "direccion":
        return txt.upper()
    return txt


def exportar_muros(muros: list[Muro], ruta: str | Path) -> None:
    with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow([c[0] for c in COLUMNAS_MURO])
        for m in muros:
            w.writerow([a_texto(m, c) for c in COLUMNAS_MURO])


def importar_muros(ruta: str | Path) -> list[Muro]:
    muros = []
    with open(ruta, newline="", encoding="utf-8-sig") as fh:
        muestra = fh.read(2048)
        fh.seek(0)
        sep = ";" if muestra.count(";") >= muestra.count(",") else ","
        r = csv.DictReader(fh, delimiter=sep)
        for fila in r:
            kw = {}
            for c in COLUMNAS_MURO:
                if c[0] in fila:
                    kw[c[1]] = de_texto(fila[c[0]], c)
            muros.append(Muro(**kw))
    return muros


def exportar_tabla(encabezados: list[str], filas: list[list], ruta: str | Path) -> None:
    with open(ruta, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(encabezados)
        w.writerows(filas)


__all__ = ["COLUMNAS_MURO", "COLUMNAS_NP", "exportar_muros", "importar_muros", "exportar_tabla", "ElementoNoPortante"]
