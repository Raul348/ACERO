"""Clasificación de piers: albañilería (se verifica) o concreto/otros (sólo se listan)."""
from __future__ import annotations

PALABRAS = ("ALB", "MURO", "KK", "KING", "LADRILLO", "MASON", "BRICK")


def es_albanileria(tipo_material: str, material: str, seccion: str = "", pier: str = "",
                   lista_usuario: list[str] | None = None) -> tuple[bool, str]:
    if lista_usuario is not None and pier in lista_usuario:
        return True, "marcado por el usuario"
    if (tipo_material or "").lower() == "masonry":
        return True, "material de tipo Masonry"
    texto = f"{material} {seccion}".upper()
    for p in PALABRAS:
        if p in texto:
            return True, f"el nombre contiene '{p}'"
    if (tipo_material or "").lower() == "concrete":
        return False, "placa de concreto (se verifica con el programa de concreto)"
    return False, f"material '{material}' ({tipo_material or 'tipo desconocido'})"


def clasificar_muros(muros, lista_usuario: list[str] | None = None) -> None:
    for m in muros:
        m.es_albanileria, m.motivo = es_albanileria(m.tipo_material, m.material, m.seccion, m.pier, lista_usuario)
