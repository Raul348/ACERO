"""Prueba EN VIVO (opcional, Windows + ETABS 22): crea un edificio pequeño de albañilería en una instancia
oculta, lo analiza, lo lee con el lector del programa y compara Pg, Ve y Me.

Edificio: 2 pisos de 2.60 m, planta 6 x 4 m, 4 muros de albañilería t = 13 cm (2 en X de 6 m, 2 en Y de 4 m),
losa rígida de 20 cm, CM = 3 kPa (Dead + SuperDead), CV = 2 kPa, sismo estático SX/SY con C = 0.375.

Uso:  python tools/test_etabs_live.py
Comprueba:
  - Pg (PG = D + 0.25 L) de cada pier ≈ carga tributaria de losa + peso propio (error en %).
  - Σ V2 de los piers de una dirección ≈ cortante del piso (Story Forces) del caso SX/SY.
  - Me (M3 en Bottom) > 0 y α = Ve·L/Me entre 0.2 y 3.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from albanileria.etabs.batch import ConfigEtabs, verificar_snapshot  # noqa: E402
from albanileria.etabs.client import SALIDA_REAL, ClienteEtabs  # noqa: E402
from albanileria.etabs.reader import leer_modelo  # noqa: E402

H, NP = 2600.0, 2
LX, LY, T = 6000.0, 4000.0, 130.0


def crear_modelo(cli: ClienteEtabs) -> Path:
    m = cli.m
    cli.nuevo_modelo()
    m.File.NewGridOnly(NP, H, H, 2, 2, LX, LY)
    m.SetPresentUnits(9)
    m.PropMaterial.SetMaterial("ALB_KK", 8)                       # 8 = Masonry
    m.PropMaterial.SetMPIsotropic("ALB_KK", 3187.0, 0.25, 6.5e-6)
    m.PropMaterial.SetWeightAndMass("ALB_KK", 1, 18e-6)           # 18 kN/m³ en N/mm³
    m.PropMaterial.SetMaterial("C210", 2)
    m.PropMaterial.SetMPIsotropic("C210", 21737.0, 0.2, 9.9e-6)
    m.PropArea.SetWall("MURO13", 1, 1, "ALB_KK", T)               # Specified, ShellThin
    m.PropArea.SetSlab("LOSA20", 0, 1, "C210", 200.0)
    m.Diaphragm.SetDiaphragm("D1", False)
    m.LoadPatterns.Add("Live", 3, 0, True)
    m.LoadPatterns.Add("SuperDead", 2, 0, True)
    muros = {"X1": [(0, 0), (LX, 0)], "X2": [(0, LY), (LX, LY)], "Y1": [(0, 0), (0, LY)], "Y2": [(LX, 0), (LX, LY)]}
    for nombre in muros:
        m.PierLabel.SetPier(nombre)
    for i in range(1, NP + 1):
        z0, z1 = (i - 1) * H, i * H
        for nombre, ((xa, ya), (xb, yb)) in muros.items():
            r = m.AreaObj.AddByCoord(4, [xa, xb, xb, xa], [ya, yb, yb, ya], [z0, z0, z1, z1], "", "MURO13")
            area = list(r)[3]
            m.AreaObj.SetPier(area, nombre)
        r = m.AreaObj.AddByCoord(4, [0, LX, LX, 0], [0, 0, LY, LY], [z1] * 4, "", "LOSA20")
        losa = list(r)[3]
        m.AreaObj.SetDiaphragm(losa, "D1")
        m.AreaObj.SetLoadUniform(losa, "SuperDead", 1.0e-3 if i < NP else 0.5e-3, 6, True, "Global")
        m.AreaObj.SetLoadUniform(losa, "Live", 2.0e-3 if i < NP else 1.0e-3, 6, True, "Global")
    # empotrar la base
    nombres = list(m.PointObj.GetNameList())[1]
    for p in nombres:
        c = list(m.PointObj.GetCoordCartesian(p))
        if abs(c[2]) < 1e-6:
            m.PointObj.SetRestraint(p, [True] * 6)
    for caso, d in (("SX", 1), ("SY", 2)):
        m.LoadPatterns.Add(caso, 5, 0, True)                      # 5 = Quake
        m.LoadPatterns.AutoSeismic.SetUserCoefficient(caso, d, 0.05, False, 0, 0, 0.375, 1.0)
    cli.crear_combo("PG", {"Dead": 1.0, "SuperDead": 1.0, "Live": 0.25})
    cli.crear_combo("SERV", {"Dead": 1.0, "SuperDead": 1.0, "Live": 1.0})
    SALIDA_REAL.mkdir(parents=True, exist_ok=True)
    ruta = SALIDA_REAL / "prueba_viva_albanileria.EDB"
    m.File.Save(str(ruta))
    cli.ruta_modelo = ruta
    return ruta


def main() -> int:
    with ClienteEtabs() as cli:
        cli.iniciar()
        ruta = crear_modelo(cli)
        print(f"Modelo creado: {ruta}")
        cli.analizar()
        snap = leer_modelo(cli, {"X": "SX", "Y": "SY"}, "PG", "SERV")
    snap.guardar(SALIDA_REAL / "prueba_viva_snapshot.json")
    ed, res = verificar_snapshot(snap, ConfigEtabs(casos_sismo={"X": "SX", "Y": "SY"}, combo_gravedad="PG",
                                                   combo_servicio="SERV", Ap=LX * LY))
    ok = True
    print("\nMuro  Piso  Pg ETABS (kN)  Pg mano (kN)  dif %   Ve (kN)  Me (kN·m)  α")
    for m in ed.muros:
        niveles = NP - m.piso + 1
        trib = LX * LY / 4
        losa = 200 * 24e-6
        D = sum(trib * (losa + (1.0e-3 if j < NP else 0.5e-3)) for j in range(m.piso, NP + 1))
        L = sum(trib * (2.0e-3 if j < NP else 1.0e-3) for j in range(m.piso, NP + 1))
        pp = 18e-6 * T * m.L * H * niveles
        Pg_mano = D + pp + 0.25 * L
        dif = (m.Pg - Pg_mano) / Pg_mano * 100 if m.Pg else float("nan")
        alfa = m.Ve * m.L / m.Me if m.Me else float("nan")
        print(f"{m.nombre:4s}  {m.piso:4d}  {m.Pg / 1000:12.1f}  {Pg_mano / 1000:12.1f}  {dif:6.1f}  "
              f"{m.Ve / 1000:7.1f}  {m.Me / 1e6:9.1f}  {alfa:5.2f}")
        if not (abs(dif) < 35 and m.Me > 0 and 0.2 < alfa < 3):
            ok = False
    for pd in res.pisos:
        suma = sum(m.Ve for m in ed.muros if m.piso == pd.piso and m.direccion == pd.direccion) * 2
        dif = (suma - pd.VE) / pd.VE * 100 if pd.VE else float("nan")
        print(f"Piso {pd.piso} {pd.direccion}: ΣV2 piers = {suma / 1000:.1f} kN, Story Forces = {pd.VE / 1000:.1f} kN "
              f"({dif:+.1f} %)")
        if abs(dif) > 10:
            ok = False
    print("\nRESULTADO:", "OK" if ok else "REVISAR (diferencias grandes)")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
