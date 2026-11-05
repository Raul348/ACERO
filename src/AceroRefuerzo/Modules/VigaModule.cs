using System;
using System.Collections.Generic;
using System.Linq;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace AceroRefuerzo.Modules
{
    /// <summary>
    /// Vigas: acero corrido superior e inferior (con patas en los apoyos), bastones superiores en apoyos,
    /// bastones inferiores al centro y estribos cerrados con ganchos de 135° distribuidos por zonas.
    /// Con "Vigas continuas", las vigas alineadas de un eje se unen: el acero corrido pasa sobre los apoyos
    /// intermedios y se traslapa en tramos óptimos dentro de la zona permitida de cada luz.
    /// </summary>
    internal sealed class VigaModule : IRebarModule, IGroupModule
    {
        public ElementKind Kind => ElementKind.Viga;
        public string Title => "Vigas";
        public string Prompt => "Seleccione las vigas a armar y pulse Finalizar";

        private static readonly string[] Zonas =
        {
            "Tercio central de cada luz",
            "Cuartos extremos, fuera de 2h del apoyo",
            "Fuera de 2h de la cara del apoyo",
            "Cualquier punto"
        };

        public bool Accepts(Element e) =>
            e is FamilyInstance && e.Category?.BuiltInCategory == BuiltInCategory.OST_StructuralFraming;

        private static LocalFrame Frame(Element e)
        {
            Line line = (e.Location as LocationCurve)?.Curve as Line
                        ?? throw new InvalidOperationException("Solo se admiten vigas rectas.");
            XYZ p0 = line.GetEndPoint(0);
            XYZ x = (line.GetEndPoint(1) - p0).Normalize();
            if (Math.Abs(x.Z) > 0.99) throw new InvalidOperationException("La viga es vertical.");
            XYZ y = XYZ.BasisZ.CrossProduct(x).Normalize();
            XYZ z = x.CrossProduct(y).Normalize();
            return LocalFrame.Measure(e, p0, x, y, z);
        }

        // ------------------------------------------------------------------ Vigas continuas (cadenas)

        /// <summary>Una viga dentro de la cadena: su marco propio y sus extremos en el eje de la cadena.</summary>
        private sealed class Luz
        {
            public Element H;
            public LocalFrame F;
            public double Xa, Xb;   // en el marco de referencia de la cadena
        }

        /// <summary>Centro de la viga en coordenadas globales.</summary>
        private static XYZ Centro(LocalFrame f) => f.P((f.MinX + f.MaxX) / 2, (f.MinY + f.MaxY) / 2, (f.MinZ + f.MaxZ) / 2);

        /// <summary>Extremos de la viga <paramref name="f"/> proyectados en el eje X del marco <paramref name="r"/>.</summary>
        private static (double xa, double xb) Extremos(LocalFrame r, LocalFrame f)
        {
            double cy = (f.MinY + f.MaxY) / 2, cz = (f.MinZ + f.MaxZ) / 2;
            double a = r.ToLocal(f.P(f.MinX, cy, cz)).X, b = r.ToLocal(f.P(f.MaxX, cy, cz)).X;
            return (Math.Min(a, b), Math.Max(a, b));
        }

        /// <summary>Misma línea (eje y cota) y misma sección.</summary>
        private static bool MismaLinea(LocalFrame r, LocalFrame f)
        {
            if (Math.Abs(r.X.DotProduct(f.X)) < 0.999) return false;
            XYZ c = r.ToLocal(Centro(f));
            double cy = (r.MinY + r.MaxY) / 2, cz = (r.MinZ + r.MaxZ) / 2;
            return Math.Abs(c.Y - cy) < Un.Cm(3) && Math.Abs(c.Z - cz) < Un.Cm(3)
                   && Math.Abs(r.SizeY - f.SizeY) < Un.Cm(1.5) && Math.Abs(r.SizeZ - f.SizeZ) < Un.Cm(1.5);
        }

        /// <summary>
        /// Agrupa las vigas alineadas (misma línea y sección, separadas por un apoyo de hasta 1.5 m) en cadenas
        /// ordenadas a lo largo del eje. Cada cadena usa como referencia el marco de su primera viga.
        /// </summary>
        private static List<List<Luz>> Cadenas(IList<Element> hosts)
        {
            var cadenas = new List<List<Luz>>();
            foreach (Element h in hosts)
            {
                LocalFrame f = Frame(h);
                List<Luz> destino = cadenas.FirstOrDefault(c => MismaLinea(c[0].F, f));
                if (destino == null) cadenas.Add(destino = new List<Luz>());
                var (xa, xb) = Extremos(destino[0].F, f);
                destino.Add(new Luz { H = h, F = f, Xa = xa, Xb = xb });
            }

            // Ordenar por posición y cortar la cadena donde el vacío (apoyo) sea mayor que 1.5 m.
            var res = new List<List<Luz>>();
            foreach (List<Luz> c in cadenas)
            {
                var ord = c.OrderBy(l => l.Xa).ToList();
                var actual = new List<Luz> { ord[0] };
                for (int i = 1; i < ord.Count; i++)
                {
                    if (ord[i].Xa - actual[actual.Count - 1].Xb > Un.Cm(150)) { res.Add(actual); actual = new List<Luz>(); }
                    actual.Add(ord[i]);
                }
                res.Add(actual);
            }
            return res;
        }

        /// <summary>Luces de la cadena en cm, medidas desde el inicio de la primera viga.</summary>
        private static List<(double xa, double xb)> LucesCm(List<Luz> cadena)
        {
            double x0 = cadena[0].Xa;
            return cadena.Select(l => (Un.ToCm(l.Xa - x0), Un.ToCm(l.Xb - x0))).ToList();
        }

        public List<List<Element>> Agrupar(IList<Element> hosts, FormValues v) =>
            v != null && v.I("modoLong") == 1
                ? Cadenas(hosts).Select(c => c.Select(l => l.H).ToList()).ToList()
                : hosts.Select(h => new List<Element> { h }).ToList();

        // ------------------------------------------------------------------ Formulario

        public ParamForm CreateForm(Document doc, IList<Element> hosts, IList<BarItem> bars)
        {
            Element sample = hosts[0];
            LocalFrame f = Frame(sample);
            double b = Un.ToCm(f.SizeY), h = Un.ToCm(f.SizeZ), L = Un.ToCm(f.SizeX);

            // Cadena de la viga de muestra (para la figura de vigas continuas).
            List<Luz> cadena = Cadenas(hosts).First(c => c.Any(l => l.H.Id == sample.Id));
            List<(double xa, double xb)> luces = LucesCm(cadena);
            string eje = cadena.Count > 1 ? $"   ·   {cadena.Count} vigas alineadas en el eje" : "";

            var form = new ParamForm(App.Name, "viga", "Acero en VIGAS",
                $"Sección {b:0.#} × {h:0.#} cm   ·   Luz libre {L / 100:0.00} m{eje}", bars);

            form.Section("Recubrimiento");
            form.Num("rec", "Recubrimiento libre", 4, "cm");

            form.Section("Acero corrido superior");
            form.Bar("barSup", "Diámetro", 15.875);
            form.Num("nSup", "Número de barras", 2, "u", 0, 1, 20);
            form.Num("pataSup", "Pata en apoyos, medida por fuera (0 = recta)", 30, "cm");

            form.Section("Acero corrido inferior");
            form.Bar("barInf", "Diámetro", 15.875);
            form.Num("nInf", "Número de barras", 3, "u", 0, 1, 20);
            form.Num("pataInf", "Pata en apoyos, medida por fuera (0 = recta)", 30, "cm");
            form.Num("prol", "Prolongación dentro de apoyos", 0, "cm");

            form.Section("Longitud de las barras corridas");
            form.Choice("modoLong", "Barras corridas", new[]
            {
                "Una barra por viga (patas en ambos apoyos)",
                "Vigas continuas: tramos óptimos con traslapes"
            }, 0);
            form.Choice("zonaSup", "Zona de traslape del acero superior", Zonas, 0);
            form.Choice("zonaInf", "Zona de traslape del acero inferior", Zonas, 1);
            form.Num("ls", "Traslape (0 = 40 diámetros)", 0, "cm", 0, 0, 500);
            form.Txt("comerciales", "Longitudes comerciales de barra (m)", "9, 12");
            form.Note("Con \"Vigas continuas\" seleccione todas las vigas alineadas del eje: se unen, el acero corrido pasa sobre los " +
                      "apoyos intermedios (patas solo en los extremos) y el programa ubica los traslapes dentro de la zona permitida " +
                      "para que los tramos aprovechen las barras comerciales con el menor desperdicio. Verifique la zona con su norma.");

            form.Section("Bastones (2ª capa)");
            form.Num("bsN", "Bastones superiores por apoyo", 0, "u", 0, 0, 20);
            form.Bar("bsBar", "Diámetro bastón superior", 15.875);
            form.Num("bsFr", "Longitud bastón sup. (fracción de L)", 0.30, "×L", 2, 0.05, 0.5);
            form.Num("biN", "Bastones inferiores al centro", 0, "u", 0, 0, 20);
            form.Bar("biBar", "Diámetro bastón inferior", 15.875);
            form.Num("biFr", "Longitud bastón inf. (fracción de L)", 0.60, "×L", 2, 0.1, 1.0);

            form.Section("Estribos");
            form.Bar("barEst", "Diámetro", 9.525);
            form.Txt("dist", "Distribución (desde cada extremo)", "1@5, 10@10, R@20");
            form.Num("gancho", "Gancho a 135°: extensión (0 = la de Revit)", 7.5, "cm");
            form.Note("Ejemplo: 1@5, 10@10, R@20 (cm)  ó  1@0.05, 10@0.10, R@0.20 (m).");

            form.Section("Opciones");
            form.Chk("cover", "Asignar el recubrimiento a la viga en Revit", true);

            form.Validator = v =>
            {
                string e = Distribution.Check(v.S("dist"));
                if (e != null) return e;
                try { Corte.ParseComerciales(v.S("comerciales")); }
                catch (FormatException ex) { return ex.Message; }
                return null;
            };
            form.Figure = (g, r, v) => Figures.Viga(g, r, v, b, h, L, v.I("modoLong") == 1 ? luces : new List<(double, double)> { (0, L) });
            return form;
        }

        // ------------------------------------------------------------------ Colocación

        public void Apply(Document doc, Element host, FormValues v, View view, RunReport report) =>
            ApplyGroup(doc, new List<Element> { host }, v, view, report);

        public void ApplyGroup(Document doc, List<Element> grupo, FormValues v, View view, RunReport report)
        {
            // Cadena ordenada; el marco de referencia es el de la primera viga (todas tienen la misma sección y eje).
            List<Luz> cadena = Cadenas(grupo).OrderBy(c => c[0].Xa).First();
            if (cadena.Count != grupo.Count)
                throw new InvalidOperationException("Las vigas del grupo no están alineadas ni tienen la misma sección.");
            LocalFrame f = cadena[0].F;
            bool continua = v.I("modoLong") == 1;

            double c = v.Cm("rec");
            if (v.B("cover")) foreach (Luz l in cadena) RebarTools.SetHostCover(doc, l.H, c);

            RebarBarType tSup = v.Bar(doc, "barSup"), tInf = v.Bar(doc, "barInf"), tEst = v.Bar(doc, "barEst");
            double ds = RebarTools.Diameter(tEst);
            double ext = v.Cm("prol");
            double xIni = cadena[0].Xa, xFin = cadena[cadena.Count - 1].Xb;
            double x0 = ext > 0 ? xIni - ext : xIni + c;
            double x1 = ext > 0 ? xFin + ext : xFin - c;
            double innerH = f.SizeZ - 2 * c - 2 * ds;

            // --- Acero corrido superior e inferior ---------------------------------------------------
            double dbS = RebarTools.Diameter(tSup);
            double zTop = f.MaxZ - c - ds - dbS / 2;
            Corrido(doc, cadena, f, v, tSup, v.I("nSup"), x0, x1, zTop, -1, Leg(Empalmes.PataEje(v.Cm("pataSup"), dbS), innerH - dbS),
                ds, c, continua, "zonaSup", view, report, "viga sup. corrido");

            double dbI = RebarTools.Diameter(tInf);
            double zBot = f.MinZ + c + ds + dbI / 2;
            Corrido(doc, cadena, f, v, tInf, v.I("nInf"), x0, x1, zBot, +1, Leg(Empalmes.PataEje(v.Cm("pataInf"), dbI), innerH - dbI),
                ds, c, continua, "zonaInf", view, report, "viga inf. corrido");

            // --- Bastones superiores en cada apoyo (segunda capa) -------------------------------------
            int bsN = v.I("bsN");
            if (bsN > 0)
            {
                RebarBarType tb = v.Bar(doc, "bsBar");
                double db = RebarTools.Diameter(tb);
                double z = zTop - dbS / 2 - Math.Max(Un.Cm(2.5), Math.Max(db, dbS)) - db / 2;
                double leg = Leg(Empalmes.PataEje(v.Cm("pataSup"), db), z - (f.MinZ + c + ds) - db);
                double Len(Luz l) => v.D("bsFr") * (l.Xb - l.Xa);

                if (continua)
                {
                    // Apoyos extremos con pata; apoyos intermedios con un bastón continuo que abraza las dos luces.
                    Luz a = cadena[0], z2 = cadena[cadena.Count - 1];
                    Longitudinal(doc, a.H, f, tb, bsN, x0, a.Xa + Len(a), z, -1, leg, ds, c, view, report, "viga bastón sup.", legEnd: false);
                    Longitudinal(doc, z2.H, f, tb, bsN, z2.Xb - Len(z2), x1, z, -1, leg, ds, c, view, report, "viga bastón sup.", legStart: false);
                    for (int i = 0; i + 1 < cadena.Count; i++)
                        Longitudinal(doc, cadena[i].H, f, tb, bsN, cadena[i].Xb - Len(cadena[i]), cadena[i + 1].Xa + Len(cadena[i + 1]), z, -1, 0,
                            ds, c, view, report, "viga bastón sup. (apoyo intermedio)");
                }
                else
                {
                    Luz l = cadena[0];
                    Longitudinal(doc, l.H, f, tb, bsN, x0, l.Xa + Len(l), z, -1, leg, ds, c, view, report, "viga bastón sup.", legEnd: false);
                    Longitudinal(doc, l.H, f, tb, bsN, l.Xb - Len(l), x1, z, -1, leg, ds, c, view, report, "viga bastón sup.", legStart: false);
                }
            }

            // --- Bastones inferiores al centro de cada luz (segunda capa) -----------------------------
            int biN = v.I("biN");
            if (biN > 0)
            {
                RebarBarType tb = v.Bar(doc, "biBar");
                double db = RebarTools.Diameter(tb);
                double z = zBot + dbI / 2 + Math.Max(Un.Cm(2.5), Math.Max(db, dbI)) + db / 2;
                foreach (Luz l in cadena)
                {
                    double len = v.D("biFr") * (l.Xb - l.Xa), mid = (l.Xa + l.Xb) / 2;
                    Longitudinal(doc, l.H, f, tb, biN, mid - len / 2, mid + len / 2, z, +1, 0, ds, c, view, report, "viga bastón inf.");
                }
            }

            // --- Estribos de cada viga (en su propio marco) ------------------------------------------
            RebarHookType h135 = RebarTools.Hook(doc, 135, RebarStyle.StirrupTie);
            RebarTools.SetHookLength(tEst, h135, v.Cm("gancho"));
            foreach (Luz l in cadena) Estribos(doc, l.H, l.F, v, tEst, h135, c, ds, view, report);
        }

        /// <summary>
        /// Una fila de acero corrido: una barra de punta a punta, o (vigas continuas) los tramos óptimos con los
        /// traslapes dentro de la zona permitida; los tramos impares se bajan un diámetro para quedar junto al anterior.
        /// </summary>
        private static void Corrido(Document doc, List<Luz> cadena, LocalFrame f, FormValues v, RebarBarType t, int n,
            double x0, double x1, double z, int legDir, double leg, double ds, double c, bool continua, string zonaKey,
            View view, RunReport report, string comment)
        {
            if (n <= 0) return;
            double db = RebarTools.Diameter(t);
            double extra = leg > 0 ? leg - Empalmes.Ahorro90(db, t.StandardBendDiameter) : 0;
            double ls = Traslapes.Ls(v, "ls", db, Un.Cm);
            var luces = cadena.Select(l => (l.Xa, l.Xb)).ToList();

            List<(double Ini, double Fin)> tramos;
            if (continua)
            {
                Traslapes.Resultado opt = Traslapes.Viga(v, luces, f.SizeZ, x0, x1, extra, extra, db, zonaKey, Un.Cm);
                if (opt == null)
                    throw new InvalidOperationException($"{comment}: no hay forma de ubicar los traslapes dentro de la zona permitida con las barras " +
                                                        "comerciales indicadas. Amplíe la zona o agregue una longitud comercial mayor.");
                tramos = opt.Tramos;
            }
            else tramos = new List<(double, double)> { (x0, x1) };

            double lcom = Traslapes.Comerciales(v, Un.Cm).Max();
            Element h0 = cadena[0].H;
            for (int k = 0; k < tramos.Count; k++)
            {
                var (ini, fin) = tramos[k];
                bool primero = k == 0, ultimo = k == tramos.Count - 1;
                double desarrollo = fin - ini + (primero ? extra : 0) + (ultimo ? extra : 0);
                if (desarrollo > lcom + Un.Mm(5))
                    report.Warnings.Add($"{h0.Name} [Id {h0.Id}]: {comment} mide {Un.ToCm(desarrollo) / 100:0.00} m y supera la barra comercial de {Un.ToCm(lcom) / 100:0.00} m.");

                Element host = Anfitrion(cadena, (ini + fin) / 2);
                double zk = z + (k % 2 == 1 ? legDir * db : 0);
                string nombre = tramos.Count > 1 ? $"{comment} tramo {k + 1}/{tramos.Count}" : comment;
                Longitudinal(doc, host, f, t, n, ini, fin, zk, legDir, leg, ds, c, view, report, nombre, legStart: primero, legEnd: ultimo);
            }
            if (tramos.Count > 1)
                report.Warnings.Add($"{h0.Name} [Id {h0.Id}]: {comment} en {tramos.Count} tramos de " +
                                    string.Join(" + ", tramos.Select((tr, k) => $"{Un.ToCm(tr.Fin - tr.Ini + (k == 0 ? extra : 0) + (k == tramos.Count - 1 ? extra : 0)) / 100:0.00}")) +
                                    " m; traslapes en " + string.Join(", ", tramos.Skip(1).Select(tr => $"{Un.ToCm(tr.Ini - cadena[0].Xa) / 100:0.00}–{Un.ToCm(tr.Ini + ls - cadena[0].Xa) / 100:0.00} m")) +
                                    " desde el inicio de la viga.");
        }

        /// <summary>Viga de la cadena que contiene la posición indicada (anfitrión de cada tramo).</summary>
        private static Element Anfitrion(List<Luz> cadena, double x)
        {
            foreach (Luz l in cadena) if (x <= l.Xb) return l.H;
            return cadena[cadena.Count - 1].H;
        }

        private static void Estribos(Document doc, Element host, LocalFrame f, FormValues v, RebarBarType tEst, RebarHookType h135,
            double c, double ds, View view, RunReport report)
        {
            List<double> pos = Distribution.Positions(v.S("dist"), f.SizeX);
            double yL = f.MinY + c + ds / 2, yR = f.MaxY - c - ds / 2;
            double zb = f.MinZ + c + ds / 2, zt = f.MaxZ - c - ds / 2;
            if (yR - yL < ds * 2 || zt - zb < ds * 2)
                throw new InvalidOperationException("La sección es demasiado pequeña para el recubrimiento indicado.");

            // Polígono del estribo en el plano local Y-Z (se usa X=Y local, Y=Z local para la verificación).
            var poly = new List<XYZ> { new XYZ(yL, zb, 0), new XYZ(yR, zb, 0), new XYZ(yR, zt, 0), new XYZ(yL, zt, 0) };
            Func<Rebar, bool> inside = r =>
            {
                IList<Curve> cl = r.GetCenterlineCurves(false, false, false, MultiplanarOption.IncludeOnlyPlanarCurves, 0);
                if (cl.Count == 0) return true;
                XYZ a = f.ToLocal(cl[0].GetEndPoint(0)), e = f.ToLocal(cl[cl.Count - 1].GetEndPoint(1));
                return Geo.Contains(poly, a.Y, a.Z) && Geo.Contains(poly, e.Y, e.Z);
            };

            foreach (Distribution.Run run in Distribution.Group(pos))
            {
                double x = f.MinX + run.Start;
                IList<Curve> loop = Geo.Poly(f.P(x, yL, zt), f.P(x, yR, zt), f.P(x, yR, zb), f.P(x, yL, zb), f.P(x, yL, zt));
                Rebar r = RebarTools.Create(doc, host, RebarStyle.StirrupTie, tEst, h135, h135, f.X, loop, inside);
                RebarTools.LayoutFixed(r, run.Count, run.Length, f.X);
                RebarTools.Finish(r, tEst, "viga estribo", view, report);
            }
        }

        private static double Leg(double wanted, double max) => wanted <= 0 ? 0 : Math.Max(0, Math.Min(wanted, max));

        /// <summary>Crea un conjunto de barras longitudinales repartidas en el ancho de la viga.</summary>
        private static void Longitudinal(Document doc, Element host, LocalFrame f, RebarBarType t, int n,
            double x0, double x1, double z, int legDir, double leg, double ds, double c, View view, RunReport report,
            string comment, bool legStart = true, bool legEnd = true)
        {
            if (n <= 0 || x1 - x0 < Un.Cm(5)) return;
            double db = RebarTools.Diameter(t);
            double yA = f.MinY + c + ds + db / 2, yB = f.MaxY - c - ds - db / 2;
            if (yB < yA) throw new InvalidOperationException("La viga es demasiado angosta para el recubrimiento indicado.");
            double y = n == 1 ? (yA + yB) / 2 : yA;

            var pts = new List<XYZ>();
            if (leg > 0 && legStart) pts.Add(f.P(x0, y, z + legDir * leg));
            pts.Add(f.P(x0, y, z));
            pts.Add(f.P(x1, y, z));
            if (leg > 0 && legEnd) pts.Add(f.P(x1, y, z + legDir * leg));

            Rebar r = RebarTools.Create(doc, host, RebarStyle.Standard, t, null, null, f.Y, Geo.Poly(pts.ToArray()));
            RebarTools.LayoutFixed(r, n, yB - yA, f.Y);
            RebarTools.Finish(r, t, comment, view, report);
        }
    }
}
