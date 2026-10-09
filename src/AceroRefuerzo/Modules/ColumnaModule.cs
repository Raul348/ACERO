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
    /// Columnas rectangulares e irregulares (L, T, cruz, circulares, poligonales).
    ///  · Rectangular: número de barras por cara, estribo perimetral y grapas en barras intermedias.
    ///  · Irregular / circular: se lee el contorno real de la sección; barras en cada esquina y en cada lado
    ///    con separación máxima, estribo cerrado que sigue el contorno.
    /// Estribos con ganchos de 135° distribuidos por zonas (confinamiento arriba y abajo).
    /// </summary>
    internal sealed class ColumnaModule : IRebarModule, IGroupModule
    {
        public ElementKind Kind => ElementKind.Columna;
        public string Title => "Columnas";
        public string Prompt => "Seleccione las columnas a armar y pulse Finalizar";

        public bool Accepts(Element e) =>
            e is FamilyInstance && e.Category?.BuiltInCategory == BuiltInCategory.OST_StructuralColumns;

        /// <summary>Sección horizontal de la columna.</summary>
        private sealed class Section
        {
            public LocalFrame F;
            public CurveLoop Outline;   // contorno real (coordenadas globales) en la cara inferior
            public double OutlineZ;     // Z local de esa cara
            public List<P2> Poly;       // contorno en coordenadas locales (pies)
            public bool IsRect;
        }

        private static Section Read(Element e)
        {
            var fi = (FamilyInstance)e;
            Transform t = fi.GetTransform();
            XYZ x = Geo.Horizontal(t.BasisX) ?? XYZ.BasisX;
            XYZ y = XYZ.BasisZ.CrossProduct(x).Normalize();
            LocalFrame f = LocalFrame.Measure(e, t.Origin, x, y, XYZ.BasisZ);
            var sec = new Section { F = f };

            // Cara inferior horizontal más baja → contorno exterior (el de mayor área).
            PlanarFace bottom = null;
            double bottomZ = double.MaxValue;
            foreach (Solid s in Geo.Solids(e))
                foreach (Face face in s.Faces)
                    if (face is PlanarFace pf && pf.FaceNormal.DotProduct(XYZ.BasisZ) < -0.99)
                    {
                        double z = f.ToLocal(pf.Origin).Z;
                        if (z < bottomZ - 1e-6) { bottomZ = z; bottom = pf; }
                    }

            if (bottom != null)
            {
                double best = 0;
                foreach (CurveLoop loop in bottom.GetEdgesAsCurveLoops())
                {
                    List<P2> pts = Tessellate(loop, f);
                    double a = Math.Abs(Polygon2D.Area(pts));
                    if (a > best) { best = a; sec.Poly = pts; sec.Outline = loop; }
                }
                sec.OutlineZ = bottomZ;
            }

            if (sec.Poly == null || sec.Poly.Count < 3)
            {
                sec.Poly = new List<P2> { new P2(f.MinX, f.MinY), new P2(f.MaxX, f.MinY), new P2(f.MaxX, f.MaxY), new P2(f.MinX, f.MaxY) };
                sec.Outline = null;
                sec.OutlineZ = f.MinZ;
            }

            sec.Poly = Polygon2D.Clean(sec.Poly, Un.Mm(1));
            if (Polygon2D.Area(sec.Poly) < 0) sec.Poly.Reverse();
            sec.IsRect = sec.Poly.Count == 4 && Math.Abs(Polygon2D.Area(sec.Poly) - f.SizeX * f.SizeY) < 0.02 * f.SizeX * f.SizeY;
            return sec;
        }

        private static List<P2> Tessellate(CurveLoop loop, LocalFrame f)
        {
            var pts = new List<P2>();
            foreach (Curve c in loop)
            {
                IList<XYZ> tp = c.Tessellate();
                for (int i = 0; i < tp.Count - 1; i++)
                {
                    XYZ l = f.ToLocal(tp[i]);
                    pts.Add(new P2(l.X, l.Y));
                }
            }
            return pts;
        }

        private static bool UseContour(FormValues v, Section s) =>
            v.I("modo") == 2 || (v.I("modo") == 0 && !s.IsRect);

        public ParamForm CreateForm(Document doc, IList<Element> hosts, IList<BarItem> bars)
        {
            Element sample = hosts[0];
            Section s = Read(sample);
            LocalFrame f = s.F;
            List<P2> polyCm = s.Poly.Select(p => new P2(Un.ToCm(p.X), Un.ToCm(p.Y))).ToList();
            double hCm = Un.ToCm(f.SizeZ);
            string forma = s.IsRect ? $"Rectangular {Un.ToCm(f.SizeX):0.#} × {Un.ToCm(f.SizeY):0.#} cm" : $"Irregular ({s.Poly.Count} vértices)";

            // Pisos de la columna seleccionada (si se eligieron columnas de varios pisos del mismo eje).
            List<Element> pila = Pilas(hosts).First(g => g.Contains(sample));
            List<Section> secs = pila.Select(Read).ToList();
            double z0 = Abs(secs[0], secs[0].F.MinZ);
            List<(double zb, double zt)> pisosCm = secs.Select(x => (Un.ToCm(Abs(x, x.F.MinZ) - z0), Un.ToCm(Abs(x, x.F.MaxZ) - z0))).ToList();
            string pisos = pila.Count > 1 ? $"   ·   {pila.Count} pisos en el eje" : "";

            var form = new ParamForm(App.Name, "columna", "Acero en COLUMNAS", $"{forma}   ·   Altura {hCm / 100:0.00} m{pisos}", bars);

            form.Section("Sección");
            form.Choice("modo", "Tipo de armado", new[] { "Automático (detecta la forma)", "Rectangular (barras por cara)", "Contorno irregular / circular" }, 0);
            form.Num("rec", "Recubrimiento libre", 4, "cm");

            form.Section("Acero longitudinal");
            form.Bar("barLong", "Diámetro", 15.875);
            form.Num("nX", "Barras por cara en X (con esquinas)", 3, "u", 0, 2, 20);
            form.Num("nY", "Barras por cara en Y (con esquinas)", 3, "u", 0, 2, 20);
            form.Num("smax", "Irregular: separación máx. entre barras", 15, "cm", 1, 5, 100);
            form.Num("nMin", "Circular: número mínimo de barras", 6, "u", 0, 4, 60);

            form.Section("Longitud de las barras longitudinales");
            form.Num("anc", "Anclaje inferior (dentro de la zapata)", 0, "cm");
            form.Num("pata", "Pata inferior a 90°, medida por fuera (0 = sin pata)", 0, "cm");
            form.Choice("modoLong", "Longitud de las barras", new[]
            {
                "Automática: del anclaje al empalme superior",
                "Tramos que yo indico (columnas de varios pisos)"
            }, 0);
            form.Num("emp", "Automática: empalme superior (sobre el elemento)", 0, "cm");
            form.Txt("tramos", "Tramos: longitudes de abajo hacia arriba (m)", "9");
            form.Num("ls", "Tramos: traslape entre tramos", 60, "cm");
            form.Num("lcom", "Longitud comercial de la barra", 9, "m", 2, 1, 30);
            form.Note("Con \"Tramos que yo indico\", seleccione las columnas de todos los pisos del eje: se unen y las barras se " +
                      "colocan de abajo hacia arriba con la longitud exacta escrita (incluye la pata y el doblez del primer tramo). " +
                      "Cada tramo empieza un traslape antes del final del anterior; el traslape queda donde usted lo decida.");

            form.Section("Estribos");
            form.Bar("barEst", "Diámetro", 9.525);
            form.Txt("dist", "Distribución (desde cada extremo)", "1@5, 8@10, R@25");
            form.Num("gancho", "Gancho a 135°: extensión (0 = la de Revit)", 7.5, "cm");
            form.Chk("grapas", "Rectangular: grapas en barras intermedias", true);
            form.Choice("tipoEst", "Irregular: estribos", new[] { "Uno por cada ala (L, T, cruz)", "Uno perimetral (sigue el contorno)" }, 0);
            form.Note("Ejemplo: 1@5, 8@10, R@25 (cm). Se aplica desde abajo y desde arriba de cada piso.");
            form.Note("El gancho se guarda en Revit, en la tabla \"Longitudes de gancho\" del tipo de barra del estribo.");
            form.Note("\"Uno por cada ala\": en secciones L, T o cruz coloca un estribo rectangular por ala, traslapados en la unión. " +
                      "En secciones circulares o con lados inclinados se usa el perimetral.");

            form.Section("Opciones");
            form.Chk("cover", "Asignar el recubrimiento a la columna en Revit", true);

            form.Validator = v =>
            {
                string e = Distribution.Check(v.S("dist"));
                if (e != null || v.I("modoLong") != 1) return e;
                try { return Empalmes.LeerTramos(v.S("tramos")) == null ? "Escriba la longitud de los tramos, por ejemplo: 9, 9, 4.5" : null; }
                catch (FormatException ex) { return ex.Message; }
            };
            form.Figure = (g, r, v) => Figures.Columna(g, r, v, polyCm, v.I("modo") == 2 || (v.I("modo") == 0 && !s.IsRect),
                v.I("modoLong") == 1 ? pisosCm : new List<(double, double)> { (0, hCm) });
            return form;
        }

        public void Apply(Document doc, Element host, FormValues v, View view, RunReport report) =>
            ApplyGroup(doc, new List<Element> { host }, v, view, report);

        // ------------------------------------------------------------------ Columnas de varios pisos

        /// <summary>
        /// Con "tramos indicados", agrupa las columnas apiladas de un mismo eje (misma sección y posición en planta,
        /// una sobre otra) para armarlas con barras continuas de las longitudes que indique el usuario.
        /// Cada grupo va ordenado de abajo hacia arriba. En modo automático cada columna va sola.
        /// </summary>
        public List<List<Element>> Agrupar(IList<Element> hosts, FormValues v) =>
            v != null && v.I("modoLong") == 1 ? Pilas(hosts) : hosts.Select(h => new List<Element> { h }).ToList();

        private static List<List<Element>> Pilas(IList<Element> hosts)
        {
            var datos = hosts.Select(h => (h, s: Read(h))).ToList();
            var grupos = new List<List<(Element h, Section s)>>();
            foreach (var d in datos.OrderBy(x => Abs(x.s, x.s.F.MinZ)))
            {
                List<(Element h, Section s)> destino = null;
                foreach (var g in grupos)
                {
                    var ultimo = g[g.Count - 1];
                    if (MismoEje(ultimo.s, d.s) && Abs(d.s, d.s.F.MinZ) - Abs(ultimo.s, ultimo.s.F.MaxZ) < Un.Cm(150)
                        && Abs(d.s, d.s.F.MinZ) > Abs(ultimo.s, ultimo.s.F.MinZ) + Un.Cm(10))
                    {
                        destino = g;
                        break;
                    }
                }
                if (destino == null) grupos.Add(destino = new List<(Element, Section)>());
                destino.Add(d);
            }
            return grupos.Select(g => g.Select(x => x.h).ToList()).ToList();
        }

        /// <summary>Cota absoluta (Z del proyecto) de una cota local del marco de la sección.</summary>
        private static double Abs(Section s, double zLocal) => s.F.Origin.Z + zLocal;

        private static bool MismoEje(Section a, Section b)
        {
            P2 ca = Polygon2D.Centroid(a.Poly), cb = Polygon2D.Centroid(b.Poly);
            XYZ wa = a.F.P(ca.X, ca.Y, 0), wb = b.F.P(cb.X, cb.Y, 0);
            double dxy = Math.Sqrt(Math.Pow(wa.X - wb.X, 2) + Math.Pow(wa.Y - wb.Y, 2));
            double aa = Math.Abs(Polygon2D.Area(a.Poly)), ab = Math.Abs(Polygon2D.Area(b.Poly));
            return dxy < Un.Cm(3) && Math.Abs(aa - ab) < 0.02 * Math.Max(aa, ab) && Math.Abs(a.F.X.DotProduct(b.F.X)) > 0.999;
        }

        /// <summary>Una "línea" de barras longitudinales: una barra o un conjunto repartido en una cara.</summary>
        private sealed class Linea
        {
            public P2 P;          // posición de la primera barra (local de la columna inferior)
            public P2 Dir;        // dirección de reparto (si Count > 1)
            public int Count = 1;
            public double Largo;  // longitud de reparto
            public P2 Adentro;    // hacia el interior de la sección (pata y desplazamiento del traslape)
        }

        /// <summary>
        /// Arma un grupo de columnas apiladas: barras longitudinales de punta a punta (automático) o en los tramos que
        /// indique el usuario, unidos con el traslape indicado; estribos y grapas por piso.
        /// </summary>
        public void ApplyGroup(Document doc, List<Element> grupo, FormValues v, View view, RunReport report)
        {
            List<Section> secs = grupo.Select(Read).ToList();
            Section s0 = secs[0];
            LocalFrame f = s0.F;
            double c = v.Cm("rec");
            if (v.B("cover")) foreach (Element h in grupo) RebarTools.SetHostCover(doc, h, c);

            RebarBarType tL = v.Bar(doc, "barLong"), tE = v.Bar(doc, "barEst");
            double db = RebarTools.Diameter(tL), ds = RebarTools.Diameter(tE);
            RebarHookType h135 = RebarTools.Hook(doc, 135, RebarStyle.StirrupTie);
            RebarTools.SetHookLength(tE, h135, v.Cm("gancho"));

            bool contorno = UseContour(v, s0);
            List<Linea> lineas = contorno ? LineasContorno(s0, v, c, db, ds, out var alas) : LineasRect(f, v, c, db, ds);

            // Cotas de cada piso en el marco de la columna inferior.
            var pisos = secs.Select(x => (zb: Abs(x, x.F.MinZ) - f.Origin.Z, zt: Abs(x, x.F.MaxZ) - f.Origin.Z)).ToList();
            double anc = v.Cm("anc"), emp = v.Cm("emp");
            double zIni = anc > 0 ? pisos[0].zb - anc : pisos[0].zb + c;
            double zFin = emp > 0 ? pisos[pisos.Count - 1].zt + emp : pisos[pisos.Count - 1].zt - c;

            // Tramos: automático (una barra de punta a punta) o las longitudes que indique el usuario.
            double pataEje = Empalmes.PataEje(v.Cm("pata"), db);
            double extra = pataEje > 0 ? pataEje - Empalmes.Ahorro90(db, tL.StandardBendDiameter) : 0;
            List<double> largos = v.I("modoLong") == 1 ? Empalmes.LeerTramos(v.S("tramos")) : null;
            List<(double Ini, double Fin)> tramos = largos != null
                ? Empalmes.UbicarTramos(zIni, largos.Select(m => Un.Cm(m * 100)).ToList(), extra, v.Cm("ls"))
                : new List<(double, double)> { (zIni, zFin) };

            double lcom = Un.Cm(v.D("lcom") * 100);
            for (int k = 0; k < tramos.Count; k++)
            {
                double desarrollo = tramos[k].Fin - tramos[k].Ini + (k == 0 ? extra : 0);
                if (desarrollo > lcom + Un.Mm(5))
                    report.Warnings.Add($"{grupo[0].Name} [Id {grupo[0].Id}]: el tramo {k + 1} mide {Un.ToCm(desarrollo) / 100:0.00} m y supera la barra comercial de {v.D("lcom"):0.00} m.");
            }
            if (largos != null && tramos[tramos.Count - 1].Fin < zFin - Un.Mm(5))
                report.Warnings.Add($"{grupo[0].Name} [Id {grupo[0].Id}]: los tramos indicados no llegan al tope; faltan {Un.ToCm(zFin - tramos[tramos.Count - 1].Fin) / 100:0.00} m.");

            for (int k = 0; k < tramos.Count; k++)
            {
                var (ini, fin) = tramos[k];
                Element host = Anfitrion(grupo, pisos, (ini + fin) / 2);
                double desplazar = k % 2 == 1 ? db : 0; // las barras que traslapan van una al lado de la otra
                string nombre = tramos.Count > 1 ? $"columna longitudinal tramo {k + 1}/{tramos.Count}" : "columna longitudinal";

                foreach (Linea ln in lineas)
                {
                    P2 p = ln.P + ln.Adentro * desplazar;
                    XYZ adentro = f.X * ln.Adentro.X + f.Y * ln.Adentro.Y;
                    XYZ dir = f.X * ln.Dir.X + f.Y * ln.Dir.Y;
                    var pts = new List<XYZ>();
                    double pata = k == 0 ? Math.Min(pataEje, Un.Cm(150)) : 0;
                    if (pata > Un.Mm(5)) pts.Add(f.P(p.X, p.Y, ini) + adentro * pata);
                    pts.Add(f.P(p.X, p.Y, ini));
                    pts.Add(f.P(p.X, p.Y, fin));

                    XYZ normal = ln.Count > 1 ? dir : adentro.CrossProduct(XYZ.BasisZ);
                    if (normal.GetLength() < 1e-9) normal = f.X;
                    Rebar r = RebarTools.Create(doc, host, RebarStyle.Standard, tL, null, null, normal.Normalize(), Geo.Poly(pts.ToArray()));
                    RebarTools.LayoutFixed(r, ln.Count, ln.Largo, dir);
                    RebarTools.Finish(r, tL, nombre, view, report);
                }
            }

            // Estribos (y grapas) de cada piso.
            for (int i = 0; i < grupo.Count; i++)
            {
                if (contorno) EstribosContorno(doc, grupo[i], secs[i], v, tE, h135, c, db, ds, view, report);
                else EstribosRect(doc, grupo[i], secs[i].F, v, tE, h135, c, db, ds, view, report);
            }
        }

        /// <summary>Columna del grupo que contiene la cota indicada (anfitrión de cada tramo de barra).</summary>
        private static Element Anfitrion(List<Element> grupo, List<(double zb, double zt)> pisos, double z)
        {
            for (int i = 0; i < pisos.Count; i++)
                if (z <= pisos[i].zt) return grupo[i];
            return grupo[grupo.Count - 1];
        }

        // ------------------------------------------------------------------ Rectangular

        private static List<Linea> LineasRect(LocalFrame f, FormValues v, double c, double db, double ds)
        {
            int nX = Math.Max(2, v.I("nX")), nY = Math.Max(2, v.I("nY"));
            double xA = f.MinX + c + ds + db / 2, xB = f.MaxX - c - ds - db / 2;
            double yA = f.MinY + c + ds + db / 2, yB = f.MaxY - c - ds - db / 2;
            if (xB <= xA || yB <= yA) throw new InvalidOperationException("La sección es demasiado pequeña para el recubrimiento indicado.");

            var l = new List<Linea>
            {
                // Caras paralelas a X (incluyen las esquinas).
                new Linea { P = new P2(xA, yA), Dir = new P2(1, 0), Count = nX, Largo = xB - xA, Adentro = new P2(0, 1) },
                new Linea { P = new P2(xA, yB), Dir = new P2(1, 0), Count = nX, Largo = xB - xA, Adentro = new P2(0, -1) },
            };
            if (nY > 2)
            {
                // Caras paralelas a Y (solo barras intermedias).
                double dy = (yB - yA) / (nY - 1);
                l.Add(new Linea { P = new P2(xA, yA + dy), Dir = new P2(0, 1), Count = nY - 2, Largo = dy * (nY - 3), Adentro = new P2(1, 0) });
                l.Add(new Linea { P = new P2(xB, yA + dy), Dir = new P2(0, 1), Count = nY - 2, Largo = dy * (nY - 3), Adentro = new P2(-1, 0) });
            }
            return l;
        }

        private static void EstribosRect(Document doc, Element host, LocalFrame f, FormValues v, RebarBarType tE,
            RebarHookType h135, double c, double db, double ds, View view, RunReport report)
        {
            int nX = Math.Max(2, v.I("nX")), nY = Math.Max(2, v.I("nY"));
            double xA = f.MinX + c + ds + db / 2, xB = f.MaxX - c - ds - db / 2;
            double yA = f.MinY + c + ds + db / 2, yB = f.MaxY - c - ds - db / 2;
            double xL = f.MinX + c + ds / 2, xR = f.MaxX - c - ds / 2;
            double yL = f.MinY + c + ds / 2, yR = f.MaxY - c - ds / 2;
            var poly = new List<XYZ> { new XYZ(xL, yL, 0), new XYZ(xR, yL, 0), new XYZ(xR, yR, 0), new XYZ(xL, yR, 0) };
            Func<Rebar, bool> inside = RebarTools.TipsInside(f, poly);

            var xs = Enumerable.Range(0, nX).Select(i => xA + (xB - xA) * i / (nX - 1)).ToList();
            var ys = Enumerable.Range(0, nY).Select(i => yA + (yB - yA) * i / (nY - 1)).ToList();
            double off = db / 2 + ds / 2;

            foreach (Distribution.Run run in Distribution.Group(Distribution.Positions(v.S("dist"), f.SizeZ)))
            {
                double z = f.MinZ + run.Start;
                IList<Curve> loop = Geo.Poly(f.P(xL, yR, z), f.P(xR, yR, z), f.P(xR, yL, z), f.P(xL, yL, z), f.P(xL, yR, z));
                Rebar r = RebarTools.Create(doc, host, RebarStyle.StirrupTie, tE, h135, h135, f.Z, loop, inside);
                RebarTools.LayoutFixed(r, run.Count, run.Length, f.Z);
                RebarTools.Finish(r, tE, "columna estribo", view, report);

                if (!v.B("grapas")) continue;
                double zg = z + ds;
                for (int i = 1; i < nX - 1; i++)
                    Grapa(doc, host, f, tE, h135, f.P(xs[i] + off, yL, zg), f.P(xs[i] + off, yR, zg), run, view, report);
                for (int j = 1; j < nY - 1; j++)
                    Grapa(doc, host, f, tE, h135, f.P(xL, ys[j] + off, zg), f.P(xR, ys[j] + off, zg), run, view, report);
            }
        }

        private static void Grapa(Document doc, Element host, LocalFrame f, RebarBarType t, RebarHookType hook,
            XYZ a, XYZ b, Distribution.Run run, View view, RunReport report)
        {
            Rebar r = RebarTools.Create(doc, host, RebarStyle.StirrupTie, t, hook, hook, f.Z, Geo.Poly(a, b));
            RebarTools.LayoutFixed(r, run.Count, run.Length, f.Z);
            RebarTools.Finish(r, t, "columna grapa", view, report);
        }

        // ------------------------------------------------------------------ Irregular / circular

        private static List<Linea> LineasContorno(Section s, FormValues v, double c, double db, double ds,
            out List<(double x0, double y0, double x1, double y1)> alas)
        {
            double area0 = Polygon2D.Area(s.Poly);
            List<P2> barPoly = Polygon2D.OffsetInward(s.Poly, c + ds + db / 2);
            if (Polygon2D.Area(barPoly) <= 0 || Polygon2D.Area(barPoly) >= area0)
                throw new InvalidOperationException("La sección es demasiado pequeña para el recubrimiento indicado.");

            alas = Alas(s, v);
            List<P2> pts = Polygon2D.PerimeterPoints(barPoly, v.Cm("smax"), Math.Max(4, v.I("nMin")),
                alas != null ? Polygon2D.RectCorners(alas, c + ds + db / 2) : null);
            P2 cen = Polygon2D.Centroid(barPoly);
            return pts.Select(p => new Linea { P = p, Dir = new P2(1, 0), Adentro = (cen - p).Unit }).ToList();
        }

        /// <summary>Estribos por ala (secciones ortogonales: L, T, cruz) o null si va un estribo perimetral.</summary>
        private static List<(double x0, double y0, double x1, double y1)> Alas(Section s, FormValues v)
        {
            var alas = v.I("tipoEst") == 0 ? Polygon2D.CoverRectangles(s.Poly) : null;
            return alas != null && alas.Count >= 2 ? alas : null;
        }

        private static void EstribosContorno(Document doc, Element host, Section s, FormValues v, RebarBarType tE,
            RebarHookType h135, double c, double db, double ds, View view, RunReport report)
        {
            LocalFrame f = s.F;
            var alas = Alas(s, v);
            if (alas != null)
            {
                EstribosPorAla(doc, host, f, v, tE, h135, alas, c + ds / 2, ds, view, report);
                return;
            }

            // Estribo que sigue el contorno (arcos reales cuando Revit puede desfasar el contorno).
            List<P2> tiePoly = Polygon2D.OffsetInward(s.Poly, c + ds / 2);
            CurveLoop tieLoop = OffsetLoop(s.Outline, c + ds / 2);
            Func<Rebar, bool> inside = RebarTools.TipsInside(f, tiePoly.Select(q => new XYZ(q.X, q.Y, 0)).ToList());

            foreach (Distribution.Run run in Distribution.Group(Distribution.Positions(v.S("dist"), f.SizeZ)))
            {
                double z = f.MinZ + run.Start;
                IList<Curve> loop;
                if (tieLoop != null)
                {
                    Transform move = Transform.CreateTranslation(f.Z * (z - s.OutlineZ));
                    loop = tieLoop.Select(cv => cv.CreateTransformed(move)).ToList();
                }
                else
                {
                    var corners = tiePoly.Select(q => f.P(q.X, q.Y, z)).ToList();
                    corners.Add(corners[0]);
                    loop = Geo.Poly(corners.ToArray());
                }
                Rebar r = RebarTools.Create(doc, host, RebarStyle.StirrupTie, tE, h135, h135, f.Z, loop, inside);
                RebarTools.LayoutFixed(r, run.Count, run.Length, f.Z);
                RebarTools.Finish(r, tE, "columna estribo", view, report);
            }
        }

        /// <summary>
        /// Un estribo rectangular cerrado por cada ala de la sección. En cada nivel los estribos se separan
        /// un diámetro en altura para que no ocupen el mismo lugar donde se traslapan.
        /// </summary>
        private static void EstribosPorAla(Document doc, Element host, LocalFrame f, FormValues v, RebarBarType tE,
            RebarHookType h135, List<(double x0, double y0, double x1, double y1)> alas, double inset, double ds,
            View view, RunReport report)
        {
            foreach (Distribution.Run run in Distribution.Group(Distribution.Positions(v.S("dist"), f.SizeZ)))
            {
                for (int k = 0; k < alas.Count; k++)
                {
                    var (x0, y0, x1, y1) = alas[k];
                    double xL = x0 + inset, xR = x1 - inset, yB = y0 + inset, yT = y1 - inset;
                    if (xR - xL < ds * 2 || yT - yB < ds * 2) continue;
                    double z = f.MinZ + run.Start + k * ds;
                    var rect = new List<XYZ> { new XYZ(xL, yB, 0), new XYZ(xR, yB, 0), new XYZ(xR, yT, 0), new XYZ(xL, yT, 0) };
                    IList<Curve> loop = Geo.Poly(f.P(xL, yT, z), f.P(xR, yT, z), f.P(xR, yB, z), f.P(xL, yB, z), f.P(xL, yT, z));
                    Rebar r = RebarTools.Create(doc, host, RebarStyle.StirrupTie, tE, h135, h135, f.Z, loop, RebarTools.TipsInside(f, rect));
                    RebarTools.LayoutFixed(r, run.Count, run.Length, f.Z);
                    RebarTools.Finish(r, tE, $"columna estribo ala {k + 1}", view, report);
                }
            }
        }

        /// <summary>Desfasa el contorno hacia adentro con la API de Revit (null si no es posible).</summary>
        private static CurveLoop OffsetLoop(CurveLoop outline, double d)
        {
            if (outline == null) return null;
            double area0 = LoopArea(outline);
            foreach (double sign in new[] { 1.0, -1.0 })
            {
                try
                {
                    CurveLoop off = CurveLoop.CreateViaOffset(outline, sign * d, XYZ.BasisZ);
                    if (off != null && LoopArea(off) < area0) return off;
                }
                catch (Exception) { }
            }
            return null;
        }

        private static double LoopArea(CurveLoop loop)
        {
            var pts = new List<XYZ>();
            foreach (Curve c in loop)
            {
                IList<XYZ> tp = c.Tessellate();
                for (int i = 0; i < tp.Count - 1; i++) pts.Add(tp[i]);
            }
            return Math.Abs(Geo.SignedArea(pts));
        }
    }
}
