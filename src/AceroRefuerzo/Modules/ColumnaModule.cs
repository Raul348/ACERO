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
    internal sealed class ColumnaModule : IRebarModule
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

        public ParamForm CreateForm(Document doc, Element sample, IList<BarItem> bars)
        {
            Section s = Read(sample);
            LocalFrame f = s.F;
            List<P2> polyCm = s.Poly.Select(p => new P2(Un.ToCm(p.X), Un.ToCm(p.Y))).ToList();
            double hCm = Un.ToCm(f.SizeZ);
            string forma = s.IsRect ? $"Rectangular {Un.ToCm(f.SizeX):0.#} × {Un.ToCm(f.SizeY):0.#} cm" : $"Irregular ({s.Poly.Count} vértices)";

            var form = new ParamForm("columna", "Acero en COLUMNAS", $"{forma}   ·   Altura {hCm / 100:0.00} m", bars);

            form.Section("Sección");
            form.Choice("modo", "Tipo de armado", new[] { "Automático (detecta la forma)", "Rectangular (barras por cara)", "Contorno irregular / circular" }, 0);
            form.Num("rec", "Recubrimiento libre", 4, "cm");

            form.Section("Acero longitudinal");
            form.Bar("barLong", "Diámetro", 15.875);
            form.Num("nX", "Barras por cara en X (con esquinas)", 3, "u", 0, 2, 20);
            form.Num("nY", "Barras por cara en Y (con esquinas)", 3, "u", 0, 2, 20);
            form.Num("smax", "Irregular: separación máx. entre barras", 15, "cm", 1, 5, 100);
            form.Num("nMin", "Circular: número mínimo de barras", 6, "u", 0, 4, 60);

            form.Section("Anclaje y empalme");
            form.Num("anc", "Anclaje inferior (dentro de la zapata)", 0, "cm");
            form.Num("pata", "Pata inferior a 90° (0 = sin pata)", 0, "cm");
            form.Num("emp", "Empalme superior (sobre el elemento)", 0, "cm");

            form.Section("Estribos");
            form.Bar("barEst", "Diámetro", 9.525);
            form.Txt("dist", "Distribución (desde cada extremo)", "1@5, 8@10, R@25");
            form.Chk("grapas", "Rectangular: grapas en barras intermedias", true);
            form.Note("Ejemplo: 1@5, 8@10, R@25 (cm). Se aplica desde abajo y desde arriba.");

            form.Section("Opciones");
            form.Chk("cover", "Asignar el recubrimiento a la columna en Revit", true);

            form.Figure = (g, r, v) => Figures.Columna(g, r, v, polyCm, v.I("modo") == 2 || (v.I("modo") == 0 && !s.IsRect), hCm);
            return form;
        }

        public void Apply(Document doc, Element host, FormValues v, View view, RunReport report)
        {
            Section s = Read(host);
            double c = v.Cm("rec");
            if (v.B("cover")) RebarTools.SetHostCover(doc, host, c);

            if (UseContour(v, s)) ApplyContour(doc, host, s, v, view, report);
            else ApplyRect(doc, host, s.F, v, view, report);
        }

        private static (double zA, double zB) BarEnds(LocalFrame f, FormValues v, double c)
        {
            double anc = v.Cm("anc"), emp = v.Cm("emp");
            return (anc > 0 ? f.MinZ - anc : f.MinZ + c, emp > 0 ? f.MaxZ + emp : f.MaxZ - c);
        }

        // ------------------------------------------------------------------ Rectangular
        private static void ApplyRect(Document doc, Element host, LocalFrame f, FormValues v, View view, RunReport report)
        {
            double c = v.Cm("rec");
            RebarBarType tL = v.Bar(doc, "barLong"), tE = v.Bar(doc, "barEst");
            double db = RebarTools.Diameter(tL), ds = RebarTools.Diameter(tE);
            int nX = Math.Max(2, v.I("nX")), nY = Math.Max(2, v.I("nY"));

            double xA = f.MinX + c + ds + db / 2, xB = f.MaxX - c - ds - db / 2;
            double yA = f.MinY + c + ds + db / 2, yB = f.MaxY - c - ds - db / 2;
            if (xB <= xA || yB <= yA) throw new InvalidOperationException("La sección es demasiado pequeña para el recubrimiento indicado.");

            var (zA, zB) = BarEnds(f, v, c);
            double leg = Math.Min(v.Cm("pata"), Math.Min(xB - xA, yB - yA));

            // Caras paralelas a X (incluyen las esquinas).
            Vertical(doc, host, f, tL, f.P(xA, yA, zA), f.P(xA, yA, zB), f.Y, leg, f.X, nX, xB - xA, view, report);
            Vertical(doc, host, f, tL, f.P(xA, yB, zA), f.P(xA, yB, zB), -f.Y, leg, f.X, nX, xB - xA, view, report);

            // Caras paralelas a Y (solo barras intermedias).
            if (nY > 2)
            {
                double dy = (yB - yA) / (nY - 1);
                Vertical(doc, host, f, tL, f.P(xA, yA + dy, zA), f.P(xA, yA + dy, zB), f.X, leg, f.Y, nY - 2, dy * (nY - 3), view, report);
                Vertical(doc, host, f, tL, f.P(xB, yA + dy, zA), f.P(xB, yA + dy, zB), -f.X, leg, f.Y, nY - 2, dy * (nY - 3), view, report);
            }

            // Estribos y grapas.
            double xL = f.MinX + c + ds / 2, xR = f.MaxX - c - ds / 2;
            double yL = f.MinY + c + ds / 2, yR = f.MaxY - c - ds / 2;
            RebarHookType h135 = RebarTools.Hook(doc, 135, RebarStyle.StirrupTie);
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

        /// <summary>Barra vertical con pata opcional; se reparte <paramref name="count"/> veces en <paramref name="dir"/>.</summary>
        private static void Vertical(Document doc, Element host, LocalFrame f, RebarBarType t, XYZ bottom, XYZ top,
            XYZ legDir, double leg, XYZ dir, int count, double length, View view, RunReport report)
        {
            var pts = new List<XYZ>();
            if (leg > Un.Mm(5)) pts.Add(bottom + legDir * leg);
            pts.Add(bottom);
            pts.Add(top);
            Rebar r = RebarTools.Create(doc, host, RebarStyle.Standard, t, null, null, dir, Geo.Poly(pts.ToArray()));
            RebarTools.LayoutFixed(r, count, length, dir);
            RebarTools.Finish(r, t, "columna longitudinal", view, report);
        }

        // ------------------------------------------------------------------ Irregular / circular
        private static void ApplyContour(Document doc, Element host, Section s, FormValues v, View view, RunReport report)
        {
            LocalFrame f = s.F;
            double c = v.Cm("rec");
            RebarBarType tL = v.Bar(doc, "barLong"), tE = v.Bar(doc, "barEst");
            double db = RebarTools.Diameter(tL), ds = RebarTools.Diameter(tE);
            double area0 = Polygon2D.Area(s.Poly);

            List<P2> tiePoly = Polygon2D.OffsetInward(s.Poly, c + ds / 2);
            List<P2> barPoly = Polygon2D.OffsetInward(s.Poly, c + ds + db / 2);
            if (Polygon2D.Area(barPoly) <= 0 || Polygon2D.Area(barPoly) >= area0)
                throw new InvalidOperationException("La sección es demasiado pequeña para el recubrimiento indicado.");

            // Barras longitudinales.
            List<P2> pts = Polygon2D.PerimeterPoints(barPoly, v.Cm("smax"), Math.Max(4, v.I("nMin")));
            P2 cen = Polygon2D.Centroid(barPoly);
            var (zA, zB) = BarEnds(f, v, c);
            double legWanted = v.Cm("pata");

            foreach (P2 p in pts)
            {
                P2 toC = cen - p;
                XYZ legDir = (f.X * toC.X + f.Y * toC.Y);
                double leg = Math.Min(legWanted, toC.Length);
                var curve = new List<XYZ>();
                XYZ normal = f.X;
                if (leg > Un.Mm(5) && legDir.GetLength() > 1e-9)
                {
                    legDir = legDir.Normalize();
                    curve.Add(f.P(p.X, p.Y, zA) + legDir * leg);
                    normal = legDir.CrossProduct(XYZ.BasisZ).Normalize();
                }
                curve.Add(f.P(p.X, p.Y, zA));
                curve.Add(f.P(p.X, p.Y, zB));
                Rebar r = RebarTools.Create(doc, host, RebarStyle.Standard, tL, null, null, normal, Geo.Poly(curve.ToArray()));
                RebarTools.Finish(r, tL, "columna longitudinal", view, report);
            }

            // Estribo que sigue el contorno (arcos reales cuando Revit puede desfasar el contorno).
            CurveLoop tieLoop = OffsetLoop(s.Outline, c + ds / 2);
            RebarHookType h135 = RebarTools.Hook(doc, 135, RebarStyle.StirrupTie);
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
