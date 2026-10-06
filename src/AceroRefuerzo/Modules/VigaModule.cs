using System;
using System.Collections.Generic;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace AceroRefuerzo.Modules
{
    /// <summary>
    /// Vigas: acero corrido superior e inferior (con patas en los apoyos), bastones superiores en apoyos,
    /// bastones inferiores al centro y estribos cerrados con ganchos de 135° distribuidos por zonas.
    /// </summary>
    internal sealed class VigaModule : IRebarModule
    {
        public ElementKind Kind => ElementKind.Viga;
        public string Title => "Vigas";
        public string Prompt => "Seleccione las vigas a armar y pulse Finalizar";

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

        public ParamForm CreateForm(Document doc, Element sample, IList<BarItem> bars)
        {
            LocalFrame f = Frame(sample);
            double b = Un.ToCm(f.SizeY), h = Un.ToCm(f.SizeZ), L = Un.ToCm(f.SizeX);

            var form = new ParamForm("viga", "Acero en VIGAS",
                $"Sección {b:0.#} × {h:0.#} cm   ·   Luz libre {L / 100:0.00} m", bars);

            form.Section("Recubrimiento");
            form.Num("rec", "Recubrimiento libre", 4, "cm");

            form.Section("Acero corrido superior");
            form.Bar("barSup", "Diámetro", 15.875);
            form.Num("nSup", "Número de barras", 2, "u", 0, 1, 20);
            form.Num("pataSup", "Pata en apoyos (0 = recta)", 30, "cm");

            form.Section("Acero corrido inferior");
            form.Bar("barInf", "Diámetro", 15.875);
            form.Num("nInf", "Número de barras", 3, "u", 0, 1, 20);
            form.Num("pataInf", "Pata en apoyos (0 = recta)", 30, "cm");
            form.Num("prol", "Prolongación dentro de apoyos", 0, "cm");

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
            form.Note("Ejemplo: 1@5, 10@10, R@20 (cm)  ó  1@0.05, 10@0.10, R@0.20 (m).");

            form.Section("Opciones");
            form.Chk("cover", "Asignar el recubrimiento a la viga en Revit", true);

            form.Figure = (g, r, v) => Figures.Viga(g, r, v, b, h, L);
            return form;
        }

        public void Apply(Document doc, Element host, FormValues v, View view, RunReport report)
        {
            LocalFrame f = Frame(host);
            double c = v.Cm("rec");
            if (v.B("cover")) RebarTools.SetHostCover(doc, host, c);

            RebarBarType tSup = v.Bar(doc, "barSup"), tInf = v.Bar(doc, "barInf"), tEst = v.Bar(doc, "barEst");
            double ds = RebarTools.Diameter(tEst);
            double ext = v.Cm("prol");
            double x0 = ext > 0 ? f.MinX - ext : f.MinX + c;
            double x1 = ext > 0 ? f.MaxX + ext : f.MaxX - c;
            double innerH = f.SizeZ - 2 * c - 2 * ds;

            // --- Acero corrido superior e inferior ---------------------------------------------------
            double dbS = RebarTools.Diameter(tSup);
            double zTop = f.MaxZ - c - ds - dbS / 2;
            Longitudinal(doc, host, f, tSup, v.I("nSup"), x0, x1, zTop, -1, Leg(v.Cm("pataSup"), innerH - dbS), ds, c, view, report, "viga sup. corrido");

            double dbI = RebarTools.Diameter(tInf);
            double zBot = f.MinZ + c + ds + dbI / 2;
            Longitudinal(doc, host, f, tInf, v.I("nInf"), x0, x1, zBot, +1, Leg(v.Cm("pataInf"), innerH - dbI), ds, c, view, report, "viga inf. corrido");

            // --- Bastones superiores en cada apoyo (segunda capa) -------------------------------------
            int bsN = v.I("bsN");
            if (bsN > 0)
            {
                RebarBarType tb = v.Bar(doc, "bsBar");
                double db = RebarTools.Diameter(tb);
                double z = zTop - dbS / 2 - Math.Max(Un.Cm(2.5), Math.Max(db, dbS)) - db / 2;
                double len = v.D("bsFr") * f.SizeX;
                double leg = Leg(v.Cm("pataSup"), z - (f.MinZ + c + ds) - db);
                Longitudinal(doc, host, f, tb, bsN, x0, f.MinX + len, z, -1, leg, ds, c, view, report, "viga bastón sup.", legEnd: false);
                Longitudinal(doc, host, f, tb, bsN, f.MaxX - len, x1, z, -1, leg, ds, c, view, report, "viga bastón sup.", legStart: false);
            }

            // --- Bastones inferiores al centro (segunda capa) -----------------------------------------
            int biN = v.I("biN");
            if (biN > 0)
            {
                RebarBarType tb = v.Bar(doc, "biBar");
                double db = RebarTools.Diameter(tb);
                double z = zBot + dbI / 2 + Math.Max(Un.Cm(2.5), Math.Max(db, dbI)) + db / 2;
                double len = v.D("biFr") * f.SizeX;
                double mid = (f.MinX + f.MaxX) / 2;
                Longitudinal(doc, host, f, tb, biN, mid - len / 2, mid + len / 2, z, +1, 0, ds, c, view, report, "viga bastón inf.");
            }

            // --- Estribos ----------------------------------------------------------------------------
            List<double> pos = Distribution.Positions(v.S("dist"), f.SizeX);
            double yL = f.MinY + c + ds / 2, yR = f.MaxY - c - ds / 2;
            double zb = f.MinZ + c + ds / 2, zt = f.MaxZ - c - ds / 2;
            if (yR - yL < ds * 2 || zt - zb < ds * 2)
                throw new InvalidOperationException("La sección es demasiado pequeña para el recubrimiento indicado.");

            RebarHookType h135 = RebarTools.Hook(doc, 135, RebarStyle.StirrupTie);
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
