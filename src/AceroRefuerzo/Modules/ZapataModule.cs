using System;
using System.Collections.Generic;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace AceroRefuerzo.Modules
{
    /// <summary>
    /// Zapatas aisladas (y losas de cimentación rectangulares): parrilla inferior en ambas direcciones
    /// con patas hacia arriba y parrilla superior opcional con patas hacia abajo.
    /// </summary>
    internal sealed class ZapataModule : IRebarModule
    {
        public ElementKind Kind => ElementKind.Zapata;
        public string Title => "Zapatas";
        public string Prompt => "Seleccione las zapatas a armar y pulse Finalizar";

        public bool Accepts(Element e) =>
            e.Category?.BuiltInCategory == BuiltInCategory.OST_StructuralFoundation && (e is FamilyInstance || e is Floor);

        private static LocalFrame Frame(Element e)
        {
            if (e is FamilyInstance fi)
            {
                Transform t = fi.GetTransform();
                XYZ x = Geo.Horizontal(t.BasisX) ?? XYZ.BasisX;
                return LocalFrame.Measure(e, t.Origin, x, XYZ.BasisZ.CrossProduct(x), XYZ.BasisZ);
            }
            return LocalFrame.Measure(e, XYZ.Zero, XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ);
        }

        public ParamForm CreateForm(Document doc, IList<Element> hosts, IList<BarItem> bars)
        {
            Element sample = hosts[0];
            LocalFrame f = Frame(sample);
            double lx = Un.ToCm(f.SizeX), ly = Un.ToCm(f.SizeY), h = Un.ToCm(f.SizeZ);

            var form = new ParamForm(App.Name, "zapata", "Acero en ZAPATAS",
                $"Planta {lx / 100:0.00} × {ly / 100:0.00} m   ·   Peralte {h:0.#} cm", bars);

            form.Section("Recubrimiento");
            form.Num("rec", "Recubrimiento libre", 7.5, "cm");

            form.Section("Parrilla inferior");
            form.Bar("barX", "Diámetro barras en X", 15.875);
            form.Num("sepX", "Separación barras en X", 20, "cm", 1, 5, 100);
            form.Bar("barY", "Diámetro barras en Y", 15.875);
            form.Num("sepY", "Separación barras en Y", 20, "cm", 1, 5, 100);
            form.Num("pata", "Pata hacia arriba, medida por fuera (0 = recta)", 15, "cm");

            form.Section("Parrilla superior (opcional)");
            form.Chk("sup", "Colocar parrilla superior", false);
            form.Bar("barSX", "Diámetro barras en X", 12.7);
            form.Num("sepSX", "Separación barras en X", 25, "cm", 1, 5, 100);
            form.Bar("barSY", "Diámetro barras en Y", 12.7);
            form.Num("sepSY", "Separación barras en Y", 25, "cm", 1, 5, 100);
            form.Num("pataS", "Pata hacia abajo, medida por fuera (0 = recta)", 15, "cm");

            form.Section("Opciones");
            form.Chk("cover", "Asignar el recubrimiento a la zapata en Revit", true);
            form.Note("Barras en X: corren en la dirección X de la zapata y se reparten a lo largo de Y.");

            form.Figure = (g, r, v) => Figures.Zapata(g, r, v, lx, ly, h);
            return form;
        }

        public void Apply(Document doc, Element host, FormValues v, View view, RunReport report)
        {
            LocalFrame f = Frame(host);
            double c = v.Cm("rec");
            if (v.B("cover")) RebarTools.SetHostCover(doc, host, c);

            RebarBarType tX = v.Bar(doc, "barX"), tY = v.Bar(doc, "barY");
            double dx = RebarTools.Diameter(tX), dy = RebarTools.Diameter(tY);
            double maxLeg = f.SizeZ - 2 * c - dx - dy;

            // Capa 1: barras en X apoyadas sobre el recubrimiento inferior.
            double z1 = f.MinZ + c + dx / 2;
            Mesh(doc, host, f, tX, true, z1, +1, Clamp(Empalmes.PataEje(v.Cm("pata"), dx), maxLeg), v.Cm("sepX"), c, view, report, "zapata inf. X");
            // Capa 2: barras en Y sobre la capa 1.
            double z2 = f.MinZ + c + dx + dy / 2;
            Mesh(doc, host, f, tY, false, z2, +1, Clamp(Empalmes.PataEje(v.Cm("pata"), dy), maxLeg - dx), v.Cm("sepY"), c, view, report, "zapata inf. Y");

            if (!v.B("sup")) return;

            RebarBarType sX = v.Bar(doc, "barSX"), sY = v.Bar(doc, "barSY");
            double dsx = RebarTools.Diameter(sX), dsy = RebarTools.Diameter(sY);
            double zs1 = f.MaxZ - c - dsx / 2;
            double zs2 = f.MaxZ - c - dsx - dsy / 2;
            double maxLegS = zs2 - z2 - dy;
            Mesh(doc, host, f, sX, true, zs1, -1, Clamp(Empalmes.PataEje(v.Cm("pataS"), dsx), maxLegS), v.Cm("sepSX"), c, view, report, "zapata sup. X");
            Mesh(doc, host, f, sY, false, zs2, -1, Clamp(Empalmes.PataEje(v.Cm("pataS"), dsy), maxLegS - dsx), v.Cm("sepSY"), c, view, report, "zapata sup. Y");
        }

        private static double Clamp(double wanted, double max) => wanted <= 0 ? 0 : Math.Max(0, Math.Min(wanted, max));

        /// <summary>Una dirección de la parrilla: barras con patas repartidas a separación máxima.</summary>
        private static void Mesh(Document doc, Element host, LocalFrame f, RebarBarType t, bool alongX, double z,
            int legDir, double leg, double spacing, double c, View view, RunReport report, string comment)
        {
            double db = RebarTools.Diameter(t);
            double a0, a1, b0, b1;
            if (alongX)
            {
                a0 = f.MinX + c; a1 = f.MaxX - c;
                b0 = f.MinY + c + db / 2; b1 = f.MaxY - c - db / 2;
            }
            else
            {
                a0 = f.MinY + c; a1 = f.MaxY - c;
                b0 = f.MinX + c + db / 2; b1 = f.MaxX - c - db / 2;
            }
            if (a1 <= a0 || b1 <= b0) throw new InvalidOperationException("La zapata es demasiado pequeña para el recubrimiento indicado.");

            Func<double, double, double, XYZ> P = alongX
                ? (Func<double, double, double, XYZ>)((a, b, zz) => f.P(a, b, zz))
                : (a, b, zz) => f.P(b, a, zz);

            var pts = new List<XYZ>();
            if (leg > Un.Mm(5)) pts.Add(P(a0, b0, z + legDir * leg));
            pts.Add(P(a0, b0, z));
            pts.Add(P(a1, b0, z));
            if (leg > Un.Mm(5)) pts.Add(P(a1, b0, z + legDir * leg));

            XYZ dir = alongX ? f.Y : f.X;
            Rebar r = RebarTools.Create(doc, host, RebarStyle.Standard, t, null, null, dir, Geo.Poly(pts.ToArray()));
            RebarTools.LayoutSpacing(r, spacing, b1 - b0, dir);
            RebarTools.Finish(r, t, comment, view, report);
        }
    }
}
