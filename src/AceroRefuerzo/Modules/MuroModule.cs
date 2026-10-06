using System;
using System.Collections.Generic;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace AceroRefuerzo.Modules
{
    /// <summary>
    /// Muros estructurales (placas): malla simple o doble de acero vertical y horizontal,
    /// anclaje inferior, empalme superior y ganchos en los extremos del acero horizontal.
    /// </summary>
    internal sealed class MuroModule : IRebarModule
    {
        public ElementKind Kind => ElementKind.Muro;
        public string Title => "Muros";
        public string Prompt => "Seleccione los muros estructurales a armar y pulse Finalizar";

        public bool Accepts(Element e) => e is Wall w && w.WallType?.Kind == WallKind.Basic;

        private static LocalFrame Frame(Element e)
        {
            var wall = (Wall)e;
            Line line = (wall.Location as LocationCurve)?.Curve as Line
                        ?? throw new InvalidOperationException("Solo se admiten muros rectos.");
            XYZ p0 = line.GetEndPoint(0);
            XYZ x = Geo.Horizontal(line.GetEndPoint(1) - p0) ?? throw new InvalidOperationException("Muro sin longitud.");
            XYZ y = XYZ.BasisZ.CrossProduct(x).Normalize();
            if (y.DotProduct(wall.Orientation) < 0) y = -y; // +Y = cara exterior
            return LocalFrame.Measure(e, p0, x, y, XYZ.BasisZ);
        }

        public ParamForm CreateForm(Document doc, Element sample, IList<BarItem> bars)
        {
            LocalFrame f = Frame(sample);
            double L = Un.ToCm(f.SizeX), t = Un.ToCm(f.SizeY), H = Un.ToCm(f.SizeZ);

            var form = new ParamForm(App.Name, "muro", "Acero en MUROS ESTRUCTURALES",
                $"Longitud {L / 100:0.00} m   ·   Espesor {t:0.#} cm   ·   Altura {H / 100:0.00} m", bars);

            form.Section("General");
            form.Num("rec", "Recubrimiento libre", 4, "cm");
            form.Chk("doble", "Doble malla (una en cada cara)", true);

            form.Section("Acero vertical");
            form.Bar("barV", "Diámetro", 12.7);
            form.Num("sepV", "Separación", 20, "cm", 1, 5, 100);
            form.Num("anc", "Anclaje inferior (en cimentación)", 0, "cm");
            form.Num("emp", "Empalme superior (sobre el muro)", 0, "cm");

            form.Section("Acero horizontal");
            form.Bar("barH", "Diámetro", 9.525);
            form.Num("sepH", "Separación", 20, "cm", 1, 5, 100);
            form.Num("gancho", "Gancho a 90° en extremos (0 = recto)", 10, "cm");

            form.Section("Opciones");
            form.Chk("cover", "Asignar el recubrimiento al muro en Revit", true);
            form.Note("El muro debe tener la opción \"Estructural\" activada para alojar armadura.");

            form.Figure = (g, r, v) => Figures.Muro(g, r, v, L, t, H);
            return form;
        }

        public void Apply(Document doc, Element host, FormValues v, View view, RunReport report)
        {
            if (!RebarHostData.IsValidHost(host))
                throw new InvalidOperationException("El muro no es estructural (active el parámetro \"Estructural\").");

            LocalFrame f = Frame(host);
            double c = v.Cm("rec");
            if (v.B("cover")) RebarTools.SetHostCover(doc, host, c);

            RebarBarType tV = v.Bar(doc, "barV"), tH = v.Bar(doc, "barH");
            double dv = RebarTools.Diameter(tV), dh = RebarTools.Diameter(tH);
            double mid = (f.MinY + f.MaxY) / 2;

            var faces = new List<int>();
            if (v.B("doble")) { faces.Add(+1); faces.Add(-1); } else faces.Add(0);

            double anc = v.Cm("anc"), emp = v.Cm("emp");
            double zA = anc > 0 ? f.MinZ - anc : f.MinZ + c;
            double zB = emp > 0 ? f.MaxZ + emp : f.MaxZ - c;
            double xA = f.MinX + c + dh + dv / 2, xB = f.MaxX - c - dh - dv / 2;
            double hx0 = f.MinX + c, hx1 = f.MaxX - c;
            double hz0 = f.MinZ + c + dh / 2, hz1 = f.MaxZ - c - dh / 2;
            if (xB <= xA || hz1 <= hz0) throw new InvalidOperationException("El muro es demasiado pequeño para el recubrimiento indicado.");

            foreach (int s in faces)
            {
                // Vertical: capa exterior de cada cara.
                double yv = s > 0 ? f.MaxY - c - dv / 2 : s < 0 ? f.MinY + c + dv / 2 : mid;
                Rebar rv = RebarTools.Create(doc, host, RebarStyle.Standard, tV, null, null, f.X, Geo.Poly(f.P(xA, yv, zA), f.P(xA, yv, zB)));
                RebarTools.LayoutSpacing(rv, v.Cm("sepV"), xB - xA, f.X);
                RebarTools.Finish(rv, tV, s == 0 ? "muro vertical" : s > 0 ? "muro vertical ext." : "muro vertical int.", view, report);

                // Horizontal: por dentro del acero vertical; ganchos hacia el interior del muro.
                double yh = s > 0 ? f.MaxY - c - dv - dh / 2 : s < 0 ? f.MinY + c + dv + dh / 2 : mid + dv / 2 + dh / 2;
                double maxLeg = s == 0 ? (f.SizeY / 2 - c - dh) : f.SizeY - 2 * c - 2 * dv - dh;
                double leg = Math.Max(0, Math.Min(v.Cm("gancho"), maxLeg));
                int legSign = s > 0 ? -1 : +1;

                var pts = new List<XYZ>();
                if (leg > Un.Mm(5)) pts.Add(f.P(hx0, yh + legSign * leg, hz0));
                pts.Add(f.P(hx0, yh, hz0));
                pts.Add(f.P(hx1, yh, hz0));
                if (leg > Un.Mm(5)) pts.Add(f.P(hx1, yh + legSign * leg, hz0));

                Rebar rh = RebarTools.Create(doc, host, RebarStyle.Standard, tH, null, null, f.Z, Geo.Poly(pts.ToArray()));
                RebarTools.LayoutSpacing(rh, v.Cm("sepH"), hz1 - hz0, f.Z);
                RebarTools.Finish(rh, tH, s == 0 ? "muro horizontal" : s > 0 ? "muro horizontal ext." : "muro horizontal int.", view, report);
            }
        }
    }
}
