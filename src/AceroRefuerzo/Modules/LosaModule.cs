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
    /// Losas de techo / entrepiso macizas: refuerzo de área de Revit que sigue el contorno real de la losa,
    /// con malla inferior y superior (temperatura) en dos direcciones.
    /// </summary>
    internal sealed class LosaModule : IRebarModule
    {
        public ElementKind Kind => ElementKind.Losa;
        public string Title => "Losas";
        public string Prompt => "Seleccione las losas a armar y pulse Finalizar";

        public bool Accepts(Element e) =>
            e is Floor && e.Category?.BuiltInCategory == BuiltInCategory.OST_Floors;

        /// <summary>Contorno exterior de la losa (boceto o, si no hay, cara superior).</summary>
        private static IList<Curve> Boundary(Document doc, Floor floor)
        {
            var loops = new List<IList<Curve>>();
            try
            {
                if (doc.GetElement(floor.SketchId) is Sketch sk)
                    foreach (CurveArray arr in sk.Profile)
                        loops.Add(arr.Cast<Curve>().ToList());
            }
            catch (Exception) { }

            if (loops.Count == 0)
            {
                foreach (Reference rf in HostObjectUtils.GetTopFaces(floor))
                    if (floor.GetGeometryObjectFromReference(rf) is Face face)
                        foreach (CurveLoop cl in face.GetEdgesAsCurveLoops())
                            loops.Add(cl.ToList());
            }

            IList<Curve> best = null;
            double bestArea = 0;
            foreach (IList<Curve> loop in loops)
            {
                var pts = new List<XYZ>();
                foreach (Curve c in loop)
                {
                    IList<XYZ> tp = c.Tessellate();
                    for (int i = 0; i < tp.Count - 1; i++) pts.Add(tp[i]);
                }
                double a = Math.Abs(Geo.SignedArea(pts));
                if (a > bestArea) { bestArea = a; best = loop; }
            }
            return best ?? throw new InvalidOperationException("No se pudo leer el contorno de la losa.");
        }

        private static XYZ MajorDirection(IList<Curve> boundary, int mode)
        {
            if (mode == 1) return XYZ.BasisX;
            if (mode == 2) return XYZ.BasisY;
            Curve longest = boundary.Where(c => c is Line).OrderByDescending(c => c.Length).FirstOrDefault();
            if (longest == null) return XYZ.BasisX;
            return Geo.Horizontal(longest.GetEndPoint(1) - longest.GetEndPoint(0)) ?? XYZ.BasisX;
        }

        public ParamForm CreateForm(Document doc, IList<Element> hosts, IList<BarItem> bars)
        {
            Element sample = hosts[0];
            var floor = (Floor)sample;
            IList<Curve> boundary = Boundary(doc, floor);
            var pts = new List<XYZ>();
            foreach (Curve c in boundary)
            {
                IList<XYZ> tp = c.Tessellate();
                for (int i = 0; i < tp.Count - 1; i++) pts.Add(tp[i]);
            }
            double minX = pts.Min(p => p.X), minY = pts.Min(p => p.Y);
            List<P2> outline = pts.Select(p => new P2(Un.ToCm(p.X - minX), Un.ToCm(p.Y - minY))).ToList();
            double e = Un.ToCm(floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? Un.Cm(20));
            double lx = outline.Max(p => p.X), ly = outline.Max(p => p.Y);

            var form = new ParamForm(App.Name, "losa", "Acero en LOSAS DE TECHO",
                $"Losa {lx / 100:0.00} × {ly / 100:0.00} m (envolvente)   ·   Espesor {e:0.#} cm", bars);

            form.Section("Recubrimientos y dirección");
            form.Num("recInf", "Recubrimiento inferior", 2.5, "cm");
            form.Num("recSup", "Recubrimiento superior", 2.5, "cm");
            form.Choice("dir", "Dirección principal (barras en X)", new[] { "Borde más largo de la losa", "Eje X del proyecto", "Eje Y del proyecto" }, 0);

            form.Section("Malla inferior");
            form.Chk("iX", "Barras inferiores en X (dirección principal)", true);
            form.Bar("bIX", "Diámetro", 9.525);
            form.Num("sIX", "Separación", 20, "cm", 1, 5, 100);
            form.Chk("iY", "Barras inferiores en Y", true);
            form.Bar("bIY", "Diámetro", 9.525);
            form.Num("sIY", "Separación", 20, "cm", 1, 5, 100);

            form.Section("Malla superior / temperatura");
            form.Chk("sX", "Barras superiores en X", false);
            form.Bar("bSX", "Diámetro", 6);
            form.Num("sSX", "Separación", 25, "cm", 1, 5, 100);
            form.Chk("sY", "Barras superiores en Y", false);
            form.Bar("bSY", "Diámetro", 6);
            form.Num("sSY", "Separación", 25, "cm", 1, 5, 100);
            form.Note("Se crea un Refuerzo de Área de Revit que sigue el contorno real de la losa.");

            form.Figure = (g, r, v) => Figures.Losa(g, r, v, outline, e);
            return form;
        }

        public void Apply(Document doc, Element host, FormValues v, View view, RunReport report)
        {
            var floor = (Floor)host;
            bool iX = v.B("iX"), iY = v.B("iY"), sX = v.B("sX"), sY = v.B("sY");
            if (!(iX || iY || sX || sY)) throw new InvalidOperationException("Active al menos una capa de acero.");

            RebarTools.SetParamId(host, BuiltInParameter.CLEAR_COVER_BOTTOM, RebarTools.Cover(doc, v.Cm("recInf")).Id);
            RebarTools.SetParamId(host, BuiltInParameter.CLEAR_COVER_TOP, RebarTools.Cover(doc, v.Cm("recSup")).Id);

            IList<Curve> boundary = Boundary(doc, floor);
            XYZ major = MajorDirection(boundary, v.I("dir"));

            ElementId areaType = new FilteredElementCollector(doc).OfClass(typeof(AreaReinforcementType)).FirstElementId();
            if (areaType == null || areaType == ElementId.InvalidElementId)
                areaType = AreaReinforcementType.CreateDefaultAreaReinforcementType(doc);

            ElementId firstBar = v.Bar(doc, iX ? "bIX" : iY ? "bIY" : sX ? "bSX" : "bSY").Id;
            AreaReinforcement ar = AreaReinforcement.Create(doc, host, boundary, major, areaType, firstBar, ElementId.InvalidElementId);

            Layer(doc, ar, v, iX, "bIX", "sIX", BuiltInParameter.REBAR_SYSTEM_ACTIVE_BOTTOM_DIR_1, BuiltInParameter.REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_1, BuiltInParameter.REBAR_SYSTEM_SPACING_BOTTOM_DIR_1, BuiltInParameter.REBAR_SYSTEM_HOOK_TYPE_BOTTOM_DIR_1);
            Layer(doc, ar, v, iY, "bIY", "sIY", BuiltInParameter.REBAR_SYSTEM_ACTIVE_BOTTOM_DIR_2, BuiltInParameter.REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_2, BuiltInParameter.REBAR_SYSTEM_SPACING_BOTTOM_DIR_2, BuiltInParameter.REBAR_SYSTEM_HOOK_TYPE_BOTTOM_DIR_2);
            Layer(doc, ar, v, sX, "bSX", "sSX", BuiltInParameter.REBAR_SYSTEM_ACTIVE_TOP_DIR_1, BuiltInParameter.REBAR_SYSTEM_BAR_TYPE_TOP_DIR_1, BuiltInParameter.REBAR_SYSTEM_SPACING_TOP_DIR_1, BuiltInParameter.REBAR_SYSTEM_HOOK_TYPE_TOP_DIR_1);
            Layer(doc, ar, v, sY, "bSY", "sSY", BuiltInParameter.REBAR_SYSTEM_ACTIVE_TOP_DIR_2, BuiltInParameter.REBAR_SYSTEM_BAR_TYPE_TOP_DIR_2, BuiltInParameter.REBAR_SYSTEM_SPACING_TOP_DIR_2, BuiltInParameter.REBAR_SYSTEM_HOOK_TYPE_TOP_DIR_2);

            Parameter cm = ar.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (cm != null && !cm.IsReadOnly) cm.Set("ACERO - losa");

            doc.Regenerate();
            foreach (ElementId id in ar.GetRebarInSystemIds())
            {
                if (!(doc.GetElement(id) is RebarInSystem ris)) continue;
                var bt = doc.GetElement(ris.GetTypeId()) as RebarBarType;
                report.Add(ris, bt != null ? RebarTools.Diameter(bt) : 0, ris.NumberOfBarPositions, ris.TotalLength);
                if (view != null && !view.IsTemplate)
                {
                    try { ris.SetUnobscuredInView(view, true); } catch (Exception) { }
                }
            }
        }

        private static void Layer(Document doc, AreaReinforcement ar, FormValues v, bool active, string barKey, string sepKey,
            BuiltInParameter pActive, BuiltInParameter pBar, BuiltInParameter pSpacing, BuiltInParameter pHook)
        {
            Parameter a = ar.get_Parameter(pActive);
            if (a != null && !a.IsReadOnly) a.Set(active ? 1 : 0);
            if (!active) return;

            RebarTools.SetParamId(ar, pBar, v.Bar(doc, barKey).Id);
            RebarTools.SetParamId(ar, pHook, ElementId.InvalidElementId);
            Parameter s = ar.get_Parameter(pSpacing);
            if (s != null && !s.IsReadOnly) s.Set(v.Cm(sepKey));
        }
    }
}
