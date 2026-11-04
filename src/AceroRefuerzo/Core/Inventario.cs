using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace AceroRefuerzo.Core
{
    /// <summary>Lee las barras de refuerzo del modelo y las convierte en piezas a cortar (diámetro, longitud, cantidad).</summary>
    internal static class Inventario
    {
        public static List<PiezaCorte> Todas(Document doc) =>
            Desde(doc, new FilteredElementCollector(doc));

        /// <summary>Barras visibles en la vista (vacío si la vista no es gráfica).</summary>
        public static List<PiezaCorte> EnVista(Document doc, View view)
        {
            if (view == null || view.IsTemplate) return new List<PiezaCorte>();
            try { return Desde(doc, new FilteredElementCollector(doc, view.Id)); }
            catch (Autodesk.Revit.Exceptions.ArgumentException) { return new List<PiezaCorte>(); }
        }

        /// <summary>Barras seleccionadas, más las alojadas en los elementos seleccionados (vigas, columnas, losas...).</summary>
        public static List<PiezaCorte> Seleccion(UIDocument uidoc)
        {
            Document doc = uidoc.Document;
            var elems = new Dictionary<ElementId, Element>();
            foreach (ElementId id in uidoc.Selection.GetElementIds())
            {
                Element e = doc.GetElement(id);
                if (e == null) continue;
                if (e is Rebar || e is RebarInSystem) { elems[e.Id] = e; continue; }
                if (e is AreaReinforcement ar) { foreach (ElementId r in ar.GetRebarInSystemIds()) Agregar(doc, elems, r); continue; }
                if (e is PathReinforcement pr) { foreach (ElementId r in pr.GetRebarInSystemIds()) Agregar(doc, elems, r); continue; }

                RebarHostData host = RebarHostData.GetRebarHostData(e);
                if (host == null) continue;
                foreach (Rebar r in host.GetRebarsInHost()) elems[r.Id] = r;
                foreach (AreaReinforcement a in host.GetAreaReinforcementsInHost())
                    foreach (ElementId r in a.GetRebarInSystemIds()) Agregar(doc, elems, r);
                foreach (PathReinforcement p in host.GetPathReinforcementsInHost())
                    foreach (ElementId r in p.GetRebarInSystemIds()) Agregar(doc, elems, r);
            }
            return elems.Values.SelectMany(e => Piezas(doc, e)).ToList();
        }

        private static void Agregar(Document doc, Dictionary<ElementId, Element> dic, ElementId id)
        {
            Element e = doc.GetElement(id);
            if (e != null) dic[e.Id] = e;
        }

        private static List<PiezaCorte> Desde(Document doc, FilteredElementCollector col) =>
            col.OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType()
               .Where(e => e is Rebar || e is RebarInSystem)
               .SelectMany(e => Piezas(doc, e))
               .ToList();

        /// <summary>Piezas de un conjunto de barras. Si las barras varían de longitud, una pieza por barra.</summary>
        public static IEnumerable<PiezaCorte> Piezas(Document doc, Element e)
        {
            var tipo = doc.GetElement(e.GetTypeId()) as RebarBarType;
            if (tipo == null) yield break;

            int n = e.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS)?.AsInteger() ?? 0;
            double total = e.get_Parameter(BuiltInParameter.REBAR_ELEM_TOTAL_LENGTH)?.AsDouble() ?? 0;
            double largo = e.get_Parameter(BuiltInParameter.REBAR_ELEM_LENGTH)?.AsDouble() ?? 0;
            if (n <= 0) n = e is Rebar rb ? rb.NumberOfBarPositions : e is RebarInSystem rs ? rs.NumberOfBarPositions : 0;
            if (n <= 0) yield break;
            if (total <= 0) total = e is Rebar r2 ? r2.TotalLength : e is RebarInSystem s2 ? s2.TotalLength : 0;
            if (largo <= 0 && total > 0) largo = total / n;
            if (largo <= 0) yield break;

            string comentario = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.AsString() ?? "";
            bool deAcero = comentario.StartsWith("ACERO - ", StringComparison.Ordinal);
            string origen = deAcero ? comentario.Substring(8) : !string.IsNullOrWhiteSpace(comentario) ? comentario : Anfitrion(doc, e);

            // Conjuntos de longitud variable: la longitud de cada barra se mide por su eje
            // y se le suma lo que Revit agrega a la primera barra (ganchos, extensiones).
            if (e is Rebar rebar && EsVariable(rebar) && n > 1)
            {
                var largos = new List<double>();
                try
                {
                    double eje0 = Eje(rebar, 0);
                    double extra = eje0 > 0 ? Math.Max(0, largo - eje0) : 0;
                    for (int i = 0; i < n; i++)
                    {
                        double l = Eje(rebar, i);
                        largos.Add(l > 0 ? l + extra : largo);
                    }
                }
                catch (Exception) { largos.Clear(); }

                if (largos.Count == n)
                {
                    foreach (var g in largos.GroupBy(l => Math.Round(Un.ToCm(l), 1)))
                        yield return new PiezaCorte { Bar = tipo.Name, Mm = Un.ToMm(tipo.BarModelDiameter), LargoCm = g.Key, Cantidad = g.Count(), Origen = origen, DeAcero = deAcero };
                    yield break;
                }
            }

            yield return new PiezaCorte { Bar = tipo.Name, Mm = Un.ToMm(tipo.BarModelDiameter), LargoCm = Un.ToCm(largo), Cantidad = n, Origen = origen, DeAcero = deAcero };
        }

        private static bool EsVariable(Rebar r)
        {
            try { return r.DistributionType == DistributionType.VaryingLength; }
            catch (Exception) { return false; }
        }

        private static double Eje(Rebar r, int i)
        {
            IList<Curve> cs = r.GetCenterlineCurves(false, false, false, MultiplanarOption.IncludeAllMultiplanarCurves, i);
            return cs == null ? 0 : cs.Sum(c => c.Length);
        }

        private static string Anfitrion(Document doc, Element e)
        {
            try
            {
                ElementId id = e is Rebar r ? r.GetHostId() : e is RebarInSystem s ? s.GetHostId() : ElementId.InvalidElementId;
                Element h = id == ElementId.InvalidElementId ? null : doc.GetElement(id);
                if (h == null) return "";
                return (h.Category?.Name ?? "") + ": " + h.Name;
            }
            catch (Exception) { return ""; }
        }
    }
}
