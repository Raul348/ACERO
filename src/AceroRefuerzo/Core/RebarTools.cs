using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Comun;

namespace AceroRefuerzo.Core
{
    /// <summary>Resumen de lo creado en una ejecución.</summary>
    internal sealed class RunReport
    {
        public int Sets;
        public int Bars;
        public double LengthFt;
        public double KgSteel;
        public int Elements;
        public readonly List<string> Errors = new List<string>();
        public readonly List<string> Warnings = new List<string>();

        public void Add(Element rebar, double diameterFt, int positions, double totalLength)
        {
            if (rebar == null) return;
            Sets++;
            Bars += positions;
            LengthFt += totalLength;
            double dM = Un.ToMm(diameterFt) / 1000.0;
            double lM = Un.ToMm(totalLength) / 1000.0;
            KgSteel += Math.PI * dM * dM / 4.0 * lM * 7850.0;
        }
    }

    internal static class RebarTools
    {
        /// <summary>Diámetros estándar (nombre, mm) que se crean si el proyecto no tiene tipos de barra.</summary>
        public static readonly (string Name, double Mm)[] Standard =
        {
            ("6mm", 6.0), ("8mm", 8.0), ("3/8\"", 9.525), ("12mm", 12.0), ("1/2\"", 12.7),
            ("5/8\"", 15.875), ("3/4\"", 19.05), ("1\"", 25.4), ("1 3/8\"", 34.925)
        };

        public static double Diameter(RebarBarType t) => t.BarModelDiameter;

        public static List<BarItem> BarItems(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>()
                .Select(t => new BarItem { Id = t.Id, Name = t.Name, Mm = Un.ToMm(Diameter(t)) })
                .OrderBy(b => b.Mm).ThenBy(b => b.Name).ToList();

        /// <summary>Crea los diámetros estándar que falten. Devuelve cuántos creó.</summary>
        public static int EnsureStandardBars(Document doc)
        {
            var existing = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType))
                .Cast<RebarBarType>().ToList();
            int created = 0;
            foreach (var (name, mm) in Standard)
            {
                string fullName = "Ø " + name;
                double d = Un.Mm(mm);
                if (existing.Any(t => t.Name == fullName || Math.Abs(Diameter(t) - d) < Un.Mm(0.05)))
                    continue;

                RebarBarType t = RebarBarType.Create(doc);
                t.Name = fullName;
                t.BarNominalDiameter = d;
                t.BarModelDiameter = d;
                t.StandardBendDiameter = d * (mm <= 25.4 ? 6 : 8);
                t.StandardHookBendDiameter = d * (mm <= 25.4 ? 6 : 8);
                t.StirrupTieBendDiameter = d * 4;
                existing.Add(t);
                created++;
            }
            return created;
        }

        /// <summary>Busca (o crea) un tipo de gancho con el ángulo indicado.</summary>
        public static RebarHookType Hook(Document doc, double angleDeg, RebarStyle style)
        {
            double ang = angleDeg * Math.PI / 180.0;
            var hooks = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType))
                .Cast<RebarHookType>().Where(h => Math.Abs(h.HookAngle - ang) < 0.01).ToList();

            RebarHookType hook = hooks.FirstOrDefault(h => h.Style == style) ?? hooks.FirstOrDefault();
            if (hook != null) return hook;

            hook = RebarHookType.Create(doc, ang, style == RebarStyle.StirrupTie ? 6.0 : 12.0);
            hook.Style = style;
            try { hook.Name = $"ACERO - Gancho {angleDeg:0}°" + (style == RebarStyle.StirrupTie ? " estribo" : ""); }
            catch (Exception) { /* nombre repetido: se deja el automático */ }
            return hook;
        }

        /// <summary>Busca o crea un tipo de recubrimiento con la distancia indicada.</summary>
        public static RebarCoverType Cover(Document doc, double distance)
        {
            var found = new FilteredElementCollector(doc).OfClass(typeof(RebarCoverType)).Cast<RebarCoverType>()
                .FirstOrDefault(c => Math.Abs(c.CoverDistance - distance) < Un.Mm(0.5));
            if (found != null) return found;

            string name = $"ACERO {Un.ToMm(distance):0.#} mm";
            int i = 1;
            var names = new HashSet<string>(new FilteredElementCollector(doc).OfClass(typeof(RebarCoverType)).Select(c => c.Name));
            while (names.Contains(name)) name = $"ACERO {Un.ToMm(distance):0.#} mm ({++i})";
            return RebarCoverType.Create(doc, name, distance);
        }

        /// <summary>Asigna el recubrimiento a todas las caras del anfitrión.</summary>
        public static void SetHostCover(Document doc, Element host, double distance)
        {
            RebarHostData data = RebarHostData.GetRebarHostData(host);
            if (data == null) return;
            try { data.SetCommonCoverType(Cover(doc, distance)); }
            catch (Exception) { /* algunos anfitriones no admiten recubrimiento común */ }
        }

        public static void SetParamId(Element e, BuiltInParameter bip, ElementId id)
        {
            Parameter p = e.get_Parameter(bip);
            if (p != null && !p.IsReadOnly) p.Set(id);
        }

        /// <summary>
        /// Crea una barra desde curvas. Si lleva ganchos prueba las cuatro orientaciones posibles y se queda
        /// con la primera que cumpla <paramref name="accept"/> (por ejemplo, ganchos hacia el interior del estribo).
        /// </summary>
        public static Rebar Create(Document doc, Element host, RebarStyle style, RebarBarType barType,
            RebarHookType hookStart, RebarHookType hookEnd, XYZ normal, IList<Curve> curves,
            Func<Rebar, bool> accept = null)
        {
            if (curves == null || curves.Count == 0)
                throw new InvalidOperationException("Geometría de barra vacía (revise recubrimientos y dimensiones).");

            // (inicio a la izquierda, fin a la izquierda) para las cuatro combinaciones posibles.
            var orientations = new[] { (true, true), (false, false), (true, false), (false, true) };
            if (hookStart == null && hookEnd == null) orientations = new[] { orientations[0] };

            Exception last = null;
            Rebar fallback = null;
            foreach (var (leftStart, leftEnd) in orientations)
            {
                Rebar r = null;
                try { r = FromCurves(doc, style, barType, hookStart, hookEnd, host, normal, curves, leftStart, leftEnd); }
                catch (Exception ex) { last = ex; }

                if (r == null) continue;
                if (accept == null) return r;

                bool ok;
                try { doc.Regenerate(); ok = accept(r); }
                catch (Exception) { ok = true; }

                if (ok)
                {
                    if (fallback != null) doc.Delete(fallback.Id);
                    return r;
                }
                if (fallback == null) fallback = r; else doc.Delete(r.Id);
            }

            if (fallback != null) return fallback;

            // Último intento: sin ganchos.
            if (hookStart != null || hookEnd != null)
            {
                try
                {
                    Rebar r = FromCurves(doc, style, barType, null, null, host, normal, curves, true, true);
                    if (r != null) return r;
                }
                catch (Exception ex) { last = ex; }
            }

            throw new InvalidOperationException("Revit no pudo crear la barra: " + (last?.Message ?? "geometría no válida"), last);
        }

        /// <summary>Rebar.CreateFromCurves con la firma que corresponde a cada versión de Revit.</summary>
        private static Rebar FromCurves(Document doc, RebarStyle style, RebarBarType barType, RebarHookType hookStart,
            RebarHookType hookEnd, Element host, XYZ normal, IList<Curve> curves, bool leftStart, bool leftEnd)
        {
#if REVIT2026_OR_GREATER
            var terms = new BarTerminationsData(doc);
            if (hookStart != null) terms.HookTypeIdAtStart = hookStart.Id;
            if (hookEnd != null) terms.HookTypeIdAtEnd = hookEnd.Id;
            terms.TerminationOrientationAtStart = leftStart ? RebarTerminationOrientation.Left : RebarTerminationOrientation.Right;
            terms.TerminationOrientationAtEnd = leftEnd ? RebarTerminationOrientation.Left : RebarTerminationOrientation.Right;
            return Rebar.CreateFromCurves(doc, style, barType, host, normal, curves, terms, true, true);
#else
            return Rebar.CreateFromCurves(doc, style, barType, hookStart, hookEnd, host, normal, curves,
                leftStart ? RebarHookOrientation.Left : RebarHookOrientation.Right,
                leftEnd ? RebarHookOrientation.Left : RebarHookOrientation.Right, true, true);
#endif
        }

        /// <summary>Acepta la barra si los extremos de sus ganchos quedan dentro del polígono indicado.</summary>
        public static Func<Rebar, bool> TipsInside(LocalFrame f, IList<XYZ> polygonLocal)
        {
            return r =>
            {
                IList<Curve> cl = r.GetCenterlineCurves(false, false, false, MultiplanarOption.IncludeOnlyPlanarCurves, 0);
                if (cl == null || cl.Count == 0) return true;
                XYZ a = f.ToLocal(cl[0].GetEndPoint(0));
                XYZ b = f.ToLocal(cl[cl.Count - 1].GetEndPoint(1));
                return Geo.Contains(polygonLocal, a.X, a.Y) && Geo.Contains(polygonLocal, b.X, b.Y);
            };
        }

        /// <summary>Distribuye la barra como "número fijo" a lo largo de <paramref name="length"/> en la dirección indicada.</summary>
        public static void LayoutFixed(Rebar r, int count, double length, XYZ direction)
        {
            if (r == null || count <= 1 || length <= Un.Mm(1)) return;
            RebarShapeDrivenAccessor acc = r.GetShapeDrivenAccessor();
            acc.SetLayoutAsFixedNumber(count, length, acc.Normal.DotProduct(direction) > 0, true, true);
        }

        /// <summary>Distribuye la barra con separación máxima a lo largo de <paramref name="length"/>.</summary>
        public static void LayoutSpacing(Rebar r, double spacing, double length, XYZ direction)
        {
            if (r == null || spacing <= 0 || length <= spacing * 0.5) return;
            RebarShapeDrivenAccessor acc = r.GetShapeDrivenAccessor();
            acc.SetLayoutAsMaximumSpacing(spacing, length, acc.Normal.DotProduct(direction) > 0, true, true);
        }

        /// <summary>Marca la barra (comentario), la deja visible sin oclusión en la vista activa y la suma al reporte.</summary>
        public static void Finish(Rebar r, RebarBarType type, string comment, View view, RunReport report)
        {
            if (r == null) return;
            Parameter p = r.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (p != null && !p.IsReadOnly) p.Set("ACERO - " + comment);

            if (view != null && !view.IsTemplate)
            {
                try { r.SetUnobscuredInView(view, true); } catch (Exception) { }
            }

            report.Add(r, Diameter(type), r.NumberOfBarPositions, r.TotalLength);
        }
    }
}
