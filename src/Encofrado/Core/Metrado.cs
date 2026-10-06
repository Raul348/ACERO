using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;

namespace Encofrado.Core
{
    /// <summary>Metrado del encofrado: tabla de planificación en Revit y archivo CSV.</summary>
    internal static class Metrado
    {
        public const string NombreTabla = "ENCOFRADO - Metrado";

        /// <summary>Crea (o devuelve) la tabla agrupada por elemento y cara, con subtotales y total.</summary>
        public static ViewSchedule Tabla(Document doc)
        {
            ViewSchedule existente = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule)).Cast<ViewSchedule>()
                .FirstOrDefault(v => v.Name == NombreTabla);
            if (existente != null) return existente;

            ViewSchedule vs = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            vs.Name = NombreTabla;
            ScheduleDefinition def = vs.Definition;

            var campos = new Dictionary<string, ScheduleField>();
            foreach (string n in new[] { Almacen.PElemento, Almacen.PCara, Almacen.PNivel, Almacen.PId, Almacen.PArea })
            {
                SchedulableField sf = def.GetSchedulableFields().FirstOrDefault(f => f.GetName(doc) == n);
                if (sf != null) campos[n] = def.AddField(sf);
            }
            if (campos.TryGetValue(Almacen.PElemento, out ScheduleField fe))
            {
                fe.ColumnHeading = "Elemento";
                def.AddFilter(new ScheduleFilter(fe.FieldId, ScheduleFilterType.HasValue));
                def.AddSortGroupField(new ScheduleSortGroupField(fe.FieldId) { ShowHeader = true, ShowFooter = true });
            }
            if (campos.TryGetValue(Almacen.PCara, out ScheduleField fc))
            {
                fc.ColumnHeading = "Cara";
                def.AddSortGroupField(new ScheduleSortGroupField(fc.FieldId) { ShowFooter = true });
            }
            if (campos.TryGetValue(Almacen.PNivel, out ScheduleField fn))
            {
                fn.ColumnHeading = "Nivel";
                def.AddSortGroupField(new ScheduleSortGroupField(fn.FieldId));
            }
            if (campos.TryGetValue(Almacen.PId, out ScheduleField fi)) fi.ColumnHeading = "Id elemento";
            if (campos.TryGetValue(Almacen.PArea, out ScheduleField fa))
            {
                fa.ColumnHeading = "Área (m²)";
                fa.DisplayType = ScheduleFieldDisplayType.Totals;
            }
            def.IsItemized = true;
            def.ShowGrandTotal = true;
            def.ShowGrandTotalTitle = true;
            def.GrandTotalTitle = "TOTAL ENCOFRADO";
            return vs;
        }

        /// <summary>Exporta el detalle y el resumen a CSV (separador ;) en Documentos. Devuelve la ruta.</summary>
        public static string Csv(Document doc, out Reporte resumen)
        {
            resumen = new Reporte();
            var filas = new List<string[]>();
            foreach (DirectShape ds in Almacen.Todos(doc))
            {
                string el = Almacen.Texto(ds, Almacen.PElemento);
                string cara = Almacen.Texto(ds, Almacen.PCara);
                double a = Almacen.Area(ds);
                filas.Add(new[] { el, cara, Almacen.Texto(ds, Almacen.PNivel), Almacen.Texto(ds, Almacen.PId), Almacen.M2(a).ToString("0.00", CultureInfo.CurrentCulture) });

                Pieza? p = Piezas.Todas.Cast<Pieza?>().FirstOrDefault(x => Piezas.Nombre(x.Value) == el);
                if (p.HasValue)
                    resumen.Sumar(p.Value, cara == Piezas.NombreCara(p.Value, Cara.Fondo) ? Cara.Fondo : Cara.Lateral, a);
                resumen.Tableros++;
            }

            var sb = new StringBuilder();
            sb.AppendLine("Elemento;Cara;Nivel;Id elemento;Área (m²)");
            foreach (string[] f in filas.OrderBy(f => f[0]).ThenBy(f => f[1]).ThenBy(f => f[2]))
                sb.AppendLine(string.Join(";", f.Select(Escapar)));
            sb.AppendLine();
            sb.AppendLine("RESUMEN;;;;");
            foreach (var g in resumen.Areas.OrderBy(k => k.Key.Item1).ThenBy(k => k.Key.Item2))
                sb.AppendLine($"{Piezas.Nombre(g.Key.Item1)};{Piezas.NombreCara(g.Key.Item1, g.Key.Item2)};;;{Almacen.M2(g.Value).ToString("0.00", CultureInfo.CurrentCulture)}");
            sb.AppendLine($"TOTAL;;;;{resumen.TotalM2.ToString("0.00", CultureInfo.CurrentCulture)}");

            string nombre = string.Concat((doc.Title ?? "Proyecto").Split(Path.GetInvalidFileNameChars()));
            string ruta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                $"Encofrado - {nombre} - {DateTime.Now:yyyyMMdd-HHmm}.csv");
            File.WriteAllText(ruta, sb.ToString(), new UTF8Encoding(true));
            return ruta;
        }

        private static string Escapar(string s) =>
            s.IndexOfAny(new[] { ';', '"', '\n' }) >= 0 ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;
    }
}
