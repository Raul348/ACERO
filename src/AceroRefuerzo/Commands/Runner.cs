using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AceroRefuerzo.Core;
using AceroRefuerzo.Modules;
using AceroRefuerzo.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using TaskDialogCommonButtons = Autodesk.Revit.UI.TaskDialogCommonButtons;
using TaskDialogResult = Autodesk.Revit.UI.TaskDialogResult;
using Autodesk.Revit.UI.Selection;

namespace AceroRefuerzo.Commands
{
    /// <summary>Flujo común: seleccionar elementos → formulario con figura → crear acero → resumen.</summary>
    internal static class Runner
    {
        public static Result Run(UIApplication uiapp, ElementKind kind, ref string message)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc == null || uidoc.Document.IsFamilyDocument)
            {
                TaskDialog.Show("ACERO Refuerzo", "Abra un proyecto de Revit (no una familia) para colocar acero.");
                return Result.Cancelled;
            }

            Document doc = uidoc.Document;
            IRebarModule module = ModuleRegistry.Get(kind);
            var owner = new RevitWindow(uiapp.MainWindowHandle);

            // 1. Diámetros disponibles.
            List<BarItem> bars = RebarTools.BarItems(doc);
            if (bars.Count == 0)
            {
                var td = new TaskDialog("ACERO Refuerzo")
                {
                    MainInstruction = "El proyecto no tiene tipos de barra de refuerzo.",
                    MainContent = "¿Desea crear los diámetros estándar (6 mm, 8 mm, 3/8\", 12 mm, 1/2\", 5/8\", 3/4\", 1\", 1 3/8\")?",
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                };
                if (td.Show() != TaskDialogResult.Yes) return Result.Cancelled;
                using (var t = new Transaction(doc, "ACERO - Diámetros estándar"))
                {
                    t.Start();
                    RebarTools.EnsureStandardBars(doc);
                    t.Commit();
                }
                bars = RebarTools.BarItems(doc);
            }

            // 2. Selección.
            List<Element> hosts = Select(uidoc, module);
            if (hosts.Count == 0) return Result.Cancelled;

            // 3. Formulario con la figura del elemento.
            FormValues values;
            using (ParamForm form = module.CreateForm(doc, hosts[0], bars))
            {
                if (form.ShowDialog(owner) != DialogResult.OK) return Result.Cancelled;
                values = form.Values();
            }

            // 4. Creación del acero (cada elemento en su propia subtransacción).
            var report = new RunReport();
            using (var t = new Transaction(doc, "ACERO - " + module.Title))
            {
                t.Start();
                FailureHandlingOptions opts = t.GetFailureHandlingOptions();
                opts.SetFailuresPreprocessor(new WarningSwallower());
                t.SetFailureHandlingOptions(opts);

                foreach (Element host in hosts)
                {
                    using (var st = new SubTransaction(doc))
                    {
                        st.Start();
                        try
                        {
                            if (!RebarHostData.IsValidHost(host))
                                throw new InvalidOperationException("no puede alojar armadura (verifique que sea estructural).");
                            module.Apply(doc, host, values, uidoc.ActiveView, report);
                            st.Commit();
                            report.Elements++;
                        }
                        catch (Exception ex)
                        {
                            st.RollBack();
                            report.Errors.Add($"{host.Name} [Id {host.Id}]: {ex.Message}");
                        }
                    }
                }

                if (report.Elements > 0) t.Commit();
                else t.RollBack();
            }

            // 5. Resumen.
            var sum = new TaskDialog("ACERO Refuerzo - " + module.Title)
            {
                MainInstruction = report.Elements > 0
                    ? $"Acero colocado en {report.Elements} de {hosts.Count} elemento(s)."
                    : "No se pudo colocar acero.",
                MainContent =
                    $"Conjuntos de barras: {report.Sets}\n" +
                    $"Barras individuales: {report.Bars}\n" +
                    $"Longitud total: {Un.ToMm(report.LengthFt) / 1000.0:N2} m\n" +
                    $"Peso aproximado: {report.KgSteel:N1} kg",
            };
            if (report.Errors.Count > 0)
                sum.ExpandedContent = "Errores:\n" + string.Join("\n", report.Errors.Take(30));
            sum.Show();

            if (report.Elements == 0)
            {
                message = report.Errors.FirstOrDefault() ?? "No se creó acero.";
                return Result.Failed;
            }
            return Result.Succeeded;
        }

        private static List<Element> Select(UIDocument uidoc, IRebarModule module)
        {
            Document doc = uidoc.Document;
            List<Element> pre = uidoc.Selection.GetElementIds()
                .Select(id => doc.GetElement(id))
                .Where(e => e != null && module.Accepts(e))
                .ToList();
            if (pre.Count > 0) return pre;

            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(ObjectType.Element, new KindFilter(module), module.Prompt);
                return refs.Select(r => doc.GetElement(r)).Where(e => e != null).ToList();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return new List<Element>();
            }
        }

        private sealed class KindFilter : ISelectionFilter
        {
            private readonly IRebarModule _module;
            public KindFilter(IRebarModule module) { _module = module; }
            public bool AllowElement(Element elem) => _module.Accepts(elem);
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        /// <summary>Oculta las advertencias (no los errores) durante la creación masiva de barras.</summary>
        private sealed class WarningSwallower : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                foreach (FailureMessageAccessor f in accessor.GetFailureMessages())
                    if (f.GetSeverity() == FailureSeverity.Warning)
                        accessor.DeleteWarning(f);
                return FailureProcessingResult.Continue;
            }
        }
    }
}
