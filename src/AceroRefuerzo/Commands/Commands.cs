using System;
using System.Windows.Forms;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using TaskDialogCommonButtons = Autodesk.Revit.UI.TaskDialogCommonButtons;
using TaskDialogResult = Autodesk.Revit.UI.TaskDialogResult;

namespace AceroRefuerzo.Commands
{
    public abstract class KindCommand : IExternalCommand
    {
        protected abstract ElementKind Kind { get; }

        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            try
            {
                return Runner.Run(data.Application, Kind, ref message);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show(App.Name, "Error inesperado:\n" + ex.Message);
                return Result.Failed;
            }
        }
    }

    [Transaction(TransactionMode.Manual)]
    public sealed class CmdVigas : KindCommand { protected override ElementKind Kind => ElementKind.Viga; }

    [Transaction(TransactionMode.Manual)]
    public sealed class CmdColumnas : KindCommand { protected override ElementKind Kind => ElementKind.Columna; }

    [Transaction(TransactionMode.Manual)]
    public sealed class CmdZapatas : KindCommand { protected override ElementKind Kind => ElementKind.Zapata; }

    [Transaction(TransactionMode.Manual)]
    public sealed class CmdLosas : KindCommand { protected override ElementKind Kind => ElementKind.Losa; }

    [Transaction(TransactionMode.Manual)]
    public sealed class CmdMuros : KindCommand { protected override ElementKind Kind => ElementKind.Muro; }

    /// <summary>Menú principal con las cinco figuras.</summary>
    [Transaction(TransactionMode.Manual)]
    public sealed class CmdIntegral : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            ElementKind kind;
            using (var f = new LauncherForm())
            {
                if (f.ShowDialog(new RevitWindow(data.Application.MainWindowHandle)) != DialogResult.OK)
                    return Result.Cancelled;
                kind = f.Selected;
            }
            try
            {
                return Runner.Run(data.Application, kind, ref message);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show(App.Name, "Error inesperado:\n" + ex.Message);
                return Result.Failed;
            }
        }
    }

    /// <summary>Crea los diámetros estándar que falten en el proyecto.</summary>
    [Transaction(TransactionMode.Manual)]
    public sealed class CmdDiametros : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            Document doc = data.Application.ActiveUIDocument?.Document;
            if (doc == null || doc.IsFamilyDocument)
            {
                TaskDialog.Show(App.Name, "Abra un proyecto de Revit.");
                return Result.Cancelled;
            }
            int n;
            using (var t = new Transaction(doc, "ACERO - Diámetros estándar"))
            {
                t.Start();
                n = RebarTools.EnsureStandardBars(doc);
                t.Commit();
            }
            TaskDialog.Show(App.Name, n > 0
                ? $"Se crearon {n} tipos de barra (Ø 6 mm a Ø 1 3/8\")."
                : "El proyecto ya tiene todos los diámetros estándar.");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.ReadOnly)]
    public sealed class CmdAyuda : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            var td = new TaskDialog(App.Name)
            {
                MainInstruction = App.Name + " " + typeof(CmdAyuda).Assembly.GetName().Version.ToString(3),
                MainContent =
                    "Colocación de acero de refuerzo para Revit 2024 - 2027.\n\n" +
                    "1. Seleccione los elementos (o pulse el botón y selecciónelos).\n" +
                    "2. Complete los datos; la figura muestra el armado en tiempo real.\n" +
                    "3. Pulse \"Colocar acero\".\n\n" +
                    "Distribución de estribos: 1@5, 10@10, R@20 (cm) se aplica desde cada extremo.\n" +
                    "Columnas: rectangulares o irregulares (L, T, cruz, circulares, poligonales).\n" +
                    "Losas: refuerzo de área que sigue el contorno real.\n" +
                    "Los últimos valores usados se recuerdan para la próxima vez.",
            };
            td.Show();
            return Result.Succeeded;
        }
    }
}
