using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Encofrado.Core;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using TaskDialogCommonButtons = Autodesk.Revit.UI.TaskDialogCommonButtons;
using TaskDialogResult = Autodesk.Revit.UI.TaskDialogResult;

namespace Encofrado.Commands
{
    /// <summary>Base: captura errores inesperados y cancelaciones.</summary>
    public abstract class Comando : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            try
            {
                return Ejecutar(data.Application, ref message);
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

        protected abstract Result Ejecutar(UIApplication app, ref string message);

        protected static Document Doc(UIApplication app)
        {
            Document doc = app.ActiveUIDocument?.Document;
            if (doc == null || doc.IsFamilyDocument)
            {
                TaskDialog.Show(App.Name, "Abra un proyecto de Revit.");
                return null;
            }
            return doc;
        }
    }

    public abstract class ComandoPieza : Comando
    {
        protected abstract Pieza Pieza { get; }
        protected override Result Ejecutar(UIApplication app, ref string message) => Runner.PorPieza(app, Pieza, ref message);
    }

    [Transaction(TransactionMode.Manual)] public sealed class CmdColumnas : ComandoPieza { protected override Pieza Pieza => Pieza.Columna; }
    [Transaction(TransactionMode.Manual)] public sealed class CmdVigas : ComandoPieza { protected override Pieza Pieza => Pieza.Viga; }
    [Transaction(TransactionMode.Manual)] public sealed class CmdZapatas : ComandoPieza { protected override Pieza Pieza => Pieza.Zapata; }
    [Transaction(TransactionMode.Manual)] public sealed class CmdLosas : ComandoPieza { protected override Pieza Pieza => Pieza.Losa; }
    [Transaction(TransactionMode.Manual)] public sealed class CmdMuros : ComandoPieza { protected override Pieza Pieza => Pieza.Muro; }
    [Transaction(TransactionMode.Manual)] public sealed class CmdEscaleras : ComandoPieza { protected override Pieza Pieza => Pieza.Escalera; }

    /// <summary>Encofra todos los elementos de la vista activa o del proyecto.</summary>
    [Transaction(TransactionMode.Manual)]
    public sealed class CmdTodo : Comando
    {
        protected override Result Ejecutar(UIApplication app, ref string message) => Runner.Todo(app, ref message);
    }

    /// <summary>Tabla de metrado en Revit + exportación CSV.</summary>
    [Transaction(TransactionMode.Manual)]
    public sealed class CmdMetrado : Comando
    {
        protected override Result Ejecutar(UIApplication app, ref string message)
        {
            Document doc = Doc(app);
            if (doc == null) return Result.Cancelled;
            if (!Almacen.Todos(doc).Any())
            {
                TaskDialog.Show(App.Name, "Todavía no hay encofrado en el proyecto. Genere primero el encofrado de algún elemento.");
                return Result.Cancelled;
            }

            ViewSchedule tabla;
            using (var t = new Transaction(doc, "ENCOFRADO - Metrado"))
            {
                t.Start();
                Almacen.AsegurarParametros(doc, app.Application);
                tabla = Metrado.Tabla(doc);
                t.Commit();
            }

            string ruta = Metrado.Csv(doc, out Reporte resumen);
            app.ActiveUIDocument.ActiveView = tabla;

            new TaskDialog(App.Name)
            {
                MainInstruction = $"Metrado de encofrado: {resumen.TotalM2:N2} m²",
                MainContent = resumen.Tabla() + $"\nTabla: \"{Metrado.NombreTabla}\"\nCSV (Excel): {ruta}"
            }.Show();
            return Result.Succeeded;
        }
    }

    /// <summary>Elimina el encofrado de los elementos seleccionados (o de todo el proyecto).</summary>
    [Transaction(TransactionMode.Manual)]
    public sealed class CmdEliminar : Comando
    {
        protected override Result Ejecutar(UIApplication app, ref string message)
        {
            Document doc = Doc(app);
            if (doc == null) return Result.Cancelled;

            var sel = app.ActiveUIDocument.Selection.GetElementIds().Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
            var ids = new HashSet<ElementId>();
            foreach (Element e in sel)
            {
                if (e is DirectShape ds && ds.ApplicationId == Almacen.AppId) ids.Add(e.Id);
                else if (Piezas.De(e).HasValue) foreach (ElementId id in Almacen.DeElemento(doc, e)) ids.Add(id);
            }

            string alcance = "de los elementos seleccionados";
            if (ids.Count == 0)
            {
                foreach (DirectShape ds in Almacen.Todos(doc)) ids.Add(ds.Id);
                alcance = "de TODO el proyecto";
            }
            if (ids.Count == 0)
            {
                TaskDialog.Show(App.Name, "No hay encofrado para eliminar.");
                return Result.Cancelled;
            }

            var td = new TaskDialog(App.Name)
            {
                MainInstruction = $"¿Eliminar {ids.Count} pieza(s) de encofrado {alcance}?",
                MainContent = "Seleccione antes elementos estructurales si solo quiere borrar su encofrado.",
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
            };
            if (td.Show() != TaskDialogResult.Yes) return Result.Cancelled;

            using (var t = new Transaction(doc, "ENCOFRADO - Eliminar"))
            {
                t.Start();
                doc.Delete(ids.ToList());
                t.Commit();
            }
            return Result.Succeeded;
        }
    }

    /// <summary>Muestra u oculta el encofrado en la vista activa.</summary>
    [Transaction(TransactionMode.Manual)]
    public sealed class CmdVisibilidad : Comando
    {
        protected override Result Ejecutar(UIApplication app, ref string message)
        {
            Document doc = Doc(app);
            if (doc == null) return Result.Cancelled;
            View view = app.ActiveUIDocument.ActiveView;

            List<Element> piezas = Almacen.Todos(doc).Cast<Element>().Where(e => e.CanBeHidden(view)).ToList();
            if (piezas.Count == 0)
            {
                TaskDialog.Show(App.Name, "No hay encofrado que mostrar u ocultar en esta vista.");
                return Result.Cancelled;
            }
            bool algunaVisible = piezas.Any(e => !e.IsHidden(view));
            using (var t = new Transaction(doc, algunaVisible ? "ENCOFRADO - Ocultar" : "ENCOFRADO - Mostrar"))
            {
                t.Start();
                var ids = piezas.Select(e => e.Id).ToList();
                if (algunaVisible) view.HideElements(ids); else view.UnhideElements(ids);
                t.Commit();
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.ReadOnly)]
    public sealed class CmdAyuda : Comando
    {
        protected override Result Ejecutar(UIApplication app, ref string message)
        {
            new TaskDialog(App.Name)
            {
                MainInstruction = App.Name + " " + typeof(CmdAyuda).Assembly.GetName().Version.ToString(3),
                MainContent =
                    "Encofrado de elementos estructurales para Revit 2024 - 2027.\n\n" +
                    "1. Pulse el botón del elemento (o seleccione antes los elementos).\n" +
                    "2. Elija las caras a encofrar; la figura muestra tableros, puntales y yugos.\n" +
                    "3. \"Generar encofrado\" crea los tableros (Modelo genérico) descontando las zonas en contacto.\n" +
                    "4. \"Metrado\" crea la tabla de planificación y exporta un CSV para Excel.\n\n" +
                    "Caras habituales: columnas y zapatas → costados · vigas → costados y fondo · " +
                    "losas → fondo y frisos · muros → ambas caras y dinteles · escaleras → fondo y contrapasos.\n" +
                    "\"Mostrar/Ocultar\" alterna la visibilidad del encofrado en la vista activa."
            }.Show();
            return Result.Succeeded;
        }
    }
}
