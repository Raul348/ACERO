using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using TaskDialogCommandLinkId = Autodesk.Revit.UI.TaskDialogCommandLinkId;
using TaskDialogCommonButtons = Autodesk.Revit.UI.TaskDialogCommonButtons;
using TaskDialogResult = Autodesk.Revit.UI.TaskDialogResult;

namespace AceroRefuerzo.Commands
{
    /// <summary>
    /// Optimizador de corte: lee las barras del modelo, agrupa las piezas por diámetro y longitud,
    /// y calcula cómo cortarlas de las barras comerciales (9 m, 12 m...) con el menor desperdicio.
    /// No modifica el modelo; genera una planilla de corte para Excel.
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    public sealed class CmdCorte : IExternalCommand
    {
        private static readonly string[] Ambitos = { "Todo el proyecto", "Solo los elementos seleccionados", "Solo la vista activa" };

        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            UIDocument uidoc = data.Application.ActiveUIDocument;
            if (uidoc == null || uidoc.Document.IsFamilyDocument)
            {
                TaskDialog.Show(App.Name, "Abra un proyecto de Revit con barras de refuerzo.");
                return Result.Cancelled;
            }
            Document doc = uidoc.Document;
            var owner = new RevitWindow(data.Application.MainWindowHandle);

            try
            {
                // 1. Inventario de barras en los tres alcances posibles.
                var alcances = new List<PiezaCorte>[]
                {
                    Inventario.Todas(doc),
                    Inventario.Seleccion(uidoc),
                    Inventario.EnVista(doc, uidoc.ActiveView)
                };
                if (alcances[0].Count == 0)
                {
                    TaskDialog.Show(App.Name, "El proyecto no tiene barras de refuerzo.\nColoque acero primero (pestaña ACERO).");
                    return Result.Cancelled;
                }

                // 2. Ventana de datos con la figura del resultado (se recalcula con cada cambio).
                int totalBarras = alcances[0].Sum(p => p.Cantidad);
                var form = new ParamForm(App.Name, "corte", "CORTE de barras",
                    $"{alcances[0].Count} conjuntos · {totalBarras} barras en el proyecto · {alcances[1].Sum(p => p.Cantidad)} seleccionadas");
                form.AcceptText = "Generar planilla";

                form.Section("Barras comerciales");
                form.Txt("largos", "Longitudes disponibles (m)", "9, 12");
                form.Note("Separe con comas. El programa elige para cada barra la longitud comercial que menos desperdicia.");
                form.Num("kerf", "Pérdida por cada corte", 0, "cm", 1, 0, 5);

                form.Section("Piezas más largas que la barra comercial");
                form.Chk("empalmar", "Partirlas y empalmarlas con traslape", true);
                form.Num("traslape", "Traslape del empalme (0 = 40 diámetros)", 60, "cm", 0, 0, 500);

                form.Section("Alcance");
                form.Choice("ambito", "Barras a considerar", Ambitos, alcances[1].Count > 0 ? 1 : 0);
                form.Chk("soloAcero", "Solo las barras creadas por ACERO", false);
                form.Note("Con la selección se incluyen las barras alojadas en los elementos seleccionados (vigas, columnas, losas...).");

                form.Section("Resultado");
                form.Num("retazo", "Retazo aprovechable desde", 100, "cm", 0, 0, 1000);
                form.Chk("csv", "Guardar la planilla de corte (CSV para Excel)", true);
                form.Note("Las longitudes son las que calcula Revit para cada barra, con patas y dobleces.");

                // Caché: el mismo alcance y opciones dan el mismo resultado.
                string cacheKey = null;
                ResultadoCorte cache = null;
                ResultadoCorte Calcular(FormValues v)
                {
                    OpcionesCorte o = Opciones(v);
                    int amb = v.I("ambito");
                    string key = amb + "|" + v.B("soloAcero") + "|" + string.Join(",", o.ComercialesCm) + "|" + o.KerfCm + "|" + o.Empalmar + "|" + o.TraslapeCm + "|" + o.RetazoMinCm;
                    if (key == cacheKey) return cache;
                    List<PiezaCorte> piezas = Piezas(alcances, v);
                    cache = piezas.Count == 0 ? null : Corte.Optimizar(piezas, o);
                    cacheKey = key;
                    return cache;
                }

                form.Figure = (g, r, v) =>
                {
                    ResultadoCorte res;
                    string titulo = Ambitos[Math.Max(0, Math.Min(2, v.I("ambito")))].ToLowerInvariant();
                    try { res = Calcular(v); }
                    catch (Exception ex) { Dibujo.Error(g, new RectangleF(16, 44, r.Width - 32, 60), ex.Message); return; }
                    FiguraCorte.Dibujar(g, r, res, titulo);
                };
                form.Validator = v =>
                {
                    try { Opciones(v); }
                    catch (Exception ex) { return ex.Message; }
                    if (Piezas(alcances, v).Count == 0) return "No hay barras de refuerzo en el alcance elegido.";
                    return null;
                };

                FormValues values;
                using (form)
                {
                    if (form.ShowDialog(owner) != DialogResult.OK) return Result.Cancelled;
                    values = form.Values();
                }

                // 3. Resultado final y planilla.
                List<PiezaCorte> lista = Piezas(alcances, values);
                ResultadoCorte resultado = Corte.Optimizar(lista, Opciones(values));
                string ruta = null;
                if (values.B("csv"))
                {
                    using (var sfd = new SaveFileDialog
                    {
                        Title = "Guardar planilla de corte",
                        Filter = "Planilla CSV para Excel (*.csv)|*.csv",
                        FileName = Path.GetFileNameWithoutExtension(doc.Title) + " - Corte de barras.csv",
                        InitialDirectory = Directorio(doc)
                    })
                    {
                        if (sfd.ShowDialog(owner) == DialogResult.OK)
                        {
                            File.WriteAllText(sfd.FileName, resultado.Csv(lista), new UTF8Encoding(true));
                            ruta = sfd.FileName;
                        }
                    }
                }

                var td = new TaskDialog(App.Name + " - Corte de barras")
                {
                    MainInstruction = $"{resultado.BarrasComerciales} barras comerciales · desperdicio {resultado.Porcentaje:0.0} %",
                    MainContent =
                        $"Piezas: {resultado.Piezas}\n" +
                        $"Acero comercial: {resultado.LargoComercialCm / 100:N1} m ({resultado.KgComercial:N0} kg)\n" +
                        $"Desperdicio: {resultado.DesperdicioCm / 100:N1} m ({resultado.KgDesperdicio:N0} kg), " +
                        $"de los cuales {resultado.RetazosCm / 100:N1} m son retazos aprovechables",
                    ExpandedContent = resultado.Texto(),
                    CommonButtons = TaskDialogCommonButtons.Close
                };
                if (ruta != null)
                {
                    td.FooterText = "Planilla guardada en " + ruta;
                    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Abrir la planilla en Excel");
                }
                if (td.Show() == TaskDialogResult.CommandLink1 && ruta != null)
                {
                    try { Process.Start(new ProcessStartInfo(ruta) { UseShellExecute = true }); }
                    catch (Exception ex) { TaskDialog.Show(App.Name, "No se pudo abrir la planilla:\n" + ex.Message); }
                }
                return Result.Succeeded;
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

        private static OpcionesCorte Opciones(FormValues v) => new OpcionesCorte
        {
            ComercialesCm = Corte.ParseComerciales(v.S("largos")),
            KerfCm = v.D("kerf"),
            Empalmar = v.B("empalmar"),
            TraslapeCm = v.D("traslape"),
            RetazoMinCm = v.D("retazo"),
            RedondeoCm = 1
        };

        private static List<PiezaCorte> Piezas(List<PiezaCorte>[] alcances, FormValues v)
        {
            List<PiezaCorte> p = alcances[Math.Max(0, Math.Min(alcances.Length - 1, v.I("ambito")))];
            return v.B("soloAcero") ? p.Where(x => x.DeAcero).ToList() : p;
        }

        private static string Directorio(Document doc)
        {
            try
            {
                string dir = Path.GetDirectoryName(doc.PathName);
                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) return dir;
            }
            catch (Exception) { }
            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }
    }
}
