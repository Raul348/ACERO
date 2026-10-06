using System;
using System.Reflection;
using Autodesk.Revit.UI;
using Encofrado.Commands;
using Encofrado.Core;
using Encofrado.UI;

namespace Encofrado
{
    /// <summary>Crea la pestaña "ENCOFRADO" con un botón (y su figura) por elemento estructural.</summary>
    public sealed class App : IExternalApplication
    {
        public const string Name = "ENCOFRADO";
        public const string TabName = "ENCOFRADO";

        public Result OnStartup(UIControlledApplication app)
        {
            try { app.CreateRibbonTab(TabName); }
            catch (Autodesk.Revit.Exceptions.ArgumentException) { /* la pestaña ya existe */ }

            string path = Assembly.GetExecutingAssembly().Location;

            RibbonPanel panel = app.CreateRibbonPanel(TabName, "Encofrado por elemento");
            Add(panel, path, typeof(CmdColumnas), Pieza.Columna, "Caras laterales de columnas rectangulares, circulares e irregulares.");
            Add(panel, path, typeof(CmdVigas), Pieza.Viga, "Costados y fondo de vigas, descontando losas y columnas.");
            Add(panel, path, typeof(CmdZapatas), Pieza.Zapata, "Costados de zapatas aisladas, cimientos corridos y losas de cimentación.");
            Add(panel, path, typeof(CmdLosas), Pieza.Losa, "Fondo de losa y frisos (bordes), descontando vigas, muros y columnas.");
            Add(panel, path, typeof(CmdMuros), Pieza.Muro, "Ambas caras, extremos, jambas y fondos de vanos de muros y placas.");
            Add(panel, path, typeof(CmdEscaleras), Pieza.Escalera, "Fondo (losa inclinada y descansos), costados y contrapasos.");

            RibbonPanel todo = app.CreateRibbonPanel(TabName, "Edificación");
            todo.AddItem(new PushButtonData("EncTodo", "Encofrar\ntodo", path, typeof(CmdTodo).FullName)
            {
                ToolTip = "Encofra todos los elementos estructurales de la vista activa o del proyecto.",
                LargeImage = Iconos.Ribbon(Iconos.Extra.Todo, 32),
                Image = Iconos.Ribbon(Iconos.Extra.Todo, 16)
            });

            RibbonPanel herr = app.CreateRibbonPanel(TabName, "Metrado y vista");
            herr.AddItem(new PushButtonData("EncMetrado", "Metrado", path, typeof(CmdMetrado).FullName)
            {
                ToolTip = "Tabla de planificación del encofrado (m² por elemento, cara y nivel) y exportación a CSV para Excel.",
                LargeImage = Iconos.Ribbon(Iconos.Extra.Metrado, 32),
                Image = Iconos.Ribbon(Iconos.Extra.Metrado, 16)
            });
            herr.AddStackedItems(
                new PushButtonData("EncVer", "Mostrar/Ocultar", path, typeof(CmdVisibilidad).FullName)
                {
                    ToolTip = "Muestra u oculta el encofrado en la vista activa.",
                    Image = Iconos.Ribbon(Iconos.Extra.Visibilidad, 16),
                    LargeImage = Iconos.Ribbon(Iconos.Extra.Visibilidad, 32)
                },
                new PushButtonData("EncEliminar", "Eliminar", path, typeof(CmdEliminar).FullName)
                {
                    ToolTip = "Elimina el encofrado de los elementos seleccionados (o de todo el proyecto).",
                    Image = Iconos.Ribbon(Iconos.Extra.Eliminar, 16),
                    LargeImage = Iconos.Ribbon(Iconos.Extra.Eliminar, 32)
                },
                new PushButtonData("EncAyuda", "Ayuda", path, typeof(CmdAyuda).FullName)
                {
                    ToolTip = "Instrucciones de uso.",
                    Image = Iconos.Ribbon(Iconos.Extra.Ayuda, 16),
                    LargeImage = Iconos.Ribbon(Iconos.Extra.Ayuda, 32)
                });

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;

        private static void Add(RibbonPanel panel, string path, Type cmd, Pieza p, string tip)
        {
            panel.AddItem(new PushButtonData("Enc" + Piezas.Plural(p), Piezas.Plural(p), path, cmd.FullName)
            {
                ToolTip = tip,
                LongDescription = "Seleccione los elementos (o pulse el botón y selecciónelos), elija las caras y el espesor del tablero. " +
                                  "El encofrado se crea como Modelo genérico con su área en el parámetro ENC_Area.",
                LargeImage = Iconos.Ribbon(p, 32),
                Image = Iconos.Ribbon(p, 16),
                ToolTipImage = Iconos.Tooltip(p)
            });
        }
    }
}
