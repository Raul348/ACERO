using System;
using System.Reflection;
using AceroRefuerzo.Commands;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.UI;
using Comun;

namespace AceroRefuerzo
{
    /// <summary>Crea la pestaña "ACERO" en la cinta de Revit con un botón (y su figura) por elemento.</summary>
    public sealed class App : IExternalApplication
    {
        public const string Name = "ACERO Refuerzo";
        public const string TabName = "ACERO";

        public Result OnStartup(UIControlledApplication app)
        {
            try { app.CreateRibbonTab(TabName); }
            catch (Autodesk.Revit.Exceptions.ArgumentException) { /* la pestaña ya existe */ }

            string path = Assembly.GetExecutingAssembly().Location;

            RibbonPanel panel = app.CreateRibbonPanel(TabName, "Colocación de acero");
            Add(panel, path, typeof(CmdVigas), "Vigas", ElementKind.Viga,
                "Acero corrido superior e inferior, bastones y estribos por zonas.",
                "Seleccione una o varias vigas rectas. Crea acero corrido con patas en los apoyos, bastones superiores " +
                "en apoyos e inferiores al centro, y estribos cerrados con ganchos a 135° según la distribución (1@5, 10@10, R@20).");
            Add(panel, path, typeof(CmdColumnas), "Columnas", ElementKind.Columna,
                "Columnas rectangulares e irregulares (L, T, cruz, circulares).",
                "Rectangulares: barras por cara, estribo perimetral y grapas. Irregulares: lee el contorno real, " +
                "coloca barras en las esquinas y en los lados, y un estribo que sigue la forma. Incluye anclaje, pata y empalme.");
            Add(panel, path, typeof(CmdZapatas), "Zapatas", ElementKind.Zapata,
                "Parrilla inferior y superior con patas.",
                "Seleccione zapatas aisladas. Crea parrilla inferior en X e Y con patas hacia arriba y, opcionalmente, parrilla superior.");
            Add(panel, path, typeof(CmdLosas), "Losas", ElementKind.Losa,
                "Mallas inferior y superior que siguen el contorno de la losa.",
                "Seleccione losas (suelos estructurales). Crea un refuerzo de área de Revit con malla inferior y de temperatura.");
            Add(panel, path, typeof(CmdMuros), "Muros", ElementKind.Muro,
                "Malla simple o doble de acero vertical y horizontal.",
                "Seleccione muros estructurales rectos. Crea acero vertical y horizontal en una o dos caras, con anclaje, empalme y ganchos.");

            RibbonPanel tools = app.CreateRibbonPanel(TabName, "Herramientas");
            var integral = new PushButtonData("AceroIntegral", "Acero\nintegral", path, typeof(CmdIntegral).FullName)
            {
                ToolTip = "Menú principal: elija el elemento a armar desde su figura.",
                LargeImage = Icons.Ribbon(Icons.Extra.Integral, 32),
                Image = Icons.Ribbon(Icons.Extra.Integral, 16)
            };
            tools.AddItem(integral);
            var diam = new PushButtonData("AceroDiametros", "Diámetros", path, typeof(CmdDiametros).FullName)
            {
                ToolTip = "Crea los diámetros estándar que falten (6 mm, 8 mm, 3/8\", 12 mm, 1/2\", 5/8\", 3/4\", 1\", 1 3/8\").",
                LargeImage = Icons.Ribbon(Icons.Extra.Diametros, 32),
                Image = Icons.Ribbon(Icons.Extra.Diametros, 16)
            };
            var help = new PushButtonData("AceroAyuda", "Ayuda", path, typeof(CmdAyuda).FullName)
            {
                ToolTip = "Instrucciones de uso.",
                LargeImage = Icons.Ribbon(Icons.Extra.Ayuda, 32),
                Image = Icons.Ribbon(Icons.Extra.Ayuda, 16)
            };
            tools.AddStackedItems(diam, help);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;

        private static void Add(RibbonPanel panel, string path, Type cmd, string text, ElementKind kind, string tip, string longTip)
        {
            var data = new PushButtonData("Acero" + text, text, path, cmd.FullName)
            {
                ToolTip = tip,
                LongDescription = longTip,
                LargeImage = Icons.Ribbon(kind, 32),
                Image = Icons.Ribbon(kind, 16),
                ToolTipImage = Icons.Tooltip(kind)
            };
            panel.AddItem(data);
        }
    }
}
