using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Encofrado.Core
{
    /// <summary>Resumen de una ejecución: áreas por elemento y cara.</summary>
    internal sealed class Reporte
    {
        public int Elementos;
        public int Tableros;
        public readonly Dictionary<(Pieza, Cara), double> Areas = new Dictionary<(Pieza, Cara), double>();
        public readonly List<string> Errores = new List<string>();
        public readonly List<string> Avisos = new List<string>();
        public readonly List<string> Omitidos = new List<string>();

        public void Sumar(Pieza p, Cara c, double areaPies2)
        {
            Areas.TryGetValue((p, c), out double a);
            Areas[(p, c)] = a + areaPies2;
        }

        public double TotalM2 => Almacen.M2(Areas.Values.Sum());

        public string Tabla()
        {
            var sb = new StringBuilder();
            foreach (var g in Areas.OrderBy(k => k.Key.Item1).ThenBy(k => k.Key.Item2))
                sb.AppendLine($"{Piezas.Nombre(g.Key.Item1),-10} {Piezas.NombreCara(g.Key.Item1, g.Key.Item2),-34} {Almacen.M2(g.Value),10:N2} m²");
            sb.AppendLine($"{"TOTAL",-45} {TotalM2,10:N2} m²");
            return sb.ToString();
        }
    }
}
