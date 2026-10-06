namespace Comun
{
    /// <summary>Conversión entre unidades internas de Revit (pies) y unidades métricas.</summary>
    internal static class Un
    {
        private const double MmPorPie = 304.8;

        public static double Mm(double mm) => mm / MmPorPie;
        public static double Cm(double cm) => cm * 10.0 / MmPorPie;
        public static double ToMm(double pies) => pies * MmPorPie;
        public static double ToCm(double pies) => pies * MmPorPie / 10.0;
    }
}
