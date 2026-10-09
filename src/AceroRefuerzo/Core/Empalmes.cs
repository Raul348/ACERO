using System;
using System.Collections.Generic;

namespace AceroRefuerzo.Core
{
    /// <summary>
    /// Longitud desarrollada de barras con dobleces y ubicación de traslapes cuando la barra supera la
    /// longitud comercial (9 m en Perú). Todas las funciones son independientes de las unidades
    /// (se usan pies en Revit y centímetros en las figuras).
    /// </summary>
    public static class Empalmes
    {
        /// <summary>Radio del doblez medido al eje de la barra.</summary>
        public static double RadioEje(double db, double diametroDoblez) => diametroDoblez / 2 + db / 2;

        /// <summary>
        /// Lo que se descuenta por cada doblez a 90° cuando los tramos se miden al eje hasta la intersección:
        /// la barra real sigue el arco en vez de la esquina viva (2R − πR/2).
        /// </summary>
        public static double Ahorro90(double db, double diametroDoblez) => RadioEje(db, diametroDoblez) * (2 - Math.PI / 2);

        /// <summary>Pata medida por fuera (como en los planos) → tramo al eje hasta la intersección.</summary>
        public static double PataEje(double pataExterior, double db) => pataExterior <= 0 ? 0 : Math.Max(0, pataExterior - db / 2);

        /// <summary>Longitud desarrollada: tramos al eje menos el ahorro de cada doblez a 90°.</summary>
        public static double Desarrollo(double sumaTramosEje, int dobleces90, double db, double diametroDoblez) =>
            sumaTramosEje - dobleces90 * Ahorro90(db, diametroDoblez);

        /// <summary>
        /// Lee las longitudes de los tramos escritas por el usuario, en metros: "9, 9, 4.5" (también separadas por
        /// punto y coma, espacios o "+"). Devuelve null si el texto está vacío.
        /// </summary>
        public static List<double> LeerTramos(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto)) return null;
            var largos = new List<double>();
            foreach (string t in texto.Split(new[] { ',', ';', ' ', '+', '/' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!double.TryParse(t.Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double m) || m <= 0)
                    throw new FormatException($"Longitud de tramo inválida: \"{t.Trim()}\". Escriba metros con punto decimal, por ejemplo: 9, 9, 4.5");
                largos.Add(m);
            }
            return largos.Count > 0 ? largos : null;
        }

        /// <summary>
        /// Ubica los tramos indicados por el usuario, de abajo hacia arriba y sin modificarlos: el primero empieza en
        /// <paramref name="zIni"/> y su longitud desarrollada incluye la pata y el doblez (<paramref name="extraPrimero"/>);
        /// cada tramo siguiente empieza <paramref name="ls"/> antes del final del anterior (traslape).
        /// </summary>
        public static List<(double Ini, double Fin)> UbicarTramos(double zIni, IList<double> largos, double extraPrimero, double ls)
        {
            var tramos = new List<(double, double)>();
            double s = zIni;
            for (int k = 0; k < largos.Count; k++)
            {
                double recto = largos[k] - (k == 0 ? extraPrimero : 0);
                if (recto <= (k == 0 ? 0 : ls))
                    throw new InvalidOperationException($"El tramo {k + 1} es demasiado corto (debe ser mayor que la pata y el traslape).");
                tramos.Add((s, s + recto));
                s = s + recto - ls;
            }
            return tramos;
        }
    }
}
