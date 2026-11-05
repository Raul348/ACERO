using System;
using System.Collections.Generic;
using System.Linq;

namespace AceroRefuerzo.Core
{
    /// <summary>
    /// Ubicación óptima de los traslapes de una barra larga (columnas de varios pisos, vigas continuas):
    /// elige dónde cortarla, solo dentro de las zonas permitidas, de modo que los tramos aprovechen
    /// las barras comerciales (9 m, 12 m...) con el menor desperdicio y el menor acero de traslape.
    /// Todas las funciones son independientes de las unidades (pies en Revit, centímetros en las figuras).
    /// </summary>
    public static class Traslapes
    {
        public sealed class Opciones
        {
            /// <summary>Longitudes comerciales disponibles (misma unidad que las cotas).</summary>
            public double[] Comerciales;
            /// <summary>Longitud del traslape.</summary>
            public double Ls;
            /// <summary>Lo que el primer tramo suma por su pata y doblez (y el último, si también lleva pata).</summary>
            public double ExtraPrimero, ExtraUltimo;
            /// <summary>Tramo más corto admisible (longitud recta, sin contar la pata).</summary>
            public double TramoMinimo;
            /// <summary>Paso de la grilla de puntos candidatos (por ejemplo 5 cm).</summary>
            public double Paso;
            /// <summary>Costo adicional por cada empalme (mano de obra), expresado en longitud de barra.</summary>
            public double PenalizacionEmpalme;
        }

        /// <summary>Un tramo de barra: desde <see cref="Ini"/> hasta <see cref="Fin"/> (cotas rectas, sin la pata).</summary>
        public sealed class Resultado
        {
            public List<(double Ini, double Fin)> Tramos = new List<(double, double)>();
            /// <summary>Longitud desarrollada de cada tramo (con la pata y el doblez del primero y del último).</summary>
            public List<double> Largos = new List<double>();
            /// <summary>Desperdicio estimado por juego de barras (suma de las sobras al cortar cada tramo de una barra comercial).</summary>
            public double Desperdicio;
            public int Empalmes => Math.Max(0, Tramos.Count - 1);
        }

        /// <summary>
        /// Zonas permitidas en columnas apiladas: una por piso, según el criterio elegido.
        /// 0 = tercio central, 1 = mitad central, 2 = personalizada (desde / hasta medidos desde la base del piso),
        /// 3 = cualquier punto.
        /// </summary>
        public static List<(double a, double b)> ZonasColumna(IList<(double zb, double zt)> pisos, int criterio, double desde, double hasta)
        {
            var z = new List<(double, double)>();
            foreach (var (zb, zt) in pisos)
            {
                double h = zt - zb;
                switch (criterio)
                {
                    case 0: z.Add((zb + h / 3, zb + 2 * h / 3)); break;
                    case 1: z.Add((zb + h / 4, zb + 3 * h / 4)); break;
                    case 2: z.Add((zb + Math.Max(0, desde), zb + Math.Min(h, hasta))); break;
                    default: z.Add((zb, zt)); break;
                }
            }
            return z.Where(t => t.Item2 > t.Item1).ToList();
        }

        /// <summary>
        /// Zonas permitidas en vigas continuas: por cada luz (xa, xb) de peralte <paramref name="h"/>.
        /// 0 = tercio central de la luz, 1 = cuartos extremos fuera de 2h de la cara del apoyo,
        /// 2 = fuera de 2h de la cara del apoyo (todo lo demás), 3 = cualquier punto.
        /// </summary>
        public static List<(double a, double b)> ZonasViga(IList<(double xa, double xb)> luces, double h, int criterio)
        {
            var z = new List<(double, double)>();
            foreach (var (xa, xb) in luces)
            {
                double L = xb - xa;
                switch (criterio)
                {
                    case 0: z.Add((xa + L / 3, xb - L / 3)); break;
                    case 1:
                        z.Add((xa + 2 * h, xa + L / 4));
                        z.Add((xb - L / 4, xb - 2 * h));
                        break;
                    case 2: z.Add((xa + 2 * h, xb - 2 * h)); break;
                    default: z.Add((xa, xb)); break;
                }
            }
            return z.Where(t => t.Item2 > t.Item1).ToList();
        }

        // ------------------------------------------------------------------ Desde el formulario

        /// <summary>Traslape del formulario (clave <paramref name="key"/>, en cm); con 0 se usan 40 diámetros.</summary>
        public static double Ls(FormValues v, string key, double db, Func<double, double> cm) =>
            v.Has(key) && v.D(key) > 0 ? cm(v.D(key)) : 40 * db;

        /// <summary>Longitudes comerciales del formulario (clave "comerciales", en m) convertidas con <paramref name="cm"/>.</summary>
        public static double[] Comerciales(FormValues v, Func<double, double> cm)
        {
            string txt = v.Has("comerciales") ? v.S("comerciales") : "";
            if (string.IsNullOrWhiteSpace(txt)) txt = v.Has("lcom") ? v.D("lcom").ToString(System.Globalization.CultureInfo.InvariantCulture) : "9";
            return Corte.ParseComerciales(txt).Select(cm).ToArray();
        }

        private static Opciones OpcionesForm(FormValues v, string lsKey, double db, double extraIni, double extraFin, Func<double, double> cm) => new Opciones
        {
            Comerciales = Comerciales(v, cm),
            Ls = Ls(v, lsKey, db, cm),
            ExtraPrimero = extraIni,
            ExtraUltimo = extraFin,
            TramoMinimo = cm(100),
            Paso = cm(5),
            PenalizacionEmpalme = cm(30)
        };

        /// <summary>
        /// Tramos óptimos de las barras longitudinales de una columna de varios pisos, con los datos del formulario
        /// (zonaCol, zonaIni, zonaFin, ls, comerciales). <paramref name="cm"/> convierte centímetros a la unidad de trabajo.
        /// </summary>
        public static Resultado Columna(FormValues v, IList<(double zb, double zt)> pisos, double zIni, double zFin, double extra, double db, Func<double, double> cm)
        {
            var zonas = ZonasColumna(pisos, v.Has("zonaCol") ? v.I("zonaCol") : 0, cm(v.Has("zonaIni") ? v.D("zonaIni") : 0), cm(v.Has("zonaFin") ? v.D("zonaFin") : 1e6));
            return Optimizar(zIni, zFin, zonas, OpcionesForm(v, "ls", db, extra, 0, cm));
        }

        /// <summary>
        /// Tramos óptimos de una fila de barras corridas de una viga continua (superior o inferior), con los datos del
        /// formulario (<paramref name="zonaKey"/>, ls, comerciales).
        /// </summary>
        public static Resultado Viga(FormValues v, IList<(double xa, double xb)> luces, double h, double x0, double x1,
            double extraIni, double extraFin, double db, string zonaKey, Func<double, double> cm)
        {
            var zonas = ZonasViga(luces, h, v.Has(zonaKey) ? v.I(zonaKey) : 0);
            return Optimizar(x0, x1, zonas, OpcionesForm(v, "ls", db, extraIni, extraFin, cm));
        }

        /// <summary>
        /// Divide la barra que va de <paramref name="ini"/> a <paramref name="fin"/> en tramos cuyos traslapes caen
        /// completos dentro de alguna zona permitida. Devuelve null si no hay solución (por ejemplo, zonas muy
        /// separadas para la barra comercial disponible).
        /// </summary>
        public static Resultado Optimizar(double ini, double fin, IList<(double a, double b)> zonas, Opciones o)
        {
            if (o == null || o.Comerciales == null || o.Comerciales.Length == 0) throw new ArgumentException("Indique las longitudes comerciales.");
            double lmax = o.Comerciales.Max();
            double ls = Math.Max(0, o.Ls);
            double paso = o.Paso > 0 ? o.Paso : lmax / 200;
            double minTramo = Math.Max(o.TramoMinimo, ls + paso);
            double total = fin - ini;
            if (total <= 0) return null;

            // Una sola barra si cabe: menos empalmes siempre que sea posible.
            if (total + o.ExtraPrimero + o.ExtraUltimo <= lmax + 1e-9)
                return Armar(new List<double> { fin }, ini, fin, o, ls);

            // Puntos candidatos donde puede terminar un tramo: el traslape [e − ls, e] debe caber en una zona.
            var candidatos = new List<double>();
            foreach (var (a, b) in zonas ?? new List<(double, double)>())
            {
                double e0 = Math.Max(a + ls, ini + minTramo), e1 = Math.Min(b, fin - minTramo);
                if (e1 < e0) continue;
                // Grilla alineada al inicio de la barra para que los tramos salgan en números redondos.
                double first = ini + Math.Ceiling((e0 - ini) / paso - 1e-9) * paso;
                for (double e = first; e <= e1 + 1e-9; e += paso) candidatos.Add(e);
                if (candidatos.Count == 0 || Math.Abs(candidatos[candidatos.Count - 1] - e1) > 1e-9) candidatos.Add(e1);
            }
            candidatos = candidatos.Where(e => e > ini + minTramo - 1e-9 && e < fin - 1e-9).Distinct().OrderBy(e => e).ToList();
            candidatos.Add(fin);
            int n = candidatos.Count;

            // Programación dinámica: costo mínimo para terminar un tramo en cada candidato.
            var costo = new double[n];
            var prev = new int[n];
            for (int j = 0; j < n; j++) { costo[j] = double.PositiveInfinity; prev[j] = -2; }

            for (int j = 0; j < n; j++)
            {
                double e = candidatos[j];
                bool ultimo = j == n - 1;
                // Desde el inicio de la barra.
                {
                    double L = e - ini;
                    double dev = L + o.ExtraPrimero + (ultimo ? o.ExtraUltimo : 0);
                    if (L >= minTramo - 1e-9 && dev <= lmax + 1e-9)
                    {
                        double c = Sobra(dev, o.Comerciales);
                        if (c < costo[j]) { costo[j] = c; prev[j] = -1; }
                    }
                }
                // Desde cada candidato anterior (el tramo empieza ls antes del final del anterior).
                for (int i = 0; i < j; i++)
                {
                    if (double.IsInfinity(costo[i])) continue;
                    double s = candidatos[i] - ls;
                    double L = e - s;
                    if (L < minTramo - 1e-9) continue;
                    double dev = L + (ultimo ? o.ExtraUltimo : 0);
                    if (dev > lmax + 1e-9) continue; // demasiado lejos; los candidatos siguientes están más cerca
                    double c = costo[i] + Sobra(dev, o.Comerciales) + ls + o.PenalizacionEmpalme;
                    if (c < costo[j]) { costo[j] = c; prev[j] = i; }
                }
            }
            if (double.IsInfinity(costo[n - 1])) return null;

            var fines = new List<double>();
            for (int j = n - 1; j >= 0; j = prev[j]) fines.Add(candidatos[j]);
            fines.Reverse();
            return Armar(fines, ini, fin, o, ls);
        }

        private static Resultado Armar(List<double> fines, double ini, double fin, Opciones o, double ls)
        {
            var r = new Resultado();
            double s = ini;
            for (int k = 0; k < fines.Count; k++)
            {
                double e = fines[k];
                r.Tramos.Add((s, e));
                double dev = e - s + (k == 0 ? o.ExtraPrimero : 0) + (k == fines.Count - 1 ? o.ExtraUltimo : 0);
                r.Largos.Add(dev);
                r.Desperdicio += Sobra(dev, o.Comerciales);
                s = e - ls;
            }
            return r;
        }

        /// <summary>
        /// Sobra por pieza al cortar piezas iguales de longitud <paramref name="L"/> de la barra comercial que mejor
        /// las aprovecha: min sobre S y k de (S/k − L), con k piezas enteras por barra.
        /// </summary>
        public static double Sobra(double L, double[] comerciales)
        {
            double best = double.PositiveInfinity;
            foreach (double S in comerciales)
            {
                int k = (int)Math.Floor(S / L + 1e-9);
                if (k < 1) continue;
                best = Math.Min(best, S / k - L);
            }
            return double.IsInfinity(best) ? L : best;
        }
    }
}
