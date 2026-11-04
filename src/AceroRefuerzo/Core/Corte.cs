using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace AceroRefuerzo.Core
{
    /// <summary>Pieza a cortar: un grupo de barras iguales (mismo diámetro y longitud).</summary>
    public sealed class PiezaCorte
    {
        /// <summary>Nombre del tipo de barra en Revit (por ejemplo, Ø 5/8").</summary>
        public string Bar;
        public double Mm;
        /// <summary>Longitud de cada barra, en cm (con patas y dobleces, como la calcula Revit).</summary>
        public double LargoCm;
        public int Cantidad;
        /// <summary>De dónde viene (comentario o anfitrión), solo para la planilla.</summary>
        public string Origen;
        /// <summary>Barra creada por este programa (comentario "ACERO - ...").</summary>
        public bool DeAcero;
    }

    public sealed class OpcionesCorte
    {
        /// <summary>Longitudes comerciales disponibles, en cm (por ejemplo 900 y 1200).</summary>
        public double[] ComercialesCm = { 900 };
        /// <summary>Lo que se pierde en cada corte (disco), en cm.</summary>
        public double KerfCm;
        /// <summary>Si una pieza es más larga que la barra comercial, partirla con traslape.</summary>
        public bool Empalmar = true;
        /// <summary>Traslape del empalme, en cm. Con 0 se usan 40 diámetros.</summary>
        public double TraslapeCm = 60;
        /// <summary>Sobrante mínimo para considerarlo retazo aprovechable, en cm.</summary>
        public double RetazoMinCm = 100;
        /// <summary>Las longitudes se redondean hacia arriba a este múltiplo, en cm.</summary>
        public double RedondeoCm = 1;

        public double Traslape(double mm) => TraslapeCm > 0 ? TraslapeCm : 40 * mm / 10.0;
    }

    /// <summary>Una forma de cortar una barra comercial, y cuántas veces se repite.</summary>
    public sealed class PatronCorte
    {
        public double ComercialCm;
        /// <summary>Longitudes de las piezas (cm), de mayor a menor.</summary>
        public List<double> Piezas = new List<double>();
        public int Veces;
        /// <summary>Sobrante físico por barra: comercial − piezas − pérdidas de corte.</summary>
        public double SobraCm;

        public string Texto(double kerf)
        {
            string piezas = string.Join(" + ", Piezas.Select(p => (p / 100).ToString("0.00", CultureInfo.InvariantCulture)));
            string sobra = SobraCm > 0.5 ? $"  (sobra {SobraCm / 100:0.00} m)" : "  (sin sobra)";
            return $"{Veces} × barra de {ComercialCm / 100:0.##} m: {piezas}{sobra}";
        }
    }

    /// <summary>Resultado del corte para un diámetro.</summary>
    public sealed class CorteDiametro
    {
        public string Bar;
        public double Mm;
        public int PiezasPedidas;
        public int PiezasCortadas;
        public int PiezasOmitidas;
        public int Empalmes;
        public double LargoTraslapesCm;
        public List<PatronCorte> Patrones = new List<PatronCorte>();
        public readonly SortedDictionary<double, int> BarrasPorComercial = new SortedDictionary<double, int>();
        /// <summary>Retazos aprovechables (largo en cm → cantidad).</summary>
        public readonly SortedDictionary<double, int> Retazos = new SortedDictionary<double, int>();
        public readonly List<string> Avisos = new List<string>();

        public int BarrasComerciales => BarrasPorComercial.Values.Sum();
        public double LargoComercialCm => BarrasPorComercial.Sum(kv => kv.Key * kv.Value);
        public double LargoUtilCm => Patrones.Sum(p => p.Piezas.Sum() * p.Veces);
        public double DesperdicioCm => Math.Max(0, LargoComercialCm - LargoUtilCm);
        public double RetazosCm => Retazos.Sum(kv => kv.Key * kv.Value);
        public double Porcentaje => LargoComercialCm > 0 ? 100.0 * DesperdicioCm / LargoComercialCm : 0;
        public double KgPorM => Math.PI * Mm * Mm / 4.0 / 1e6 * 7850.0;
        public double KgComercial => LargoComercialCm / 100 * KgPorM;
        public double KgDesperdicio => DesperdicioCm / 100 * KgPorM;

        public string Resumen()
        {
            string barras = string.Join(" + ", BarrasPorComercial.Select(kv => $"{kv.Value} de {kv.Key / 100:0.##} m"));
            return $"Ø {Bar}: {PiezasCortadas} piezas → {BarrasComerciales} barras ({barras}) = {LargoComercialCm / 100:0.0} m · " +
                   $"{KgComercial:0} kg · desperdicio {Porcentaje:0.0} % ({DesperdicioCm / 100:0.0} m, {KgDesperdicio:0} kg)";
        }
    }

    public sealed class ResultadoCorte
    {
        public readonly List<CorteDiametro> Diametros = new List<CorteDiametro>();
        public readonly List<string> Avisos = new List<string>();
        public OpcionesCorte Opciones;

        public int BarrasComerciales => Diametros.Sum(d => d.BarrasComerciales);
        public int Piezas => Diametros.Sum(d => d.PiezasCortadas);
        public double LargoComercialCm => Diametros.Sum(d => d.LargoComercialCm);
        public double DesperdicioCm => Diametros.Sum(d => d.DesperdicioCm);
        public double RetazosCm => Diametros.Sum(d => d.RetazosCm);
        public double Porcentaje => LargoComercialCm > 0 ? 100.0 * DesperdicioCm / LargoComercialCm : 0;
        public double KgComercial => Diametros.Sum(d => d.KgComercial);
        public double KgDesperdicio => Diametros.Sum(d => d.KgDesperdicio);

        /// <summary>Informe completo en texto: resumen, patrones por diámetro, retazos y avisos.</summary>
        public string Texto()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Barras comerciales: {BarrasComerciales} ({LargoComercialCm / 100:0.0} m, {KgComercial:0} kg)");
            sb.AppendLine($"Desperdicio: {Porcentaje:0.0} % = {DesperdicioCm / 100:0.0} m ({KgDesperdicio:0} kg), " +
                          $"de los cuales {RetazosCm / 100:0.0} m son retazos aprovechables (≥ {Opciones.RetazoMinCm / 100:0.##} m)");
            foreach (CorteDiametro d in Diametros)
            {
                sb.AppendLine();
                sb.AppendLine(d.Resumen());
                if (d.Empalmes > 0)
                    sb.AppendLine($"  {d.Empalmes} empalme(s) por piezas más largas que la barra comercial (traslape {Opciones.Traslape(d.Mm) / 100:0.00} m)");
                foreach (PatronCorte p in d.Patrones)
                    sb.AppendLine("  " + p.Texto(Opciones.KerfCm));
                if (d.Retazos.Count > 0)
                    sb.AppendLine("  Retazos: " + string.Join(", ", d.Retazos.Reverse().Select(kv => $"{kv.Value} × {kv.Key / 100:0.00} m")));
                foreach (string a in d.Avisos) sb.AppendLine("  ⚠ " + a);
            }
            if (Avisos.Count > 0)
            {
                sb.AppendLine();
                foreach (string a in Avisos) sb.AppendLine("⚠ " + a);
            }
            return sb.ToString().TrimEnd();
        }

        /// <summary>
        /// Planilla para Excel. Usa el separador de listas y el formato de números de Windows,
        /// para que Excel la abra directamente en el idioma del usuario.
        /// </summary>
        public string Csv(IEnumerable<PiezaCorte> piezas)
        {
            CultureInfo cu = CultureInfo.CurrentCulture;
            string sep = cu.TextInfo.ListSeparator;
            if (string.IsNullOrEmpty(sep)) sep = ";";
            string N(double v, string f = "0.00") => v.ToString(f, cu);
            string Q(string s) => "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
            string Fila(params string[] celdas) => string.Join(sep, celdas);

            var sb = new StringBuilder();
            sb.AppendLine(Fila(Q("ACERO - Optimización de corte de barras"), Q(DateTime.Now.ToString("g", cu))));
            sb.AppendLine(Fila(Q("Barras comerciales (m)"), Q(string.Join(", ", Opciones.ComercialesCm.Select(c => N(c / 100, "0.##"))))));
            sb.AppendLine(Fila(Q("Pérdida por corte (cm)"), N(Opciones.KerfCm, "0.#")));
            sb.AppendLine();

            sb.AppendLine(Fila(Q("RESUMEN"), "", "", "", "", "", "", ""));
            sb.AppendLine(Fila(Q("Diámetro"), Q("Piezas"), Q("Barras comerciales"), Q("Largo comercial (m)"), Q("Peso (kg)"),
                Q("Desperdicio (m)"), Q("Desperdicio (%)"), Q("Retazos aprovechables (m)")));
            foreach (CorteDiametro d in Diametros)
                sb.AppendLine(Fila(Q(d.Bar), d.PiezasCortadas.ToString(cu), d.BarrasComerciales.ToString(cu), N(d.LargoComercialCm / 100),
                    N(d.KgComercial, "0.0"), N(d.DesperdicioCm / 100), N(d.Porcentaje, "0.0"), N(d.RetazosCm / 100)));
            sb.AppendLine(Fila(Q("TOTAL"), Piezas.ToString(cu), BarrasComerciales.ToString(cu), N(LargoComercialCm / 100),
                N(KgComercial, "0.0"), N(DesperdicioCm / 100), N(Porcentaje, "0.0"), N(RetazosCm / 100)));
            sb.AppendLine();

            int maxPiezas = Diametros.SelectMany(d => d.Patrones).Select(p => p.Piezas.Count).DefaultIfEmpty(1).Max();
            sb.AppendLine(Fila(Q("PATRONES DE CORTE (longitudes en m)")));
            var cab = new List<string> { Q("Diámetro"), Q("Barra comercial"), Q("Veces") };
            for (int i = 1; i <= maxPiezas; i++) cab.Add(Q("Pieza " + i));
            cab.Add(Q("Sobra"));
            sb.AppendLine(Fila(cab.ToArray()));
            foreach (CorteDiametro d in Diametros)
                foreach (PatronCorte p in d.Patrones)
                {
                    var f = new List<string> { Q(d.Bar), N(p.ComercialCm / 100, "0.##"), p.Veces.ToString(cu) };
                    f.AddRange(p.Piezas.Select(x => N(x / 100)));
                    for (int i = p.Piezas.Count; i < maxPiezas; i++) f.Add("");
                    f.Add(N(p.SobraCm / 100));
                    sb.AppendLine(Fila(f.ToArray()));
                }
            sb.AppendLine();

            sb.AppendLine(Fila(Q("LISTA DE PIEZAS (tal como están en el modelo)")));
            sb.AppendLine(Fila(Q("Diámetro"), Q("Longitud (m)"), Q("Cantidad"), Q("Total (m)"), Q("Origen")));
            foreach (var g in piezas.Where(p => p.Cantidad > 0)
                         .GroupBy(p => (p.Bar, p.Mm, cm: Corte.Redondear(p.LargoCm, Opciones.RedondeoCm)))
                         .OrderBy(g => g.Key.Mm).ThenBy(g => g.Key.Bar).ThenByDescending(g => g.Key.cm))
            {
                int n = g.Sum(p => p.Cantidad);
                string origen = string.Join(", ", g.Select(p => p.Origen).Where(o => !string.IsNullOrEmpty(o)).Distinct().Take(4));
                sb.AppendLine(Fila(Q(g.Key.Bar), N(g.Key.cm / 100), n.ToString(cu), N(g.Key.cm / 100 * n), Q(origen)));
            }

            var avisos = Diametros.SelectMany(d => d.Avisos.Select(a => $"Ø {d.Bar}: {a}")).Concat(Avisos).ToList();
            if (avisos.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine(Fila(Q("AVISOS")));
                foreach (string a in avisos) sb.AppendLine(Q(a));
            }
            return sb.ToString();
        }
    }

    /// <summary>
    /// Optimizador de corte (problema de corte unidimensional). Para cada diámetro prueba varias heurísticas
    /// (primer ajuste y mejor ajuste decreciente, con cada longitud comercial como barra base) y se queda
    /// con la que menos acero comercial consume. Luego cada barra se reduce a la longitud comercial más corta
    /// en la que caben sus piezas.
    /// </summary>
    public static class Corte
    {
        private sealed class Barra
        {
            public double Stock;
            public double Usado;        // piezas + pérdidas de corte
            public readonly List<double> Piezas = new List<double>();
            public double Libre => Stock - Usado;
        }

        public static double Redondear(double cm, double paso)
        {
            if (paso <= 0) return cm;
            return Math.Ceiling(cm / paso - 1e-9) * paso;
        }

        /// <summary>Interpreta "9, 12" (metros) o "900 1200" (centímetros) como longitudes comerciales en cm.</summary>
        public static double[] ParseComerciales(string texto)
        {
            var list = new List<double>();
            foreach (string tok in (texto ?? "").Split(new[] { ',', ';', ' ', '\t', '/', '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string t = tok.Replace("m", "").Trim();
                if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out double v) &&
                    !double.TryParse(t, NumberStyles.Float, CultureInfo.CurrentCulture, out v))
                    throw new FormatException($"No se entiende la longitud comercial \"{tok}\". Escriba por ejemplo: 9, 12");
                if (v <= 0) continue;
                list.Add(v < 50 ? v * 100 : v);
            }
            if (list.Count == 0) throw new FormatException("Indique al menos una longitud comercial, por ejemplo: 9, 12");
            return list.Distinct().OrderBy(x => x).ToArray();
        }

        public static ResultadoCorte Optimizar(IEnumerable<PiezaCorte> piezas, OpcionesCorte o)
        {
            var res = new ResultadoCorte { Opciones = o };
            double[] stocks = (o.ComercialesCm ?? new double[0]).Where(s => s > 0).Distinct().OrderBy(s => s).ToArray();
            if (stocks.Length == 0) throw new ArgumentException("Indique al menos una longitud comercial.");
            double smax = stocks[stocks.Length - 1];
            double kerf = Math.Max(0, o.KerfCm);

            foreach (var grupo in piezas.Where(p => p != null && p.Cantidad > 0 && p.LargoCm > 0)
                         .GroupBy(p => p.Bar ?? "?").OrderBy(g => g.First().Mm).ThenBy(g => g.Key))
            {
                var d = new CorteDiametro { Bar = grupo.Key, Mm = grupo.First().Mm };
                double traslape = o.Traslape(d.Mm);

                // 1. Demanda: longitud (redondeada hacia arriba) → cantidad. Las piezas largas se parten con traslape.
                var demanda = new Dictionary<double, int>();
                foreach (PiezaCorte p in grupo)
                {
                    double L = Redondear(p.LargoCm, o.RedondeoCm);
                    d.PiezasPedidas += p.Cantidad;
                    if (L + kerf <= smax + 1e-6)
                    {
                        Sumar(demanda, L, p.Cantidad);
                        continue;
                    }
                    if (!o.Empalmar)
                    {
                        d.PiezasOmitidas += p.Cantidad;
                        d.Avisos.Add($"{p.Cantidad} pieza(s) de {L / 100:0.00} m superan la barra comercial de {smax / 100:0.##} m y no se incluyen (active los empalmes).");
                        continue;
                    }
                    double util = smax - kerf;                 // lo que aprovecha cada tramo comercial
                    double neto = util - traslape;             // avance real de cada tramo completo
                    if (neto <= 0)
                        throw new ArgumentException("El traslape es mayor que la barra comercial.");
                    int k = (int)Math.Ceiling((L - util) / neto - 1e-9);
                    double resto = Redondear(L + k * traslape - k * util, o.RedondeoCm);
                    if (resto + kerf > smax + 1e-6) { k++; resto = Redondear(L + k * traslape - k * util, o.RedondeoCm); }
                    Sumar(demanda, util, k * p.Cantidad);
                    if (resto > 0) Sumar(demanda, resto, p.Cantidad);
                    d.Empalmes += k * p.Cantidad;
                    d.LargoTraslapesCm += k * traslape * p.Cantidad;
                }
                if (demanda.Count == 0)
                {
                    if (d.Avisos.Count > 0) res.Diametros.Add(d);
                    continue;
                }

                // 2. Varias heurísticas; se queda con la de menor acero comercial.
                List<double> lista = Expandir(demanda);
                double mayor = lista[0];
                List<Barra> mejor = null;
                double mejorTotal = double.MaxValue;
                foreach (double s in stocks)
                {
                    if (mayor + kerf > s + 1e-6) continue;
                    foreach (bool bestFit in new[] { false, true })
                    {
                        List<Barra> barras = Empacar(lista, s, kerf, bestFit);
                        Encoger(barras, stocks);
                        double total = barras.Sum(b => b.Stock);
                        if (mejor == null || total < mejorTotal - 1e-6 ||
                            (Math.Abs(total - mejorTotal) < 1e-6 && barras.Count < mejor.Count))
                        {
                            mejor = barras;
                            mejorTotal = total;
                        }
                    }
                }
                if (mejor == null)
                {
                    d.Avisos.Add("No hay una longitud comercial en la que quepan las piezas.");
                    res.Diametros.Add(d);
                    continue;
                }

                // 3. Patrones (barras iguales agrupadas), barras por longitud comercial y retazos.
                foreach (var g in mejor.GroupBy(b => b.Stock.ToString("0.0", CultureInfo.InvariantCulture) + "|" +
                                                     string.Join(",", b.Piezas.Select(x => x.ToString("0.0", CultureInfo.InvariantCulture))))
                             .OrderByDescending(g => g.Count()).ThenByDescending(g => g.First().Piezas.Sum()))
                {
                    Barra b0 = g.First();
                    var p = new PatronCorte
                    {
                        ComercialCm = b0.Stock,
                        Piezas = b0.Piezas.ToList(),
                        Veces = g.Count(),
                        SobraCm = Math.Max(0, b0.Stock - b0.Piezas.Sum() - kerf * b0.Piezas.Count)
                    };
                    d.Patrones.Add(p);
                    Sumar(d.BarrasPorComercial, p.ComercialCm, p.Veces);
                    if (p.SobraCm >= o.RetazoMinCm && o.RetazoMinCm > 0)
                        Sumar(d.Retazos, Math.Floor(p.SobraCm), p.Veces);
                }
                d.PiezasCortadas = lista.Count;
                res.Diametros.Add(d);
            }

            if (res.Diametros.Count == 0) res.Avisos.Add("No hay piezas que cortar.");
            return res;
        }

        private static void Sumar<TKey>(IDictionary<TKey, int> dic, TKey key, int n)
        {
            dic.TryGetValue(key, out int v);
            dic[key] = v + n;
        }

        private static List<double> Expandir(Dictionary<double, int> demanda)
        {
            var lista = new List<double>();
            foreach (var kv in demanda.OrderByDescending(kv => kv.Key))
                for (int i = 0; i < kv.Value; i++) lista.Add(kv.Key);
            return lista;
        }

        /// <summary>Primer ajuste decreciente (o mejor ajuste decreciente) con una sola longitud de barra.</summary>
        private static List<Barra> Empacar(List<double> piezasDesc, double stock, double kerf, bool bestFit)
        {
            var barras = new List<Barra>();
            foreach (double L in piezasDesc)
            {
                double need = L + kerf;
                Barra dest = null;
                if (bestFit)
                {
                    double libre = double.MaxValue;
                    foreach (Barra b in barras)
                        if (b.Libre + 1e-6 >= need && b.Libre < libre) { dest = b; libre = b.Libre; }
                }
                else
                {
                    foreach (Barra b in barras)
                        if (b.Libre + 1e-6 >= need) { dest = b; break; }
                }
                if (dest == null)
                {
                    dest = new Barra { Stock = stock };
                    barras.Add(dest);
                }
                dest.Piezas.Add(L);
                dest.Usado += need;
            }
            return barras;
        }

        /// <summary>Cada barra pasa a la longitud comercial más corta en la que caben sus piezas.</summary>
        private static void Encoger(List<Barra> barras, double[] stocksAsc)
        {
            foreach (Barra b in barras)
                foreach (double s in stocksAsc)
                    if (b.Usado <= s + 1e-6) { b.Stock = s; break; }
        }
    }
}
