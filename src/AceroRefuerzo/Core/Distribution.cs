using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Comun;

namespace AceroRefuerzo.Core
{
    /// <summary>
    /// Distribución de estribos en la notación usada en planos:
    ///   "1@5, 10@10, R@25"  →  1 a 5 cm, 10 a 10 cm y el resto a 25 cm,
    /// aplicada simétricamente desde ambos extremos del elemento.
    /// Las separaciones van en centímetros; si todas son menores que 1 se interpretan en metros
    /// (por ejemplo "1@0.05, 10@0.10, R@0.25").
    /// </summary>
    internal static class Distribution
    {
        internal sealed class Item
        {
            public int Count;
            public double SpacingCm;
            public bool Rest;
        }

        /// <summary>Grupo de barras equiespaciadas (se crea como un solo conjunto de Revit).</summary>
        internal sealed class Run
        {
            public double Start;
            public int Count;
            public double Spacing;
            public double Length => Count > 1 ? (Count - 1) * Spacing : 0.0;
        }

        public static List<Item> Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new FormatException("La distribución de estribos está vacía.");

            var items = new List<Item>();
            foreach (string raw in text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string t = raw.Replace(" ", "").Trim();
                if (t.Length == 0) continue;

                int at = t.IndexOf('@');
                if (at <= 0 || at == t.Length - 1)
                    throw new FormatException($"Término inválido \"{raw.Trim()}\". Use n@s, por ejemplo 10@10 o R@20.");

                string left = t.Substring(0, at);
                string right = t.Substring(at + 1);
                if (!double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out double s) || s <= 0)
                    throw new FormatException($"Separación inválida en \"{raw.Trim()}\" (use punto decimal).");

                var item = new Item { SpacingCm = s };
                string l = left.ToUpperInvariant();
                if (l == "R" || l == "RTO" || l == "RESTO" || l == "REST")
                    item.Rest = true;
                else if (int.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) && n > 0)
                    item.Count = n;
                else
                    throw new FormatException($"Cantidad inválida en \"{raw.Trim()}\".");

                items.Add(item);
                if (item.Rest) break; // "R" siempre cierra la distribución
            }

            if (items.Count == 0)
                throw new FormatException("La distribución de estribos está vacía.");

            if (items.All(i => i.SpacingCm < 1.0))
                foreach (var i in items) i.SpacingCm *= 100.0;

            return items;
        }

        /// <summary>Valida el texto de la distribución: devuelve el mensaje de error o null.</summary>
        public static string Check(string text)
        {
            try { Parse(text); return null; }
            catch (FormatException ex) { return ex.Message; }
        }

        /// <summary>Posiciones (en pies, medidas desde la cara inicial) de todos los estribos.</summary>
        public static List<double> Positions(string text, double length)
        {
            var items = Parse(text);
            double tol = Un.Mm(1);
            double half = length / 2.0;

            var left = new List<double>();
            double acc = 0, lastSpacing = 0;
            double? rest = null;
            bool full = false;

            foreach (var it in items)
            {
                double s = Un.Cm(it.SpacingCm);
                if (it.Rest) { rest = s; break; }
                lastSpacing = s;
                for (int i = 0; i < it.Count && !full; i++)
                {
                    acc += s;
                    if (acc > half + tol) { full = true; break; }
                    left.Add(acc);
                }
                if (full) break;
            }

            if (!rest.HasValue && lastSpacing > 0) rest = lastSpacing;
            if (left.Count == 0) left.Add(Math.Min(Un.Cm(5), half));

            double a = left[left.Count - 1];
            double b = length - a;
            var pos = new List<double>(left);

            if (rest.HasValue && b - a > rest.Value * 0.5 + tol)
            {
                int n = (int)Math.Ceiling((b - a) / rest.Value - 1e-6);
                double step = (b - a) / n;
                for (int i = 1; i < n; i++) pos.Add(a + i * step);
            }

            foreach (double p in left)
            {
                double q = length - p;
                if (q > a + tol) pos.Add(q);
            }

            pos.Sort();
            var result = new List<double>();
            foreach (double p in pos)
                if (result.Count == 0 || p - result[result.Count - 1] > tol)
                    result.Add(p);
            return result;
        }

        /// <summary>Agrupa posiciones consecutivas con la misma separación.</summary>
        public static List<Run> Group(IList<double> positions)
        {
            double tol = Un.Mm(2);
            var runs = new List<Run>();
            Run cur = null;
            double prev = 0;

            foreach (double p in positions)
            {
                if (cur == null)
                {
                    cur = new Run { Start = p, Count = 1 };
                }
                else if (cur.Count == 1)
                {
                    cur.Spacing = p - prev;
                    cur.Count = 2;
                }
                else if (Math.Abs(p - prev - cur.Spacing) < tol)
                {
                    cur.Count++;
                }
                else
                {
                    runs.Add(cur);
                    cur = new Run { Start = p, Count = 1 };
                }
                prev = p;
            }
            if (cur != null) runs.Add(cur);
            return runs;
        }
    }
}
