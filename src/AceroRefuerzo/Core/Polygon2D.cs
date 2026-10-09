using System;
using System.Collections.Generic;
using System.Linq;
using AceroRefuerzo.UI;

namespace AceroRefuerzo.Core
{
    /// <summary>Punto 2D sencillo (sin dependencias de Revit, se usa también para dibujar).</summary>
    public struct P2
    {
        public double X, Y;
        public P2(double x, double y) { X = x; Y = y; }
        public static P2 operator +(P2 a, P2 b) => new P2(a.X + b.X, a.Y + b.Y);
        public static P2 operator -(P2 a, P2 b) => new P2(a.X - b.X, a.Y - b.Y);
        public static P2 operator *(P2 a, double k) => new P2(a.X * k, a.Y * k);
        public double Length => Math.Sqrt(X * X + Y * Y);
        public P2 Unit { get { double l = Length; return l < 1e-12 ? new P2(0, 0) : new P2(X / l, Y / l); } }
    }

    /// <summary>Utilidades para secciones poligonales (columnas irregulares).</summary>
    public static class Polygon2D
    {
        public static double Area(IList<P2> p)
        {
            double a = 0;
            for (int i = 0; i < p.Count; i++)
            {
                P2 u = p[i], w = p[(i + 1) % p.Count];
                a += u.X * w.Y - w.X * u.Y;
            }
            return a / 2.0;
        }

        public static P2 Centroid(IList<P2> p)
        {
            double a = Area(p);
            if (Math.Abs(a) < 1e-12) return new P2(p.Average(q => q.X), p.Average(q => q.Y));
            double cx = 0, cy = 0;
            for (int i = 0; i < p.Count; i++)
            {
                P2 u = p[i], w = p[(i + 1) % p.Count];
                double k = u.X * w.Y - w.X * u.Y;
                cx += (u.X + w.X) * k;
                cy += (u.Y + w.Y) * k;
            }
            return new P2(cx / (6 * a), cy / (6 * a));
        }

        public static bool Contains(IList<P2> poly, P2 q)
        {
            bool inside = false;
            for (int i = 0, j = poly.Count - 1; i < poly.Count; j = i++)
            {
                P2 a = poly[i], b = poly[j];
                if ((a.Y > q.Y) != (b.Y > q.Y) && q.X < (b.X - a.X) * (q.Y - a.Y) / (b.Y - a.Y) + a.X)
                    inside = !inside;
            }
            return inside;
        }

        /// <summary>Quita vértices repetidos y colineales.</summary>
        public static List<P2> Clean(IList<P2> poly, double tol)
        {
            var pts = new List<P2>();
            foreach (P2 p in poly)
                if (pts.Count == 0 || (p - pts[pts.Count - 1]).Length > tol) pts.Add(p);
            if (pts.Count > 1 && (pts[0] - pts[pts.Count - 1]).Length <= tol) pts.RemoveAt(pts.Count - 1);

            bool changed = true;
            while (changed && pts.Count > 3)
            {
                changed = false;
                for (int i = 0; i < pts.Count; i++)
                {
                    P2 a = pts[(i - 1 + pts.Count) % pts.Count], b = pts[i], c = pts[(i + 1) % pts.Count];
                    double cross = (b - a).Unit.X * (c - b).Unit.Y - (b - a).Unit.Y * (c - b).Unit.X;
                    double dot = (b - a).Unit.X * (c - b).Unit.X + (b - a).Unit.Y * (c - b).Unit.Y;
                    if (Math.Abs(cross) < 1e-4 && dot > 0) { pts.RemoveAt(i); changed = true; break; }
                }
            }
            return pts;
        }

        /// <summary>Desfase hacia el interior una distancia d (intersección de bordes desplazados).</summary>
        public static List<P2> OffsetInward(IList<P2> poly, double d)
        {
            int n = poly.Count;
            var result = new List<P2>(n);
            double sign = Area(poly) > 0 ? 1 : -1; // antihorario: el interior queda a la izquierda
            for (int i = 0; i < n; i++)
            {
                P2 a = poly[(i - 1 + n) % n], b = poly[i], c = poly[(i + 1) % n];
                P2 t1 = (b - a).Unit, t2 = (c - b).Unit;
                P2 n1 = new P2(-t1.Y, t1.X) * sign, n2 = new P2(-t2.Y, t2.X) * sign;
                P2 p1 = a + n1 * d, p2 = b + n2 * d;
                double den = t1.X * t2.Y - t1.Y * t2.X;
                if (Math.Abs(den) < 1e-6)
                {
                    result.Add(b + n1 * d);
                    continue;
                }
                double s = ((p2.X - p1.X) * t2.Y - (p2.Y - p1.Y) * t2.X) / den;
                P2 q = p1 + t1 * s;
                // Limita picos excesivos en ángulos muy agudos.
                if ((q - b).Length > 4 * Math.Abs(d)) q = b + (n1 + n2).Unit * Math.Abs(d) * Math.Sign(d);
                result.Add(q);
            }
            return result;
        }

        /// <summary>
        /// Posiciones de las barras longitudinales sobre un contorno: una en cada esquina, una en cada punto
        /// obligatorio de <paramref name="anchors"/> (por ejemplo, las esquinas de los estribos de cada ala) y
        /// barras intermedias con separación no mayor que <paramref name="maxSpacing"/>. Para contornos sin
        /// esquinas (circulares) reparte uniformemente con un mínimo de <paramref name="minBars"/> barras.
        /// </summary>
        public static List<P2> PerimeterPoints(IList<P2> poly, double maxSpacing, int minBars,
            IList<P2> anchors = null, double cornerDeg = 25)
        {
            int n = poly.Count;
            var result = new List<P2>();
            if (n < 2 || maxSpacing <= 0) return result;

            // Longitud acumulada del contorno en cada vértice.
            var cum = new double[n + 1];
            for (int i = 0; i < n; i++) cum[i + 1] = cum[i] + (poly[(i + 1) % n] - poly[i]).Length;
            double per = cum[n];
            if (per <= 0) return result;
            double tol = per * 1e-4;

            var must = new List<double>();
            for (int i = 0; i < n; i++)
            {
                P2 a = poly[(i - 1 + n) % n], b = poly[i], c = poly[(i + 1) % n];
                P2 t1 = (b - a).Unit, t2 = (c - b).Unit;
                double cos = Math.Max(-1, Math.Min(1, t1.X * t2.X + t1.Y * t2.Y));
                if (Math.Acos(cos) * 180 / Math.PI > cornerDeg) must.Add(cum[i]);
            }

            if (anchors != null)
                foreach (P2 q in anchors)
                {
                    double best = double.MaxValue, at = 0;
                    for (int i = 0; i < n; i++)
                    {
                        P2 a = poly[i], b = poly[(i + 1) % n], ab = b - a;
                        double l2 = ab.X * ab.X + ab.Y * ab.Y;
                        if (l2 < 1e-18) continue;
                        double t = Math.Max(0, Math.Min(1, ((q.X - a.X) * ab.X + (q.Y - a.Y) * ab.Y) / l2));
                        double d = (a + ab * t - q).Length;
                        if (d < best) { best = d; at = cum[i] + t * Math.Sqrt(l2); }
                    }
                    if (best < per * 2e-3) must.Add(at);
                }

            if (must.Count == 0)
            {
                int count = Math.Max(minBars, (int)Math.Ceiling(per / maxSpacing - 1e-6));
                for (int k = 0; k < count; k++) result.Add(PointAt(poly, cum, per * k / count));
                return result;
            }

            must.Sort();
            var marks = new List<double>();
            foreach (double m in must)
                if (marks.Count == 0 || m - marks[marks.Count - 1] > tol) marks.Add(m);
            if (marks.Count > 1 && marks[0] + per - marks[marks.Count - 1] <= tol) marks.RemoveAt(marks.Count - 1);

            for (int k = 0; k < marks.Count; k++)
            {
                double s0 = marks[k];
                double s1 = k + 1 < marks.Count ? marks[k + 1] : marks[0] + per;
                double len = s1 - s0;
                int seg = Math.Max(1, (int)Math.Ceiling(len / maxSpacing - 1e-6));
                for (int j2 = 0; j2 < seg; j2++) result.Add(PointAt(poly, cum, (s0 + len * j2 / seg) % per));
            }
            return result;
        }

        private static P2 PointAt(IList<P2> poly, double[] cum, double s)
        {
            int n = poly.Count;
            for (int i = 0; i < n; i++)
            {
                double l = cum[i + 1] - cum[i];
                if (s <= cum[i + 1] || i == n - 1)
                    return l < 1e-12 ? poly[i] : poly[i] + (poly[(i + 1) % n] - poly[i]) * (Math.Max(0, Math.Min(l, s - cum[i])) / l);
            }
            return poly[0];
        }

        /// <summary>
        /// Rectángulos (x0, y0, x1, y1) que cubren una sección ortogonal (L, T, cruz, U...), uno por cada ala.
        /// Son rectángulos máximos que se traslapan en las uniones, como los estribos de cada ala.
        /// Devuelve null si la sección no es ortogonal (circular, poligonal inclinada).
        /// </summary>
        public static List<(double x0, double y0, double x1, double y1)> CoverRectangles(IList<P2> poly)
        {
            int n = poly.Count;
            if (n < 4) return null;
            double size = Math.Max(poly.Max(p => p.X) - poly.Min(p => p.X), poly.Max(p => p.Y) - poly.Min(p => p.Y));
            double tol = size * 1e-4;
            for (int i = 0; i < n; i++)
            {
                P2 d = poly[(i + 1) % n] - poly[i];
                if (Math.Abs(d.X) > tol && Math.Abs(d.Y) > tol) return null;
            }

            List<double> xs = Unique(poly.Select(p => p.X), tol), ys = Unique(poly.Select(p => p.Y), tol);
            int nx = xs.Count - 1, ny = ys.Count - 1;
            if (nx < 1 || ny < 1) return null;
            var inside = new bool[nx, ny];
            for (int i = 0; i < nx; i++)
                for (int j = 0; j < ny; j++)
                    inside[i, j] = Contains(poly, new P2((xs[i] + xs[i + 1]) / 2, (ys[j] + ys[j + 1]) / 2));

            bool Full(int i0, int j0, int i1, int j1)
            {
                if (i0 < 0 || j0 < 0 || i1 >= nx || j1 >= ny) return false;
                for (int i = i0; i <= i1; i++)
                    for (int j = j0; j <= j1; j++)
                        if (!inside[i, j]) return false;
                return true;
            }

            // Rectángulos máximos: llenos y que no se pueden ampliar hacia ningún lado.
            var maximos = new List<(int i0, int j0, int i1, int j1)>();
            for (int i0 = 0; i0 < nx; i0++)
                for (int i1 = i0; i1 < nx; i1++)
                    for (int j0 = 0; j0 < ny; j0++)
                        for (int j1 = j0; j1 < ny; j1++)
                            if (Full(i0, j0, i1, j1) && !Full(i0 - 1, j0, i1, j1) && !Full(i0, j0, i1 + 1, j1)
                                && !Full(i0, j0 - 1, i1, j1) && !Full(i0, j0, i1, j1 + 1))
                                maximos.Add((i0, j0, i1, j1));

            // Cobertura voraz: en cada paso el rectángulo que cubre más celdas libres (y luego el de mayor área).
            var libres = new HashSet<(int, int)>();
            for (int i = 0; i < nx; i++)
                for (int j = 0; j < ny; j++)
                    if (inside[i, j]) libres.Add((i, j));

            var result = new List<(double, double, double, double)>();
            while (libres.Count > 0)
            {
                int bestCount = 0;
                double bestArea = 0;
                (int i0, int j0, int i1, int j1) best = default;
                foreach (var r in maximos)
                {
                    int c = 0;
                    for (int i = r.i0; i <= r.i1; i++)
                        for (int j = r.j0; j <= r.j1; j++)
                            if (libres.Contains((i, j))) c++;
                    double area = (xs[r.i1 + 1] - xs[r.i0]) * (ys[r.j1 + 1] - ys[r.j0]);
                    if (c > bestCount || (c == bestCount && c > 0 && area > bestArea)) { bestCount = c; bestArea = area; best = r; }
                }
                if (bestCount == 0) break;
                for (int i = best.i0; i <= best.i1; i++)
                    for (int j = best.j0; j <= best.j1; j++)
                        libres.Remove((i, j));
                result.Add((xs[best.i0], ys[best.j0], xs[best.i1 + 1], ys[best.j1 + 1]));
            }
            return result;
        }

        /// <summary>Esquinas de los rectángulos reducidos hacia adentro una distancia <paramref name="inset"/>.</summary>
        public static List<P2> RectCorners(IEnumerable<(double x0, double y0, double x1, double y1)> rects, double inset)
        {
            var pts = new List<P2>();
            foreach (var r in rects)
            {
                pts.Add(new P2(r.x0 + inset, r.y0 + inset));
                pts.Add(new P2(r.x1 - inset, r.y0 + inset));
                pts.Add(new P2(r.x1 - inset, r.y1 - inset));
                pts.Add(new P2(r.x0 + inset, r.y1 - inset));
            }
            return pts;
        }

        private static List<double> Unique(IEnumerable<double> values, double tol)
        {
            var list = new List<double>();
            foreach (double v in values.OrderBy(v => v))
                if (list.Count == 0 || v - list[list.Count - 1] > tol) list.Add(v);
            return list;
        }
    }
}
