using System;
using System.Collections.Generic;
using System.Linq;

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
        /// Posiciones de las barras longitudinales sobre un contorno: una en cada esquina y barras intermedias
        /// en cada lado con separación no mayor que <paramref name="maxSpacing"/>. Para contornos sin esquinas
        /// (circulares) reparte uniformemente con un mínimo de <paramref name="minBars"/> barras.
        /// </summary>
        public static List<P2> PerimeterPoints(IList<P2> poly, double maxSpacing, int minBars, double cornerDeg = 25)
        {
            int n = poly.Count;
            var result = new List<P2>();
            if (n < 2 || maxSpacing <= 0) return result;

            var corners = new List<int>();
            for (int i = 0; i < n; i++)
            {
                P2 a = poly[(i - 1 + n) % n], b = poly[i], c = poly[(i + 1) % n];
                P2 t1 = (b - a).Unit, t2 = (c - b).Unit;
                double cos = Math.Max(-1, Math.Min(1, t1.X * t2.X + t1.Y * t2.Y));
                if (Math.Acos(cos) * 180 / Math.PI > cornerDeg) corners.Add(i);
            }

            if (corners.Count == 0)
            {
                double per = 0;
                for (int i = 0; i < n; i++) per += (poly[(i + 1) % n] - poly[i]).Length;
                int count = Math.Max(minBars, (int)Math.Ceiling(per / maxSpacing - 1e-6));
                for (int k = 0; k < count; k++) result.Add(AlongPath(poly, 0, n, per * k / count));
                return result;
            }

            for (int ci = 0; ci < corners.Count; ci++)
            {
                int start = corners[ci];
                int end = corners[(ci + 1) % corners.Count];
                int steps = (end - start + n) % n;
                if (steps == 0) steps = n;
                double len = 0;
                for (int k = 0; k < steps; k++) len += (poly[(start + k + 1) % n] - poly[(start + k) % n]).Length;
                int seg = Math.Max(1, (int)Math.Ceiling(len / maxSpacing - 1e-6));
                for (int k = 0; k < seg; k++) result.Add(AlongPath(poly, start, steps, len * k / seg));
            }
            return result;
        }

        private static P2 AlongPath(IList<P2> poly, int start, int steps, double dist)
        {
            int n = poly.Count;
            for (int k = 0; k < steps; k++)
            {
                P2 a = poly[(start + k) % n], b = poly[(start + k + 1) % n];
                double l = (b - a).Length;
                if (dist <= l || k == steps - 1) return l < 1e-12 ? a : a + (b - a) * (Math.Min(dist, l) / l);
                dist -= l;
            }
            return poly[start % n];
        }
    }
}
