using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace AceroRefuerzo.Core
{
    /// <summary>
    /// Sistema de coordenadas local de un elemento con las dimensiones reales de su geometría
    /// (mínimos y máximos medidos sobre los sólidos, ya recortados por las uniones).
    /// </summary>
    internal sealed class LocalFrame
    {
        public XYZ Origin { get; }
        public XYZ X { get; }
        public XYZ Y { get; }
        public XYZ Z { get; }

        public double MinX = double.MaxValue, MaxX = double.MinValue;
        public double MinY = double.MaxValue, MaxY = double.MinValue;
        public double MinZ = double.MaxValue, MaxZ = double.MinValue;

        public double SizeX => MaxX - MinX;
        public double SizeY => MaxY - MinY;
        public double SizeZ => MaxZ - MinZ;

        private LocalFrame(XYZ origin, XYZ x, XYZ y, XYZ z)
        {
            Origin = origin;
            X = x.Normalize();
            Y = y.Normalize();
            Z = z.Normalize();
        }

        /// <summary>Punto global a partir de coordenadas locales.</summary>
        public XYZ P(double x, double y, double z) => Origin + X * x + Y * y + Z * z;

        /// <summary>Coordenadas locales de un punto global.</summary>
        public XYZ ToLocal(XYZ p)
        {
            XYZ v = p - Origin;
            return new XYZ(v.DotProduct(X), v.DotProduct(Y), v.DotProduct(Z));
        }

        private void Add(XYZ p)
        {
            XYZ l = ToLocal(p);
            MinX = Math.Min(MinX, l.X); MaxX = Math.Max(MaxX, l.X);
            MinY = Math.Min(MinY, l.Y); MaxY = Math.Max(MaxY, l.Y);
            MinZ = Math.Min(MinZ, l.Z); MaxZ = Math.Max(MaxZ, l.Z);
        }

        public static LocalFrame Measure(Element e, XYZ origin, XYZ x, XYZ y, XYZ z)
        {
            var f = new LocalFrame(origin, x, y, z);
            int count = 0;
            foreach (Solid s in Geo.Solids(e))
                foreach (Edge ed in s.Edges)
                    foreach (XYZ p in ed.Tessellate())
                    {
                        f.Add(p);
                        count++;
                    }

            if (count == 0)
            {
                BoundingBoxXYZ bb = e.get_BoundingBox(null)
                    ?? throw new InvalidOperationException("El elemento no tiene geometría.");
                for (int i = 0; i < 8; i++)
                    f.Add(bb.Transform.OfPoint(new XYZ(
                        (i & 1) == 0 ? bb.Min.X : bb.Max.X,
                        (i & 2) == 0 ? bb.Min.Y : bb.Max.Y,
                        (i & 4) == 0 ? bb.Min.Z : bb.Max.Z)));
            }
            return f;
        }
    }

    internal static class Geo
    {
        public static IEnumerable<Solid> Solids(Element e)
        {
            var opt = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false };
            GeometryElement ge = e.get_Geometry(opt);
            if (ge == null) yield break;
            foreach (Solid s in Solids(ge)) yield return s;
        }

        private static IEnumerable<Solid> Solids(GeometryElement ge)
        {
            foreach (GeometryObject o in ge)
            {
                if (o is Solid s && s.Volume > 1e-9)
                    yield return s;
                else if (o is GeometryInstance gi)
                    foreach (Solid si in Solids(gi.GetInstanceGeometry()))
                        yield return si;
            }
        }

        /// <summary>Lista de segmentos rectos entre puntos consecutivos (omite tramos nulos).</summary>
        public static IList<Curve> Poly(params XYZ[] pts)
        {
            var list = new List<Curve>();
            double tol = Un.Mm(1);
            for (int i = 0; i + 1 < pts.Length; i++)
                if (pts[i].DistanceTo(pts[i + 1]) > tol)
                    list.Add(Line.CreateBound(pts[i], pts[i + 1]));
            return list;
        }

        /// <summary>Vector horizontal unitario (proyección en XY); null si es vertical.</summary>
        public static XYZ Horizontal(XYZ v)
        {
            var h = new XYZ(v.X, v.Y, 0);
            return h.GetLength() < 1e-9 ? null : h.Normalize();
        }

        /// <summary>Área con signo de un polígono 2D (X,Y).</summary>
        public static double SignedArea(IList<XYZ> pts)
        {
            double a = 0;
            for (int i = 0; i < pts.Count; i++)
            {
                XYZ p = pts[i], q = pts[(i + 1) % pts.Count];
                a += p.X * q.Y - q.X * p.Y;
            }
            return a / 2.0;
        }

        /// <summary>Punto dentro de un polígono 2D (X,Y) por cruce de rayos.</summary>
        public static bool Contains(IList<XYZ> poly, double x, double y)
        {
            bool inside = false;
            for (int i = 0, j = poly.Count - 1; i < poly.Count; j = i++)
            {
                XYZ a = poly[i], b = poly[j];
                if ((a.Y > y) != (b.Y > y) && x < (b.X - a.X) * (y - a.Y) / (b.Y - a.Y) + a.X)
                    inside = !inside;
            }
            return inside;
        }
    }
}
