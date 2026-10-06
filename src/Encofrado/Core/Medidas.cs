using System;
using Autodesk.Revit.DB;
using Comun;

namespace Encofrado.Core
{
    /// <summary>Medidas principales (cm) del elemento de muestra, para dibujar su figura.</summary>
    internal struct Medidas
    {
        public double A, B, C; // columna: bx, by, H · viga: b, h, L · zapata: Lx, Ly, H · losa: Lx, Ly, e · muro: L, t, H · escalera: L, ancho, H

        public static Medidas De(Element e, Pieza p)
        {
            LocalFrame f = Marco(e, p);
            double x = Un.ToCm(f.SizeX), y = Un.ToCm(f.SizeY), z = Un.ToCm(f.SizeZ);
            switch (p)
            {
                case Pieza.Viga: return new Medidas { A = y, B = z, C = x };
                case Pieza.Muro: return new Medidas { A = x, B = y, C = z };
                case Pieza.Escalera: return new Medidas { A = Math.Max(x, y), B = Math.Min(x, y), C = z };
                default: return new Medidas { A = x, B = y, C = z };
            }
        }

        private static LocalFrame Marco(Element e, Pieza p)
        {
            if ((e.Location as LocationCurve)?.Curve is Line line)
            {
                XYZ x = Geo.Horizontal(line.GetEndPoint(1) - line.GetEndPoint(0));
                if (x != null)
                {
                    XYZ y = XYZ.BasisZ.CrossProduct(x).Normalize();
                    return LocalFrame.Measure(e, line.GetEndPoint(0), x, y, XYZ.BasisZ);
                }
            }
            if (e is FamilyInstance fi)
            {
                Transform t = fi.GetTransform();
                XYZ x = Geo.Horizontal(t.BasisX) ?? XYZ.BasisX;
                return LocalFrame.Measure(e, t.Origin, x, XYZ.BasisZ.CrossProduct(x), XYZ.BasisZ);
            }
            return LocalFrame.Measure(e, XYZ.Zero, XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ);
        }
    }
}
