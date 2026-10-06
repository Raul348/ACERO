using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Comun;

namespace Encofrado.Core
{
    /// <summary>Opciones de cálculo del encofrado.</summary>
    internal sealed class Opciones
    {
        public bool Lateral = true;
        public bool Fondo = true;
        public bool OmitirBase;
        public bool Descontar = true;
        public double Espesor = Un.Cm(1.8);
    }

    /// <summary>Tablero de encofrado de una cara (o grupo de caras) de un elemento.</summary>
    internal sealed class Tablero
    {
        public Cara Cara;
        public Solid Solido;
        public double Area; // pies²
    }

    /// <summary>
    /// Genera los tableros de encofrado de un elemento:
    ///  1. clasifica sus caras (laterales, inferiores, superiores);
    ///  2. extruye cada cara hacia afuera con el espesor del tablero;
    ///  3. resta los sólidos de los elementos vecinos (zonas en contacto: no llevan encofrado);
    ///  4. el área neta es volumen / espesor.
    /// </summary>
    internal static class Constructor
    {
        private const double Vertical = 0.17; // |nz| menor que esto → cara vertical (± 10°)

        /// <summary>Área bruta (sin descuentos) por tipo de cara, para la figura del formulario.</summary>
        public static (double lateral, double fondo) AreaBruta(Element host, Pieza p)
        {
            var o = new Opciones { OmitirBase = Piezas.OmitirBase(p) };
            double lat = 0, fon = 0;
            foreach (var (face, cara) in Caras(host, o, null))
                if (cara == Cara.Lateral) lat += face.Area; else fon += face.Area;
            return (lat, fon);
        }

        public static List<Tablero> Construir(Document doc, Element host, Opciones o, Vecinos vecinos,
            ElementId material, List<string> avisos)
        {
            var result = new List<Tablero>();
            var so = new SolidOptions(material ?? ElementId.InvalidElementId, ElementId.InvalidElementId);
            List<Solid> propios = Geo.Solids(host).ToList();
            List<Solid> vecinosSol = o.Descontar ? vecinos.Alrededor(host, o.Espesor) : new List<Solid>();
            int curvas = 0;

            foreach (var (face, cara) in Caras(host, o, () => curvas++))
            {
                Solid panel = null;
                try
                {
                    if (face is PlanarFace pf) panel = ExtruirPlana(pf, o.Espesor, so, propios);
                    else if (face is CylindricalFace cf) panel = ExtruirCilindro(cf, o.Espesor, so);
                }
                catch (Exception) { panel = null; }

                if (panel == null || panel.Volume < 1e-9) { curvas++; continue; }

                foreach (Solid v in vecinosSol)
                {
                    try
                    {
                        Solid d = BooleanOperationsUtils.ExecuteBooleanOperation(panel, v, BooleanOperationsType.Difference);
                        if (d != null) panel = d;
                    }
                    catch (Exception) { /* intersección degenerada: se conserva el tablero */ }
                    if (panel.Volume < 1e-9) break;
                }

                double area = panel.Volume / o.Espesor;
                if (area < Un.Cm(1) * Un.Cm(1) * 25) continue; // menos de 25 cm²: es contacto total
                result.Add(new Tablero { Cara = cara, Solido = panel, Area = area });
            }

            if (curvas > 0) avisos.Add($"{host.Name} [Id {host.Id}]: {curvas} cara(s) curva(s) no soportada(s) se omitieron.");
            return result;
        }

        /// <summary>Caras que llevan encofrado según las opciones.</summary>
        private static IEnumerable<(Face face, Cara cara)> Caras(Element host, Opciones o, Action noSoportada)
        {
            List<Solid> solids = Geo.Solids(host).ToList();
            if (solids.Count == 0) yield break;

            double minZ = double.MaxValue;
            foreach (Solid s in solids)
                foreach (Edge e in s.Edges)
                    foreach (XYZ p in e.Tessellate())
                        minZ = Math.Min(minZ, p.Z);

            foreach (Solid s in solids)
                foreach (Face f in s.Faces)
                {
                    if (f is PlanarFace pf)
                    {
                        XYZ n = pf.FaceNormal;
                        if (Math.Abs(n.Z) < Vertical)
                        {
                            if (o.Lateral) yield return (f, Cara.Lateral);
                        }
                        else if (n.Z < 0 && o.Fondo)
                        {
                            if (o.OmitirBase && MaxZ(f) <= minZ + Un.Cm(1)) continue; // apoyada en el suelo o cimiento
                            yield return (f, Cara.Fondo);
                        }
                    }
                    else if (f is CylindricalFace cf && Math.Abs(cf.Axis.Z) > 0.99)
                    {
                        if (o.Lateral) yield return (f, Cara.Lateral);
                    }
                    else
                    {
                        BoundingBoxUV bb = f.GetBoundingBox();
                        XYZ n = f.ComputeNormal((bb.Min + bb.Max) / 2);
                        if (n.Z > Vertical) continue;                 // superior: no se encofra
                        if (n.Z < -Vertical && !o.Fondo) continue;
                        if (Math.Abs(n.Z) <= Vertical && !o.Lateral) continue;
                        noSoportada?.Invoke();
                    }
                }
        }

        private static double MaxZ(Face f)
        {
            double z = double.MinValue;
            foreach (EdgeArray loop in f.EdgeLoops)
                foreach (Edge e in loop)
                    foreach (XYZ p in e.Tessellate())
                        z = Math.Max(z, p.Z);
            return z;
        }

        /// <summary>Extruye una cara plana hacia afuera del elemento.</summary>
        private static Solid ExtruirPlana(PlanarFace f, double t, SolidOptions so, List<Solid> propios)
        {
            XYZ n = f.FaceNormal;
            Solid s = Extruir(f.GetEdgesAsCurveLoops(), n, t, so);
            // Comprobación: si el tablero quedó dentro del propio elemento, la normal estaba invertida.
            if (s != null && DentroDe(s, propios))
                s = Extruir(f.GetEdgesAsCurveLoops(), -n, t, so);
            return s;
        }

        private static Solid Extruir(IList<CurveLoop> loops, XYZ dir, double t, SolidOptions so)
        {
            try { return GeometryCreationUtilities.CreateExtrusionGeometry(loops, dir, t, so); }
            catch (Exception)
            {
                foreach (CurveLoop l in loops) l.Flip();
                return GeometryCreationUtilities.CreateExtrusionGeometry(loops, dir, t, so);
            }
        }

        private static bool DentroDe(Solid panel, List<Solid> propios)
        {
            foreach (Solid h in propios)
            {
                try
                {
                    Solid i = BooleanOperationsUtils.ExecuteBooleanOperation(panel, h, BooleanOperationsType.Intersect);
                    if (i != null && i.Volume > panel.Volume * 0.5) return true;
                }
                catch (Exception) { }
            }
            return false;
        }

        /// <summary>Anillo (o sector) alrededor de una cara cilíndrica vertical: columnas circulares.</summary>
        private static Solid ExtruirCilindro(CylindricalFace f, double t, SolidOptions so)
        {
            var arcs = new List<Arc>();
            foreach (EdgeArray loop in f.EdgeLoops)
                foreach (Edge e in loop)
                    if (e.AsCurve() is Arc a && Math.Abs(a.Normal.Z) > 0.99)
                        arcs.Add(a);
            if (arcs.Count < 2) return null;

            double z0 = arcs.Min(a => a.Center.Z), z1 = arcs.Max(a => a.Center.Z);
            if (z1 - z0 < Un.Cm(1)) return null;
            Arc baseArc = arcs.First(a => Math.Abs(a.Center.Z - z0) < 1e-6);

            var loops = new List<CurveLoop>();
            if (!baseArc.IsBound || baseArc.Length >= 2 * Math.PI * baseArc.Radius - 1e-6)
            {
                // Círculo completo: dos anillos (exterior e interior) partidos en semicírculos.
                loops.Add(Circulo(baseArc.Center, baseArc.Radius + t));
                CurveLoop inner = Circulo(baseArc.Center, baseArc.Radius);
                inner.Flip();
                loops.Add(inner);
            }
            else
            {
                Curve off = baseArc.CreateOffset(t, XYZ.BasisZ);
                if (!(off is Arc oa) || oa.Radius < baseArc.Radius)
                    off = baseArc.CreateOffset(-t, XYZ.BasisZ);
                var loop = new CurveLoop();
                loop.Append(baseArc);
                loop.Append(Line.CreateBound(baseArc.GetEndPoint(1), off.GetEndPoint(1)));
                loop.Append(off.CreateReversed());
                loop.Append(Line.CreateBound(off.GetEndPoint(0), baseArc.GetEndPoint(0)));
                loops.Add(loop);
            }
            return Extruir(loops, XYZ.BasisZ, z1 - z0, so);
        }

        private static CurveLoop Circulo(XYZ c, double r)
        {
            var loop = new CurveLoop();
            loop.Append(Arc.Create(c, r, 0, Math.PI, XYZ.BasisX, XYZ.BasisY));
            loop.Append(Arc.Create(c, r, Math.PI, 2 * Math.PI, XYZ.BasisX, XYZ.BasisY));
            return loop;
        }
    }

    /// <summary>Sólidos de los elementos vecinos (con caché) para descontar las zonas de contacto.</summary>
    internal sealed class Vecinos
    {
        private static readonly BuiltInCategory[] Categorias =
        {
            BuiltInCategory.OST_StructuralColumns, BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralFoundation, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Stairs, BuiltInCategory.OST_Roofs, BuiltInCategory.OST_Columns
        };

        private readonly Document _doc;
        private readonly Dictionary<ElementId, List<Solid>> _cache = new Dictionary<ElementId, List<Solid>>();

        public Vecinos(Document doc) { _doc = doc; }

        public List<Solid> Alrededor(Element host, double pad)
        {
            var result = new List<Solid>();
            BoundingBoxXYZ bb = host.get_BoundingBox(null);
            if (bb == null) return result;
            XYZ d = new XYZ(1, 1, 1) * (pad + Un.Cm(2));
            var outline = new Outline(bb.Min - d, bb.Max + d);

            var ids = new FilteredElementCollector(_doc)
                .WherePasses(new ElementMulticategoryFilter(Categorias))
                .WhereElementIsNotElementType()
                .WherePasses(new BoundingBoxIntersectsFilter(outline))
                .Excluding(new List<ElementId> { host.Id })
                .ToElementIds();

            foreach (ElementId id in ids)
            {
                if (!_cache.TryGetValue(id, out List<Solid> solids))
                {
                    Element e = _doc.GetElement(id);
                    solids = e == null ? new List<Solid>() : Geo.Solids(e).ToList();
                    _cache[id] = solids;
                }
                result.AddRange(solids);
            }
            return result;
        }
    }
}
