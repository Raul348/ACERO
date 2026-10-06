using System.Drawing;
using Autodesk.Revit.DB;

namespace Encofrado.Core
{
    /// <summary>Elementos estructurales que se pueden encofrar.</summary>
    public enum Pieza
    {
        Columna,
        Viga,
        Zapata,
        Losa,
        Muro,
        Escalera
    }

    /// <summary>Cara de encofrado (se metra por separado).</summary>
    public enum Cara
    {
        Lateral,
        Fondo
    }

    /// <summary>Datos y reglas de cada tipo de elemento.</summary>
    internal static class Piezas
    {
        public static readonly Pieza[] Todas = { Pieza.Columna, Pieza.Viga, Pieza.Zapata, Pieza.Losa, Pieza.Muro, Pieza.Escalera };

        public static string Nombre(Pieza p)
        {
            switch (p)
            {
                case Pieza.Columna: return "Columna";
                case Pieza.Viga: return "Viga";
                case Pieza.Zapata: return "Zapata";
                case Pieza.Losa: return "Losa";
                case Pieza.Muro: return "Muro";
                default: return "Escalera";
            }
        }

        public static string Plural(Pieza p)
        {
            switch (p)
            {
                case Pieza.Columna: return "Columnas";
                case Pieza.Viga: return "Vigas";
                case Pieza.Zapata: return "Zapatas";
                case Pieza.Losa: return "Losas";
                case Pieza.Muro: return "Muros";
                default: return "Escaleras";
            }
        }

        /// <summary>Nombre de cada cara según el elemento (como se metra en obra).</summary>
        public static string NombreCara(Pieza p, Cara c)
        {
            if (c == Cara.Fondo)
            {
                switch (p)
                {
                    case Pieza.Viga: return "Fondo de viga";
                    case Pieza.Losa: return "Fondo de losa";
                    case Pieza.Muro: return "Fondo de vanos (dinteles)";
                    case Pieza.Escalera: return "Fondo de escalera y descansos";
                    default: return "Fondo";
                }
            }
            switch (p)
            {
                case Pieza.Columna: return "Caras laterales";
                case Pieza.Viga: return "Costados de viga";
                case Pieza.Zapata: return "Costados de zapata";
                case Pieza.Losa: return "Frisos (bordes de losa)";
                case Pieza.Muro: return "Caras del muro, extremos y jambas";
                default: return "Costados y contrapasos";
            }
        }

        /// <summary>Qué caras se encofran por defecto.</summary>
        public static bool LateralPorDefecto(Pieza p) => true;
        public static bool FondoPorDefecto(Pieza p) => p != Pieza.Columna && p != Pieza.Zapata;
        public static bool TieneFondo(Pieza p) => p != Pieza.Columna && p != Pieza.Zapata;

        /// <summary>Las caras inferiores apoyadas (base del muro o de la escalera) no se encofran.</summary>
        public static bool OmitirBase(Pieza p) => p == Pieza.Muro || p == Pieza.Escalera;

        /// <summary>Color del tablero en el modelo (y en la figura).</summary>
        public static System.Drawing.Color Color(Pieza p)
        {
            switch (p)
            {
                case Pieza.Columna: return System.Drawing.Color.FromArgb(232, 140, 40);
                case Pieza.Viga: return System.Drawing.Color.FromArgb(238, 196, 50);
                case Pieza.Zapata: return System.Drawing.Color.FromArgb(160, 105, 60);
                case Pieza.Losa: return System.Drawing.Color.FromArgb(105, 175, 90);
                case Pieza.Muro: return System.Drawing.Color.FromArgb(80, 145, 210);
                default: return System.Drawing.Color.FromArgb(165, 95, 190);
            }
        }

        /// <summary>Categoría y clase del elemento (sin filtros de material).</summary>
        public static bool Es(Pieza p, Element e)
        {
            BuiltInCategory? cat = e.Category?.BuiltInCategory;
            switch (p)
            {
                case Pieza.Columna: return e is FamilyInstance && cat == BuiltInCategory.OST_StructuralColumns;
                case Pieza.Viga: return e is FamilyInstance && cat == BuiltInCategory.OST_StructuralFraming;
                case Pieza.Zapata:
                    return cat == BuiltInCategory.OST_StructuralFoundation && (e is FamilyInstance || e is Floor || e is WallFoundation);
                case Pieza.Losa: return e is Floor && cat == BuiltInCategory.OST_Floors;
                case Pieza.Muro: return e is Wall w && w.WallType?.Kind == WallKind.Basic;
                default: return cat == BuiltInCategory.OST_Stairs && !(e is ElementType);
            }
        }

        /// <summary>Tipo de pieza de un elemento cualquiera (null si no se encofra).</summary>
        public static Pieza? De(Element e)
        {
            foreach (Pieza p in Todas)
                if (Es(p, e)) return p;
            return null;
        }

        /// <summary>Filtros opcionales: solo concreto vaciado en obra / solo muros estructurales.</summary>
        public static string Excluir(Pieza p, Element e, bool soloConcreto, bool soloMurosEstructurales)
        {
            if (soloConcreto && e is FamilyInstance fi)
            {
                string mt = fi.StructuralMaterialType.ToString();
                if (mt == "Steel" || mt == "Wood" || mt == "Aluminum" || mt == "PrecastConcrete")
                    return "no es de concreto vaciado en obra";
            }
            if (p == Pieza.Muro && soloMurosEstructurales)
            {
                Parameter s = e.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT);
                if (s != null && s.AsInteger() == 0) return "no es un muro estructural";
            }
            return null;
        }
    }
}
