using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Comun;

namespace Encofrado.Core
{
    /// <summary>
    /// Guarda el encofrado en el modelo: un DirectShape (Modelo genérico) por elemento y tipo de cara,
    /// con parámetros compartidos para el metrado (ENC_Elemento, ENC_Cara, ENC_Area, ENC_Nivel, ENC_IdElemento).
    /// </summary>
    internal static class Almacen
    {
        public const string AppId = "ENCOFRADO-Revit";

        public const string PElemento = "ENC_Elemento";
        public const string PCara = "ENC_Cara";
        public const string PArea = "ENC_Area";
        public const string PNivel = "ENC_Nivel";
        public const string PId = "ENC_IdElemento";

        // GUID fijos: el mismo parámetro en todos los proyectos y computadoras.
        private static readonly (string name, Guid guid, bool area)[] Parametros =
        {
            (PElemento, new Guid("3b1f2a10-6c1e-4a0e-9a51-0c1d2e3f4a01"), false),
            (PCara, new Guid("3b1f2a10-6c1e-4a0e-9a51-0c1d2e3f4a02"), false),
            (PArea, new Guid("3b1f2a10-6c1e-4a0e-9a51-0c1d2e3f4a03"), true),
            (PNivel, new Guid("3b1f2a10-6c1e-4a0e-9a51-0c1d2e3f4a04"), false),
            (PId, new Guid("3b1f2a10-6c1e-4a0e-9a51-0c1d2e3f4a05"), false),
        };

        /// <summary>Crea y vincula los parámetros compartidos a Modelos genéricos (solo la primera vez).</summary>
        public static void AsegurarParametros(Document doc, Application app)
        {
            var existentes = new HashSet<string>();
            DefinitionBindingMapIterator it = doc.ParameterBindings.ForwardIterator();
            while (it.MoveNext()) existentes.Add(it.Key.Name);
            if (Parametros.All(p => existentes.Contains(p.name))) return;

            string archivo = Path.Combine(Path.GetTempPath(), "Encofrado_ParametrosCompartidos.txt");
            if (!File.Exists(archivo)) File.WriteAllText(archivo, "");

            string anterior = app.SharedParametersFilename;
            try
            {
                app.SharedParametersFilename = archivo;
                DefinitionFile df = app.OpenSharedParameterFile();
                DefinitionGroup grupo = df.Groups.get_Item("Encofrado") ?? df.Groups.Create("Encofrado");

                CategorySet cats = app.Create.NewCategorySet();
                cats.Insert(Category.GetCategory(doc, BuiltInCategory.OST_GenericModel));
                InstanceBinding binding = app.Create.NewInstanceBinding(cats);

                foreach (var (name, guid, area) in Parametros)
                {
                    if (existentes.Contains(name)) continue;
                    Definition def = grupo.Definitions.get_Item(name);
                    if (def == null)
                    {
                        var opt = new ExternalDefinitionCreationOptions(name, area ? SpecTypeId.Area : SpecTypeId.String.Text)
                        {
                            GUID = guid,
                            Description = "Metrado de encofrado (programa ENCOFRADO)"
                        };
                        def = grupo.Definitions.Create(opt);
                    }
                    doc.ParameterBindings.Insert(def, binding, GroupTypeId.Data);
                }
            }
            finally
            {
                try { if (!string.IsNullOrEmpty(anterior)) app.SharedParametersFilename = anterior; }
                catch (Exception) { }
            }
        }

        /// <summary>Material del tablero (un color por tipo de elemento).</summary>
        public static ElementId Material(Document doc, Pieza p)
        {
            string name = "Encofrado - " + Piezas.Nombre(p);
            Material m = new FilteredElementCollector(doc).OfClass(typeof(Material)).Cast<Material>()
                .FirstOrDefault(x => x.Name == name);
            if (m != null) return m.Id;

            ElementId id = Autodesk.Revit.DB.Material.Create(doc, name);
            m = (Material)doc.GetElement(id);
            System.Drawing.Color c = Piezas.Color(p);
            m.Color = new Color(c.R, c.G, c.B);
            m.SurfaceForegroundPatternColor = new Color(c.R, c.G, c.B);
            m.MaterialClass = "Madera";
            return id;
        }

        public static IEnumerable<DirectShape> Todos(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(DirectShape)).Cast<DirectShape>()
                .Where(d => d.ApplicationId == AppId);

        public static List<ElementId> DeElemento(Document doc, Element host) =>
            Todos(doc).Where(d => d.ApplicationDataId == host.UniqueId).Select(d => d.Id).ToList();

        public static DirectShape Crear(Document doc, Element host, Pieza p, Cara cara, IList<Tablero> tableros)
        {
            var ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
            ds.ApplicationId = AppId;
            ds.ApplicationDataId = host.UniqueId;
            ds.SetShape(tableros.Select(t => (GeometryObject)t.Solido).ToList());
            try { ds.SetName($"Encofrado {Piezas.Nombre(p)} - {Piezas.NombreCara(p, cara)}"); } catch (Exception) { }

            double area = tableros.Sum(t => t.Area);
            Set(ds, PElemento, Piezas.Nombre(p));
            Set(ds, PCara, Piezas.NombreCara(p, cara));
            Set(ds, PNivel, Nivel(doc, host));
            Set(ds, PId, host.Id.ToString());
            Parameter pa = ds.LookupParameter(PArea);
            if (pa != null && !pa.IsReadOnly) pa.Set(area);

            Parameter com = ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (com != null && !com.IsReadOnly) com.Set($"ENCOFRADO {Piezas.Nombre(p)} - {M2(area):0.00} m²");
            return ds;
        }

        private static void Set(Element e, string name, string value)
        {
            Parameter p = e.LookupParameter(name);
            if (p != null && !p.IsReadOnly) p.Set(value);
        }

        public static string Texto(Element e, string name) => e.LookupParameter(name)?.AsString() ?? "";
        public static double Area(Element e) => e.LookupParameter(PArea)?.AsDouble() ?? 0;

        /// <summary>Nombre del nivel del elemento (si tiene).</summary>
        public static string Nivel(Document doc, Element e)
        {
            ElementId id = e.LevelId;
            if (id == null || id == ElementId.InvalidElementId)
            {
                foreach (BuiltInParameter bip in new[]
                {
                    BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM, BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                    BuiltInParameter.WALL_BASE_CONSTRAINT, BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                    BuiltInParameter.LEVEL_PARAM, BuiltInParameter.STAIRS_BASE_LEVEL_PARAM
                })
                {
                    Parameter p = e.get_Parameter(bip);
                    if (p != null && p.StorageType == StorageType.ElementId && p.AsElementId() != ElementId.InvalidElementId)
                    {
                        id = p.AsElementId();
                        break;
                    }
                }
            }
            return (doc.GetElement(id) as Level)?.Name ?? "";
        }

        /// <summary>m² a partir de pies².</summary>
        public static double M2(double pies2) => pies2 * 0.09290304;
    }
}
