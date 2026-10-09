using System.Collections.Generic;
using AceroRefuerzo.Core;
using AceroRefuerzo.UI;
using Autodesk.Revit.DB;

namespace AceroRefuerzo.Modules
{
    /// <summary>Un tipo de elemento que el programa sabe armar.</summary>
    internal interface IRebarModule
    {
        ElementKind Kind { get; }
        string Title { get; }
        string Prompt { get; }
        bool Accepts(Element e);
        ParamForm CreateForm(Document doc, IList<Element> hosts, IList<BarItem> bars);
        void Apply(Document doc, Element host, FormValues v, View view, RunReport report);
    }

    /// <summary>Módulo que arma varios elementos juntos (por ejemplo, columnas de varios pisos de un mismo eje).</summary>
    internal interface IGroupModule
    {
        List<List<Element>> Agrupar(IList<Element> hosts, FormValues v);
        void ApplyGroup(Document doc, List<Element> grupo, FormValues v, View view, RunReport report);
    }

    internal static class ModuleRegistry
    {
        public static IRebarModule Get(ElementKind kind)
        {
            switch (kind)
            {
                case ElementKind.Viga: return new VigaModule();
                case ElementKind.Columna: return new ColumnaModule();
                case ElementKind.Zapata: return new ZapataModule();
                case ElementKind.Losa: return new LosaModule();
                default: return new MuroModule();
            }
        }
    }
}
