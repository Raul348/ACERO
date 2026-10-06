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
        ParamForm CreateForm(Document doc, Element sample, IList<BarItem> bars);
        void Apply(Document doc, Element host, FormValues v, View view, RunReport report);
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
