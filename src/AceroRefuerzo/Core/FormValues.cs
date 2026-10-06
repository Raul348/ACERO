using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace AceroRefuerzo.Core
{
    /// <summary>Diámetro disponible en el proyecto (para las listas desplegables).</summary>
    public sealed class BarItem
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public double Mm { get; set; }
        public override string ToString() => $"{Name}  (Ø {Mm:0.#} mm)";
    }

    /// <summary>Valores capturados en el formulario (independientes de la ventana).</summary>
    public sealed class FormValues
    {
        private readonly Dictionary<string, object> _v;

        public FormValues(Dictionary<string, object> values) { _v = values; }

        public bool Has(string key) => _v.ContainsKey(key);

        /// <summary>Número tal como se escribió (normalmente en cm).</summary>
        public double D(string key) => Convert.ToDouble(_v[key], CultureInfo.InvariantCulture);

        /// <summary>Número escrito en centímetros convertido a pies.</summary>
        public double Cm(string key) => Un.Cm(D(key));

        public int I(string key) => (int)Math.Round(D(key));
        public bool B(string key) => _v.TryGetValue(key, out object o) && o is bool b && b;
        public string S(string key) => _v.TryGetValue(key, out object o) ? o as string ?? "" : "";
        public BarItem BarItem(string key) => _v.TryGetValue(key, out object o) ? o as BarItem : null;

        public RebarBarType Bar(Document doc, string key)
        {
            BarItem it = BarItem(key) ?? throw new InvalidOperationException("Seleccione un diámetro de barra.");
            return doc.GetElement(it.Id) as RebarBarType
                   ?? throw new InvalidOperationException("El tipo de barra seleccionado ya no existe.");
        }

        /// <summary>Diámetro en milímetros (para dibujar la figura).</summary>
        public double BarMm(string key, double fallback = 12) => BarItem(key)?.Mm ?? fallback;
        public string BarName(string key) => BarItem(key)?.Name ?? "?";
    }
}
