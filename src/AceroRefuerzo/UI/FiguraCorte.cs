using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using AceroRefuerzo.Core;
using static AceroRefuerzo.UI.Dibujo;

namespace AceroRefuerzo.UI
{
    /// <summary>Figura del optimizador de corte: resumen por diámetro y patrones de corte a escala.</summary>
    internal static class FiguraCorte
    {
        private static readonly Color[] Piezas =
        {
            Color.FromArgb(200, 35, 35), Color.FromArgb(225, 120, 20), Color.FromArgb(25, 105, 200),
            Color.FromArgb(40, 150, 90), Color.FromArgb(130, 70, 180), Color.FromArgb(0, 140, 150)
        };
        private static readonly Color Sobra = Color.FromArgb(215, 218, 222);
        private static readonly Color Retazo = Color.FromArgb(170, 200, 120);

        public static void Dibujar(Graphics g, RectangleF r, ResultadoCorte res, string titulo)
        {
            SplitV(r, 0.36f, out RectangleF a, out RectangleF b);
            Caption(g, a, "RESUMEN POR DIÁMETRO" + (string.IsNullOrEmpty(titulo) ? "" : "  ·  " + titulo));
            Caption(g, b, "PATRONES DE CORTE (barra comercial a escala)");

            if (res == null || res.Diametros.Count == 0 || res.Diametros.All(d => d.BarrasComerciales == 0))
            {
                Error(g, new RectangleF(a.X + 14, a.Y + 44, a.Width - 28, 60),
                    res == null ? "No hay barras de refuerzo en el alcance elegido." : string.Join("\n", res.Avisos.Concat(res.Diametros.SelectMany(d => d.Avisos)).Take(4)));
                return;
            }

            // --- Resumen: una barra horizontal por diámetro (útil + desperdicio), a escala del mayor.
            RectangleF ra = Inset(a, 14, 42, 14, 8);
            var dias = res.Diametros.Where(d => d.BarrasComerciales > 0).ToList();
            float rowH = Math.Max(14f, Math.Min(26f, (ra.Height - 40) / Math.Max(1, dias.Count)));
            double maxL = dias.Max(d => d.LargoComercialCm);
            float x0 = ra.Left + 70, wMax = Math.Max(40, ra.Width - 70 - 250);
            float y = ra.Top;
            using (var ink = new SolidBrush(Ink))
            using (var util = new SolidBrush(Piezas[0]))
            using (var des = new SolidBrush(Theme.Accent))
            using (var edge = new Pen(ConcreteEdge, 0.8f))
            {
                foreach (CorteDiametro d in dias)
                {
                    if (y + rowH > ra.Bottom - 36) break;
                    g.DrawString("Ø " + d.Bar, Normal, ink, ra.Left, y + rowH / 2 - 8);
                    float w = (float)(wMax * d.LargoComercialCm / maxL);
                    float wu = (float)(w * (1 - d.DesperdicioCm / Math.Max(1, d.LargoComercialCm)));
                    var rc = new RectangleF(x0, y + 3, w, rowH - 6);
                    g.FillRectangle(util, x0, rc.Y, wu, rc.Height);
                    g.FillRectangle(des, x0 + wu, rc.Y, w - wu, rc.Height);
                    g.DrawRectangle(edge, rc.X, rc.Y, rc.Width, rc.Height);
                    string txt = $"{d.BarrasComerciales} barras · {d.LargoComercialCm / 100:0} m · {d.KgComercial:0} kg · desperdicio {d.Porcentaje:0.0} %";
                    g.DrawString(txt, Small, ink, x0 + w + 8, y + rowH / 2 - 7);
                    y += rowH;
                }
            }
            Lines(g, new RectangleF(ra.Left, ra.Bottom - 34, ra.Width, 36),
                ($"Total: {res.BarrasComerciales} barras comerciales · {res.LargoComercialCm / 100:0} m · {res.KgComercial:0} kg", Piezas[0]),
                ($"Desperdicio: {res.Porcentaje:0.0} % = {res.DesperdicioCm / 100:0.0} m ({res.KgDesperdicio:0} kg) · retazos aprovechables {res.RetazosCm / 100:0.0} m", Theme.Accent));

            // --- Patrones: cada barra comercial dibujada a escala con sus piezas.
            RectangleF rb = Inset(b, 14, 42, 14, 10);
            double smax = res.Opciones.ComercialesCm.Max();
            float barH = 20f, gap = 7f, labelW = 64f, countW = 60f;
            float wBar = Math.Max(60, rb.Width - labelW - countW);
            int filas = (int)Math.Floor(rb.Height / (barH + gap));
            if (filas <= 0) return;

            // Reparto de filas: cada diámetro muestra al menos un patrón, y el resto se reparte por orden de frecuencia.
            var mostrar = new List<(CorteDiametro d, PatronCorte p)>();
            foreach (CorteDiametro d in dias) if (d.Patrones.Count > 0) mostrar.Add((d, d.Patrones[0]));
            var resto = dias.SelectMany(d => d.Patrones.Skip(1).Select(p => (d, p))).OrderByDescending(t => t.p.Veces).ToList();
            foreach (var t in resto) { if (mostrar.Count >= filas) break; mostrar.Add(t); }
            mostrar = mostrar.OrderBy(t => dias.IndexOf(t.d)).ThenByDescending(t => t.p.Veces).Take(filas).ToList();
            int ocultos = dias.Sum(d => d.Patrones.Count) - mostrar.Count;

            y = rb.Top;
            using (var ink = new SolidBrush(Ink))
            using (var white = new SolidBrush(Color.White))
            using (var edge = new Pen(ConcreteEdge, 1f))
            using (var sobra = new HatchBrush(HatchStyle.WideUpwardDiagonal, Color.FromArgb(160, 164, 170), Sobra))
            using (var retazo = new HatchBrush(HatchStyle.WideUpwardDiagonal, Color.FromArgb(110, 150, 70), Retazo))
            {
                CorteDiametro prev = null;
                foreach (var (d, p) in mostrar)
                {
                    if (d != prev)
                    {
                        g.DrawString("Ø " + d.Bar, Bold, ink, rb.Left, y + 2);
                        prev = d;
                    }
                    float w = (float)(wBar * p.ComercialCm / smax);
                    var rc = new RectangleF(rb.Left + labelW, y, w, barH);
                    float x = rc.X;
                    int i = 0;
                    foreach (double L in p.Piezas)
                    {
                        float wp = (float)(w * L / p.ComercialCm);
                        using (var br = new SolidBrush(Piezas[i % Piezas.Length]))
                            g.FillRectangle(br, x, rc.Y, wp, barH);
                        string s = (L / 100).ToString("0.00");
                        SizeF sz = g.MeasureString(s, Small);
                        if (sz.Width + 4 < wp)
                            g.DrawString(s, Small, white, x + wp / 2 - sz.Width / 2, rc.Y + barH / 2 - sz.Height / 2);
                        x += wp;
                        i++;
                    }
                    if (x < rc.Right - 0.5f)
                        g.FillRectangle(p.SobraCm >= res.Opciones.RetazoMinCm && res.Opciones.RetazoMinCm > 0 ? retazo : sobra, x, rc.Y, rc.Right - x, barH);
                    g.DrawRectangle(edge, rc.X, rc.Y, rc.Width, rc.Height);
                    g.DrawString($"× {p.Veces}", Bold, ink, rc.Right + 6, y + 2);
                    string info = $"{p.ComercialCm / 100:0.##} m" + (p.SobraCm > 0.5 ? $" · sobra {p.SobraCm / 100:0.00}" : "");
                    SizeF si = g.MeasureString(info, Small);
                    if (rc.Right + 6 + 40 + si.Width < rb.Right)
                        g.DrawString(info, Small, ink, rc.Right + 44, y + 3);
                    y += barH + gap;
                }
                if (ocultos > 0)
                    using (var dim = new SolidBrush(Dim))
                        g.DrawString($"… y {ocultos} patrón(es) más en la planilla.", Small, dim, rb.Left, Math.Min(y, rb.Bottom - 14));
            }
        }
    }
}
