using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using AceroRefuerzo.Core;

namespace AceroRefuerzo.UI
{
    /// <summary>Íconos ("figuritas") dibujados por código para la cinta de Revit y el menú principal.</summary>
    internal static partial class Icons
    {
        public enum Extra { Integral, Diametros, Ayuda }

        private static readonly Color Concrete = Color.FromArgb(200, 204, 210);
        private static readonly Color Edge = Color.FromArgb(90, 96, 105);
        private static readonly Color Red = Color.FromArgb(205, 35, 35);
        private static readonly Color Blue = Color.FromArgb(25, 105, 200);

        public static Bitmap Large(ElementKind kind, int size) => Bitmap(size, g => Draw(kind, g, size));

        private static Bitmap Bitmap(int size, Action<Graphics> draw)
        {
            var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);
                draw(g);
            }
            return bmp;
        }

        public static void Draw(ElementKind kind, Graphics g, float s)
        {
            float k = s / 32f; // todo se dibuja en una grilla de 32 x 32
            float w = Math.Max(1f, 1.2f * k), wb = Math.Max(1f, 1.6f * k);
            using (var cb = new SolidBrush(Concrete))
            using (var ep = new Pen(Edge, Math.Max(1f, k)))
            using (var rp = new Pen(Red, wb))
            using (var bp = new Pen(Blue, w))
            using (var rb = new SolidBrush(Red))
            {
                switch (kind)
                {
                    case ElementKind.Viga:
                        Box(g, cb, ep, 1 * k, 9 * k, 30 * k, 14 * k);
                        for (int i = 0; i < 8; i++) g.DrawLine(bp, (4 + i * 3.5f) * k, 11 * k, (4 + i * 3.5f) * k, 21 * k);
                        g.DrawLines(rp, new[] { P(3, 16, k), P(3, 12, k), P(29, 12, k), P(29, 16, k) });
                        g.DrawLines(rp, new[] { P(3, 16, k), P(3, 20, k), P(29, 20, k), P(29, 16, k) });
                        break;

                    case ElementKind.Columna:
                        Box(g, cb, ep, 10 * k, 1 * k, 12 * k, 30 * k);
                        for (int i = 0; i < 8; i++) g.DrawLine(bp, 12 * k, (3.5f + i * 3.6f) * k, 20 * k, (3.5f + i * 3.6f) * k);
                        g.DrawLine(rp, 13 * k, 2 * k, 13 * k, 30 * k);
                        g.DrawLine(rp, 19 * k, 2 * k, 19 * k, 30 * k);
                        break;

                    case ElementKind.Zapata:
                        Box(g, cb, ep, 12 * k, 1 * k, 8 * k, 16 * k);
                        Box(g, cb, ep, 1 * k, 17 * k, 30 * k, 13 * k);
                        g.DrawLines(rp, new[] { P(3.5f, 22, k), P(3.5f, 27.5f, k), P(28.5f, 27.5f, k), P(28.5f, 22, k) });
                        for (int i = 0; i < 7; i++) g.FillEllipse(rb, (5f + i * 3.6f) * k, 24.2f * k, 2.2f * k, 2.2f * k);
                        g.DrawLine(rp, 14 * k, 2 * k, 14 * k, 27.5f * k);
                        g.DrawLine(rp, 18 * k, 2 * k, 18 * k, 27.5f * k);
                        break;

                    case ElementKind.Losa:
                        var slab = new[] { P(1, 22, k), P(10, 9, k), P(31, 9, k), P(22, 22, k) };
                        g.FillPolygon(cb, slab);
                        g.DrawPolygon(ep, slab);
                        g.FillPolygon(new SolidBrush(Color.FromArgb(170, 175, 182)), new[] { P(1, 22, k), P(22, 22, k), P(22, 25, k), P(1, 25, k) });
                        g.DrawPolygon(ep, new[] { P(1, 22, k), P(22, 22, k), P(22, 25, k), P(1, 25, k) });
                        for (int i = 1; i < 5; i++)
                        {
                            float t = i / 5f;
                            g.DrawLine(rp, Lerp(slab[0], slab[1], t), Lerp(slab[3], slab[2], t));
                            g.DrawLine(bp, Lerp(slab[0], slab[3], t), Lerp(slab[1], slab[2], t));
                        }
                        break;

                    case ElementKind.Muro:
                        Box(g, cb, ep, 5 * k, 1 * k, 22 * k, 30 * k);
                        for (int i = 0; i < 6; i++) g.DrawLine(rp, (8 + i * 3.2f) * k, 2 * k, (8 + i * 3.2f) * k, 30 * k);
                        for (int i = 0; i < 7; i++) g.DrawLine(bp, 6.5f * k, (4.5f + i * 3.8f) * k, 25.5f * k, (4.5f + i * 3.8f) * k);
                        break;
                }
            }
        }

        public static void Draw(Extra kind, Graphics g, float s)
        {
            float k = s / 32f;
            switch (kind)
            {
                case Extra.Integral:
                    using (var cb = new SolidBrush(Color.FromArgb(34, 47, 68)))
                        g.FillEllipse(cb, 1 * k, 1 * k, 30 * k, 30 * k);
                    using (var rp = new Pen(Color.FromArgb(235, 70, 60), 3.2f * k) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                    {
                        g.DrawLine(rp, 8 * k, 24 * k, 24 * k, 8 * k);
                        g.DrawLine(rp, 8 * k, 16 * k, 16 * k, 8 * k);
                        g.DrawLine(rp, 16 * k, 24 * k, 24 * k, 16 * k);
                    }
                    using (var wp = new Pen(Color.White, 1f * k))
                        for (int i = 0; i < 5; i++)
                        {
                            float t = 10 + i * 3.2f;
                            g.DrawLine(wp, (t - 1.5f) * k, (32 - t - 1.5f) * k, (t + 1.5f) * k, (32 - t + 1.5f) * k);
                        }
                    break;

                case Extra.Diametros:
                    using (var rb = new SolidBrush(Red))
                    using (var ep = new Pen(Edge, Math.Max(1f, k)))
                    {
                        float x = 2;
                        foreach (float d in new[] { 4f, 6f, 8f, 11f })
                        {
                            g.FillEllipse(rb, x * k, (24 - d) * k, d * k, d * k);
                            g.DrawEllipse(ep, x * k, (24 - d) * k, d * k, d * k);
                            x += d + 1.2f;
                        }
                        g.DrawLine(ep, 1 * k, 27 * k, 31 * k, 27 * k);
                    }
                    break;

                case Extra.Ayuda:
                    using (var cb = new SolidBrush(Blue))
                        g.FillEllipse(cb, 2 * k, 2 * k, 28 * k, 28 * k);
                    using (var f = new Font("Segoe UI", 17 * k, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (var wb = new SolidBrush(Color.White))
                    {
                        var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        g.DrawString("?", f, wb, new RectangleF(0, 0, s, s), sf);
                    }
                    break;
            }
        }

        private static void Box(Graphics g, Brush b, Pen p, float x, float y, float w, float h)
        {
            g.FillRectangle(b, x, y, w, h);
            g.DrawRectangle(p, x, y, w, h);
        }

        private static PointF P(float x, float y, float k) => new PointF(x * k, y * k);
        private static PointF Lerp(PointF a, PointF b, float t) => new PointF(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t);
    }
}
