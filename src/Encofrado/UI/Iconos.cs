using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using Encofrado.Core;

namespace Encofrado.UI
{
    /// <summary>Íconos ("figuritas") del encofrado dibujados por código: concreto gris y tableros de color.</summary>
    internal static partial class Iconos
    {
        public enum Extra { Todo, Metrado, Eliminar, Visibilidad, Ayuda }

        private static readonly Color Concreto = Color.FromArgb(196, 200, 206);
        private static readonly Color Borde = Color.FromArgb(90, 96, 105);
        private static readonly Color Madera = Color.FromArgb(120, 78, 36);

        public static Bitmap Large(Pieza p, int size) => Bitmap(size, g => Draw(p, g, size));

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

        public static void Draw(Pieza p, Graphics g, float s)
        {
            float k = s / 32f;
            Color c = Piezas.Color(p);
            using (var cb = new SolidBrush(Concreto))
            using (var wb = new SolidBrush(c))
            using (var ep = new Pen(Borde, Math.Max(1f, k)))
            using (var mp = new Pen(Madera, Math.Max(1f, 1.4f * k)))
            {
                switch (p)
                {
                    case Pieza.Columna:
                        Box(g, cb, ep, 11, 2, 10, 28, k);
                        Box(g, wb, ep, 7, 2, 4, 28, k);
                        Box(g, wb, ep, 21, 2, 4, 28, k);
                        for (int i = 0; i < 4; i++) g.DrawLine(mp, 5 * k, (6 + i * 7) * k, 27 * k, (6 + i * 7) * k);
                        break;

                    case Pieza.Viga:
                        Box(g, cb, ep, 1, 6, 30, 5, k);   // losa
                        Box(g, cb, ep, 10, 11, 12, 9, k); // viga
                        Box(g, wb, ep, 7, 11, 3, 12, k);
                        Box(g, wb, ep, 22, 11, 3, 12, k);
                        Box(g, wb, ep, 7, 20, 18, 3, k);
                        g.DrawLine(new Pen(Madera, 2.4f * k), 16 * k, 23 * k, 16 * k, 31 * k);
                        g.DrawLine(mp, 12 * k, 31 * k, 20 * k, 31 * k);
                        break;

                    case Pieza.Zapata:
                        Box(g, cb, ep, 12, 3, 8, 13, k);
                        Box(g, cb, ep, 4, 16, 24, 10, k);
                        Box(g, wb, ep, 1, 15, 3, 12, k);
                        Box(g, wb, ep, 28, 15, 3, 12, k);
                        g.DrawLine(mp, 1 * k, 30 * k, 31 * k, 30 * k);
                        break;

                    case Pieza.Losa:
                        var top = new[] { P(1, 12, k), P(9, 4, k), P(31, 4, k), P(23, 12, k) };
                        g.FillPolygon(cb, top); g.DrawPolygon(ep, top);
                        Box(g, cb, ep, 1, 12, 22, 4, k);
                        Box(g, wb, ep, 1, 16, 22, 3, k);
                        foreach (float x in new[] { 4f, 12f, 20f })
                            g.DrawLine(new Pen(Madera, 2.2f * k), x * k, 19 * k, x * k, 30 * k);
                        g.DrawLine(mp, 1 * k, 30.5f * k, 25 * k, 30.5f * k);
                        break;

                    case Pieza.Muro:
                        Box(g, wb, ep, 4, 2, 24, 28, k);
                        using (var lp = new Pen(Madera, Math.Max(1f, k)))
                        {
                            g.DrawLine(lp, 16 * k, 2 * k, 16 * k, 30 * k);
                            g.DrawLine(lp, 4 * k, 16 * k, 28 * k, 16 * k);
                        }
                        using (var db = new SolidBrush(Color.FromArgb(60, 60, 65)))
                            foreach (float x in new[] { 10f, 22f })
                                foreach (float y in new[] { 9f, 23f })
                                    g.FillEllipse(db, (x - 1.2f) * k, (y - 1.2f) * k, 2.4f * k, 2.4f * k);
                        break;

                    default:
                        var st = new[] { P(2, 30, k), P(2, 25, k), P(9, 25, k), P(9, 19, k), P(16, 19, k), P(16, 13, k), P(23, 13, k), P(23, 7, k), P(30, 7, k), P(30, 14, k), P(9, 30, k) };
                        g.FillPolygon(cb, st); g.DrawPolygon(ep, st);
                        using (var fp = new Pen(c, 3f * k)) g.DrawLine(fp, 11 * k, 31 * k, 31 * k, 16.5f * k);
                        foreach (var (x, y) in new[] { (2f, 25f), (9f, 19f), (16f, 13f), (23f, 7f) })
                            Box(g, wb, ep, x - 2, y, 2, 6, k);
                        break;
                }
            }
        }

        public static void Draw(Extra e, Graphics g, float s)
        {
            float k = s / 32f;
            switch (e)
            {
                case Extra.Todo:
                    using (var cb = new SolidBrush(Concreto))
                    using (var ep = new Pen(Borde, Math.Max(1f, k)))
                    using (var wp = new Pen(Piezas.Color(Pieza.Columna), 3f * k))
                    using (var bp = new Pen(Piezas.Color(Pieza.Viga), 3f * k))
                    using (var lp = new Pen(Piezas.Color(Pieza.Losa), 2.4f * k))
                    {
                        foreach (float y in new[] { 9f, 19f })
                        {
                            Box(g, cb, ep, 2, y, 28, 3, k);
                            g.DrawLine(lp, 2 * k, (y + 4) * k, 30 * k, (y + 4) * k);
                        }
                        foreach (float x in new[] { 4f, 15f, 26f })
                        {
                            g.DrawLine(wp, x * k, 12 * k, x * k, 19 * k);
                            g.DrawLine(wp, x * k, 22 * k, x * k, 30 * k);
                        }
                        Box(g, new SolidBrush(Piezas.Color(Pieza.Zapata)), ep, 1, 29, 30, 2.5f, k);
                        g.DrawLine(bp, 4 * k, 6 * k, 15 * k, 2 * k);
                        g.DrawLine(bp, 15 * k, 2 * k, 26 * k, 6 * k);
                    }
                    break;

                case Extra.Metrado:
                    using (var pb = new SolidBrush(Color.White))
                    using (var ep = new Pen(Borde, Math.Max(1f, k)))
                    using (var hb = new SolidBrush(Color.FromArgb(34, 47, 68)))
                    using (var lp = new Pen(Color.FromArgb(160, 165, 172), Math.Max(1f, k)))
                    {
                        Box(g, pb, ep, 4, 2, 24, 28, k);
                        g.FillRectangle(hb, 4 * k, 2 * k, 24 * k, 6 * k);
                        for (int i = 0; i < 4; i++) g.DrawLine(lp, 7 * k, (12 + i * 4) * k, 25 * k, (12 + i * 4) * k);
                        using (var f = new Font("Segoe UI", 7.5f * k, FontStyle.Bold, GraphicsUnit.Pixel))
                        using (var tb = new SolidBrush(Color.FromArgb(196, 38, 38)))
                            g.DrawString("m²", f, tb, 12 * k, 22 * k);
                    }
                    break;

                case Extra.Eliminar:
                    using (var wb = new SolidBrush(Piezas.Color(Pieza.Columna)))
                    using (var ep = new Pen(Borde, Math.Max(1f, k)))
                    using (var xp = new Pen(Color.FromArgb(196, 38, 38), 3.4f * k) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                    {
                        Box(g, wb, ep, 6, 3, 20, 26, k);
                        g.DrawLine(xp, 9 * k, 8 * k, 23 * k, 24 * k);
                        g.DrawLine(xp, 23 * k, 8 * k, 9 * k, 24 * k);
                    }
                    break;

                case Extra.Visibilidad:
                    using (var ep = new Pen(Color.FromArgb(34, 47, 68), 2.2f * k))
                    using (var ib = new SolidBrush(Piezas.Color(Pieza.Columna)))
                    using (var pb = new SolidBrush(Color.FromArgb(34, 47, 68)))
                    {
                        var path = new GraphicsPath();
                        path.AddBezier(2 * k, 16 * k, 9 * k, 5 * k, 23 * k, 5 * k, 30 * k, 16 * k);
                        path.AddBezier(30 * k, 16 * k, 23 * k, 27 * k, 9 * k, 27 * k, 2 * k, 16 * k);
                        g.DrawPath(ep, path);
                        g.FillEllipse(ib, 10 * k, 10 * k, 12 * k, 12 * k);
                        g.FillEllipse(pb, 13.5f * k, 13.5f * k, 5 * k, 5 * k);
                    }
                    break;

                default:
                    using (var cb = new SolidBrush(Color.FromArgb(25, 105, 200)))
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

        private static void Box(Graphics g, Brush b, Pen p, float x, float y, float w, float h, float k)
        {
            g.FillRectangle(b, x * k, y * k, w * k, h * k);
            g.DrawRectangle(p, x * k, y * k, w * k, h * k);
        }

        private static PointF P(float x, float y, float k) => new PointF(x * k, y * k);
    }
}
