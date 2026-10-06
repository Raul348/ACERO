using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace Comun
{
    /// <summary>Primitivas de dibujo compartidas por las figuras de los programas (estilo común).</summary>
    public static class Dibujo
    {
        public static readonly Color ConcreteFill = Color.FromArgb(226, 228, 231);
        public static readonly Color ConcreteEdge = Color.FromArgb(105, 110, 118);
        public static readonly Color Dim = Color.FromArgb(120, 120, 120);
        public static readonly Color Ink = Color.FromArgb(45, 45, 50);
        public static readonly Font Small = new Font("Segoe UI", 8f);
        public static readonly Font Normal = new Font("Segoe UI", 9f);
        public static readonly Font Bold = new Font("Segoe UI Semibold", 9.5f);


        public sealed class Map
        {
            public float S;
            private float _ox, _oy;
            private double _x0, _y0;

            public static Map Fit(RectangleF r, double x0, double y0, double w, double h)
            {
                w = Math.Max(w, 1e-6); h = Math.Max(h, 1e-6);
                float s = (float)Math.Min(r.Width / w, r.Height / h);
                float pw = (float)(w * s), ph = (float)(h * s);
                return new Map
                {
                    S = s, _x0 = x0, _y0 = y0,
                    _ox = r.Left + (r.Width - pw) / 2,
                    _oy = r.Top + (r.Height - ph) / 2 + ph
                };
            }

            public PointF P(double x, double y) => new PointF(_ox + (float)((x - _x0) * S), _oy - (float)((y - _y0) * S));
            public float L(double d) => (float)(d * S);
        }

        public static RectangleF Inset(RectangleF r, float l, float t, float rr, float b) =>
            new RectangleF(r.Left + l, r.Top + t, Math.Max(10, r.Width - l - rr), Math.Max(10, r.Height - t - b));

        public static void SplitH(RectangleF r, float k, out RectangleF a, out RectangleF b)
        {
            a = new RectangleF(r.Left, r.Top, r.Width * k, r.Height);
            b = new RectangleF(r.Left + r.Width * k, r.Top, r.Width * (1 - k), r.Height);
        }

        public static void SplitV(RectangleF r, float k, out RectangleF a, out RectangleF b)
        {
            a = new RectangleF(r.Left, r.Top, r.Width, r.Height * k);
            b = new RectangleF(r.Left, r.Top + r.Height * k, r.Width, r.Height * (1 - k));
        }

        public static void Caption(Graphics g, RectangleF r, string text)
        {
            using (var b = new SolidBrush(Theme.Header))
                g.DrawString(text, Bold, b, r.Left + 12, r.Top + 10);
            using (var p = new Pen(Theme.Accent, 2f))
                g.DrawLine(p, r.Left + 13, r.Top + 30, r.Left + 46, r.Top + 30);
        }

        public static void Lines(Graphics g, RectangleF r, params (string text, Color color)[] lines)
        {
            float y = r.Top;
            foreach (var (text, color) in lines)
            {
                using (var b = new SolidBrush(color)) g.FillEllipse(b, r.Left, y + 5, 7, 7);
                using (var b = new SolidBrush(Ink)) g.DrawString(text, Normal, b, r.Left + 12, y);
                y += 18;
            }
        }

        public static void Error(Graphics g, RectangleF r, string text)
        {
            using (var b = new SolidBrush(Theme.Accent)) g.DrawString("⚠ " + text, Normal, b, r);
        }

        public static void Fill(Graphics g, RectangleF r)
        {
            using (var b = new HatchBrush(HatchStyle.Percent10, Color.FromArgb(175, 178, 184), ConcreteFill)) g.FillRectangle(b, r);
            using (var p = new Pen(ConcreteEdge, 1.5f)) g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height);
        }

        public static void ConcreteRect(Graphics g, Map m, double x, double y, double w, double h)
        {
            PointF p0 = m.P(x, y + h);
            Fill(g, new RectangleF(p0.X, p0.Y, m.L(w), m.L(h)));
        }

        public static void Polygon(Graphics g, PointF[] pts)
        {
            using (var b = new HatchBrush(HatchStyle.Percent10, Color.FromArgb(175, 178, 184), ConcreteFill)) g.FillPolygon(b, pts);
            using (var p = new Pen(ConcreteEdge, 1.5f)) g.DrawPolygon(p, pts);
        }



        public static void Dot(Graphics g, PointF c, float r, Color col)
        {
            using (var b = new SolidBrush(col)) g.FillEllipse(b, c.X - r, c.Y - r, 2 * r, 2 * r);
            using (var p = new Pen(Color.FromArgb(120, 0, 0, 0), 0.8f)) g.DrawEllipse(p, c.X - r, c.Y - r, 2 * r, 2 * r);
        }

        public static void DimH(Graphics g, PointF a, PointF b, float off, string text)
        {
            float y = a.Y + off;
            using (var p = new Pen(Dim, 1f))
            {
                g.DrawLine(p, a.X, a.Y + 3, a.X, y + 4);
                g.DrawLine(p, b.X, b.Y + 3, b.X, y + 4);
                g.DrawLine(p, a.X, y, b.X, y);
                g.DrawLine(p, a.X - 3, y + 3, a.X + 3, y - 3);
                g.DrawLine(p, b.X - 3, y + 3, b.X + 3, y - 3);
            }
            SizeF sz = g.MeasureString(text, Small);
            using (var bg = new SolidBrush(Color.White)) g.FillRectangle(bg, (a.X + b.X) / 2 - sz.Width / 2, y - sz.Height / 2, sz.Width, sz.Height);
            using (var br = new SolidBrush(Ink)) g.DrawString(text, Small, br, (a.X + b.X) / 2 - sz.Width / 2, y - sz.Height / 2);
        }

        public static void DimV(Graphics g, PointF a, PointF b, float off, string text)
        {
            float x = a.X + off;
            using (var p = new Pen(Dim, 1f))
            {
                g.DrawLine(p, a.X, a.Y, x, a.Y);
                g.DrawLine(p, b.X, b.Y, x, b.Y);
                g.DrawLine(p, x, a.Y, x, b.Y);
                g.DrawLine(p, x - 3, a.Y + 3, x + 3, a.Y - 3);
                g.DrawLine(p, x - 3, b.Y + 3, x + 3, b.Y - 3);
            }
            GraphicsState st = g.Save();
            g.TranslateTransform(x, (a.Y + b.Y) / 2);
            g.RotateTransform(-90);
            SizeF sz = g.MeasureString(text, Small);
            using (var bg = new SolidBrush(Color.White)) g.FillRectangle(bg, -sz.Width / 2, -sz.Height / 2, sz.Width, sz.Height);
            using (var br = new SolidBrush(Ink)) g.DrawString(text, Small, br, -sz.Width / 2, -sz.Height / 2);
            g.Restore(st);
        }
    }
}
