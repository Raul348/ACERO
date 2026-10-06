using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using Comun;
using Encofrado.Core;
using static Comun.Dibujo;

namespace Encofrado.UI
{
    /// <summary>
    /// "Figuritas" del encofrado: corte y vista de cada elemento con sus tableros, yugos, puntales,
    /// barrotes y estacas, según las caras elegidas en el formulario.
    /// </summary>
    internal static class Figuras
    {
        private static readonly Color Madera = Color.FromArgb(120, 78, 36);
        private static readonly Color Suelo = Color.FromArgb(150, 140, 125);

        public static void Dibujar(Graphics g, RectangleF r, FormValues v, Pieza p, Medidas m, double latM2, double fonM2, int n)
        {
            bool lat = v.B("lat"), fon = v.Has("fondo") && v.B("fondo");
            Color c = Piezas.Color(p);
            switch (p)
            {
                case Pieza.Columna: Columna(g, r, c, lat, m); break;
                case Pieza.Viga: Viga(g, r, c, lat, fon, m); break;
                case Pieza.Zapata: Zapata(g, r, c, lat, m); break;
                case Pieza.Losa: Losa(g, r, c, lat, fon, m); break;
                case Pieza.Muro: Muro(g, r, c, lat, fon, m); break;
                default: Escalera(g, r, c, lat, fon, m); break;
            }

            var lines = new List<(string, Color)>();
            if (lat) lines.Add(($"{Piezas.NombreCara(p, Cara.Lateral)}: ≈ {latM2:N2} m²", c));
            if (fon) lines.Add(($"{Piezas.NombreCara(p, Cara.Fondo)}: ≈ {fonM2:N2} m²", Color.FromArgb(c.R * 3 / 4, c.G * 3 / 4, c.B * 3 / 4)));
            lines.Add(($"Tablero e = {v.D("esp"):0.#} cm · {n} elemento(s) · áreas del 1.º sin descontar contactos", Ink));
            Lines(g, new RectangleF(r.X + 14, r.Bottom - 22 - 18 * lines.Count, r.Width - 20, 18 * lines.Count + 4), lines.ToArray());
        }

        // ================================================================== COLUMNA

        private static void Columna(Graphics g, RectangleF r, Color c, bool lat, Medidas d)
        {
            SplitH(r, 0.5f, out RectangleF a, out RectangleF e);
            Caption(g, a, "CORTE (planta)");
            Caption(g, e, "ELEVACIÓN");
            double bx = d.A, by = d.B, H = d.C;

            double pad = Math.Max(bx, by) * 0.22;
            Map m = Map.Fit(Inset(a, 40, 50, 30, 100), -pad, -pad, bx + 2 * pad, by + 2 * pad);
            double t = T(m, 7);
            ConcreteRect(g, m, 0, 0, bx, by);
            if (lat)
            {
                Board(g, m, -t, 0, t, by, c, true);
                Board(g, m, bx, 0, t, by, c, true);
                Board(g, m, -t, -t, bx + 2 * t, t, c, false);
                Board(g, m, -t, by, bx + 2 * t, t, c, false);
                double y = T(m, 4);
                using (var pen = new Pen(Madera, 3f))
                {
                    PointF p0 = m.P(-t - y, by + t + y), p1 = m.P(bx + t + y, -t - y);
                    g.DrawRectangle(pen, p0.X, p0.Y, p1.X - p0.X, p1.Y - p0.Y);
                    float k = 9;
                    g.DrawLine(pen, p0.X - k, p0.Y, p0.X, p0.Y); g.DrawLine(pen, p1.X, p1.Y, p1.X + k, p1.Y);
                    g.DrawLine(pen, p0.X, p1.Y, p0.X, p1.Y + k); g.DrawLine(pen, p1.X, p0.Y, p1.X, p0.Y - k);
                }
            }
            DimH(g, m.P(0, -pad * 0.6), m.P(bx, -pad * 0.6), 0, $"{bx:0.#} cm");
            DimV(g, m.P(-pad * 0.6, 0), m.P(-pad * 0.6, by), 0, $"{by:0.#} cm");

            // Elevación
            RectangleF ea = Inset(e, 30, 50, 30, 100);
            float w = Math.Min(ea.Width * 0.3f, 60f);
            var col = new RectangleF(ea.Left + (ea.Width - w) / 2, ea.Top + 10, w, ea.Height - 30);
            Floor(g, ea.Left, ea.Right, col.Bottom);
            Fill(g, col);
            if (lat)
            {
                var face = RectangleF.Inflate(col, 5, 0);
                BoardPx(g, face, c, true);
                using (var pen = new Pen(Madera, 3f))
                {
                    int n = Math.Max(3, (int)(H / 60));
                    for (int i = 1; i <= n; i++)
                    {
                        float yy = col.Bottom - col.Height * i / (n + 1);
                        g.DrawLine(pen, face.Left - 6, yy, face.Right + 6, yy);
                    }
                }
                float spread = Math.Min(col.Height * 0.3f, face.Left - ea.Left - 4);
                Prop(g, new PointF(face.Left - 3, col.Top + col.Height * 0.3f), new PointF(face.Left - spread, col.Bottom));
                Prop(g, new PointF(face.Right + 3, col.Top + col.Height * 0.3f), new PointF(face.Right + spread, col.Bottom));
            }
            DimV(g, new PointF(col.Right + 8, col.Bottom), new PointF(col.Right + 8, col.Top), 22, $"{H / 100:0.00} m");
        }

        // ================================================================== VIGA

        private static void Viga(Graphics g, RectangleF r, Color c, bool lat, bool fon, Medidas d)
        {
            SplitH(r, 0.42f, out RectangleF a, out RectangleF e);
            Caption(g, a, "CORTE");
            Caption(g, e, "ELEVACIÓN");
            double b = d.A, h = d.B, L = d.C;

            double es = Math.Max(12, h * 0.25), hp = h * 1.6;
            Map m = Map.Fit(Inset(a, 30, 50, 30, 100), -b * 1.4, -hp, b * 3.8, hp + h + es);
            double t = T(m, 7);
            Floor(g, m.P(-b * 1.4, 0).X, m.P(b * 2.4, 0).X, m.P(0, -hp).Y);
            ConcreteRect(g, m, -b * 1.4, h, b * 3.8, es);
            ConcreteRect(g, m, 0, 0, b, h);
            if (lat)
            {
                Board(g, m, -t, 0, t, h, c, true);
                Board(g, m, b, 0, t, h, c, true);
            }
            if (fon) Board(g, m, -t, -t, b + 2 * t, t, Shade(c), false);
            Board(g, m, -b * 1.4, h - t, b * 1.4 - t, t, Color.FromArgb(200, 210, 190), false);
            Board(g, m, b + t, h - t, b * 1.4, t, Color.FromArgb(200, 210, 190), false);
            Board(g, m, -t * 3, -t * 2.5, b + 6 * t, t * 1.5, Madera, false);
            Prop(g, m.P(b / 2, -t * 2.5), m.P(b / 2, -hp));
            DimH(g, m.P(0, h + es), m.P(b, h + es), -14, $"b = {b:0.#} cm");
            DimV(g, m.P(b + t, 0), m.P(b + t, h), 18, $"h = {h:0.#} cm");

            RectangleF ea = Inset(e, 30, 60, 30, 100);
            float hb = Math.Min(ea.Height * 0.3f, Math.Max(36f, ea.Width * (float)(h / Math.Max(L, 1)) * 2.5f));
            float floorY = ea.Bottom - 10;
            var beam = new RectangleF(ea.Left + 16, ea.Top + 30, ea.Width - 32, hb);
            Floor(g, ea.Left, ea.Right, floorY);
            using (var cb = new SolidBrush(Color.FromArgb(200, 203, 208)))
            {
                g.FillRectangle(cb, beam.Left - 16, beam.Top - 20, 16, floorY - beam.Top + 20);
                g.FillRectangle(cb, beam.Right, beam.Top - 20, 16, floorY - beam.Top + 20);
            }
            Fill(g, beam);
            if (lat) BoardPx(g, beam, c, false);
            if (fon) BoardPx(g, new RectangleF(beam.Left, beam.Bottom, beam.Width, 6), Shade(c), false);
            int np = Math.Max(3, (int)(L / 90));
            for (int i = 1; i <= np; i++)
            {
                float x = beam.Left + beam.Width * i / (np + 1);
                Prop(g, new PointF(x, beam.Bottom + 6), new PointF(x, floorY));
            }
            DimH(g, new PointF(beam.Left, beam.Top), new PointF(beam.Right, beam.Top), -16, $"L = {L / 100:0.00} m");
        }

        // ================================================================== ZAPATA

        private static void Zapata(Graphics g, RectangleF r, Color c, bool lat, Medidas d)
        {
            SplitV(r, 0.52f, out RectangleF a, out RectangleF e);
            Caption(g, a, "CORTE");
            Caption(g, e, "PLANTA");
            double lx = d.A, ly = d.B, H = d.C;
            double cs = Math.Min(lx, ly) * 0.25;

            Map m = Map.Fit(Inset(a, 50, 46, 50, 16), -lx * 0.3, -H * 0.4, lx * 1.6, H * 2.2);
            double t = T(m, 7);
            using (var sb = new HatchBrush(HatchStyle.WideUpwardDiagonal, Suelo, Color.FromArgb(235, 230, 222)))
            {
                PointF p0 = m.P(-lx * 0.3, 0), p1 = m.P(lx * 1.3, -H * 0.4);
                g.FillRectangle(sb, p0.X, p0.Y, p1.X - p0.X, p1.Y - p0.Y);
            }
            ConcreteRect(g, m, lx / 2 - cs / 2, H, cs, H * 0.8);
            ConcreteRect(g, m, 0, 0, lx, H);
            if (lat)
            {
                Board(g, m, -t, 0, t, H, c, true);
                Board(g, m, lx, 0, t, H, c, true);
                Prop(g, m.P(-t, H * 0.8), m.P(-t - H * 0.7, 0));
                Prop(g, m.P(lx + t, H * 0.8), m.P(lx + t + H * 0.7, 0));
            }
            DimV(g, m.P(lx + t, 0), m.P(lx + t, H), 26, $"{H:0.#} cm");

            Map k = Map.Fit(Inset(e, 50, 40, 50, 80), -lx * 0.1, -ly * 0.1, lx * 1.2, ly * 1.2);
            double tk = T(k, 6);
            ConcreteRect(g, k, 0, 0, lx, ly);
            using (var cb = new SolidBrush(Color.FromArgb(150, 120, 125, 135)))
            {
                PointF p0 = k.P(lx / 2 - cs / 2, ly / 2 + cs / 2);
                g.FillRectangle(cb, p0.X, p0.Y, k.L(cs), k.L(cs));
            }
            if (lat)
            {
                Board(g, k, -tk, -tk, tk, ly + 2 * tk, c, true);
                Board(g, k, lx, -tk, tk, ly + 2 * tk, c, true);
                Board(g, k, 0, -tk, lx, tk, c, false);
                Board(g, k, 0, ly, lx, tk, c, false);
            }
            DimH(g, k.P(0, -tk), k.P(lx, -tk), 14, $"{lx / 100:0.00} m");
            DimV(g, k.P(-tk, 0), k.P(-tk, ly), -14, $"{ly / 100:0.00} m");
        }

        // ================================================================== LOSA

        private static void Losa(Graphics g, RectangleF r, Color c, bool lat, bool fon, Medidas d)
        {
            SplitV(r, 0.5f, out RectangleF a, out RectangleF e);
            Caption(g, a, "CORTE");
            Caption(g, e, "PLANTA");
            double lx = d.A, ly = d.B, es = d.C;

            double w = Math.Min(Math.Max(lx, 100), 400), hp = Math.Max(es * 6, 120);
            Map m = Map.Fit(Inset(a, 40, 46, 40, 16), -w * 0.06, -hp, w * 1.12, hp + es * 1.6);
            double t = T(m, 6);
            Floor(g, m.P(-w * 0.06, 0).X, m.P(w * 1.06, 0).X, m.P(0, -hp).Y);
            ConcreteRect(g, m, 0, 0, w, es);
            if (fon)
            {
                Board(g, m, 0, -t, w, t, Shade(c), false);
                int nv = Math.Max(4, (int)(w / 50));
                for (int i = 0; i <= nv; i++)
                    Board(g, m, w * i / nv - t, -t * 3, t * 2, t * 2, Madera, false);
                int np = Math.Max(2, (int)(w / 100));
                for (int i = 0; i <= np; i++)
                {
                    double x = w * (i + 0.5) / (np + 1);
                    Prop(g, m.P(x, -t * 3), m.P(x, -hp));
                }
            }
            if (lat)
            {
                Board(g, m, -t, -t, t, es + t * 2, c, true);
                Board(g, m, w, -t, t, es + t * 2, c, true);
            }
            DimV(g, m.P(w + t, 0), m.P(w + t, es), 16, $"e = {es:0.#} cm");

            Map k = Map.Fit(Inset(e, 40, 40, 40, 96), 0, 0, lx, ly);
            double tk = T(k, 5);
            ConcreteRect(g, k, 0, 0, lx, ly);
            if (fon)
            {
                PointF p0 = k.P(0, ly);
                BoardPx(g, new RectangleF(p0.X, p0.Y, k.L(lx), k.L(ly)), Color.FromArgb(150, Shade(c)), false);
            }
            if (lat)
                using (var pen = new Pen(c, Math.Max(4f, k.L(tk))))
                {
                    PointF p0 = k.P(0, ly);
                    g.DrawRectangle(pen, p0.X, p0.Y, k.L(lx), k.L(ly));
                }
            DimH(g, k.P(0, 0), k.P(lx, 0), 14, $"{lx / 100:0.00} m");
            DimV(g, k.P(0, 0), k.P(0, ly), -14, $"{ly / 100:0.00} m");
        }

        // ================================================================== MURO

        private static void Muro(Graphics g, RectangleF r, Color c, bool lat, bool fon, Medidas d)
        {
            SplitH(r, 0.42f, out RectangleF a, out RectangleF e);
            Caption(g, a, "CORTE (planta)");
            Caption(g, e, "ELEVACIÓN");
            double L = d.A, tw = d.B, H = d.C;

            double w = Math.Min(L, Math.Max(4 * tw, 90));
            Map m = Map.Fit(Inset(a, 30, 60, 30, 110), -tw * 0.5, -tw * 1.2, w + tw * 0.5, tw * 3.4);
            double t = T(m, 6);
            ConcreteRect(g, m, 0, 0, w, tw);
            if (lat)
            {
                Board(g, m, 0, -t, w, t, c, false);
                Board(g, m, 0, tw, w, t, c, false);
                Board(g, m, -t, -t, t, tw + 2 * t, c, true);
                double s = T(m, 7);
                for (double x = 12; x < w; x += 30)
                {
                    Board(g, m, x, -t - s, s, s, Madera, false);
                    Board(g, m, x, tw + t, s, s, Madera, false);
                }
                using (var pen = new Pen(Color.FromArgb(70, 70, 75), 1.6f) { DashStyle = DashStyle.Dash })
                    for (double x = 27; x < w; x += 60)
                        g.DrawLine(pen, m.P(x, -t - s * 1.6), m.P(x, tw + t + s * 1.6));
            }
            using (var pen = new Pen(Dim, 1f) { DashStyle = DashStyle.Dash })
                g.DrawLine(pen, m.P(w, -tw), m.P(w, tw * 2));
            DimV(g, m.P(-t, 0), m.P(-t, tw), -18, $"t = {tw:0.#} cm");

            Map k = Map.Fit(Inset(e, 40, 50, 30, 110), 0, 0, L, H);
            ConcreteRect(g, k, 0, 0, L, H);
            if (lat)
            {
                PointF p0 = k.P(0, H);
                var rc = new RectangleF(p0.X, p0.Y, k.L(L), k.L(H));
                BoardPx(g, rc, c, true);
                using (var pen = new Pen(Madera, 1.4f))
                {
                    for (double x = 120; x < L; x += 120) g.DrawLine(pen, k.P(x, 0), k.P(x, H));
                    for (double z = 240; z < H; z += 240) g.DrawLine(pen, k.P(0, z), k.P(L, z));
                }
                using (var b = new SolidBrush(Color.FromArgb(60, 60, 65)))
                    for (double x = 60; x < L; x += 60)
                        for (double z = 30; z < H; z += 60)
                        {
                            PointF q = k.P(x, z);
                            g.FillEllipse(b, q.X - 2, q.Y - 2, 4, 4);
                        }
            }
            DimH(g, k.P(0, 0), k.P(L, 0), 16, $"{L / 100:0.00} m");
            DimV(g, k.P(L, 0), k.P(L, H), 16, $"{H / 100:0.00} m");
            if (fon)
                Lines(g, new RectangleF(e.X + 12, e.Bottom - 104, e.Width - 16, 20), ("Incluye fondos de vanos (dinteles)", Shade(c)));
        }

        // ================================================================== ESCALERA

        private static void Escalera(Graphics g, RectangleF r, Color c, bool lat, bool fon, Medidas d)
        {
            Caption(g, r, "CORTE LONGITUDINAL");
            double L = Math.Max(d.A, 100), H = Math.Max(d.C, 60);
            int n = Math.Max(2, (int)Math.Round(H / 17.5));
            double cp = H / n, ps = L / n, garganta = 15;

            Map m = Map.Fit(Inset(r, 50, 50, 50, 110), -20, -garganta * 3, L + 40, H + garganta * 3.5);
            double t = T(m, 6);
            Floor(g, m.P(-20, 0).X, m.P(L + 20, 0).X, m.P(0, 0).Y);

            var pts = new List<PointF> { m.P(0, 0) };
            for (int i = 0; i < n; i++)
            {
                pts.Add(m.P(i * ps, (i + 1) * cp));
                pts.Add(m.P((i + 1) * ps, (i + 1) * cp));
            }
            double ang = Math.Atan2(H, L), off = garganta / Math.Cos(ang);
            pts.Add(m.P(L, H - off));
            pts.Add(m.P(ps * 1.0, 0));
            Polygon(g, pts.ToArray());

            if (fon)
            {
                using (var pen = new Pen(Shade(c), Math.Max(5f, m.L(t))))
                    g.DrawLine(pen, m.P(ps + t, -t), m.P(L + t, H - off - t * 1.2));
                for (int i = 1; i < n; i++)
                {
                    double x = ps + (L - ps) * i / n;
                    double y = (x - ps) * (H - off) / (L - ps) - t * 1.6;
                    Prop(g, m.P(x, y), m.P(x, 0));
                }
            }
            if (lat)
                for (int i = 0; i < n; i++)
                    Board(g, m, i * ps - t, i * cp, t, cp, c, true);

            DimH(g, m.P(0, 0), m.P(L, 0), 18, $"{L / 100:0.00} m");
            DimV(g, m.P(L, 0), m.P(L, H), 24, $"{H / 100:0.00} m");
        }

        // ================================================================== TODO

        public static void Todo(Graphics g, RectangleF r, FormValues v, Func<Pieza, int> count)
        {
            Caption(g, r, "ELEMENTOS A ENCOFRAR");
            float cw = (r.Width - 40) / 3, ch = (r.Height - 90) / 2;
            int i = 0;
            foreach (Pieza p in Piezas.Todas)
            {
                float x = r.Left + 20 + (i % 3) * cw, y = r.Top + 50 + (i / 3) * ch;
                bool on = v.B("p_" + p);
                float s = Math.Min(cw, ch) - 46;
                GraphicsState st = g.Save();
                g.TranslateTransform(x + (cw - s) / 2, y + 4);
                Iconos.Draw(p, g, s);
                g.Restore(st);
                if (!on)
                    using (var veil = new SolidBrush(Color.FromArgb(185, 255, 255, 255)))
                        g.FillRectangle(veil, x, y, cw, ch);
                using (var b = new SolidBrush(on ? Ink : Color.Silver))
                {
                    var sf = new StringFormat { Alignment = StringAlignment.Center };
                    g.DrawString($"{(on ? "☑" : "☐")} {Piezas.Plural(p)} ({count(p)})", Bold, b, new RectangleF(x, y + s + 10, cw, 22), sf);
                }
                i++;
            }
        }

        // ================================================================== Primitivas

        /// <summary>Espesor dibujado del tablero: al menos <paramref name="px"/> píxeles.</summary>
        private static double T(Map m, float px) => px / Math.Max(m.S, 1e-6f);

        private static Color Shade(Color c) => Color.FromArgb(c.A, c.R * 3 / 4, c.G * 3 / 4, c.B * 3 / 4);

        private static void Board(Graphics g, Map m, double x, double y, double w, double h, Color c, bool vertical)
        {
            PointF p0 = m.P(x, y + h);
            BoardPx(g, new RectangleF(p0.X, p0.Y, m.L(w), m.L(h)), c, vertical);
        }

        /// <summary>Tablero de madera: color del elemento con vetas en el sentido del tablón.</summary>
        private static void BoardPx(Graphics g, RectangleF rc, Color c, bool vertical)
        {
            if (rc.Width <= 0 || rc.Height <= 0) return;
            using (var b = new SolidBrush(c)) g.FillRectangle(b, rc);
            using (var p = new Pen(Color.FromArgb(70, 60, 30, 0), 1f))
            {
                float step = 9;
                if (vertical) for (float x = rc.Left + step; x < rc.Right - 1; x += step) g.DrawLine(p, x, rc.Top, x, rc.Bottom);
                else for (float y = rc.Top + step; y < rc.Bottom - 1; y += step) g.DrawLine(p, rc.Left, y, rc.Right, y);
            }
            using (var p = new Pen(Madera, 1.2f)) g.DrawRectangle(p, rc.X, rc.Y, rc.Width, rc.Height);
        }

        /// <summary>Puntal (pie derecho) con placas en los extremos.</summary>
        private static void Prop(Graphics g, PointF top, PointF bottom)
        {
            using (var p = new Pen(Madera, 4f)) g.DrawLine(p, top, bottom);
            using (var p = new Pen(Madera, 2f))
            {
                g.DrawLine(p, top.X - 7, top.Y, top.X + 7, top.Y);
                g.DrawLine(p, bottom.X - 7, bottom.Y, bottom.X + 7, bottom.Y);
            }
        }

        private static void Floor(Graphics g, float x0, float x1, float y)
        {
            using (var p = new Pen(Suelo, 2f)) g.DrawLine(p, x0, y, x1, y);
            using (var p = new Pen(Suelo, 1f))
                for (float x = x0 + 4; x < x1; x += 10) g.DrawLine(p, x, y, x - 6, y + 6);
        }
    }
}
