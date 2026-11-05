using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using AceroRefuerzo.Core;
using static AceroRefuerzo.UI.Dibujo;

namespace AceroRefuerzo.UI
{
    /// <summary>
    /// "Figuritas" de cada elemento: secciones y elevaciones dibujadas a escala con los valores
    /// del formulario (recubrimiento, diámetros, cantidades, separaciones y distribución de estribos).
    /// </summary>
    internal static class Figures
    {
        private static readonly Color Steel = Color.FromArgb(200, 35, 35);
        private static readonly Color Steel2 = Color.FromArgb(225, 120, 20);
        private static readonly Color Tie = Color.FromArgb(25, 105, 200);

        // ================================================================== VIGA

        public static void Viga(Graphics g, RectangleF r, FormValues v, double b, double h, double L, List<(double xa, double xb)> luces)
        {
            if (luces == null || luces.Count == 0) luces = new List<(double, double)> { (0, L) };
            bool continua = v.Has("modoLong") && v.I("modoLong") == 1 && luces.Count >= 1;
            SplitH(r, 0.40f, out RectangleF a, out RectangleF e);
            Caption(g, a, "SECCIÓN TRANSVERSAL");
            Caption(g, e, luces.Count > 1 ? $"ELEVACIÓN: VIGA CONTINUA DE {luces.Count} LUCES (esquema)" : "ELEVACIÓN (esquema)");

            double c = v.D("rec"), ds = v.BarMm("barEst") / 10, dbS = v.BarMm("barSup") / 10, dbI = v.BarMm("barInf") / 10;
            int nS = v.I("nSup"), nI = v.I("nInf"), bsN = v.I("bsN"), biN = v.I("biN");
            double dbBs = v.BarMm("bsBar") / 10, dbBi = v.BarMm("biBar") / 10;

            // --- Sección
            RectangleF sa = Inset(a, 40, 46, 30, 112);
            Map m = Map.Fit(sa, 0, 0, b, h);
            ConcreteRect(g, m, 0, 0, b, h);
            Stirrup(g, m, c + ds / 2, c + ds / 2, b - c - ds / 2, h - c - ds / 2, ds);
            Row(g, m, nS, c + ds + dbS / 2, b - c - ds - dbS / 2, h - c - ds - dbS / 2, dbS, Steel);
            Row(g, m, nI, c + ds + dbI / 2, b - c - ds - dbI / 2, c + ds + dbI / 2, dbI, Steel);
            if (bsN > 0)
                Row(g, m, bsN, c + ds + dbBs / 2, b - c - ds - dbBs / 2, h - c - ds - dbS - Math.Max(2.5, dbBs) - dbBs / 2, dbBs, Steel2);
            if (biN > 0)
                Row(g, m, biN, c + ds + dbBi / 2, b - c - ds - dbBi / 2, c + ds + dbI + Math.Max(2.5, dbBi) + dbBi / 2, dbBi, Steel2);
            DimH(g, m.P(0, 0), m.P(b, 0), 18, $"b = {b:0.#} cm");
            DimV(g, m.P(0, 0), m.P(0, h), -18, $"h = {h:0.#} cm");

            // --- Barras corridas: una pieza, o tramos óptimos con traslapes en la zona permitida (vigas continuas).
            double prol = v.D("prol");
            double xr0 = luces[0].xa, xr1 = luces[luces.Count - 1].xb;
            double x0 = prol > 0 ? xr0 - prol : xr0 + c, x1 = prol > 0 ? xr1 + prol : xr1 - c;
            double[] comerciales;
            try { comerciales = Traslapes.Comerciales(v, x => x); } catch (Exception) { comerciales = new[] { 900.0 }; }
            double lcom = comerciales.Max();
            string errTramos = null;

            (List<(double Ini, double Fin)> tramos, List<double> largos, double leg, double db, double ls) Fila(string barKey, string pataKey, string zonaKey)
            {
                double dbc = v.BarMm(barKey) / 10;
                double legc = Math.Min(Empalmes.PataEje(v.D(pataKey), dbc), h - 2 * c);
                double extra = legc > 0 ? legc - Empalmes.Ahorro90(dbc, v.BendMm(barKey) / 10) : 0;
                double lsc = Traslapes.Ls(v, "ls", dbc, x => x);
                List<(double, double)> tr = null;
                if (continua)
                {
                    try
                    {
                        Traslapes.Resultado opt = Traslapes.Viga(v, luces, h, x0, x1, extra, extra, dbc, zonaKey, x => x);
                        if (opt == null) errTramos = errTramos ?? "No hay forma de ubicar los traslapes en la zona permitida con esas barras comerciales.";
                        else tr = opt.Tramos;
                    }
                    catch (Exception ex) { errTramos = errTramos ?? ex.Message; }
                }
                if (tr == null) tr = new List<(double, double)> { (x0, x1) };
                var lg = tr.Select((t, k) => t.Item2 - t.Item1 + (k == 0 ? extra : 0) + (k == tr.Count - 1 ? extra : 0)).ToList();
                return (tr, lg, legc, dbc, lsc);
            }
            var sup = Fila("barSup", "pataSup", "zonaSup");
            var inf = Fila("barInf", "pataInf", "zonaInf");

            (string, Color) Linea(string titulo, int n, string barKey, (List<(double Ini, double Fin)> tramos, List<double> largos, double leg, double db, double ls) fila)
            {
                bool pasa = fila.largos.Any(l => l > lcom + 0.5);
                string txt = fila.tramos.Count == 1
                    ? $"{titulo}: {n} Ø {v.BarName(barKey)} · L = {fila.largos[0] / 100:0.00} m" + (pasa ? $" > {lcom / 100:0.#} m" : "")
                    : $"{titulo}: {n} Ø {v.BarName(barKey)} · {fila.tramos.Count} tramos: " + string.Join(" + ", fila.largos.Select(l => $"{l / 100:0.00}")) + " m";
                if (fila.tramos.Count > 1 && !pasa)
                {
                    string corte = CorteTexto(fila.largos, n, comerciales, v.BarName(barKey), fila.db);
                    int i = corte.IndexOf("desperdicio", StringComparison.Ordinal);
                    if (i >= 0) txt += " · " + corte.Substring(i);
                }
                return (txt, pasa ? Theme.Accent : Steel);
            }
            Lines(g, new RectangleF(a.X + 12, a.Bottom - 100, a.Width - 16, 98),
                Linea("Superior", nS, "barSup", sup),
                Linea("Inferior", nI, "barInf", inf),
                (bsN + biN > 0 ? $"Bastones: {bsN} Ø {v.BarName("bsBar")} sup. · {biN} Ø {v.BarName("biBar")} inf." : "Sin bastones", Steel2),
                ($"Estribos Ø {v.BarName("barEst")}  ·  rec. {c:0.#} cm", Tie));

            // --- Elevación
            RectangleF ea = Inset(e, 34, 60, 34, 118);
            double Ltot = Math.Max(1, xr1 - xr0);
            float hpx = Math.Min(ea.Height * 0.55f, Math.Max(60f, ea.Width * (float)(h / Ltot) * 2.5f));
            float y0 = ea.Top + (ea.Height - hpx) / 2;
            var beam = new RectangleF(ea.Left, y0, ea.Width, hpx);
            double sx = beam.Width / Ltot, sz = hpx / h;
            float col = 16;
            Func<double, float> X = x => beam.Left + (float)((x - xr0) * sx);
            Func<double, float> Z = z => beam.Bottom - (float)(z * sz);

            using (var colB = new SolidBrush(Color.FromArgb(200, 203, 208)))
            {
                g.FillRectangle(colB, beam.Left - col, beam.Top - 30, col, hpx + 60);
                g.FillRectangle(colB, beam.Right, beam.Top - 30, col, hpx + 60);
                for (int i = 0; i + 1 < luces.Count; i++)
                {
                    float xa = X(luces[i].xb), xb = X(luces[i + 1].xa);
                    g.FillRectangle(colB, xa, beam.Top - 30, Math.Max(3, xb - xa), hpx + 60);
                }
            }
            foreach (var (xa, xb) in luces)
                Fill(g, new RectangleF(X(xa), beam.Top, X(xb) - X(xa), hpx));

            // Zonas permitidas para los traslapes (verde): superior en la parte alta, inferior en la baja.
            if (continua)
                using (var zb2 = new SolidBrush(Color.FromArgb(90, 120, 200, 120)))
                {
                    foreach (var (za, zc) in Traslapes.ZonasViga(luces, h, v.I("zonaSup")))
                        g.FillRectangle(zb2, X(za), beam.Top + 1, Math.Max(1, X(zc) - X(za)), hpx * 0.35f);
                    foreach (var (za, zc) in Traslapes.ZonasViga(luces, h, v.I("zonaInf")))
                        g.FillRectangle(zb2, X(za), beam.Bottom - hpx * 0.35f - 1, Math.Max(1, X(zc) - X(za)), hpx * 0.35f);
                }

            List<double> pos = null;
            string err = null;
            try
            {
                foreach (var (xa, xb) in luces)
                {
                    var p0 = Distribution.Positions(v.S("dist"), Un.Cm(xb - xa)).Select(Un.ToCm).ToList();
                    if (pos == null) pos = p0;
                    using (var p = new Pen(Tie, 1.2f))
                        foreach (double x in p0) g.DrawLine(p, X(xa + x), Z(c), X(xa + x), Z(h - c));
                }
            }
            catch (Exception ex) { err = ex.Message; }

            double zt = h - c - ds - dbS / 2, zb = c + ds + dbI / 2;
            void Corrida((List<(double Ini, double Fin)> tramos, List<double> largos, double leg, double db, double ls) fila, double z, int dir)
            {
                for (int k = 0; k < fila.tramos.Count; k++)
                {
                    var (ini, fin) = fila.tramos[k];
                    double zk = z + (k % 2 == 1 ? dir * fila.db : 0);
                    using (var p = new Pen(k % 2 == 0 ? Steel : Steel2, 2.6f))
                        LongBar(g, p, X, Z, ini, fin, zk, dir * fila.leg, k == 0, k == fila.tramos.Count - 1);
                    if (k > 0)
                        using (var p = new Pen(Theme.Accent, 1.5f))
                        {
                            float ya = Z(zk) + dir * 5, xa = X(ini), xb = X(ini + fila.ls);
                            g.DrawLine(p, xa, ya, xb, ya);
                            g.DrawLine(p, xa, ya - 3, xa, ya + 3);
                            g.DrawLine(p, xb, ya - 3, xb, ya + 3);
                        }
                }
            }
            Corrida(sup, zt, -1);
            Corrida(inf, zb, +1);

            using (var p = new Pen(Steel2, 2.2f))
            {
                if (bsN > 0)
                {
                    double z2 = zt - Math.Max(2.5, dbBs) - dbS / 2 - dbBs / 2;
                    double Len(int i) => v.D("bsFr") * (luces[i].xb - luces[i].xa);
                    int last = luces.Count - 1;
                    LongBar(g, p, X, Z, x0, luces[0].xa + Len(0), z2, -Math.Min(v.D("pataSup"), z2 - c), true, false);
                    LongBar(g, p, X, Z, luces[last].xb - Len(last), x1, z2, -Math.Min(v.D("pataSup"), z2 - c), false, true);
                    for (int i = 0; i + 1 < luces.Count; i++)
                        LongBar(g, p, X, Z, luces[i].xb - Len(i), luces[i + 1].xa + Len(i + 1), z2, 0, false, false);
                }
                if (biN > 0)
                {
                    double z2 = zb + Math.Max(2.5, dbBi) + dbI / 2 + dbBi / 2;
                    foreach (var (xa, xb) in luces)
                    {
                        double len = v.D("biFr") * (xb - xa), mid = (xa + xb) / 2;
                        LongBar(g, p, X, Z, mid - len / 2, mid + len / 2, z2, 0, false, false);
                    }
                }
            }
            DimH(g, new PointF(beam.Left, beam.Bottom), new PointF(beam.Right, beam.Bottom), 40,
                luces.Count > 1 ? $"L total = {Ltot / 100:0.00} m ({luces.Count} luces)" : $"L libre = {L / 100:0.00} m");

            var abajo = new List<(string, Color)> { ($"Estribos: {v.S("dist")}", Tie) };
            if (pos != null) abajo.Add((luces.Count > 1 ? $"{pos.Count} estribos en la primera viga" : $"{pos.Count} estribos por viga", Ink));
            if (continua && errTramos == null)
            {
                string Pos(List<(double Ini, double Fin)> tr) => tr.Count > 1
                    ? string.Join(", ", tr.Skip(1).Select(t => $"{(t.Ini - xr0) / 100:0.00}"))
                    : "sin traslape";
                abajo.Add(($"Traslapes (m desde el inicio): sup. {Pos(sup.tramos)} · inf. {Pos(inf.tramos)}", Theme.Accent));
            }
            string error = err ?? errTramos;
            if (error != null) Error(g, new RectangleF(e.X + 12, e.Bottom - 74, e.Width - 20, 70), error);
            else Lines(g, new RectangleF(e.X + 12, e.Bottom - 74, e.Width - 20, 70), abajo.ToArray());
        }

        private static void LongBar(Graphics g, Pen p, Func<double, float> X, Func<double, float> Z,
            double x0, double x1, double z, double leg, bool legStart, bool legEnd)
        {
            var pts = new List<PointF>();
            if (Math.Abs(leg) > 0.1 && legStart) pts.Add(new PointF(X(x0), Z(z + leg)));
            pts.Add(new PointF(X(x0), Z(z)));
            pts.Add(new PointF(X(x1), Z(z)));
            if (Math.Abs(leg) > 0.1 && legEnd) pts.Add(new PointF(X(x1), Z(z + leg)));
            g.DrawLines(p, pts.ToArray());
        }

        // ================================================================== COLUMNA

        public static void Columna(Graphics g, RectangleF r, FormValues v, List<P2> polyCm, bool contour, List<(double zb, double zt)> pisos)
        {
            SplitH(r, 0.62f, out RectangleF a, out RectangleF e);
            Caption(g, a, contour ? "SECCIÓN (contorno real)" : "SECCIÓN TRANSVERSAL");
            Caption(g, e, "ELEVACIÓN");

            double c = v.D("rec"), ds = v.BarMm("barEst") / 10, db = v.BarMm("barLong") / 10;
            double minX = polyCm.Min(p => p.X), maxX = polyCm.Max(p => p.X), minY = polyCm.Min(p => p.Y), maxY = polyCm.Max(p => p.Y);
            double bx = maxX - minX, by = maxY - minY;

            RectangleF sa = Inset(a, 44, 46, 30, 80);
            Map m = Map.Fit(sa, minX, minY, bx, by);
            int nBars;
            string estribos = $"Estribos Ø {v.BarName("barEst")}: {v.S("dist")}";
            int porNivel = 1;

            if (!contour)
            {
                ConcreteRect(g, m, minX, minY, bx, by);
                int nX = Math.Max(2, v.I("nX")), nY = Math.Max(2, v.I("nY"));
                double xA = minX + c + ds + db / 2, xB = maxX - c - ds - db / 2, yA = minY + c + ds + db / 2, yB = maxY - c - ds - db / 2;
                Stirrup(g, m, minX + c + ds / 2, minY + c + ds / 2, maxX - c - ds / 2, maxY - c - ds / 2, ds);
                if (v.B("grapas"))
                    using (var p = new Pen(Tie, Math.Max(1.2f, m.L(ds))))
                    {
                        for (int i = 1; i < nX - 1; i++)
                        {
                            double x = xA + (xB - xA) * i / (nX - 1) + db / 2 + ds / 2;
                            g.DrawLine(p, m.P(x, minY + c + ds / 2), m.P(x, maxY - c - ds / 2));
                        }
                        for (int j = 1; j < nY - 1; j++)
                        {
                            double y = yA + (yB - yA) * j / (nY - 1) + db / 2 + ds / 2;
                            g.DrawLine(p, m.P(minX + c + ds / 2, y), m.P(maxX - c - ds / 2, y));
                        }
                    }
                Row(g, m, nX, xA, xB, yA, db, Steel);
                Row(g, m, nX, xA, xB, yB, db, Steel);
                for (int j = 1; j < nY - 1; j++)
                {
                    double y = yA + (yB - yA) * j / (nY - 1);
                    Dot(g, m.P(xA, y), Math.Max(2.5f, m.L(db / 2)), Steel);
                    Dot(g, m.P(xB, y), Math.Max(2.5f, m.L(db / 2)), Steel);
                }
                nBars = 2 * nX + 2 * (nY - 2);
            }
            else
            {
                Polygon(g, polyCm.Select(p => m.P(p.X, p.Y)).ToArray());
                List<P2> bars = Polygon2D.OffsetInward(polyCm, c + ds + db / 2);
                var alas = v.Has("tipoEst") && v.I("tipoEst") == 0 ? Polygon2D.CoverRectangles(polyCm) : null;
                if (alas != null && alas.Count >= 2)
                {
                    Color[] tonos = { Tie, Color.FromArgb(20, 70, 150), Color.FromArgb(70, 150, 230), Color.FromArgb(10, 50, 110) };
                    for (int k = 0; k < alas.Count; k++)
                    {
                        var (x0, y0, x1, y1) = alas[k];
                        Stirrup(g, m, x0 + c + ds / 2, y0 + c + ds / 2, x1 - c - ds / 2, y1 - c - ds / 2, ds, tonos[k % tonos.Length]);
                    }
                    estribos = $"Estribos Ø {v.BarName("barEst")}: {alas.Count} por nivel (uno por ala) · {v.S("dist")}";
                    porNivel = alas.Count;
                }
                else
                {
                    alas = null;
                    List<P2> tie = Polygon2D.OffsetInward(polyCm, c + ds / 2);
                    using (var p = new Pen(Tie, Math.Max(1.5f, m.L(ds))) { LineJoin = LineJoin.Round })
                        g.DrawPolygon(p, tie.Select(q => m.P(q.X, q.Y)).ToArray());
                }
                List<P2> pts = Polygon2D.PerimeterPoints(bars, v.D("smax"), Math.Max(4, v.I("nMin")),
                    alas != null ? Polygon2D.RectCorners(alas, c + ds + db / 2) : null);
                foreach (P2 q in pts) Dot(g, m.P(q.X, q.Y), Math.Max(2.5f, m.L(db / 2)), Steel);
                nBars = pts.Count;
            }
            DimH(g, m.P(minX, minY), m.P(maxX, minY), 18, $"{bx:0.#} cm");
            DimV(g, m.P(minX, minY), m.P(minX, maxY), -18, $"{by:0.#} cm");

            Lines(g, new RectangleF(a.X + 12, a.Bottom - 60, a.Width - 16, 58),
                ($"{nBars} barras Ø {v.BarName("barLong")}", Steel),
                (estribos, Tie),
                ($"Recubrimiento {c:0.#} cm", Ink));

            // --- Elevación (uno o varios pisos) con los tramos de barra y sus traslapes
            double H = pisos[pisos.Count - 1].zt - pisos[0].zb;
            double anc = v.D("anc"), emp = v.D("emp"), pata = v.D("pata");
            double bend = v.BendMm("barLong") / 10;
            double pataEje = Empalmes.PataEje(pata, db);
            double extra = pataEje > 0 ? pataEje - Empalmes.Ahorro90(db, bend) : 0;
            double zIni = anc > 0 ? -anc : c, zTope = emp > 0 ? H + emp : H - c;
            int modoLong = v.Has("modoLong") ? v.I("modoLong") : 0;
            bool manual = modoLong == 1, optimo = modoLong == 2;
            double ls = Traslapes.Ls(v, "ls", db, x => x);

            string errTramos = null;
            List<(double Ini, double Fin)> tramos;
            List<(double a, double b)> zonas = null;
            try
            {
                if (optimo)
                {
                    zonas = Traslapes.ZonasColumna(pisos, v.I("zonaCol"), v.D("zonaIni"), v.D("zonaFin"));
                    Traslapes.Resultado opt = Traslapes.Columna(v, pisos, zIni, zTope, extra, db, x => x);
                    if (opt == null) throw new InvalidOperationException("No hay solución con esa zona permitida y esas barras comerciales.");
                    tramos = opt.Tramos;
                }
                else
                {
                    List<double> largos = manual ? Empalmes.LeerTramos(v.S("tramos")) : null;
                    if (manual && largos == null) throw new FormatException("Escriba la longitud de los tramos (m).");
                    tramos = largos != null
                        ? Empalmes.UbicarTramos(zIni, largos.Select(m => m * 100).ToList(), extra, ls)
                        : new List<(double, double)> { (zIni, zTope) };
                }
            }
            catch (Exception ex)
            {
                errTramos = ex.Message;
                tramos = new List<(double, double)> { (zIni, zTope) };
            }

            double zMin = Math.Min(0, zIni) - 10, zMax = Math.Max(zTope, tramos.Max(t => t.Fin)) + 10;
            RectangleF ea = Inset(e, 34, 50, 40, 86);
            float w = Math.Min(ea.Width * 0.35f, 60f);
            float sc = (float)(ea.Height / (zMax - zMin));
            float cx = ea.Left + ea.Width * 0.42f;
            Func<double, float> Y = z => ea.Bottom - (float)((z - zMin) * sc);

            if (anc > 0)
                using (var fb = new SolidBrush(Color.FromArgb(210, 213, 218)))
                    g.FillRectangle(fb, cx - w * 1.4f, Y(0), w * 2.8f, Y(zMin) - Y(0));

            List<double> pos = null;
            string err = null;
            for (int i = 0; i < pisos.Count; i++)
            {
                var (zb, zt) = pisos[i];
                var colR = new RectangleF(cx - w / 2, Y(zt), w, Y(zb) - Y(zt));
                Fill(g, colR);
                if (pisos.Count > 1)
                    using (var tb = new SolidBrush(Dim)) g.DrawString($"P{i + 1}", Small, tb, colR.Left - 24, (colR.Top + colR.Bottom) / 2 - 7);
                if (zonas != null)
                    using (var zb2 = new SolidBrush(Color.FromArgb(90, 120, 200, 120)))
                        foreach (var (za, zc) in zonas.Where(q => q.a >= zb - 0.5 && q.b <= zt + 0.5))
                            g.FillRectangle(zb2, colR.Left + 1, Y(zc), colR.Width - 2, Y(za) - Y(zc));
                try
                {
                    var p0 = Distribution.Positions(v.S("dist"), Un.Cm(zt - zb)).Select(Un.ToCm).ToList();
                    if (i == 0) pos = p0;
                    using (var p = new Pen(Tie, 1f))
                        foreach (double z in p0) g.DrawLine(p, colR.Left + 4, Y(zb + z), colR.Right - 4, Y(zb + z));
                }
                catch (Exception ex) { err = ex.Message; }
            }

            // Barras: cada tramo a un lado del anterior; zona de traslape resaltada.
            for (int k = 0; k < tramos.Count; k++)
            {
                var (ini, fin) = tramos[k];
                float dx = k % 2 == 1 ? 4 : 0;
                using (var p = new Pen(k % 2 == 0 ? Steel : Steel2, 2.4f))
                    foreach (float x in new[] { cx - w / 2 + 7 + dx, cx + w / 2 - 7 - dx })
                    {
                        g.DrawLine(p, x, Y(ini), x, Y(fin));
                        if (k == 0 && pata > 0)
                        {
                            float dir = x < cx ? 1 : -1;
                            g.DrawLine(p, x, Y(ini), x + dir * Math.Min((float)(pata * sc) + 6, w * 0.8f), Y(ini));
                        }
                    }
                if (k > 0)
                {
                    double t0 = ini, t1 = tramos[k - 1].Fin;
                    float xr = cx + w / 2 + 6;
                    using (var p = new Pen(Theme.Accent, 2f)) g.DrawLine(p, xr, Y(t0), xr, Y(t1));
                    using (var tb = new SolidBrush(Theme.Accent))
                        g.DrawString($"traslape\n{t0 / 100:0.00}–{t1 / 100:0.00} m", Small, tb, xr + 3, (Y(t0) + Y(t1)) / 2 - 14);
                }
            }
            DimV(g, new PointF(cx - w / 2, Y(pisos[0].zb)), new PointF(cx - w / 2, Y(H)), pisos.Count > 1 ? -30 : -18, $"{H / 100:0.00} m");

            // Longitud desarrollada de cada tramo (incluye la pata y el doblez del primero).
            double[] comerciales;
            try { comerciales = Traslapes.Comerciales(v, x => x); } catch (Exception) { comerciales = new[] { 900.0 }; }
            double lcom = comerciales.Max();
            var largosTxt = tramos.Select((t, k) => t.Fin - t.Ini + (k == 0 ? extra : 0)).ToList();
            var lineas = new List<(string, Color)>
            {
                ((tramos.Count > 1 ? "Tramos: " : "Barra: ") + string.Join(" · ", largosTxt.Select(l => $"{l / 100:0.00}")) + " m", Steel)
            };
            if (largosTxt.Any(l => l > lcom + 0.5)) lineas.Add(($"Supera la barra comercial de {lcom / 100:0.00} m", Theme.Accent));
            else if (optimo && errTramos == null)
                lineas.Add((CorteTexto(largosTxt, nBars, comerciales, v.BarName("barLong"), db), Color.FromArgb(40, 130, 70)));
            if (manual && errTramos == null && tramos[tramos.Count - 1].Fin < zTope - 0.5)
                lineas.Add(($"No llega al tope: faltan {(zTope - tramos[tramos.Count - 1].Fin) / 100:0.00} m", Theme.Accent));
            if (pos != null)
                lineas.Add((porNivel > 1 ? $"Estribos: {pos.Count} niveles × {porNivel} por piso" : $"Estribos: {pos.Count} por piso", Tie));

            string error = errTramos ?? err;
            if (error != null) Error(g, new RectangleF(e.X + 6, e.Bottom - 80, e.Width - 10, 76), error);
            else Lines(g, new RectangleF(e.X + 6, e.Bottom - 80, e.Width - 10, 76), lineas.ToArray());
        }

        /// <summary>Resumen del corte de los tramos de un elemento: barras comerciales por juego y desperdicio.</summary>
        private static string CorteTexto(IList<double> largosCm, int nBarras, double[] comercialesCm, string barName, double dbCm)
        {
            var piezas = largosCm.Select(l => new PiezaCorte { Bar = barName, Mm = dbCm * 10, LargoCm = l, Cantidad = Math.Max(1, nBarras) }).ToList();
            ResultadoCorte r = Corte.Optimizar(piezas, new OpcionesCorte { ComercialesCm = comercialesCm, Empalmar = false, RetazoMinCm = 0 });
            if (r.Diametros.Count == 0 || r.Diametros[0].BarrasComerciales == 0) return "";
            CorteDiametro d = r.Diametros[0];
            string barras = string.Join(" + ", d.BarrasPorComercial.Select(kv => $"{kv.Value} de {kv.Key / 100:0.##} m"));
            return $"Corte: {barras} · desperdicio {d.Porcentaje:0.0} %";
        }

        // ================================================================== ZAPATA

        public static void Zapata(Graphics g, RectangleF r, FormValues v, double lx, double ly, double hz)
        {
            SplitV(r, 0.55f, out RectangleF a, out RectangleF e);
            Caption(g, a, "PLANTA - parrilla inferior");
            Caption(g, e, "CORTE");

            double c = v.D("rec"), dx = v.BarMm("barX") / 10, dy = v.BarMm("barY") / 10;
            double sX = Math.Max(1, v.D("sepX")), sY = Math.Max(1, v.D("sepY"));

            Map m = Map.Fit(Inset(a, 40, 40, 40, 30), 0, 0, lx, ly);
            ConcreteRect(g, m, 0, 0, lx, ly);
            int nX = Count(ly - 2 * c - dx, sX), nY = Count(lx - 2 * c - dy, sY);
            using (var p = new Pen(Steel, 1.6f))
                for (int i = 0; i < nX; i++)
                {
                    double y = c + dx / 2 + (nX > 1 ? (ly - 2 * c - dx) * i / (nX - 1) : (ly - 2 * c - dx) / 2);
                    g.DrawLine(p, m.P(c, y), m.P(lx - c, y));
                }
            using (var p = new Pen(Tie, 1.6f))
                for (int i = 0; i < nY; i++)
                {
                    double x = c + dy / 2 + (nY > 1 ? (lx - 2 * c - dy) * i / (nY - 1) : (lx - 2 * c - dy) / 2);
                    g.DrawLine(p, m.P(x, c), m.P(x, ly - c));
                }
            double cs = Math.Min(lx, ly) * 0.25;
            using (var cb = new SolidBrush(Color.FromArgb(150, 120, 125, 135)))
            {
                PointF p0 = m.P(lx / 2 - cs / 2, ly / 2 + cs / 2);
                g.FillRectangle(cb, p0.X, p0.Y, m.L(cs), m.L(cs));
            }
            DimH(g, m.P(0, 0), m.P(lx, 0), 18, $"{lx / 100:0.00} m");
            DimV(g, m.P(0, 0), m.P(0, ly), -18, $"{ly / 100:0.00} m");

            // --- Corte (a lo largo de X)
            RectangleF ca = Inset(e, 40, 40, 40, 70);
            double colH = hz * 0.8;
            Map k = Map.Fit(ca, 0, 0, lx, hz + colH);
            using (var cb = new SolidBrush(Color.FromArgb(205, 208, 214)))
            {
                PointF p0 = k.P(lx / 2 - cs / 2, hz + colH);
                g.FillRectangle(cb, p0.X, p0.Y, k.L(cs), k.L(colH));
            }
            ConcreteRect(g, k, 0, 0, lx, hz);
            double maxLeg = hz - 2 * c - dx - dy;
            double leg = Math.Min(v.D("pata"), maxLeg);
            using (var p = new Pen(Steel, 2.4f))
                Bent(g, p, k, c, lx - c, c + dx / 2, leg);
            for (int i = 0; i < nY; i++)
            {
                double x = c + dy / 2 + (nY > 1 ? (lx - 2 * c - dy) * i / (nY - 1) : (lx - 2 * c - dy) / 2);
                Dot(g, k.P(x, c + dx + dy / 2), Math.Max(2.2f, k.L(dy / 2)), Tie);
            }
            int nSX = 0, nSY = 0;
            if (v.B("sup"))
            {
                double dsx = v.BarMm("barSX") / 10, dsy = v.BarMm("barSY") / 10;
                double legS = Math.Min(v.D("pataS"), hz - 2 * c - dx - dy - dsx - dsy);
                using (var p = new Pen(Steel2, 2.2f))
                    Bent(g, p, k, c, lx - c, hz - c - dsx / 2, -legS);
                nSX = Count(ly - 2 * c - dsx, Math.Max(1, v.D("sepSX")));
                nSY = Count(lx - 2 * c - dsy, Math.Max(1, v.D("sepSY")));
                for (int i = 0; i < nSY; i++)
                {
                    double x = c + dsy / 2 + (nSY > 1 ? (lx - 2 * c - dsy) * i / (nSY - 1) : (lx - 2 * c - dsy) / 2);
                    Dot(g, k.P(x, hz - c - dsx - dsy / 2), Math.Max(2.2f, k.L(dsy / 2)), Steel2);
                }
            }
            DimV(g, k.P(lx, 0), k.P(lx, hz), 18, $"{hz:0.#} cm");

            Lines(g, new RectangleF(e.X + 12, e.Bottom - 62, e.Width - 16, 60),
                ($"Inferior X: {nX} Ø {v.BarName("barX")} @ {sX:0.#} cm", Steel),
                ($"Inferior Y: {nY} Ø {v.BarName("barY")} @ {sY:0.#} cm", Tie),
                (v.B("sup") ? $"Superior: {nSX} Ø {v.BarName("barSX")} (X) · {nSY} Ø {v.BarName("barSY")} (Y)" : "Sin parrilla superior", Steel2));
        }

        private static void Bent(Graphics g, Pen p, Map k, double x0, double x1, double z, double leg)
        {
            var pts = new List<PointF>();
            if (Math.Abs(leg) > 0.1) pts.Add(k.P(x0, z + leg));
            pts.Add(k.P(x0, z));
            pts.Add(k.P(x1, z));
            if (Math.Abs(leg) > 0.1) pts.Add(k.P(x1, z + leg));
            g.DrawLines(p, pts.ToArray());
        }

        // ================================================================== LOSA

        public static void Losa(Graphics g, RectangleF r, FormValues v, List<P2> outline, double e)
        {
            SplitV(r, 0.62f, out RectangleF a, out RectangleF b);
            Caption(g, a, "PLANTA - mallas");
            Caption(g, b, "CORTE TÍPICO");

            double lx = outline.Max(p => p.X), ly = outline.Max(p => p.Y);
            Map m = Map.Fit(Inset(a, 34, 40, 34, 28), 0, 0, lx, ly);
            PointF[] poly = outline.Select(p => m.P(p.X, p.Y)).ToArray();
            Polygon(g, poly);

            using (var path = new GraphicsPath())
            {
                path.AddPolygon(poly);
                GraphicsState st = g.Save();
                g.SetClip(path);
                Grid(g, m, lx, ly, v.B("iX"), v.D("sIX"), true, Steel, false);
                Grid(g, m, lx, ly, v.B("iY"), v.D("sIY"), false, Tie, false);
                Grid(g, m, lx, ly, v.B("sX"), v.D("sSX"), true, Steel2, true);
                Grid(g, m, lx, ly, v.B("sY"), v.D("sSY"), false, Color.FromArgb(130, 60, 170), true);
                g.Restore(st);
            }
            DimH(g, m.P(0, 0), m.P(lx, 0), 16, $"{lx / 100:0.00} m");
            DimV(g, m.P(0, 0), m.P(0, ly), -16, $"{ly / 100:0.00} m");

            // --- Corte
            double w = Math.Max(4 * e, 100);
            Map k = Map.Fit(Inset(b, 30, 36, 30, 64), 0, 0, w, e);
            ConcreteRect(g, k, 0, 0, w, e);
            double ci = v.D("recInf"), cs = v.D("recSup");
            double dIX = v.BarMm("bIX") / 10, dIY = v.BarMm("bIY") / 10, dSX = v.BarMm("bSX") / 10, dSY = v.BarMm("bSY") / 10;
            if (v.B("iX")) using (var p = new Pen(Steel, Math.Max(2f, k.L(dIX)))) g.DrawLine(p, k.P(2, ci + dIX / 2), k.P(w - 2, ci + dIX / 2));
            if (v.B("iY")) DotsAlong(g, k, w, v.D("sIY"), ci + (v.B("iX") ? dIX : 0) + dIY / 2, dIY, Tie);
            if (v.B("sX")) using (var p = new Pen(Steel2, Math.Max(2f, k.L(dSX)))) g.DrawLine(p, k.P(2, e - cs - dSX / 2), k.P(w - 2, e - cs - dSX / 2));
            if (v.B("sY")) DotsAlong(g, k, w, v.D("sSY"), e - cs - (v.B("sX") ? dSX : 0) - dSY / 2, dSY, Color.FromArgb(130, 60, 170));
            DimV(g, k.P(w, 0), k.P(w, e), 16, $"e = {e:0.#} cm");

            string Layer(string chk, string bar, string sep) =>
                v.B(chk) ? $"Ø {v.BarName(bar)} @ {v.D(sep):0.#} cm" : "—";
            Lines(g, new RectangleF(b.X + 10, b.Bottom - 62, b.Width - 14, 60),
                ($"Inf. X: {Layer("iX", "bIX", "sIX")}   Inf. Y: {Layer("iY", "bIY", "sIY")}", Steel),
                ($"Sup. X: {Layer("sX", "bSX", "sSX")}   Sup. Y: {Layer("sY", "bSY", "sSY")}", Steel2),
                ("Líneas continuas: inferior · discontinuas: superior", Ink));
        }

        private static void Grid(Graphics g, Map m, double lx, double ly, bool on, double sep, bool alongX, Color col, bool dashed)
        {
            if (!on || sep <= 0) return;
            using (var p = new Pen(col, 1.2f))
            {
                if (dashed) p.DashStyle = DashStyle.Dash;
                double max = alongX ? ly : lx;
                double step = sep;
                // Si la separación es muy pequeña en pantalla se dibuja una de cada n barras (esquema).
                while (m.L(step) < 9) step *= 2;
                for (double t = step / 2 + (dashed ? step / 4 : 0); t < max; t += step)
                {
                    if (alongX) g.DrawLine(p, m.P(0, t), m.P(lx, t));
                    else g.DrawLine(p, m.P(t, 0), m.P(t, ly));
                }
            }
        }

        private static void DotsAlong(Graphics g, Map k, double w, double sep, double z, double d, Color col)
        {
            if (sep <= 0) return;
            for (double x = sep / 2; x < w; x += sep) Dot(g, k.P(x, z), Math.Max(2.2f, k.L(d / 2)), col);
        }

        // ================================================================== MURO

        public static void Muro(Graphics g, RectangleF r, FormValues v, double L, double t, double H)
        {
            SplitH(r, 0.42f, out RectangleF a, out RectangleF e);
            Caption(g, a, "CORTE HORIZONTAL (detalle)");
            Caption(g, e, "ELEVACIÓN");

            double c = v.D("rec"), dv = v.BarMm("barV") / 10, dh = v.BarMm("barH") / 10;
            double sv = Math.Max(1, v.D("sepV")), sh = Math.Max(1, v.D("sepH"));
            bool doble = v.B("doble");

            // --- Detalle de planta (extremo del muro)
            double w = Math.Min(L, Math.Max(4 * t, 80));
            Map m = Map.Fit(Inset(a, 30, 60, 30, 110), 0, 0, w, t);
            ConcreteRect(g, m, 0, 0, w, t);
            double leg = v.D("gancho");
            var faces = doble ? new[] { 1, -1 } : new[] { 0 };
            foreach (int s in faces)
            {
                double yv = s > 0 ? t - c - dv / 2 : s < 0 ? c + dv / 2 : t / 2;
                double yh = s > 0 ? t - c - dv - dh / 2 : s < 0 ? c + dv + dh / 2 : t / 2 + dv / 2 + dh / 2;
                double maxLeg = s == 0 ? t / 2 - c - dh : t - 2 * c - 2 * dv - dh;
                double lg = Math.Max(0, Math.Min(leg, maxLeg)) * (s > 0 ? -1 : 1);
                using (var p = new Pen(Tie, Math.Max(2f, m.L(dh))))
                {
                    var pts = new List<PointF>();
                    if (Math.Abs(lg) > 0.1) pts.Add(m.P(c, yh + lg));
                    pts.Add(m.P(c, yh));
                    pts.Add(m.P(w, yh));
                    g.DrawLines(p, pts.ToArray());
                }
                for (double x = c + dh + dv / 2; x < w; x += sv)
                    Dot(g, m.P(x, yv), Math.Max(2.5f, m.L(dv / 2)), Steel);
            }
            DimV(g, m.P(0, 0), m.P(0, t), -16, $"t = {t:0.#} cm");
            using (var p = new Pen(Dim, 1f) { DashStyle = DashStyle.Dash })
                g.DrawLine(p, m.P(w, -2), m.P(w, t + 2));

            Lines(g, new RectangleF(a.X + 12, a.Bottom - 80, a.Width - 16, 78),
                (doble ? "Doble malla" : "Malla simple (al centro)", Ink),
                ($"Vertical: Ø {v.BarName("barV")} @ {sv:0.#} cm", Steel),
                ($"Horizontal: Ø {v.BarName("barH")} @ {sh:0.#} cm", Tie),
                ($"Recubrimiento {c:0.#} cm", Ink));

            // --- Elevación
            double anc = v.D("anc"), emp = v.D("emp");
            Map k = Map.Fit(Inset(e, 40, 46, 30, 60), 0, -Math.Max(anc, 0), L, H + Math.Max(anc, 0) + Math.Max(emp, 0));
            ConcreteRect(g, k, 0, 0, L, H);
            double stepV = sv; while (k.L(stepV) < 5) stepV *= 2;
            double stepH = sh; while (k.L(stepH) < 5) stepH *= 2;
            using (var p = new Pen(Steel, 1.3f))
                for (double x = c + dv / 2; x <= L - c; x += stepV)
                    g.DrawLine(p, k.P(x, anc > 0 ? -anc : c), k.P(x, emp > 0 ? H + emp : H - c));
            using (var p = new Pen(Tie, 1.3f))
                for (double z = c + dh / 2; z <= H - c; z += stepH)
                    g.DrawLine(p, k.P(c, z), k.P(L - c, z));
            DimH(g, k.P(0, 0), k.P(L, 0), anc > 0 ? 18 + k.L(anc) : 18, $"{L / 100:0.00} m");
            DimV(g, k.P(L, 0), k.P(L, H), 18, $"{H / 100:0.00} m");

            int nV = Count(L - 2 * c - 2 * dh - dv, sv) * faces.Length;
            int nH = Count(H - 2 * c - dh, sh) * faces.Length;
            Lines(g, new RectangleF(e.X + 12, e.Bottom - 40, e.Width - 16, 36),
                ($"≈ {nV} barras verticales y {nH} horizontales", Ink));
        }

        // ================================================================== Primitivas propias del acero

        private static int Count(double length, double sep) =>
            length <= 0 ? 0 : (int)Math.Ceiling(length / sep - 1e-6) + 1;

        private static void Stirrup(Graphics g, Map m, double x0, double y0, double x1, double y1, double ds) =>
            Stirrup(g, m, x0, y0, x1, y1, ds, Tie);

        private static void Stirrup(Graphics g, Map m, double x0, double y0, double x1, double y1, double ds, Color color)
        {
            PointF a = m.P(x0, y1), b = m.P(x1, y0);
            float w = Math.Max(1.6f, m.L(ds));
            float rad = Math.Max(3f, m.L(ds * 2));
            using (var path = new GraphicsPath())
            {
                var rc = new RectangleF(a.X, a.Y, b.X - a.X, b.Y - a.Y);
                float d = Math.Min(rad * 2, Math.Min(rc.Width, rc.Height) / 2);
                path.AddArc(rc.X, rc.Y, d, d, 180, 90);
                path.AddArc(rc.Right - d, rc.Y, d, d, 270, 90);
                path.AddArc(rc.Right - d, rc.Bottom - d, d, d, 0, 90);
                path.AddArc(rc.X, rc.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                using (var p = new Pen(color, w)) g.DrawPath(p, path);
                // Ganchos a 135° en la esquina superior izquierda.
                float hk = Math.Max(8f, m.L(ds * 6));
                using (var p = new Pen(color, w) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    g.DrawLine(p, rc.X + d / 3, rc.Y + d / 3, rc.X + d / 3 + hk * 0.7f, rc.Y + d / 3 + hk * 0.7f);
                    g.DrawLine(p, rc.X + d / 2.2f, rc.Y + d / 5, rc.X + d / 2.2f + hk * 0.7f, rc.Y + d / 5 + hk * 0.7f);
                }
            }
        }

        private static void Row(Graphics g, Map m, int n, double x0, double x1, double y, double db, Color col)
        {
            if (n <= 0) return;
            float rad = Math.Max(2.5f, m.L(db / 2));
            for (int i = 0; i < n; i++)
            {
                double x = n == 1 ? (x0 + x1) / 2 : x0 + (x1 - x0) * i / (n - 1);
                Dot(g, m.P(x, y), rad, col);
            }
        }
    }
}
