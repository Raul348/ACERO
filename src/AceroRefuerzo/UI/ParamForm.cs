using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AceroRefuerzo.Core;

namespace AceroRefuerzo.UI
{
    /// <summary>
    /// Ventana de datos genérica: a la izquierda la figura del elemento (se redibuja con cada cambio)
    /// y a la derecha los parámetros. Recuerda los últimos valores usados en %AppData%\AceroRefuerzo.
    /// </summary>
    public sealed class ParamForm : Form
    {
        private enum Kind { Num, Txt, Bar, Chk, Choice }

        private readonly string _id;
        private readonly IList<BarItem> _bars;
        private readonly TableLayoutPanel _grid;
        private readonly Panel _preview;
        private readonly Dictionary<string, (Kind kind, Control ctrl)> _inputs = new Dictionary<string, (Kind, Control)>();

        /// <summary>Dibuja la figura del elemento con los valores actuales.</summary>
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Action<Graphics, RectangleF, FormValues> Figure { get; set; }

        public ParamForm(string id, string title, string subtitle, IList<BarItem> bars)
        {
            _id = id;
            _bars = bars;

            Text = "ACERO Refuerzo - " + title;
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(1100, 700);
            MinimumSize = new Size(900, 560);
            BackColor = Color.White;

            // Cabecera
            var header = new Panel { Dock = DockStyle.Top, Height = 66, BackColor = Theme.Header };
            header.Controls.Add(new Label
            {
                Text = title, Font = new Font("Segoe UI Semibold", 15f), ForeColor = Color.White,
                AutoSize = true, Location = new Point(16, 8), BackColor = Color.Transparent
            });
            header.Controls.Add(new Label
            {
                Text = subtitle, ForeColor = Color.FromArgb(205, 215, 230), AutoSize = true,
                Location = new Point(18, 42), BackColor = Color.Transparent
            });

            // Botones
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom, Height = 54, FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 10, 12, 10), BackColor = Color.FromArgb(244, 245, 247)
            };
            var ok = new Button
            {
                Text = "Colocar acero", Width = 150, Height = 32, DialogResult = DialogResult.OK,
                BackColor = Theme.Accent, ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9.5f)
            };
            ok.FlatAppearance.BorderSize = 0;
            var cancel = new Button { Text = "Cancelar", Width = 110, Height = 32, DialogResult = DialogResult.Cancel };
            buttons.Controls.Add(ok);
            buttons.Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            ok.Click += OnOk;

            // Cuerpo: figura | parámetros
            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 57));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            _preview = new BufferedPanel { Dock = DockStyle.Fill, BackColor = Color.White, Margin = new Padding(10) };
            _preview.Paint += PaintPreview;
            _preview.Resize += (s, e) => _preview.Invalidate();

            var scroller = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(6, 8, 14, 8), BackColor = Color.FromArgb(250, 250, 251) };
            _grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 3, RowCount = 0
            };
            _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 235));
            _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 175));
            _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40));
            scroller.Controls.Add(_grid);

            body.Controls.Add(_preview, 0, 0);
            body.Controls.Add(scroller, 1, 0);

            // El último control agregado se acopla primero: Fill debe ir antes que Top/Bottom.
            Controls.Add(body);
            Controls.Add(buttons);
            Controls.Add(header);

            Load += (s, e) => { LoadSettings(); _preview.Invalidate(); };
        }

        // ------------------------------------------------------------------ Construcción de campos

        public void Section(string text)
        {
            var l = new Label
            {
                Text = text.ToUpperInvariant(), AutoSize = true, ForeColor = Theme.Header,
                Font = new Font("Segoe UI Semibold", 9f), Margin = new Padding(0, 12, 0, 2)
            };
            int r = NewRow();
            _grid.Controls.Add(l, 0, r);
            _grid.SetColumnSpan(l, 3);
        }

        public void Note(string text)
        {
            var l = new Label
            {
                Text = text, AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(440, 0),
                Font = new Font("Segoe UI", 8f), Margin = new Padding(2, 0, 0, 4)
            };
            int r = NewRow();
            _grid.Controls.Add(l, 0, r);
            _grid.SetColumnSpan(l, 3);
        }

        public void Num(string key, string label, double def, string unit, int decimals = 1, double min = 0, double max = 100000)
        {
            var n = new NumericUpDown
            {
                DecimalPlaces = decimals,
                Minimum = (decimal)min,
                Maximum = (decimal)max,
                Increment = decimals == 0 ? 1m : decimals == 1 ? 0.5m : 0.05m,
                Width = 165, TextAlign = HorizontalAlignment.Right
            };
            n.Value = Clamp(n, def);
            n.ValueChanged += Changed;
            AddRow(label, n, unit);
            _inputs[key] = (Kind.Num, n);
        }

        public void Txt(string key, string label, string def)
        {
            var t = new TextBox { Text = def, Width = 205 };
            t.TextChanged += Changed;
            int r = AddRow(label, t, null);
            _grid.SetColumnSpan(t, 2);
            _inputs[key] = (Kind.Txt, t);
        }

        public void Bar(string key, string label, double defMm)
        {
            var cb = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 205 };
            foreach (BarItem b in _bars) cb.Items.Add(b);
            if (_bars.Count > 0)
                cb.SelectedItem = _bars.OrderBy(b => Math.Abs(b.Mm - defMm)).First();
            cb.SelectedIndexChanged += Changed;
            AddRow(label, cb, null);
            _grid.SetColumnSpan(cb, 2);
            _inputs[key] = (Kind.Bar, cb);
        }

        public void Chk(string key, string label, bool def)
        {
            var c = new CheckBox { Text = label, Checked = def, AutoSize = true, Margin = new Padding(2, 5, 0, 3) };
            c.CheckedChanged += Changed;
            int r = NewRow();
            _grid.Controls.Add(c, 0, r);
            _grid.SetColumnSpan(c, 3);
            _inputs[key] = (Kind.Chk, c);
        }

        public void Choice(string key, string label, string[] options, int def)
        {
            var cb = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 205 };
            cb.Items.AddRange(options);
            cb.SelectedIndex = Math.Max(0, Math.Min(def, options.Length - 1));
            cb.SelectedIndexChanged += Changed;
            AddRow(label, cb, null);
            _grid.SetColumnSpan(cb, 2);
            _inputs[key] = (Kind.Choice, cb);
        }

        private int NewRow()
        {
            int r = _grid.RowCount;
            _grid.RowCount = r + 1;
            _grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return r;
        }

        private int AddRow(string label, Control input, string unit)
        {
            int r = NewRow();
            _grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(2, 6, 4, 4) }, 0, r);
            input.Margin = new Padding(2, 3, 2, 3);
            _grid.Controls.Add(input, 1, r);
            if (!string.IsNullOrEmpty(unit))
                _grid.Controls.Add(new Label { Text = unit, AutoSize = true, ForeColor = Color.DimGray, Anchor = AnchorStyles.Left, Margin = new Padding(2, 6, 0, 4) }, 2, r);
            return r;
        }

        private static decimal Clamp(NumericUpDown n, double v) =>
            Math.Max(n.Minimum, Math.Min(n.Maximum, (decimal)v));

        private void Changed(object sender, EventArgs e) => _preview.Invalidate();

        // ------------------------------------------------------------------ Valores

        public FormValues Values()
        {
            var d = new Dictionary<string, object>();
            foreach (var kv in _inputs)
            {
                Control c = kv.Value.ctrl;
                switch (kv.Value.kind)
                {
                    case Kind.Num: d[kv.Key] = (double)((NumericUpDown)c).Value; break;
                    case Kind.Txt: d[kv.Key] = ((TextBox)c).Text; break;
                    case Kind.Bar: d[kv.Key] = ((ComboBox)c).SelectedItem as BarItem; break;
                    case Kind.Chk: d[kv.Key] = ((CheckBox)c).Checked; break;
                    case Kind.Choice: d[kv.Key] = ((ComboBox)c).SelectedIndex; break;
                }
            }
            return new FormValues(d);
        }

        private void OnOk(object sender, EventArgs e)
        {
            try
            {
                foreach (var kv in _inputs)
                {
                    if (kv.Value.kind == Kind.Bar && ((ComboBox)kv.Value.ctrl).SelectedItem == null)
                        throw new FormatException("Seleccione todos los diámetros de barra.");
                    if (kv.Value.kind == Kind.Txt && kv.Key == "dist")
                        Distribution.Parse(((TextBox)kv.Value.ctrl).Text);
                }
                SaveSettings();
            }
            catch (FormatException ex)
            {
                MessageBox.Show(this, ex.Message, "ACERO Refuerzo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
            }
        }

        // ------------------------------------------------------------------ Figura

        private void PaintPreview(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            var r = new RectangleF(0, 0, _preview.ClientSize.Width, _preview.ClientSize.Height);
            using (var border = new Pen(Color.FromArgb(225, 228, 232)))
                g.DrawRectangle(border, 0, 0, r.Width - 1, r.Height - 1);
            try
            {
                Figure?.Invoke(g, r, Values());
            }
            catch (Exception ex)
            {
                using (var b = new SolidBrush(Theme.Accent))
                    g.DrawString("No se puede dibujar la figura:\n" + ex.Message, Font, b, new RectangleF(16, 16, r.Width - 32, r.Height - 32));
            }
        }

        // ------------------------------------------------------------------ Memoria de valores

        private string SettingsPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AceroRefuerzo", _id + ".ini");

        private void SaveSettings()
        {
            try
            {
                var lines = new List<string>();
                foreach (var kv in _inputs)
                {
                    Control c = kv.Value.ctrl;
                    string val;
                    switch (kv.Value.kind)
                    {
                        case Kind.Num: val = ((NumericUpDown)c).Value.ToString(CultureInfo.InvariantCulture); break;
                        case Kind.Txt: val = ((TextBox)c).Text; break;
                        case Kind.Bar: val = (((ComboBox)c).SelectedItem as BarItem)?.Name ?? ""; break;
                        case Kind.Chk: val = ((CheckBox)c).Checked ? "1" : "0"; break;
                        default: val = ((ComboBox)c).SelectedIndex.ToString(CultureInfo.InvariantCulture); break;
                    }
                    lines.Add(kv.Key + "=" + val.Replace("\r", "").Replace("\n", " "));
                }
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
                File.WriteAllLines(SettingsPath, lines);
            }
            catch (Exception) { /* no es crítico */ }
        }

        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return;
                foreach (string line in File.ReadAllLines(SettingsPath))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq), val = line.Substring(eq + 1);
                    if (!_inputs.TryGetValue(key, out var input)) continue;
                    Control c = input.ctrl;
                    switch (input.kind)
                    {
                        case Kind.Num:
                            if (decimal.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal d))
                                ((NumericUpDown)c).Value = Math.Max(((NumericUpDown)c).Minimum, Math.Min(((NumericUpDown)c).Maximum, d));
                            break;
                        case Kind.Txt: ((TextBox)c).Text = val; break;
                        case Kind.Bar:
                            var cb = (ComboBox)c;
                            BarItem b = cb.Items.Cast<BarItem>().FirstOrDefault(x => x.Name == val);
                            if (b != null) cb.SelectedItem = b;
                            break;
                        case Kind.Chk: ((CheckBox)c).Checked = val == "1"; break;
                        case Kind.Choice:
                            var ch = (ComboBox)c;
                            if (int.TryParse(val, out int i) && i >= 0 && i < ch.Items.Count) ch.SelectedIndex = i;
                            break;
                    }
                }
            }
            catch (Exception) { /* valores por defecto */ }
        }

        private sealed class BufferedPanel : Panel
        {
            public BufferedPanel()
            {
                DoubleBuffered = true;
                ResizeRedraw = true;
            }
        }
    }

    internal static class Theme
    {
        public static readonly Color Header = Color.FromArgb(34, 47, 68);
        public static readonly Color Accent = Color.FromArgb(196, 38, 38);
    }

    /// <summary>Permite usar la ventana de Revit como propietaria de los formularios.</summary>
    internal sealed class RevitWindow : IWin32Window
    {
        public RevitWindow(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; }
    }
}
