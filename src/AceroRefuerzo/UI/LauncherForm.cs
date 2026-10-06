using System.Drawing;
using System.Windows.Forms;
using AceroRefuerzo.Core;

namespace AceroRefuerzo.UI
{
    /// <summary>Menú principal "Acero integral": elige el elemento a armar desde su figura.</summary>
    internal sealed class LauncherForm : Form
    {
        public ElementKind Selected { get; private set; }

        public LauncherForm()
        {
            Text = "ACERO Refuerzo";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = MaximizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.White;
            ClientSize = new Size(760, 300);

            var header = new Panel { Dock = DockStyle.Top, Height = 62, BackColor = Theme.Header };
            header.Controls.Add(new Label
            {
                Text = "Colocación de acero de refuerzo", Font = new Font("Segoe UI Semibold", 14f),
                ForeColor = Color.White, AutoSize = true, Location = new Point(16, 8)
            });
            header.Controls.Add(new Label
            {
                Text = "Elija el tipo de elemento estructural a armar",
                ForeColor = Color.FromArgb(205, 215, 230), AutoSize = true, Location = new Point(18, 38)
            });

            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14, 18, 14, 10), WrapContents = false };
            Add(flow, ElementKind.Viga, "Vigas");
            Add(flow, ElementKind.Columna, "Columnas");
            Add(flow, ElementKind.Zapata, "Zapatas");
            Add(flow, ElementKind.Losa, "Losas de techo");
            Add(flow, ElementKind.Muro, "Muros estructurales");

            Controls.Add(flow);
            Controls.Add(header);

            var cancel = new Button { DialogResult = DialogResult.Cancel, Size = new Size(1, 1), Location = new Point(-10, -10) };
            Controls.Add(cancel);
            CancelButton = cancel;
        }

        private void Add(FlowLayoutPanel flow, ElementKind kind, string text)
        {
            var b = new Button
            {
                Text = text,
                Image = Icons.Large(kind, 96),
                TextImageRelation = TextImageRelation.ImageAboveText,
                Size = new Size(138, 168),
                Margin = new Padding(4),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(247, 248, 250),
                Font = new Font("Segoe UI Semibold", 9.5f),
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderColor = Color.FromArgb(220, 223, 228);
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(253, 236, 236);
            b.Click += (s, e) => { Selected = kind; DialogResult = DialogResult.OK; Close(); };
            flow.Controls.Add(b);
        }
    }
}
