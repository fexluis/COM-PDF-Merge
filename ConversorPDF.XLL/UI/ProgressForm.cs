using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Drawing;
using System.Windows.Forms;
using ConversorPDF.XLL;
namespace ConversorPDF.XLL.UI
{
    public partial class ProgressForm : Form
    {
        private Panel pnlBar;
        private Label lblTitle;
        private Label lblPercentage;
        private Label lblETA;
        private Panel pnlBarBackground;
        private PictureBox picLogo;
        private double maxProgress;

        public ProgressForm(string title, double max = 100)
        {
            this.maxProgress = max;
            InitializeComponent(title);
        }

        private void InitializeComponent(string title)
        {
            this.Text = "ESAP - " + title;
            this.Size = new Size(450, 180);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;

            // Logotipo de ESAP
            picLogo = new PictureBox()
            {
                Size = new Size(100, 40),
                Location = new Point(320, 10),
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = GetImageResource("ConversorPDF.XLL.img.esap.png")
            };

            lblTitle = new Label()
            {
                Text = title,
                Location = new Point(20, 20),
                Size = new Size(300, 25),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 153) // ESAP Blue
            };

            pnlBarBackground = new Panel()
            {
                Location = new Point(20, 60),
                Size = new Size(400, 20),
                BackColor = Color.FromArgb(230, 230, 230),
                BorderStyle = BorderStyle.FixedSingle
            };

            // Usamos un Panel en lugar de ProgressBar para tener control total del color
            pnlBar = new Panel()
            {
                Location = new Point(0, 0),
                Size = new Size(0, 18),
                BackColor = Color.FromArgb(0, 51, 153) // Mismo azul que el texto
            };
            pnlBarBackground.Controls.Add(pnlBar);

            lblPercentage = new Label()
            {
                Text = "0%",
                Location = new Point(20, 90),
                Size = new Size(50, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 153)
            };

            lblETA = new Label()
            {
                Text = "Calculando tiempo restante...",
                Location = new Point(150, 90),
                Size = new Size(270, 20),
                TextAlign = ContentAlignment.TopRight,
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                ForeColor = Color.FromArgb(0, 51, 153)
            };

            this.Controls.Add(picLogo);
            this.Controls.Add(lblTitle);
            this.Controls.Add(pnlBarBackground);
            this.Controls.Add(lblPercentage);
            this.Controls.Add(lblETA);
        }

        private Image GetImageResource(string name)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(name))
                {
                    if (stream == null) return null;
                    
                    // Clone the image to avoid GDI+ errors when stream is disposed
                    using (var originalImage = Image.FromStream(stream))
                    {
                        return new Bitmap(originalImage);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"No se pudo cargar el recurso de imagen {name}: {ex.Message}");
                return null;
            }
        }

        public void UpdateProgress(double current, string etaText)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(current, etaText)));
                return;
            }

            int percent = (int)((current / maxProgress) * 100);
            if (percent > 100) percent = 100;

            pnlBar.Width = (int)((current / maxProgress) * pnlBarBackground.Width);
            lblPercentage.Text = percent + "%";
            lblETA.Text = etaText;

            if (current >= maxProgress)
            {
                this.Hide();
            }
        }
    }
}