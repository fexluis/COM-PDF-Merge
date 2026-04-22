using System;
using System.Drawing;
using System.Windows.Forms;

namespace ConversorPDF.XLL.UI
{
    /// <summary>
    /// Formulario modal simplificado para mostrar el progreso de operaciones pesadas.
    /// Diseñado para mantener vivo el message pump de Windows en el hilo STA de Excel,
    /// evitando bloqueos de COM y diálogos OLE.
    /// </summary>
    public partial class ProgressDialog : Form
    {
        private Label lblMessage;
        private ProgressBar pgbProgress;
        private Panel pnlBackground;

        public IProgress<string> Progress { get; }

        public ProgressDialog(string initialTitle = "Procesando...")
        {
            InitializeComponent();
            this.Text = initialTitle;
            this.Progress = new Progress<string>(UpdateStatus);
        }

        private void InitializeComponent()
        {
            this.lblMessage = new Label();
            this.pgbProgress = new ProgressBar();
            this.pnlBackground = new Panel();

            // pnlBackground
            this.pnlBackground.Dock = DockStyle.Fill;
            this.pnlBackground.Padding = new Padding(20);
            this.pnlBackground.BackColor = Color.White;

            // lblMessage
            this.lblMessage.AutoSize = false;
            this.lblMessage.Dock = DockStyle.Top;
            this.lblMessage.Height = 40;
            this.lblMessage.Text = "Iniciando operación...";
            this.lblMessage.TextAlign = ContentAlignment.MiddleLeft;
            this.lblMessage.Font = new Font("Segoe UI", 9F, FontStyle.Regular);

            // pgbProgress
            this.pgbProgress.Dock = DockStyle.Top;
            this.pgbProgress.Height = 23;
            this.pgbProgress.Style = ProgressBarStyle.Marquee;
            this.pgbProgress.MarqueeAnimationSpeed = 30;

            // ProgressDialog
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(400, 120);
            this.ControlBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;

            this.Controls.Add(this.pnlBackground);
            this.pnlBackground.Controls.Add(this.pgbProgress);
            this.pnlBackground.Controls.Add(this.lblMessage);
        }

        private void UpdateStatus(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action<string>(UpdateStatus), message);
                return;
            }
            lblMessage.Text = message;
        }

        /// <summary>
        /// Cierra el formulario de forma segura desde cualquier hilo.
        /// </summary>
        public void SafeClose()
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(SafeClose));
                return;
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
