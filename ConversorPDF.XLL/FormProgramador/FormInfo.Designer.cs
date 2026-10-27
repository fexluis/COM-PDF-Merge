using System;
using System.Drawing;
using System.Windows.Forms;
using ConversorPDF.XLL;
using ConversorPDF.XLL.Core.Helpers;
namespace ConversorPDF.XLL.FormProgramador
{
    partial class FormInfo
    {
        private System.ComponentModel.IContainer components = null;
        private PictureBox picHero;
        private Label lblAppName;
        private Label lblSlogan;
        private Label lblVersion;
        private Label lblDetails;
        private Button btnSync;
        private Button btnClose;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.picHero = new PictureBox();
            this.lblAppName = new Label();
            this.lblSlogan = new Label();
            this.lblVersion = new Label();
            this.lblDetails = new Label();
            this.btnSync = new Button();
            this.btnClose = new Button();
            this.SuspendLayout();

            // picHero
            this.picHero.Image = App.GetResourceImage("ConversorPDF.XLL.img.hero.png");
            this.picHero.Location = new Point(0, 0);
            this.picHero.Size = new Size(360, 180);
            this.picHero.SizeMode = PictureBoxSizeMode.StretchImage;

            // lblAppName
            this.lblAppName.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            this.lblAppName.Location = new Point(20, 195);
            this.lblAppName.Size = new Size(320, 25);
            this.lblAppName.Text = "ConversorPDF";
            this.lblAppName.TextAlign = ContentAlignment.MiddleCenter;

            // lblSlogan
            this.lblSlogan.Font = new Font("Segoe UI", 8F, FontStyle.Italic);
            this.lblSlogan.Location = new Point(20, 215);
            this.lblSlogan.Size = new Size(320, 25);
            this.lblSlogan.Text = "Programas menos. Gestionas más.";
            this.lblSlogan.TextAlign = ContentAlignment.MiddleCenter;


            // lblVersion
            this.lblVersion.Font = new Font("Segoe UI", 8F);
            this.lblVersion.Location = new Point(20, 245);
            this.lblVersion.Size = new Size(320, 20);
            this.lblVersion.Text = $"Versión {App.APP_VERSION}";
            this.lblVersion.TextAlign = ContentAlignment.MiddleCenter;

            // lblDetails
            this.lblDetails.Font = new Font("Segoe UI", 9F);
            this.lblDetails.Location = new Point(20, 280);
            this.lblDetails.Size = new Size(320, 150);
            this.lblDetails.Text = $"Usuario: {App.DEFAULT_USER}\n" +
                               "Autor: Luis Eduardo Bechara Polanía\n" +
                               "Contacto: luis.bechara@esap.edu.co\n" +
                               "Territorial Valle del Cauca\n" +
                               "Escuela Superior de Administración Pública\n" +
                               "ESAP 2026";
            this.lblDetails.TextAlign = ContentAlignment.TopCenter;

            // btnSync
            this.btnSync.Location = new Point(70, 435);
            this.btnSync.Size = new Size(100, 35);
            this.btnSync.Text = "SYNC";
            this.btnSync.FlatStyle = FlatStyle.Flat;
            this.btnSync.BackColor = Color.White;
            this.btnSync.FlatAppearance.BorderColor = Color.LightGray;
            this.btnSync.Cursor = Cursors.Hand;
            this.btnSync.Click += new EventHandler(this.btnSync_Click);

            // btnClose
            this.btnClose.Location = new Point(190, 435);
            this.btnClose.Size = new Size(100, 35);
            this.btnClose.Text = "CERRAR";
            this.btnClose.FlatStyle = FlatStyle.Flat;
            this.btnClose.BackColor = Color.White;
            this.btnClose.FlatAppearance.BorderColor = Color.LightGray;
            this.btnClose.Cursor = Cursors.Hand;
            this.btnClose.Click += (s, e) => this.Close();

            // FormInfo
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(360, 485);
            this.Controls.Add(this.picHero);
            this.Controls.Add(this.lblAppName);
            this.Controls.Add(this.lblSlogan);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.lblDetails);
            this.Controls.Add(this.btnSync);
            this.Controls.Add(this.btnClose);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "Información del Aplicativo";
            this.ResumeLayout(false);
        }
    }
}