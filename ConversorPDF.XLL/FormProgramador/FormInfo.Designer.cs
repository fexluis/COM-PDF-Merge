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
namespace ConversorPDF.XLL.FormProgramador
{
    partial class FormInfo
    {
        private System.ComponentModel.IContainer components = null;
        private PictureBox picHero;
        private Label lblAppName;
        private Label lblVersion;
        private Label lblDetails;
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
            this.lblVersion = new Label();
            this.lblDetails = new Label();
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

            // lblVersion
            this.lblVersion.Font = new Font("Segoe UI", 10F);
            this.lblVersion.Location = new Point(20, 220);
            this.lblVersion.Size = new Size(320, 20);
            this.lblVersion.Text = $"Versión {App.APP_VERSION}";
            this.lblVersion.TextAlign = ContentAlignment.MiddleCenter;

            // lblDetails
            this.lblDetails.Font = new Font("Segoe UI", 9F);
            this.lblDetails.Location = new Point(20, 250);
            this.lblDetails.Size = new Size(320, 150);
            this.lblDetails.Text = $"Usuario: {App.DEFAULT_USER}\n" +
                               "Plataforma: 64-bit x86 Windows\n" +
                               "Registro: 80babf47-af47-4518-a228-9258c54f8726\n" +
                               "Autor: Luis Eduardo Bechara Polanía\n" +
                               "Contacto: luis.bechara@esap.edu.co\n" +
                               "Escuela Superior de Administración Pública\n\n" +
                               "ESAP 2026";
            this.lblDetails.TextAlign = ContentAlignment.TopCenter;

            // btnClose
            this.btnClose.Location = new Point(130, 410);
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
            this.ClientSize = new Size(360, 460);
            this.Controls.Add(this.picHero);
            this.Controls.Add(this.lblAppName);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.lblDetails);
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